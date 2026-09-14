"""Fit inline pictures to their actual section/cell without altering image data."""

import re

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shape import InlineShape
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph


def _length(parent, path, attribute, default):
    node = parent.find(path)
    if node is None:
        return default
    return int(node.get(qn(attribute), default // 635)) * 635


def _bounds(section, node):
    width = _length(section, qn('w:pgSz'), 'w:w', Cm(21))
    height = _length(section, qn('w:pgSz'), 'w:h', Cm(29.7))
    for side in ('left', 'right', 'gutter'):
        width -= _length(section, qn('w:pgMar'), 'w:' + side, 0)
    for side in ('top', 'bottom'):
        height -= _length(section, qn('w:pgMar'), 'w:' + side, 0)
    cols = section.find(qn('w:cols'))
    if cols is not None:
        explicit = cols.findall(qn('w:col'))
        if explicit:
            width = min(width, min(int(c.get(qn('w:w'), '1')) * 635 for c in explicit))
        else:
            count = max(1, int(cols.get(qn('w:num'), '1')))
            gap = int(cols.get(qn('w:space'), '720')) * 635
            width = (width - gap * (count - 1)) / count
    # Walk every containing cell, including nested tables; never enlarge a cell.
    for cell in node.iterancestors(qn('w:tc')):
        props = cell.find(qn('w:tcPr'))
        if props is None:
            continue
        cell_width = props.find(qn('w:tcW'))
        if cell_width is None or cell_width.get(qn('w:type')) != 'dxa':
            continue
        available = int(cell_width.get(qn('w:w'))) * 635
        table = next(cell.iterancestors(qn('w:tbl')))
        for side in ('left', 'right'):
            margin = props.find(qn('w:tcMar') + '/' + qn('w:' + side))
            if margin is None:
                margin = table.find(qn('w:tblPr') + '/' + qn('w:tblCellMar') + '/' + qn('w:' + side))
            available -= int(margin.get(qn('w:w'), '108')) * 635 if margin is not None else 108 * 635
        width = min(width, available)
    return max(Pt(1), width), max(Pt(1), height)


def _captions(paragraph):
    """Recognize only adjacent figure captions; do not join unrelated body text."""
    found, pending = [], []
    node = paragraph._p.getnext()
    while node is not None and node.tag == qn('w:p'):
        p = Paragraph(node, paragraph._parent)
        text = p.text.strip()
        if next(node.iter(qn('w:sectPr')), None) is not None or next(node.iter(qn('w:br')), None) is not None:
            break
        if next(node.iter(qn('w:drawing')), None) is not None:
            break
        if p.style.name in ('PS caption_figure', 'PS caption_en') or re.match(r'^(?:图\s*\d|Fig(?:ure|\.)?\s*\d|\{\{fig:)', text, re.I):
            found.extend(pending + [p])
            pending = []
        elif not text and len(pending) < 2:
            pending.append(p)
        else:
            break
        node = node.getnext()
    return found


def format_images(doc, template):
    config = template['figures']
    section_for = {}
    section = doc.sections[-1]._sectPr
    # A section's properties occur at its end, not at its start.
    for block in reversed(list(doc.element.body)):
        end = block.find(qn('w:pPr') + '/' + qn('w:sectPr'))
        if end is not None:
            section = end
        section_for[block] = section
    resized = 0
    for node in doc.element.body.iter(qn('w:p')):
        if next(node.iterancestors(qn('w:txbxContent')), None) is not None:
            continue
        inlines = list(node.iter(qn('wp:inline')))
        if not inlines:
            continue
        shapes = [InlineShape(n) for n in inlines if n.find('.//' + qn('pic:pic')) is not None]
        if not shapes:
            continue
        p = Paragraph(node, doc._body)
        block = node
        while block.getparent() is not doc.element.body:
            block = block.getparent()
        width, height = _bounds(section_for[block], node)
        standalone = not ''.join(node.itertext()).strip() and not any(
            next(node.iter(qn(tag)), None) is not None
            for tag in ('m:oMath', 'w:object', 'wp:anchor', 'w:fldChar', 'w:instrText', 'w:br')
        )
        captions = _captions(p) if standalone and config['keep_caption'] else []
        reserve = Pt(config['before_pt'] + config['after_pt'] + 12)
        for caption in captions:
            key = 'caption_en' if caption.style.name == 'PS caption_en' else 'caption_figure'
            style = template['styles'][key]
            # Allow at least two caption lines, plus longer caption wrapping.
            lines = max(2, (len(caption.text) * Pt(style['size']) / max(width, 1)) + 1)
            reserve += Pt(style['before'] + style['after'] + lines * style['size'] * 1.5)
        limit_width = min(Cm(config['max_width_cm']), width)
        limit_height = max(Pt(1), height - reserve)
        if config['max_height_cm']:
            limit_height = min(limit_height, Cm(config['max_height_cm']))
        # Pictures in one paragraph can be side by side; fit their total width.
        total_width = sum(s.width for s in shapes) if standalone else 0
        for shape in shapes:
            if shape.width <= 0 or shape.height <= 0:
                continue
            ratio = min(1, limit_width / (total_width or shape.width), limit_height / shape.height)
            if ratio < 1:
                shape.width, shape.height = round(shape.width * ratio), round(shape.height * ratio)
                resized += 1
        fmt = p.paragraph_format
        spacing = fmt if fmt.line_spacing_rule is not None else p.style.paragraph_format
        if spacing.line_spacing_rule == WD_LINE_SPACING.EXACTLY:
            fmt.line_spacing = spacing.line_spacing
            fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        if standalone:
            if config['center']:
                fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
                fmt.first_line_indent = fmt.left_indent = fmt.right_indent = 0
                ind = node.get_or_add_pPr().find(qn('w:ind'))
                for name in list(ind.attrib):
                    if name.endswith('Chars') or name == qn('w:hanging'):
                        del ind.attrib[name]
            fmt.space_before, fmt.space_after = Pt(config['before_pt']), Pt(config['after_pt'])
            fmt.keep_together = True
            if captions:
                fmt.keep_with_next = True
                for caption in captions[:-1]:
                    caption.paragraph_format.keep_with_next = True
                captions[-1].paragraph_format.keep_with_next = False
    return resized

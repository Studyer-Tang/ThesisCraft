"""Explicit removal of source emphasis, without removing text or document objects."""

from docx.oxml import parse_xml
from docx.oxml.ns import qn
from lxml import etree

from .xmlutil import element


def clear_source_emphasis(doc, options):
    selected = [key for key in ('bold', 'italic', 'underline') if options.get(key)]
    if not selected:
        return 0
    roots = {qn('w:' + name) for name in
             ('document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments')}
    count = 0
    for part in doc.part.package.parts:
        if not str(part.partname).startswith('/word/') or not str(part.partname).endswith('.xml'):
            continue
        root = getattr(part, '_element', None)
        if root is None:
            root = parse_xml(part.blob)
        if root.tag not in roots:
            continue
        for node in root.iter():
            if node.tag not in (qn('w:r'), qn('w:p')):
                continue
            # Native equations carry semantic math styling. Keep their XML intact.
            if any(a.tag in (qn('m:oMath'), qn('m:oMathPara')) for a in node.iterancestors()):
                continue
            owner = node
            if node.tag == qn('w:p'):
                owner = node.find(qn('w:pPr'))
                if owner is None:
                    continue
            rpr = owner.find(qn('w:rPr'))
            if rpr is None:
                if node.tag == qn('w:p'):
                    continue
                rpr = element('w:rPr')
                owner.insert(0, rpr)
            for key in selected:
                tags = {'bold': ('b', 'bCs'), 'italic': ('i', 'iCs'),
                        'underline': ('u',)}[key]
                for tag in tags:
                    # Explicit off also overrides inherited character/paragraph styles.
                    child = getattr(rpr, 'get_or_add_' + tag)()
                    child.set(qn('w:val'), 'none' if tag == 'u' else '0')
            count += 1
        if hasattr(part, '_element'):
            part._element = root
        else:
            part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return count

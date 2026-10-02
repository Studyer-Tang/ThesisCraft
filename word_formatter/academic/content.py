"""Literal body text, separated from updateable field results and native math."""

from docx.oxml.ns import qn


def body_text_blocks(root):
    """Keep paragraph boundaries, including tables/links/textboxes, across split runs.

    Field results can span paragraphs (TOCs do); their caches are deliberately not
    treated as authored text. Math and deleted revisions have separate object
    protection. This is a text preservation check, not a rendering/text order proof.
    """
    blocks = {}
    depth = 0
    for node in root.iter():
        if node.tag == qn("w:fldChar"):
            flag = node.get(qn("w:fldCharType"))
            if flag == "begin":
                depth += 1
            elif flag == "end":
                depth = max(0, depth - 1)
        if node.tag not in {qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")}:
            continue
        ancestors = list(node.iterancestors())
        # w:tab also names paragraph tab-stop definitions; those are layout only.
        if not any(a.tag == qn("w:r") for a in ancestors):
            continue
        if depth or any(a.tag in {qn("w:fldSimple"), qn("w:del"),
                                  qn("m:oMath"), qn("m:oMathPara")}
                        for a in ancestors):
            continue
        paragraph = next((a for a in ancestors if a.tag == qn("w:p")), None)
        if paragraph is None:
            continue
        if node.tag == qn("w:t"):
            value = node.text or ""
        elif node.tag == qn("w:tab"):
            value = "\t"
        elif node.tag == qn("w:br") and node.get(qn("w:type"), "textWrapping") != "textWrapping":
            continue  # Page/column breaks change pagination, not authored text.
        else:
            value = "\n"
        blocks.setdefault(paragraph, []).append(value)
    return ["".join(value) for value in blocks.values() if "".join(value)]


def paragraph_text(paragraph):
    return "".join(body_text_blocks(paragraph._p))


def content_change(before, expected):
    """Record an independently derived expectation, never a mutator's actual output."""
    return dict(content_before=before, content_after=expected)


def replace_marker_text(text, marker, replacement):
    """Markers are processed right-to-left; field caches never enter this string."""
    position = text.rfind(marker)
    if position < 0:
        raise RuntimeError("预期正文标记不存在，未放行文字改写。")
    return text[:position] + replacement + text[position + len(marker):]

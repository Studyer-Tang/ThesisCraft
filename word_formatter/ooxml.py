"""Text-only edits that preserve drawings, links, bookmarks and field nodes."""

from copy import deepcopy
from difflib import SequenceMatcher
from docx.oxml.ns import qn
from docx.text.run import Run


def paragraph_runs(paragraph):
    for element in paragraph._p.iter(qn("w:r")):
        parent = element.getparent()
        while parent is not None and parent.tag != qn("w:p"):
            parent = parent.getparent()
        if parent is paragraph._p:
            yield Run(element, paragraph)


def replace_text_nodes(runs, new_text):
    nodes = [child for run in runs for child in run._r if child.tag == qn("w:t")]
    old_text = "".join(node.text or "" for node in nodes)
    if not nodes or old_text == new_text:
        return
    owners = [i for i, node in enumerate(nodes) for _ in (node.text or "")]
    updated = [""] * len(nodes)
    for operation, i, j, a, b in SequenceMatcher(
        None, old_text, new_text, autojunk=False
    ).get_opcodes():
        if operation == "equal":
            for position in range(i, j):
                updated[owners[position]] += old_text[position]
        elif operation in ("replace", "insert"):
            owner = owners[min(i, len(owners) - 1)] if owners else 0
            updated[owner] += new_text[a:b]
    for node, value in zip(nodes, updated):
        node.text = value
        if value[:1].isspace() or value[-1:].isspace():
            node.set(qn("xml:space"), "preserve")


def split_run_at(run, position):
    """Only split plain-text runs; copy formatting inside the existing parent."""
    if any(child.tag not in (qn("w:rPr"), qn("w:t")) for child in run._r):
        return None
    clone = deepcopy(run._r)
    run._r.addnext(clone)
    right = Run(clone, run._parent)
    text = run.text
    run.text, right.text = text[:position], text[position:]
    return right


def iter_tables(parent):
    for table in parent.tables:
        yield table
        seen = set()
        for row in table.rows:
            for cell in row.cells:
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                yield from iter_tables(cell)

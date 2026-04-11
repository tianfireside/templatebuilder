"""
BC SCCR — Removal of unnecessary content
"""

import re

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _elem_text(elem):
    return "".join(t.text or "" for t in elem.iter(f"{{{W}}}t"))


def _is_pure_instruction(text):
    """True if the entire paragraph is a [bracketed instruction] with no Rule citation."""
    t = text.strip()
    if t == "[Style of Proceeding]":
        return False  # expanded by template_change
    if re.match(r"^\[.*\]$", t, re.DOTALL):
        return not re.search(r"Rule\s+\d", t)
    return False


def _clear_cell(cell):
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""


def run(doc):
    # 1. Delete purely instructional paragraphs — stop at Appendix
    to_delete = []
    for p in doc.paragraphs:
        if p.text.strip() == "Appendix":
            break
        if _is_pure_instruction(p.text):
            to_delete.append(p._element)
    for elem in to_delete:
        elem.getparent().remove(elem)

    # 2. Clear [type or print name] table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if re.search(r"type or print name", cell.text, re.IGNORECASE):
                    _clear_cell(cell)

    # 3. Remove trailing empty table
    for elem in reversed(list(doc.element.body)):
        tag = elem.tag.split("}")[-1]
        if tag == "tbl":
            if not _elem_text(elem).strip():
                doc.element.body.remove(elem)
            break
        elif tag == "p":
            if _elem_text(elem).strip():
                break

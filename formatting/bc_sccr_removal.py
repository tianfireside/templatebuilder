"""
BC SCCR — Removal of unnecessary content
"""

import re

KEEP_IF_CONTAINS = [
    "Rule 22-3",
    "Style of Proceeding",
]


def clear_cell(cell):
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""


def rewrite_cell(cell, new_text):
    """Clear all runs and write new_text into the first run of the first paragraph."""
    first_run = None
    for para in cell.paragraphs:
        for run in para.runs:
            if first_run is None:
                first_run = run
                first_run.text = new_text
            else:
                run.text = ""


def rewrite_para(para, new_text):
    """Clear all runs and write new_text into the first run."""
    first_run = None
    for run in para.runs:
        if first_run is None:
            first_run = run
            first_run.text = new_text
        else:
            run.text = ""


def run(doc):
    # Remove [brackets] from paragraphs — stop when Appendix is found
    for para in doc.paragraphs:
        t = para.text
        if t.strip() == "Appendix":
            break
        if re.search(r'\[.*?\]', t, re.DOTALL):
            if not any(keep in t for keep in KEEP_IF_CONTAINS):
                new_text = re.sub(r'\[.*?\]', '', t, flags=re.DOTALL).strip()
                rewrite_para(para, new_text)

    # Handle tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()

                # Only dots — remove entire cell
                if re.match(r'^\.*$', t):
                    clear_cell(cell)

                # Date cell — keep label only e.g. "Date: "
                elif re.search(r'dd/mmm/yyyy', t):
                    label = re.match(r'^[^\[\.]*', t).group(0)
                    rewrite_cell(cell, label)

                # Checkbox cell — keep only "lawyer for plaintiff(s)"
                elif re.search(r'\[[\s\xa0]*\].*\[[\s\xa0]*\]', t):
                    kept = re.sub(r'.*\[[\s\xa0]*\]\s*', '', t).strip()
                    rewrite_cell(cell, kept)

                # [type or print name] — remove entire cell
                elif re.search(r'type or print name', t):
                    clear_cell(cell)

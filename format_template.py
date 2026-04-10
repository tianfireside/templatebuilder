#!/usr/bin/env python3
"""
format_template.py — Format and prepare a raw template for use
Usage: python format_template.py
"""

import os
import importlib
from docx import Document

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# Map (province, court) -> formatting modules to run in order
FORMAT_HANDLERS = {
    ("BC", "SCCR"): [
        "formatting.bc_sccr_removal",
        "formatting.bc_sccr_formatting",
        "formatting.bc_sccr_template_change",
    ]
}


def parse_filename(filename):
    """Extract province and court from filename e.g. FORM1_BC_SCCR_Notice_of_Civil_Claim.docx"""
    name = os.path.splitext(filename)[0]
    parts = name.split("_")
    if len(parts) < 3:
        return None, None
    province = parts[1].upper()
    court    = parts[2].upper()
    return province, court


def select_template():
    """List all templates and let user pick one."""
    templates = []
    for court_folder in sorted(os.listdir(TEMPLATES_DIR)):
        court_dir = os.path.join(TEMPLATES_DIR, court_folder)
        if not os.path.isdir(court_dir):
            continue
        for f in sorted(os.listdir(court_dir)):
            if f.endswith(".docx"):
                templates.append(os.path.join(court_dir, f))

    if not templates:
        print("\n  Error: no templates found.\n")
        exit(1)

    print("\n  Which template?")
    for i, path in enumerate(templates, 1):
        print(f"    ({i}) {os.path.basename(path)}")

    choice = input("\n  Enter number: ").strip()
    if not choice.isdigit() or not (1 <= int(choice) <= len(templates)):
        print("\n  Invalid choice.\n")
        exit(1)

    return templates[int(choice) - 1]


def main():
    template_path = select_template()
    filename      = os.path.basename(template_path)
    province, court = parse_filename(filename)

    if not province or not court:
        print(f"\n  Error: could not parse province/court from filename: {filename}\n")
        exit(1)

    print(f"\n  Province: {province} | Court: {court}")

    key = (province, court)
    if key not in FORMAT_HANDLERS:
        print(f"\n  Error: no formatting handler found for {province} {court}.\n")
        exit(1)

    doc = Document(template_path)

    for module_name in FORMAT_HANDLERS[key]:
        handler = importlib.import_module(module_name)
        handler.run(doc)
        print(f"  ✓ {module_name}")

    doc.save(template_path)
    print(f"\n  Saved → {template_path}\n")


if __name__ == "__main__":
    main()

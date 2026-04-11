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


def process_file(template_path):
    filename        = os.path.basename(template_path)
    province, court = parse_filename(filename)

    if not province or not court:
        print(f"  SKIP  {filename}  (cannot parse province/court)")
        return False

    key = (province, court)
    if key not in FORMAT_HANDLERS:
        print(f"  SKIP  {filename}  (no handler for {province} {court})")
        return False

    try:
        doc = Document(template_path)
        for module_name in FORMAT_HANDLERS[key]:
            handler = importlib.import_module(module_name)
            handler.run(doc)
        doc.save(template_path)
        print(f"  ✓  {filename}")
        return True
    except Exception as e:
        print(f"  ERR  {filename}  → {e}")
        return False


def main():
    print("\n" + "=" * 58)
    print("  Template Formatter")
    print("=" * 58)
    print("\n  (1) Process a single template")
    print("  (2) Batch process all files in data/split/")
    mode = input("\n  Choice: ").strip()

    if mode == "2":
        split_dir = os.path.join(BASE_DIR, "data", "split")
        files = []
        for court_folder in sorted(os.listdir(split_dir)):
            court_dir = os.path.join(split_dir, court_folder)
            if not os.path.isdir(court_dir):
                continue
            for f in sorted(os.listdir(court_dir)):
                if f.endswith(".docx") and not f.startswith("~$"):
                    files.append(os.path.join(court_dir, f))

        if not files:
            print("\n  No .docx files found in data/split/\n")
            exit(1)

        print(f"\n  Found {len(files)} files.\n")
        ok = err = skip = 0
        for path in files:
            result = process_file(path)
            if result is True:   ok   += 1
            elif result is False: skip += 1

        print(f"\n  Done. {ok} processed, {skip} skipped.\n")

    else:
        template_path = select_template()
        process_file(template_path)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
add_template.py — Auto-rename and organize a raw SCCR Word template
Usage: python add_template.py
"""

import os
import re
from docx import Document

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

COURT_MAP = {
    "supreme court of british columbia": ("BC", "SCCR"),
    "provincial court of british columbia": ("BC", "BCPC"),
}


def extract_info(doc):
    form_number = None
    form_name   = None
    province    = None
    court       = None

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t:
            continue

        # Extract form number e.g. "Form 1 (Rule 3-1 (1))"
        if form_number is None:
            match = re.search(r'\bForm\s+(\d+)\b', t, re.IGNORECASE)
            if match:
                form_number = match.group(1)

        # Extract court and province
        if court is None:
            for key, (prov, court_code) in COURT_MAP.items():
                if key in t.lower():
                    province = prov
                    court    = court_code
                    break

        # Extract form name — all caps line, not a boilerplate line
        if form_name is None and t.isupper() and len(t) > 5 and "RULE" not in t:
            form_name = t.title().replace(" ", "_")

        if all([form_number, form_name, province, court]):
            break

    return form_number, form_name, province, court


def main():
    # Ask for the raw file path
    raw_path = input("\n  Path to raw Word doc: ").strip().strip('"')
    if not os.path.exists(raw_path):
        print("\n  Error: file not found.\n")
        exit(1)

    doc = Document(raw_path)
    form_number, form_name, province, court = extract_info(doc)

    # Ask for any fields that couldn't be extracted
    if form_number is None:
        form_number = input("  Form number not found. Enter it: ").strip()
    if form_name is None:
        form_name = input("  Form name not found. Enter it: ").strip().replace(" ", "_")
    if province is None:
        province = input("  Province not found. Enter it (e.g. BC): ").strip().upper()
    if court is None:
        court = input("  Court not found. Enter it (e.g. SCCR, BCPC): ").strip().upper()

    # Confirm all four fields
    print(f"\n  Please confirm:")
    print(f"    Province:    {province}")
    print(f"    Court:       {court}")
    print(f"    Form number: {form_number}")
    print(f"    Form name:   {form_name}")

    confirm = input("\n  Look correct? (y/n): ").strip().lower()
    if confirm != "y":
        print("\n  Cancelled.\n")
        exit(0)

    # Build new filename and destination
    new_name = f"FORM{form_number}_{province}_{court}_{form_name}.docx"
    dest_dir = os.path.join(TEMPLATES_DIR, court)
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, new_name)

    os.rename(raw_path, dest_path)
    print(f"\n  Moved → {dest_path}\n")


if __name__ == "__main__":
    main()

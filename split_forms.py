#!/usr/bin/env python3
"""
split_forms.py — Split a combined court forms doc into individual files
Usage: python split_forms.py
"""

import os
import re
from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "..", "Formfiller", "templates")


def is_form_header(text):
    """Returns form number if line is a Form N header, else None."""
    match = re.match(r'^Form\s+(\d+)\b', text.strip(), re.IGNORECASE)
    return match.group(1) if match else None


def extract_form_name(paragraphs):
    """Scan paragraphs for an all-caps line to use as the form name."""
    for para in paragraphs:
        t = para.text.strip()
        if t.isupper() and len(t) > 5 and "RULE" not in t:
            return t.title().replace(" ", "_")
    return "Unknown"


def split_forms(doc):
    """Split doc into list of (form_number, [paragraphs])."""
    forms = []
    current_num = None
    current_paras = []

    for para in doc.paragraphs:
        num = is_form_header(para.text)
        if num:
            if current_num is not None:
                forms.append((current_num, current_paras))
            current_num = num
            current_paras = [para]
        elif current_num is not None:
            current_paras.append(para)

    if current_num is not None:
        forms.append((current_num, current_paras))

    return forms


def save_form(source_doc, paragraphs, out_path):
    """Create a new doc with the given paragraphs and save it."""
    new_doc = Document()

    # Remove default empty paragraph
    for para in new_doc.paragraphs:
        p = para._element
        p.getparent().remove(p)

    body = new_doc.element.body
    for para in paragraphs:
        body.append(deepcopy(para._element))

    new_doc.save(out_path)


def main():
    print("\n" + "=" * 58)
    print("  SCCR Form Splitter")
    print("=" * 58)

    raw_path = input("\n  Path to combined Word doc: ").strip().strip('"')
    if not os.path.exists(raw_path):
        print("\n  Error: file not found.\n")
        exit(1)

    province = input("  Province (e.g. BC): ").strip().upper()
    court    = input("  Court (e.g. SCCR): ").strip().upper()

    dest_dir = os.path.join(BASE_DIR, "data", "split", court.lower())
    os.makedirs(dest_dir, exist_ok=True)

    doc   = Document(raw_path)
    forms = split_forms(doc)

    print(f"\n  Found {len(forms)} forms. Saving...\n")

    for form_num, paragraphs in forms:
        form_name = extract_form_name(paragraphs)
        form_name = re.sub(r'[\n\r]+', '_', form_name).strip()
        filename  = f"FORM{form_num}_{province}_{court}_{form_name}.docx"
        out_path  = os.path.join(dest_dir, filename)
        save_form(doc, paragraphs, out_path)
        print(f"  ✓ {filename}")

    print(f"\n  Done. Files saved to {dest_dir}\n")


if __name__ == "__main__":
    main()

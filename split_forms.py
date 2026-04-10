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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def get_paragraph_text(elem):
    """Get text from a w:p element."""
    return "".join(t.text or "" for t in elem.iter(f"{{{W}}}t"))


def is_form_header(text):
    """Returns form number if line is a Form N header, else None."""
    match = re.match(r'^Form\s+([\d.]+)\b', text.strip(), re.IGNORECASE)
    return match.group(1) if match else None


def extract_form_name(elements):
    """Scan elements for an all-caps paragraph to use as the form name."""
    for elem in elements:
        if elem.tag == f"{{{W}}}p":
            t = get_paragraph_text(elem).strip()
            if t.isupper() and len(t) > 5 and "RULE" not in t:
                return t.title().replace(" ", "_")
    return "Unknown"


def split_forms(doc):
    """Split doc body elements into list of (form_number, [elements])."""
    forms = []
    current_num = None
    current_elems = []

    for elem in doc.element.body:
        # Check if this is a paragraph with a Form N header
        if elem.tag == f"{{{W}}}p":
            text = get_paragraph_text(elem)
            num  = is_form_header(text)
            if num:
                if current_num is not None:
                    forms.append((current_num, current_elems))
                current_num   = num
                current_elems = [elem]
                continue

        if current_num is not None:
            current_elems.append(elem)

    if current_num is not None:
        forms.append((current_num, current_elems))

    return forms


def save_form(elements, out_path):
    """Create a new doc with the given elements and save it."""
    new_doc = Document()

    # Remove default empty paragraph
    body = new_doc.element.body
    for child in list(body):
        body.remove(child)

    for elem in elements:
        body.append(deepcopy(elem))

    new_doc.save(out_path)


def main():
    print("\n" + "=" * 58)
    print("  Form Splitter")
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

    for form_num, elements in forms:
        form_name = extract_form_name(elements)
        form_name = re.sub(r'[\n\r]+', '_', form_name).strip()
        filename  = f"FORM{form_num}_{province}_{court}_{form_name}.docx"
        out_path  = os.path.join(dest_dir, filename)
        save_form(elements, out_path)
        print(f"  ✓ {filename}")

    print(f"\n  Done. Files saved to {dest_dir}\n")


if __name__ == "__main__":
    main()

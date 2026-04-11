#!/usr/bin/env python3
"""
split_forms.py — Split a combined court forms doc into individual files
Usage: python split_forms.py
"""

import io
import os
import re
import urllib.request
from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
import pdfplumber

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


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


def get_hyperlink_urls(elements, src_part):
    """Return list of external URLs found in w:hyperlink elements."""
    urls = []
    for elem in elements:
        for hl in elem.iter(f"{{{W}}}hyperlink"):
            rid = hl.get(f"{{{R}}}id")
            if rid:
                rel = src_part.rels.get(rid)
                if rel and rel.reltype == HYPERLINK_REL_TYPE:
                    urls.append(rel.target_ref)
    return urls


def fetch_pdf_text(url):
    """Download a PDF from url and return its text content page by page."""
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = resp.read()
    pages = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages.append(text.strip())
    return pages


def collect_hyperlink_rids(elements):
    """Return the set of r:id values used by w:hyperlink elements."""
    rids = set()
    for elem in elements:
        for hl in elem.iter(f"{{{W}}}hyperlink"):
            rid = hl.get(f"{{{R}}}id")
            if rid:
                rids.add(rid)
    return rids


def is_pdf_link_only(elements, src_part):
    """Return the PDF URL if the form body is just a hyperlink to a PDF, else None."""
    urls = get_hyperlink_urls(elements, src_part)
    if not urls:
        return None
    # Check that non-header paragraphs have no meaningful text outside the hyperlink
    content_text = ""
    for elem in elements[1:]:  # skip the Form N header element
        if elem.tag == f"{{{W}}}p":
            for hl in elem.iter(f"{{{W}}}hyperlink"):
                # blank out hyperlink text so we can check surrounding text
                for t in hl.iter(f"{{{W}}}t"):
                    t.text = ""
            content_text += get_paragraph_text(elem).strip()
    if content_text == "":
        return urls[0]
    return None


def save_form(elements, out_path, src_part):
    """Create a new doc with the given elements and save it.

    If the form body is just a hyperlink to a PDF, fetch the PDF and write its
    text into the document instead.  Otherwise copy elements as normal and fix
    up hyperlink relationship IDs.
    """
    new_doc = Document()
    body = new_doc.element.body
    for child in list(body):
        body.remove(child)

    pdf_url = is_pdf_link_only(elements, src_part)
    if pdf_url:
        # Write the Form N header paragraph first
        body.append(deepcopy(elements[0]))
        # Then append PDF text as plain paragraphs
        pages = fetch_pdf_text(pdf_url)
        for page_text in pages:
            for line in page_text.splitlines():
                new_doc.add_paragraph(line)
    else:
        for elem in elements:
            body.append(deepcopy(elem))

        # Copy hyperlink relationships
        rids_needed = collect_hyperlink_rids(elements)
        rid_map = {}
        for rid in rids_needed:
            try:
                rel = src_part.rels[rid]
            except KeyError:
                continue
            if rel.reltype == HYPERLINK_REL_TYPE:
                new_rid = new_doc.part.relate_to(rel.target_ref, HYPERLINK_REL_TYPE, is_external=True)
                rid_map[rid] = new_rid

        if rid_map:
            for hl in body.iter(f"{{{W}}}hyperlink"):
                old = hl.get(f"{{{R}}}id")
                if old and old in rid_map:
                    hl.set(f"{{{R}}}id", rid_map[old])

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
        pdf_url = is_pdf_link_only(elements, doc.part)
        save_form(elements, out_path, doc.part)
        tag = " (PDF extracted)" if pdf_url else ""
        print(f"  ✓ {filename}{tag}")

    print(f"\n  Done. Files saved to {dest_dir}\n")


if __name__ == "__main__":
    main()

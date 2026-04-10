# Form System Design

## Overview

Two separate projects:

1. **Formfiller** (this repo) — the application lawyers use daily
2. **Template Builder** (separate repo) — a dev tool for processing raw court documents into usable templates

---

## Project 1: Formfiller

The end-user application. Lawyers run this to fill out court forms quickly.

### What it does
- Loads firm info from `firm.json`
- Presents a menu to select court and form
- Asks matter-specific questions
- Outputs a filled `.docx` to the matter folder

### Structure
```
fill_form.py          ← entry point
edit_firm.py          ← manage firm.json
forms/                ← one file per court+form (e.g. bc_sccr_form2.py)
templates/            ← ready-to-use .docx templates
data/                 ← firm.json, client files (gitignored)
output/               ← filled forms (gitignored)
```

---

## Project 2: Template Builder (separate repo)

A foundational dev tool. Not seen by end users. Run once per jurisdiction to build the template library.

### What it does
1. **Split** — takes a raw Word doc containing multiple forms, detects each `Form N` heading, and splits into individual files
2. **Remove** — strips unnecessary content (instructions, placeholders, boilerplate)
3. **Format** — applies consistent formatting
4. **Template change** — inserts `{{placeholders}}` for dynamic fields

### Key design decisions
- Works on any jurisdiction, not just SCCR
- Each jurisdiction gets its own processing modules (e.g. `bc_sccr_removal.py`)
- Split script detects new forms by finding `Form N` pattern in paragraphs — reactive, not line-number dependent
- Output goes directly into Formfiller's `templates/` folder

### Structure
```
split_forms.py            ← splits a combined doc into individual form files
format_template.py        ← orchestrator: runs removal → formatting → template change
add_template.py           ← renames and organizes a single raw template
formatting/
├── bc_sccr_removal.py
├── bc_sccr_formatting.py
└── bc_sccr_template_change.py
```

### Supported jurisdictions (planned)
- BC SCCR
- BC BCPC
- (more to be added)

---

## Workflow

1. Download raw court forms (combined Word doc)
2. Run `split_forms.py` → individual raw files
3. Run `format_template.py` on each → cleaned, formatted, placeholdered templates
4. Drop templates into Formfiller's `templates/` folder
5. Lawyers use Formfiller to fill them out

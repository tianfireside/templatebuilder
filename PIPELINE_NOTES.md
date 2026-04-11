# Pipeline Build Notes — BC SCCR Forms
## How to replicate this for a new jurisdiction

---

## What the pipeline does

Takes a single combined Word document containing all court forms (e.g. Forms 1–125)
and produces individually formatted `.docx` template files ready for use in Formfiller.

Four stages, run in order:
1. **Split** — detect `Form N` headings, cut into individual files
2. **Remove** — strip instructions and boilerplate
3. **Format** — apply alignment, list styles, bold
4. **Template change** — insert `{{placeholders}}`

---

## Step 1: Get the mock-up first

**The most important lesson.** Before writing any formatting code, ask the user to
manually produce one finished example form exactly as they want it. Then compare it
line-by-line to the raw source form. Every transformation rule comes from that diff.

Without the mock-up you are guessing. With it you are just encoding.

---

## Step 2: Split (`split_forms.py`)

**How it works:**
- Iterates all body elements (paragraphs AND tables) in the source doc
- Detects a new form when a paragraph matches `^Form\s+[\d.]+`
- Deepcopies elements into a new `Document()` for each form
- Saves as `FORM{N}_{PROVINCE}_{COURT}_{NAME}.docx`

**Lessons learned:**

1. **Hyperlinks break in split files.** When you deepcopy elements, `r:id` attributes
   reference relationships that don't exist in the new document. Fix: copy hyperlink
   relationships from the source `doc.part` into the new doc using `part.relate_to()`,
   then rewrite `r:id` attributes with the new IDs.

2. **Some forms are just a hyperlink to a PDF.** For BC SCCR, 17 forms had no body
   content — just a hyperlink to a PDF on bclaws.gov.bc.ca. Fix: detect forms where
   the only content is a hyperlink, fetch the PDF, extract text with `pdfplumber`,
   write it as plain paragraphs. Install: `pip install pdfplumber`.

3. **Form names.** The script scans elements for an all-caps paragraph to use as the
   form name. Forms with no all-caps paragraph get named `Unknown` — acceptable,
   fix manually later.

4. **Decimal form numbers** (e.g. `Form 30.001`) work fine with `[\d.]+` in the regex.

---

## Step 3: Removal (`bc_sccr_removal.py`)

**What to remove:**
- Paragraphs whose ENTIRE text is `[bracketed instruction]` — but NOT if they contain
  a Rule citation (e.g. `Rule 22-3`)
- `[type or print name]` table cells — clear them
- Trailing empty tables at end of document

**Critical:** Stop at `Appendix`. Everything after that is statutory text, leave it alone.

**Lesson:** The `[Style of Proceeding]` placeholder looks like an instruction but must
NOT be deleted — it gets expanded in the template change stage. Add it as an explicit
exception.

---

## Step 4: Formatting (`bc_sccr_formatting.py`)

**Alignment:**
- Form header line (`Form N...`) → centered
- All-caps title → centered
- `[Rule N-N...]` notice → centered
- `File Number:` and `Registry:` lines → right-aligned (set in template_change
  when creating these elements, not here)

**Lists:**
- Paragraphs starting with `N  text` (number + non-breaking spaces + content)
  → strip the leading number, apply `List Number` style
- Standalone `1`, `2` blank lines → clear text, apply `List Number` style
- **Each Division/Part must restart at 1.** Use a separate `w:numId` per list group.
  Track `in_list` state: when the first numbered item follows a non-list paragraph,
  create a new `w:num` entry in the numbering XML with `<w:startOverride w:val="1"/>`.

**Dot-only cells:** Replace with `________________________` (underscores).

**Lesson:** `List Paragraph` gives indentation only. `List Number` gives actual
auto-numbering. Use `List Number`. Restarting numbering requires manipulating
`doc.part.numbering_part` XML directly — find the `abstractNumId`, create a new
`w:num` element, attach it to the first paragraph of each new list.

---

## Step 5: Template change (`bc_sccr_template_change.py`)

**Style of Proceeding expansion:**
Replace the single `[Style of Proceeding]` paragraph with three new paragraphs
(created as raw XML elements, not via python-docx API):
```
File Number:{{file_number}}       ← right-aligned
Registry:{{registry_location}}    ← right-aligned
{{Style_of_Cause}}                ← centered
```

**Placeholder rules:**
- `.................[party(ies)].................` → `{{client_name}}` (always fixed)
- `address for service: [long instruction]` → `address for service: {{address}}`
- Bare `Fax number...:` → append `{{fax_number}}`
- Bare `E-mail...:` → append `{{email}}`
- Mid-sentence dot-fills `......[desc]......` → `________________` (underscores, not
  a placeholder — the lawyer fills these in manually)
- Signature checkbox `[  ] defendant [  ] lawyer for X` → `{{lawyer_name}} lawyer for X`
- Date cell → `Date: _________________`

**Bold split:**
For address/fax/email/filed-by lines: label up to and including `:` is bold,
value after `:` is plain. Requires `_clear_runs()` then `para.add_run()` twice.
Do NOT just set `run.bold` on existing runs — the runs from deepcopy can behave
unexpectedly. Clear and rebuild.

**Placeholder name inference:**
Strip filler words, snake_case the remainder, take first 3 words. Good enough for
80% of cases. The remaining 20% are judgment calls — do those manually.

---

## What to do manually (the 20%)

- `{{lawyer_name}}` appearing in non-standard positions
- Underscore lines that are unclear or misplaced
- Forms with unusual table structures
- Any form that genuinely requires reading the legal context

Do not try to automate this. It requires domain knowledge.

---

## Applying to a new jurisdiction

1. Get the combined Word doc for the new jurisdiction
2. Ask the user to mock up one finished form
3. Run `split_forms.py` — works for any jurisdiction
4. Create new modules: `{province}_{court}_removal.py`, `_formatting.py`, `_template_change.py`
5. Add the new key to `FORMAT_HANDLERS` in `format_template.py`
6. The structure will be similar to BC SCCR — expect ~80% reuse of patterns,
   ~20% jurisdiction-specific rules

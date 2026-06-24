# Office File Inspector

[![tests](https://github.com/seunghoonchoi-phd/llm-office-qa/actions/workflows/test.yml/badge.svg)](https://github.com/seunghoonchoi-phd/llm-office-qa/actions/workflows/test.yml)
[![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![DOI](https://zenodo.org/badge/1271079447.svg)](https://zenodo.org/badge/latestdoi/1271079447)

**Check PowerPoint, Excel, and Word files for objective defects without lowering the ceiling of the AI-generated result.**

Office File Inspector catches the kind of problems that are wrong no matter what style you prefer:

- PowerPoint text pushed off the slide
- Excel `#REF!` / `#DIV/0!` errors
- Word broken fields and ragged tables
- leftover markdown or placeholders in delivered Office files

It deliberately avoids judging taste. It does not reject a deck because it has "too much text", "too many fonts", or an unusual layout. Those may be bad choices, or they may be the right choices for a dense technical document. A checker should remove obvious failures, not force every result into the same old template.

> The rule: catch the floor-breaking mistakes, but do not cap the ceiling.

The repository slug is still `llm-office-qa` for continuity, but the public product name is **Office File Inspector**.

---

## Which file should I use?

Most users only need one command:

```bash
python check_office_file.py report.pptx
python check_office_file.py model.xlsx
python check_office_file.py proposal.docx
```

If you want format-specific commands:

```bash
python tools/check_powerpoint.py report.pptx
python tools/check_excel.py model.xlsx
python tools/check_word.py proposal.docx
```

If you use an agent workflow and want automatic checks after generated Office files:

```bash
python integrations/auto_check_office_files.py
```

---

## Project structure

```text
check_office_file.py                 # main command for users
tools/
  check_powerpoint.py                # .pptx checker
  check_excel.py                     # .xlsx checker
  check_word.py                      # .docx checker
  render_powerpoint.py               # optional PPTX render helper
integrations/
  auto_check_office_files.py         # agent hook for automatic checks
examples/
  make_bad_powerpoint.py             # generate flawed test files
  make_bad_excel.py
  make_bad_word.py
docs/
  quality-philosophy.md              # design principle and taxonomy
```

---

## What it catches

### PowerPoint

| ID | Severity | Catches |
|---|---|---|
| `GEO-OFFSLIDE` | ERROR | a shape pushed off the slide canvas |
| `TXT-OVERFLOW` | WARN | text likely taller than its own box |
| `GEO-OVERLAP` | WARN | text boxes stacked so one hides another |
| `IMG-DISTORT` | WARN | image stretched off its native aspect ratio |
| `ART-MARKDOWN` | WARN | literal markdown left in the text |
| `ART-PLACEHOLDER` | WARN | placeholder or boilerplate text |
| `TXT-TINY` | WARN | explicit font below 8pt |
| `TXT-INVISIBLE` | WARN | text nearly invisible against its fill |

### Excel

| ID | Severity | Catches |
|---|---|---|
| `XL-FORMULA-ERR` | ERROR | formula errors or `#REF!` baked into formulas |
| `XL-MERGE-OVERLAP` | ERROR | overlapping merged ranges |
| `XL-BORDER-GAP` | WARN | partial borders around a data block |
| `XL-NUM-AS-TEXT` | WARN | numbers stored as text in numeric columns |
| `XL-ARTIFACT` | WARN | literal markdown or placeholder text |
| `XL-CLIP` | WARN | text likely clipped by an occupied neighbor |

### Word

| ID | Severity | Catches |
|---|---|---|
| `DOC-FIELD-ERR` | ERROR | broken field references or mojibake |
| `DOC-TABLE-RAGGED` | ERROR | rows in a table with different cell counts |
| `DOC-MARKDOWN` | WARN | literal markdown left in prose |
| `DOC-PLACEHOLDER` | WARN | unresolved placeholders or template tokens |
| `DOC-EMPTY-HEADING` | WARN | heading-styled paragraph with no text |
| `DOC-HEADING-SKIP` | WARN | outline jumps such as H1 to H3 |
| `DOC-EMPTY-TABLE` | WARN | table whose cells are all empty |
| `DOC-IMG-DISTORT` | WARN | inline image stretched off native ratio |

`WARN` means "verify this." It is not an automatic guilty verdict.

---

## What it refuses to flag

These are not hard failures in this project:

- word count or bullet count
- font-family count
- color palette taste
- layout elegance
- margin or alignment style
- prose tone
- information density
- a creative aspect-ratio or layout choice

Some teams may need strict brand rules. That is valid, but it is a separate policy. This tool is for objective defects that a stronger model would also want to avoid.

---

## Quickstart

Install dependencies:

```bash
pip install -r requirements.txt
```

Check your files:

```bash
python check_office_file.py deck.pptx
python check_office_file.py book.xlsx --json report.json
python check_office_file.py doc.docx --strict
```

Generate deliberately flawed fixtures and confirm the checks fire:

```bash
python examples/make_bad_powerpoint.py flawed.pptx
python examples/make_bad_excel.py flawed.xlsx
python examples/make_bad_word.py flawed.docx

python check_office_file.py flawed.pptx flawed.xlsx flawed.docx
```

Each flawed fixture should exit with code `1` because it contains objective defects.

Optional render check for PowerPoint:

```bash
python tools/render_powerpoint.py deck.pptx --png
```

This requires LibreOffice. PNG export also uses PyMuPDF.

---

## Agent auto-check hook

`integrations/auto_check_office_files.py` can be used as an agent `PostToolUse` hook. After a shell command creates or modifies `.pptx`, `.xlsx`, or `.docx` files, the hook checks recent files:

- `ERROR` findings exit `2` so the agent can fix the file before delivery.
- `WARN` findings are non-blocking and ask for review.

```text
python integrations/auto_check_office_files.py
```

---

## Design principle

The design note is in [`docs/quality-philosophy.md`](docs/quality-philosophy.md).

Short version:

1. A check must be objective.
2. A stronger model should still want to avoid the defect.
3. If a rule could block a better result, it should not be an error.

## Cite / License

MIT. If this is useful in your work, please cite it using [`CITATION.cff`](CITATION.cff).

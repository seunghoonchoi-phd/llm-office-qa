# llm-office-qa

[![tests](https://github.com/seunghoonchoi-phd/llm-office-qa/actions/workflows/test.yml/badge.svg)](https://github.com/seunghoonchoi-phd/llm-office-qa/actions/workflows/test.yml)
[![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**Catch an LLM's _objective_ mistakes in PowerPoint, Excel, and Word — without lobotomizing the model.**

Most "AI output checkers" quietly enforce *taste*: density limits, font-count caps, palette rules, margin grids. That is a mistake. The moment your QA layer polices style, it becomes a **shackle on a smarter model** — it punishes the very choices a more capable model makes to do better work.

`llm-office-qa` does the opposite. It flags **only objective, unambiguous defects** — the kind of thing that is wrong regardless of taste, that a strictly *more* capable model would also never want to ship. Everything that is a matter of capability or judgment, it deliberately leaves alone.

> **The thesis:** Removing obvious mistakes is QA's job. Raising the ceiling is the model's job. A good checker never confuses the two.

---

## The two tests every check must pass

A rule earns its place here only if it passes **both**:

1. **Objective** — a violation is wrong independent of style, context, or intent. (If it's arguable, it's out.)
2. **No-shackle** — a *more capable* model would still always avoid it. (If a smarter model might break it *on purpose to do better*, it's out.)

Fail either test and the rule is dropped, or narrowed to the core that passes.

---

## What it catches

Deterministic, file-only checks. No network, no model calls — it reads the file and measures.

### PowerPoint (`pptx_lint.py`)
| ID | Severity | Catches |
|---|---|---|
| `GEO-OFFSLIDE` | ERROR | a shape pushed off the slide canvas |
| `TXT-OVERFLOW` | WARN* | text taller than its own box (CJK-aware estimate) |
| `GEO-OVERLAP` | WARN | two text boxes stacked so one hides the other |
| `IMG-DISTORT` | WARN | an image stretched off its native aspect ratio |
| `ART-MARKDOWN` | WARN | literal `**`, `#`, `- `, `\|` markdown left in the text |
| `ART-PLACEHOLDER` | WARN | "click to add", empty placeholders, boilerplate |
| `TXT-TINY` | WARN | explicit font below 8pt (unreadable, not a style floor) |
| `TXT-INVISIBLE` | WARN | text-on-fill contrast below 2:1 (effectively invisible) |

### Excel (`xlsx_lint.py`)
| ID | Severity | Catches |
|---|---|---|
| `XL-FORMULA-ERR` | ERROR | `#REF!`/`#DIV/0!`/error values, or `#REF!` baked into a formula |
| `XL-MERGE-OVERLAP` | ERROR | merged ranges that overlap |
| `XL-BORDER-GAP` | WARN | a data block whose border outline survives on *some cells only* |
| `XL-NUM-AS-TEXT` | WARN | a number stored as text inside a numeric column (breaks math/sort) |
| `XL-ARTIFACT` | WARN | literal markdown / leftover placeholder in a cell |
| `XL-CLIP` | WARN | long text clipped by an occupied neighbor (wrap off) |

### Word (`docx_lint.py`)
| ID | Severity | Catches |
|---|---|---|
| `DOC-FIELD-ERR` | ERROR | broken field ("Error! Reference source not found"), mojibake (�) |
| `DOC-TABLE-RAGGED` | ERROR | a table whose rows have different cell counts |
| `DOC-MARKDOWN` | WARN | literal markdown left in the prose |
| `DOC-PLACEHOLDER` | WARN | leftover placeholder / unresolved `{{token}}` / `${var}` |
| `DOC-EMPTY-HEADING` | WARN | a heading-styled paragraph with no text |
| `DOC-HEADING-SKIP` | WARN | an outline that jumps a level (H1 → H3) |
| `DOC-EMPTY-TABLE` | WARN | a table whose cells are all empty |
| `DOC-IMG-DISTORT` | WARN | an inline image stretched off its native ratio |

\* `WARN` items marked *verify on render* are conservative heuristics. The renderer + your eyes are the ground truth for those.

---

## What it refuses to flag

These are **not** in the linter, on purpose:

> word/bullet density · font-family count · color palette · layout elegance · margins · alignment grids · aspect-ratio choice · prose quality · tone · length

These are not mistakes. They are **the model's ceiling**. A checker that polices them is shackling the next, better model. We don't.

---

## The one root cause: open-loop generation

The defects above look unrelated, but inside the machine they are **one mechanism**:

> An LLM emits a spec (XML coordinates, a cell write) for something **it cannot see** (the rendered slide, the cell's real shape) — and instead of reading the current state, it fills in its own assumption and **never closes the loop**: it doesn't read ground truth before writing, and doesn't verify the result after.

So the real discipline isn't a style rule. It's: **read the truth before you write, verify the result after.** A smarter model follows this *better*, not worse — which is exactly why it's no shackle. See [`PHILOSOPHY.md`](PHILOSOPHY.md) for the full taxonomy (A: blind geometry, B: ground-truth destruction, C: imposing your own formatting over the document's).

---

## Three layers

```
① Prevent     follow the process discipline while generating   (don't err in the first place)
② Deterministic   *_lint.py — measured straight from the file       (cheap, 100% precise)
③ Judgment    render + look with your eyes                       (the final ground truth)
```

---

## Quickstart

```bash
pip install -r requirements.txt   # python-pptx, openpyxl, python-docx, Pillow

python pptx_lint.py deck.pptx              # exit 1 if any ERROR  (--strict: WARN too)
python xlsx_lint.py book.xlsx --json r.json
python docx_lint.py doc.docx

# see it for yourself — render to PDF/PNG (needs LibreOffice)
python pptx_render.py deck.pptx --png      # -> _render/ : PDF + slide-NN.png
```

Try it on the deliberately-flawed fixtures, which the generators build from scratch (no sample files needed):

```bash
python examples/make_test_deck.py flawed.pptx && python pptx_lint.py flawed.pptx
python examples/make_test_book.py flawed.xlsx && python xlsx_lint.py flawed.xlsx
python examples/make_test_doc.py  flawed.docx && python docx_lint.py flawed.docx
```

Each linter exits `1` on these — proof the deterministic checks fire on real defects.

## Auto-lint hook (optional)

`qa_hook.py` is a drop-in [Claude Code](https://docs.claude.com/en/docs/claude-code) `PostToolUse` hook. After any shell call, it finds freshly created `.pptx/.xlsx/.docx` files, lints them, and on an **objective ERROR** exits `2` — feeding the defect back to the model to fix *before it delivers*. WARN-only findings pass non-blocking. Wire it under `~/.claude/settings.json`.

## Honest limits

- **PowerPoint:** inherited/theme font sizes and theme colors can't be read from the file — only explicit values are checked. Overflow/overlap are heuristics; the render is the truth.
- **Excel:** "a color that breaks our house style" is **not** objectively detectable from the file alone (it needs your original as a baseline) — that's a process rule, not a lint. Formula errors need Excel's cached result; a freshly written file only has the `#REF!` string.
- **Word:** reflows, so there's almost no "layout" to break — the real defects are generation artifacts and structure.

## Cite / License

MIT. If it's useful in your work, please cite it — see [`CITATION.cff`](CITATION.cff).

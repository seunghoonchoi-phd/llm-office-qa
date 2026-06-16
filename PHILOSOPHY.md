# Philosophy — "remove only the obvious mistakes"

> The single purpose of this guideline is to **remove obvious mistakes**.
> Raising the **ceiling** of the output is the job of the **model's underlying capability** — not of a rule.
> Therefore this guideline must **never become a shackle on a more capable future model.**

## 0. The two tests every rule must pass

Before adding a rule, confirm it satisfies **both**. If it fails either, drop it — or keep only the narrow core that passes.

1. **Objectivity test** — is the violation *unambiguously* wrong, independent of taste and context? (If it's arguable, it fails.)
2. **No-shackle test** — would a *more capable* model still always avoid it? (If a smarter model might break it on purpose to do better work, it fails.)

**Out of scope (these are decided by model capability; the guideline does not touch them):** information density / word count, color palette, number of font families, margins / alignment grids, aspect-ratio choice, layout elegance, length. These are not "mistakes" — they are the **ceiling**.

---

## 1. The common root: open-loop generation

The failure cases below look different on the surface, but inside the machine they are **one mechanism**:

> An LLM emits a spec (coordinate XML, a cell write) for something **it cannot see** (the rendered slide, the real shape of a cell). Instead of the **actual current state**, it inserts **its own assumption / its own template**, and **never closes the loop** — it doesn't read the ground truth before writing, and doesn't check the result after.

Almost every obvious mistake is a variant of this open loop. So the essence of the guideline is not a style rule — it is **"close the loop": read the truth before you write, verify the result after.** A smarter model follows this *more* reliably, so it is never a shackle.

---

## 2. MECE taxonomy — three categories

| | Category | LLM-language mechanism | Detection layer |
|---|---|---|---|
| **A** | Blind space / containment errors | no render-feedback or text-metric engine, so it *guesses* size and position | output lint (deterministic) |
| **B** | Ground-truth destruction | statelessness → instead of reading the user's current file, it regenerates from *its own earlier version* | process discipline |
| **C** | Imposing its own formatting over the existing one | doesn't infer/respect the document's existing formatting grammar; injects *learned default decoration* | process discipline (some lint) |

### A. Blind space / containment errors — *the model cannot see geometry*
Put N characters into a box of font F and width W, and the model has **no reflow engine** in its head for how much vertical space that takes after wrapping. So it places a box, pours text in, and assumes "it'll fit."
- **A1 off-canvas** — a shape/text runs off the slide.
- **A2 box overflow** — text is bigger than *its own box* and gets clipped or spills.
- **A3 unintended collision** — text/elements overlap so content is hidden.
- **A4 canvas/aspect mismatch** — content built for 16:9 dropped onto a 4:3 default canvas → runs off the right (detected as A1).
- **A5 media distortion** — an image stretched off its native ratio.
- **A6 (Excel sibling)** — merged-cell overflow, row height that doesn't fit wrapped text, content past the print/visible area.
→ all objective and no-shackle. **Detected by output lint** (overflow uses a conservative estimate + render confirmation).

### B. Ground-truth destruction — *regenerating its own version instead of reading the current file*
The deepest and most damaging. It **silently discards the user's manual work.**
- **B1** ignores the user's edits and reverts to *the model's own earlier version*.
- **B2** overwrites user settings/structure outside the requested scope.
- **B3** doesn't read the current file/state as the authoritative baseline before editing.
> **Rule:** the user's **current artifact is the single source of truth.** Edit it **in place, with the minimum diff.** **Do not reconstruct from your own earlier output.** **Preserve everything** outside the explicit change scope.
→ destroying the user's work is a pure "obvious mistake," and a smarter model follows the rule better → pure no-shackle. **Process discipline.**

### C. Imposing its own formatting over the existing one
The user isn't saying "never use color" — they're saying "**don't fight our document's formatting conventions.**"
- **C1** adds formatting that wasn't there / wasn't requested (arbitrary cell color), breaking the document's conventions.
- **C2** drops or partially applies existing cell styles while writing values — openpyxl's *style-blind write*, plus the classic **partial-border** artifact (borders are per-cell-edge, so applying them to only some cells of a region leaves gaps).
- **C3** unrequested changes to number format / font / column width / merges.
> **Rule:** **conform to the artifact's existing formatting system.** When you write a value, **preserve the existing cell style.** Structural formatting like borders must be applied **completely and consistently across the whole block**, or not touched at all. **Don't inject the model's default decoration** into an existing document.
→ C2's "incomplete border" is partly lintable (inconsistent edges across a contiguous block = an objective artifact). C1/C3 need a baseline → process discipline.

---

## 3. The two enforcement layers

| Layer | What | Where |
|---|---|---|
| **Output lint** (deterministic, after the fact) | objective defects measurable straight from the finished file | `pptx_lint.py`, `xlsx_lint.py`, `docx_lint.py` |
| **Process discipline** (behavior, during generation/editing) | what output alone can't catch (needs prior state or intent) | followed by the generator / a review subagent |

The key point: **B, C1, C3 cannot be caught by scoring a file.** They are *behavior rules* — they must be honored **at generation time**, which is why a linter alone is not enough.

---

## 4. Process discipline (B · C — honor while generating/editing)

1. **Read first** — when editing an artifact, *open the user's current file first.* Never rebuild from your own earlier output.
2. **Minimum diff** — change only what was asked. Leave the value/style/border of untouched cells/shapes exactly as they were.
3. **Conform to the existing style** — follow the surrounding/existing formatting grammar (color, border, number format). No unrequested decoration.
4. **Apply completely** — structural formatting like borders goes across the *whole* region with no gaps. "Some cells only" is forbidden.

---

## 5. Why these checks, and not others

Each linted check maps back to category A or C2 — the parts that are *objectively measurable from the file*. The items deliberately excluded from the linters (density, palette, font count, aspect ratio, prose quality, tone) all fail the **no-shackle test**: a more capable model may legitimately choose any of them to do better work. Policing them would punish capability. So they live nowhere in this tool — not as an ERROR, not as a WARN.

Thresholds for every objective check live in one place: the `CFG` dict at the top of each linter.

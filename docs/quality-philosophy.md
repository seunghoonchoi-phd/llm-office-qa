# Office File Inspector Quality Philosophy

Office File Inspector has one job: remove obvious Office-file defects without
turning taste into a hard rule.

The project exists because AI-generated PowerPoint, Excel, and Word files often
fail in ways that are objectively broken:

- text is outside the slide
- a formula contains `#REF!`
- a table has missing borders on only some cells
- a Word document has broken fields
- literal markdown or placeholders are left in a delivered file

Those failures should be caught early. But a checker should not reject a result
just because it uses a denser layout, more text, a different color choice, or a
structure that does not match an old template.

## The Two Tests

Every check must pass both tests.

### 1. Objectivity Test

Is the violation wrong independent of style, context, or personal preference?

Examples that pass:

- `#REF!` in an Excel formula
- text pushed outside the PowerPoint slide
- an unresolved `{{placeholder}}` token in a Word file

Examples that fail:

- too many bullets
- unusual color palette
- dense technical slide
- unconventional layout

### 2. No-Ceiling Test

Would a more capable model still want to avoid this defect?

If a stronger model might break a rule on purpose to make a better result, that
rule should not be an `ERROR`. It may be a warning, or it may be outside the
tool entirely.

The checker should raise the floor, not lower the ceiling.

## ERROR vs WARN

`ERROR` means the file has a defect that should almost always be fixed before
delivery.

`WARN` means "verify this." The issue may be a real problem, but it may also be
intentional or context-dependent.

This split prevents the tool from becoming too weak or too forceful:

- too weak: obvious broken files slip through
- too forceful: good choices are blocked because they do not match an old style

## The Common Root Cause: Open-Loop Generation

Many Office-file defects come from open-loop generation.

The model writes coordinates, cells, or document structures without seeing the
final rendered result. It guesses that text will fit. It assumes the current
file still matches an earlier version. It writes values without preserving the
existing formatting system.

The practical rule is simple:

1. Read the current file before writing.
2. Change only what was requested.
3. Preserve the surrounding style.
4. Verify the result after writing.

Office File Inspector helps with the fourth step. It cannot replace good
generation discipline, and it should not pretend to understand every design
choice.

## What Belongs in This Tool

Good candidates:

- objective geometry errors
- broken formulas and fields
- leftover generation artifacts
- structural corruption
- nonblocking warnings that ask for human review

Bad candidates:

- prose quality
- tone
- color taste
- font count
- information density
- house style without an explicit baseline

Those may matter, but they are not universal file defects.

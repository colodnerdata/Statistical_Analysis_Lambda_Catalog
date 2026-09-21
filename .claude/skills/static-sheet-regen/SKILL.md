---
name: static-sheet-regen
description: Reminder that editing a static sheet writer — any write_sheet_*.py module registered in scripts/rebuild_static_sheets.py, e.g. write_sheet_regression_instructions.py, write_sheet_modeling_concepts.py, write_sheet_diagnostic_guide.py, write_sheet_comparison_guide.py — has zero effect on any build until templates/static_sheets.xlsx is regenerated and committed. Use immediately after editing a static sheet writer.
---

# Regenerate the static template after editing a static sheet writer

The rule in full, including the per-module-CLI trap it exists to prevent, is `AGENTS.md` → *Static reference sheets — regenerate via `rebuild_static_sheets.py`, not the per-module CLI*. Read it there; this file is the reminder, not the rule.

The short version, because this one is easy to lose at the moment you finish an edit:

```
python scripts/rebuild_static_sheets.py
```

Then commit the regenerated `templates/static_sheets.xlsx` **in the same PR** as the Python change. Editing one of those writers changes nothing any build can see until that template is regenerated and committed. The authoritative list is the registry in `scripts/rebuild_static_sheets.py` — a writer missing from it is edited but never baked.

**A reminder is not a guard — so one now exists.** `tests/test_static_template_freshness.py` reads every registered writer's source (AST) and the committed template (zipfile — no Excel required) and fails when text a writer produces is absent from its sheet, so a forgotten regeneration is caught by the suite rather than by a reviewer's memory. It runs in CI. This file is still worth reading: the test tells you *after* the work, this tells you *before* it.

---
name: static-sheet-regen
description: One-step reminder that editing write_sheet_regression_instructions.py, write_sheet_modeling_concepts.py, or write_sheet_diagnostic_guide.py has zero effect on any build until templates/static_sheets.xlsx is regenerated and committed. Use immediately after editing any of those three files.
---

# Static reference sheets must be regenerated

`write_sheet_regression_instructions.py`, `write_sheet_modeling_concepts.py`, and `write_sheet_diagnostic_guide.py` write their content (`_ROWS` / `_write_template_sheet`) **only into `templates/static_sheets.xlsx`**. `build_production.py` never executes that Python content directly — it copies the already-baked sheet via `copy_static_sheet`.

**Editing one of these three modules has zero effect on any build until the template is regenerated and committed.** This has already shipped stale doc text at least twice.

## Required after any edit to these files

```
python scripts/rebuild_static_sheets.py
```

Then commit the regenerated `templates/static_sheets.xlsx` **in the same PR** as the Python change.

The per-module CLIs (for single-sheet debugging) still exist, but running one of those instead of the combined script is the exact failure mode this skill exists to prevent — it writes to the wrong place and leaves the template stale.

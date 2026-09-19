---
name: sheet-writer-conventions
description: Index of recurring mistakes when writing or editing any write_sheet_*.py module — cell colours, A1 addresses in formulas, named-range scope, spill-reader references, charts, activation/freeze calls, calculation mode. Points at the canonical rule for each rather than restating it. Use before and while editing any sheet-writer file in lambda_catalog/.
---

# Sheet writer conventions — where each rule lives

Every rule below is stated in full in `AGENTS.md`, which is always in context, and at more length in `CONTRIBUTING.md`. **This file is an index, not a second copy:** read the canonical section rather than trusting a paraphrase here. A paraphrase of a rule rots silently — nothing checks this file, and a stale line number in it is exactly how the previous version went wrong.

| About to… | Headline rule | Canonical section |
|---|---|---|
| set a cell colour | Never hard-code an RGB tuple; import from `sheet_styles.py` with the `as _NAME` alias | AGENTS.md → *Cell styling* |
| put an A1 address in a formula, CF rule, chart title, or OFFSET name | Never spell it — build from the `_C_*` constants via `_abs_ref` / `_band` / the `_A_*` anchors; the same applies to anything reading the sheet from outside | AGENTS.md → *Regression sheet heading hierarchy* |
| add a defined name | Sheet-scoped `sheet.api.Names.Add`, never `book.names.add`; every name gets a `.Comment` at the add site | AGENTS.md → *Workbook scope belongs to the catalog* |
| add a workbook-scoped catalog body | Sheet-agnostic (unqualified spec refs) unless deliberately declared sheet-scoped | AGENTS.md → same section |
| reference a spill reader (`Fit_Context`, `Fit_Design_Columns`, `Fit_Sample_Include`) from a RefersTo written before the materialization zone | Qualify it with the owning sheet at install time — `qualify_spill_reader_references` | AGENTS.md → same section |
| add or change a chart | xlwings COM only, never openpyxl; `.Text` vs `.Formula`; identity lines as a real series; selective labels as a masked overlay; title cells outside the `try/except` | AGENTS.md → *Charts — patterns and pitfalls* |
| activate a sheet or freeze panes | `safe_activate` / `safe_freeze_top_row`; freeze at the selection, never `SplitRow`; the Regression sheet's three-row freeze is its own inline `try/except` | AGENTS.md → *Guard headless/no-focus Excel calls* |
| set the calculation mode | Two different phases — see below | — |

## Calculation mode — the one piece of detail this index carries

- **Within a sheet-writing or verification session** (`build_test_models.py`, `build_demo_workbook.py`, `deep_verify.py`, `analyze_model_construction.py`): `XL_CALCULATION_MANUAL` during writes, then `XL_CALCULATION_SEMIAUTOMATIC` before that session's own save.
- **The production build is different and unconditional**: `build_common._recalculate_and_save` always runs `CalculateFullRebuild()`, and its `calc_mode` default is `XL_CALCULATION_AUTOMATIC`. Cite that function and the test that pins it — **never a line number**.

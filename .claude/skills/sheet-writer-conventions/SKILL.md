---
name: sheet-writer-conventions
description: Consolidated checklist of recurring mistakes when writing or editing any write_sheet_*.py module (colors, cell addresses, named-range scope, charts, activation/freeze calls, calculation mode, workbook-scope catalog bodies). Use before and while editing any sheet-writer file in lambda_catalog/.
---

# Sheet writer conventions

This repo's sheet writers (`write_sheet_regression.py`, `write_sheet_univariate.py`, etc.) have several conventions that are easy to violate without anything failing loudly at write time. Check all of these before finishing an edit.

## Colors
All cell colors live in `lambda_catalog/sheet_styles.py` (`HEADER_COLOR`, `SUBHDR_COLOR`, `INPUT_COLOR`, `CF_LIGHT_RED_FILL`, `CF_DARK_RED_TEXT`, `CF_YELLOW_FILL`, `CF_DARK_YELLOW_TEXT`). **Never hard-code an RGB tuple in a sheet writer.** Import with the `as _NAME` alias pattern used elsewhere (`from .sheet_styles import HEADER_COLOR as _HEADER`).

## Cell addresses in formulas
**Never spell an A1 address into a formula string.** Hand-written letters silently read the wrong cell after a column insertion — the formula still parses. Build every address from the `_C_*` column constants via `_abs_ref(row, col)` / `_band(col)` and the `_A_*` anchors at the top of the relevant writer module. Anything reading the sheet from outside (`tools/inspect_regression_sheet.py`, `lambda_catalog/analyze_regression_spec_block.py`) must import those same constants rather than keeping a parallel copy.

## Named ranges
- Sheet-scoped: `sheet.api.Names.Add(...)` — used for wiring names, `RegChart*` chart-range names, `UV_*` names, spill readers.
- Workbook-scoped: **only** `lambda_functions.json` catalog LAMBDAs, via `sync_workbook_names`. Never call `book.names.add` from a sheet writer — every build drops any workbook-scoped name that isn't a catalog function.
- **Every defined name needs a `.Comment`**, set immediately after `Names.Add`: `_nm = sheet.api.Names.Add(...); _nm.Comment = "..."`. For a constructor-closure site, the comment is the catalog entry's `notes` field verbatim. `tests/test_workbook_invariants.py::test_every_defined_name_carries_a_comment` checks the committed dist for this, so a comment-less name fails the suite, not just the Name Manager UX.
- A **workbook-scoped catalog LAMBDA must be sheet-agnostic** (unqualified spec references) unless deliberately declared sheet-scoped (`"scope": "Regression"` style, like `Base_Period_Delta`). A body that hardcodes `'Regression'!` breaks in a workbook with several Regression-shaped sheets (every sheet reads whichever is literally named "Regression") and in a workbook with none (`#NAME?`).
- A RefersTo body written **before** the materialization zone that references a spill reader (`Fit_Context`, `Fit_Design_Columns`, `Fit_Sample_Include`) must qualify it with the owning sheet at install time — see `qualify_spill_reader_references` / `SPILL_READER_NAMES` in `regression_materialization.py`. Unqualified resolution against the calling formula's sheet does NOT hold inside a RefersTo for a late-created name.

## Charts
- **xlwings COM only** — `sheet.api.ChartObjects().Add(...)`. Never `openpyxl` for a workbook that has Excel-created charts: its `load_workbook()`/`save()` rewrites the whole package and silently drops chart parts, VML, and chartUserShapes it didn't create.
- `.Text` for a static chart title; `.Formula` (pointed at a dedicated cell) for a title that needs to update dynamically. Never pass a formula string to `.Text`.
- Reference lines (e.g. `y=x`): a real data series with matching `XValues`/`Values` against a named range, `ChartType = xlXYScatterLinesNoMarkers`. Never `chart.Shapes.AddLine(...)` — a shape sits at fixed pixel coordinates and goes wrong on resize/rescale.
- Selective data labels: a masked overlay series/column, not per-point COM loops. Token is `NA()` under `ShowValue`/`ShowCategoryName`, `""` under Value-From-Cells (and if `""`, every other label element must be off or it prints on every point).
- Write chart-title formula cells **outside** the try/except guard (plain cell writes, testable via the `RecordingSheet` mock); only the actual `ChartObjects().Add(...)` call needs the guard.

## Activation / freeze panes
Use `safe_activate(sheet)` / `safe_freeze_top_row(sheet)` from `workbook_helpers.py` — both are `try/except Exception: pass` wrappers, because `Application.ActiveWindow` raises with no interactive desktop/focus and that failure must not abort `build_production_workbook()`. Exception: the Regression sheet freezes its top **three** rows via its own inline `try/except` (not `safe_freeze_top_row`, which is single-row only) — follow that inline pattern if a future sheet needs a multi-row freeze. Always freeze at the selection (`select` the cell below the frozen band, then `ActiveWindow.FreezePanes = True`) and clear any stale `FreezePanes`/`Split` first — never set `SplitRow`, which persists as `frozenSplit` instead of a true `frozen` pane state.

## Calculation mode — two different phases, don't conflate them
- **Within a sheet-writing/verification session** (`build_test_models.py`, `build_demo_workbook.py`, `deep_verify.py`, `analyze_model_construction.py`): `XL_CALCULATION_MANUAL` during writes, then `XL_CALCULATION_SEMIAUTOMATIC` before that session's own save.
- **The production build is different and unconditional**: `build_common._recalculate_and_save` sets `XL_CALCULATION_MANUAL`, runs `CalculateFullRebuild()`, then sets `calc_mode` — whose **default is `XL_CALCULATION_AUTOMATIC`** (Excel's real "Automatic", constant `-4105`) — before the final save. Confirmed at `build_production.py:395` and pinned by `test_build_production.py`. The shipped `dist/Lambda_Library.xlsx` ships in full Automatic, not semiautomatic. Don't assume the sheet-writer MANUAL→SEMIAUTOMATIC pattern also describes the final production artifact's saved calc mode.

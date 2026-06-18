# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

Generates `Lambda_Library.xlsx` — an Excel workbook containing workbook-scoped LAMBDA functions that reproduce Analysis ToolPak-style multilinear regression statistics, plus smoke-test sheets that verify each LAMBDA against Python-computed OLS expected values.

## Commands

```powershell
# Install dependencies
uv sync

# Production build: 6 user-facing sheets + LAMBDA name sync
python build_production.py

# QC build: all production sheets + 3 MLR test sheets + verification pass
python build_qc.py

# Common flags (both scripts)
python build_production.py --validate-reopen   # reopen in Excel to verify XML patch
python build_production.py --verbose           # print per-phase timing

# Write a single sheet standalone
python -m lambda_catalog.write_sheet_lambda_functions Lambda_Library.xlsx
python -m lambda_catalog.write_sheet_life_expectancy_data Lambda_Library.xlsx

# Inspect Calc columns in test sheets vs. Python-computed expected values
python tools/inspect_test_sheets.py Lambda_Library.xlsx

# Inspect the Regression sheet against Python-computed expected values
python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx

# Print Name Manager comment character counts
python tools/check_lengths.py
```

Requires Windows and desktop Excel. If the workbook is open in Excel when the build runs, the script prompts to close it and retries.

## Architecture

### Source of truth: `lambda_functions.json`

Top-level structure is `{ "functions": [...] }` — a dict with a `"functions"` key, not a bare array. Each entry has `name`, `formula_display` (human-readable, may be multi-line), `arguments`, `test_table`, `number_format`, `yields`, and `description`. The `test_table` value must be valid as both an Excel worksheet name and an Excel table name simultaneously; set it to `null` for utility functions that are not directly tested.

### Two build scripts, two outputs

| Script | Output | Sheets |
|---|---|---|
| `build_production.py` | `Lambda_Library.xlsx` (committed) | LAMBDA_functions, Life Expectancy Data, Regression Instructions, Diagnostic Guide, Version History, Regression |
| `build_qc.py` | `Lambda_Library_QC.xlsx` (gitignored) | All production sheets except Version History, plus the 3 MLR test sheets, plus verification pass |

Run the QC build after any LAMBDA addition or modification.

### Build phases

**Production** (`build_production.py`) runs two phases:

1. **xlwings phase** (`xw.App`): Opens or creates the workbook, writes all worksheets via COM, saves.
2. **XML patch phase** (`sync_workbook_names`): Directly rewrites `xl/workbook.xml` inside the `.xlsx` ZIP using `lxml`, injecting `<definedName>` elements with the LAMBDA formula text. A subsequent `_write_name_comments` pass reopens via COM just to set Name Manager comment strings.

**QC** (`build_qc.py`) adds a third phase after the XML patch:

3. **Verify phase**: Reopens the workbook, forces full recalculation, reads Calc columns via `inspect_test_sheets`, and compares against Python-computed expected values. Mismatches print as `WARNING` lines (workbook is still written). `TOLERANCE_DECIMALS = 6` is defined in `tools/inspect_test_sheets.py`.

### Key modules in `lambda_catalog/`

Beyond the `write_sheet_*.py` writers:

- `workbook_builder.py` — `load_lambda_definitions` (parses JSON), `sync_workbook_names` (XML patch), `_validate_workbook_reopen`
- `workbook_helpers.py` — COM error handling and Excel lock detection; `raise_excel_access_error` recognizes "currently in use"/"sharing violation"/"read-only" phrases to show a friendly retry prompt
- `make_test_sheet.py` — shared infrastructure for writing test tables (used by all three `write_sheet_mlr_*` writers)
- `regression_shared.py` — shared dataclasses (`RegressionVectors`, `RegressionPredictorSummary`, etc.) and the canonical 18-column `FEATURE_COLUMNS` list
- `inspection_compare.py` — shared numeric helpers (`to_float_or_none`, `compare_values`, `first_digit_deviation`) used by both inspector scripts
- `analyze_regression_sheet.py` — computes Python-expected values for all output zones of the Regression worksheet; used by `build_qc.py` verify phase and `tools/inspect_regression_sheet.py`
- `analyze_univariate.py` — scipy/numpy oracle for univariate statistics: descriptive stats, histogram binning (Sturges/Scott/FD), NLL functions for 7 distributions, MLE parameter fitting, AIC/BIC

There is no pytest suite. The QC build's verification pass is the only automated test.

### Formula translation: `lambda_catalog/lambda_formula_parser.py`

`lambda_functions.json` stores human-readable formulas. The XML patch requires a different syntax: `_xlfn.LAMBDA(...)`, required parameter references as `_xlpm.<name>`, **optional parameter references (declared with `[brackets]` in `arguments`) as `_xlop.<name>`**, and function prefixes like `_xlfn.BYROW`. The parser handles stripping whitespace outside string literals, splitting LAMBDA signatures, translating `LET` bindings (which introduce new bound names mid-expression that must be tracked to avoid double-prefixing), and applying the `XML_FUNCTION_PREFIXES` map.

Mixing up `_xlpm` vs `_xlop`, or forgetting brackets on optional arguments, causes silent Excel parse errors.

### OLS analysis and cache: `lambda_catalog/analyze_life_expectancy.py` + `lambda_catalog/analysis_cache.py`

Python-computed OLS expected values (via statsmodels) are used to populate the `(Exp.)` columns in smoke-test sheets. These are cached in `.analysis_cache.json` (gitignored) keyed on CSV SHA-256 + `_CACHE_SCHEMA_VERSION`. Bump `_CACHE_SCHEMA_VERSION` in `analysis_cache.py` whenever regression methodology changes (k values, alpha, filter logic, cached fields, or dataclass structure). Delete `.analysis_cache.json` to force a full recompute.

The `Full_Data` filter always checks all 18 feature columns for non-null, regardless of how many predictors (k) are actually used in a given regression configuration. This matches the Excel `Full_Data` boolean column in the `LifeExpectancyData` table. Do not change this to filter on only the active predictors — it will cause Excel Calc columns to diverge from Python expected values.

### Worksheets written

| Sheet | Writer | Purpose |
|---|---|---|
| `LAMBDA_functions` | `write_sheet_lambda_functions.py` | Human-readable catalog table |
| `Life Expectancy Data` | `write_sheet_life_expectancy_data.py` | WHO CSV as `LifeExpectancyData` ListObject with `Full_Data` column |
| `Regression Instructions` | `write_sheet_regression_instructions.py` | How-to guide for the Regression sheet |
| `Diagnostic Guide` | `write_sheet_diagnostic_guide.py` | Reference table of diagnostics, thresholds, and interpretation |
| `Version History` | `write_sheet_version_history.py` | Release notes (production only) |
| `Regression` | `write_sheet_regression.py` | ToolPak-style interactive regression sheet |
| `MLR_Scalar_Test` | `write_sheet_mlr_scalar_test.py` | Scalar LAMBDA smoke tests vs. OLS expected values (QC only) |
| `MLR_Vector_Outputs_Test` | `write_sheet_mlr_vector_outputs_test.py` | Vector LAMBDA smoke tests — spill formulas (QC only) |
| `MLR_Observation_Test` | `write_sheet_mlr_observation_test.py` | Observation-level LAMBDA smoke tests (QC only) |
| `Univariate` | `write_sheet_univariate.py` | Univariate analysis: descriptive stats, histogram binning, distribution fitting |

**Regression sheet layout**: three horizontal zones — inputs (cols A–B), analysis output (cols C–J), diagnostics/residual table (cols L–X), hidden helpers (cols Y–AB), charts (cols AC+). The `Allow_Intercept` toggle in D2 controls whether a VSTACK("") row is prepended to keep predictor rows aligned when intercept is FALSE; do not break this padding logic. Sheet-scoped names (`x_s`, `y`, `fil`, `Allow_Intercept`, `pred_input`) are defined fresh on every sheet write.

**Univariate sheet layout**: four horizontal zones — Col A (data input, 2000 rows, orange), Cols C–D (descriptive stats, 12 rows), Cols F–M (three side-by-side histogram bin tables: Sturges/Scott/FD), Cols P–Z (distribution fitting summary: Normal, Lognormal, Exponential, Triangular, BetaPERT with NLL/AIC/BIC; lowest-AIC row highlighted green). Sheet-scoped named ranges: `UV_Data`, `UV_n`, plus OFFSET-based chart series ranges `UV_Sturges_Edges/Counts`, `UV_Scott_Edges/Counts`, `UV_FD_Edges/Counts`.

**Test sheet row parameterization**: each test sheet has a `build_mlr_row_configs` function returning `(k, allow_intercept, expected_values)` tuples. Rows iterate over `_MLR_K_VALUES` (typically k=1,5,10,18). `x_s` is dynamic per row via `OFFSET`. Adding a new k-value requires updating `_MLR_K_VALUES` in all test sheet writers and bumping `_CACHE_SCHEMA_VERSION`.

### Naming conventions

- `build_*.py` — workbook-level entry points (open/create workbook, orchestrate sheets)
- `write_sheet_*.py` — single-sheet writers; each can also be run standalone as a module
- `lambda_catalog/` — installable package (`uv sync` installs it editable)
- `tools/` — standalone diagnostic scripts that don't affect the workbook

## Adding a new LAMBDA function

1. Add an entry to `lambda_functions.json` with `name`, `formula_display`, `arguments`, `yields`, `description`, and optionally `test_table` and `number_format`. Wrap optional parameters in `[brackets]` in `arguments`.
2. If scalar, set `"test_table": "MLR_Scalar_Test"` and add expected values to `analyze_life_expectancy.py` → `RegressionSummary` / `calculate_regression_summary`.
3. If vector, set `"test_table": "MLR_Vector_Outputs_Test"` and add to `RegressionVectors` / `calculate_regression_vectors`.
4. If observation-level, set `"test_table": "MLR_Observation_Test"` and add to the corresponding dataclass and calculation function.
5. Bump `_CACHE_SCHEMA_VERSION` in `analysis_cache.py`.
6. Update the relevant `write_sheet_mlr_*.py` to include a Calc column for the new function.
7. Run `python build_qc.py` — confirm no WARNING lines.
8. Run `python build_production.py` to rebuild the distributable.

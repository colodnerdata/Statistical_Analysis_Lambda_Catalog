# Contributing

## Setup

Requires Python 3.10+, [uv](https://github.com/astral-sh/uv). Building the Excel workbook also requires desktop Excel on Windows (xlwings uses COM automation), but running the Python test suite does not.

```powershell
uv sync
```

This installs the `lambda_catalog` package in editable mode along with all dependencies: `lxml`, `numpy`, `pandas`, `pywin32` (Windows only), `scipy`, `statsmodels`, `xlwings`, plus dev tools (`pytest`, `pytest-cov`, `pylint`).

## Running tests

The Python test suite covers the pure-Python analysis engine, formula parser, and serialization helpers. It runs on any platform without Excel.

```powershell
# Run all tests
uv run pytest

# Run with coverage (omits xlwings-dependent workbook writers)
uv run pytest --cov --cov-report=term-missing
```

Tests live in `tests/`. The current test files are:

| File | What it covers |
|---|---|
| `test_univariate.py` | NLL functions, MLE estimators, binning rules, GoF metrics |
| `test_regression_summary_metrics.py` | AIC, BIC, AICc, QQ-correlation formulas |
| `test_regression_vectors.py` | Per-coefficient statistics (SE, t-stats, p-values, CI, beta weights) |
| `test_regression_observation_vectors.py` | Observation-level diagnostics (rank fraction, normal scores, residuals) |
| `test_internal_helpers.py` | `_parse_float`, `_normalize_header`, `_validate_required_headers`, `_build_training_arrays`, `_predict_single_row` |
| `test_formula_parser.py` | LAMBDA formula → workbook.xml XML token translation |
| `test_cache_serialization.py` | JSON serialization round-trips for `RegressionVectors` and `RegressionObservationVectors` |
| `test_data_completeness_qc.py` | `calculate_data_completeness_flags` against the sample dataset |
| `test_catalog_schema.py` | `CatalogDocument` loading, validation, duplicate rejection, `test_table` rules, projection methods |
| `test_dummy_functions.py` | `Dummy_Levels`/`Dummy_Code` NA()-based error contract: formula statics, parser translation, and the pure-Python mirrors behind the `Dummy_Test` QC sheet |
| `test_lambda_catalog_plain_language.py` | All LAMBDA functions have a `plain_language_summary` in `lambda_functions.json` |
| `test_sheet_writers.py` | Sheet writer integration (conditional formatting, named ranges) |
| `test_weibull_grid_excel.py` | Weibull grid-search mechanics validation |
| `test_inspection_compare.py` | QC value comparison logic (`to_float_or_none`, `first_digit_deviation`, `compare_values`) |
| `test_independent_verification.py` | Independent numpy/scipy verification of all LAMBDA function outputs (scalars, vectors, observation diagnostics, predictor summary, prediction interval) |
| `test_qc_configs.py` | QC config generation, cross-consistency between scalar/vector/observation configs, regression sheet diagnostics, cache round-trips |

### Coverage scope

The coverage configuration in `pyproject.toml` tracks only the modules that are testable without Excel:

- `analyze_life_expectancy.py`
- `analyze_univariate.py`
- `catalog_schema.py`
- `lambda_formula_parser.py`
- `regression_shared.py`
- `analysis_cache.py`

The `write_sheet_*.py` modules, `workbook_builder.py`, `workbook_helpers.py`, `make_test_sheet.py`, `sheet_styles.py`, `inspection_compare.py`, `analyze_regression_sheet.py`, and other xlwings-dependent modules are omitted from CI coverage measurement. They are validated by the QC build instead (see below).

### CI

GitHub Actions runs the test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request. See `.github/workflows/ci.yml`. Coverage must stay at or above 70% on the tracked modules.

## Building

There are two separate build scripts with distinct purposes.

### Production build

```powershell
python build_production.py
```

Produces `Lambda_Library.xlsx` — the distributable artifact committed to the repo. Writes seven sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions
- **Life Expectancy Data** — WHO dataset as a structured table
- **Univariate Analysis** — descriptive statistics, histogram binning, and Weibull grid-search fitting
- **Regression Instructions** — step-by-step guide for adapting the sheet to new datasets
- **Diagnostic Guide** — interpretation guide for regression diagnostics
- **Version History** — changelog that travels with the workbook
- **Regression** — ToolPak-style analysis interface

No test sheets, no OLS analysis, no cache dependency.

Optional flags:

```powershell
python build_production.py --validate-reopen   # re-open workbook to verify XML patch
python build_production.py --verbose           # print per-phase timing
```

### QC build

```powershell
python build_qc.py
```

Produces `Lambda_Library_QC.xlsx` (gitignored). Writes all eleven sheets (the seven above plus `MLR_Scalar_Test`, `MLR_Vector_Outputs_Test`, `MLR_Observation_Test`, `Dummy_Test`), updates `.analysis_cache.json`, and runs the expected-vs-actual verification pass.

The `Dummy_Test` sheet is self-checking: every case is a boolean Pass formula (e.g. `=ISNA(Dummy_Levels(...))`) evaluated by Excel, and the verification pass reads the Pass cells back and reports any that are not TRUE.

The verification step forces Excel to recalculate all test formulas, reads the Calc columns, compares them against Python-computed expected values, and prints a `WARNING` line for any value that diverges beyond the tolerance band. No warnings means the LAMBDA implementations agree with statsmodels.

Optional flags:

```powershell
python build_qc.py --validate-reopen
python build_qc.py --verbose
python build_qc.py --cache path/to/.analysis_cache.json   # non-default cache location
```

Run the QC build whenever you add or modify a LAMBDA function.

## File structure

```
build_production.py          # production entry point → Lambda_Library.xlsx
build_qc.py                  # QC entry point → Lambda_Library_QC.xlsx
lambda_functions.json         # LAMBDA definitions (source of truth)
sample_data/
  Life Expectancy Data.csv   # WHO life expectancy dataset
lambda_catalog/
  catalog_schema.py          # typed document model: CatalogArgument, CatalogFunction, CatalogDocument
  regression_shared.py       # shared regression dataclasses: RegressionSummary, RegressionVectors, etc.
  sheet_styles.py            # shared cell-formatting constants (colors, conditional formatting)
  workbook_builder.py        # shared core: sync_workbook_names, workbook XML patching
  workbook_helpers.py        # shared xlwings utilities and cell formatting helpers
  analyze_life_expectancy.py # OLS engine: calculate_regression_summary, vectors, observations
  analyze_regression_sheet.py # full Regression sheet QC oracle (predictor summary, residuals, prediction interval)
  analyze_univariate.py      # univariate analysis: NLL functions, MLE estimators, binning, GoF
  analysis_cache.py          # disk cache keyed on CSV SHA-256 + schema version
  lambda_formula_parser.py   # converts display formulas to workbook XML syntax
  inspection_compare.py      # numeric comparison helpers for QC value verification
  make_test_sheet.py         # shared helpers for Excel ListObject test tables
  write_sheet_lambda_functions.py
  write_sheet_life_expectancy_data.py
  write_sheet_univariate.py
  write_sheet_regression_instructions.py
  write_sheet_diagnostic_guide.py
  write_sheet_version_history.py
  write_sheet_regression.py
  write_sheet_mlr_scalar_test.py
  write_sheet_mlr_vector_outputs_test.py
  write_sheet_mlr_observation_test.py
  write_sheet_dummy_test.py
tools/
  inspect_test_sheets.py     # scalar/vector/observation test sheet comparison (used by build_qc.py)
  inspect_regression_sheet.py # Regression sheet QC comparison (used by build_qc.py)
  inspect_univariate_sheet.py # Univariate sheet QC comparison (used by build_qc.py)
  inspect_xlsx.py            # workbook inspection utility
  check_lengths.py           # print Name Manager comment lengths
```

## File naming conventions

- `build_*.py` — workbook-level entry points that open or create an Excel workbook
- `write_sheet_*.py` — worksheet writers, each responsible for one sheet; can also be run standalone
- `lambda_catalog/` — installable package containing all writers and shared helpers

## Analysis cache

`build_qc.py` caches OLS expected values in `.analysis_cache.json` (gitignored) to avoid rerunning statsmodels on every build. The cache is keyed on the SHA-256 hash of the CSV file and a schema version constant.

The cache is invalidated automatically when:

- the CSV file content changes (SHA-256 hash mismatch)
- `_CACHE_SCHEMA_VERSION` in `analysis_cache.py` is bumped

Bump `_CACHE_SCHEMA_VERSION` whenever analysis configuration changes — k values, alpha, regression methodology, or the set of cached output fields. Delete `.analysis_cache.json` to force a full recompute at any time.

## Writing individual sheets

Each `write_sheet_*.py` module can be run standalone against any open workbook:

```powershell
python -m lambda_catalog.write_sheet_lambda_functions Lambda_Library.xlsx --definitions lambda_functions.json
python -m lambda_catalog.write_sheet_life_expectancy_data Lambda_Library.xlsx
```

## Adding a new LAMBDA function

1. Add an entry to `lambda_functions.json` with `name`, `formula_display`, `arguments`, `yields`, `description`, and optionally `test_table` and `number_format`.
2. If the function is scalar, set `"test_table": "MLR_Scalar_Test"` and add its expected value to `analyze_life_expectancy.py` → `RegressionSummary` / `calculate_regression_summary`.
3. If the function returns a vector, set `"test_table": "MLR_Vector_Outputs_Test"` and add expected values to `RegressionVectors` / `calculate_regression_vectors`.
4. Update `_CACHE_SCHEMA_VERSION` in `analysis_cache.py`.
5. Update the relevant `write_sheet_mlr_*.py` to include a Calc column for the new function.
6. Run `python build_qc.py` and confirm no WARNING lines appear.
7. Run `python build_production.py` to rebuild the distributable.

## Cell styling

All cell colors are defined once in `lambda_catalog/sheet_styles.py` and imported by every sheet writer. Never hard-code RGB tuples directly in a sheet writer.

| Constant | RGB | Usage |
|---|---|---|
| `HEADER_COLOR` | `(202, 237, 251)` | Zone / section headings — light blue, used on every sheet |
| `SUBHDR_COLOR` | `(220, 230, 241)` | Column sub-header rows within a zone |
| `INPUT_COLOR`  | `(251, 226, 213)` | User-editable input cells — light orange |
| `CF_LIGHT_RED_FILL` | `(255, 199, 206)` | Conditional formatting — failed significance / diagnostic flag |
| `CF_DARK_RED_TEXT` | `(156, 0, 6)` | Conditional formatting — text color for red-flagged cells |
| `CF_YELLOW_FILL` | `(255, 235, 156)` | Conditional formatting — borderline diagnostic flag |
| `CF_DARK_YELLOW_TEXT` | `(156, 101, 0)` | Conditional formatting — text color for yellow-flagged cells |

Import pattern:

```python
from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR
```

The `as _NAME` alias keeps existing private helpers (`_section_heading`, `_subheader_row`, etc.) unchanged.

A **section heading** is bold text with `HEADER_COLOR` fill at the default font size. The sheet title ("Univariate Analysis") is 14 pt bold with no fill — that is the only cell with a custom font size. Use the private `_section_heading(sheet, row, col, label)` helper defined in each sheet writer; do not apply the style inline.

Sheet-specific colors that differ from the shared palette (e.g., `_SUBHEADER_COLOR` in `write_sheet_diagnostic_guide.py`) remain as local constants in the relevant file.

## Regression sheet conventions

### Chart series data ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns (`$Y$3:$Y$1048576`) degrades Excel's recalculation performance and can crash the workbook on large datasets.

Instead, all chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in `$M$8`:

```python
sheet.api.Names.Add(
    Name="RegChartFitY",
    RefersTo=f"=OFFSET('{sname}'!$Y$2,1,0,'{sname}'!$M$8,1)",
)
```

This starts one row below the column header (row 2) and extends exactly `$M$8` rows — the number of filtered observations.

**Naming convention** — all OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix, distinguishing them from formula-helper names (`All_Xs`, `pred_input`, etc.):

| Name | Column | Contents |
|---|---|---|
| `RegChartQQX` | AE | Normal Scores Ranked (QQ theoretical axis) |
| `RegChartQQY` | AF | Studentized Residuals Ranked (QQ actual axis) |
| `RegChartFitY` | Y | Predicted Y — shared by multiple charts |
| `RegChartResid` | Z | Residuals |
| `RegChartActY` | X | Actual Y |
| `RegChartScaleLoc` | AG | Scale-Location |
| `RegChartCookDist` | AD | Cook's Distance |
| `RegChartLeverage` | AB | Hat Diagonal |
| `RegChartStudResid` | AC | Studentized Residuals |
| `RegChartPRESSResid` | AH | PRESS Residual |

**Scope:** all names are worksheet-scoped (created via `sheet.api.Names.Add`). Chart `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer:

```excel
Series X values: ='Regression'!RegChartFitY
Series Y values: ='Regression'!RegChartResid
```

In code, use the `_name_ref` helper in `_write_diagnostic_charts`:

```python
def _name_ref(local_name: str) -> str:
    return f"='{sname}'!{local_name}"
```

When adding a new diagnostic column or chart, add the corresponding `RegChart`-prefixed named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

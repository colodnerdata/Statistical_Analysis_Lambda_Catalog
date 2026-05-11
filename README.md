# Statistical_Analysis_Lambda_Catalog

Reproduce Analysis ToolPak-style statistics with workbook-scoped Excel LAMBDA functions.

## Build the workbook

Run `build_lambda_library.py` to open or create `Lambda_Library.xlsx`, sync the workbook Name Manager entries from `lambda_functions.json`, and write all worksheets.

The script:
- creates `Lambda_Library.xlsx` if it does not exist, or opens it if it does
- converts each JSON `formula_display` value into workbook XML syntax and writes workbook-scoped LAMBDA names to the Name Manager
- writes four worksheets: `MLR_Scalar_Test`, `MLR_Vector_Outputs_Test`, `LAMBDA_functions`, and `Life Expectancy Data`
- uses a disk cache (`.analysis_cache.json`) to skip recomputing OLS expected values when the input CSV has not changed

This script requires desktop Excel because it uses `xlwings`.

```powershell
python build_lambda_library.py
```

Use `--validate-reopen` to confirm Excel can reopen the saved workbook after the XML name patch.

```powershell
python build_lambda_library.py --validate-reopen
```

## File naming conventions

- `build_*.py` — workbook-level scripts that open or create an Excel workbook and orchestrate one or more worksheet writes
- `write_sheet_*.py` — worksheet-level scripts that write a single worksheet and can also be run standalone against any target workbook
- `lambda_catalog/` — installable Python package containing all worksheet writers and shared helpers

## Project structure

```
build_lambda_library.py          # main entry point
lambda_functions.json            # LAMBDA definitions (source of truth)
sample_data/
  Life Expectancy Data.csv       # WHO life expectancy dataset
lambda_catalog/
  analyze_life_expectancy.py     # OLS regression: RegressionSummary, RegressionVectors
  analysis_cache.py              # disk cache keyed on CSV SHA-256 + schema version
  lambda_formula_parser.py       # converts display formulas to workbook XML syntax
  make_test_sheet.py             # shared helpers for Excel ListObject test tables
  workbook_helpers.py            # shared xlwings helpers
  write_sheet_lambda_functions.py
  write_sheet_life_expectancy_data.py
  write_sheet_mlr_scalar_test.py
  write_sheet_mlr_vector_outputs_test.py
```

## Worksheets

### MLR_Scalar_Test

Smoke-tests scalar regression LAMBDA functions (e.g. `R_squared`, `SE_Regression`, `Observations`) against OLS expected values computed by `analyze_life_expectancy.py`. Each row represents one regression configuration (k predictors × intercept on/off). Columns show expected values, calculated values, per-metric match booleans, and an overall `Smoke Test` AND column.

### MLR_Vector_Outputs_Test

Smoke-tests vector regression LAMBDA functions (`Coefficients`, `SE_Coefficients`, `T_Stats`, `P_Values`, `CI_Lower`, `CI_Upper`, `CI_Excludes_Zero`) against OLS expected values. Each section represents one regression configuration; calculated values use Excel 365 dynamic-array spill formulas (`.Formula2`).

### LAMBDA_functions

Human-readable catalog of all LAMBDA definitions loaded from `lambda_functions.json`, written as a structured Excel table (`LAMBDAFunctionsCatalog`).

### Life Expectancy Data

The WHO life expectancy CSV imported as a structured Excel table (`LifeExpectancyData`) with a computed `Full_Data` boolean column indicating rows with no missing values across all 18 feature columns.

## Analysis cache

`build_lambda_library.py` caches OLS expected values in `.analysis_cache.json` (gitignored) to avoid rerunning statsmodels on every build. The cache is invalidated automatically when:

- the CSV file content changes (SHA-256 hash)
- `_CACHE_SCHEMA_VERSION` in `analysis_cache.py` is bumped

Bump `_CACHE_SCHEMA_VERSION` whenever analysis configuration changes (k values, `alpha`, regression methodology, or cached output fields). Delete `.analysis_cache.json` to force a full recompute at any time.

## Write individual sheets

Each `write_sheet_*.py` module can be run standalone against any workbook:

```powershell
python -m lambda_catalog.write_sheet_lambda_functions Lambda_Library.xlsx
python -m lambda_catalog.write_sheet_life_expectancy_data Lambda_Library.xlsx
```

## Setup

Requires Python 3.10+ and desktop Excel (Windows). Install dependencies with [uv](https://github.com/astral-sh/uv):

```powershell
uv sync
```

This installs the project as an editable package (`lambda_catalog`) along with all dependencies: `lxml`, `numpy`, `pywin32`, `statsmodels`, `xlwings`.


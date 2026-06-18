# Contributing

## Setup

Requires Python 3.10+, [uv](https://github.com/astral-sh/uv), and desktop Excel on Windows (xlwings uses COM automation).

```powershell
uv sync
```

This installs the `lambda_catalog` package in editable mode along with all dependencies: `lxml`, `numpy`, `pywin32`, `statsmodels`, `xlwings`.

## Building

There are two separate build scripts with distinct purposes.

### Production build

```powershell
python build_production.py
```

Produces `Lambda_Library.xlsx` — the distributable artifact committed to the repo. Writes three sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions
- **Life Expectancy Data** — WHO dataset as a structured table
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

Produces `Lambda_Library_QC.xlsx` (gitignored). Writes all six sheets (the three above plus `MLR_Scalar_Test`, `MLR_Vector_Outputs_Test`, `MLR_Observation_Test`), updates `.analysis_cache.json`, and runs the expected-vs-actual verification pass.

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
  workbook_builder.py        # shared core: LambdaDefinition, sync_workbook_names
  analyze_life_expectancy.py # OLS engine: RegressionSummary, RegressionVectors, etc.
  analysis_cache.py          # disk cache keyed on CSV SHA-256 + schema version
  lambda_formula_parser.py   # converts display formulas to workbook XML syntax
  make_test_sheet.py         # shared helpers for Excel ListObject test tables
  workbook_helpers.py        # shared xlwings utilities
  write_sheet_lambda_functions.py
  write_sheet_life_expectancy_data.py
  write_sheet_regression.py
  write_sheet_mlr_scalar_test.py
  write_sheet_mlr_vector_outputs_test.py
  write_sheet_mlr_observation_test.py
tools/
  inspect_test_sheets.py     # standalone comparison tool (also used by build_qc.py)
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

## Regression sheet conventions

### Chart series data ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns (`$Y$3:$Y$1048576`) degrades Excel's recalculation performance and can crash the workbook on large datasets.

Instead, all chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in `$M$8`:

```python
sheet.api.Names.Add(
    Name="FittedY",
    RefersTo=f"=OFFSET('{sname}'!$Y$2,1,0,'{sname}'!$M$8,1)",
)
```

This starts one row below the column header (row 2) and extends exactly `$M$8` rows — the number of filtered observations.

**Naming convention** — names identify the data, not the chart that uses them:

| Name | Column | Contents |
|---|---|---|
| `QQPlotX` | AE | Normal Scores Ranked (QQ theoretical axis) |
| `QQPlotY` | AF | Studentized Residuals Ranked (QQ actual axis) |
| `FittedY` | Y | Predicted Y — shared by multiple charts |
| `ResidData` | Z | Residuals |
| `ActualY` | X | Actual Y |
| `ScaleLocData` | AG | Scale-Location |
| `CooksDistData` | AD | Cook's Distance |
| `LeverageData` | AB | Hat Diagonal |
| `StudResidData` | AC | Studentized Residuals |
| `PRESSResidData` | AH | PRESS Residual |

**Scope:** all names are worksheet-scoped (created via `sheet.api.Names.Add`). Chart `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer:

```excel
Series X values: ='Regression'!FittedY
Series Y values: ='Regression'!ResidData
```

In code, use the `_name_ref` helper in `_write_diagnostic_charts`:

```python
def _name_ref(local_name: str) -> str:
    return f"='{sname}'!{local_name}"
```

When adding a new diagnostic column or chart, add the corresponding named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

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

Produces `Lambda_Library.xlsx` — the distributable artifact committed to the repo. Writes these sheets (in order):

- **LAMBDA_functions** — browsable catalog of all function definitions
- **Life Expectancy Data** — WHO dataset as a structured table
- **Regression Instructions** — user guide for the Regression sheet
- **Diagnostic Guide** — interpretation guide for regression diagnostics
- **Version History** — changelog
- **Univariate Analysis** — histogram tables, descriptive statistics, distribution fitting, Weibull grid-search MLE
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

Produces `Lambda_Library_QC.xlsx` (gitignored). Writes all production sheets plus `MLR_Scalar_Test`, `MLR_Vector_Outputs_Test`, and `MLR_Observation_Test`, updates `.analysis_cache.json`, and runs the expected-vs-actual verification pass.

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
  sheet_styles.py            # shared color constants (HEADER_COLOR, SUBHDR_COLOR, INPUT_COLOR)
  write_sheet_lambda_functions.py
  write_sheet_life_expectancy_data.py
  write_sheet_univariate.py
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

## Cell styling

All cell colors are defined once in `lambda_catalog/sheet_styles.py` and imported by every sheet writer. Never hard-code RGB tuples directly in a sheet writer.

| Constant | RGB | Usage |
|---|---|---|
| `HEADER_COLOR` | `(202, 237, 251)` | Zone / section headings — light blue, used on every sheet |
| `SUBHDR_COLOR` | `(220, 230, 241)` | Column sub-header rows within a zone |
| `INPUT_COLOR`  | `(251, 226, 213)` | User-editable input cells — light orange |

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

**Naming convention** — all chart-series named ranges carry the `RegChart` prefix. This distinguishes them from formula-helper names (`All_Xs`, `pred_input`, `alpha`, etc.) which use plain noun names:

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

When adding a new diagnostic column or chart, add the corresponding `RegChart*` named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

## Univariate sheet conventions

### Sheet layout overview

The Univariate Analysis sheet is written by `write_sheet_univariate.py`. It has five horizontal zones:

| Cols | Zone | Contents |
|---|---|---|
| A | Data input | User data column, rows 4–2003 (2 000 rows max) |
| C–D | Descriptive Statistics | 12 stat label/value rows |
| F–M | Histograms | Sturges (F–G), Scott (I–J), Freedman-Diaconis (L–M) |
| O–Z | Distribution Fitting | Per-distribution MLE params, NLL, AIC, BIC |
| AB–BR | Weibull Grid-Search | Two 20×20 stages (Stage 1: AB–AV; Stage 2: AX–BR) |

### Sheet-scoped named ranges

All named ranges on this sheet carry the `UV_` prefix:

| Name | Range | Purpose |
|---|---|---|
| `UV_Data` | `$A$4:$A$2003` | Input data column — referenced by all LAMBDA formulas |
| `UV_n` | scalar | `COUNT(UV_Data)` — used for AIC/BIC denominators |
| `UV_Sturges_Edges` | OFFSET-based | Chart series for Sturges histogram |
| `UV_Sturges_Counts` | OFFSET-based | Chart series for Sturges histogram |
| `UV_Scott_Edges` | OFFSET-based | Chart series for Scott histogram |
| `UV_Scott_Counts` | OFFSET-based | Chart series for Scott histogram |
| `UV_FD_Edges` | OFFSET-based | Chart series for Freedman-Diaconis histogram |
| `UV_FD_Counts` | OFFSET-based | Chart series for Freedman-Diaconis histogram |
| `UV_WB_S1` | `AC5:AV24` | Weibull Stage 1 NLL grid body |
| `UV_WB_S2` | `AY5:BR24` | Weibull Stage 2 NLL grid body |

### Weibull two-stage grid-search (Zone 5)

The Weibull fit uses a two-stage 20×20 grid-search MLE implemented entirely in Excel formulas. Stage 1 searches a coarse range; Stage 2 zooms in on the Stage 1 best estimate using a narrower range seeded from Stage 1's result. The distribution fitting table (Zone 4) reads the best shape and scale from Stage 2.

Each stage is written by `_write_grid_stage(sheet, row_start, col_start, title, ...)`. Both stages share the same compact 3-row header layout:

| Row offset | Contents |
|---|---|
| dr=0 (row 1) | Zone title (merged c0:c0+2, `_HEADER` fill); "Min NLL:" label at c0+3, MIN formula at c0+4 |
| dr=1 (row 2) | "shape (k) range:"; min/max bounds (editable, `_INPUT` fill); "Best shape:" label + INDEX formula; "Best scale:" label + INDEX formula |
| dr=2 (row 3) | "scale (λ) range:"; min/max bounds; row_offset formula; col_offset formula; dt_row_input placeholder; dt_col_input placeholder |
| dr=3 (row 4) | Corner cell: `=NLL_Weibull(UV_Data, dt_row_ref, dt_col_ref)`; param1 SEQUENCE spills right 20 cols |
| dr=4–23 (rows 5–24) | Param2 SEQUENCE in col c0; Data Table body in c0+1:c0+20 |

The auxiliary cells in dr=2 serve two purposes:

- **row_offset / col_offset** — `IFERROR(MIN(IF(...)), 1)` array formulas that locate the minimum within the grid body. They feed the "Best shape" and "Best scale" `INDEX` formulas in dr=1.
- **dt_row_input / dt_col_input** — plain numeric placeholders (`1.0`) that Excel substitutes during Data Table evaluation. They connect the corner NLL formula to the parameter SEQUENCE headers.

#### Excel Data Table setup

The two-input Data Table is created by calling `.api.Table(RowInput, ColumnInput)` on the full range from the corner cell through the last body cell. The corner cell contains the NLL formula; the param1 SEQUENCE is the column header (spilling right); the param2 SEQUENCE is the row header (spilling down). Excel substitutes `dt_row_input` with each param1 value and `dt_col_input` with each param2 value to populate the body.

**Important:** `_drop_local_name(sheet, body_name)` must be called and the named range for the body (`UV_WB_S1` or `UV_WB_S2`) must be registered via `sheet.api.Names.Add` before calling `.api.Table()`. The Data Table write must happen before the Min NLL formula that references the body name.

#### Layout constants

All row and column positions within a stage are expressed as `_GS_R_*` (row offset from `row_start`) and `_GS_C_*` (column offset from `col_start`) constants. Never hard-code positions inside `_write_grid_stage`.

| Constant | Value | Meaning |
|---|---|---|
| `_ROW_GS_WB` | 1 | Row anchor for both grid stages (`row_start`) |
| `_GS_R_P1_BND` | 1 | Shape bounds row offset |
| `_GS_R_P2_BND` | 2 | Scale bounds / auxiliary row offset |
| `_GS_R_HDR` | 3 | Grid header row offset (corner + param1 SEQUENCE) |
| `_GS_R_BODY` | 4 | First grid body row offset |
| `_GS_C_MINNLL_LBL` | 3 | "Min NLL:" label col offset |
| `_GS_C_MINNLL_VAL` | 4 | Min NLL value col offset |
| `_GS_C_BEST_P1_LBL` | 3 | "Best shape:" label col offset |
| `_GS_C_BEST_P1_VAL` | 4 | Best shape value col offset |
| `_GS_C_BEST_P2_LBL` | 5 | "Best scale:" label col offset |
| `_GS_C_BEST_P2_VAL` | 6 | Best scale value col offset |
| `_GS_C_ROW_OFF` | 3 | row_offset formula col offset |
| `_GS_C_COL_OFF` | 4 | col_offset formula col offset |
| `_GS_C_DT_ROW` | 5 | dt_row_input placeholder col offset |
| `_GS_C_DT_COL` | 6 | dt_col_input placeholder col offset |

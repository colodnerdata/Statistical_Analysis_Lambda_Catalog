# Contributing

## Setup

Requires Python 3.10+, [uv](https://github.com/astral-sh/uv). Building the Excel workbook also requires desktop Excel on Windows or Mac (xlwings uses COM automation on Windows, AppleScript bridges on Mac), but running the Python test suite does not.

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
| `test_model_construction_writer.py` | Model Construction sheet writer: sheet-scoped name definitions and order, T0 default-spec prefill, dropdowns, conditional formats, `X_s`/`Constructed_Column_Names` twin invariants |
| `test_analyze_model_construction.py` | Model Construction QC analyzer: default-spec expectations pinned against the sample CSV (mask size, k, level-qualified names), the stratified-Filter degeneracy case, and the observed-vs-expected comparison layer |
| `test_weibull_grid_excel.py` | Weibull grid-search mechanics validation |
| `test_inspection_compare.py` | QC value comparison logic (`to_float_or_none`, `first_digit_deviation`, `compare_values`) |
| `test_independent_verification.py` | Independent numpy/scipy verification of all LAMBDA function outputs (scalars, vectors, observation diagnostics, predictor summary, prediction interval) |
| `test_qc_configs.py` | QC config generation, cross-consistency between scalar/vector/observation configs, regression sheet diagnostics, cache round-trips |
| `test_bfn_panel_durbin_watson_verification.py` | `BFN_Panel_Durbin_Watson` against the WHO panel — within-group differencing via `Difference_By`, mutual gating with `Durbin_Watson_By` |
| `test_serial_correlation_group_resolver.py` | `Serial_Correlation_Group()` SWITCH, including the dormant Cluster branch (the v2.6+ reserved-spec-column pattern) |
| `test_difference_by_verification.py` | Gap-aware `Difference_By` (WHO exact counts plus the punched-out-year and calendar-date synthetic cases per `HUMAN_TEST_PLAN_v3_model_construction.md` T17–T19) |
| `test_analyze_regression_spec_block.py` | Post-changeover spec-block QC analyzer (predicted counts and values, regression sheet spec state) |

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

Produces `Lambda_Library.xlsx` — the distributable artifact committed to the repo. Writes eight sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions
- **Life Expectancy Data** — WHO dataset as a structured table
- **Univariate Analysis** — descriptive statistics, histogram binning, and Weibull grid-search fitting
- **Regression Instructions** — step-by-step guide for adapting the sheet to new datasets
- **Diagnostic Guide** — interpretation guide for regression diagnostics
- **Version History** — changelog that travels with the workbook
- **Regression** — ToolPak-style analysis interface
- **Model Construction** — declarative variable-specification block and the sheet-scoped names that assemble the design matrix from it. The wiring names (`Source_Data`, `Header_Names`, `Spec_*`) hardcode this sheet's cell addresses and are defined in `write_sheet_model_construction.py`; the constructor closures (`Sample_Include`, `Response_Column`, `Row_Labels`, `X_s`, `Constructed_Column_Names`) live in `lambda_functions.json` with `"scope": "Model Construction"`, so they are the single source of truth and appear on the LAMBDA_functions catalog sheet (Scope column) like any other function — they are just installed on this sheet rather than workbook-wide

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

Produces `Lambda_Library_QC.xlsx` (gitignored). Writes all twelve sheets (the eight above plus `MLR_Scalar_Test`, `MLR_Vector_Outputs_Test`, `MLR_Observation_Test`, `Dummy_Test`), updates `.analysis_cache.json`, and runs the expected-vs-actual verification pass.

The `Dummy_Test` sheet is self-checking: every case is a boolean Pass formula (e.g. `=ISNA(Dummy_Levels(...))`) evaluated by Excel, and the verification pass reads the Pass cells back and reports any that are not TRUE.

The verification step forces Excel to recalculate all test formulas, reads the Calc columns, compares them against Python-computed expected values, and prints a `WARNING` line for any value that diverges beyond the tolerance band. No warnings means the LAMBDA implementations agree with statsmodels.

Optional flags:

```powershell
python build_qc.py --validate-reopen
python build_qc.py --verbose
python build_qc.py --cache path/to/.analysis_cache.json   # non-default cache location
python build_qc.py --no-verify                           # build QC sheets but skip the spec-driven pass
```

Run the QC build whenever you add or modify a LAMBDA function.

## Verifying builds

There are two verifier layers with different speeds and different scopes. Run them in this order when in doubt; either can be skipped if the other has been run recently.

### Layer 1 — headless structural check

Pure `zipfile` + `lxml` reads of the produced `.xlsx`. Runs in <1 s on Linux CI without Excel. Catches packaging regressions the unit-test suite misses: dangling defined names, `#REF!`/`#NAME?` cached-value literals, broken `[Content_Types].xml`/`workbook.xml.rels`, orphan chart-relationship targets, `localSheetId` out of range, sheet drift.

```powershell
# All-in-one (also runs the opt-in real-workbook tests):
make verify-headless

# Or directly:
uv run pytest tests/test_workbook_invariants.py -v
RUN_EXCEL_INTEGRATION=1 uv run pytest tests/test_workbook_invariants.py -v   # + real-workbook tests
```

This is a fast screen. A green run does **not** mean the workbook calculates correctly — that is what Layer 2 is for.

### Layer 2 — spec-driven deep check

Reuses `build_qc.verify_test_sheets` against the production sheets. Same machinery the QC build runs against `MLR_*_Test` and `Dummy_Test`, gated off the `Dummy_Test` block via `skip_dummy=True` because production workbooks do not contain a `Dummy_Test` sheet. This is the source of truth for cell-level correctness.

```powershell
# Run the production build, recalculate, then verify against the spec oracle.
# On drift: print a structured VerifyReport, sys.exit(1), and do NOT open
# Excel (so a stale build cannot be launched in place of a fresh one).
python build_production.py --verify --no-launch --skip-data-table-calculations

# Or, on the just-built workbook without rebuilding:
python tools/verify_workbook.py Lambda_Library.xlsx
python tools/verify_workbook.py Lambda_Library.xlsx --json   # agentic consumption
```

The `VerifyReport` (in `lambda_catalog/verify_report.py`) is emitted both in a human-readable form (`Verify: passed (spec mode, …)` / `ERROR Verify mismatch totals: …`) and as JSON (stable schema: `passed`, `categories`, `failures`, `elapsed_seconds`, `mode`, `workbook`).

### `make verify`

```powershell
make verify             # both layers (Layer 1 on any platform, Layer 2 needs Excel)
make verify-headless    # Layer 1 only
make verify-deep        # Layer 2 only
```

`make verify-deep` shells out to `build_production.py --verify --no-launch --skip-data-table-calculations`, so it both rebuilds and verifies. To verify an already-built workbook, use `python tools/verify_workbook.py Lambda_Library.xlsx` instead.

### CI

GitHub Actions runs the unit-test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request via `.github/workflows/ci.yml`. The spec-driven verifier (Layer 2) is **not** run in CI: the GitHub-hosted `windows-latest` runner image does not include Microsoft Office, so xlwings fails to dispatch `Excel.Application` (`pywintypes.com_error: (-2147221005, 'Invalid class string')`). Until a self-hosted runner with Office is wired in, Layer 2 must be run on a developer machine (or any Windows box with desktop Excel) — the agentic workflow runs it before pushing. The `windows-verify` job was removed for that reason; see the comment block at the bottom of `ci.yml`. Layer 1 (the headless `tests/test_workbook_invariants.py` suite) is auto-discovered by the existing Linux job once it lands.

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
  write_sheet_model_construction.py
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
  verify_workbook.py         # standalone CLI wrapping the spec-driven verifier
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

All cell colors are defined once in `lambda_catalog/sheet_styles.py` and imported by every sheet writer. Never hard-code RGB tuples directly in a sheet writer. The constant table, the import pattern, and the section-heading convention are in [CLAUDE.md](CLAUDE.md) / [AGENTS.md](AGENTS.md) at the project-instructions tier — consult that for the canonical definitions. Sheet-specific colors that differ from the shared palette (e.g., `_SUBHEADER_COLOR` in `write_sheet_diagnostic_guide.py`) remain as local constants in the relevant file.

## Regression sheet conventions

### Chart series data ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns (`$AI$3:$AI$1048576`) degrades Excel's recalculation performance and can crash the workbook on large datasets.

Instead, all chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in `$W$8` (the `Observations` cell in the Regression Outputs zone):

```python
sheet.api.Names.Add(
    Name="RegChartFitY",
    RefersTo=f"=OFFSET('{sname}'!$AK$2,1,0,MAX(IFERROR('{sname}'!$W$8,1),1),1)",
)
```

This starts one row below the column header (row 2) and extends exactly `$W$8` rows — the number of filtered observations. The `MAX(IFERROR(...,1),1)` guard keeps the range one row tall (instead of erroring) when `$W$8` cannot resolve. Each name also carries a Name Manager `Comment` identifying the chart it feeds — see the loop in `_setup_local_names`.

**Naming convention** — all OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix, distinguishing them from the constructor closures (`X_s`, `Sample_Include`, etc.) and formula-helper names:

| Name | Column | Contents |
|---|---|---|
| `RegChartQQX` | AP | Normal Scores Ranked (QQ theoretical axis) |
| `RegChartQQY` | AQ | Studentized Residuals Ranked (QQ actual axis) |
| `RegChartFitY` | AJ | Predicted Y — shared by multiple charts |
| `RegChartResid` | AK | Residuals |
| `RegChartActY` | AI | Actual Y |
| `RegChartScaleLoc` | AR | Scale-Location |
| `RegChartCookDist` | AO | Cook's Distance |
| `RegChartLeverage` | AM | Hat Diagonal |
| `RegChartStudResid` | AN | Studentized Residuals |
| `RegChartPRESSResid` | AS | PRESS Residual |

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

# Contributing

## Setup

Requires Python 3.10+, [uv](https://github.com/astral-sh/uv). Building the Excel workbook also requires desktop Excel on Windows or Mac (xlwings uses COM automation on Windows, AppleScript bridges on Mac), but running the Python test suite does not.

```powershell
uv sync
```

This installs the `lambda_catalog` package in editable mode along with all dependencies: `lxml`, `numpy`, `pandas`, `pywin32` (Windows only), `scipy`, `statsmodels`, `xlwings`, plus dev tools (`pytest`, `pytest-cov`, `pylint`).

## Quick start

The most common tasks, copy-pasteable. Commands marked **(needs Excel)** dispatch xlwings COM automation and only work on Windows or Mac with desktop Excel installed; everything else runs on any platform, including Linux CI.

```powershell
# One-time: install the package and all dependencies
uv sync

# Run the full test suite (no Excel needed)
uv run pytest

# Run tests with a coverage report
uv run pytest --cov --cov-report=term-missing

# Fast headless structural check of the committed workbook (no Excel needed)
uv run pytest tests/test_workbook_invariants.py -v      # or: make verify-headless

# Build the distributable workbook, Lambda_Library.xlsx  (needs Excel)
uv run python build_production.py

# Build + verify with no Excel window popping up — the one-shot CI-style flow  (needs Excel)
uv run python build_production.py --verify --no-launch --skip-data-table-calculations

# Build the QC workbook and run the full expected-vs-actual pass
# (do this whenever you add or change a LAMBDA function)  (needs Excel)
uv run python build_qc.py

# Verify an already-built workbook without rebuilding it  (needs Excel)
uv run python tools/verify_workbook.py Lambda_Library.xlsx
```

New to the repo? A typical loop is: edit code → `uv run pytest` → `uv run python build_qc.py` (to confirm the workbook still calculates) → `uv run python build_production.py` (to regenerate the committed `Lambda_Library.xlsx`). The full flag reference for each script is under [Building](#building) and [Verifying builds](#verifying-builds) below.

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
| `test_difference_by_verification.py` | Gap-aware `Difference_By` (WHO exact counts plus the punched-out-year and calendar-date synthetic cases per `HUMAN_TEST_PLAN_v20_model_construction.md` T17–T19) |
| `test_analyze_regression_spec_block.py` | Post-changeover spec-block QC analyzer (predicted counts and values, regression sheet spec state) |
| `test_regression_spec_qc.py` | Spec-driven Regression QC oracle (`analyze_regression_spec.py` case definitions) |
| `test_csv_dataset_loader.py` | `load_csv_rows` (`write_sheet_csv_dataset.py`) against all three `CsvDatasetConfig`s and the committed sample CSVs |
| `test_mileage_completeness_qc.py` | `calculate_mileage_completeness_flags` (`analyze_mileage.py`) against the Auto MPG dataset |
| `test_within_estimator.py` | v2.1 Fixed Effects phase 2 — the fit-time demeaned pair `y_s()`/`X_s_Within()`, against an independent `statsmodels` LSDV fit |
| `test_group_panel_transforms.py` | v2.1 Fixed Effects phase 1 — `Group_Mean`, `Demean_By`, `Is_Balanced_Panel`, `Absorbed_Degrees_Of_Freedom` |
| `test_df_absorbed_threading.py` | v2.1 Fixed Effects phase 3 — `[DF_Absorbed]` threaded through SE/t/p/CI/MS-Residual/AIC/BIC/AICc, against an independent `statsmodels` LSDV fit |
| `test_group_prediction_interval.py` | v2.1 Fixed Effects phase 5 — `Group_Mean_At`, `Group_Count_At`, `Prediction_Group_Column`, `Group_Prediction_Interval` (the group-mean-recovery CI+PI form), against an explicit LSDV `get_prediction()` reference |
| `test_workbook_helpers.py` | `safe_activate()` / `safe_freeze_top_row()` against stub sheet/window objects (headless/no-focus Excel session guards) |
| `test_workbook_builder.py` | Workbook package-patching helpers (`sync_workbook_names` and friends) that don't require Excel |
| `test_build_production.py` | `build_production.py`'s pure-Python logic (CLI flag handling, dataset selection, tab order/color assignment) that doesn't require Excel |
| `test_workbook_invariants.py` | Layer 1 headless structural check of a built `.xlsx` package (`zipfile` + `lxml`): dangling defined names, `#REF!`/`#NAME?` cached-value literals, broken package parts, orphan chart-relationship targets, sheet drift — see [Verifying builds](#verifying-builds) |
| `test_ln_positive_verification.py` | v2.2 Transform=Log — `Ln_Positive` pure-Python mirror (the `NA()`-exception contract, the geometric-mean round-trip the Prediction Inputs fix relies on) and implementation-shape assertions on the catalog formula |
| `test_transform_threading.py` | v2.2 Transform=Log wiring end to end — cross-checks the new `production_lots_log_transform` QC case (raw columns, `transform="Log"`) against the pre-existing precomputed-log-column case to floating-point precision; Categorical×Log inertness |

### Coverage scope

The coverage configuration in `pyproject.toml` tracks only the modules that are testable without Excel:

- `analyze_life_expectancy.py`
- `analyze_mileage.py`
- `analyze_production_lots.py`
- `analyze_model_construction.py`
- `analyze_regression_spec.py`
- `analyze_regression_spec_block.py`
- `analyze_univariate.py`
- `catalog_schema.py`
- `lambda_formula_parser.py`
- `regression_shared.py`
- `analysis_cache.py`
- `verify_report.py`

The `write_sheet_*.py` modules, `workbook_builder.py`, `workbook_helpers.py`, `make_test_sheet.py`, `sheet_styles.py`, `inspection_compare.py`, `analyze_regression_sheet.py`, and other xlwings-dependent modules are omitted from CI coverage measurement. They are validated by the QC build instead (see below).

### CI

GitHub Actions runs the test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request. See `.github/workflows/ci.yml`. Coverage must stay at or above 70% on the tracked modules.

## Building

There are two separate build scripts with distinct purposes.

### Production build

```powershell
uv run python build_production.py
```

Produces `Lambda_Library.xlsx` — the distributable artifact committed to the repo. Writes nine sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions
- **Life Expectancy Data** — WHO dataset as a structured table (a second sample dataset for practicing the Source_Table retarget workflow)
- **Mileage Data** — Auto MPG dataset as a structured table; this is the dataset the Regression sheet's `Source_Table` targets by default
- **Production Lots** — a small unbalanced learning-curve panel (3 facilities, 51 lots) as a structured table; a third sample dataset, and the only one with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year)
- **Univariate Analysis** — descriptive statistics, histogram binning, and Weibull grid-search fitting
- **Regression Instructions** — step-by-step guide for adapting the sheet to new datasets
- **Diagnostic Guide** — interpretation guide for regression diagnostics
- **Version History** — changelog that travels with the workbook
- **Regression** — ToolPak-style analysis interface driven by a declarative variable-specification block (the spec block) and the sheet-scoped names that assemble the design matrix from it. The wiring names (`Source_Data`, `Header_Names`, `Spec_*`) hardcode the spec block's cell addresses and are defined in `write_sheet_model_construction.py` (imported by `write_sheet_regression.py`); the constructor closures (`Sample_Include`, `Response_Column`, `Row_Labels`, `X_s`, `Constructed_Column_Names`) live in `lambda_functions.json` with `"scope": "Regression"`, so they are the single source of truth and appear on the LAMBDA_functions catalog sheet (Scope column) like any other function — they are just installed on this sheet rather than workbook-wide

No test sheets, no OLS analysis, no cache dependency.

**All `build_production.py` options:**

| Flag | Default | What it does |
|---|---|---|
| `--workbook PATH` | `Lambda_Library.xlsx` | Path to the workbook to create or update. |
| `--definitions PATH` | `lambda_functions.json` | Path to the JSON catalog of LAMBDA definitions. |
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV written to the **Life Expectancy Data** sheet. (The **Mileage Data** and **Production Lots** sources are fixed committed sample files with no CLI override.) |
| `--regression-dataset {auto_mpg,life_expectancy,production_lots}` | `auto_mpg` | Which dataset the Regression sheet's `Source_Table` targets. `production_lots` is the one to pick for a ready-made Fixed Effects example (Facility as the FE role, Fiscal_Year as Sequence). |
| `--skip-univariate` | off | Skip writing the Univariate Analysis sheet to speed up iteration on other sheets. An existing Univariate sheet is left as-is; a from-scratch build omits it. |
| `--skip-data-table-calculations` | off | Skip the final Excel `CalculateFullRebuild` phase that evaluates Data Tables. The workbook is still written and names synced; formulas/Data Tables recalc later when opened in Excel. Big speedup for iteration. |
| `--verify` | off | After the build, run the spec-driven verifier (`build_qc.verify_test_sheets`) against the production sheets. On any drift, print a structured `VerifyReport` and `sys.exit(1)`. The Excel handoff only fires when verify passes, so a stale build can't launch in place of a fresh one. |
| `--no-verify` | (default) | Explicitly disable the verifier pass. Mainly for wrapper scripts that default to `--verify`. |
| `--no-launch` | off | Suppress the post-build `cmd /c start <workbook>` Excel handoff. Use in agentic/automated loops where no Excel window should pop up. |
| `--validate-reopen` | off | Reopen the workbook after syncing names to confirm Excel accepts the result. |
| `--verbose` | off | Print per-phase timing checkpoints to stdout. |

Common combinations:

```powershell
# Plain build → opens Lambda_Library.xlsx in Excel when done
uv run python build_production.py

# The one-shot automated flow: build, sync names, run the spec-driven verifier,
# exit non-zero on drift, and never open Excel. This is what `make verify-deep` runs.
uv run python build_production.py --verify --no-launch --skip-data-table-calculations

# Fast Regression-only iteration (skip the slow Univariate sheet and Data Table rebuild)
uv run python build_production.py --skip-univariate --skip-data-table-calculations --no-launch
```

### QC build

```powershell
uv run python build_qc.py
```

Produces `Lambda_Library_QC.xlsx` (gitignored). Writes all thirteen sheets (the nine above plus `MLR_Scalar_Test`, `MLR_Vector_Outputs_Test`, `MLR_Observation_Test`, `Dummy_Test`), updates `.analysis_cache.json`, and runs the expected-vs-actual verification pass.

The `Dummy_Test` sheet is self-checking: every case is a boolean Pass formula (e.g. `=ISNA(Dummy_Levels(...))`) evaluated by Excel, and the verification pass reads the Pass cells back and reports any that are not TRUE.

The verification step forces Excel to recalculate all required sheets, reads the Calc columns, compares them against Python-computed expected values, and emits `ERROR ...` lines plus a mismatch summary when drift is found. A clean run produces no mismatch errors.

**All `build_qc.py` options:**

| Flag | Default | What it does |
|---|---|---|
| `--workbook PATH` | `Lambda_Library_QC.xlsx` | Path to the QC workbook to create or update. |
| `--definitions PATH` | `lambda_functions.json` | Path to the JSON catalog of LAMBDA definitions. |
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV used for both the data sheet and the `Full_Data` QC comparison. |
| `--cache PATH` | `.analysis_cache.json` | Retained for compatibility; spec-driven QC computes on demand, so this rarely matters. |
| `--no-verify` | off | Skip the spec-driven verify pass. Escape hatch for iterating on a known-broken sheet; the skip is logged to `qc_log.txt` so the absence is visible. |
| `--validate-reopen` | off | Reopen the workbook after syncing names to confirm Excel accepts the result. |
| `--verbose` | off | Print per-phase timing checkpoints to stdout. |

```powershell
# Full QC build + verification (the usual invocation)
uv run python build_qc.py

# Build the QC sheets but skip the spec-driven pass (iterating on a known-broken sheet)
uv run python build_qc.py --no-verify
```

`build_qc.py` mirrors terminal output to `qc_log.txt` and includes end-of-run timing lines (`Timing: prep`, `write sheets`, `sync names`, `verify`, `total`) in both places.

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
python build_production.py --verify --no-launch --skip-data-table-calculations --skip-univariate

# Or, on the just-built workbook without rebuilding:
uv run python tools/verify_workbook.py Lambda_Library.xlsx
uv run python tools/verify_workbook.py Lambda_Library.xlsx --json   # agentic consumption
```

**All `tools/verify_workbook.py` options** (positional `workbook` path is required):

| Flag | Default | What it does |
|---|---|---|
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV used for the `Full_Data` comparison. |
| `--json` | off | Emit the report as JSON (stable schema, for agentic consumption) instead of the human-readable form. |
| `--verbose` | off | Print per-phase checkpoints from the spec-driven verifier. |

The `VerifyReport` (in `lambda_catalog/verify_report.py`) is emitted both in a human-readable form (`Verify: passed (spec mode, …)` / `ERROR Verify mismatch totals: …`) and as JSON (stable schema: `passed`, `categories`, `failures`, `elapsed_seconds`, `mode`, `workbook`).

Expected terminal flow for the one-shot command:

1. Sheet update summary (`Sheet updated: ...`, `Created names: ...`, `Updated names: ...`).
2. Spec verifier result (`Verify: passed ...` or `ERROR Verify mismatch totals: ...`).
3. Timing summary lines (`Timing: build+sync`, `Timing: recalculate` or `skipped`, `Timing: verify`, `Timing: total`).

When the production verifier runs, the deep-check pre-calc list excludes `Dummy_Test` (`skip_dummy=True`) because production workbooks do not include that sheet.

### `make verify`

```powershell
make verify             # both layers (Layer 1 on any platform, Layer 2 needs Excel)
make verify-headless    # Layer 1 only
make verify-deep        # Layer 2 only
```

`make verify-deep` shells out to `build_production.py --verify --no-launch --skip-data-table-calculations --skip-univariate`, so it both rebuilds and verifies while leaving the existing Univariate sheet untouched for faster regression-focused loops. To verify an already-built workbook, use `python tools/verify_workbook.py Lambda_Library.xlsx` instead.

### CI

GitHub Actions runs the unit-test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request via `.github/workflows/ci.yml`. The spec-driven verifier (Layer 2) is **not** run in CI: the GitHub-hosted `windows-latest` runner image does not include Microsoft Office, so xlwings fails to dispatch `Excel.Application` (`pywintypes.com_error: (-2147221005, 'Invalid class string')`). Until a self-hosted runner with Office is wired in, Layer 2 must be run on a developer machine (or any Windows box with desktop Excel) — the agentic workflow runs it before pushing. The `windows-verify` job was removed for that reason; see the comment block at the bottom of `ci.yml`. Layer 1 (the headless `tests/test_workbook_invariants.py` suite) is auto-discovered by the existing Linux job once it lands.

## File structure

```
build_production.py          # production entry point → Lambda_Library.xlsx
build_qc.py                  # QC entry point → Lambda_Library_QC.xlsx
rebuild_static_sheets.py      # regenerates templates/static_sheets.xlsx from its Python source —
                              # see "Static reference sheets" below
lambda_functions.json         # LAMBDA definitions (source of truth)
sample_data/
  Life Expectancy Data.csv   # WHO life expectancy dataset
  auto_mpg_data.csv          # Auto MPG dataset (second sample dataset, "Mileage Data" sheet)
  production_lots.csv        # Learning-curve panel (third sample dataset, "Production Lots" sheet) —
                              # the only shipped dataset with a natural Fixed Effects grouping column
  auto_mpg_data.xlsx         # retired source file for auto_mpg_data.csv; kept for reference, unused by writers
  production_lots.xlsx       # retired source file for production_lots.csv; kept for reference, unused by writers
templates/
  static_sheets.xlsx         # pre-built copies of dataset-independent reference sheets
                              # (Regression Instructions, Diagnostic Guide) — see
                              # "Static reference sheets" below
lambda_catalog/
  catalog_schema.py          # typed document model: CatalogArgument, CatalogFunction, CatalogDocument
  regression_shared.py       # shared regression dataclasses: RegressionSummary, RegressionVectors, etc.
  sheet_styles.py            # shared cell-formatting constants (colors, conditional formatting)
  workbook_builder.py        # shared core: sync_workbook_names, workbook XML patching
  workbook_helpers.py        # shared xlwings utilities and cell formatting helpers
  analyze_life_expectancy.py # OLS engine: calculate_regression_summary, vectors, observations
  analyze_mileage.py         # Auto MPG QC oracle: calculate_mileage_completeness_flags
  analyze_production_lots.py # Production Lots QC oracle: calculate_production_lots_completeness_flags
  analyze_regression_spec.py # spec-driven Regression QC cases, incl. the Fixed Effects case
  analyze_regression_spec_block.py # post-changeover spec-block QC analyzer (predicted counts/values, spec state)
  analyze_regression_sheet.py # full Regression sheet QC oracle (predictor summary, residuals, prediction interval,
                              # and the Fixed Effects within-transform/DF_Absorbed correction)
  analyze_model_construction.py # Model Construction QC analyzer: default-spec expectations, mask/level checks
  analyze_univariate.py      # univariate analysis: NLL functions, MLE estimators, binning, GoF
  analysis_cache.py          # disk cache keyed on CSV SHA-256 + schema version
  lambda_formula_parser.py   # converts display formulas to workbook XML syntax
  inspection_compare.py      # numeric comparison helpers for QC value verification
  verify_report.py           # VerifyReport: structured pass/fail result for the spec-driven verifier
  make_test_sheet.py         # shared helpers for Excel ListObject test tables
  write_sheet_lambda_functions.py
  write_sheet_csv_dataset.py # unified loader/writer/CLI for Life Expectancy, Mileage, and Production Lots
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
python -m lambda_catalog.write_sheet_csv_dataset life_expectancy Lambda_Library.xlsx
python -m lambda_catalog.write_sheet_csv_dataset mileage Lambda_Library.xlsx
python -m lambda_catalog.write_sheet_csv_dataset production_lots Lambda_Library.xlsx
```

`write_sheet_csv_dataset.py` is a single config-driven module backing all three sample datasets (Life Expectancy, Mileage, Production Lots) — one `CsvDatasetConfig` per dataset (`LIFE_EXPECTANCY`, `MILEAGE`, `PRODUCTION_LOTS`) captures its sheet/table name, `Full_Data` formula, and CSV-parsing quirks (missing-value markers, header normalization), and the shared `load_csv_rows` / `write_csv_dataset_sheet` functions do the actual work. All three CSVs are real committed sample files (`sample_data/Life Expectancy Data.csv`, `sample_data/auto_mpg_data.csv`, `sample_data/production_lots.csv`) that can be pointed at a different CSV via `--csv`, or via `build_production.py`'s `--csv` / `--mileage-csv` / `--production-lots-csv` flags. The loader (`load_csv_rows`) has no Excel dependency, so the Python QC oracles (`analyze_mileage.calculate_mileage_completeness_flags`, `analyze_production_lots.calculate_production_lots_completeness_flags`) can run on any platform.

### Static reference sheets

`write_sheet_regression_instructions.py` and `write_sheet_diagnostic_guide.py` write sheets whose content never depends on the target dataset — a fixed how-to guide and a fixed diagnostics reference. Rebuilding hundreds of styled cells with COM calls for unchanging text on every production/QC build is wasted work, so these two modules instead copy an already-styled sheet out of `templates/static_sheets.xlsx` via `workbook_helpers.copy_static_sheet` (Excel's native `Sheet.Copy` between two workbooks open in the same Excel instance — not an openpyxl round-trip; see CLAUDE.md's "Use xlwings COM API for all chart creation — never openpyxl" for why openpyxl is unsafe for anything Excel-native like this). `write_regression_instructions_sheet(workbook)` / `write_diagnostic_guide_sheet(workbook)` keep their original call signature, so `build_production.py` / `build_qc.py` and their tests are unaffected.

The authored content still lives in Python — `_ROWS` in `write_sheet_regression_instructions.py`, the body of `_write_template_sheet` in `write_sheet_diagnostic_guide.py` — but neither `build_production.py` nor `build_qc.py` ever executes it; they only call the copy-from-template functions above. Regenerating the template is a separate, manual step.

Run **`python rebuild_static_sheets.py`** after editing either sheet's content, then commit the updated `templates/static_sheets.xlsx` alongside the Python change. It opens the template once, calls every static sheet's `_write_template_sheet(workbook)` (so nothing is skipped or forgotten), and saves once. This is the standard command — prefer it over the per-module CLIs below, which exist only for regenerating a single sheet in isolation while debugging:

```powershell
python rebuild_static_sheets.py                          # regenerates every static sheet (standard)
python -m lambda_catalog.write_sheet_regression_instructions  # single-sheet debugging only
python -m lambda_catalog.write_sheet_diagnostic_guide         # single-sheet debugging only
```

**Why a dedicated command exists:** before it was added, each sheet's own CLI was the only way to regenerate the template, so editing `_ROWS` or `_write_template_sheet` and forgetting to also run that specific CLI would silently ship stale reference text — the template drifted from its Python source with no error anywhere in the build. That happened at least twice (see DECISIONS.md → "Static template drift"). `rebuild_static_sheets.py` collapses "which CLI do I need to remember to run" into a single always-correct command.

All of this — the per-module CLIs and `rebuild_static_sheets.py` alike — requires a real Excel COM engine (`xlwings.App`); none of it runs in a headless/CI environment.

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

Instead, all chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in `$Y$8` (the `Observations` cell in the Regression Outputs zone):

```python
sheet.api.Names.Add(
    Name="RegChartFitY",
    RefersTo=f"=OFFSET('{sname}'!$AM$2,1,0,MAX(IFERROR('{sname}'!$Y$8,1),1),1)",
)
```

This starts one row below the column header (row 2) and extends exactly `$Y$8` rows — the number of filtered observations. The `MAX(IFERROR(...,1),1)` guard keeps the range one row tall (instead of erroring) when `$Y$8` cannot resolve. Each name also carries a Name Manager `Comment` identifying the chart it feeds — see the loop in `_setup_local_names`.

**Naming convention** — all OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix, distinguishing them from the constructor closures (`X_s`, `Sample_Include`, etc.) and formula-helper names:

| Name | Column | Contents |
|---|---|---|
| `RegChartQQX` | AR | Normal Scores Ranked (QQ theoretical axis) |
| `RegChartQQY` | AS | Studentized Residuals Ranked (QQ actual axis) |
| `RegChartFitY` | AM | Predicted Y — shared by multiple charts |
| `RegChartResid` | AN | Residuals |
| `RegChartActY` | AL | Actual Y |
| `RegChartScaleLoc` | AT | Scale-Location |
| `RegChartCookDist` | AQ | Cook's Distance |
| `RegChartLeverage` | AO | Hat Diagonal |
| `RegChartStudResid` | AP | Studentized Residuals |
| `RegChartPRESSResid` | AU | PRESS Residual — the leave-one-out (LOOCV) residual; there is no separate "LOOCV Residual" column |
| `RegChartCookDistFlag` | AV | Cook's Distance, `NA()`'d below both influence cutoffs (`D > 4/n` or `D > 0.9`) — feeds the Cook's Distance chart's data-label overlay series |
| `RegChartObsLabel` | AK | Row identifier (Row_Labels()) — the flagged-point overlay series' `XValues`, so its data labels read as the observation identifier rather than a bare index |

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

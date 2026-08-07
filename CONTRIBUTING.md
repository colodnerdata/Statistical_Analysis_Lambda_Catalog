# Contributing

## Setup

Requires Python 3.10+, [uv](https://github.com/astral-sh/uv). Building the Excel workbook also requires desktop Excel on Windows or Mac (xlwings uses COM automation on Windows, AppleScript bridges on Mac), but running the Python test suite does not.

```powershell
uv sync
```

This installs the `lambda_catalog` package in editable mode along with all dependencies: `lxml`, `numpy`, `pandas`, `pywin32` (Windows only), `scipy`, `statsmodels`, `xlwings`, plus dev tools (`pytest`, `pytest-cov`, `pylint`, `poethepoet`).

### Task shorthands (`poe`)

Every routine command in this file also exists as a [poethepoet](https://poethepoet.natn.io/) task defined in `[tool.poe.tasks]` in `pyproject.toml`. `poethepoet` is a standalone task runner, not a Poetry plugin or a Poetry-specific workflow. It is included in this project's default uv development environment (`[dependency-groups].dev` plus `[tool.uv].default-groups`), so `uv sync` installs the `poe` executable into `.venv` by default.

```powershell
uv sync                         # one-time, installs poe into the project .venv
.venv\Scripts\Activate.ps1     # Windows PowerShell
# source .venv/bin/activate     # macOS/Linux shells
poe verify-headless
```

Bare `poe <task>` is the default form in this document. It uses the repository's uv-managed environment after you activate `.venv` (or if your shell otherwise puts `.venv` on `PATH`). If you choose not to activate the environment, use `uv run poe <task>` as the equivalent explicit form. A global `uv tool install poethepoet` is optional and only useful when you want `poe` available outside this repository.

`[tool.poe.env]` sets `UV_FROZEN=1`, so tasks fail on a stale lockfile exactly as the old `--frozen` flag did.

Both forms are shown throughout this document: the task where one exists, and the underlying command, which always remains runnable directly.

#### Common poe shortcuts

Use these shortcuts for day-to-day work from an activated `.venv`. If the environment is not activated, prefix the task with `uv run`, for example `uv run poe verify-headless`.

| Goal | poe task | Equivalent direct command | Notes |
|---|---|---|---|
| Run the full unit-test suite | `poe test` | `uv run pytest` | No Excel required; this is the default local test pass. |
| Run tests with coverage | `poe test-cov` | `uv run pytest --cov --cov-report=term-missing --cov-report=xml` | Matches the CI coverage step. |
| Run the fast workbook invariant screen | `poe verify-headless` | `uv run pytest tests/test_workbook_invariants.py -v` | No Excel required; catches packaging/name/cache drift. |
| Run pylint's CI check | `poe lint` | `uv run pylint --errors-only lambda_catalog scripts tools` | Error-only lint, matching CI. |
| Run both CI checks at once | `poe check` | — | `test-cov` and `lint` in parallel; the local pre-push screen. |
| Build the Regression artifact | `poe build` | `uv run python scripts/build_production.py` | Needs desktop Excel. |
| Build the Univariate artifact | `poe build-univariate` | `uv run python scripts/build_univariate.py` | Needs desktop Excel. |
| Build + verify Regression | `poe verify-deep` | `uv run python scripts/build_production.py --verify --no-launch` | Needs desktop Excel; archives a transcript in `excel-only-runs/`. |
| Build + verify Univariate | `poe verify-deep-univariate` | `uv run python scripts/build_univariate.py --verify --no-launch` | Needs desktop Excel; archives a transcript in `excel-only-runs/`. |
| Build + verify test models | `poe verify-test-models` | `uv run python scripts/build_test_models.py --verify --no-launch --verbose` | Needs desktop Excel; append `--include-heavy` to include the heavy cases (`L05`, `L08`). |
| Run the whole verification ladder | `poe verify` | Run the three deep `verify-*` tasks concurrently, then `verify-headless` over their output | Needs desktop Excel; stops on first failure. |
| Resync workbook-scoped catalog names | `poe resync-names -- <workbook.xlsx>` | `uv run python tools/resync_workbook_names.py <workbook.xlsx>` | Use the `--` separator before positional args. |
| Rebuild static reference sheets | `poe static-sheets` | `uv run python scripts/rebuild_static_sheets.py` | Needs desktop Excel; manual template maintenance. |

To pass extra arguments through a poe task, put them after the task name. For example, `poe verify-test-models --include-heavy` includes the heavy `L05` and `L08` cases, and `poe verify-deep --log scratch/regression-verify.log` overrides the archived transcript path. Use `poe --help` for task-runner options and bare `poe` for the current project task list.

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
uv run pytest tests/test_workbook_invariants.py -v      # or: poe verify-headless

# Build the distributable Regression workbook, Lambda_Library.xlsx  (needs Excel)
uv run python scripts/build_production.py

# Build the standalone Univariate workbook, Lambda_Library_Univariate.xlsx  (needs Excel)
uv run python scripts/build_univariate.py

# Recommended verification path (headless first, then artifact-specific deep checks).
# The three build_* --verify commands need desktop Excel; poe verify-headless does not.
poe verify-headless
uv run python scripts/build_production.py --verify --no-launch
uv run python scripts/build_univariate.py --verify --no-launch
uv run python scripts/build_test_models.py --verify --no-launch

# Verify an already-built workbook without rebuilding it  (needs Excel)
uv run python tools/verify_workbook.py Lambda_Library.xlsx
```

New to the repo? A typical loop is: edit code → `uv run pytest` → `poe verify-headless` → the artifact-specific deep verifier for the surface you touched (`uv run python scripts/build_production.py --verify --no-launch`, `uv run python scripts/build_univariate.py --verify --no-launch`, and/or `uv run python scripts/build_test_models.py --verify --no-launch`) → rebuild the committed distributable(s) with `uv run python scripts/build_production.py` and/or `uv run python scripts/build_univariate.py`. The full flag reference for each script is under [Building](#building) and [Verifying builds](#verifying-builds) below.

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
| `test_data_completeness_qc.py` | `calculate_data_completeness_flags` against the sample dataset |
| `test_catalog_schema.py` | `CatalogDocument` loading, validation, duplicate rejection, projection methods |
| `test_dummy_functions.py` | `Dummy_Levels`/`Dummy_Code` NA()-based error contract: formula statics, signatures, and parser translation to workbook XML |
| `test_lambda_catalog_plain_language.py` | All LAMBDA functions have a `plain_language_summary` in `lambda_functions.json` |
| `test_sheet_writers.py` | Sheet writer integration (conditional formatting, named ranges) |
| `test_spec_block_writer.py` | Spec-block component library: sheet-scoped name definitions and order, T0 default-spec prefill, dropdowns, conditional formats, `Predictor_Columns`/`Constructed_Column_Names` twin invariants |
| `test_analyze_model_construction.py` | Spec-block QC analyzer: default-spec expectations pinned against the sample CSV (mask size, k, level-qualified names), the stratified-Filter degeneracy case, and the observed-vs-expected comparison layer |
| `test_weibull_grid_excel.py` | Weibull grid-search mechanics validation |
| `test_inspection_compare.py` | QC value comparison logic (`to_float_or_none`, `first_digit_deviation`, `compare_values`) |
| `test_independent_verification.py` | Independent numpy/scipy verification of all LAMBDA function outputs (scalars, vectors, observation diagnostics, predictor summary, prediction interval) |
| `test_qc_configs.py` | Internal-consistency invariants of the shared OLS oracle across six model shapes — hat diagonal sums to *p*, residuals sum to zero under an intercept, SS decomposition, prediction-interval symmetry |
| `test_bfn_panel_durbin_watson_verification.py` | `BFN_Panel_Durbin_Watson` against the WHO panel — within-group differencing via `Difference_By`, mutual gating with `Durbin_Watson_By` |
| `test_serial_correlation_group_resolver.py` | `Serial_Correlation_Group()` SWITCH, including the dormant Cluster branch (the reserved-spec-column pattern) |
| `test_difference_by_verification.py` | Gap-aware `Difference_By` — WHO exact counts plus the punched-out-year and calendar-date synthetic cases (the automated form of the retired v2.0 test plan's T17–T19) |
| `test_analyze_regression_spec_block.py` | Post-changeover spec-block QC analyzer (predicted counts and values, regression sheet spec state) |
| `test_regression_spec_qc.py` | Spec-driven Regression QC oracle (`analyze_regression_spec.py` case definitions) |
| `test_csv_dataset_loader.py` | `load_csv_rows` (`write_sheet_csv_dataset.py`) against all three `CsvDatasetConfig`s and the committed sample CSVs |
| `test_mileage_completeness_qc.py` | `calculate_mileage_completeness_flags` (`analyze_mileage.py`) against the Auto MPG dataset |
| `test_intercept_relocation.py` | v3.0 stage 1 — the relocated intercept read back through the context-accessor path (200 datasets, both intercept states), the FE correction routed through element 2 of the same context array, plus the contract assertions: only `Model_Context` declares `Has_Intercept`/`DF_Absorbed`, no carrier reads the context with a bare positional index, `ROWS(Model_Context())` is 4 |
| `test_recording_sheet.py` | The `RecordingSheet` test double itself (`tests/recording_sheet.py`) — the mock every Excel-free sheet-writer test is built on |
| `test_within_estimator.py` | v2.1 Fixed Effects phase 2 — the constructor pipeline — the fit-time pair `Design_Response()`/`Design_Columns()` and its stage order, against an independent `statsmodels` LSDV fit |
| `test_group_panel_transforms.py` | v2.1 Fixed Effects phase 1 — `Group_Mean`, `Demean_By`, `Is_Balanced_Panel`, `Absorbed_Degrees_Of_Freedom` |
| `test_df_absorbed_threading.py` | v2.1 Fixed Effects phase 3 — `[DF_Absorbed]` threaded through SE/t/p/CI/MS-Residual/AIC/BIC/AICc, against an independent `statsmodels` LSDV fit |
| `test_group_prediction_interval.py` | v2.1 Fixed Effects phase 5 — `Group_Mean_At`, `Group_Count_At`, `Prediction_Group_Column`, `Group_Prediction_Interval` (the group-mean-recovery CI+PI form), against an explicit LSDV `get_prediction()` reference |
| `test_doc_links.py` | Every relative `](*.md)` link in the repo's markdown resolves to a file that exists, relative to the linking file's own directory |
| `test_workbook_helpers.py` | `safe_activate()` / `safe_freeze_top_row()` against stub sheet/window objects (headless/no-focus Excel session guards) |
| `test_workbook_builder.py` | Workbook package-patching helpers (`sync_workbook_names` and friends) that don't require Excel |
| `test_build_common.py` | Shared build scaffolding (`lambda_catalog.build_common`: recalculate-and-save calc-mode handling, retry-on-open) that doesn't require Excel |
| `test_build_production.py` | `build_production.py`'s pure-Python logic (Regression-only sheet set, dataset selection, tab order/color, verify forwards `skip_univariate=True`) that doesn't require Excel |
| `test_build_univariate.py` | `build_univariate.py`'s pure-Python logic (Univariate four-sheet set, default output path, Automatic calc mode, `--skip-data-table-calculations` and `--no-calculation` gates, verify forwards `skip_regression=True`) that doesn't require Excel |
| `test_version_history_writer.py` | `write_sheet_version_history`'s per-artifact version lineage (`artifact="regression"` vs `"univariate"`) and the bad-artifact guard |
| `test_workbook_invariants.py` | Layer 1 headless structural check of a built `.xlsx` package (`zipfile` + `lxml`): dangling defined names, `#REF!`/`#NAME?` cached-value literals, broken package parts, orphan chart-relationship targets, sheet drift — for both the Regression and Univariate artifacts — see [Verifying builds](#verifying-builds) |
| `test_ln_positive_verification.py` | v2.2 Transform=Log — `Ln_Positive` pure-Python mirror (the `NA()`-exception contract, the geometric-mean round-trip the Prediction Inputs fix relies on) and implementation-shape assertions on the catalog formula |
| `test_transform_threading.py` | v2.2 Transform=Log wiring end to end — cross-checks the new `production_lots_log_transform` QC case (raw columns, `transform="Log"`) against the pre-existing precomputed-log-column case to floating-point precision; Categorical×Log inertness |
| `test_interaction_wiring.py` | v3.1 interaction wiring — the spec block's M/N pair against the Python mirror in `analyze_regression_spec.build_spec_design`: the three width regimes (1 / L−1 / (L₁−1)(L₂−1)), the closed Product/Difference/Ratio arithmetic, the four operand Role/Include cases, the two-way limit, the documented quadratic, and Ratio's zero-denominator refusal |

### Coverage scope

Coverage measures **all** of `lambda_catalog/` except the modules that drive Excel's COM API. Those are named in the `omit` list in `pyproject.toml`: the `write_sheet_*.py` writers, `workbook_helpers.py`, and `deep_verify.py`. They are validated by the artifact-specific Excel verification commands instead (see [Verifying builds](#verifying-builds)).

**That list is an explicit set of exceptions, not a rule you can re-derive.** Sixteen modules in `lambda_catalog/` import `xlwings`, and only those three are omitted — `workbook_builder.py` (87 %), `build_common.py` (92 %) and `analyze_model_construction.py` (76 %) are all measured, because the `RecordingSheet` COM double exercises most of what they do without Excel. Coverage level is not the boundary either: omitted `write_sheet_version_history.py` reports 100 %, while measured `regression_spec_sheet_io.py` reports 43 %. If you are wondering whether a module is measured, read `pyproject.toml`; there is nothing to infer.

**What *is* a rule is negative: a module that imports neither `xlwings` nor `pywin32` does not belong in `omit` at all.** That is the one this section previously stated too loosely — as "is it about the workbook" — and four modules sat in `omit` under that reading while importing neither, reporting 92 / 87 / 100 / 48 % the moment they were let back in. The costly one was `analyze_regression_sheet.py`. It is the OLS core the live spec oracle calls (`analyze_regression_spec` → `calculate_regression_results_from_matrix`), so the headline coverage number was excluding the arithmetic the whole Regression suite is checking. Do not re-add a module because it *sounds* Excel-shaped; if it does not import a COM binding, it stays measured.

The nine `write_sheet_*.py` writers stay as one glob rather than nine entries. Their individual coverage ranges from 23 % to 100 % depending on how much of each is `RecordingSheet`-reachable, but they are one family with one reason to be listed, and per-file entries would be nine things to keep current instead of one.

This section deliberately no longer enumerates what *is* measured. The list it used to carry had drifted to omit four in-scope modules of its own, which is the failure mode a hand-maintained inventory always reaches — `pyproject.toml` is the source of truth, and it is short because it names only exceptions.

### CI

GitHub Actions runs the test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request. See `.github/workflows/ci.yml`. Coverage must stay at or above 70% on the tracked modules.

## The regression test-model suite

The unit tests above check functions. The **test-model suite** checks *model configurations* — a fixed list of specs, each fitted by the workbook and by an independent Python oracle, compared cell for cell. It is what catches a spec-block change that computes a different model without erroring anywhere.

**The plan of record is [docs/MODEL_TESTING_ASSETS.md](docs/MODEL_TESTING_ASSETS.md).** It names every model the suite should cover, the corner each one is there for, and the datasets future milestones need. Read it before adding a case; add to it before adding a case that isn't listed.

### The regime, in four rules

1. **It is a covering array, not a full factorial.** Every implemented corner case is exercised by at least one model, and every model earns its place by covering something no other model does. Target size: ~25–30 fittable models plus ~10 guard-state configurations. Full crosses (every transform × every interaction × every role) are explicitly out of scope — they multiply sheet count without adding information.
2. **A case is a `RegressionSpecCase`, not a sheet fixture.** Cases are declared in `lambda_catalog/analyze_regression_spec.py` (`SpecVariable` rows built with `_spec_var(...)`, `source_csv_path` / `row_loader` / `source_table_ref` for a non-default dataset, `prediction_group` for the FE prediction box, `back_transform` for the Duan/Naive toggle) and expected values come from `calculate_regression_spec_case`, which fits with NumPy/statsmodels rather than reading the workbook back. A **guard-rail** configuration is a `GuardStateCase` in `lambda_catalog/analyze_regression_guard_states.py` instead — most of them make `calculate_regression_spec_case` raise by design, and what they assert is status text, the per-row Design Columns audit and which CF rules fire, not fit statistics.
3. **Every case is pinned by name.** `_EXPECTED_CASE_NAMES` in `tests/test_regression_spec_qc.py` is an ordered list asserted against `build_regression_spec_cases()`, so a case cannot be added, renamed, reordered, or silently dropped without the test failing. Add the name there in the same commit. Guard cases have the same regime in `_EXPECTED_GUARD_NAMES` (`tests/test_regression_guard_states.py`).
4. **The suite's growth rate orders the roadmap, under one constraint.** From v3.4 on the ladder sorts first by track — all remaining Regression work before the two milestones that open a new analysis surface (v3.10 Two-sample, v3.11 Resampling) — and then, within the Regression track, by how much a milestone forces this suite to grow: additive first, axis-wideners last. See [ROADMAP.md § Ladder order](docs/ROADMAP.md#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).

### Adding a case

1. Write the spec builder in `analyze_regression_spec.py` — a `list[SpecVariable]` with a docstring naming **the corner it covers** and why no existing case covers it.
2. Register it in `build_regression_spec_cases()`. If it targets a dataset other than Auto MPG, set `source_csv_path`, `row_loader`, **and** `source_table_ref` together — `Source_Table` is the one name that retargets which data sheet the spec block reads, and a case that forgets it lands its spec rows on the wrong table's columns.
3. Give it an identity in `_CASE_SHEET_IDENTITY`: the plan ID from [docs/MODEL_TESTING_ASSETS.md § 1](docs/MODEL_TESTING_ASSETS.md) and a worksheet name. The name states **the concept under test, not the variables** (`M05 Log-Log NA Masking`, not `MPG ~ Ln(Weight) + Ln(HP)`) and must satisfy the contract in `lambda_catalog/test_model_sheets.py` — 31 characters, legal charset, `<PlanID> <Concept>`, unique across model *and* guard cases. It is validated at registry-build time, so a bad name fails in the unit suite rather than mid-build.
4. Add the case name to `_EXPECTED_CASE_NAMES` in `tests/test_regression_spec_qc.py`, in position.
5. Add the assertions that make the case worth having — the design-matrix facts (`constructed_column_names`, the row mask, k) plus whatever the corner is actually about.
6. Run `uv run pytest tests/test_regression_spec_qc.py tests/test_test_model_sheets.py` (no Excel), then, on a machine with Excel, `python scripts/build_test_models.py --verify --no-launch` for the case's own sheet and `python scripts/build_production.py --verify --no-launch` for the shipped Regression sheet. See [Verifying builds](#verifying-builds).
7. Update the coverage matrix in [docs/MODEL_TESTING_ASSETS.md § 1.5](docs/MODEL_TESTING_ASSETS.md#15-coverage-matrix) and flip the case's status from **new** to **existing**.
8. Commit the build+verify transcripts on the same branch. Each run writes its own into [`excel-only-runs/`](excel-only-runs/) — `build_test_models verify no launch.log` and `build_production verify no launch.log` — so this step is `git add`, not a copy-paste out of the terminal. The transcript is item 4 of the [PR-shape rules](#the-pr-shape-rules--what-every-regression-pr-must-contain) below; a Regression-track PR without it has no verifiable paper trail because the spec-driven verifier cannot run on Linux CI.

### The PR-shape rules — what every Regression PR must contain

1. **The feature itself.** The sheet writer, lambda, name, or engine call site that delivers the new behavior.
2. **An oracle** in the spec-driven verifier chain — a new branch in `calculate_regression_spec_case` / `calculate_guard_state_case` (or a new `GuardStateCase`) that compares the workbook to an independent NumPy/statsmodels fit. Reading the cell back is not an oracle.
3. **At least one new test-model case** in [docs/MODEL_TESTING_ASSETS.md § 1](docs/MODEL_TESTING_ASSETS.md) + the case registry, pinned in `_EXPECTED_CASE_NAMES` / `_EXPECTED_GUARD_NAMES`, with the § 1.5 coverage-matrix row flipped from **new** to **existing**.
4. **A transcript in `excel-only-runs/`** for the new test-model sheet, archived by `lambda_catalog.build_common.run_log_path` and committed. The spec-driven verifier cannot run on the GitHub-hosted Linux CI (no Excel), so the transcript is the only verifiable artifact a reviewer or future agent has for the new case's first successful build + verify.

A Regression PR that lands items 1–3 but whose transcript shows the new sheet failing `--verify` is still mergeable only if the failure mode is itself an open issue with a tracking link; otherwise the PR is incomplete and the verification commit (item 4) ships in a follow-up PR once the failure is closed.

**Mark a case `heavy=True`** if its sheet is expensive — L08's 173 Fixed Effects groups and L05's k = 19 design matrix on n = 2117 are the current two. The first is gated on build cost; the second is gated on the statsmodels-vs-Excel floating-point floor at fdd = 5/6 that both implementations agree on (it lives here as a deliberate showcase for the floor, not as a defect). The Python oracle always runs; only the sheet build is gated, behind `build_test_models.py --include-heavy`. L07 was the third candidate until the live run showed the workbook cannot fit a 205-column design at all — it is a guard state now, and guard sheets are cheap.

**Check the case actually fits before writing assertions.** Three of the models this plan listed did not survive contact with the data: a full Cylinders × Origin cross has two all-zero columns (singular), `Displacement + Horsepower + (Displacement − Horsepower)` is exactly collinear, and an interaction operand with `Include = FALSE` imposes no mask condition — so a sparse operand poisons the design with `#N/A` rather than dropping rows. A spec that Excel answers with `#NUM!` while NumPy quietly returns a minimum-norm solution cannot be compared at all.

### The test-model workbook

Every case is also materialized as its own worksheet in `Lambda_Library_TestModels.xlsx` (gitignored — a fixture, not an artifact). `build_test_models.py` builds it; `tools/inspect_test_model_sheets.py` verifies it by **reading only**. Both halves — writing a case onto a sheet and reading one back — live in `lambda_catalog/regression_spec_sheet_io.py`, shared with the single-sheet verifier so the two cannot disagree about what a case is. See [docs/MODEL_TESTING_ASSETS.md § 1b](docs/MODEL_TESTING_ASSETS.md#section-1b--one-worksheet-per-test-model).

### Datasets

Three are wired: Auto MPG (406 rows — baseline, categoricals, interactions), Life Expectancy (2938 rows — transform dispatch, scale, missingness), Production Lots (51 rows — learning curves, fixed effects, sequence). A new dataset must buy a corner those three cannot; keep additions at ~250 rows or fewer, prefer CSVs already in `sample_data/`, and expect to write one `CsvDatasetConfig` (`lambda_catalog/write_sheet_csv_dataset.py`) plus one `SpecDatasetProfile` registry entry (`SPEC_DATASET_PROFILES` in `write_spec_block.py`). The shortlist and what each one buys is [§ 3](docs/MODEL_TESTING_ASSETS.md#section-3--supplemental-datasets-kept-minimal).

## Building

There are two separate build scripts with distinct purposes. From v3.0 the production script emits **two artifacts** rather than one.

### The two production artifacts

| Target | Produces | Calculation mode | Sheets |
|---|---|---|---|
| Regression | `Lambda_Library.xlsx` | **Automatic** (full) | Catalog, three sample datasets, Regression, the two reference sheets, Version History |
| Univariate | `Lambda_Library_Univariate.xlsx` | **Automatic** (full) | Catalog, Life Expectancy Data, Univariate Analysis, Version History |

**Both artifacts carry the complete function library.** All 141 LAMBDA definitions are written into both Name Managers. There is no bundling step, no dependency closure, and no per-artifact function subsetting — the artifacts differ only in which sheets they contain. When you add a function, it lands in both; there is no list to update.

**Two named build targets, not flags.** `build_production.py` (Regression) and `build_univariate.py` (Univariate) are separate driver scripts sharing one scaffolding module (`lambda_catalog/build_common.py`). The old `--skip-univariate` flag is retired — the Regression target simply never writes the Univariate sheet, and the Univariate target never writes the Regression sheets. `--skip-data-table-calculations` survives (flag name kept for CLI stability) and is the Univariate target's primary fast-iteration flag (it skips the slow Beta `Full_Factorial` grid-spill rebuild).

### Which version number moves

Two numbers, and a change usually moves exactly one:

| You changed | Moves |
|---|---|
| A LAMBDA definition in `lambda_functions.json` — added, renamed, or different return | **The library version.** It affects *both* artifacts, because both ship the whole catalog |
| A sheet's layout, input cells, or control block | **That workbook's version only** |
| Both | Both |

**A change to a shared function is a shared event.** There is no such thing as "a Univariate-only function change" — every function ships in both workbooks, so a catalog edit reaches every user of either artifact and moves the one library version. Conversely, a Univariate sheet-layout change must **not** move the Regression workbook version: that number is what a Regression user reads to answer "do my existing inputs still work?"

The `Breaking?` flag in each Version History sheet attaches to the **workbook** version, never the library version. Full conventions in [ROADMAP.md § Versioning](docs/ROADMAP.md#versioning--release-conventions).

### Production build

```powershell
uv run python scripts/build_production.py
```

Produces `Lambda_Library.xlsx` — the distributable Regression artifact committed to the repo. Writes eight sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions
- **Life Expectancy Data** — WHO dataset as a structured table (a second sample dataset for practicing the Source_Table retarget workflow)
- **Mileage Data** — Auto MPG dataset as a structured table; this is the dataset the Regression sheet's `Source_Table` targets by default
- **Production Lots** — a small unbalanced learning-curve panel (3 facilities, 51 lots) as a structured table; a third sample dataset, and the only one with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year)
- **Regression Instructions** — step-by-step guide for adapting the sheet to new datasets
- **Diagnostic Guide** — interpretation guide for regression diagnostics
- **Version History** — changelog that travels with the workbook
- **Regression** — ToolPak-style analysis interface driven by a declarative variable-specification block (the spec block) and the sheet-scoped names that assemble the design matrix from it. The wiring names (`Source_Data`, `Header_Names`, `Spec_*`) hardcode the spec block's cell addresses and are defined in `write_spec_block.py` (imported by `write_sheet_regression.py`); the constructor closures (`Sample_Include`, `Response_Column`, `Row_Labels`, `Predictor_Columns`, `Design_Columns`, `Design_Response`, `Constructed_Column_Names`) live in `lambda_functions.json` with `"scope": "Regression"`, so they are the single source of truth and appear on the LAMBDA_functions catalog sheet (Scope column) like any other function — they are just installed on this sheet rather than workbook-wide

No Univariate sheet (it ships in its own workbook — see [Univariate build](#univariate-build) below), no test sheets, no OLS analysis, no cache dependency.

**All `build_production.py` options:**

| Flag | Default | What it does |
|---|---|---|
| `--workbook PATH` | `Lambda_Library.xlsx` | Path to the workbook to create or update. |
| `--definitions PATH` | `lambda_functions.json` | Path to the JSON catalog of LAMBDA definitions. |
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV written to the **Life Expectancy Data** sheet. (The **Mileage Data** and **Production Lots** sources are fixed committed sample files with no CLI override.) |
| `--regression-dataset {auto_mpg,life_expectancy,production_lots}` | `auto_mpg` | Which dataset the Regression sheet's `Source_Table` targets, **and** which shipped default spec pre-fills the MODEL SPECIFICATION block (`SPEC_DATASET_PROFILES` in `write_spec_block.py`) — every column starts with a real Role/Include/Type instead of falling back to an un-flagged Predictor. The profile decides which rows arrive pre-filled, not how many spec rows exist: the block sizes itself from `COLUMNS(Source_Data)`, so retargeting `Source_Table` by hand afterwards resizes it too. `life_expectancy` ships Response=`Life expectancy`, the 18-column `FEATURE_COLUMNS` predictor set, `Country` as Identifier, `Status` as a Categorical predictor, and `Year` as the Sequence axis. `production_lots` is the one to pick for a ready-made Fixed Effects example (Facility as the FE role, Fiscal_Year as Sequence) — its default spec is the QC-validated Crawford/Wright learning-curve model (`log Unit Cost` ~ `log Cum Units`). |
| `--skip-data-table-calculations` | off | No effect for the Regression workbook (it has no Data Tables — never did). The final `CalculateFullRebuild` is cheap and the Regression sheet needs it (the verifier's per-sheet `Calculate` doesn't rebuild the dependency tree after a name sync), so the rebuild always runs regardless of this flag. The flag matters for `build_univariate.py`, whose Beta `Full_Factorial` grid spills make the rebuild slow (flag name kept for CLI stability; the Univariate artifact no longer uses any Excel Data Table). |
| `--verify` | off | After the build, run the spec-driven verifier (`lambda_catalog.deep_verify.verify_test_sheets` with `skip_univariate=True`) against the production sheets. On any drift, print a structured `VerifyReport` and `sys.exit(1)`. The Excel handoff only fires when verify passes, so a stale build can't launch in place of a fresh one. |
| `--no-verify` | (default) | Explicitly disable the verifier pass. Mainly for wrapper scripts that default to `--verify`. |
| `--no-launch` | off | Suppress the post-build `cmd /c start <workbook>` Excel handoff. Use in agentic/automated loops where no Excel window should pop up. |
| `--validate-reopen` | off | Reopen the workbook after syncing names to confirm Excel accepts the result. |
| `--verbose` | off | Print per-phase timing checkpoints to stdout. |
| `--log PATH` | `excel-only-runs/<script> <flags>.log` | Where to archive this run's transcript. Every run is teed to this file — stdout, stderr, the `VerifyReport` on drift, and the traceback of anything that aborts the build. This target cannot run in CI, so the transcript is the branch's only paper trail; commit it. |

Common combinations:

```powershell
# Plain build → opens Lambda_Library.xlsx in Excel when done
uv run python scripts/build_production.py

# The one-shot automated flow: build, sync names, run the spec-driven verifier,
# exit non-zero on drift, and never open Excel. This is what `poe verify-deep` runs.
# The Regression workbook has no Data Tables, so the rebuild is cheap and always runs.
# The transcript lands in `excel-only-runs/build_production verify no launch.log`.
uv run python scripts/build_production.py --verify --no-launch
```

### Univariate build

```powershell
uv run python scripts/build_univariate.py
```

Produces `Lambda_Library_Univariate.xlsx` — the distributable Univariate artifact committed to the repo. Writes four sheets: **LAMBDA_functions**, **Life Expectancy Data** (the dataset the Univariate data zone reads via `LifeExpectancyData[Life expectancy]`), **Univariate** (descriptive statistics, histogram binning, and the two-stage MLE fitting — Weibull and Gamma via 1-D profile-NLL searches, Beta via two `Full_Factorial` dynamic-array grid spills, the other five distributions in closed form; **no Excel Data Tables**), and **Version History** (the Univariate artifact's own lineage, starting at 1.0.0). Carries the complete 141-function library; no Regression-side sheets.

**All `build_univariate.py` options:**

| Flag | Default | What it does |
|---|---|---|
| `--workbook PATH` | `Lambda_Library_Univariate.xlsx` | Path to the workbook to create or update. |
| `--definitions PATH` | `lambda_functions.json` | Path to the JSON catalog of LAMBDA definitions. |
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV written to the **Life Expectancy Data** sheet. |
| `--skip-data-table-calculations` | off | Skip the final Excel `CalculateFullRebuild`. Beta's two `Full_Factorial` grid spills make that rebuild slower than a plain formula recalc, so this is the primary fast-iteration flag for this artifact (flag name kept for CLI stability — the Univariate artifact no longer uses any Excel Data Table). Note: the verifier's per-sheet `Calculate` does not reliably resolve the Beta spills after a name sync, so combining this with `--verify` may report stale-fit mismatches a real rebuild would not. |
| `--no-calculation` | off | **Never calculate.** Excel stays in Manual for the whole run: the build never switches to Automatic, suppresses Excel's recalculate-before-saving for the duration, and skips the final rebuild. Stronger than `--skip-data-table-calculations`, which still pays a full calculation when the build switches to Automatic ahead of the save — see the note below. For inspecting structure (name manager, sheet layout, spill anchors) without paying for the grid searches. The workbook it leaves has stale computed cells and is saved in Manual mode, so **do not ship it**; the next ordinary build fixes both. |
| `--verify` | off | After the build, run the spec-driven verifier (`lambda_catalog.deep_verify.verify_test_sheets` with `skip_regression=True`) against the Life Expectancy and Univariate sheets. On any drift, print a structured `VerifyReport` and `sys.exit(1)`. |
| `--no-verify` | (default) | Explicitly disable the verifier pass. |
| `--no-launch` | off | Suppress the post-build `cmd /c start <workbook>` Excel handoff. |
| `--validate-reopen` | off | Reopen the workbook after syncing names to confirm Excel accepts the result. |
| `--verbose` | off | Print per-phase timing checkpoints to stdout. |
| `--log PATH` | `excel-only-runs/<script> <flags>.log` | Where to archive this run's transcript, exactly as for `build_production.py` — including the `--no-calculation` + `--verify` warning, which goes to stderr. |

Common combinations:

```powershell
# Build + verify the standalone Univariate workbook. This is `poe verify-deep-univariate`.
# The rebuild runs by default so the shipped Beta Full_Factorial spills are computed, not stale.
uv run python scripts/build_univariate.py --verify --no-launch

# Fast iteration: write the sheets and sync names, skip the slow Beta grid-spill rebuild.
uv run python scripts/build_univariate.py --skip-data-table-calculations --no-launch

# Structure only: write the sheets and sync names, calculating nothing at all.
# Use this to inspect the name manager or the layout; the result is not shippable.
uv run python scripts/build_univariate.py --no-calculation --no-launch
```

**Why `--no-calculation` exists when `--skip-data-table-calculations` already
skips the rebuild.** The rebuild is not the only calculation in the build.
Before saving, `build_univariate_workbook` sets `Application.Calculation` to
Automatic — and setting Automatic on an open workbook calculates it *there and
then*, Beta's `Full_Factorial` grid spills included. `--skip-data-table-calculations` skips the
phase-2 `CalculateFullRebuild` and still pays that one.
`--no-calculation` is the flag that pays neither: Automatic is never set, and
because Excel under Manual still recalculates on save unless told otherwise, it
also turns `Application.CalculateBeforeSave` off for the duration and restores
it afterwards (it is an Excel-wide setting, not a per-workbook one, so leaving
it off would change every later session on that machine).

If all you need is the **workbook-scoped** name manager, you do not need Excel
at all — `python tools/resync_workbook_names.py Lambda_Library_Univariate.xlsx`
rewrites and reports those names from the catalog in about a second. Reach for
`--no-calculation` when you need the *sheet-scoped* names or the layout, which
only the sheet writers produce.

## Verifying builds

There are two verifier layers with different speeds and different scopes. Run them in this order when in doubt; either can be skipped if the other has been run recently.

### Layer 1 — headless structural check

Pure `zipfile` + `lxml` reads of the produced `.xlsx`. Runs in <1 s on Linux CI without Excel. Catches packaging regressions the unit-test suite misses: dangling defined names, `#REF!`/`#NAME?` cached-value literals, broken `[Content_Types].xml`/`workbook.xml.rels`, orphan chart-relationship targets, `localSheetId` out of range, sheet drift.

```powershell
poe verify-headless
poe test-excel               # + the opt-in real-workbook tests

# Or directly:
uv run pytest tests/test_workbook_invariants.py -v
RUN_EXCEL_INTEGRATION=1 uv run pytest tests/test_workbook_invariants.py -v   # + real-workbook tests
```

This is a fast screen. A green run does **not** mean the workbook calculates correctly — that is what Layer 2 is for.

### Layer 2 — spec-driven deep check

Reuses `lambda_catalog.deep_verify.verify_test_sheets` against the production sheets. This is the source of truth for cell-level correctness.

```powershell
# Run the Regression production build, recalculate, then verify against the
# spec oracle. On drift: print a structured VerifyReport, sys.exit(1), and do
# NOT open Excel (so a stale build cannot be launched in place of a fresh one).
# The rebuild is cheap (no Data Tables — Regression never had any) and always runs — it is the source of
# truth the verifier reads; do not pair --verify with --skip-data-table-calculations.
python scripts/build_production.py --verify --no-launch

# Same, for the standalone Univariate workbook:
python scripts/build_univariate.py --verify --no-launch

# Same, for the Regression model-case fixture workbook:
python scripts/build_test_models.py --verify --no-launch

# Or, on the just-built workbook without rebuilding:
uv run python tools/verify_workbook.py Lambda_Library.xlsx
uv run python tools/verify_workbook.py Lambda_Library_Univariate.xlsx --skip-regression
uv run python tools/verify_workbook.py Lambda_Library.xlsx --json   # agentic consumption
```

**All `tools/verify_workbook.py` options** (positional `workbook` path is required):

| Flag | Default | What it does |
|---|---|---|
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV used for the `Full_Data` comparison. |
| `--mileage PATH` | `sample_data/auto_mpg_data.csv` | Auto MPG CSV for the Mileage Data `Full_Data` comparison and the Regression spec oracle. |
| `--skip-regression` | off | Skip every Regression / Mileage / Production Lots check. Use to verify the standalone Univariate workbook, which carries none of those sheets; the Life Expectancy and Univariate checks still run. |
| `--json` | off | Emit the report as JSON (stable schema, for agentic consumption) instead of the human-readable form. |
| `--verbose` | off | Print per-phase checkpoints from the spec-driven verifier. |

The `VerifyReport` (in `lambda_catalog/verify_report.py`) is emitted both in a human-readable form (`Verify: passed (spec mode, …)` / `ERROR Verify mismatch totals: …`) and as JSON (stable schema: `passed`, `categories`, `failures`, `elapsed_seconds`, `mode`, `workbook`).

Expected terminal flow for the one-shot command:

1. Sheet update summary (`Sheet updated: ...`, `Created names: ...`, `Updated names: ...`).
2. Spec verifier result (`Verify: passed ...` or `ERROR Verify mismatch totals: ...`).
3. Timing summary lines (`Timing: build+sync`, `Timing: recalculate` or `skipped`, `Timing: verify`, `Timing: total`).


### `poe verify`

```powershell
poe verify                  # both layers, both artifacts, plus the test-model suite
poe verify-headless         # Layer 1 only (any platform)
poe verify-deep             # Layer 2, Regression artifact (needs Excel)
poe verify-deep-univariate  # Layer 2, Univariate artifact (needs Excel)
poe verify-test-models      # Layer 2, the ~48-sheet test-model suite (needs Excel)
```

**`poe verify` runs the three builds concurrently, then screens their output.** It stops at the first stage that exits non-zero, and the stage boundary is what makes the order matter: `verify-headless` reads whatever `.xlsx` files are sitting in `dist/`, and the deep tasks *rewrite* those files. The task used to run the screen first, which meant it validated the previously committed artifacts and never looked at the ones the run had just built — a rebuild that broke a defined name or orphaned a chart relationship passed `verify` clean. Builds first, screen last.

The three builds overlap safely because they share nothing: each driver opens its own `xw.App(visible=False, add_book=False)` and reaches every workbook through that instance's `app.books` handle (there is no bare `xw.Book()` or `xw.apps.active` anywhere in the package), and they write three different artifacts and three differently-named transcripts. The one file two of them could have contended over is `templates/static_sheets.xlsx`, which `copy_static_sheet` opens read-only. Output is buffered per task rather than interleaved, so each transcript stays contiguous.

Wall time becomes roughly the longest build — `verify-test-models`, minutes — instead of the sum of all three. The cost is three Excel instances competing for CPU; on a constrained machine, run the `verify-*` tasks one at a time instead. **None of this is checkable in CI** (no GitHub-hosted runner has Office), so a change to the `verify` task needs a developer-machine run archived to `excel-only-runs/`.

`poe verify-deep` shells out to `build_production.py --verify --no-launch` and `poe verify-deep-univariate` to `build_univariate.py --verify --no-launch`, so each both rebuilds and verifies its own artifact. Both tee their run into [`excel-only-runs/`](excel-only-runs/) (`<script> <flags>.log`, via `lambda_catalog.build_common.run_log_path`) — stderr and any traceback included — so a failed deep verify is a file you can commit and hand over rather than terminal scrollback; override the destination with `--log PATH`. To verify an already-built workbook, use `python tools/verify_workbook.py <workbook>` instead (with `--skip-regression` for the Univariate artifact).

The `verify-test-models` task passes `--verbose` because that run takes minutes across ~48 sheets: it names each sheet *before* writing it, so an interrupted run leaves the offending case on screen. It archives its transcript the same way — all three Excel-required drivers do. The heavy `L08` case is excluded by default; append `--include-heavy` (`poe verify-test-models --include-heavy`) to include it. Its Python oracle runs in the unit suite regardless.

### CI

GitHub Actions runs the unit-test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request via `.github/workflows/ci.yml`. Its two check steps invoke the poe tasks (`uv run poe test-cov`, `uv run poe lint`) rather than spelling out their own pytest and pylint invocations, so CI and a developer machine cannot drift apart — changing a task in `pyproject.toml` changes both. The spec-driven verifier (Layer 2) is **not** run in CI: the GitHub-hosted `windows-latest` runner image does not include Microsoft Office, so xlwings fails to dispatch `Excel.Application` (`pywintypes.com_error: (-2147221005, 'Invalid class string')`). Until a self-hosted runner with Office is wired in, Layer 2 must be run on a developer machine (or any Windows box with desktop Excel) — the agentic workflow runs it before pushing. The `windows-verify` job was removed for that reason; see the comment block at the bottom of `ci.yml`. Layer 1 (the headless `tests/test_workbook_invariants.py` suite) needs no Excel and is auto-discovered by the existing Linux job, so it runs on every push and pull request.

## File structure

```
scripts/
  build_production.py        # Regression production entry point → dist/Lambda_Library.xlsx
  build_univariate.py        # Univariate production entry point → dist/Lambda_Library_Univariate.xlsx
  build_test_models.py       # Regression model-case fixture builder → Lambda_Library_TestModels.xlsx
                              # (gitignored; verify with --verify --no-launch)
  rebuild_static_sheets.py   # regenerates templates/static_sheets.xlsx from its Python source —
                              # see "Static reference sheets" below
dist/                        # the two shipped .xlsx artifacts (build output, committed)
excel-only-runs/             # archived --verify transcripts from developer-machine runs
lambda_functions.json         # LAMBDA definitions (source of truth)
sample_data/
  Life Expectancy Data.csv   # WHO life expectancy dataset
  auto_mpg_data.csv          # Auto MPG dataset (second sample dataset, "Mileage Data" sheet)
  production_lots.csv        # Learning-curve panel (third sample dataset, "Production Lots" sheet) —
                              # the only shipped dataset with a natural Fixed Effects grouping column
templates/
  static_sheets.xlsx         # pre-built copies of dataset-independent reference sheets
                              # (Regression Instructions, Diagnostic Guide) — see
                              # "Static reference sheets" below
lambda_catalog/
  catalog_schema.py          # typed document model: CatalogArgument, CatalogFunction, CatalogDocument
  build_common.py            # shared build scaffolding (retry-on-open, recalculate-and-save) for the two production scripts
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
  lambda_formula_parser.py   # converts display formulas to workbook XML syntax
  inspection_compare.py      # numeric comparison helpers for QC value verification
  deep_verify.py             # shared xlwings spec-driven verifier used by build scripts and tools/verify_workbook.py
  verify_report.py           # VerifyReport: structured pass/fail result for the spec-driven verifier
  write_sheet_lambda_functions.py
  write_sheet_csv_dataset.py # unified loader/writer/CLI for Life Expectancy, Mileage, and Production Lots
  write_sheet_univariate.py
  write_sheet_regression_instructions.py
  write_sheet_diagnostic_guide.py
  write_sheet_version_history.py
  write_sheet_regression.py
  write_spec_block.py
tools/
  inspect_regression_sheet.py # Regression sheet comparison (loaded by lambda_catalog.deep_verify)
  inspect_univariate_sheet.py # Univariate sheet comparison (loaded by lambda_catalog.deep_verify)
  inspect_xlsx.py            # workbook inspection utility
  verify_workbook.py         # standalone CLI wrapping the spec-driven verifier
```

## File naming conventions

- `build_*.py` — workbook-level entry points that open or create an Excel workbook
- `write_sheet_*.py` — worksheet writers, each responsible for one sheet; can also be run standalone
- `lambda_catalog/` — installable package containing all writers and shared helpers

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

`write_sheet_regression_instructions.py` and `write_sheet_diagnostic_guide.py` write sheets whose content never depends on the target dataset — a fixed how-to guide and a fixed diagnostics reference. Rebuilding hundreds of styled cells with COM calls for unchanging text on every production build is wasted work, so these two modules instead copy an already-styled sheet out of `templates/static_sheets.xlsx` via `workbook_helpers.copy_static_sheet` (Excel's native `Sheet.Copy` between two workbooks open in the same Excel instance — not an openpyxl round-trip; see CLAUDE.md's "Use xlwings COM API for all chart creation — never openpyxl" for why openpyxl is unsafe for anything Excel-native like this). `write_regression_instructions_sheet(workbook)` / `write_diagnostic_guide_sheet(workbook)` keep their original call signature, so the production build scripts and their tests are unaffected.

The authored content still lives in Python — `_ROWS` in `write_sheet_regression_instructions.py`, the body of `_write_template_sheet` in `write_sheet_diagnostic_guide.py` — but the production build scripts never execute it; they only call the copy-from-template functions above. Regenerating the template is a separate, manual step.

Run **`python scripts/rebuild_static_sheets.py`** after editing either sheet's content, then commit the updated `templates/static_sheets.xlsx` alongside the Python change. It opens the template once, calls every static sheet's `_write_template_sheet(workbook)` (so nothing is skipped or forgotten), and saves once. This is the standard command — prefer it over the per-module CLIs below, which exist only for regenerating a single sheet in isolation while debugging:

```powershell
python scripts/rebuild_static_sheets.py                       # regenerates every static sheet (standard)
python -m lambda_catalog.write_sheet_regression_instructions  # single-sheet debugging only
python -m lambda_catalog.write_sheet_diagnostic_guide         # single-sheet debugging only
```

**Why a dedicated command exists:** before it was added, each sheet's own CLI was the only way to regenerate the template, so editing `_ROWS` or `_write_template_sheet` and forgetting to also run that specific CLI would silently ship stale reference text — the template drifted from its Python source with no error anywhere in the build. That happened at least twice (see DECISIONS.md → "Static template drift"). `rebuild_static_sheets.py` collapses "which CLI do I need to remember to run" into a single always-correct command.

All of this — the per-module CLIs and `rebuild_static_sheets.py` alike — requires a real Excel COM engine (`xlwings.App`); none of it runs in a headless/CI environment.

## Documentation drift

`lambda_functions.json` is the source of truth for functions, but nothing is the source of truth for the *documented* state, and the planning docs have drifted from the code more than once. Examples caught by hand: `ROADMAP.md` listed a milestone as planned that was fully built; `ARCHITECTURE.md` documented Role dropdown values without the parenthetical suffixes that formulas actually string-compare against.

Three mechanical checks would catch most of this class. All are pure Python and need no Excel, so they run in the existing Linux CI job. **Two and a half are built:**

1. **Link targets — built, `tests/test_doc_links.py`.** Every relative `](target.md)` link resolves to a file that exists, relative to the *linking file's own directory*. This is the check that would have caught two real breakages: the deletion of `REVIEW.md` while four documents still linked to it, and a docs-reorganization pass that prefixed every relative link with `docs/` — including links already inside `docs/`, where the prefix is one level too many. `docs/TODOs.md` → `docs/ROADMAP.md` resolves to `docs/docs/ROADMAP.md`; that single commit broke 171 links and nothing failed.
2. **Cross-document anchors — built, same file.** Every `](target.md#anchor)` and `](#anchor)` resolves to a heading that exists in the target file, matched by reproducing GitHub's slug rules. Heading renames silently break these: the `ARCHITECTURE.md` §4 renames from `(A–L)` to `(A–N)` (v2.1, adding the Sequence Period / Period In Use pair) and from `(A–N)` to `(A–O)` (v3.0, adding the Design Columns audit column) each broke at least one `ROADMAP.md` link this way. It found one live break the moment it was written — retitling *this very section* from "Documentation drift (proposed check — not yet implemented)" broke the `TODOs.md` link into it, in the same commit that built check 1.
3. **Function names — the count half is built, `tests/test_doc_catalog_counts.py`.** The built half asserts that every *count* of catalog functions stated in README / CONTRIBUTING / ROADMAP matches `lambda_functions.json` — total, workbook-scoped, and Regression-scoped. It found four stale numbers on its first run. The unbuilt half is the harder one: every *name* written as a function reference resolves to a catalog entry, unless it is a native Excel function or explicitly tagged as planned. That is what would have caught an older review's claim that `Interact` was shipping when it was only specified, and the stale-rename list before that.

The name half of item 3 is recorded as a scoped follow-up in [docs/TODOs.md](docs/TODOs.md#documentation), not a claim.

**Both built checks pin their inputs rather than sniffing them,** and both will fail when the docs grow a phrasing they do not know — that is the intended bargain, the same one `_EXPECTED_TASK_NAMES` and `_EXPECTED_CASE_NAMES` make. The count check in particular matches three exact phrasings instead of "a number near the word LAMBDA", because README's first line reads *"Excel 365 LAMBDA functions…"* and a looser rule would fail the build on the Office release number.

## Adding a new LAMBDA function

1. Add an entry to `lambda_functions.json` with `name`, `formula_display`, `arguments`, `yields`, `description`, and `plain_language_summary`. Add `notes` for the Name Manager tooltip (255 characters max), and `scope` only when the function is a sheet-scoped closure rather than a portable workbook name.
2. Add or update the relevant Python oracle when the function feeds a production analysis surface (for example, Regression outputs in `analyze_regression_sheet.py` / `analyze_regression_spec.py` or Univariate outputs in `analyze_univariate.py`).
3. Run the appropriate verifier (`poe verify-headless`, `python scripts/build_production.py --verify --no-launch`, `python scripts/build_univariate.py --verify --no-launch`, and/or `python scripts/build_test_models.py --verify --no-launch`) and confirm no unexpected WARNING lines appear.
4. Run `python scripts/build_production.py` and/or `python scripts/build_univariate.py` to rebuild the distributables that carry the function.
5. Move the **library version**, not a workbook version — a new function ships through the catalog. See [Which version number moves](#which-version-number-moves).

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

**Naming convention** — all OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix, distinguishing them from the constructor closures (`Design_Columns`, `Sample_Include`, etc.) and formula-helper names:

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

**Scope:** all names are worksheet-scoped (created via `sheet.api.Names.Add`), and so is every other range a sheet writer creates. Workbook scope is the catalog's alone — see [Workbook scope belongs to the catalog](#workbook-scope-belongs-to-the-catalog) below. Chart `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer:

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

### Workbook scope belongs to the catalog

Two scopes, two owners, no overlap:

| Scope | Owner | Created by | Examples |
|---|---|---|---|
| Worksheet | the sheet writers | `sheet.api.Names.Add` during the write phase | `RegChart*`, `UV_*`, `Source_Data`, `Spec_*`, the constructor closures |
| Workbook | `lambda_functions.json` | `sync_workbook_names` patching `xl/workbook.xml` after the write phase | every catalog LAMBDA with `scope: "workbook"` |

`sync_workbook_names` enforces the second row literally. On every build it removes **every** workbook-scoped `<definedName>` that is neither a catalog function nor one of Excel's reserved `_xlnm.*` names, then writes the catalog entries fresh. Sheet-scoped entries are never touched.

Anything workbook-scoped and outside the catalog is residue from an earlier build, and the v3.0 split produced a lot of it: each artifact was carrying the other one's chart ranges at workbook scope, pointing at `#REF!` (twelve `RegChart*` entries in the Univariate workbook, forty-two `UV_*` entries in the Regression workbook), plus twenty-one LAMBDA names the catalog had retired. The sheet-scoped originals were correct the whole time — only the stale workbook-scoped copies were broken, which is why the workbooks still rendered their charts.

**A catalog function that names a missing worksheet is skipped, not written.** Excel does not leave such a reference unresolved: it rebinds it to an external workbook (`Regression!Source_Data` becomes `[1]!Source_Data`), writes an `xlExternalLinkPath/xlPathMissing` external-link part, and prompts about broken links every time the file is opened. No catalog function is sheet-qualified today. `Base_Period_Delta` was the one — its body read `'Regression'!Source_Data` / `Spec_Sequence` / `Spec_Sequence_Period`, so the standalone Univariate artifact did not carry it — and it is now **sheet-scoped** with unqualified references, one definition per Regression-shaped sheet. The guard stays because it is what stops the next sheet-qualified body shipping a broken link. The build prints `Skipped names: …` when it happens. When the last external reference goes, the orphaned external-link parts, relationships and content-type overrides are stripped with it.

Keep a new workbook-scoped catalog LAMBDA sheet-agnostic unless it is deliberately Regression-only. If it must read the spec block, expect it to be skipped in the Univariate artifact.

**Checks and repair.** `tests/test_workbook_invariants.py::TestRealWorkbookNameScope` asserts workbook-scope ownership, error-free defined-name bodies, and the absence of external links against both committed artifacts on every commit — pure zipfile + lxml, so it runs in CI without Excel. To re-apply the cleanup to a built artifact without a full rebuild (also Excel-free):

```bash
python tools/resync_workbook_names.py Lambda_Library.xlsx Lambda_Library_Univariate.xlsx
```

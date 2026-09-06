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
| Build the workbook | `poe build` | `uv run python scripts/build_production.py` | Needs desktop Excel. |
| Build + verify the workbook | `poe verify-deep` | `uv run python scripts/build_production.py --verify --no-launch` | Needs desktop Excel; archives a transcript in `excel-only-runs/`. |
| Build + verify test models | `poe verify-test-models` | `uv run python scripts/build_test_models.py --verify --no-launch --verbose` | Needs desktop Excel; append `--include-heavy` to include the heavy cases (`L05`, `L08`). |
| Build + verify the guard states | `poe verify-guards` | `uv run python scripts/build_test_models.py --verify --no-launch --verbose --kind guards --exclude L07` | Needs desktop Excel; every `GuardStateCase` except the 205-column `L07`. |
| Build + verify the spec-block error surfaces | `poe verify-spec-errors` | `uv run python scripts/build_test_models.py --verify --no-launch --verbose --cases <20 IDs>` | Needs desktop Excel; the five row-2 status cells and the ten CF flag rules. Includes `L07`, so it is the slower of the two guard slices. |
| Build + verify the happy-path fits | `poe verify-models` | `uv run python scripts/build_test_models.py --verify --no-launch --verbose --kind models` | Needs desktop Excel; the 33 fittable cases, heavy excluded. No guard sheets. |
| Finish a run `verify-spec-errors` started | `poe verify-models-rest` | `uv run python scripts/build_test_models.py --verify --no-launch --verbose --kind models --exclude G08,M15,L12` | Needs desktop Excel; the 30 fittable cases `verify-spec-errors` did **not** already build. Disjoint from it — see *Narrower slices*. |
| Run the whole verification ladder | `poe verify` | Run `verify-deep` + `verify-test-models` concurrently, then `verify-headless` over their output | Needs desktop Excel; stops on first failure. |
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

# Build the distributable workbook, Lambda_Library.xlsx  (needs Excel)
uv run python scripts/build_production.py

# Recommended verification path (headless first, then deep checks).
# The build_* --verify command needs desktop Excel; poe verify-headless does not.
poe verify-headless
uv run python scripts/build_production.py --verify --no-launch
uv run python scripts/build_test_models.py --verify --no-launch

# Verify an already-built workbook without rebuilding it  (needs Excel)
uv run python tools/verify_workbook.py Lambda_Library.xlsx
```

New to the repo? A typical loop is: edit code → `uv run pytest` → `poe verify-headless` → the deep verifier for the surface you touched (`uv run python scripts/build_production.py --verify --no-launch` and/or `uv run python scripts/build_test_models.py --verify --no-launch`) → rebuild the committed distributable with `uv run python scripts/build_production.py`. The full flag reference for each script is under [Building](#building) and [Verifying builds](#verifying-builds) below.

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
| `test_write_sheet_univariate_cartesian.py` | The cartesian Beta grid-search writer (`write_sheet_univariate_cartesian.py`) — separate from `test_sheet_writers.py` so the cartesian writer can grow its own coverage without swelling the standard-writer suite |
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
| `test_apply_case_inputs.py` | `apply_case_inputs` — the one helper that composes the three writes every fittable case needs (`apply_spec_case`, the typed Sequence Period into column I, `set_prediction_inputs`) so a second call site cannot forget the middle step |
| `test_spec_input_columns_are_cleared.py` | `apply_spec_case` clears every writable spec column before rewriting it, so a case is never evaluated against whatever the previous write left behind |
| `test_regression_spec_qc.py` | Spec-driven Regression QC oracle (`analyze_regression_spec.py` case definitions) |
| `test_regression_guard_states.py` | Guard-state QC oracle (`analyze_regression_guard_states.py` case definitions) — status text, the per-row Design Columns audit, the model formula, and which CF rules fire; pinned in `_EXPECTED_GUARD_NAMES` |
| `test_test_model_sheets.py` | The one-sheet-per-test-model framework — the sheet-naming contract (`<PlanID> <Concept>`, Excel's limits, uniqueness), coverage in both directions across both registries, and the writer's per-sheet parameterization |
| `test_csv_dataset_loader.py` | `load_csv_rows` (`write_sheet_csv_dataset.py`) against all three `CsvDatasetConfig`s and the committed sample CSVs |
| `test_mileage_completeness_qc.py` | `calculate_mileage_completeness_flags` (`analyze_mileage.py`) against the Auto MPG dataset |
| `test_intercept_relocation.py` | v3.0 stage 1 — the relocated intercept read back through the context-accessor path (200 datasets, both intercept states), the FE correction routed through element 2 of the same context array, plus the contract assertions: only `Model_Context` declares `Has_Intercept`/`DF_Absorbed`, no carrier reads the context with a bare positional index, `ROWS(Model_Context())` is 4 |
| `test_recording_sheet.py` | The `RecordingSheet` test double itself (`tests/recording_sheet.py`) — the mock every Excel-free sheet-writer test is built on |
| `test_within_estimator.py` | v2.1 Fixed Effects phase 2 — the constructor pipeline — the fit-time pair `Design_Response()`/`Design_Columns()` and its stage order, against an independent `statsmodels` LSDV fit |
| `test_group_panel_transforms.py` | v2.1 Fixed Effects phase 1 — `Group_Mean`, `Demean_By`, `Is_Balanced_Panel`, `Absorbed_Degrees_Of_Freedom` |
| `test_df_absorbed_threading.py` | v2.1 Fixed Effects phase 3 — `[DF_Absorbed]` threaded through SE/t/p/CI/MS-Residual/AIC/BIC/AICc, against an independent `statsmodels` LSDV fit |
| `test_group_prediction_interval.py` | v2.1 Fixed Effects phase 5 — `Group_Mean_At`, `Group_Count_At`, `Prediction_Group_Column`, `Group_Prediction_Interval` (the group-mean-recovery CI+PI form), against an explicit LSDV `get_prediction()` reference |
| `test_doc_links.py` | Every relative `](*.md)` link in the repo's markdown resolves to a file that exists, relative to the linking file's own directory |
| `test_doc_catalog_counts.py` | Catalog-function counts stated in prose (README, CONTRIBUTING, ROADMAP) match `lambda_functions.json` — the count half of the documentation-drift proposal, with the phrasings pinned, not sniffed |
| `test_doc_function_names.py` | Every function name the docs *call* (catalog naming convention, inline span or fenced block) resolves to `lambda_functions.json` or one of five pinned exclusion lists — the name half of the documentation-drift proposal, with the exclusions guarded for dead entries and catalog overlap |
| `test_poe_tasks.py` | Pins the `[tool.poe.tasks]` table in `pyproject.toml` so a silent task rename or deletion breaks the unit suite, not the CI workflow or the docs |
| `test_workbook_helpers.py` | `safe_activate()` / `safe_freeze_top_row()` against stub sheet/window objects (headless/no-focus Excel session guards) |
| `test_workbook_builder.py` | Workbook package-patching helpers (`sync_workbook_names` and friends) that don't require Excel |
| `test_build_common.py` | Shared build scaffolding (`lambda_catalog.build_common`: recalculate-and-save calc-mode handling, retry-on-open) that doesn't require Excel |
| `test_build_production.py` | `build_production.py`'s pure-Python logic (Regression-only sheet set, dataset selection, tab order/color, verify forwards `skip_univariate=True`) that doesn't require Excel |
| `test_build_univariate.py` | `build_univariate.py`'s pure-Python logic (Univariate sheet set, Automatic calc mode, the always-on recalc rebuild) that doesn't require Excel — *retained for historical coverage; the script itself is no longer in the build* |
| `test_version_history_writer.py` | `write_sheet_version_history`'s per-artifact version lineage (`artifact="regression"` vs `"univariate"`) and the bad-artifact guard |
| `test_workbook_invariants.py` | Layer 1 headless structural check of a built `.xlsx` package (`zipfile` + `lxml`): dangling defined names, `#REF!`/`#NAME?` cached-value literals, broken package parts, orphan chart-relationship targets, charts reading a sheet other than the one they sit on, sheet drift — see [Verifying builds](#verifying-builds) |
| `test_ln_positive_verification.py` | v2.2 Transform=Log — `Ln_Positive` pure-Python mirror (the `NA()`-exception contract, the geometric-mean round-trip the Prediction Inputs fix relies on) and implementation-shape assertions on the catalog formula |
| `test_transform_threading.py` | v2.2 Transform=Log wiring end to end — cross-checks the new `production_lots_log_transform` QC case (raw columns, `transform="Log"`) against the pre-existing precomputed-log-column case to floating-point precision; Categorical×Log inertness |
| `test_unit_space_dispatch.py` | v3.3 unit-space dispatch (`Smearing_Factor`, `Back_Transform_Response`, the `Unit_Space_*` family) — pure-Python mirrors cross-checked against a NumPy OLS reference, plus catalog `formula_display` shape assertions |
| `test_interaction_wiring.py` | v3.1 interaction wiring — the spec block's M/N pair against the Python mirror in `analyze_regression_spec.build_spec_design`: the three width regimes (1 / L−1 / (L₁−1)(L₂−1)), the closed Product/Difference/Ratio arithmetic, the four operand Role/Include cases, the two-way limit, the documented quadratic, and Ratio's zero-denominator refusal |
| `test_numeric_complete_cases_verification.py` | v3.9 `Numeric_Complete_Cases` — pure-Python mirror (Excel `ISNUMBER` semantics, not pandas coercion) and implementation-shape assertions on the catalog formula |
| `test_categorical_model_construction_verification.py` | v3.9 trio (`Dummy_Column`/`Interact`/`Model_Matrix`) — pure-Python mirrors with Excel broadcasting and `""`/`#N/A` propagation, plus catalog-formula shape assertions (workbook-scoped, document order, none calls `LINEST`) |
| `test_categorical_model_construction_excel.py` | v3.9 trio — Excel COM evaluation of the catalog formulas (gated on `RUN_EXCEL_INTEGRATION=1`); the workbook-backed companion to the pure-Python mirror, reading the spills back after a full recalculation |

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
4. **The suite's growth rate orders the roadmap, under one constraint.** From v3.4 on the ladder sorts first by track — all remaining Regression work before the three milestones that open a new analysis surface (v3.10 Two-sample, v3.11 Resampling, v3.12 Time Series) — and then, within the Regression track, by how much a milestone forces this suite to grow: additive first, axis-wideners last. See [ROADMAP.md § Ladder order](docs/ROADMAP.md#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).

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

One build script produces one workbook. `build_production.py` emits the unified `Lambda_Library.xlsx` — the function library plus the pre-built templates (Regression and Univariate), reference sheets, and data sheets.

### The production artifact

| Target | Produces | Calculation mode | Sheets |
|---|---|---|---|
| `Lambda_Library.xlsx` | The unified workbook | **Automatic** (full) | Regression, Regression Instructions, Diagnostic Guide, Univariate, LAMBDA_functions, Version History, Production Lots, Life Expectancy Data, Mileage Data |

**The workbook carries the complete function library.** All 152 LAMBDA definitions are written into the Name Manager. There is no bundling step, no dependency closure, and no per-function subsetting. When you add a function, it lands in the workbook; there is no list to update.

**The production constructor always runs the full `CalculateFullRebuild` and saves in full Automatic** — there is no skip-calculation flag.

### Which version number moves

One workbook means one version number, plus the library version that tracks the catalog:

| You changed | Moves |
|---|---|
| A LAMBDA definition in `lambda_functions.json` — added, renamed, or different return | **The library version.** It affects the workbook, because the workbook ships the whole catalog. |
| A sheet's layout, input cells, or control block | **The workbook version only** |
| Both | Both |

**A change to a shared function is a shared event.** Every function ships in the workbook, so a catalog edit reaches every user and moves the one library version. A sheet-layout change on one template moves the workbook version only.

The `Breaking?` flag in the Version History sheet attaches to the **workbook** version, never the library version. Full conventions in [ROADMAP.md § Versioning](docs/ROADMAP.md#versioning--release-conventions).

### Production build

```powershell
uv run python scripts/build_production.py
```

Produces `Lambda_Library.xlsx` — the distributable workbook committed to the repo. Writes ten sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions (the library)
- **Life Expectancy Data** — WHO dataset as a structured table; this is one of the datasets the Regression template's `Source_Table` can target (a curated four-driver model: Adult Mortality, Alcohol, percentage expenditure, and `Status`)
- **Mileage Data** — Auto MPG dataset as a structured table (a second sample dataset for the multi-level categorical-encoding demo, and for practicing the Source_Table retarget workflow)
- **Production Lots** — a small unbalanced learning-curve panel (3 facilities, 51 lots) as a structured table; a third sample dataset, and the only one with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year)
- **Regression Instructions** — step-by-step guide for adapting the template to new datasets (reference sheet)
- **Modeling Concepts** — per-feature explainers (Fixed Effects, reference levels, Sequence, Log transforms, interactions, sample filtering, intercept control): the point of each feature, the statistical method it enables, and a use case (reference sheet)
- **Diagnostic Guide** — interpretation guide for regression diagnostics (reference sheet)
- **Univariate** — descriptive statistics, histogram binning, and two-stage MLE distribution fitting (a pre-built template)
- **Version History** — changelog that travels with the workbook
- **Regression** — ToolPak-style analysis interface driven by a declarative variable-specification block (the spec block) and the sheet-scoped names that assemble the design matrix from it. A pre-built template. The wiring names (`Source_Data`, `Header_Names`, `Spec_*`) hardcode the spec block's cell addresses and are defined in `write_spec_block.py` (imported by `write_sheet_regression.py`); the constructor closures (`Sample_Include`, `Response_Column`, `Row_Labels`, `Predictor_Columns`, `Design_Columns`, `Design_Response`, `Constructed_Column_Names`) live in `lambda_functions.json` with `"scope": "Regression"`, so they are the single source of truth and appear on the LAMBDA_functions catalog sheet (Scope column) like any other function — they are just installed on this sheet rather than workbook-wide. Sheet scope is not the same as being a constructor: the row-2 status readouts (`Role_Status`, `Sequence_Status`, `Log_Domain_Status`, `Design_Width_Status`) are sheet-scoped for the same reason — each Regression sheet validates its own spec — but they feed no fit, they only say what is wrong with the specification. A new status cell is a new sheet-scoped catalog entry, not an inline formula in a writer

No test sheets, no OLS analysis, no cache dependency.

**All `build_production.py` options:**

| Flag | Default | What it does |
|---|---|---|
| `--workbook PATH` | `Lambda_Library.xlsx` | Path to the workbook to create or update. |
| `--definitions PATH` | `lambda_functions.json` | Path to the JSON catalog of LAMBDA definitions. |
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV written to the **Life Expectancy Data** sheet. (The **Mileage Data** and **Production Lots** sources are fixed committed sample files with no CLI override.) |
| `--regression-dataset {auto_mpg,life_expectancy,production_lots}` | `life_expectancy` | Which dataset the Regression sheet's `Source_Table` targets, **and** which shipped default spec pre-fills the MODEL SPECIFICATION block (`SPEC_DATASET_PROFILES` in `write_spec_block.py`) — every column starts with a real Role/Include/Type instead of falling back to an un-flagged Predictor. The profile decides which rows arrive pre-filled, not how many spec rows exist: the block sizes itself from `COLUMNS(Source_Data)`, so retargeting `Source_Table` by hand afterwards resizes it too. `life_expectancy` (the default) ships the curated four-driver model both presentation decks headline — `Life expectancy` as Response, `Adult Mortality` + `Alcohol` + `percentage expenditure` as Continuous predictors and `Status` as a Categorical predictor, `Country` as Identifier, `Year` as the Sequence axis (the remaining `FEATURE_COLUMNS` predictors are present with Include off, ready to toggle on). `auto_mpg` targets Mileage Data for a multi-level categorical-encoding demo (MPG ~ Horsepower + Weight + C(Model Year) + C(Origin)). `production_lots` is the one to pick for a ready-made Fixed Effects example (Facility as the FE role, Fiscal_Year as Sequence) — its default spec is the QC-validated Crawford/Wright learning-curve model (`log Unit Cost` ~ `log Cum Units`). |
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

## Verifying builds

There are two verifier layers with different speeds and different scopes. Run them in this order when in doubt; either can be skipped if the other has been run recently.

### Layer 1 — headless structural check

Pure `zipfile` + `lxml` reads of the produced `.xlsx`. Runs in <1 s on Linux CI without Excel. Catches packaging regressions the unit-test suite misses: dangling defined names, `#REF!`/`#NAME?` cached-value literals, broken `[Content_Types].xml`/`workbook.xml.rels`, orphan chart-relationship targets, charts whose references name a sheet other than the one they sit on (SERIES formulas and the `c15:datalabelsRange` "Value From Cells" extension alike), `localSheetId` out of range, sheet drift.

```powershell
poe verify-headless          # includes the committed-artifact checks

# Or directly:
uv run pytest tests/test_workbook_invariants.py -v
```

Every check in this layer is always-on, the committed-artifact ones included: they read `dist/Lambda_Library.xlsx` as a zip, so a stale or hand-edited workbook fails here rather than shipping. If it does fail, rebuild and commit the artifact (`python scripts/build_production.py --verify --no-launch`, needs Excel) — the check is not the thing to relax.

`RUN_EXCEL_INTEGRATION=1` now gates two suites only, the Excel COM checks for the grid-search helpers and the v3.9 categorical & model-construction trio:

```powershell
poe test-excel               # needs desktop Excel; skips without it
```

This is a fast screen. A green run does **not** mean the workbook calculates correctly — that is what Layer 2 is for.

### Layer 2 — spec-driven deep check

Reuses `lambda_catalog.deep_verify.verify_test_sheets` against the production sheets. This is the source of truth for cell-level correctness.

```powershell
# Run the production build, recalculate, then verify against the
# spec oracle. On drift: print a structured VerifyReport, sys.exit(1), and do
# NOT open Excel (so a stale build cannot be launched in place of a fresh one).
# The rebuild is cheap and always runs — it is the source of truth the verifier reads.
python scripts/build_production.py --verify --no-launch

# Same, for the Regression model-case fixture workbook:
python scripts/build_test_models.py --verify --no-launch

# Or, on the just-built workbook without rebuilding:
uv run python tools/verify_workbook.py Lambda_Library.xlsx
uv run python tools/verify_workbook.py Lambda_Library.xlsx --json   # agentic consumption
```

#### Narrower slices of the test-model suite

`verify-test-models` builds all ~50 sheets and runs for minutes. Three tasks cut
it along the axes worth iterating on:

```powershell
poe verify-guards        # every guard state except the oversized L07
poe verify-spec-errors   # every spec-block status line and CF flag
poe verify-models        # the 33 fittable cases, heavy excluded
poe verify-models-rest   # verify-models minus what verify-spec-errors built
```

**The slices overlap, and one containment is total.** Every one of the 17
`GuardStateCase` IDs appears in `verify-spec-errors`' list, so **`verify-guards`
is a strict subset of `verify-spec-errors`** — run one or the other, never both.
`verify-guards` is the fast slice when a guard-state change is all you touched;
`verify-spec-errors` is the one to run when a status cell or CF rule changed,
and it also builds `L07` and three fittable cases (`G08`, `M15`, `L12`, each of
which produces a real fit *and* an error surface).

**Finishing a run you started.** After `verify-spec-errors`, the remainder is
`poe verify-models-rest` — `--kind models` minus those same three IDs, 30
sheets, disjoint from what you already built. The two together cover every
registered case exactly once, heavy excluded;
`tests/test_poe_tasks.py::test_the_two_resumable_slices_partition_the_registry`
derives that exclusion from the registries rather than restating it, so the pair
cannot drift.

**Two of them select structurally, and that is the point.** `--kind guards` and
`--kind models` ask the registry which half a case belongs to, so neither task
carries a list of plan IDs that goes stale the next time a case is registered.
Only `verify-spec-errors` names cases, because *which* cases are an error
surface is a judgement rather than something derivable — and
`tests/test_poe_tasks.py` resolves every ID it names against the live registry,
so it cannot rot into naming a case that no longer exists.

**`L07` is the one documented hole.** Its whole assertion is that a 205-column
model is too wide to build, so the sheet materializes a 206 × 2909 design
matrix and then fails to invert the Gram — minutes of work to confirm a
`WARNING` string. `verify-guards` excludes it; `verify-spec-errors` keeps it,
because `O2` is the width guard's status cell and no other case reaches it.

**`verify-models` is the happy path.** `--include-heavy` stays off, which is
what keeps the runtime sane: `L05` (k = 19 over n = 2117) and `L08` (173 Fixed
Effects groups over 2909 rows) are the two Gram matrices large enough to
dominate a run, and both have Python oracles in the unit suite regardless. Add
`--include-heavy` when the fit arithmetic itself is what changed.

The same selection flags work directly on the script:

```powershell
python scripts/build_test_models.py --verify --no-launch --kind guards
python scripts/build_test_models.py --verify --no-launch --exclude L05,L07,L08
python scripts/build_test_models.py --verify --no-launch --cases L06,L12   # the Log-domain pair
```

`--exclude` is validated against the whole registry, so a typo is an error
rather than a silent no-op that excludes nothing.

**All `tools/verify_workbook.py` options** (positional `workbook` path is required):

| Flag | Default | What it does |
|---|---|---|
| `--csv PATH` | `sample_data/Life Expectancy Data.csv` | Life Expectancy CSV used for the `Developed Country after 2013` derived-column comparison. |
| `--mileage PATH` | `sample_data/auto_mpg_data.csv` | Auto MPG CSV for the Regression spec oracle (the Regression sheet's `Source_Table` default). |
| `--json` | off | Emit the report as JSON (stable schema, for agentic consumption) instead of the human-readable form. |
| `--verbose` | off | Print per-phase checkpoints from the spec-driven verifier. |

The `VerifyReport` (in `lambda_catalog/verify_report.py`) is emitted both in a human-readable form (`Verify: passed (spec mode, …)` / `ERROR Verify mismatch totals: …`) and as JSON (stable schema: `passed`, `categories`, `failures`, `elapsed_seconds`, `mode`, `workbook`).

Expected terminal flow for the one-shot command:

1. Sheet update summary (`Sheet updated: ...`, `Created names: ...`, `Updated names: ...`).
2. Spec verifier result (`Verify: passed ...` or `ERROR Verify mismatch totals: ...`).
3. Timing summary lines (`Timing: build+sync`, `Timing: recalculate` or `skipped`, `Timing: verify`, `Timing: total`).


### How a comparison is scored — the comparison-scale convention

The verifiers score each compared value with `first_digit_deviation`, which
returns the DECIMAL PLACE where the oracle and the sheet first differ, and
fail anything at or below `TOLERANCE_DECIMALS`. Because that is an absolute
measure, it silently gets stricter as a statistic's magnitude grows: six
decimal places on a quantity of order 70 is a demand for eight or nine
significant figures. The convention that keeps it meaningful is:

> **Score a statistic against the magnitude its error comes from, not against
> its own.**

Every compared statistic falls into one of three cases, all declared in
`lambda_catalog/regression_spec_sheet_io.py`:

| Statistic's error tracks... | Divisor | Declared in |
|---|---|---|
| its own value | `max(\|expected\|, 1.0)` | `SCALE_FREE_STATS` |
| the fitted values, in response units | response RMS | `_RESPONSE_UNIT_STATS` |
| the fitted values, over `SE_Regression` | response RMS / `SE_Regression` | `_STANDARDIZED_RESIDUAL_STATS` |
| nothing larger than itself | none — absolute | (default) |

`compare_values` takes an explicit `scale` for the response-derived cases. It
divides BOTH sides by the same factor, so the relative difference is untouched
and a genuinely wrong number still fails; and it floors every divisor at 1.0,
so the convention can only loosen a comparison, never tighten one.

**Two rules when you add a compared statistic:**

1. **Choose its case deliberately.** The residual band is the inherited-error
   one — every statistic there is built from the predictions, so it carries the
   response's absolute precision floor whatever its own magnitude. A residual is
   the clearest example: it is the difference of two response-sized numbers, so
   it can be of order 0.1 while carrying the error of numbers of order 70, and
   self-scaling divides it by 1.0 and changes nothing.
2. **Derive the divisor from the fit, never from a constant.** A response in the
   tens and one in the billions must be treated proportionately, and one below
   1.0 must get no adjustment at all.

`T_Statistics` is deliberately in none of the response-derived sets: it is
dimensionless and O(1) and its error comes from the coefficient (relative error
on the order of `eps · cond(X)`), so a response-derived divisor is not a scale
it has. Where a t-statistic and the sheet disagree, the design matrix's
conditioning is what widened it, and conditioning is where it is addressed —
not the tolerance.

Related: the OLS oracle is pinned to `method="qr"` in `_fit_ols_model` rather
than the statsmodels `"pinv"` default. The oracle is the reference the workbook
is scored against, so it should be the more accurate side by as wide a margin
as the choice allows; LINEST is QR-based too, so the two now solve the same
problem the same way. Full rationale for all of this: `docs/DECISIONS.md` →
*QC comparison scale, the clear-list invariant, and the OLS solver*.

### `poe verify`

```powershell
poe verify                  # both layers, plus the test-model suite
poe verify-headless         # Layer 1 only (any platform)
poe verify-deep             # Layer 2, the workbook (needs Excel)
poe verify-test-models      # Layer 2, the ~50-sheet test-model suite (needs Excel)
```

**`poe verify` runs the two builds concurrently, then screens their output.** It stops at the first stage that exits non-zero, and the stage boundary is what makes the order matter: `verify-headless` reads whatever `.xlsx` files are sitting in `dist/`, and the deep tasks *rewrite* those files. The task used to run the screen first, which meant it validated the previously committed artifacts and never looked at the ones the run had just built — a rebuild that broke a defined name or orphaned a chart relationship passed `verify` clean. Builds first, screen last.

The two builds overlap safely because they share nothing: each driver opens its own `xw.App(visible=False, add_book=False)` and reaches every workbook through that instance's `app.books` handle (there is no bare `xw.Book()` or `xw.apps.active` anywhere in the package), and they write two different artifacts and two differently-named transcripts. The one file they could have contended over is `templates/static_sheets.xlsx`, which `copy_static_sheet` opens read-only. Output is buffered per task rather than interleaved, so each transcript stays contiguous.

Wall time becomes roughly the longest build — `verify-test-models`, minutes — instead of the sum of both. The cost is two Excel instances competing for CPU; on a constrained machine, run the `verify-*` tasks one at a time instead. **None of this is checkable in CI** (no GitHub-hosted runner has Office), so a change to the `verify` task needs a developer-machine run archived to `excel-only-runs/`.

`poe verify-deep` shells out to `build_production.py --verify --no-launch`, so it both rebuilds and verifies the workbook. It tees its run into [`excel-only-runs/`](excel-only-runs/) (`<script> <flags>.log`, via `lambda_catalog.build_common.run_log_path`) — stderr and any traceback included — so a failed deep verify is a file you can commit and hand over rather than terminal scrollback; override the destination with `--log PATH`. To verify an already-built workbook, use `python tools/verify_workbook.py <workbook>` instead.

The `verify-test-models` task passes `--verbose` because that run takes minutes across ~50 sheets: it names each sheet *before* writing it, so an interrupted run leaves the offending case on screen. It archives its transcript the same way — all three Excel-required drivers do. The heavy `L08` case is excluded by default; append `--include-heavy` (`poe verify-test-models --include-heavy`) to include it. Its Python oracle runs in the unit suite regardless.

### CI

GitHub Actions runs the unit-test suite on Python 3.10–3.13 (Ubuntu) on every push and pull request via `.github/workflows/ci.yml`. Its two check steps invoke the poe tasks (`uv run poe test-cov`, `uv run poe lint`) rather than spelling out their own pytest and pylint invocations, so CI and a developer machine cannot drift apart — changing a task in `pyproject.toml` changes both. The spec-driven verifier (Layer 2) is **not** run in CI: the GitHub-hosted `windows-latest` runner image does not include Microsoft Office, so xlwings fails to dispatch `Excel.Application` (`pywintypes.com_error: (-2147221005, 'Invalid class string')`). Until a self-hosted runner with Office is wired in, Layer 2 must be run on a developer machine (or any Windows box with desktop Excel) — the agentic workflow runs it before pushing. The `windows-verify` job was removed for that reason; see the comment block at the bottom of `ci.yml`. Layer 1 (the headless `tests/test_workbook_invariants.py` suite) needs no Excel and is auto-discovered by the existing Linux job, so it runs on every push and pull request.

## File structure

```
scripts/
  build_production.py        # production entry point → dist/Lambda_Library.xlsx
  build_test_models.py       # Regression model-case fixture builder → Lambda_Library_TestModels.xlsx
                              # (gitignored; verify with --verify --no-launch)
  rebuild_static_sheets.py   # regenerates templates/static_sheets.xlsx from its Python source —
                              # see "Static reference sheets" below
dist/                        # the shipped .xlsx artifact (build output, committed)
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
  build_common.py            # shared build scaffolding (retry-on-open, recalculate-and-save) for the production script
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

`write_sheet_csv_dataset.py` is a single config-driven module backing all three sample datasets (Life Expectancy, Mileage, Production Lots) — one `CsvDatasetConfig` per dataset (`LIFE_EXPECTANCY`, `MILEAGE`, `PRODUCTION_LOTS`) captures its sheet/table name, an optional appended derived column (Life Expectancy ships `Developed Country after 2013` = `AND([@Status]="Developed",[@Year]>2013)`; Mileage and Production Lots ship none), and CSV-parsing quirks (missing-value markers, header normalization), and the shared `load_csv_rows` / `write_csv_dataset_sheet` functions do the actual work. All three CSVs are real committed sample files (`sample_data/Life Expectancy Data.csv`, `sample_data/auto_mpg_data.csv`, `sample_data/production_lots.csv`) that can be pointed at a different CSV via `--csv`, or via `build_production.py`'s `--csv` / `--mileage-csv` / `--production-lots-csv` flags. The loader (`load_csv_rows`) has no Excel dependency, so the Python QC oracle (`analyze_life_expectancy.calculate_developed_country_flags`) can run on any platform.

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

Three mechanical checks would catch most of this class. All are pure Python and need no Excel, so they run in the existing Linux CI job. **All three are built:**

1. **Link targets — built, `tests/test_doc_links.py`.** Every relative `](target.md)` link resolves to a file that exists, relative to the *linking file's own directory*. This is the check that would have caught two real breakages: the deletion of `REVIEW.md` while four documents still linked to it, and a docs-reorganization pass that prefixed every relative link with `docs/` — including links already inside `docs/`, where the prefix is one level too many. `docs/TODOs.md` → `docs/ROADMAP.md` resolves to `docs/docs/ROADMAP.md`; that single commit broke 171 links and nothing failed.
2. **Cross-document anchors — built, same file.** Every `](target.md#anchor)` and `](#anchor)` resolves to a heading that exists in the target file, matched by reproducing GitHub's slug rules. Heading renames silently break these: the `ARCHITECTURE.md` §4 renames from `(A–L)` to `(A–N)` (v2.1, adding the Sequence Period / Period In Use pair) and from `(A–N)` to `(A–O)` (v3.0, adding the Design Columns audit column) each broke at least one `ROADMAP.md` link this way. It found one live break the moment it was written — retitling *this very section* from "Documentation drift (proposed check — not yet implemented)" broke the `TODOs.md` link into it, in the same commit that built check 1.
3. **Function names — built: the count half in `tests/test_doc_catalog_counts.py`, the name half in `tests/test_doc_function_names.py`.** The count half asserts that every *count* of catalog functions stated in README / CONTRIBUTING / ROADMAP matches `lambda_functions.json` — total, workbook-scoped, and Regression-scoped; it found four stale numbers on its first run. The name half asserts that every function name the docs *call* — written as `Name(...)`, in the catalog's naming convention, inside an inline span or a fenced block — resolves to a catalog entry, a native Excel function, a name the planning docs tag as planned, one of the v3.2 spill readers, a retired name the shipped-changelog prose still cites in rename history, or a pinned doc shorthand. That is the shape that catches a stale call: the `X_s` references the 2026-08-03 review had to find by hand. Its blind spots are deliberate and guarded — bare backticked names (undistinguishable from named ranges, Roles, and Python constants) and CamelCase single words like `Interact` (the shape the docs share with prose) are not candidates, and a test derives that blind-spot boundary from the catalog so it cannot grow silently.

**The built checks pin their inputs rather than sniffing them,** and they will fail when the docs grow a phrasing, a doc page, or a native-function call they do not know — that is the intended bargain, the same one `_EXPECTED_TASK_NAMES` and `_EXPECTED_CASE_NAMES` make. The count check matches three exact phrasings instead of "a number near the word LAMBDA", because README's first line reads *"Excel 365 LAMBDA functions…"* and a looser rule would fail the build on the Office release number. The name check pins its policed-doc list, its five exclusion lists, and guards each list against rot: a planned name must still appear in ROADMAP/TODOs, an exclusion entry that stops matching a candidate fails as dead, no exclusion may shadow a catalog function (a planned name that ships must come off the list in the same commit), and a new markdown file must be consciously classified as policed or excluded.

## Adding a new LAMBDA function

1. Add an entry to `lambda_functions.json` with `name`, `formula_display`, `arguments`, `yields`, `description`, and `plain_language_summary`. Add `notes` for the Name Manager tooltip (255 characters max), and `scope` only when the function is a sheet-scoped closure rather than a portable workbook name.
2. Add or update the relevant Python oracle when the function feeds a production analysis surface (for example, Regression outputs in `analyze_regression_sheet.py` / `analyze_regression_spec.py` or Univariate outputs in `analyze_univariate.py`).
3. Run the appropriate verifier (`poe verify-headless`, `python scripts/build_production.py --verify --no-launch`, and/or `python scripts/build_test_models.py --verify --no-launch`) and confirm no unexpected WARNING lines appear.
4. Run `python scripts/build_production.py` to rebuild the distributable that carries the function.
5. Move the **library version**, not a workbook version — a new function ships through the catalog. See [Which version number moves](#which-version-number-moves).

## Cell styling

All cell colors are defined once in `lambda_catalog/sheet_styles.py` and imported by every sheet writer. Never hard-code RGB tuples directly in a sheet writer. The constant table, the import pattern, and the section-heading convention are in [CLAUDE.md](CLAUDE.md) / [AGENTS.md](AGENTS.md) at the project-instructions tier — consult that for the canonical definitions. Sheet-specific colors that differ from the shared palette (e.g., `_SUBHEADER_COLOR` in `write_sheet_diagnostic_guide.py`) remain as local constants in the relevant file.

## Regression sheet conventions

### Chart series data ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns (`$AI$3:$AI$1048576`) degrades Excel's recalculation performance and can crash the workbook on large datasets.

Instead, all chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in `$AB$9` (the `Observations` cell in the Regression Outputs zone):

```python
sheet.api.Names.Add(
    Name="RegChartFitY",
    RefersTo=f"=OFFSET('{sname}'!$AP$3,1,0,MAX(IFERROR('{sname}'!$AB$9,1),1),1)",
)
```

This starts one row below the column header (row 3) and extends exactly `$AB$9` rows — the number of filtered observations. The `MAX(IFERROR(...,1),1)` guard keeps the range one row tall (instead of erroring) when `$AB$9` cannot resolve. Each name also carries a Name Manager `Comment` identifying the chart it feeds — see the loop in `_setup_local_names`.

**Naming convention** — all OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix, distinguishing them from the constructor closures (`Design_Columns`, `Sample_Include`, etc.) and formula-helper names:

| Name | Column | Contents |
|---|---|---|
| `RegChartQQX` | AU | Normal Scores Ranked (QQ theoretical axis) |
| `RegChartQQY` | AV | Studentized Residuals Ranked (QQ actual axis) |
| `RegChartFitY` | AP | Predicted Y — shared by multiple charts |
| `RegChartResid` | AQ | Residuals |
| `RegChartActY` | AO | Actual Y |
| `RegChartScaleLoc` | AW | Scale-Location |
| `RegChartCookDist` | AT | Cook's Distance |
| `RegChartLeverage` | AR | Hat Diagonal |
| `RegChartStudResid` | AS | Studentized Residuals |
| `RegChartPRESSResid` | AX | PRESS Residual — the leave-one-out (LOOCV) residual; there is no separate "LOOCV Residual" column |
| `RegChartCookDistFlag` | AY | Cook's Distance, `""` below the `F.INV(0.5, p, n-p)` influence cutoff — feeds the Cook's Distance chart's data-label overlay series, whose labels read this range through **Value From Cells** (which is why the mask is `""` and not `NA()`: the label would print a literal `#N/A`) |
| `RegChartObsLabel` | AN | Row identifier (Row_Labels()) — the flagged-point overlay series' `XValues`, i.e. its categories |

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

Anything workbook-scoped and outside the catalog is residue from an earlier build. The sheet-scoped originals were correct the whole time — only stale workbook-scoped copies were broken, which is why the workbook still rendered its charts. (The v3.0 split briefly produced cross-artifact residue when there were two workbooks; reunification eliminated that class of issue.)

**A catalog function that names a missing worksheet is skipped, not written.** Excel does not leave such a reference unresolved: it rebinds it to an external workbook (`Regression!Source_Data` becomes `[1]!Source_Data`), writes an `xlExternalLinkPath/xlPathMissing` external-link part, and prompts about broken links every time the file is opened. No catalog function is sheet-qualified today. `Base_Period_Delta` was the one — its body read `'Regression'!Source_Data` / `Spec_Sequence` / `Spec_Sequence_Period`, so the standalone Univariate artifact did not carry it — and it is now **sheet-scoped** with unqualified references, one definition per Regression-shaped sheet. The guard stays because it is what stops the next sheet-qualified body shipping a broken link. The build prints `Skipped names: …` when it happens. When the last external reference goes, the orphaned external-link parts, relationships and content-type overrides are stripped with it.

Keep a new workbook-scoped catalog LAMBDA sheet-agnostic unless it is deliberately Regression-only (sheet-scoped). If it must read the spec block, it should be sheet-scoped to `Regression` rather than workbook-scoped.

**Checks and repair.** `tests/test_workbook_invariants.py::TestRealWorkbookNameScope` asserts workbook-scope ownership, error-free defined-name bodies, and the absence of external links against the committed artifact on every commit — pure zipfile + lxml, so it runs in CI without Excel. To re-apply the cleanup to a built artifact without a full rebuild (also Excel-free):

```bash
python tools/resync_workbook_names.py Lambda_Library.xlsx
```

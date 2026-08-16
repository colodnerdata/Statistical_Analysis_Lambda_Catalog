# Reunify the Workbook and Refactor Sheet Creation

**Branch:** `refactor/reunify-workbook` (merged; follow-on work lands off `main`)
**Date:** 2026-08-08
**Status:** structural work complete (Parts 1–5); Part 6 formula cleanup outstanding —
see the table below

| Part | State | Landed as |
|---|---|---|
| 1.1 Extract `regression_layout.py` | **Shipped** | `1fe76b5` |
| 1.2 Extract `spec_layout.py` | **Shipped** | #209 |
| 2.1 Extract `spec_dataset_profiles.py` | **Shipped** | #210 |
| 3 Reunify the build scripts | **Shipped** | #211, `fe4fddf` |
| 4 Update documentation | **Shipped** | `836da11`, `fe4fddf` |
| 5.1 Extract `regression_charts.py` | **Shipped** | #214 |
| 5.2 Extract `regression_materialization.py` | **Shipped** | #220 — writer 2,441 → 2,098 lines |
| 6.1 Audit the inline cell formulas | **Shipped** | `.hermes/plans/lambda-extraction-audit.md` |
| 6.2 Extract the status-cell LAMBDAs | **Open** | catalog is still 141 = 123 workbook + 18 Regression-scoped |
| 6.3 Split the prediction-interval VSTACK | **Open** | |
| 6.4 Prediction-input prefill spill | **Open** | follow-on, as scoped |

Cleanup the reunification itself needed, tracked here because it is not a numbered
Part: the merge committed a hand-edited `dist/Lambda_Library.xlsx` (15,369 cached error
literals) and it survived eight merges because `TestRealWorkbook` was gated behind
`RUN_EXCEL_INTEGRATION=1`. The artifact was rebuilt in #218; the gate is removed and the
last two-artifact leftovers retired in the PR that added this table.

**Before starting 6.2**, settle one thing the plan does not: the O2 thresholds
(`_DESIGN_MATRIX_MAX_COLUMNS` and the soft pair) live in `regression_layout.py` and also
size the sheet, so moving that formula into `lambda_functions.json` duplicates them. It
wants a JSON↔Python pin test in the style of
`test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform`.

## Goal

One workbook, one build script, one download. The two-artifact split (Regression vs. Univariate) was driven by a Data Tables calculation-mode conflict that no longer exists -- Weibull/Gamma use 1-D profile searches and Beta uses `Full_Factorial` spills, so both sheets run in full Automatic with no Data Tables anywhere. Reunify into a single `Lambda_Library.xlsx` that carries all sheets, and simultaneously decompose the two monolithic build scripts and the two largest sheet-writer modules into focused, navigable files.

This is the technical foundation for a v1.0 relaunch as SALTbox (Statistical Analysis Lambda Toolbox) in a fresh repo.

## Nomenclature

Adopt **library + templates** as the project's vocabulary, woven into the existing language rather than replacing it:

- **The library** is the 141 LAMBDA-function catalog. It is workbook-scoped, identical across all sheets, and shared by every analysis.
- **A template** is a pre-built worksheet that drives a subset of the library. The Regression sheet is a template. The Univariate Analysis sheet is a template. A future simplified regression sheet (a stripped-down spec block for teaching or quick fits) would also be a template. The data sheets, instructions, and diagnostic guide are reference sheets, not templates.

The terminology coexists with "sheet" and "workbook" -- a template is a worksheet, the library lives in a workbook. The distinction is conceptual (what role does this sheet play?) not structural (it's all sheets in one workbook). Use "template" when describing the role and "sheet" when describing the mechanics. This frames what the workbook *is*: a function library plus a collection of analysis templates and reference guidance. Adding a new analysis surface means adding a template, not a new workbook. The split between "library" and "template" is also the split between workbook scope and sheet scope (see Design Principle below).

## Design Principle: sheet-scoped templates, library-scoped functions

A workbook can contain **multiple instances of the same template** -- two Regression sheets for model comparison, three Univariate sheets for different variables. This works because:

- The **LAMBDA functions** are **workbook-scoped**. There is one `R_Squared`, one `Absorb_Two_Way_Fixed_Effects`, one `NLL_Beta` -- every sheet calls the same function. The catalog is the library, shared by all templates.
- The **template wiring** is **sheet-scoped**. The constructor closures (`Predictor_Columns`, `Design_Columns`, `Sample_Include`, `Model_Formula`, `Fit_Context`, the `RegChart*` named ranges, the `UV_*` named ranges) are defined per-sheet via `sheet.api.Names.Add`, so each Regression sheet gets its own spec block, its own design matrix, its own model context. The Univariate sheet gets its own `UV_Data`, `UV_Include`, `GoF_AIC`. A workbook with three Regression sheets has three independent `Predictor_Columns` names, one per sheet, and a change on one sheet does not recalculate the other two.

This is why the `sync_workbook_names` enforcement exists: workbook scope belongs exclusively to the catalog's LAMBDA functions, and sheet scope belongs to the template wiring. The two cannot mix, because a workbook-scoped template name in a multi-template workbook resolves to whichever sheet was literally named, which is the wrong model.

The reunification must preserve this: the merged workbook's Univariate sheet defines its own `UV_*` sheet-scoped names, the Regression sheet defines its own `RegChart*` and constructor names, and adding a second Regression sheet (copy-sheet in Excel) carries its own wiring automatically. Nothing in the merge changes this -- it is already how both sheets work -- but the documentation and nomenclature should make it explicit so a future contributor adding a template understands the pattern.

## Design Principle: plain-language LAMBDA functions, not in-cell formula soup

The project's philosophy is formula transparency and auditability: "any result can be interrogated by clicking the cell." A cell that contains a 200-character nested `IF(IF(IF(...)))` or a multi-line `LET(...,IF(...,IF(...)))` is not auditable -- a user clicking it sees an opaque wall of text. The current sheet has several such cells:

- **Role cardinality status (B2):** `=IF(response_count=0,"ERROR: ...",IF(response_count>1,"ERROR: ...",IF(fe_count>1,"ERROR: ...","")))` -- a triple-nested IF that should be a single LAMBDA call, e.g. `=Role_Status()`.
- **Sequence cardinality status (H2):** `=IF(seq_count>1,"ERROR: ...","")` -- simpler, but still an inline conditional that should be `=Sequence_Status()`.
- **Width guard (O2):** `=LET(k,...,n,...,IF(k>max,"ERROR: ...",IF(OR(k>soft,n*k>soft_cells),"WARNING: ...","")))` -- a LET with nested IFs computing both the threshold and the message text inline. Should be `=Design_Width_Status()`.
- **Prediction interval (AK):** `=IF(...,LET(raw,TAKE(...),trn,...,pred_input,IF(trn="Log",Ln_Positive(raw),raw),Group_Prediction_Interval(...)))` -- a cell that both constructs transformed prediction inputs and calls the PI function. The transform dispatch should be inside a LAMBDA, not in the cell.
- **Prediction input prefill (AK rows):** `=IF(ROW()-offset<=IFERROR(ROWS(means#),0),INDEX(means#,ROW()-offset),"")` -- a per-row formula that should be a single `Prediction_Input_Means()` spill or a LAMBDA that the cell calls.

**The rule for this branch:** complex inline conditionals in cells should be extracted into named LAMBDA functions in `lambda_functions.json` that express their intent in plain language. A cell should read `=Role_Status()` or `=Design_Width_Status()`, not contain the logic. This makes the cell auditable (clicking it shows a readable function name) and the logic reusable (a second Regression sheet calls the same function against its own sheet-scoped context).

This is a significant body of work and not all of it needs to land in this branch. The reunification and structural refactoring (Parts 1-4) come first. Formula cleanup (Part 6) is scoped as a follow-on that can proceed once the structure is clean.

## Scope

In scope:
- Merge `build_univariate.py` into `build_production.py` (one build, one artifact)
- Extract shared build scaffolding into `lambda_catalog/build_common.py` (already partially done)
- Extract layout constants from `write_sheet_regression.py` (3,227 lines) and `write_spec_block.py` (2,455 lines) into focused modules
- Extract `SpecDatasetProfile` (705 lines of data) from `write_spec_block.py`
- Update tests, poe tasks, and documentation to reflect one artifact
- Audit and catalog the complex inline cell formulas that should become LAMBDA functions (Part 6 scoping, not necessarily full implementation in this branch)

Out of scope (for this branch):
- The SALTbox rename / new repo creation (separate step after this lands)
- Further decomposition of the zone writers inside `write_sheet_regression.py` (can be a follow-up)
- GitHub Releases / Pages setup (separate effort)
- Full implementation of all extracted LAMBDA functions (Part 6 is scoped, not fully implemented in this branch -- but the audit and the first few extractions should land)

## Part 1 -- Extract layout constants

**Why first:** The layout constants (_C_*, _ROW_*, _A_*, _ZONES, etc.) are already a shared contract -- 6 modules import them from `write_sheet_regression` and 10+ from `write_spec_block`. Extracting them first makes the subsequent build-script merge cleaner (fewer names to reason about) and immediately benefits every reader of the file. No behavior change, no Excel needed, fully testable.

### 1.1 -- Extract `regression_layout.py`

Create `lambda_catalog/regression_layout.py` containing the layout constants from `write_sheet_regression.py` lines 221-738:
- Column constants: `_C_A` through `_C_BA`, `_C_BB`, `_LAST_CHART_COLUMN`, materialization zone columns
- Row constants: `_ROW_DATA_FIRST`, `_ROW_ADJUSTED_R_SQUARED`, `_ROW_OBSERVATIONS`, etc.
- Anchor references: `_A_ALPHA`, `_A_OBSERVATIONS`, `_A_SIGNIFICANCE_F`, etc.
- Zone/gap/width tables: `_ZONES`, `_GAP_COLUMNS`, `_COLUMN_GROUPS`, `_COLUMN_WIDTHS`
- Chart constants: `_XL_*`, `_CHART_*`, `_CHART_Y_TICK_FORMATS`, chart label columns
- Materialization constants: `_MODEL_CONTEXT_ELEMENTS`, `ModelContextElement`, `_MATERIALIZATION_*`
- Helpers: `_abs_ref(row, col)`, `_band(col, first_row)`
- The `_COOKS_CUTOFF` formula and `_BACK_TRANSFORM_METHODS` / `_BACK_TRANSFORM_DEFAULT`
- `REGRESSION_SHEET_NAME`

`write_sheet_regression.py` imports everything from `regression_layout` (explicit re-imports or `from .regression_layout import *`). All existing external importers (`analyze_regression_spec_block`, `regression_spec_sheet_io`, `write_sheet_test_model`, `analyze_regression_guard_states`, `build_demo_workbook`, tests) keep working because `write_sheet_regression` re-exports the names it imported.

**Verification:** `uv run pytest tests/test_sheet_writers.py tests/test_test_model_sheets.py tests/test_analyze_regression_spec_block.py tests/test_regression_guard_states.py` -- no behavior change, all pass.

### 1.2 -- Extract `spec_layout.py`

Create `lambda_catalog/spec_layout.py` containing the constants from `write_spec_block.py` lines 1-378 and 543-652:
- Column/row constants: `_C_SPEC_*`, `_FIRST_DATA_ROW`, `_LAST_DATA_ROW`, `_HEADER_ROW`, etc.
- Role/type/transform definitions: `_ROLE_*`, `_TYPE_*`, `_TRANSFORM_*`, `_INTERACTION_*`
- Validation lists: `_ROLE_VALIDATION_LIST`, `_TYPE_VALIDATION_LIST`, etc.
- Note text: `_RESERVED_NOTE`, `_TRANSFORM_NOTE`, `_SEQUENCE_NOTE`, etc.
- `_is_log(expr)` helper
- `SpecVariable` (the NamedTuple used by every case builder)
- `ExtraSpecColumn`

`write_spec_block.py` imports from `spec_layout` and re-exports for compatibility. All existing importers keep working.

**Verification:** `uv run pytest tests/test_spec_block_writer.py tests/test_regression_spec_qc.py tests/test_interaction_wiring.py` -- all pass.

## Part 2 -- Extract SpecDatasetProfile

**Why:** `SpecDatasetProfile` + its 3 dataset profiles (auto_mpg, life_expectancy, production_lots) is 705 lines of data, not logic. It has no dependency on the writer.

### 2.1 -- Create `spec_dataset_profiles.py`

Move `SpecDatasetProfile`, `_AUTO_MPG_PROFILE`, `_LIFE_EXPECTANCY_PROFILE`, `_PRODUCTION_LOTS_PROFILE`, `SPEC_DATASET_PROFILES`, `LIFE_EXPECTANCY`, `MILEAGE`, `PRODUCTION_LOTS` from `write_spec_block.py` lines 653-1357 into `lambda_catalog/spec_dataset_profiles.py`.

`write_spec_block.py` imports from `spec_dataset_profiles` and re-exports `SPEC_DATASET_PROFILES` and the three profile constants for compatibility. Importers: `build_production.py`, `build_demo_workbook.py`, `write_sheet_test_model.py`, `analyze_regression_spec.py`, `analyze_model_construction.py`, tests.

**Verification:** `uv run pytest tests/test_spec_block_writer.py tests/test_regression_spec_qc.py tests/test_test_model_sheets.py` -- all pass.

## Part 3 -- Reunify the build scripts

**Why:** The two build scripts share ~80% of their scaffolding. Merging eliminates the duplication and produces one artifact. The layout constant extraction (Part 1) makes the merged script cleaner because the sheet-writing functions are already the same -- only the orchestrator differs.

### 3.1 -- Merge `build_univariate.py` into `build_production.py`

The merged `build_production.py` writes all sheets in one workbook:

```
Sheet order (tab order, left to right):
1. Regression                    -- the flagship template
2. Regression Instructions       -- static reference
3. Diagnostic Guide              -- static reference
4. Univariate                    -- the distribution-fitting template
5. LAMBDA_functions              -- the function catalog (the library reference)
6. Version History               -- changelog that travels with the workbook
7. Production Lots               -- dataset (FE panel)
8. Life Expectancy Data          -- dataset (curated default for Regression)
9. Mileage Data                  -- dataset (Auto MPG, default for Univariate)
```

This order puts the analysis templates first (Regression, Univariate), then the reference sheets (Instructions, Diagnostic Guide, LAMBDA_functions, Version History), then the raw data sheets at the end. A user opening the workbook lands on the Regression template -- the flagship -- and the data they'd edit is at the back where it belongs.

Tab ordering and styling: update `_reorder_and_style_sheet_tabs` to include the Univariate sheet in its position.

Changes to `build_production_workbook()`:
- Add `write_univariate_sheet(workbook, document.univariate_sheet_notes, beta_grid_size=beta_grid_size)` to the sheet-write sequence
- Add `beta_grid_size` parameter (default 10)
- Add `UNIVARIATE_SHEET_NAME` to imports from `write_sheet_univariate`
- Add the Univariate sheet to the tab reorder/color map
- Update the `--verify` path: `deep_verify.verify_test_sheets(...)` no longer passes `skip_univariate=True` or `skip_regression=True` -- both sheets are present
- Remove `regression_dataset` parameter? No -- keep it. The Regression sheet still targets one dataset by default. The Univariate sheet always reads Life Expectancy Data.

Delete `scripts/build_univariate.py`.

### 3.2 -- Update `deep_verify.py`

Remove the `skip_univariate` flag (or keep it as a no-op for backward compat, but it should no longer be passed). The verifier now checks both Regression and Univariate sheets in the same workbook. `skip_regression` stays -- it's still useful for verifying a test-model artifact that has no Regression sheet.

### 3.3 -- Update `build_common.py` docstrings

`build_common.py` line 1-5 says "These helpers were extracted from build_production.py when the project split into two build scripts." Update to reflect one build script.

### 3.4 -- Update `tests/test_build_univariate.py`

Fold the Univariate-specific build tests into `tests/test_build_production.py`:
- The sheet-set assertion now includes `UNIVARIATE_SHEET_NAME`
- The default output path is `Lambda_Library.xlsx` (not `Lambda_Library_Univariate.xlsx`)
- The calc-mode assertion (full Automatic) stays the same
- The `_TARGET_SHEET_NAMES` assertion expands to include Univariate

Delete `tests/test_build_univariate.py`.

### 3.5 -- Update poe tasks in `pyproject.toml`

Remove:
- `build-univariate` task
- `verify-deep-univariate` task

Update:
- `build` task: now builds the single workbook (no change needed -- it already runs `build_production.py`)
- `verify-deep` task: now verifies both Regression + Univariate (no change needed -- `build_production.py --verify` does both)
- `verify` task: remove `verify-deep-univariate` from the parallel list
- `build-presentation` task: remove `build-univariate` from the sequence
- `test_poe_tasks.py`: remove references to `verify-deep-univariate` and `build-univariate`

## Part 4 -- Update documentation

Every markdown file that references the two-artifact split needs updating. This is the largest part by line count but the lowest risk -- it's prose, not code.

### 4.1 -- Files to update

| File | What changes |
|---|---|
| `README.md` | "Which workbook do I want?" section -> one workbook. Remove `Lambda_Library_Univariate.xlsx` references. Sheet inventory merges. |
| `AGENTS.md` / `CLAUDE.md` | Remove `build_univariate.py` from verify instructions. Update static-sheet reference (no longer mentions build_univariate). |
| `CONTRIBUTING.md` | Remove "Univariate build" section. Merge sheet tables. Remove `build-univariate` / `verify-deep-univariate` from poe task table. Update file tree. Update typical loop. |
| `docs/ROADMAP.md` | Remove the two-number versioning scheme (one workbook = one number). Remove the "Univariate artifact" row. Update v3.0 entry. |
| `docs/DECISIONS.md` | Mark the "Univariate becomes its own workbook" decision as **superseded** (with rationale: Data Tables removed, calculation-mode conflict gone). Update the versioning-across-two-artifacts entry. Preserve the full history -- the new repo decision is deferred, so the changelog stays intact for now. |
| `docs/ARCHITECTURE.md` | Remove two-artifact references. |
| `docs/MODEL_TESTING_ASSETS.md` | No change expected (test-model suite is unaffected). |
| `docs/TODOs.md` | Update any references to the Univariate artifact. |
| `lambda_catalog/build_common.py` | Update docstring (already in 3.3). |
| `tools/verify_workbook.py` | Keep `--skip-regression` (still useful for test-model artifact). Remove or no-op `--skip-univariate` if present. |

### 4.2 -- What NOT to delete

Keep `docs/DECISIONS.md` entries for the split and the two-number scheme -- mark them superseded with a date and rationale. The history of *why* the split existed and *why* it was reversed is valuable for future contributors. Don't erase the decision; annotate it.

### 4.3 -- Version History

The Version History sheet currently carries two lineages (one per artifact). After reunification there is one workbook, so one Version History. **Keep the full changelog intact** -- the new repo / fresh v1.0 decision is deferred, so we preserve all history in this branch. When the SALTbox relaunch happens later, the new repo can start fresh; this branch keeps everything so nothing is lost in the meantime.

## Part 5 -- Extract charts and materialization (optional, can defer)

If the 3,227-line `write_sheet_regression.py` still feels too large after Parts 1-2 remove ~900 lines of constants:

### 5.1 -- Extract `regression_charts.py`

Move `_diagnostic_chart_specs`, `_write_chart_label_cells`, `_write_diagnostic_charts`, and the chart-related constants to `lambda_catalog/regression_charts.py`. These are 278 lines (lines 2402-2519, 2865-3024) plus their chart constants. They depend only on the layout constants (now in `regression_layout.py`) and the sheet object.

### 5.2 -- Extract `regression_materialization.py`

Move `_write_materialization_zone` (345 lines, lines 2520-2864) to its own module. It depends on layout constants, `ModelContextElement`, and the sheet object -- all available from `regression_layout.py`.

After 5.1 + 5.2, `write_sheet_regression.py` is down to ~1,700 lines (zone writers + orchestrator), which is navigable.

**Decision:** defer unless the file still feels unwieldy after Parts 1-2. These extractions are low-risk but not urgent.

**Resolved — both landed.** Parts 1-2 left the writer at 3,000+ lines rather than the
projected shrink, so the deferral condition was met: 5.1 took the charts out (#214) and
5.2 the materialization band (#220), ending at 2,098 lines of zone writers +
orchestrator. Both were pure moves — every symbol they needed was already a
`regression_layout.py` constant or a leaf helper, so neither module imports back into
the writer and the writer re-exports both under `# noqa: F401` so no call site or test
import changed.

## Part 6 -- Extract complex cell formulas into named LAMBDA functions

**Why:** The project's philosophy is formula transparency -- "any result can be interrogated by clicking the cell." Several cells on the Regression sheet contain complex inline conditionals (nested IFs, LET with nested IFs, inline transform dispatch) that are not auditable. These should become named LAMBDA functions in `lambda_functions.json` so the cell reads a plain-language function name and the logic lives in the catalog where it can be tested, documented, and reused across template instances.

### 6.1 -- Audit and catalog

Before extracting anything, produce a complete list of every cell formula on the Regression sheet (and the spec block) that contains nested conditionals or inline logic that should be a LAMBDA. For each one, record:
- The cell address (e.g., B2, O2, AK3)
- The current formula (abbreviated)
- The proposed LAMBDA name (e.g., `Role_Status`, `Design_Width_Status`)
- Whether it reads sheet-scoped names (all of them do -- this determines the scope must be sheet-scoped or the LAMBDA must take the context as an argument)
- Whether a test oracle needs updating

Known candidates (from the initial scan):

| Cell | Current | Proposed LAMBDA | Notes |
|---|---|---|---|
| B2 (Role status) | Triple-nested IF on response_count and fe_count | `Role_Status()` | Reads `Spec_Role` -- sheet-scoped. Currently in `_ROLE_STATUS_FORMULA` in write_spec_block.py. |
| H2 (Sequence status) | IF on seq_flag_count > 1 | `Sequence_Status()` | Reads `Spec_Sequence` -- sheet-scoped. Currently `_SEQUENCE_STATUS_FORMULA`. |
| O2 (Width guard) | LET with nested IF on k and n*k thresholds | `Design_Width_Status()` | Reads `Predictor_Columns()` and `Source_Data` -- sheet-scoped. Currently inline in write_sheet_regression.py `_write_design_matrix_width_guard`. |
| AK3 (Prediction interval) | IF + LET + inline transform dispatch + Group_Prediction_Interval call | `Prediction_Interval_For_Spec()` or split into `Prediction_Input_Transform()` + existing `Group_Prediction_Interval` | The transform dispatch (`IF(trn="Log",Ln_Positive(raw),raw)`) should be a LAMBDA. |
| AK rows 19-62 (Prediction input prefill) | Per-row `IF(ROW()-offset<=ROWS(means#),INDEX(...),"")` | `Prediction_Input_Means()` spill | Should be a single spill, not 44 per-row formulas. |
| G2 (Log domain status) | Conditional on transform column and nonpositive count | `Log_Domain_Status()` | Currently inline in `_write_spec_feedback`. |

Additional candidates to scan for during the audit: the Cook's Distance flag column (AZ), the residual conditional formatting expressions, and any other cell that contains `IF(` nested more than one level deep or `LET(` with more than two bindings.

### 6.2 -- Extract the first batch (lowest risk, highest clarity gain)

Start with the status cells (B2, H2, G2, O2) because:
- They are the most visible to a user (row 2 of the spec block)
- They are the most complex inline formulas
- They read only sheet-scoped names, so the extracted LAMBDAs follow the existing sheet-scoped pattern (like `Base_Period_Delta`)
- Their tests (guard states, spec-block QC) are the most comprehensive, giving immediate regression coverage

For each extraction:
1. Add the LAMBDA definition to `lambda_functions.json` with scope `"Regression"` (sheet-scoped, so each Regression template gets its own). The body reads unqualified spec references (`Spec_Role`, `Spec_Sequence`, etc.) that resolve against the calling sheet.
2. Replace the inline formula in the writer with `=Function_Name()`.
3. Update the test oracles if they reference the formula text (most don't -- they read the computed value back or assert on the status text).
4. Run `uv run pytest tests/test_regression_guard_states.py tests/test_regression_spec_qc.py tests/test_spec_block_writer.py`.

### 6.3 -- Split the prediction interval VSTACK into per-row formulas

**The problem:** AK3 currently contains one ~400-character formula that branches on `Zero_Predictors_Selected()`, builds an intercept-only path with inline LET, builds a live path with inline transform dispatch, and calls `Group_Prediction_Interval` -- which itself returns a 9-element VSTACK. The result is a VSTACK-inside-a-VSTACK where AK3 spills 9 rows and AK4:AK11 are silent spill cells with no formula of their own. Clicking any cell in AK4:AK11 shows nothing auditable -- it's just the spill. The AL column (Original Units) reads individual cells from the spill, so the back-transformation cells are the only per-row formulas.

**The split:** Give each row its own formula. Two options:

**Option A -- hidden helper + INDEX (recommended):** Keep `Group_Prediction_Interval` as the engine, but call it once in a hidden helper cell (in the materialization zone, next to the design matrix). Each AK row reads one element:
```
AK3  =INDEX(Prediction_Results#, 1)   -- Point Estimate
AK4  =INDEX(Prediction_Results#, 2)   -- SE (Mean)
AK5  =INDEX(Prediction_Results#, 3)   -- SE (New Obs)
AK6  =INDEX(Prediction_Results#, 4)   -- t Critical
AK7  =INDEX(Prediction_Results#, 5)   -- CI Lower
AK8  =INDEX(Prediction_Results#, 6)   -- CI Upper
AK9  =INDEX(Prediction_Results#, 7)   -- PI Lower
AK10 =INDEX(Prediction_Results#, 8)   -- PI Upper
AK11 =INDEX(Prediction_Results#, 9)   -- Confidence Level
```
The hidden helper cell holds `=Group_Prediction_Interval(Predictor_Columns(), Response_Column(), Transform_Prediction_Input(...), Prediction_Group_Column(), $AK$12, Sample_Include(), alpha, Fit_Context())` -- the transform dispatch lives inside `Transform_Prediction_Input`, not inline. Each AK cell is now individually auditable. The `Zero_Predictors_Selected()` branch moves into the helper or into a `Prediction_Results()` sheet-scoped LAMBDA that wraps both paths.

**Option B -- per-statistic LAMBDAs:** Provide `Prediction_Point_Estimate`, `Prediction_SE_Mean`, etc. as separate catalog functions. Cleaner conceptually but each recomputes the fit independently (Excel LAMBDAs have no memoization), so 9 full OLS fits instead of 1. Not viable for a live-recalculation workbook.

**Why Option A:** it preserves single-computation, makes every cell auditable, and the hidden helper follows the same pattern as the existing materialization zone (the design matrix and Sample_Include are already hidden spills that other formulas read). The INDEX-over-spill pattern is already used elsewhere on the sheet (the prediction input prefill reads `INDEX(means#, ...)`).

**What changes:**
- The AK3 mega-formula is replaced by 9 per-row INDEX formulas (or 9 per-row calls to a `Prediction_Result(stat_name)` LAMBDA that wraps the INDEX).
- The transform dispatch (`IF(trn="Log",Ln_Positive(raw),raw)`) is extracted into `Transform_Prediction_Input` (or `Transform_Prediction_Inputs` plural -- it takes a row vector and a transform-flag vector and returns the transformed vector).
- The `Zero_Predictors_Selected()` intercept-only branch moves into the helper cell or a wrapping LAMBDA so it doesn't clutter the visible AK cells.
- The AL column (Original Units) keeps its per-row `Back_Transform_Response` calls -- those are already clean, per-row, and auditable.
- The verifier (`regression_spec_sheet_io.py`) changes from `read_col(sheet, ROW_PI_POINT, _C_AK, 9)` (reading the spill) to reading 9 individual cells -- but the values are the same, so the oracle is unchanged.
- `test_sheet_writers.py::test_prediction_interval_binds_constructed_inputs_in_the_cell_formula` changes: it currently asserts the AK3 formula text contains `Group_Prediction_Interval` and the transform dispatch. After the split, it asserts the helper cell contains those, and AK3 contains `INDEX`.

### 6.4 -- Prediction input prefill spill (follow-on)

Replace the 44 per-row `IF(ROW()-offset<=...)` formulas in AK19:AK62 with a single `Prediction_Input_Means()` spill. This is the cleanest change conceptually but the most sensitive to the spill-height mechanics that the project has carefully engineered. Defer until the structural refactoring is stable.

**Scope decision for this branch:** 6.1 (audit) + 6.2 (status cells) + 6.3 (prediction interval split) should land. 6.4 is a follow-on that can land in a subsequent PR.

## Execution order

```
Part 1.1  Extract regression_layout.py         (no Excel, pytest verifies)
Part 1.2  Extract spec_layout.py               (no Excel, pytest verifies)
Part 2.1  Extract spec_dataset_profiles.py     (no Excel, pytest verifies)
---  checkpoint: commit, run full pytest, all green  ---
Part 3.1  Merge build scripts                  (no Excel, pytest verifies)
Part 3.2  Update deep_verify.py
Part 3.3  Update build_common.py docstrings
Part 3.4  Fold test_build_univariate into test_build_production
Part 3.5  Update poe tasks + test_poe_tasks.py
---  checkpoint: commit, run full pytest, all green  ---
Part 4.1  Update documentation                  (prose only)
---  checkpoint: commit  ---
Part 6.1  Audit and catalog complex cell formulas
Part 6.2  Extract status-cell LAMBDA functions (B2, H2, G2, O2)
Part 6.3  Split prediction interval VSTACK into per-row formulas
---  checkpoint: commit, run full pytest + poe verify-deep (needs Excel)  ---
Part 5    (optional) Extract charts + materialization
Part 6.4  (follow-on) Prediction input prefill spill
---  final: poe verify-deep (needs Excel) confirms the unified workbook builds and verifies  ---
```

## Risk assessment

**Parts 1-2 (constant extraction):** Lowest risk. No behavior change. The only risk is a missed import, caught immediately by pytest. If a patch fails to apply, re-read the file and retry -- don't attempt a third.

**Part 3 (build merge):** Medium risk. The build scripts are structurally similar but have subtle differences (sheet sets, verify skip flags, tab ordering). The merged build must produce a workbook with all sheets present, all named ranges, and both Regression and Univariate verifiers passing. The `test_build_production.py` stub-based tests catch most issues without Excel; the final `poe verify-deep` (with Excel) is the gate.

**Part 4 (documentation):** No code risk. The risk is missing a reference and shipping inconsistent docs. The `test_doc_links.py` and `test_doc_catalog_counts.py` tests catch broken links and stale catalog counts.

**Part 6.1-6.2 (LAMBDA extraction):** Medium risk. Each extracted LAMBDA must produce the same status text as the inline formula it replaces. The guard-state tests assert exact status text, so a mismatch is caught immediately. The risk is a LAMBDA body that resolves differently when called from the formula bar vs. evaluated by the test's RecordingSheet mock -- mitigate by testing both paths. The extracted LAMBDAs are sheet-scoped (scope `"Regression"`), following the existing `Base_Period_Delta` pattern, so they are per-template-instance and safe in a multi-template workbook.

**Part 6.3 (prediction interval split):** Medium-high risk. The AK3 VSTACK is read as a 9-element spill by the verifier (`regression_spec_sheet_io.py`) and by the AL column (Original Units back-transformation). Splitting into per-row formulas changes the read pattern from spill-read to cell-read. The values must be identical. The hidden helper cell in the materialization zone must be writable without disturbing the existing design-matrix spill or Sample_Include spill. The `test_prediction_interval_binds_constructed_inputs_in_the_cell_formula` test must be updated to assert the new cell structure. Gate on `poe verify-deep` (Excel) -- the spec-driven verifier compares AK3:AK11 cell-by-cell against the NumPy oracle, so any mismatch is caught.

**Key invariant to check at every checkpoint:** `uv run pytest` passes with zero failures. The tests are the safety net -- 886 test functions covering the spec block, the sheet writers (via RecordingSheet mock), the build scripts (via stubs), the poe task definitions, and the workbook invariants (via zipfile, no Excel).

## Files touched (summary)

New files:
- `lambda_catalog/regression_layout.py`
- `lambda_catalog/spec_layout.py`
- `lambda_catalog/spec_dataset_profiles.py`
- (optional) `lambda_catalog/regression_charts.py`
- (optional) `lambda_catalog/regression_materialization.py`

Deleted files:
- `scripts/build_univariate.py`
- `tests/test_build_univariate.py`

Modified files:
- `lambda_catalog/write_sheet_regression.py` (loses ~520 lines of constants; loses inline O2 formula to LAMBDA)
- `lambda_catalog/write_spec_block.py` (loses ~378 + 705 lines; loses inline B2/H2/G2 formulas to LAMBDAs)
- `lambda_catalog/build_common.py` (docstring)
- `lambda_catalog/deep_verify.py` (remove skip_univariate)
- `scripts/build_production.py` (absorb univariate sheet writes)
- `tools/verify_workbook.py` (skip_univariate)
- `pyproject.toml` (poe tasks)
- `tests/test_build_production.py` (absorb univariate tests)
- `tests/test_poe_tasks.py` (remove univariate task refs)
- `lambda_functions.json` (new sheet-scoped LAMBDA definitions: Role_Status, Sequence_Status, Design_Width_Status, Log_Domain_Status, Transform_Prediction_Input)
- `lambda_catalog/regression_spec_sheet_io.py` (prediction interval read pattern: spill-read -> per-cell-read)
- `README.md`, `AGENTS.md`, `CLAUDE.md`, `CONTRIBUTING.md`
- `docs/ROADMAP.md`, `docs/DECISIONS.md`, `docs/ARCHITECTURE.md`, `docs/TODOs.md`
# CLAUDE.md — Project context for Claude Code

> **What this file is.** Session primer — the rules you cannot infer from code, read once per session. Deep reference lives in `CONTRIBUTING.md` (build/verify/test-model workflows, flag tables, conventions) and `docs/` (planning: `ROADMAP.md`, `MODEL_TESTING_ASSETS.md`, `ARCHITECTURE.md`, `DECISIONS.md`). Where a section is a paraphrase of one of those, a one-line pointer is the canonical entry.

## PR workflow

**Reply to PR comments after committing fixes.** When a review comment (Copilot or human) leads to a fix, push the fix and then reply to that comment explaining what was changed and why. This is how the repo owner spots addressed comments and resolves the threads.

**Automated verification.** Run `python scripts/build_production.py --verify --no-launch` to build and verify the unified workbook. The spec-driven verifier is **not in CI** — GitHub-hosted `windows-latest` lacks Microsoft Office, so `xw.App` fails with `pywintypes.com_error: (-2147221005, 'Invalid class string')`. Layer 1 (`poe verify-headless`) runs on every push; its committed-artifact checks (`tests/test_workbook_invariants.py`) read the `.xlsx` as a zipfile, so they need no Excel and screen the shipped workbook in CI — gate a check on whether it needs Excel, not a blanket env var. Full pipeline + flag tables: `CONTRIBUTING.md` → *Verifying builds*.

**Recalculate mode is unconditional.** The production constructor always runs `CalculateFullRebuild` and saves in full **Automatic** — there is no skip-calculation flag. (Regression's rebuild is cheap; Univariate's is slower because it materializes the Beta `Full_Factorial` spills, but it always runs.) Full flag tables: `CONTRIBUTING.md` → *Production build*.

## Testing regime

**The regression test-model suite is planned in `docs/MODEL_TESTING_ASSETS.md`** — the plan of record for which configurations the QC harness covers, the corner each exists for, the coverage matrix, and the datasets future milestones need. Read it before adding or changing a QC model case; add to it before adding a case it does not list. Step-by-step: `CONTRIBUTING.md` → *The regression test-model suite*.

**The four-rule PR shape (feature + oracle + test-model case + transcript in `excel-only-runs/`) and the work-in-progress exception:** `CONTRIBUTING.md` → *The PR-shape rules — what every Regression PR must contain*. The plan of record for test-model cases is `docs/MODEL_TESTING_ASSETS.md`.

**Covering-array regime (target ~25–30 fittable models + ~10 guard states, no full crosses):** `CONTRIBUTING.md` → *The regime, in four rules*.

**A case is a `RegressionSpecCase` in `lambda_catalog/analyze_regression_spec.py`** (not a sheet fixture), with expected values from `calculate_regression_spec_case` (NumPy/statsmodels — reading the cell back is not an oracle). **Non-default dataset: set `source_csv_path`, `row_loader`, and `source_table_ref` together** — `Source_Table` is the one name that retargets the data sheet, and omitting it lands the spec rows on the wrong columns silently. A guard-rail configuration is a `GuardStateCase` in `analyze_regression_guard_states.py` instead (cases raise in the spec oracle by design, asserting status text + Design Columns audit + CF fires, not fit stats). **`GuardFlag` is a predicate, not a pixel** — recompute the rule condition, never read `DisplayFormat.Interior.Color`. Full convention: `CONTRIBUTING.md` → *The regime, in four rules*.

**Every case materializes as its own sheet in `Lambda_Library_TestModels.xlsx`** (gitignored; built by `scripts/build_test_models.py`, verified read-only by `tools/inspect_test_model_sheets.py`). **Sheet names** are governed by `lambda_catalog/test_model_sheets.py` (31 chars, legal charset, `<PlanID> <Concept>`, unique across model + guard cases) and validated at registry-build time, so a bad name fails in the unit suite rather than mid-build. **The name states the concept under test, never the variables** (`M05 Log-Log NA Masking`, not `MPG ~ Ln(Weight) + Ln(HP)`). Both halves of the sheet contract (`apply_spec_case`/`set_prediction_inputs` write, `read_case_comparison_rows` read) live in `lambda_catalog/regression_spec_sheet_io.py`, shared with the legacy single-sheet verifier so the two cannot disagree.

**The spec block has no fixed height, and no `SpecTable`.** Every part of it sizes itself from `COLUMNS(Source_Data)`, so retargeting `Source_Table` resizes the block — the one-name edit the Instructions sheet has always promised. Three mechanisms, all in `write_spec_block.py`:

* the `Spec_*` bands are `=TAKE($X$4:$X$16000,MAX(1,COLUMNS(Source_Data)))`, built by `_spec_band`
  — `TAKE` not `OFFSET`, for the same non-volatility reason `Source_Data` and `Header_Names` use it;
* the four computed columns (J Period In Use, K Levels, L Reference In Use, O Design Columns) are
  each ONE spill at `_FIRST_DATA_ROW`, `MAP(SEQUENCE(nc),LAMBDA(i,…))`, written with `f`
  (**`Formula2`** — `.Formula` enters a dynamic array as a legacy CSE range, which does not resize);
* the input band's `INPUT_COLOR` fill is a single lowest-priority CF rule on the same predicate,
  not per-row `format_input` calls.

**Nothing may be written below the spec block in columns B–O**: the bands reach row 16000 and the spills need clear space, so stray content is a `#SPILL!` error, not a truncation. The one ceiling — `_SPEC_BAND_LAST_ROW`, the same 16000 the validations and CF rules use — is deliberately shared so they cannot disagree about how far the block can grow.

**Case-name pinning** (`_EXPECTED_CASE_NAMES` in `tests/test_regression_spec_qc.py` and `_EXPECTED_GUARD_NAMES` in `tests/test_regression_guard_states.py` — add the name in the same commit): `CONTRIBUTING.md` → *The regime, in four rules*.

**The ladder-order rule** (track-then-growth-rate, v3.10/v3.11 ship last as a block, ladder follows the MODEL_TESTING_ASSETS § 2 table): `docs/ROADMAP.md` § *Ladder order* and `CONTRIBUTING.md` → *The regime, in four rules*.

## QC comparison scale

**A compared statistic is scored against the magnitude its error comes from, not
against its own.** `first_digit_deviation` scores in DECIMAL PLACES, so an
absolute comparison silently gets stricter as a statistic's magnitude grows.
Three cases, all declared in `lambda_catalog/regression_spec_sheet_io.py`:

| Statistic's error tracks... | Divisor | Declared in |
|---|---|---|
| its own value | `max(\|expected\|, 1.0)` | `SCALE_FREE_STATS` |
| the fitted values, in response units | response RMS | `_RESPONSE_UNIT_STATS` |
| the fitted values, over `SE_Regression` | response RMS / `SE_Regression` | `_STANDARDIZED_RESIDUAL_STATS` |
| nothing larger than itself | none — absolute | (default) |

`compare_values` takes an explicit `scale` for the response-derived cases and floors every divisor at 1.0, so the convention can only loosen a comparison, never tighten one. **Derive the divisor from the fit, never a constant** — a response in the tens and one in the billions must be treated proportionately.

**Adding a compared statistic means choosing its case.** The residual band is the inherited-error one: every statistic there is built from the predictions, so it carries the response's precision floor whatever its own size (a residual of order 0.1 carries the error of numbers of order 70, and self-scaling divides it by 1.0). `T_Statistics` is in none of the response-derived sets — dimensionless and O(1), its error from the coefficient, so a response-derived divisor would be a number picked to fit. Full rationale: `docs/DECISIONS.md` → *QC comparison scale, the clear-list invariant, and the OLS solver*.

**The OLS oracle is pinned to `method="qr"`** in `_fit_ols_model`, not the statsmodels `"pinv"` default: the oracle is the reference the sheet is scored against and should be the more accurate side; LINEST is QR-based too.

## The Transform column has two Log tokens

`None` · `Log` · `Log (drop ≤ 0)`, from `_TRANSFORM_LOG` / `_TRANSFORM_LOG_DROP` in `write_spec_block.py`. **Both build the identical `Ln(x)` column and both report `"Log"` to `Constructed_Column_Transforms()`** — the unit-space / Duan dispatcher sees one transform, not two, which is why a second token cost the test suite two cases instead of a multiplier. They differ in exactly one thing: what happens to a row whose value is zero or negative.

* **`Log`** leaves the row in the sample. `Ln_Positive` returns `#N/A` and it
  propagates through every statistic — the fit is dead, not degraded. The
  Transform cell goes red and `G2` names the variable, the count and the fix.
* **`Log (drop ≤ 0)`** is the only thing that adds the positivity term to
  `Sample_Include`. `G2` reports the excluded-row count in amber.

Filtering never happens because the workbook decided to — narrowing the sample changes the model, so it is a declaration in the spec. Don't "fix" `Log` by making it filter.

**Every Excel-side test of "is this a Log row?" goes through `_is_log(expr)`, and its Python mirror is `_logs(transform)` in `analyze_regression_spec.py`.** Never compare a Transform cell against one token. `Sample_Include` is the sole deliberate exception — it tests `= "Log (drop ≤ 0)"` alone, because plain `Log` not filtering is the whole distinction.

The token also lives literally inside `lambda_functions.json`, which no import can reach; `test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform` pins the two spellings together across all seven readers so a rename cannot half-land. `Log_Domain_Status` is the second deliberate exception to the pairing rule, alongside `Sample_Include` and in the opposite direction: it equality-tests only the STRICT token, because the rows it counts are the ones a strict Log leaves in the sample to poison the fit, naming the drop token as the remedy in prose rather than as a comparison.

## Cell styling

All cell colors are defined once in `lambda_catalog/sheet_styles.py` and imported by every sheet writer. Never hard-code RGB tuples in a sheet writer.

| Constant | RGB | Usage |
|---|---|---|
| `HEADER_COLOR` | `(202, 237, 251)` | Zone / section headings — light blue, used on every sheet |
| `SUBHDR_COLOR` | `(220, 230, 241)` | Column sub-header rows within a zone |
| `INPUT_COLOR`  | `(251, 226, 213)` | User-editable input cells — light orange |
| `CF_LIGHT_RED_FILL` | `(255, 199, 206)` | Conditional formatting — failed significance / diagnostic flag |
| `CF_DARK_RED_TEXT` | `(156, 0, 6)` | Conditional formatting — text color for red-flagged cells |
| `CF_YELLOW_FILL` | `(255, 235, 156)` | Conditional formatting — borderline diagnostic flag |
| `CF_DARK_YELLOW_TEXT` | `(156, 101, 0)` | Conditional formatting — text color for yellow-flagged cells |

Import pattern in a sheet writer: `from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR` — the `as _NAME` alias keeps local helpers (`_subheader_row`, etc.) unchanged. Sheet-specific colors that differ from the shared palette (e.g., `_SUBHEADER_COLOR` in `write_sheet_diagnostic_guide.py`) remain as local constants in the relevant file.

**Exception — the Regression spec block's header row uses `HEADER_COLOR`, not `SUBHDR_COLOR`.** By the table above it is a column sub-header row and would take `SUBHDR_COLOR`; it is deliberately pinned to `HEADER_COLOR` (`#CAEDFB`) with black bold text in `_write_spec_block`, because that row heads the sheet's primary *input* surface and reads as a zone heading, not a subdivision. `test_spec_block_prefills_the_t0_default_configuration` asserts all three properties (fill, color, bold), so a competing style cannot quietly return.

### Section heading style

A **section heading** is bold text with `HEADER_COLOR` fill at the default font size. The sheet title ("Univariate Analysis") is 14 pt bold with no fill — the only cell with a custom font size. Apply this style via the shared helper `section_heading(sheet, row, col, label)` in `lambda_catalog/workbook_helpers.py` (used by both `write_sheet_regression.py` and `write_sheet_univariate.py`); do not apply it inline.

### Univariate sheet heading hierarchy

Zones 1–4 (cols A–Z) use the standard row layout: row 1 — "Univariate Analysis" (A1, 14 pt title) + "Histograms" (F1:M1, `section_heading`); row 2 — Sturges (F2:G2) / Scott (I2:J2) / Freedman-Diaconis (L2:M2) / "Distribution Fitting/Comparison" (P2:Z2), all `section_heading` + merged; row 3 — "Data" (A3) / "Descriptive Statistics" (C3), `section_heading`; row 4 — column sub-headers ("Upper Edge", "Count", "Distribution", …) via `_subheader_row`; row 5+ — data / spill formulas.

### Univariate right-hand band — one ordered zone table

Everything right of the histograms is **one zone per topic**, each followed by a single gap column, all derived from the ordered `_BAND_ZONES` table via `_derive_band_columns()`. Zone starts (`_C_QQ`, `_C_GS_WB`, `_C_GS_GAMMA`, `_C_GS_BETA`), the gap columns, and the sheet's last column all come from it. **Never hard-code a column letter in this band** — reordering it means reordering that list and nothing else. Zones: 5 Q-Q plot data (BE–BN, w10); 6 Weibull fit (BP–BS, w4, rows 1–52); 6 Gamma fit (BU–BX, w4, 1–52); 6 Beta fit (BY–CD, w6, 1+ N² body).

**The fit zones must stay last** — they are the only zones whose width is a tunable (Beta's 6 cols reflect its two-stages-side-by-side Alpha|Beta|NLL layout; editing N resizes the body **downward**, not the band **width**), so keeping them last means a resize displaces nothing. Same principle as the Regression sheet's rule that nothing may sit right of the design-matrix zone. **Beta follows Gamma with no gutter** (per-zone `gutter_after` flag on `_BAND_ZONES`; `gamma`→False) — the one deliberate exception to the uniform-gutter rule, matching the owner's hand-laid reference.

### Univariate fit zones

**One zone per distribution, two stages side by side** (S1 left, S2 right). Stacking is not viable — every body is a spill whose height follows an in-sheet `Grid Points` cell, so a stacked Stage 2's row anchor would depend on Stage 1's dynamic spill height; side-by-side stages share a body row and grow down independently. **Every stage of every fit has the same shape** — a `Full_Factorial` spill beside a `BYROW` NLL column reading it via `#`, so the NLL is exactly as tall as its grid. **The three fits use two stage writers**, sharing one row skeleton per zone: control field-list rows 4–(≤12), profile-NLL chart rows 13–30, body header row 32, body rows 33+.

**Beta — `_write_beta_fit`, a 2-D `Full_Factorial` grid plus a BYROW NLL column** in a 6-col (`BY:CD`) zone, two stages side by side. Each stage's body is **two** spills at row 33: a `Full_Factorial` grid spill (N²×2, `Alpha | Beta`) — `=Full_Factorial($<Ncell>,VSTACK($<αmin>$,$<βmin>$),VSTACK($<αmax>$,$<βmax>$))` — and a `BYROW` NLL spill reading it via `#`: `=LET(…,z,(d-MIN(d)+pad)/scale_,BYROW($<grid>$#,LAMBDA(ab,NLL_Beta(z,INDEX(ab,1,1),INDEX(ab,1,2))+COUNT(d)*LN(scale_))))`. Two spills keep the grid a pure Cartesian product, NLL by reference. Stage 1 grid at BY33 (BY,BZ) + NLL at CA33; Stage 2 grid at CB33 (CB,CC) + NLL at CD33, refined bounds (α/β Min/Max = best±step from Stage 1). Recovery is `Grid_Argument_Minimum` over the materialized NLL column. The control block is a vertical field-list with **two value columns per stage** (α and β side by side) so it fits in ≤9 rows and clears the chart space at row 13. Control rows per stage: 1 title; 3 sub-headers (S1/S2); 4 Grid Points (N — S1 default 10, S2 `=<S1 Ncell>`, live); 5–7/9 α/β Min/Max/Step (S1 wide, S2 best±step); 8 Min NLL (`=IFERROR(TAKE(Grid_Argument_Minimum(UV_BETA_S{1,2}_NLL),,1),"—")`); 10 Min NLL value (Stage 2 anchor); 11–12 Best α/β (`=IFERROR(INDEX(UV_BETA_S{1,2}_Alpha,INDEX(Grid_Argument_Minimum(UV_BETA_S{1,2}_NLL),1,2)),"—")` / `_Beta`); 32 body headers (`Alpha`/`Beta`/`NLL`); 33+ `Full_Factorial` N²×3 spill.

**Weibull and Gamma — `_write_profile_fit`, a 1-D `Full_Factorial` axis plus a BYROW profile-NLL column** in a 4-col (`_PS_W`) zone. The scale/rate parameter is profiled out in closed form (`_weibull_profile_scale`, `_gamma_profile_rate`), so a stage is N evaluations, not N². The axis is `=Full_Factorial($<Ncell>,$<Min>,$<Max>)` (the d=1 reduction of Beta's grid); the NLL beside it is `=LET(x,FILTER(UV_Data,UV_Include),BYROW($<axis>$#,LAMBDA(r,LET(p,INDEX(r,1,1),IFERROR(<NLL at p, partner profiled out>,1E+15)))))`. Three details are load-bearing: `INDEX(r,1,1)` scalarizes the 1×1 row `BYROW` hands the callback; `IFERROR` sits **inside** the `LAMBDA` so one bad trial costs its own row, not the column; and `x` is bound once per stage, not per row. The control block is a vertical field-list (label col + one value col per stage); both stages' bodies sit side by side on shared rows. Control rows per stage: 1 title; 3 sub-headers (S1 wide / S2 refined); 4–9 Grid Points/Min/Max/Start/Min NLL/Step Size (S1 col1, S2 col2); 10–11 Optimal Shape (searched) / Optimal Scale (profiled-out closed form); 32 body headers (`Shape (k)` / `Profile NLL`); 33+ `Full_Factorial` axis + `BYROW` profile NLL side by side (N rows, default 20).

Both writers share the `_GS_R_*` / `_PS_R_*` row offsets. Each fit has its own Best cells and Stage 2 anchor (one zone per fit), so **`_final_grid_best_refs` is per-distribution**, reading `_STAGE2_ANCHORS`. Side-by-side stages ⇒ Stage 2's Best cells are at **fixed control rows** in the Stage 2 value col (no N² row offset), so the old `beta_grid_size**2` row-offset logic is gone. `_STAGE2_ANCHORS` is the single agreement point between fitting table and search writers — the table cannot reference a block the writers did not produce. Row/column positions come from the `_BAND_ZONES` table and the `_GS_R_*` / `_PS_R_*` / `_PR_C_*` constants at the top of `write_sheet_univariate.py` — never hard-code them inside either stage writer. Beta's visible Alpha and Beta Input cells are its `Full_Factorial` bound cells.

**`Grid Points` (N) is a live cell in all three fits.** Editing it resizes that stage's grid (N² rows for Beta, N for a profile fit) and its NLL column follows via `#`; the OFFSET named ranges — `UV_BETA_S{1,2}_{Alpha,Beta,NLL}`, `UV_{WB,GAMMA}_S{1,2}` + their `_Axis` partners, `UV_Profile_*` — are all sized `MAX(IFERROR(<N cell>,1),1)` (squared for Beta) so Min NLL, Best-parameter recovery, the boundary guard, and the charts all track the new height. Each cell carries a whole-number `Validation` and a red CF at the `MIN_GRID_POINTS` floor (2 — the Step cell divides by N−1); `build_common.positive_grid_size` applies the same floor to `--beta-grid-size`. Static formatting (number formats, colour scale, border box) is painted over the **default-size** window (`_PROFILE_BODY_CF_ROWS_CAP` / `_BETA_BODY_CF_ROWS_CAP`); rows a live increase adds beyond it stay unshaded — cosmetic, and deliberate.

A 1-D stage cannot use `Grid_Search_Optimum` — on a single-column grid its column-parameter half reads the cell above the body (the `Profile NLL` header, not a parameter value). Use `INDEX(<axis name>, INDEX(Grid_Argument_Minimum(<body name>),1,2))`, as `_write_profile_fit` does. Beta uses `Grid_Argument_Minimum` over its materialized NLL column (`UV_BETA_S{1,2}_NLL`) for Min NLL / Best α / Best β recovery — safe because the NLL column is already materialized before the name reads it.

Zone 5 (Q-Q plot data) holds Hazen plotting positions `P`, the sorted `Sample` column, and per-distribution theoretical-quantile columns referencing the fit-table parameter cells. Charts under the fitting table — histogram combo + per-distribution Q-Q scatter charts fed by OFFSET-based `UV_QQ_*` named ranges. The two Weibull/Gamma profile-NLL charts are anchored **above the body**, rows 13–30, one zone wide (Weibull `BP13:BS30`, Gamma `BU13:BX30` — between control block and body), each plotting *both* stages from `UV_Profile_<dist>_<S1|S2>_<Axis|NLL>` OFFSET names (pointing at the live-N body from row 33), with `+` markers on Stage 2 so the refined region stays visible over the wide Stage 1 curve. **Beta has no chart yet**: `BY13:CD30` is blank, reserved for a future Beta chart.

### Regression sheet heading hierarchy

Row 1 holds the top-level zone labels ("MODEL SPECIFICATION", "PREDICTOR SUMMARY", "REGRESSION OUTPUTS", "PREDICTION OUTPUTS", "RESIDUAL OUTPUT"); lower section headings appear at relevant data rows. The MODEL SPECIFICATION zone (A–O) is the shared spec block imported from `write_spec_block.py` (headers row 3, spec rows `_FIRST_DATA_ROW`–`_LAST_DATA_ROW` — currently 4–15, sized to `len(_VARIABLES)`; H = Sequence structural flag, I = Sequence Period (typed override input), J = Period In Use (candidate-with-override display), K/L = Levels / Reference In Use displays, M/N = the Interaction Term / Interaction Operation pair (constructor-wired since v3.1), O = Design Columns audit with Σ total at O1); every other zone keeps headers on row 3 with spills from row 4, aligned with the spec block's own header and data rows. Rows 1–3 are a frozen pane (freeze at A4 via select-then-FreezePanes), so every zone's headers stay visible while scrolling. Every `_C_*` column constant in `write_sheet_regression.py` matches its actual column letter (`_C_N` is column N).

**Rows 1–3 above the data have one grammar: row 1 is labels, row 2 is the control, value or status, row 3 is the column-header row for every zone — and every status sits in the spec column it is about.** Role cardinality at B2, Log domain at G2, Sequence cardinality at H2, spacing verdict at I2, design-matrix width guard at O2 (under its own Σ total, not at M2 above Interaction Term). Status cells carry no row-1 label: blank when legal, naming its own subject when it fires. Only readouts are labelled: Intercept (C1/C2), the FE trio (J1:L2), Σ Design Columns (N1/O1), Δ/Count (P3/Q3, spectrum spilling from P4). Each status owns a column with little runway, so every one is `WrapText` with a short imperative message plus a hover Note for full guidance; row 2 stays automatic height — one line when legal, grows when a message fires. Readouts whose feature is not in play are hidden white-on-white with their labels (`_hide_when`), not `"n/a"`. Add a new status on row 2 of its own column — never in a free cell (the old E1/B1/M2 scatter). **The four status cells hold a CALL, not the logic**: B2/G2/H2/O2 are `=Role_Status()` / `=Log_Domain_Status()` / `=Sequence_Status()` / `=Design_Width_Status()`, sheet-scoped LAMBDAs in `lambda_functions.json`, so clicking the cell shows what is checked. A new status is a new sheet-scoped catalog entry, not an inline writer formula. What stays in Python: long hover Notes (>255-char catalog cap), CF rules, shared count sub-formulas. `tests/test_spec_block_writer.py` pins each body's message text to the guard-state oracle and `tests/test_sheet_writers.py` pins the width thresholds to the `regression_layout.py` constants that also size the design-matrix band. I2's spacing verdict stays inline — the one extraction did not cover.

**Never spell an A1 address into a formula string.** Conditional-formatting expressions, chart titles, and OFFSET-based named ranges need addresses as literal text, and hand-written letters turn a column insertion into a silent-wrong-answer bug — the formula still parses, it just reads a different cell. Build every one from the `_C_*` constants via `_abs_ref(row, col)` / `_band(col)` and the `_A_*` anchors (`_A_ALPHA`, `_A_OBSERVATIONS`, `_A_MEAN_LEVERAGE`, …) at the top of `write_sheet_regression.py`. The same applies to anything reading the sheet: `tools/inspect_regression_sheet.py` and `lambda_catalog/analyze_regression_spec_block.py` IMPORT the column constants rather than keeping a parallel copy.

**Column-layout paradigm — gap columns and outline groups.** The zones are Model Specification (A–Q, incl. the P/Q Δ-spectrum feedback columns), Predictor Summary (S–Y), Regression Outputs (AA–AH — `AG4:AH10` v3.3 **Unit-Space Fit** block: `Duan`/`Naive` toggle at `AH5`, smearing factor R6, R²/Adj R²/RMSE in original units R7–9, response-space readout R10), Prediction Outputs (AJ–AL — `AL` is the **Original Units** column: back-transformed point estimate at AL4, Naive-only CI/PI bounds at AL8–11; Duan/Naive caveat is a NOTE on the Back-Transform label at `AG5`, not a row), and Residual Output (AN–BA — `AZ`/`BA` carry Predicted Y / Residual in original units, dispatched on the `AH5` toggle). Between every adjacent zone pair sits one dedicated **gap column** (R, Z, AI, AM — width 2) left OUT of every outline group; that ungrouped column lets neighbouring zones collapse independently (Excel fuses contiguous same-level grouped columns, so two zones with no gap share one collapse control). `_ZONES` (the (first, last) content spans) and `_GAP_COLUMNS` (the single column between consecutive zones, asserted one wide) are the single source of truth; `_COLUMN_GROUPS = _ZONES`, with gap columns sized and left ungrouped in `write_regression_output_sheet`'s width/grouping loop. When adding or resizing a zone, edit `_ZONES` — never hard-code an outline group or gap letter.

**Content-column widths are keyed on the `_C_*` constants too — `_COLUMN_WIDTHS`.** It is `(constant, width)` pairs with a module-level assertion that every zone content column is sized exactly once and no gap column is sized at all, so the next shift fails at import, not shipping with the wrong column. The spec block (A–O) keeps its own widths in `write_spec_block._SPEC_COLUMN_WIDTHS`; column I (the Regression-only Verdict overlay) and BB (the post-zone chart gutter) are the two deliberate entries outside the zones.

The **Model Formula readout** sits on row 1 of the terminal Constructed Design Matrix zone (header at `_C_DESIGN_MATRIX + 2`, readout one column past it, both derived from the zone anchor), with `WrapText = False` — row 1 is the one row the matrix can never reach (names spill on row 3, values on row 4, both rightward), so the caption is never displaced. The cell holds `=Model_Formula()` (the sheet-scoped catalog closure); v3.4 reads the NAME (`Comparison_Model_Formula` in `_setup_local_names`), so moving the readout cost its consumer nothing.

**Past the charts sits the ARCHITECTURE §4b materialization band**: the **Model Context** block (two-column label/value pair — `_C_MODEL_CONTEXT_LABEL` / `_C_MODEL_CONTEXT` — one labelled cell per context element, headed/border-boxed (height is a build-time constant), grouped as a pair so it collapses as a unit), the reserved `Sample_Include` column, and the terminal **Constructed Design Matrix** zone. **Model Context is the only zone here grouped or collapsed** — individual cells, so hiding it hides no spill. `Sample_Include` and the design matrix are left ungrouped and expanded on purpose: both are full-height dynamic arrays, and a collapsed outline group over a spill range is the state in which Excel stops recomputing the model — hidden columns keep stale arrays and every engine reading across them refits on stale values. Do not re-add a `Group()` / `ShowDetail` call for either one; the scrolling of an expanded unbounded zone is the accepted cost. Its columns derive from `_LAST_CHART_COLUMN` (tracks the chart anchor), so a zone shift moves the whole band. **Nothing may sit right of the design-matrix zone** — its width is one dropdown from hundreds of columns, so any zone after it would be displaced by an ordinary modelling choice.

**The model context is individual cells, never a `VSTACK` spill.** `_MODEL_CONTEXT_ELEMENTS` in `regression_layout.py` (written by `_write_materialization_zone` in `regression_materialization.py`) is the single source of row order, labels, and block height (`_MODEL_CONTEXT_ROWS`, and from it `Fit_Context`'s fixed range and the `Context OK` health-check row). A spill buys nothing (height is a build-time constant) and costs correctness: one formula producing four cells is a single dependency node Excel vacates and re-spills on every spec change, transiently blanking `Fit_Context()` so ~30 engine call sites read a torn context. Don't reintroduce one.

### Regression chart named ranges

Chart `SERIES` formulas do not support the `#` spill operator, and full-column references degrade recalculation. All chart series reference **worksheet-scoped named ranges** via `OFFSET` sized to the observation count in the Regression Statistics block at `$AB$9`. Example: `RegChartFitY = =OFFSET('<sheet>'!$AP$3,1,0,MAX(IFERROR('<sheet>'!$AB$9,1),1),1)` — one row below the header, exactly `$AB$9` rows tall, the `MAX(IFERROR(...,1),1)` guard keeping it one row tall when `$AB$9` cannot resolve. Full `sheet.api.Names.Add` example + `_name_ref` helper: `CONTRIBUTING.md` → *Chart series data ranges*.

All OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix, distinguishing them from constructor closures (`Predictor_Columns`, `Design_Columns`, `Sample_Include`, etc.) and formula-helper names. The name-to-column map (and the `$AB$9` anchor) lives in the loop in `_setup_local_names` in `lambda_catalog/write_sheet_regression.py` — the single source of truth for the column letters. When adding a diagnostic column or chart, add the `RegChart`-prefixed named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

Worksheet-scoped names are created via `sheet.api.Names.Add` on the owning sheet, and `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer.

### Workbook scope belongs to the catalog

Every named range a sheet writer creates is **sheet-scoped** — `sheet.api.Names.Add` on the owning sheet, never `book.names.add`. `RegChart*`, `UV_*` and the spec wiring all live there. Workbook scope belongs exclusively to the catalog's LAMBDA functions (`lambda_functions.json`, `scope: "workbook"`); `sync_workbook_names` enforces that literally: on every build it drops **every** workbook-scoped `<definedName>` that is not a catalog function or a reserved `_xlnm.*` name, then rewrites the catalog entries.

**Workbook-scoped catalog LAMBDAs must be sheet-agnostic.** A body that hardcodes `'Regression'!` is wrong in two directions: in a workbook with SEVERAL Regression-shaped sheets (the test-model artifact has 50) every sheet reads whichever one is literally named `Regression`, and in a workbook with none it is skipped and every call site reads `#NAME?`. `Base_Period_Delta` is the example: **sheet-scoped** (`"scope": "Regression"`) with unqualified spec references, so an unqualified name resolves against the calling formula's sheet and each Regression sheet gets its own Δ. The skip mechanism (a catalog function naming a missing worksheet is skipped, not written) is what stops the next sheet-qualified body shipping a broken link. Full history: `CONTRIBUTING.md` → *Workbook scope belongs to the catalog*.

**But a RefersTo body that references a LATE-created name cannot rely on that resolution.** "Unqualified resolves against the calling formula's sheet" holds for cell formulas; it does NOT hold inside a defined name's RefersTo when the referenced name does not exist on the sheet at `Names.Add` time — the reference stays unresolved, and at calculation Excel resolves it against the workbook's WHOLE name collection, pinning whichever same-named sheet-scoped copy it finds first. The spill readers (`Fit_Context`, `Fit_Design_Columns`, `Fit_Sample_Include` — created by `_write_materialization_zone`, i.e. AFTER the spec block installs the constructor closures and after `_setup_local_names` registers `Intercept_Only_*`) are the exposed case: in the 50-sheet test-model workbook every sheet's `Log_Domain_Status` and `Intercept_Only_N` pinned `'G01 No Response Row'!Fit_Sample_Include` and reported that sheet's fitted count. Production's single Regression sheet hides the class entirely. **So every RefersTo written before the materialization zone must qualify its spill-reader references with the owning sheet at install time** — `qualify_spill_reader_references` / `SPILL_READER_NAMES` in `regression_materialization.py` (applied in `_set_sheet_scoped_names` and on the `Intercept_Only_*` adds in `_setup_local_names`); `tests/test_sheet_writers.py::test_every_registered_refers_to_qualifies_spill_reader_references` guards it. Full history: `docs/DECISIONS.md` → *RefersTo bodies qualify the late-created spill readers at install time*.

`tests/test_workbook_invariants.py::TestRealWorkbookNameScope` asserts all of this against the committed `dist/Lambda_Library.xlsx` on every commit (pure zipfile, no Excel), as does `TestRealWorkbook` — every committed-artifact check in that file is always-on, so a stale or hand-edited workbook fails the suite instead of shipping. To re-apply the cleanup to a built artifact without rebuilding: `python tools/resync_workbook_names.py <workbook.xlsx>`.

## Charts — patterns and pitfalls

### Use xlwings COM API for all chart creation — never openpyxl

openpyxl's `load_workbook()`/`save()` **rewrites the entire .xlsx package** and silently drops chart parts, VML, and chartUserShapes it didn't create — loading a workbook with Excel-created charts (e.g., the Regression diagnostic charts) and saving destroys them. This is a fundamental, unfixable limitation. All charts use `sheet.api.ChartObjects().Add(...)` via xlwings COM; follow `regression_charts.py:_write_diagnostic_charts`.

### Chart titles — `.Text` for static, `.Formula` for cell-linked

`.Text` sets a literal string (fixed titles); `.Formula` links the title to a worksheet cell so it updates dynamically (`chart.ChartTitle.Text = "..."` vs `chart.ChartTitle.Formula = "='Sheet'!$Q$14"`). When a dynamic title needs to be computed (e.g., concatenating a method name with " Histogram"), write the formula into a dedicated cell first, then point the chart title's `.Formula` at that cell. Do **not** pass a formula string to `.Text` — it renders as literal text, not evaluated.

### Histogram-specific formatting

- Histogram bars must be contiguous: `chart.ChartGroups(1).GapWidth = 0`. Do **not** enable "Vary colors by point" (single uniform bar color).
- Add explicit axis titles (`x_axis.AxisTitle.Text = "Upper Edge"`, `y_axis.AxisTitle.Text = "Count"`) and set title `overlay=False` so Excel auto-sizes the plot area.

### Chart positioning & guarding

Define chart positions via xlwings `sheet.range(...)` for `.left`/`.top`/`.width`/`.height` in points — tying placement to the cell grid so charts stay aligned when column widths change. Charts need the Excel COM API, unavailable in CI/headless; always wrap chart insertion in `try/except` so the sheet build succeeds without charts.

### Never draw reference lines as shapes — use a real data series

Do not use `chart.Shapes.AddLine(...)` to fake a reference line like `y=x`. A shape sits in fixed plot-area pixel coordinates computed at creation time and silently goes wrong when the chart is resized, moved, or its axis scaling changes. Instead, add a real data series pointing both `XValues` and `Values` at the same named range (with `ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS`) — see `_add_identity_line` in `regression_charts.py`.

### Selective data labels — a masked overlay series, not per-point COM loops

To label only points meeting some value criterion (e.g., Cook's Distance above the `F.INV(0.5, p, n-p)` influence cutoff — see `_COOKS_CUTOFF`), add a helper column returning the real value for qualifying rows and a masking token otherwise, expose it as a `RegChart`-prefixed named range, and add it as an extra series with `HasDataLabels = True`. The mask makes the selection — no per-point `Points(i).HasDataLabel` loop, and no reading calculated values back into Python during the sheet-writing phase (which runs under `XL_CALCULATION_MANUAL` and would see stale/unfit values; see "Sheet writer conventions" below).

**Which masking token depends on where the label text comes from.** With `ShowValue`/`ShowCategoryName` the token is `NA()` (Excel skips `NA()` points for plotting and labeling, so unflagged rows disappear). With **Value From Cells** — `DataLabels.Format.TextFrame2.TextRange.InsertChartField(msoChartFieldRange, "='Sheet'!SomeName", 0)` then `ShowRange = True` — the label prints the source cell verbatim, so an `NA()` row renders a literal `#N/A`; the token must be `""`. A `""`-returning formula is plotted (at zero), not skipped, so **every other label element must be off** (`ShowValue = False`, `ShowCategoryName = False`) or it prints on all n observations. The Cook's Distance overlay is the Value-From-Cells case: `AY` is `=IF(AT3#>cutoff,AT3#,"")` and the range supplies the entire label.

On a **column-chart** target, do not give the overlay the chart's own `xlColumnClustered` type — a second column series joins the cluster and narrows/shifts the real bars, misaligning labels. Instead set the overlay's `ChartType = xlLine` (constant `4`) with `Format.Line.Visible = False` and `MarkerStyle = xlMarkerStyleNone (-4142)`: a Line series shares the Column chart's category axis without joining its cluster, so it overlays in place. A per-series `ChartType` differing from the chart's own is how Excel builds a **combo chart** — expect it to become one.

The overlay's `XValues` still points at a named range over the observation-identifier column (`RegChartObsLabel`). Displaying that identifier in the label (`ShowCategoryName = True`, plus `ShowValue` if the number should show) is only available under `NA()` masking: under `""` it would print on every point.

See `RegChartCookDistFlag` / `RegChartObsLabel` in `_setup_local_names` and the Cook's Distance branch of `_write_diagnostic_charts` in `regression_charts.py`.

### Separate chart title cells from chart insertion

Write chart-title formula cells (e.g., `Q14`, `Q34`, `Q54`) **outside** the try/except guard — standard cell writes (not COM API calls), exercisable in unit tests via the `RecordingSheet` mock. Only `ChartObjects().Add(...)` needs the guard.

### Build-phase retry separation

Don't wrap `build_production_workbook()` in one retry loop. The fast (~10s) recalc/save step is most likely to fail when the user opens the workbook mid-build; give it its own retry phase so a failure there doesn't restart the multi-minute sheet-writing phase. See `_retry_on_open` and the two-phase `main()` in `build_production.py`.

## Sheet writer conventions

- **Layout single-source-of-truth**: row constants (`_ROW_TITLE`, `_ROW_DATA_START`, …) at the top of each writer; chart series via `OFFSET` named ranges (CONTRIBUTING.md § *Chart series data ranges*); `app.api.Calculation = XL_CALCULATION_MANUAL` before writes, `XL_CALCULATION_SEMIAUTOMATIC` after, before save.
- **Every defined name carries a Name Manager comment**: set `.Comment` at the `Names.Add` site immediately after the add (`_nm = sheet.api.Names.Add(...); _nm.Comment = "..."` — sheet-scoped wiring, chart-range, and spill-reader names), and the constructor-closure sites set it to the catalog entry's `notes` verbatim, so a name's purpose is stated exactly once. `tests/test_workbook_invariants.py::test_every_defined_name_carries_a_comment` (always-on, pure zipfile) fails the suite on a comment-less name in the committed dist, so a new name without a comment fails CI, not the first user who opens the Name Manager.

### Guard headless/no-focus Excel calls with the `safe_*` helpers

`Sheet.activate()` and anything touching `Application.ActiveWindow` (e.g. freezing panes) raises when Excel cannot become the active application — no interactive desktop, focus denied by the OS, an agentic/headless build host — even though the workbook write succeeds. Which sheet is on top or whether panes are frozen is cosmetic, so that failure must not abort `build_production_workbook()`.

Use `safe_activate(sheet)` and `safe_freeze_top_row(sheet)` from `workbook_helpers.py` instead of the raw calls — each is a `try/except Exception: pass` wrapper, so the cosmetic failure (sheet not on top / panes not frozen) does not abort `build_production_workbook()`. When adding a new sheet writer that activates its sheet or freezes its header row, call the `safe_*` helper, not the raw xlwings/COM call. See `tests/test_workbook_helpers.py` for the stub-based unit coverage.

**Regression sheet exception.** `write_regression_output_sheet` calls `safe_activate(sheet)` for initial activation, but its freeze-panes block keeps its own inline `try/except` rather than `safe_freeze_top_row` — it freezes the top **three** rows (freeze at A4, matching the three-row header band of labels, control/status, and column headers), where `safe_freeze_top_row` only freezes one. Follow this sheet's own pattern (clear `FreezePanes`/`Split`, `sheet.activate()` / `sheet.range("A4").select()` / `ActiveWindow.FreezePanes = True` inside a bare `try/except Exception: pass`) if a future sheet needs a multi-row freeze; don't route it through `safe_freeze_top_row`, which is single-row only.

**Freeze at the selection, never via `SplitRow`.** Every freeze must be done the way Excel's UI does it — select the cell below the frozen band, then set `ActiveWindow.FreezePanes = True` — and clear any stale `FreezePanes`/`Split` first so re-freezing a rebuilt sheet is idempotent. Setting `SplitRow` turns the window's `Split` on, and freezing a split persists in the saved workbook as `state="frozenSplit"` (draggable split bars over a pinned header) instead of a true `state="frozen"` — the saved-XML pane states are how this is checked (`<pane … state="frozen"/>` vs `state="frozenSplit"`).

### Static reference sheets — regenerate via `rebuild_static_sheets.py`, not the per-module CLI

`write_sheet_regression_instructions.py`, `write_sheet_modeling_concepts.py` and `write_sheet_diagnostic_guide.py` write their content (`_ROWS` / `_write_template_sheet`) only into `templates/static_sheets.xlsx`; `build_production.py` never executes that content — it copies the already-baked sheet via `copy_static_sheet`. Editing those modules has **zero effect on any build** until the template is regenerated and committed. After editing any of them, run **`python scripts/rebuild_static_sheets.py`** and commit the updated `templates/static_sheets.xlsx` alongside the Python change. The per-module CLIs still exist for single-sheet debugging, but using them instead of the combined script is the failure mode it exists to prevent. Full rationale: `CONTRIBUTING.md` → *Static reference sheets*.

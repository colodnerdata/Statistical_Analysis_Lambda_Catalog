# AGENTS.md — Project context for AI agents

> **What this file is.** Session primer — the rules you cannot infer from code, read once per session. Deep reference lives in `CONTRIBUTING.md` (build/verify/test-model workflows, flag tables, conventions) and `docs/` (planning: `ROADMAP.md`, `MODEL_TESTING_ASSETS.md`, `ARCHITECTURE.md`, `DECISIONS.md`). Where a section is a paraphrase of one of those, a one-line pointer is the canonical entry.

## PR workflow

**Reply to PR comments after committing fixes.** When a review comment (Copilot or human) leads to a fix, push the fix and then reply to that comment explaining what was changed and why. This is how the repo owner spots addressed comments and resolves the threads.

**Automated verification.** Run `python scripts/build_production.py --verify --no-launch` to build and verify the unified workbook. The spec-driven verifier is **not in CI** — GitHub-hosted `windows-latest` lacks Microsoft Office, so `xw.App` fails with `pywintypes.com_error: (-2147221005, 'Invalid class string')`. Layer 1 (`poe verify-headless`) runs on every push. Full pipeline + flag tables: `CONTRIBUTING.md` → *Verifying builds*.

**Recalculate mode is unconditional.** The production constructor always runs `CalculateFullRebuild` and saves in full **Automatic** — there is no skip-calculation flag. (Regression's rebuild is cheap; Univariate's is slower because it materializes the Beta `Full_Factorial` spills, but it always runs.) Full flag tables: `CONTRIBUTING.md` → *Production build*.

## Testing regime

**The regression test-model suite is planned in `docs/MODEL_TESTING_ASSETS.md`.** That
document is the plan of record: which model configurations the QC harness covers, which
corner each one exists for, the coverage matrix, and the datasets future milestones need.
Read it before adding or changing a QC model case; add to it before adding a case it does
not list. `CONTRIBUTING.md` → *The regression test-model suite* has the step-by-step for
adding one.

**The four-rule PR shape (feature + oracle + test-model case + transcript in `excel-only-runs/`) and the work-in-progress exception:** `CONTRIBUTING.md` → *The PR-shape rules — what every Regression PR must contain*. The plan of record for test-model cases is `docs/MODEL_TESTING_ASSETS.md`.

**Covering-array regime (target ~25–30 fittable models + ~10 guard states, no full crosses):** `CONTRIBUTING.md` → *The regime, in four rules*.

**A case is a `RegressionSpecCase` in `lambda_catalog/analyze_regression_spec.py`** (not a sheet fixture), with expected values from `calculate_regression_spec_case` (NumPy/statsmodels — reading the cell back is not an oracle). **Non-default dataset: set `source_csv_path`, `row_loader`, and `source_table_ref` together** — `Source_Table` is the one name that retargets the data sheet, and omitting it lands the spec rows on the wrong columns silently. A guard-rail configuration is a `GuardStateCase` in `analyze_regression_guard_states.py` instead (its cases raise in the spec oracle by design and assert status text + Design Columns audit + CF fires, not fit stats). **`GuardFlag` is a predicate, not a pixel** — recompute the rule condition, never read `DisplayFormat.Interior.Color`. Full convention: `CONTRIBUTING.md` → *The regime, in four rules*.

**Every case materializes as its own sheet in `Lambda_Library_TestModels.xlsx`** (gitignored; built by `scripts/build_test_models.py`, verified by `tools/inspect_test_model_sheets.py` read-only). **Sheet names** are governed by `lambda_catalog/test_model_sheets.py` (31 chars, legal charset, `<PlanID> <Concept>`, unique across model + guard cases) and are validated at registry-build time, so a bad name fails in the unit suite rather than mid-build. **The name states the concept under test, never the variables** (`M05 Log-Log NA Masking`, not `MPG ~ Ln(Weight) + Ln(HP)`). Both halves of the sheet contract — `apply_spec_case` / `set_prediction_inputs` (write) and `read_case_comparison_rows` (read) — live in `lambda_catalog/regression_spec_sheet_io.py` and are shared with the legacy single-sheet verifier so the two cannot disagree.

**The spec block has no fixed height, and no `SpecTable`.** Every part of it sizes itself from
`COLUMNS(Source_Data)`, so retargeting `Source_Table` resizes the block — the one-name edit the
Instructions sheet has always promised. Three mechanisms, all in
`write_spec_block.py`:

* the `Spec_*` bands are `=TAKE($X$4:$X$16000,MAX(1,COLUMNS(Source_Data)))`, built by `_spec_band`
  — `TAKE` not `OFFSET`, for the same non-volatility reason `Source_Data` and `Header_Names` use it;
* the four computed columns (J Period In Use, K Levels, L Reference In Use, O Design Columns) are
  each ONE spill at `_FIRST_DATA_ROW`, `MAP(SEQUENCE(nc),LAMBDA(i,…))`, written with `f`
  (**`Formula2`** — `.Formula` enters a dynamic array as a legacy CSE range, which does not resize);
* the input band's `INPUT_COLOR` fill is a single lowest-priority CF rule on the same predicate,
  not per-row `format_input` calls.

**Nothing may be written below the spec block in columns B–O**: the bands reach row 16000 and the
spills need clear space, so stray content is a `#SPILL!` error, not a truncation. The one ceiling —
`_SPEC_BAND_LAST_ROW`, the same 16000 the validations and CF rules use — is deliberately shared so
they cannot disagree about how far the block can grow.

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

`compare_values` takes an explicit `scale` for the response-derived cases and
floors every divisor at 1.0, so the convention can only loosen a comparison,
never tighten one. **Derive the divisor from the fit, never a constant** — a
response in the tens and one in the billions must be treated proportionately.

**Adding a compared statistic means choosing its case.** The residual band is
the inherited-error one: every statistic there is built from the predictions,
so it carries the response's precision floor whatever its own size (a residual
of order 0.1 carries the error of numbers of order 70, and self-scaling divides
it by 1.0). `T_Statistics` is deliberately in none of the response-derived sets
— it is dimensionless and O(1) and its error comes from the coefficient, so a
response-derived divisor would be a number picked to fit. Full rationale:
`docs/DECISIONS.md` → *QC comparison scale, the clear-list invariant, and the
OLS solver*.

**The OLS oracle is pinned to `method="qr"`** in `_fit_ols_model`, not the
statsmodels `"pinv"` default: the oracle is the reference the sheet is scored
against, so it should be the more accurate side, and LINEST is QR-based too.

## The Transform column has two Log tokens

`None` · `Log` · `Log (drop ≤ 0)`, from `_TRANSFORM_LOG` / `_TRANSFORM_LOG_DROP`
in `write_spec_block.py`. **Both build the identical `Ln(x)` column and both
report `"Log"` to `Constructed_Column_Transforms()`** — the unit-space / Duan
dispatcher sees one transform, not two, which is why a second token cost the
test suite two cases instead of a multiplier. They differ in exactly one thing:
what happens to a row whose value is zero or negative.

* **`Log`** leaves the row in the sample. `Ln_Positive` returns `#N/A` and it
  propagates through every statistic — the fit is dead, not degraded. The
  Transform cell goes red and `G2` names the variable, the count and the fix.
* **`Log (drop ≤ 0)`** is the only thing that adds the positivity term to
  `Sample_Include`. `G2` reports the excluded-row count in amber.

Filtering never happens because the workbook decided to — narrowing the sample
changes the model, so it is a declaration in the spec. Don't "fix" `Log` by
making it filter.

**Every Excel-side test of "is this a Log row?" goes through `_is_log(expr)`,
and its Python mirror is `_logs(transform)` in `analyze_regression_spec.py`.**
Never compare a Transform cell against one token. `Sample_Include` is the sole
deliberate exception — it tests `= "Log (drop ≤ 0)"` alone, because plain `Log`
not filtering is the whole distinction.

The token also lives literally inside `lambda_functions.json`, which no import
can reach; `test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform`
pins the two spellings together across all seven readers so a rename cannot
half-land. `Log_Domain_Status` is the second deliberate exception to the
pairing rule, alongside `Sample_Include` and in the opposite direction: it
equality-tests only the STRICT token, because the rows it counts are the ones
a strict Log leaves in the sample to poison the fit, and it names the drop
token as the remedy in prose rather than as a comparison.

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

Import pattern in a sheet writer:

```python
from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR
```

The `as _NAME` alias keeps local helper functions (`_subheader_row`, etc.) unchanged.

Sheet-specific colors that differ from the shared palette (e.g., `_SUBHEADER_COLOR` in `write_sheet_diagnostic_guide.py`) remain as local constants in the relevant file.

**Exception — the Regression spec block's header row uses `HEADER_COLOR`, not
`SUBHDR_COLOR`.** By the table above it is a column sub-header row and would
take `SUBHDR_COLOR`; it is deliberately pinned to `HEADER_COLOR` (`#CAEDFB`)
with black bold text in `_write_spec_block`, because that row heads the sheet's
primary *input* surface and reads as a zone heading rather than a subdivision of
one. `test_spec_block_prefills_the_t0_default_configuration` asserts all three
properties (fill, color, bold), so a future writer that reintroduces a competing
style cannot quietly take them back.

### Section heading style

A **section heading** is bold text with `HEADER_COLOR` fill at the default font size. The sheet title ("Univariate Analysis") is 14 pt bold with no fill — that is the only cell with a custom font size.

The shared helper `section_heading(sheet, row, col, label)` in `lambda_catalog/workbook_helpers.py` applies this style; `write_sheet_regression.py` and `write_sheet_univariate.py` both import and call it. Use that helper; do not apply the style inline.

### Univariate sheet heading hierarchy

Zones 1–4 (cols A–Z) use the standard row layout:

| Row | Content | Style |
|---|---|---|
| 1 | "Univariate Analysis" (A1), "Histograms" (F1:M1 merged) | Title: 14 pt bold; Histograms: `section_heading` |
| 2 | "Sturges Method" (F2:G2), "Scott Method" (I2:J2), "Freedman-Diaconis Method" (L2:M2), "Distribution Fitting/Comparison" (P2:Z2) | `section_heading` + merged |
| 3 | "Data" (A3), "Descriptive Statistics" (C3) | `section_heading` |
| 4 | Column sub-headers ("Upper Edge", "Count", "Distribution", …) | `_subheader_row` |
| 5+ | Data / spill formulas | — |

### Univariate right-hand band — one ordered zone table

Everything right of the histograms is **one zone per topic**, each followed by a single gap column, all derived from the ordered `_BAND_ZONES` table via `_derive_band_columns()`. Zone starts (`_C_QQ`, `_C_GS_WB`, `_C_GS_GAMMA`, `_C_GS_BETA`), the gap columns, and the sheet's last column all come from it. **Never hard-code a column letter in this band** — reordering it means reordering that list and nothing else.

| Zone | Columns | Width | Rows |
|---|---|---|---|
| 5 — Q-Q plot data | BE–BN | 10 | — |
| 6 — Weibull fit | BP–BS | 4 | 1–52 |
| 6 — Gamma fit | BU–BX | 4 | 1–52 |
| 6 — Beta fit | BY–CD | 6 | 1+ (N² body) |

**The fit zones must stay last.** They are the only zones whose width is a tunable (Beta's 6 cols reflect its two-stages-side-by-side Alpha|Beta|NLL layout; editing N resizes the body **downward**, not the band **width**, but keeping the fits at the end still means a width change displaces nothing), so keeping them at the end means a resize displaces nothing. Same principle as the Regression sheet's rule that nothing may sit right of the design-matrix zone. **Beta follows Gamma with no gutter** (per-zone `gutter_after` flag on `_BAND_ZONES`; `gamma`→False) — the one deliberate exception to the uniform-gutter rule, matching the owner's hand-laid reference.

### Univariate fit zones

**One zone per distribution, with that fit's two stages side by side inside it** (S1 left, S2 right). Stacking is not viable: every body is a spill whose height follows an in-sheet `Grid Points` cell, so a stacked Stage 2's row anchor would depend on Stage 1's dynamic spill height. Side-by-side stages both start at the same body row and grow down independently, so the dynamic size just works. **Every stage of every fit has the same shape** — a `Full_Factorial` spill beside a `BYROW` NLL column that reads it through the `#` operator, so the NLL is always exactly as tall as the grid it scores. **The three fits use two stage writers**, both sharing one row skeleton per zone: control field-list rows 4–(≤12), profile-NLL chart rows 13–30, body header row 32, body rows 33+.

**Beta — `_write_beta_fit`, a 2-D `Full_Factorial` grid plus a BYROW NLL column** in a 6-column (`BY:CD`) zone, two stages side by side. Each stage's body is **two** spills at row 33: a `Full_Factorial` grid spill (N²×2, `Alpha | Beta`) across the stage's first two cols — `=Full_Factorial($<Ncell>,VSTACK($<αmin>$,$<βmin>$),VSTACK($<αmax>$,$<βmax>$))` — and a `BYROW` NLL spill in the third col that reads the grid via the `#` operator: `=LET(…,z,(d-MIN(d)+pad)/scale_,BYROW($<grid>$#,LAMBDA(ab,NLL_Beta(z,INDEX(ab,1,1),INDEX(ab,1,2))+COUNT(d)*LN(scale_))))`. Two spills rather than one lets the grid stand alone as the pure Cartesian product and the NLL read it by reference. Stage 1 grid at BY33 (cols BY,BZ) + NLL at CA33; Stage 2 grid at CB33 (cols CB,CC) + NLL at CD33, refined bounds (α/β Min/Max = best±step from Stage 1). Recovery is `Grid_Argument_Minimum` over the materialized NLL column. The control block is a vertical field-list with **two value columns per stage** (α and β side by side) so the block fits in ≤9 rows and clears the chart space at row 13:

| Row | Contents |
|---|---|
| 1 | stage title (unmerged) |
| 3 | stage sub-headers (S1 / S2) |
| 4 | Grid Points (N) — S1 literal default 10, S2 `=<S1 Ncell>`; N editable live |
| 5–7, 9 | α / β Min, Max, Step Size (S1 wide scope; S2 = best±step) |
| 8 | Min NLL (`=IFERROR(TAKE(Grid_Argument_Minimum(UV_BETA_S{1,2}_NLL),,1),"—")`) |
| 10 | Min NLL value (Stage 2 anchor) |
| 11–12 | Best α / Best β (`=IFERROR(INDEX(UV_BETA_S{1,2}_Alpha,INDEX(Grid_Argument_Minimum(UV_BETA_S{1,2}_NLL),1,2)),"—")` / `_Beta`) |
| 32 | body headers: `Alpha` / `Beta` / `NLL` per stage |
| 33+ | `Full_Factorial` N²×3 spill per stage |

**Weibull and Gamma — `_write_profile_fit`, a 1-D `Full_Factorial` axis plus a BYROW profile-NLL column** in a 4-column (`_PS_W`) zone. The scale / rate parameter is profiled out in closed form (`_weibull_profile_scale`, `_gamma_profile_rate`), so a stage is N evaluations, not N². The axis is `=Full_Factorial($<Ncell>,$<Min>,$<Max>)` — the d=1 reduction of Beta's grid — and the NLL beside it is `=LET(x,FILTER(UV_Data,UV_Include),BYROW($<axis>$#,LAMBDA(r,LET(p,INDEX(r,1,1),IFERROR(<NLL at p, partner profiled out>,1E+15)))))`. Three details are load-bearing: `INDEX(r,1,1)` scalarizes the 1×1 row `BYROW` hands the callback; `IFERROR` sits **inside** the `LAMBDA`, so one non-evaluable trial value costs its own row instead of collapsing the column; and `x` is bound once per stage, not once per row. The control block is a vertical field-list (label col + one value col per stage); both stages' bodies sit side by side on shared rows:

| Row | Contents |
|---|---|
| 1 | stage title (unmerged) |
| 3 | stage sub-headers (S1 wide / S2 refined) |
| 4–9 | Grid Points, Min, Max, Start, Min NLL, Step Size — S1 vals in col1, S2 vals in col2 |
| 10–11 | Optimal Shape (searched) / Optimal Scale (profiled-out closed form) |
| 32 | body headers: `Shape (k)` / `Profile NLL` per stage |
| 33+ | `Full_Factorial` axis and `BYROW` profile NLL, side by side (N rows, default 20) |

Both writers share the `_GS_R_*` / `_PS_R_*` row offsets. Each fit has its own Best cells and Stage 2 anchor (one zone per fit), so **`_final_grid_best_refs` is per-distribution**, reading `_STAGE2_ANCHORS`. Side-by-side stages ⇒ Stage 2's Best cells are at **fixed control rows** in the Stage 2 value col (no N² row offset), so the old `beta_grid_size**2` row-offset logic is gone. `_STAGE2_ANCHORS` is the single point of agreement between the fitting table and the search writers — the table cannot reference a block the writers did not produce.

Row and column positions come from the `_BAND_ZONES` table and the `_GS_R_*` / `_PS_R_*` / `_PR_C_*` constants at the top of `write_sheet_univariate.py` — never hard-code them inside either stage writer. Beta's visible Alpha and Beta Input cells are its `Full_Factorial` bound cells.

**`Grid Points` (N) is a live cell in all three fits.** Editing it resizes that stage's grid (N² rows for Beta, N for a profile fit) and its NLL column follows through the `#` reference, and the OFFSET named ranges — `UV_BETA_S{1,2}_{Alpha,Beta,NLL}`, `UV_{WB,GAMMA}_S{1,2}` and their `_Axis` partners, `UV_Profile_*` — are all sized `MAX(IFERROR(<N cell>,1),1)` (squared for Beta) so Min NLL, the Best-parameter recovery, the boundary guard, and the charts all track the new height. Each cell carries a whole-number `Validation` and a red conditional format at the `MIN_GRID_POINTS` floor (2 — the Step cell divides by N−1); `build_common.positive_grid_size` applies the same floor to `--beta-grid-size`. Static formatting (number formats, colour scale, border box) is painted over the **default-size** window — `_PROFILE_BODY_CF_ROWS_CAP` / `_BETA_BODY_CF_ROWS_CAP` — so rows a live increase adds beyond it stay unshaded; cosmetic, and deliberate.

A 1-D stage cannot use `Grid_Search_Optimum` — on a single-column grid its column-parameter half reads the cell above the body, which is the `Profile NLL` header, not a parameter value. Use `INDEX(<axis name>, INDEX(Grid_Argument_Minimum(<body name>),1,2))`, as `_write_profile_fit` does. Beta uses `Grid_Argument_Minimum` over its materialized NLL column (`UV_BETA_S{1,2}_NLL`) for Min NLL / Best α / Best β recovery — safe because the NLL column is already materialized before the name reads it.

Zone 5 (Q-Q plot data) holds Hazen plotting positions `P`, the sorted `Sample` column, and the per-distribution theoretical-quantile columns referencing the fit-table parameter cells. Charts occupy the band under the fitting table — histogram combo charts and per-distribution Q-Q scatter charts fed by OFFSET-based `UV_QQ_*` named ranges. The two Weibull / Gamma profile-NLL charts are anchored **above the body**, at rows 13–30, one zone wide (Weibull `BP13:BS30`, Gamma `BU13:BX30` — between the control block and the body, matching the hand-laid reference), and each plots *both* stages from `UV_Profile_<dist>_<S1|S2>_<Axis|NLL>` OFFSET names (which point at the live-N body starting at row 33), with `+` markers on Stage 2 so the refined region stays visible where it overlaps the wide Stage 1 curve. **Beta has no chart yet**: `BY13:CD30` is left blank, reserved for a potential future Beta chart.

### Regression sheet heading hierarchy

Row 1 holds the top-level zone labels ("MODEL SPECIFICATION", "PREDICTOR SUMMARY", "REGRESSION OUTPUTS", "PREDICTION OUTPUTS", "RESIDUAL OUTPUT"). Lower section headings appear at the relevant data rows within each zone. The MODEL SPECIFICATION zone (A–O) is the shared spec block imported from `write_spec_block.py` (headers row 3, spec rows from `_FIRST_DATA_ROW` to `_LAST_DATA_ROW` — currently rows 4–15, sized to `len(_VARIABLES)`; H = Sequence structural flag, I = Sequence Period (typed override input), J = Period In Use (candidate-with-override display), K/L = Levels / Reference In Use displays, M/N = the reserved Interaction Term / Interaction Operation pair, O = the Design Columns audit with its Σ total at O1); every other zone keeps headers on row 2 with spills from row 3. Every `_C_*` column constant in `write_sheet_regression.py` matches its actual column letter (`_C_N` is column N).

**Rows 1–2 above the spec block have one grammar: row 1 is labels, row 2 is the control, value or status — and every status sits in the spec column it is about.** Role cardinality at B2, the Log domain at G2, Sequence cardinality at H2, the spacing verdict at I2, the design-matrix width guard at O2 (under its own Σ total, not at M2 above Interaction Term). Status cells carry no row-1 label: they are blank whenever the spec is legal, and the message names its own subject when it appears. Only readouts are labelled — Intercept (C1/C2), the FE trio (J1:L2), Σ Design Columns (N1/O1), Δ/Count (P3/Q3 with the spectrum spilling from P4, aligned with the spec block's own header and data rows). Because each status owns a column it has almost no overflow runway, so every one of them is `WrapText` with a short imperative message and a hover Note carrying the full guidance; row 2 stays on automatic height, so it is one line while the spec is legal and grows when a message fires. Readouts whose feature is not in play are hidden white-on-white with their labels (`_hide_when`) rather than printing `"n/a"`. Add a new status by putting it on row 2 of its own column — never by parking it in whichever cell happens to be free, which is what the old E1/B1/M2 scatter was. **The four status cells hold a CALL, not the logic** (reunify Part 6.2): B2/G2/H2/O2 are `=Role_Status()` / `=Log_Domain_Status()` / `=Sequence_Status()` / `=Design_Width_Status()`, sheet-scoped LAMBDAs in `lambda_functions.json`, so clicking the cell shows what is being checked instead of a wall of nested IFs. A new status is therefore a new sheet-scoped catalog entry, not an inline formula in a writer. What stays in Python is what the catalog cannot carry: the long hover Notes (the catalog's `notes` is capped at 255 characters), the CF rules, and the count sub-formulas other cells share. Because no import can reach a JSON string literal, `tests/test_spec_block_writer.py` pins each body's message text to the guard-state oracle and `tests/test_sheet_writers.py` pins the width thresholds to the `regression_layout.py` constants that also size the design-matrix band. I2's spacing verdict is still inline — it is the one status the extraction did not cover.

**Never spell an A1 address into a formula string.** Conditional-formatting expressions, chart titles, and OFFSET-based named ranges all need addresses as literal text, and hand-written letters are what turn a column insertion into a silent-wrong-answer bug — the formula still parses, it just reads a different cell. Build every one of them from the `_C_*` constants via the `_abs_ref(row, col)` / `_band(col)` helpers and the `_A_*` anchors (`_A_ALPHA`, `_A_OBSERVATIONS`, `_A_MEAN_LEVERAGE`, …) at the top of `write_sheet_regression.py`. The same rule applies to anything reading the sheet: `tools/inspect_regression_sheet.py` and `lambda_catalog/analyze_regression_spec_block.py` IMPORT the column constants rather than keeping a parallel copy.

**Column-layout paradigm — gap columns and outline groups.** The zones are Model Specification (A–Q, including the P/Q Δ-spectrum feedback columns), Predictor Summary (S–Y), Regression Outputs (AA–AH — `AG3:AH9` holds the v3.3 **Unit-Space Fit** block: `Duan`/`Naive` back-transformation toggle at `AH4`, smearing factor R5, R²/Adj R²/RMSE in original units R6–8, response-space readout R9), Prediction Outputs (AJ–AL — `AL` is the **Original Units** column: back-transformed point estimate at AL3, Naive-only CI/PI bounds at AL7–10; the Duan/Naive caveat is a NOTE on the Back-Transform label at `AG4`, not a row in this zone), and Residual Output (AN–BA — the two new columns `AZ`/`BA` carry Predicted Y / Residual in original units, dispatched on the `AH4` toggle). Between every pair of adjacent zones sits exactly one dedicated **gap column** (R, Z, AI, AM — width 2) that is deliberately left OUT of every outline group. That ungrouped column is what makes the neighbouring zones collapse independently: Excel fuses a contiguous run of same-level grouped columns into one outline, so two zones with no ungrouped column between them would share a single collapse control. `_ZONES` (the (first, last) content spans) and `_GAP_COLUMNS` (derived as the single column between consecutive zones, asserted one wide) are the single source of truth; `_COLUMN_GROUPS = _ZONES`, and the gap columns are sized and left ungrouped in the width/grouping loop of `write_regression_output_sheet`. When adding or resizing a zone, edit `_ZONES` — never hard-code an outline group or a gap letter.

**Content-column widths are keyed on the `_C_*` constants too — `_COLUMN_WIDTHS`.** It is `(constant, width)` pairs with a module-level assertion that every zone content column is sized exactly once and no gap column is sized at all, so the next shift fails at import instead of shipping silently with the wrong column. The spec block (A–O) keeps its own widths in `write_spec_block._SPEC_COLUMN_WIDTHS`; column I (the Regression-only Verdict overlay) and BB (the post-zone chart gutter) are the two deliberate entries outside the zones.

The **Model Formula readout** sits on row 1 of the terminal Constructed Design Matrix zone (header at `_C_DESIGN_MATRIX + 2`, readout one column past it, both derived from the zone anchor), with `WrapText = False` — row 1 is the one row the matrix itself can never reach (its names spill on row 2, values on row 3, both rightward), so the caption is never displaced. The cell holds `=Model_Formula()` (the sheet-scoped catalog closure), and v3.4 reads the NAME (`Comparison_Model_Formula` registered in `_setup_local_names`) — that is why moving the readout cost its consumer nothing.

**Past the charts sits the ARCHITECTURE §4b materialization band**, on the same gutter-per-zone principle: the **Model Context** block (a two-column label/value pair — `_C_MODEL_CONTEXT_LABEL` / `_C_MODEL_CONTEXT` — one labelled cell per context element, headed and border-boxed because its height is a build-time constant, and grouped as a pair so it collapses as a unit), the reserved `Sample_Include` column, and the terminal **Constructed Design Matrix** zone. **Model Context is the only zone in this band that is grouped or collapsed** — it is a block of individual cells, so hiding it hides no spill. `Sample_Include` and the design matrix are left ungrouped and expanded on purpose: both are full-height dynamic arrays, and a collapsed outline group over a spill range is the state in which Excel stops recomputing the model — the hidden columns keep the stale arrays and every engine reading across them refits on stale values. Do not re-add a `Group()` / `ShowDetail` call for either one; the scrolling of an expanded unbounded zone is the accepted cost. Its columns derive from `_LAST_CHART_COLUMN`, which tracks the chart anchor, so a zone shift moves the whole band automatically. **Nothing may ever be placed to the right of the design-matrix zone** — its width is one dropdown away from hundreds of columns, so any zone after it would be displaced by an ordinary modelling choice.

**The model context is individual cells, never a `VSTACK` spill.** `_MODEL_CONTEXT_ELEMENTS` in `regression_layout.py` (written by `_write_materialization_zone` in `regression_materialization.py`) is the single source of the row order, the displayed labels, and the block height (`_MODEL_CONTEXT_ROWS`, and from it `Fit_Context`'s fixed range and the `Context OK` health-check row). A spill buys nothing here — the height is a build-time constant, not data-dependent — and costs correctness: one formula producing four cells is a single dependency node that Excel vacates and re-spills whenever the spec block changes, and while it is vacated the range behind `Fit_Context()` is transiently blank, so all ~30 engine call sites read a torn context. Don't reintroduce one.

### Regression chart named ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns degrades recalculation performance. All chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in the Regression Statistics block at `$AB$8`. Example: `RegChartFitY = =OFFSET('<sheet>'!$AP$2,1,0,MAX(IFERROR('<sheet>'!$AB$8,1),1),1)` — starts one row below the column header, extends exactly `$AB$8` rows, with the `MAX(IFERROR(...,1),1)` guard keeping it one row tall when `$AB$8` cannot resolve. `CONTRIBUTING.md` → *Chart series data ranges* has the full `sheet.api.Names.Add` example and the `_name_ref` helper.

All OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix. This distinguishes them from the constructor closures (`Predictor_Columns`, `Design_Columns`, `Sample_Include`, etc.) and formula-helper names. The name-to-column map (and the `$AB$8` anchor) lives in the loop in `_setup_local_names` in `lambda_catalog/write_sheet_regression.py` — that loop is the single source of truth for the column letters. When adding a new diagnostic column or chart, add the corresponding `RegChart`-prefixed named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

Worksheet-scoped names are created via `sheet.api.Names.Add` on the owning sheet, and `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer.

### Workbook scope belongs to the catalog

Every named range a sheet writer creates is **sheet-scoped** — `sheet.api.Names.Add` on the owning sheet, never `book.names.add`. `RegChart*`, `UV_*` and the spec wiring all live there. Workbook scope belongs exclusively to the catalog's LAMBDA functions (`lambda_functions.json`, `scope: "workbook"`), and `sync_workbook_names` enforces that literally: on every build it drops **every** workbook-scoped `<definedName>` that is not a catalog function or one of Excel's reserved `_xlnm.*` names, then rewrites the catalog entries.

**Workbook-scoped catalog LAMBDAs must be sheet-agnostic.** A body that hardcodes `'Regression'!` is wrong in two directions: in a workbook with SEVERAL Regression-shaped sheets (the test-model artifact has 47) every sheet reads whichever one is literally named `Regression`, and in a workbook with none it is skipped and every call site reads `#NAME?`. `Base_Period_Delta` is the example: it is **sheet-scoped** (`"scope": "Regression"`) with unqualified spec references, so an unqualified name resolves against the sheet the calling formula lives on and each Regression sheet gets its own Δ. The skip mechanism (a catalog function that names a missing worksheet is skipped, not written) is what stops the next sheet-qualified body shipping a broken link. Full history: `CONTRIBUTING.md` → *Workbook scope belongs to the catalog*.

`tests/test_workbook_invariants.py::TestRealWorkbookNameScope` asserts all of this against the committed `dist/Lambda_Library.xlsx` on every commit (pure zipfile, no Excel), as does `TestRealWorkbook` alongside it — every committed-artifact check in that file is always-on, so a stale or hand-edited workbook fails the suite instead of shipping. To re-apply the cleanup to a built artifact without rebuilding it: `python tools/resync_workbook_names.py <workbook.xlsx>`.

## Charts — patterns and pitfalls

### Use xlwings COM API for all chart creation — never openpyxl

openpyxl's `load_workbook()`/`save()` **rewrites the entire .xlsx package** and silently drops chart parts, VML drawings, and chartUserShapes it didn't create. Loading a workbook that already has Excel-created charts (e.g., the Regression diagnostic charts) and saving it back will destroy those charts — this is a fundamental openpyxl limitation, not fixable. All charts in this project use `sheet.api.ChartObjects().Add(...)` via xlwings COM. Follow the existing pattern in `regression_charts.py:_write_diagnostic_charts`.

### Chart creation pattern

```python
co = sheet.api.ChartObjects().Add(left, top, width, height)
chart = co.Chart

# Clear any default series
while chart.SeriesCollection().Count > 0:
    chart.SeriesCollection(1).Delete()

chart.ChartType = _XL_COLUMN_CLUSTERED   # or _XL_XY_SCATTER
chart.HasLegend = False
chart.HasTitle = True

# Data series reference worksheet-scoped named ranges
series = chart.SeriesCollection().NewSeries()
series.XValues = f"='{sheet.name}'!NamedRangeEdges"
series.Values  = f"='{sheet.name}'!NamedRangeCounts"
```

### Chart titles — `.Text` for static, `.Formula` for cell-linked

`.Text` sets a literal string and is correct for fixed titles. `.Formula` links the title to a worksheet cell so it updates dynamically:

```python
# Static title (fixed label):
chart.ChartTitle.Text = "Residuals vs. Fitted"

# Dynamic title (linked to a formula cell):
chart.ChartTitle.Formula = "='Sheet'!$Q$14"
```

When a dynamic title needs to be computed (e.g., concatenating a method name with " Histogram"), write the formula into a dedicated cell first, then point the chart title's `.Formula` at that cell. Do **not** pass a formula string to `.Text` — it will be rendered as literal text, not evaluated.

### Histogram-specific formatting

- `chart.ChartGroups(1).GapWidth = 0` — histogram bars must be contiguous (no gap).
- Do **not** enable "Vary colors by point" — histograms use a single uniform bar color.
- Add explicit axis titles (`x_axis.AxisTitle.Text = "Upper Edge"`, `y_axis.AxisTitle.Text = "Count"`).
- Title `overlay=False` so Excel auto-sizes the plot area without overlap.

### Chart positioning

Define chart positions via xlwings `sheet.range(...)` to get `.left`, `.top`, `.width`, `.height` in points. This ties chart placement to the cell grid so charts stay aligned if column widths change.

### Guard chart creation with try/except

Charts require the Excel COM API, which is unavailable in CI/headless environments. Always wrap chart insertion so the sheet build succeeds without charts:

```python
try:
    _write_histogram_charts(sheet)
except Exception:
    pass
```

### Never draw reference lines as shapes — use a real data series

Do not use `chart.Shapes.AddLine(...)` to fake a reference line like `y=x`. A shape is positioned in fixed plot-area pixel coordinates computed at creation time and silently goes wrong when the chart is resized, moved, or its axis scaling changes.

Instead, add a real data series pointing both `XValues` and `Values` at the same named range:

```python
series = chart.SeriesCollection().NewSeries()
series.XValues = name_ref   # e.g. ='Sheet'!RegChartFitY
series.Values = name_ref    # same range — guarantees every point sits on y=x
series.Name = "Identity"
series.ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS
```

See `_add_identity_line` in `regression_charts.py`.

### Selective data labels — a masked overlay series, not per-point COM loops

To label only the points that meet some value-based criterion (e.g., Cook's Distance points above the `F.INV(0.5, p, n-p)` influence cutoff — see `_COOKS_CUTOFF`), add a helper column that returns the real value for qualifying rows and a masking token for everything else, expose it as its own `RegChart`-prefixed named range, and add it to the chart as an extra series with `HasDataLabels = True`. The mask is what makes the selection — no per-point `Points(i).HasDataLabel` loop, and no reading calculated values back into Python during the sheet-writing phase (which runs under `XL_CALCULATION_MANUAL` and would see stale or unfit values; see "Sheet writer conventions" below).

**Which masking token depends on where the label text comes from, and the two choices are not interchangeable.** With `ShowValue` / `ShowCategoryName` the token is `NA()`: Excel skips `NA()` points for both plotting and labeling, so the unflagged rows disappear on their own. With **Value From Cells** — `DataLabels.Format.TextFrame2.TextRange.InsertChartField(msoChartFieldRange, "='Sheet'!SomeName", 0)` followed by `ShowRange = True` — the label prints the source cell verbatim, so an `NA()` row renders a literal `#N/A`; the token has to be `""`. That trade has a tail: a `""`-returning formula is plotted (at zero), not skipped, so **every other label element must be turned off** (`ShowValue = False`, `ShowCategoryName = False`) or it prints on all n observations instead of the flagged few. The Cook's Distance overlay is the Value-From-Cells case: `AY` is `=IF(AT3#>cutoff,AT3#,"")` and the range supplies the entire label.

On a **column-chart** target, do not give the overlay series the chart's own `xlColumnClustered` type — a second column series joins the cluster group and narrows/shifts the real bars, misaligning any label from the bar it annotates. Instead set the overlay series' own `ChartType = xlLine` (constant `4`) with `Format.Line.Visible = False` and `MarkerStyle = xlMarkerStyleNone (-4142)`: a Line-type series shares the same category axis as a Column series without joining its cluster, so it overlays exactly in place. Setting a per-series `ChartType` that differs from the chart's own is how Excel builds a **combo chart** — expect the chart to become one.

The overlay series' `XValues` still points at a named range over the observation-identifier column (`RegChartObsLabel`), which is what its categories read as. Displaying that identifier in the label — `ShowCategoryName = True`, combined with `ShowValue` if the number should show too — is only available under `NA()` masking, for the reason above: under `""` masking it would print on every point.

See the `RegChartCookDistFlag` / `RegChartObsLabel` names in `_setup_local_names` and the Cook's Distance branch of `_write_diagnostic_charts` in `regression_charts.py`.

### Separate chart title cells from chart insertion

Write formula cells for chart titles (e.g., `Q14`, `Q34`, `Q54`) **outside** the try/except guard. These are standard cell writes (not COM chart API calls), so they can be exercised in unit tests via the `RecordingSheet` mock without Excel. Only the `ChartObjects().Add(...)` call needs the guard.

### Build-phase retry separation

Do not wrap the entire `build_production_workbook()` call in a single retry loop. The recalculate/save step is fast (~10s) and is the most likely to fail when the user opens the workbook to inspect progress. Give it its own retry phase so a failure there doesn't restart the multi-minute sheet-writing phase. See `_retry_on_open` and the two-phase `main()` in `build_production.py`.

## Sheet writer conventions

- **Layout single-source-of-truth**: row constants (`_ROW_TITLE`, `_ROW_DATA_START`, …) at the top of each writer; chart series via `OFFSET` named ranges (CONTRIBUTING.md § *Chart series data ranges*); `app.api.Calculation = XL_CALCULATION_MANUAL` before writes, `XL_CALCULATION_SEMIAUTOMATIC` after, before save.

### Guard headless/no-focus Excel calls with the `safe_*` helpers

`Sheet.activate()` and anything touching `Application.ActiveWindow` (e.g. freezing panes) raise when Excel cannot become the active application — no interactive desktop session, focus denied by the OS, an agentic/headless build host, etc. — even though the workbook write itself succeeds. Which sheet is on top or whether panes are frozen when the file opens is cosmetic, so that failure must not abort `build_production_workbook()`.

Use `safe_activate(sheet)` and `safe_freeze_top_row(sheet)` from `workbook_helpers.py` instead of calling `sheet.activate()` / touching `ActiveWindow` directly — each helper is a `try/except Exception: pass` wrapper, so the cosmetic failure (sheet not on top / panes not frozen) does not abort `build_production_workbook()`. When adding a new sheet writer that activates its sheet or freezes its header row, call the `safe_*` helper, not the raw xlwings/COM call. See `tests/test_workbook_helpers.py` for the stub-based unit coverage.

**Regression sheet exception.** `write_regression_output_sheet` calls `safe_activate(sheet)` for the initial activation, but its freeze-panes block keeps its own inline `try/except` rather than calling `safe_freeze_top_row` — it freezes the top **two** rows (`SplitRow = 2`, matching the sheet's two-row header), where `safe_freeze_top_row` only freezes one. Follow this sheet's own pattern (`sheet.activate()` / `sheet.range("A3").select()` / `ActiveWindow.FreezePanes` inside a bare `try/except Exception: pass`) if a future sheet needs a multi-row freeze; don't route it through `safe_freeze_top_row`, which is single-row only.

### Static reference sheets — regenerate via `rebuild_static_sheets.py`, not the per-module CLI

`write_sheet_regression_instructions.py` and `write_sheet_diagnostic_guide.py` write their content (`_ROWS` / `_write_template_sheet`) only into `templates/static_sheets.xlsx`; `build_production.py` never executes that content — it only copies the already-baked sheet via `copy_static_sheet`. Editing those modules has **zero effect on any build** until the template is regenerated and committed. After editing either module's content, run **`python scripts/rebuild_static_sheets.py`** and commit the updated `templates/static_sheets.xlsx` alongside the Python change. The per-module CLIs still exist for single-sheet debugging but using them instead of the combined script is the failure mode the combined script exists to prevent. Full rationale: `CONTRIBUTING.md` → *Static reference sheets*.

# AGENTS.md — Project context for AI agents

## PR workflow

**Reply to PR comments after committing fixes.** When a review comment (Copilot or human) leads to a fix, push the fix and then reply to that comment explaining what was changed and why. This is how the repo owner spots addressed comments and resolves the threads.

**Automated verification.** Use `python scripts/build_production.py --verify --no-launch` to build and verify the Regression workbook in one shot during regression-focused iteration, or `python scripts/build_univariate.py --verify --no-launch` for the standalone Univariate workbook. The spec-driven verifier reuses `build_qc.verify_test_sheets(..., skip_dummy=True)` (with `skip_univariate=True` for the Regression artifact and `skip_regression=True` for the Univariate artifact); on drift it prints a structured `VerifyReport` and `sys.exit(1)`, so a stale build never opens in Excel. The fast headless screen (`make verify-headless`, pure `zipfile` + `lxml`) needs no Excel and is auto-discovered by the Linux CI job, so it runs on every push. **The spec-driven verifier is not run in CI** — the GitHub-hosted `windows-latest` image does not include Microsoft Office, so xlwings fails to dispatch `Excel.Application` (`Invalid class string`). It is a developer-machine step until a self-hosted runner with Office is wired in. See `CONTRIBUTING.md` → *Verifying builds* for the full pipeline.

**Always recalculate the Regression workbook.** The build's final `CalculateFullRebuild` is what recomputes the Regression engines after a name sync; the verifier only does a per-sheet `Calculate()`, which doesn't rebuild the dependency tree, so skipping the rebuild leaves every QC value reading `nan`. The Regression workbook has no Data Tables, so the rebuild is cheap — `build_production.py` always runs it regardless of `--skip-data-table-calculations` (which is a no-op for the Regression artifact). The Univariate workbook is different: Beta's two two-input Data Tables (the only Data Tables in the artifact — Weibull and Gamma are 1-D profile-NLL columns, the other five fits are closed-form) make `CalculateFullRebuild` slow, so `build_univariate.py` runs the rebuild by default (so the shipped artifact's fits are not stale) but honors `--skip-data-table-calculations` to skip it for fast iteration — this is that flag's now-primary purpose. A third mode, `--no-calculation`, calculates *nothing*: it never sets Automatic (which would calculate the open workbook on the spot — the cost `--skip-data-table-calculations` does not avoid), suppresses `Application.CalculateBeforeSave` for the duration and restores it after, and skips the rebuild. It is for inspecting the name manager or the layout, and its output is saved in Manual mode with stale cells — never ship an artifact built with it.

## Testing regime

**The regression test-model suite is planned in `docs/MODEL_TESTING_ASSETS.md`.** That
document is the plan of record: which model configurations the QC harness covers, which
corner each one exists for, the coverage matrix, and the datasets future milestones need.
Read it before adding or changing a QC model case; add to it before adding a case it does
not list. `CONTRIBUTING.md` → *The regression test-model suite* has the step-by-step for
adding one.

**Every Regression-track feature ships with four things in the same PR.** A new
functionality targeted at the Regression engine (any ROADMAP milestone inside the
Regression block of the ladder) must land as one merged PR containing: (1) the
feature itself; (2) an oracle in `calculate_regression_spec_case` /
`calculate_guard_state_case`; (3) at least one new test-model case in
`docs/MODEL_TESTING_ASSETS.md` § 1 + the `RegressionSpecCase` / `GuardStateCase`
registry, pinned in `_EXPECTED_CASE_NAMES` or `_EXPECTED_GUARD_NAMES`; and (4) a
transcript in `excel-only-runs/` for the new test-model sheet, archived by
`lambda_catalog.build_common.run_log_path` and committed. The spec-driven verifier
cannot run on the GitHub-hosted Linux CI without Excel, so the transcript is the
only verifiable paper trail the PR leaves for a reviewer or a future agent. A PR
without the transcript is incomplete — ask the contributor for it before
reviewing the other three. See [CONTRIBUTING.md → *The regression test-model
suite*](CONTRIBUTING.md#the-regression-test-model-suite) and the four-rule block
at the top of `CLAUDE.md` § *Testing regime* for the full text.

**It is a covering array, not a full factorial.** Every implemented corner is exercised by
at least one model, and every model earns its place by covering something no other model
does — target ~25–30 fittable models plus ~10 guard-state configurations. Do not add a case
that duplicates an existing one's coverage, and do not propose full crosses (every transform
× every interaction × every role): they multiply sheet count without adding information.

**A case is a `RegressionSpecCase` in `lambda_catalog/analyze_regression_spec.py`**, not a
sheet fixture — `SpecVariable` rows built with `_spec_var(...)`, expected values from
`calculate_regression_spec_case` (NumPy/statsmodels, never read back from the workbook). A
case targeting a dataset other than Auto MPG must set `source_csv_path`, `row_loader`, **and**
`source_table_ref` together; `Source_Table` is the one name that retargets which data sheet
the spec block reads, so omitting it lands the spec rows on the wrong table's columns with no
error.

**Every case name is pinned.** `_EXPECTED_CASE_NAMES` in `tests/test_regression_spec_qc.py`
is an ordered list asserted against `build_regression_spec_cases()`. Adding, renaming,
reordering, or dropping a case means editing that list in the same commit — otherwise the
suite fails, which is the point.

**The suite's growth rate orders the roadmap — under one constraint.** From v3.4 on the ladder
sorts on two keys. First, **all remaining Regression work ships before either milestone that
opens a new analysis surface**: v3.10 Two-sample and v3.11 Resampling go last as a block, even
though both are flat-cost to test. Second, *within* the Regression track, milestones are
sequenced by how much they force this suite to grow — additive (v3.4 Model Comparison),
near-additive (v3.5 `Cluster`, v3.6 `Time`), ~2× (v3.7 WLS, v3.8 two-way FE), and the ~10×
axis-widener last (v3.9 the standalone transform library). The inversion is deliberate: v3.10
and v3.11 are cheaper to cover than four milestones ahead of them and still ship last. Do not
"fix" it back to pure test-scale order. When a change would reorder milestones, edit the
MODEL_TESTING_ASSETS Section 2 table first and the ROADMAP ladder second — the ladder follows
that table, not the other way round.

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
one. The pin has to be applied **after** `_create_spec_table`: creating the
ListObject applies a TableStyle that overrides header fill and font, so styling
it earlier is silently undone. `test_spec_block_prefills_the_t0_default_configuration`
asserts all three properties so a future TableStyle change cannot quietly take
them back.

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
| 6 — Weibull fit | BP–BX | 9 | 1–31 |
| 6 — Gamma fit | BZ–CH | 9 | 1–31 |
| 6 — Beta fit | CJ–DD | 21 | 1–51 |

**The fit zones must stay last.** They are the only zones whose width is a tunable (`_N_GRID`, `_N_PROFILE` — and Beta's pending shrink to a ~12×12 grid will change its width by nine columns), so keeping them at the end means a resize displaces nothing. Same principle as the Regression sheet's rule that nothing may sit right of the design-matrix zone. Zones 5 and 6 swapped numbers when Q-Q moved ahead of the fits, so the numbering still reads left to right.

### Univariate fit zones

**One zone per distribution, with that fit's two stages stacked inside it** — not one band per stage. **The three fits use two different stage writers.**

**Beta — `_write_grid_stage`, a 2-D two-input Data Table** spanning 21 columns (1 row-axis col + 20 body cols). These two stages are the artifact's **only** Data Tables. Stage 2 sits `_GS_R_STAGE2` (26) rows below Stage 1 — a full grid block plus a gap row:

| dr | Row | Contents |
|---|---|---|
| 0 | row 1 | stage title merged across c0:c0+20 with `_HEADER` fill |
| 1 | row 2 | `Min NLL` (c0), `Rows/Columns` (c0+1), blank spacer (c0+2), parameter headers (c0+3:c0+8) |
| 2 | row 3 | Min NLL and grid-count values; Alpha row: `Parameter | Input | Min | Max | Step Size | Best` |
| 3 | row 4 | Beta row in the same six-column parameter table |
| 4 | row 5 | corner NLL cell (c0); Alpha SEQUENCE spills right across 20 columns |
| 5–24 | rows 6–25 | Beta SEQUENCE (c0); Data Table body (c0+1:c0+20) |

**Weibull and Gamma — `_write_profile_stage`, a 1-D profile-NLL column** in a 9-column (`_PS_W`) zone. The scale / rate parameter is profiled out in closed form (`_weibull_profile_scale`, `_gamma_profile_rate`), so a stage is 20 evaluations, not 400. **A profile stage's control block and its body are positioned independently**: the two stages' control blocks stack vertically (`_PS_R_STAGE_STRIDE` = 5 rows apart) while their two bodies sit side by side on shared rows, each under one half of the control block, reusing its offset-2 spacer (`_PS_BODY_COLS`):

| dr | Row | Contents |
|---|---|---|
| 0 / 5 | rows 1, 6 | Stage 1 / Stage 2 title merged across c0:c0+8 with `_HEADER` fill |
| 1 / 6 | rows 2, 7 | `Min NLL` (c0), `Grid Points` (c0+1), blank spacer (c0+2), parameter headers (c0+3:c0+8) |
| 2 / 7 | rows 3, 8 | Min NLL and point-count values; searched row: `Parameter | Start | Min | Max | Step Size | Best` |
| 3 / 8 | rows 4, 9 | profiled-out row — `Parameter` and `Best` only; it is solved, not searched, so it has no bounds, no step, and no boundary rule |
| 10 | row 11 | body headers: Stage 1 axis + `Profile NLL` (c0+0/c0+1), Stage 2 axis + `Profile NLL` (c0+3/c0+4) |
| 11–30 | rows 12–31 | both stages' searched-parameter SEQUENCE and profile NLL, side by side |

Both writers **share the `_GS_R_*` row offsets within a stage**, and `_PS_C_*` mirrors `_GS_C_*` — only the meaning of offset 4 differs (Data Table substitution `Input` vs. closed-form `Start`). What they no longer share is a Best *column*: with one zone per fit, each distribution's Best column and Stage 2 row differ, so **`_final_grid_best_refs` is per-distribution**, reading `_STAGE2_ANCHORS`. That dict is the single point of agreement between the fitting table and the search writers — the table cannot reference a block the writers did not produce.

Row and column positions come from the `_BAND_ZONES` table and the `_GS_R_*` / `_GS_C_*` / `_PS_C_*` / `_PS_R_*` constants at the top of `write_sheet_univariate.py` — never hard-code them inside either stage writer. Beta's visible Alpha and Beta Input cells are its Data Table substitution cells. `Rows/Columns` is generated from `_N_GRID` (`Grid Points` from `_N_PROFILE`) and documents the physical body size; editing it does not resize the Data Table.

A 1-D stage cannot use `Grid_Search_Optimum` — on a single-column grid its column-parameter half reads the cell above the body, which is the `Profile NLL` header, not a parameter value. Use `INDEX(<axis name>, INDEX(Grid_Argument_Minimum(<body name>),1,2))`, as `_write_profile_stage` does.

Zone 5 (Q-Q plot data) holds Hazen plotting positions `P`, the sorted `Sample` column, and the per-distribution theoretical-quantile columns referencing the fit-table parameter cells. Charts occupy the band under the fitting table — histogram combo charts and per-distribution Q-Q scatter charts fed by OFFSET-based `UV_QQ_*` named ranges. The two Weibull / Gamma profile-NLL charts are **not** in that band: each is anchored under its own fit zone (BP33, BZ33 — one clear row below the bodies, one zone wide), and each plots *both* stages from `UV_Profile_<dist>_<S1|S2>_<Axis|NLL>`, with `+` markers on Stage 2 so the refined region stays visible where it overlaps the wide Stage 1 curve.

### Regression sheet heading hierarchy

Row 1 holds the top-level zone labels ("MODEL SPECIFICATION", "PREDICTOR SUMMARY", "REGRESSION OUTPUTS", "PREDICTION OUTPUTS", "RESIDUAL OUTPUT"). Lower section headings appear at the relevant data rows within each zone. The MODEL SPECIFICATION zone (A–O) is the shared spec block imported from `write_sheet_model_construction.py` (headers row 3, spec rows from `_FIRST_DATA_ROW` to `_LAST_DATA_ROW` — currently rows 4–15, sized to `len(_VARIABLES)` — Intercept control A2/C2, Sequence status line H2; H = Sequence structural flag, I = Sequence Period (typed override input), J = Period In Use (candidate-with-override display), K/L = Levels / Reference In Use displays, M/N = the reserved Interaction Term / Interaction Operation pair, O = the Design Columns audit with its Σ total at O1 and the width-guard status at M2); every other zone keeps headers on row 2 with spills from row 3. Every `_C_*` column constant in `write_sheet_regression.py` matches its actual column letter (`_C_N` is column N).

**Never spell an A1 address into a formula string.** Conditional-formatting expressions, chart titles, and OFFSET-based named ranges all need addresses as literal text, and hand-written letters are what turn a column insertion into a silent-wrong-answer bug — the formula still parses, it just reads a different cell. Build every one of them from the `_C_*` constants via the `_abs_ref(row, col)` / `_band(col)` helpers and the `_A_*` anchors (`_A_ALPHA`, `_A_OBSERVATIONS`, `_A_MEAN_LEVERAGE`, …) at the top of `write_sheet_regression.py`. The same rule applies to anything reading the sheet: `tools/inspect_regression_sheet.py` and `lambda_catalog/analyze_regression_spec_block.py` IMPORT the column constants rather than keeping a parallel copy.

**Column-layout paradigm — gap columns and outline groups.** The zones are Model Specification (A–Q, including the P/Q Δ-spectrum feedback columns), Predictor Summary (S–Y), Regression Outputs (AA–AH — `AG3:AH9` holds the v3.3 **Unit-Space Fit** block: `Duan`/`Naive` back-transformation toggle at `AH4`, smearing factor R5, R²/Adj R²/RMSE in original units R6–8, response-space readout R9), Prediction Outputs (AJ–AL — `AL` is the **Original Units** column: back-transformed point estimate at AL3, Naive-only CI/PI bounds at AL7–10; the Duan/Naive caveat is a NOTE on the Back-Transform label at `AG4`, not a row in this zone), and Residual Output (AN–BA — the two new columns `AZ`/`BA` carry Predicted Y / Residual in original units, dispatched on the `AH4` toggle). Between every pair of adjacent zones sits exactly one dedicated **gap column** (R, Z, AI, AM — width 2) that is deliberately left OUT of every outline group. That ungrouped column is what makes the neighbouring zones collapse independently: Excel fuses a contiguous run of same-level grouped columns into one outline, so two zones with no ungrouped column between them would share a single collapse control (the bug this layout fixes — the predictor summary used to begin at M, hard against the spec block, fusing the two outlines). `_ZONES` (the (first, last) content spans) and `_GAP_COLUMNS` (derived as the single column between consecutive zones, asserted one wide) are the single source of truth; `_COLUMN_GROUPS = _ZONES`, and the gap columns are sized and left ungrouped in the width/grouping loop of `write_regression_output_sheet`. When adding or resizing a zone, edit `_ZONES` — never hard-code an outline group or a gap letter.

The v3.3 shift that took Residual Output from `AN–AY` to `AN–BA` is a content change, not a layout break: the chart anchor moved from `_C_AZ` to `_C_BB` (the two new columns were inserted *before* it, so the chart anchor and the entire §4b materialization band derive right by 2 and BP at column 68 is preserved). The Model Formula cell at `AA2:AB2` heads the Regression Outputs zone. The sheet-scoped `Comparison_Anchor`, `Comparison_Headline_GoF`, and `Comparison_Model_Formula` named ranges are the v3.4 Model Comparison sheet's reading surface — registered in `_setup_local_names` next to the `RegChart*` entries.

**Past the charts sits the ARCHITECTURE §4b materialization band**, on the same gutter-per-zone principle: the **Model Context** block (a two-column label/value pair — `_C_MODEL_CONTEXT_LABEL` / `_C_MODEL_CONTEXT` — one labelled cell per context element, headed and border-boxed because its height is a build-time constant, and grouped as a pair so it collapses as a unit), the reserved `Sample_Include` column, and the terminal **Constructed Design Matrix** zone, which ships collapsed. Its columns derive from `_LAST_CHART_COLUMN`, which tracks the chart anchor, so a zone shift moves the whole band automatically. **Nothing may ever be placed to the right of the design-matrix zone** — its width is one dropdown away from hundreds of columns, so any zone after it would be displaced by an ordinary modelling choice.

**The model context is individual cells, never a `VSTACK` spill.** `_MODEL_CONTEXT_ELEMENTS` in `write_sheet_regression.py` is the single source of the row order, the displayed labels, and the block height (`_MODEL_CONTEXT_ROWS`, and from it `Fit_Context`'s fixed range and the `Context OK` health-check row). A spill buys nothing here — the height is a build-time constant, not data-dependent — and costs correctness: one formula producing four cells is a single dependency node that Excel vacates and re-spills whenever the spec block changes, and while it is vacated the range behind `Fit_Context()` is transiently blank, so all ~30 engine call sites read a torn context. Don't reintroduce one.

### Regression chart named ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns degrades recalculation performance. All chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in the Regression Statistics block at `$Y$8`:

```python
sheet.api.Names.Add(
    Name="RegChartFitY",
    RefersTo=f"=OFFSET('{sname}'!$AM$2,1,0,MAX(IFERROR('{sname}'!$Y$8,1),1),1)",
)
```

All `RegChart`-prefixed names are worksheet-scoped (created via `sheet.api.Names.Add`). Chart `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer. Use the `_name_ref` helper:

```python
def _name_ref(local_name: str) -> str:
    return f"='{sname}'!{local_name}"
```

All OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix. This distinguishes them from the constructor closures (`Predictor_Columns`, `Design_Columns`, `Sample_Include`, etc.) and formula-helper names. The post-v2.0 name-to-column map (and the `$Y$8` anchor) lives in the loop in `_setup_local_names` in `lambda_catalog/write_sheet_regression.py` — that loop is the single source of truth for the column letters. When adding a new diagnostic column or chart, add the corresponding `RegChart`-prefixed named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

### Workbook scope belongs to the catalog

Every named range a sheet writer creates is **sheet-scoped** — `sheet.api.Names.Add` on the owning sheet, never `book.names.add`. `RegChart*`, `UV_*` and the spec wiring all live there. Workbook scope belongs exclusively to the catalog's LAMBDA functions (`lambda_functions.json`, `scope: "workbook"`), and `sync_workbook_names` enforces that literally: on every build it drops **every** workbook-scoped `<definedName>` that is not a catalog function or one of Excel's reserved `_xlnm.*` names, then rewrites the catalog entries.

That rule exists because the v3.0 split left each artifact carrying the other one's ranges at workbook scope — twelve `RegChart*` entries reading `OFFSET(#REF!,…)` in the Univariate workbook, forty-two `UV_*` entries in the Regression workbook. The live sheet-scoped names were correct throughout; it was the stale workbook-scoped copies that produced the broken links. Don't reintroduce a narrower rule that inspects the body text (an error literal *inside* an `OFFSET(...)` is not an error-literal body) — "workbook-scoped and not in the catalog" is the exact characterization of residue.

A catalog function whose body names a worksheet the target workbook does not have is **skipped**, not written: Excel rebinds such a reference to an external workbook (`Regression!Source_Data` → `[1]!Source_Data`), writes an `xlPathMissing` external-link part, and prompts about broken links on every open. `Base_Period_Delta` is the one such function today — Regression-coupled by design, so the Univariate artifact does not carry it, and the build prints the skip. When adding a workbook-scoped catalog LAMBDA, keep it sheet-agnostic unless you intend it to be Regression-only.

`tests/test_workbook_invariants.py::TestRealWorkbookNameScope` asserts all of this against both committed artifacts on every commit (pure zipfile, no Excel). To re-apply the cleanup to a built artifact without rebuilding it, run `python tools/resync_workbook_names.py <workbook.xlsx>`.

## Charts — patterns and pitfalls

### Use xlwings COM API for all chart creation — never openpyxl

openpyxl's `load_workbook()`/`save()` **rewrites the entire .xlsx package** and silently drops chart parts, VML drawings, and chartUserShapes it didn't create. Loading a workbook that already has Excel-created charts (e.g., the Regression diagnostic charts) and saving it back will destroy those charts. This is not fixable — it's a fundamental openpyxl limitation.

The zipfile-patching workaround (generating chart XML in throwaway openpyxl workbooks, then splicing into the real zip) is fragile, hard to test, and was ultimately abandoned. **Don't go down this path.**

All charts in this project use `sheet.api.ChartObjects().Add(...)` via xlwings COM, which writes directly into the live Excel instance without round-tripping the file. Follow the existing pattern in `write_sheet_regression.py:_write_diagnostic_charts`.

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

See `_add_identity_line` in `write_sheet_regression.py`.

### Selective data labels — an `NA()`-masked overlay series, not per-point COM loops

To label only the points that meet some value-based criterion (e.g., Cook's Distance points above the standard `4/n` or `0.9` influence cutoffs), add a helper column that returns the real value for qualifying rows and `NA()` for everything else, expose it as its own `RegChart`-prefixed named range, and add it to the chart as an extra series with `HasDataLabels = True`. Excel skips `NA()` points for both plotting and labeling, so only the flagged points render a label — no per-point `Points(i).HasDataLabel` loop, and no reading calculated values back into Python during the sheet-writing phase (which runs under `XL_CALCULATION_MANUAL` and would see stale or unfit values; see "Sheet writer conventions" below).

On a **column-chart** target, do not give the overlay series the chart's own `xlColumnClustered` type — a second column series joins the cluster group and narrows/shifts the real bars, misaligning any label from the bar it annotates. Instead set the overlay series' own `ChartType = xlLine` (constant `4`) with `Format.Line.Visible = False` and `MarkerStyle = xlMarkerStyleNone (-4142)`: a Line-type series shares the same category axis as a Column series without joining its cluster, so it overlays exactly in place. Setting a per-series `ChartType` that differs from the chart's own is how Excel builds a **combo chart** — expect the chart to become one.

To label the point by something more meaningful than its raw value (e.g., the observation's row identifier instead of, or alongside, the Cook's D number), set the overlay series' `XValues` to a named range over the identifier column and turn on `ShowCategoryName` on its `DataLabels()`, combined with `ShowValue` if the number should show too.

See the `RegChartCookDistFlag` / `RegChartObsLabel` names in `_setup_local_names` and the Cook's Distance branch of `_write_diagnostic_charts` in `write_sheet_regression.py`.

### Separate chart title cells from chart insertion

Write formula cells for chart titles (e.g., `Q14`, `Q34`, `Q54`) **outside** the try/except guard. These are standard cell writes (not COM chart API calls), so they can be exercised in unit tests via the `RecordingSheet` mock without Excel. Only the `ChartObjects().Add(...)` call needs the guard.

### Build-phase retry separation

Do not wrap the entire `build_production_workbook()` call in a single retry loop. The recalculate/save step is fast (~10s) and is the most likely to fail when the user opens the workbook to inspect progress. Give it its own retry phase so a failure there doesn't restart the multi-minute sheet-writing phase. See `_retry_on_open` and the two-phase `main()` in `build_production.py`.

## Sheet writer conventions

- Row constants (`_ROW_TITLE`, `_ROW_METHOD_HDR`, `_ROW_SECTION_HDR`, `_ROW_COL_HDRS`, `_ROW_DATA_START`, …) live at the top of each writer and are the single source of truth for the layout.
- All chart series reference **worksheet-scoped named ranges** via `OFFSET` — never spill references or full-column references (see CONTRIBUTING.md for details).
- Set `app.api.Calculation = XL_CALCULATION_MANUAL` before writing any sheet and `XL_CALCULATION_SEMIAUTOMATIC` after all writes, before save. This prevents OFFSET-based named ranges from resolving prematurely during build.

### Guard headless/no-focus Excel calls with the `safe_*` helpers

`Sheet.activate()` and anything touching `Application.ActiveWindow` (e.g. freezing panes) raise when Excel cannot become the active application — no interactive desktop session, focus denied by the OS, an agentic/headless build host, etc. — even though the workbook write itself succeeds. Which sheet is on top or whether panes are frozen when the file opens is cosmetic, so that failure must not abort `build_production_workbook()`.

Use `safe_activate(sheet)` and `safe_freeze_top_row(sheet)` from `workbook_helpers.py` instead of calling `sheet.activate()` / touching `ActiveWindow` directly — every sheet writer that used to call these unguarded (lambda functions, life expectancy, mileage, production lots, dummy test, the three MLR test sheets) now goes through these two helpers, each a `try/except Exception: pass` wrapper. When adding a new sheet writer that activates its sheet or freezes its header row, call the `safe_*` helper, not the raw xlwings/COM call. See `tests/test_workbook_helpers.py` for the stub-based unit coverage.

**Regression sheet exception.** `write_regression_output_sheet` calls `safe_activate(sheet)` for the initial activation, but its freeze-panes block keeps its own inline `try/except` rather than calling `safe_freeze_top_row` — it freezes the top **two** rows (`SplitRow = 2`, matching the sheet's two-row header), where `safe_freeze_top_row` only freezes one. Follow this sheet's own pattern (`sheet.activate()` / `sheet.range("A3").select()` / `ActiveWindow.FreezePanes` inside a bare `try/except Exception: pass`) if a future sheet needs a multi-row freeze; don't route it through `safe_freeze_top_row`, which is single-row only.

### Static reference sheets — regenerate via `rebuild_static_sheets.py`, not the per-module CLI

`write_sheet_regression_instructions.py` and `write_sheet_diagnostic_guide.py` write their content (`_ROWS` / `_write_template_sheet`) only into `templates/static_sheets.xlsx`; `build_production.py`/`build_qc.py` never execute that content — they only copy the sheet already baked into the template (`copy_static_sheet`). Editing `_ROWS` or `_write_template_sheet` has **zero effect on any build** until the template is regenerated and committed. This has already shipped stale sheet text twice from someone forgetting (or only partially doing) that step — see `DECISIONS.md` → v2.2 "Static template drift".

After editing either module's content, run `python scripts/rebuild_static_sheets.py` (regenerates every static sheet in one Excel session) and commit the updated `templates/static_sheets.xlsx` alongside the Python change. Don't reach for the older per-module CLIs (`python -m lambda_catalog.write_sheet_regression_instructions`, `python -m lambda_catalog.write_sheet_diagnostic_guide`) as the primary path — they still work for regenerating one sheet in isolation while debugging, but using them instead of the combined script is exactly the failure mode this script exists to prevent. See `CONTRIBUTING.md` → "Static reference sheets" for the full rationale.

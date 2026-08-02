# CLAUDE.md — Project context for Claude Code

## PR workflow

**Reply to PR comments after committing fixes.** When a review comment (Copilot or human) leads to a fix, push the fix and then reply to that comment explaining what was changed and why. This is how the repo owner spots addressed comments and resolves the threads.

**Automated verification.** Use `python build_production.py --verify --no-launch` to build and verify the Regression workbook in one shot during regression-focused iteration, or `python build_univariate.py --verify --no-launch` for the standalone Univariate workbook. The spec-driven verifier reuses `build_qc.verify_test_sheets(..., skip_dummy=True)` (with `skip_univariate=True` for the Regression artifact and `skip_regression=True` for the Univariate artifact); on drift it prints a structured `VerifyReport` and `sys.exit(1)`, so a stale build never opens in Excel. The fast headless screen (`make verify-headless`, pure `zipfile` + `lxml`) is auto-discovered on Linux once it lands. **The spec-driven verifier is not run in CI** — the GitHub-hosted `windows-latest` image does not include Microsoft Office, so xlwings fails to dispatch `Excel.Application` (`Invalid class string`). It is a developer-machine step until a self-hosted runner with Office is wired in. See `CONTRIBUTING.md` → *Verifying builds* for the full pipeline.

**Always recalculate the Regression workbook.** The build's final `CalculateFullRebuild` is what recomputes the Regression engines after a name sync; the verifier only does a per-sheet `Calculate()`, which doesn't rebuild the dependency tree, so skipping the rebuild leaves every QC value reading `nan`. The Regression workbook has no Data Tables, so the rebuild is cheap — `build_production.py` always runs it regardless of `--skip-data-table-calculations` (which is a no-op for the Regression artifact). The Univariate workbook is different: its grid searches (two two-input Data Tables for Beta, plus Weibull/Gamma static formula grids — ~2,400 NLL evaluations per recalc in total) make `CalculateFullRebuild` slow, so `build_univariate.py` runs the rebuild by default (so the shipped artifact's fits are not stale) but honors `--skip-data-table-calculations` to skip it for fast iteration — this is that flag's now-primary purpose.

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

Zone 5 holds the two-stage grid searches (Weibull / Gamma / Beta, vertically stacked). Each stage (`_write_grid_stage`) spans 21 columns (1 row-axis col + 20 Data Table body cols):

| dr | Row | Contents |
|---|---|---|
| 0 | row 1 | stage title merged across c0:c0+20 with `_HEADER` fill |
| 1 | row 2 | `Min NLL` (c0), `Rows/Columns` (c0+1), blank spacer (c0+2), parameter headers (c0+3:c0+8) |
| 2 | row 3 | Min NLL and grid-count values; Shape row: `Parameter | Input | Min | Max | Step Size | Best` |
| 3 | row 4 | Scale row in the same six-column parameter table |
| 4 | row 5 | corner NLL cell (c0); Shape SEQUENCE spills right across 20 columns |
| 5–24 | rows 6–25 | Scale SEQUENCE (c0); Data Table body (c0+1:c0+20) |

Column letters and row anchors are defined as `_C_GS`, `_C_GS_S2`, and the `_GS_R_*` / `_GS_C_*` constants at the top of `write_sheet_univariate.py` — never hard-code row or column positions inside `_write_grid_stage`; the constants are the single source of truth for the zone layout. The visible Shape and Scale Input cells are the Data Table substitution cells. `Rows/Columns` is generated from `_N_GRID` and documents the physical table size; editing it does not resize the Data Table.

Zone 6 (Q-Q plot data) holds Hazen plotting positions `P`, the sorted `Sample` column, and the per-distribution theoretical-quantile columns referencing the fit-table parameter cells. Charts occupy the band under the fitting table — histogram combo charts and per-distribution Q-Q scatter charts fed by OFFSET-based `UV_QQ_*` named ranges.

### Regression sheet heading hierarchy

Row 1 holds the top-level zone labels ("MODEL SPECIFICATION", "PREDICTOR SUMMARY", "REGRESSION OUTPUTS", "PREDICTION OUTPUTS", "RESIDUAL OUTPUT"). Lower section headings appear at the relevant data rows within each zone. The MODEL SPECIFICATION zone (A–O) is the shared spec block imported from `write_sheet_model_construction.py` (headers row 3, spec rows from `_FIRST_DATA_ROW` to `_LAST_DATA_ROW` — currently rows 4–15, sized to `len(_VARIABLES)` — Intercept control A2/C2, Sequence status line H2; H = Sequence structural flag, I = Sequence Period (typed override input), J = Period In Use (candidate-with-override display), K/L = Levels / Reference In Use displays, M/N = the reserved Interaction Term / Interaction Operation pair, O = the Design Columns audit with its Σ total at O1 and the width-guard status at M2); every other zone keeps headers on row 2 with spills from row 3. Every `_C_*` column constant in `write_sheet_regression.py` matches its actual column letter (`_C_N` is column N).

**Never spell an A1 address into a formula string.** Conditional-formatting expressions, chart titles, and OFFSET-based named ranges all need addresses as literal text, and hand-written letters are what turn a column insertion into a silent-wrong-answer bug — the formula still parses, it just reads a different cell. Build every one of them from the `_C_*` constants via the `_abs_ref(row, col)` / `_band(col)` helpers and the `_A_*` anchors (`_A_ALPHA`, `_A_OBSERVATIONS`, `_A_MEAN_LEVERAGE`, …) at the top of `write_sheet_regression.py`. The same rule applies to anything reading the sheet: `tools/inspect_regression_sheet.py` and `lambda_catalog/analyze_regression_spec_block.py` IMPORT the column constants rather than keeping a parallel copy.

**Column-layout paradigm — gap columns and outline groups.** The zones are Model Specification (A–Q, including the P/Q Δ-spectrum feedback columns), Predictor Summary (S–Y), Regression Outputs (AA–AH), Prediction Outputs (AJ–AL), and Residual Output (AN–AY). Between every pair of adjacent zones sits exactly one dedicated **gap column** (R, Z, AI, AM — width 2) that is deliberately left OUT of every outline group. That ungrouped column is what makes the neighbouring zones collapse independently: Excel fuses a contiguous run of same-level grouped columns into one outline, so two zones with no ungrouped column between them would share a single collapse control (the bug this layout fixes — the predictor summary used to begin at M, hard against the spec block, fusing the two outlines). `_ZONES` (the (first, last) content spans) and `_GAP_COLUMNS` (derived as the single column between consecutive zones, asserted one wide) are the single source of truth; `_COLUMN_GROUPS = _ZONES`, and the gap columns are sized and left ungrouped in the width/grouping loop of `write_regression_output_sheet`. When adding or resizing a zone, edit `_ZONES` — never hard-code an outline group or a gap letter.

**Past the charts sits the ARCHITECTURE §4b materialization band**, on the same gutter-per-zone principle: `Model_Context` (4×1 spill), the reserved `Sample_Include` column, and the terminal **Constructed Design Matrix** zone, which ships collapsed. Its columns derive from `_LAST_CHART_COLUMN`, which tracks the chart anchor, so a zone shift moves the whole band automatically. **Nothing may ever be placed to the right of the design-matrix zone** — its width is one dropdown away from hundreds of columns, so any zone after it would be displaced by an ordinary modelling choice.

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

All OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix. This distinguishes them from the constructor closures (`X_s`, `Sample_Include`, etc.) and formula-helper names. The post-v2.0 name-to-column map (and the `$Y$8` anchor) lives in the loop in `_setup_local_names` in `lambda_catalog/write_sheet_regression.py` — that loop is the single source of truth for the column letters. When adding a new diagnostic column or chart, add the corresponding `RegChart`-prefixed named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

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

After editing either module's content, run `python rebuild_static_sheets.py` (regenerates every static sheet in one Excel session) and commit the updated `templates/static_sheets.xlsx` alongside the Python change. Don't reach for the older per-module CLIs (`python -m lambda_catalog.write_sheet_regression_instructions`, `python -m lambda_catalog.write_sheet_diagnostic_guide`) as the primary path — they still work for regenerating one sheet in isolation while debugging, but using them instead of the combined script is exactly the failure mode this script exists to prevent. See `CONTRIBUTING.md` → "Static reference sheets" for the full rationale.

# AGENTS.md — Project context for AI agents

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

The `as _NAME` alias keeps local helper functions (`_section_heading`, `_subheader_row`, etc.) unchanged.

Sheet-specific colors that differ from the shared palette (e.g., `_SUBHEADER_COLOR` in `write_sheet_diagnostic_guide.py`) remain as local constants in the relevant file.

### Section heading style

A **section heading** is bold text with `HEADER_COLOR` fill at the default font size. The sheet title ("Univariate Analysis") is 14 pt bold with no fill — that is the only cell with a custom font size.

`write_sheet_regression.py` and `write_sheet_univariate.py` each define a private `_section_heading(sheet, row, col, label)` that applies this style. Use those helpers; do not apply the style inline.

### Univariate sheet heading hierarchy

Zones 1–4 (cols A–Z) use the standard row layout:

| Row | Content | Style |
|---|---|---|
| 1 | "Univariate Analysis" (A1), "Histograms" (F1:M1 merged) | Title: 14 pt bold; Histograms: `_section_heading` |
| 2 | "Sturges Method" (F2:G2), "Scott Method" (I2:J2), "Freedman-Diaconis Method" (L2:M2), "Distribution Fitting/Comparison" (P2:Z2) | `_section_heading` + merged |
| 3 | "Data" (A3), "Descriptive Statistics" (C3) | `_section_heading` |
| 4 | Column sub-headers ("Upper Edge", "Count", "Distribution", …) | `_subheader_row` |
| 5+ | Data / spill formulas | — |

Zone 5 (cols AC–BS, with gap columns AB and AX) holds the two-stage Weibull grid-search. Each stage (`_write_grid_stage`) spans 21 columns (1 row-axis col + 20 Data Table body cols):

| dr | Row | Contents |
|---|---|---|
| 0 | row 1 | stage title merged across c0:c0+20 with `_HEADER` fill |
| 1 | row 2 | `Min NLL` (c0), `Rows/Columns` (c0+1), blank spacer (c0+2), parameter headers (c0+3:c0+8) |
| 2 | row 3 | Min NLL and grid-count values; Shape row: `Parameter | Input | Min | Max | Step Size | Best` |
| 3 | row 4 | Scale row in the same six-column parameter table |
| 4 | row 5 | corner NLL cell (c0); Shape SEQUENCE spills right across 20 columns |
| 5–24 | rows 6–25 | Scale SEQUENCE (c0); Data Table body (c0+1:c0+20) |

Stage 1 is `AC1:AW25` with named body `UV_WB_S1 = AD6:AW25`. Stage 2 is `AY1:BS25` with named body `UV_WB_S2 = AZ6:BS25`. The visible Shape and Scale Input cells are the Data Table substitution cells. `Rows/Columns` is generated from `_N_GRID` and documents the physical table size; editing it does not resize the Data Table.

Row and column offsets are defined as `_GS_R_*` and `_GS_C_*` constants at the top of `write_sheet_univariate.py`. Never hard-code row or column positions inside `_write_grid_stage`.

### Regression sheet heading hierarchy

Row 1 holds the top-level zone labels ("MODEL SPECIFICATION", "PREDICTOR SUMMARY", "REGRESSION OUTPUTS", "PREDICTION OUTPUTS", "RESIDUAL OUTPUT"). Lower section headings appear at the relevant data rows within each zone. The MODEL SPECIFICATION zone (A–K) is the shared spec block imported from `write_sheet_model_construction.py` (headers row 3, spec rows 4–26, Intercept control A2/C2, Sequence status line H2; H = Sequence structural flag, I = reserved Base Period Δ, J/K = Levels / Reference In Use displays); every other zone keeps headers on row 2 with spills from row 3. Every `_C_*` column constant in `write_sheet_regression.py` matches its actual column letter (`_C_M` is column M).

### Regression chart named ranges

Chart `SERIES` formulas do not support the `#` spill operator, and referencing full columns degrades recalculation performance. All chart series reference **worksheet-scoped named ranges** defined via `OFFSET` sized to the observation count in `$V$8`:

```python
sheet.api.Names.Add(
    Name="RegChartFitY",
    RefersTo=f"=OFFSET('{sname}'!$AJ$2,1,0,MAX(IFERROR('{sname}'!$V$8,1),1),1)",
)
```

All `RegChart`-prefixed names are worksheet-scoped (created via `sheet.api.Names.Add`). Chart `SERIES` formulas must include the sheet prefix even for worksheet-scoped names, because charts live above the sheet layer. Use the `_name_ref` helper:

```python
def _name_ref(local_name: str) -> str:
    return f"='{sname}'!{local_name}"
```

All OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix. This distinguishes them from the constructor closures (`X_s`, `Sample_Include`, etc.) and formula-helper names. The full set:

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

When adding a new diagnostic column or chart, add the corresponding `RegChart`-prefixed named range in `_setup_local_names` before writing the chart in `_write_diagnostic_charts`.

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

### Separate chart title cells from chart insertion

Write formula cells for chart titles (e.g., `Q14`, `Q34`, `Q54`) **outside** the try/except guard. These are standard cell writes (not COM chart API calls), so they can be exercised in unit tests via the `RecordingSheet` mock without Excel. Only the `ChartObjects().Add(...)` call needs the guard.

### Build-phase retry separation

Do not wrap the entire `build_production_workbook()` call in a single retry loop. The recalculate/save step is fast (~10s) and is the most likely to fail when the user opens the workbook to inspect progress. Give it its own retry phase so a failure there doesn't restart the multi-minute sheet-writing phase. See `_retry_on_open` and the two-phase `main()` in `build_production.py`.

## Sheet writer conventions

- Row constants (`_ROW_TITLE`, `_ROW_METHOD_HDR`, `_ROW_SECTION_HDR`, `_ROW_COL_HDRS`, `_ROW_DATA_START`, …) live at the top of each writer and are the single source of truth for the layout.
- All chart series reference **worksheet-scoped named ranges** via `OFFSET` — never spill references or full-column references (see CONTRIBUTING.md for details).
- Set `app.api.Calculation = XL_CALCULATION_MANUAL` before writing any sheet and `XL_CALCULATION_SEMIAUTOMATIC` after all writes, before save. This prevents OFFSET-based named ranges from resolving prematurely during build.

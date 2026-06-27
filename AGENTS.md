# AGENTS.md — Project context for AI agents

## Cell styling

All cell colors are defined once in `lambda_catalog/sheet_styles.py` and imported by every sheet writer. Never hard-code RGB tuples in a sheet writer.

| Constant | RGB | Usage |
|---|---|---|
| `HEADER_COLOR` | `(202, 237, 251)` | Zone / section headings — same light blue across all sheets |
| `SUBHDR_COLOR` | `(220, 230, 241)` | Column sub-header rows within a zone |
| `INPUT_COLOR`  | `(251, 226, 213)` | User-editable input cells (light orange) |

Import pattern in a sheet writer:

```python
from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR
```

The `as _NAME` alias keeps local helper functions (`_section_heading`, `_subheader_row`, etc.) unchanged.

### Section heading style

A section heading is bold text + `HEADER_COLOR` fill. No custom font size — the sheet title uses 14 pt, all other headings use the default size.

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

Row 1 holds the top-level zone labels ("MODEL SELECTION", "PREDICTOR SUMMARY", "REGRESSION OUTPUTS", "PREDICTION OUTPUTS", "RESIDUAL OUTPUT"). Lower section headings appear at the relevant data rows within each zone.

### Regression chart named ranges

All OFFSET-based named ranges used by diagnostic charts carry the `RegChart` prefix. This distinguishes them from formula-helper names (`All_Xs`, `pred_input`, etc.). The full set:

| Name | Column | Contents |
|---|---|---|
| `RegChartQQX` | AE | Normal Scores Ranked (QQ theoretical axis) |
| `RegChartQQY` | AF | Studentized Residuals Ranked (QQ actual axis) |
| `RegChartFitY` | Y | Predicted Y — shared by multiple charts |
| `RegChartResid` | Z | Residuals |
| `RegChartActY` | X | Actual Y |
| `RegChartScaleLoc` | AG | Scale-Location |
| `RegChartCookDist` | AD | Cook's Distance |
| `RegChartLeverage` | AB | Hat Diagonal |
| `RegChartStudResid` | AC | Studentized Residuals |
| `RegChartPRESSResid` | AH | PRESS Residual |

## Charts — patterns and pitfalls (lessons from PR #67)

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

`.Text` sets a literal string and is correct for fixed titles (e.g., the Regression diagnostic charts use `chart.ChartTitle.Text = title`). `.Formula` links the title to a worksheet cell so it updates dynamically — use it when the title depends on data:

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

Define chart positions via xlwings `sheet.range(...)` to get `.left`, `.top`, `.width`, `.height` in points. This ties chart placement to the cell grid so charts stay aligned if column widths change. See `_write_histogram_charts` for the pattern.

### Guard chart creation with try/except

Charts require the Excel COM API, which is unavailable in CI/headless environments. Always wrap chart insertion so the sheet build succeeds without charts:

```python
try:
    _write_histogram_charts(sheet)
except Exception:
    pass
```

### Separate chart title cells from chart insertion

Write formula cells for chart titles (e.g., `Q14`, `Q34`, `Q54`) **outside** the try/except guard. These are standard cell writes (not COM chart API calls), so they can be exercised in unit tests via the `RecordingSheet` mock without Excel. Only the `ChartObjects().Add(...)` call needs the guard. This separation keeps chart title formulas testable in CI even though actual chart insertion is not.

### Build-phase retry separation

Do not wrap the entire `build_production_workbook()` call in a single retry loop. The recalculate/save step is fast (~10s) and is the most likely to fail when the user opens the workbook to inspect progress. Give it its own retry phase so a failure there doesn't restart the multi-minute sheet-writing phase. See `_retry_on_open` and the two-phase `main()` in `build_production.py`.

## Sheet writer conventions

- Row constants (`_ROW_TITLE`, `_ROW_METHOD_HDR`, `_ROW_SECTION_HDR`, `_ROW_COL_HDRS`, `_ROW_DATA_START`, …) live at the top of each writer and are the single source of truth for the layout.
- All chart series reference **worksheet-scoped named ranges** via `OFFSET` — never spill references or full-column references (see CONTRIBUTING.md for details).
- Set `app.api.Calculation = XL_CALCULATION_MANUAL` before writing any sheet and `XL_CALCULATION_SEMIAUTOMATIC` after all writes, before save. This prevents OFFSET-based named ranges from resolving prematurely during build.

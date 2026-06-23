# AGENTS.md — Project context for Codex

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

## Sheet writer conventions

- Row constants (`_ROW_TITLE`, `_ROW_METHOD_HDR`, `_ROW_SECTION_HDR`, `_ROW_COL_HDRS`, `_ROW_DATA_START`, …) live at the top of each writer and are the single source of truth for the layout.
- All chart series reference **worksheet-scoped named ranges** via `OFFSET` — never spill references or full-column references (see CONTRIBUTING.md for details).
- Set `app.api.Calculation = XL_CALCULATION_MANUAL` before writing any sheet and `XL_CALCULATION_SEMIAUTOMATIC` after all writes, before save. This prevents OFFSET-based named ranges from resolving prematurely during build.

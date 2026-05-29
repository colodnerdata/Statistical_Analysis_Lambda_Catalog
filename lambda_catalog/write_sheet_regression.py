"""
write_sheet_regression.py
Writes the ToolPak-style Regression sheet into any target workbook.

Layout (three horizontal zones):
  Col A          — input labels: "Intercept" in row 2, predictor names spill in rows 3+
  Col B          — input values: intercept=IF(Allow_Intercept,1,0) in row 2 (auto),
                   user-editable predictor values in rows 3+ (k rows)
                   pred_input (B2:B{k+2}) aligns with Coefficients() for SUMPRODUCT
  Cols C–J       — main analysis: Regression Statistics, Diagnostics,
                   Prediction Interval (rows 3-9), ANOVA (rows 10-14),
                   Coefficients (rows 18+)
  Cols L–X       — residual table (fixed columns, independent of k)
  Cols Y–AB      — hidden helpers: filtered actual Y (Y), top predictor (Z),
                   reference line endpoints for charts (AA–AB)
  Cols AC+        — diagnostic charts (2-column grid)
"""
from __future__ import annotations

import xlwings as xw

from .workbook_helpers import group_and_hide_columns


REGRESSION_SHEET_NAME = "Regression"

PREDICTOR_NAMES: list[str] = [
    "Adult Mortality",
    "infant deaths",
    "Alcohol",
    "percentage expenditure",
    "Hepatitis B",
    "Measles",
    "BMI",
    "under-five deaths",
    "Polio",
    "Total expenditure",
    "Diphtheria",
    "HIV/AIDS",
    "GDP",
    "Population",
    "thinness 1-19 years",
    "thinness 5-9 years",
    "Income composition of resources",
    "Schooling",
]

# 1-based column indices
_C_A = 1    # prediction input labels (auto-populated via spill formula)
_C_B = 2    # prediction input values (user-editable)
_C_C = 3    # section labels / ANOVA row labels / coefficient row labels
_C_D = 4    # stat values / ANOVA df / coefficients / Allow_Intercept toggle
_C_E = 5    # ANOVA SS / coefficient standard errors
_C_F = 6    # ANOVA MS / coefficient t-stats / diagnostics labels
_C_G = 7    # ANOVA F / coefficient p-values / diagnostics values
_C_H = 8    # ANOVA Significance F / coefficient CI lower
_C_I = 9    # prediction interval labels / coefficient CI upper
_C_J = 10   # prediction interval values
_C_L = 12   # residual: observation number
_C_M = 13   # residual: predicted Y
_C_N = 14   # residual: residuals
_C_O = 15   # residual: LOOCV prediction
_C_P = 16   # residual: percentile
_C_Q = 17   # residual: Y ranked
_C_R = 18   # residual: normal scores
_C_S = 19   # residual: scaled residuals
_C_T = 20   # residual: scaled residuals ranked
_C_U = 21   # residual: hat diagonal
_C_V = 22   # residual: studentized residuals
_C_W = 23   # residual: studentized residuals ranked
_C_X = 24   # residual: Cook's distance

# Hidden helper columns (grouped and collapsed after writing)
_C_Y = 25   # filtered actual Y — spills from row 3
_C_Z = 26   # top-predictor values — spills from row 3
_C_AA = 27  # reference line X endpoints (rows 1–12, static scalars)
_C_AB = 28  # reference line Y endpoints (rows 1–12, static scalars)

# XL chart constants (XlChartType / XlMarkerStyle / line style)
_XL_XY_SCATTER = -4169         # xlXYScatter: markers only, no connecting lines
_XL_XY_SCATTER_LINES = 75     # xlXYScatterLinesNoMarkers: line, no markers
_XL_MARKER_NONE = -4142        # xlMarkerStyleNone
_XL_MARKER_CIRCLE = 8          # xlMarkerStyleCircle
_XL_DASH = -4115               # xlDash line style
_XL_COLOR_BLUE = 0xC47244      # OLE BGR for Excel theme blue RGB(68, 114, 196)
_XL_COLOR_GRAY = 0x808080      # OLE BGR for gray RGB(128, 128, 128)

_CHART_W = 320.0    # chart width in points
_CHART_H = 240.0    # chart height in points
_CHART_GAP = 20.0   # gap between adjacent charts in points
_REF_MAX_ROW = 3001  # upper row bound for chart series ranges (covers any real dataset)


# ── Cell helpers ──────────────────────────────────────────────────────────────

def _rc(row: int, col: int) -> tuple[int, int]:
    return (row, col)


def _col_letter(col: int) -> str:
    result = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        result = chr(ord("A") + rem) + result
    return result


def _a1(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"


def _v(sheet: xw.Sheet, row: int, col: int, value: object) -> None:
    sheet.range(_rc(row, col)).value = value


def _f(sheet: xw.Sheet, row: int, col: int, formula: str) -> None:
    sheet.range(_rc(row, col)).api.Formula2 = formula


def _bold(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(_rc(row, col)).api.Font.Bold = True


def _bold_row(sheet: xw.Sheet, row: int, col1: int, col2: int) -> None:
    sheet.range(_rc(row, col1), _rc(row, col2)).api.Font.Bold = True


# ── Local name management ─────────────────────────────────────────────────────

def _drop_local_name(sheet: xw.Sheet, name: str) -> None:
    for idx in range(sheet.api.Names.Count, 0, -1):
        local = sheet.api.Names(idx).Name.split("!", 1)[-1]
        if local.lower() == name.lower():
            sheet.api.Names(idx).Delete()


def _setup_local_names(sheet: xw.Sheet, k: int) -> None:
    """Register sheet-scoped names used by every formula on this sheet."""
    for name, ref in [
        ("x_s", "=LifeExpectancyData[[Adult Mortality]:[Schooling]]"),
        ("y",   "=LifeExpectancyData[Life expectancy]"),
        ("fil", "=LifeExpectancyData[Full_Data]"),
    ]:
        _drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=ref)

    # Allow_Intercept toggle lives in D2
    _drop_local_name(sheet, "Allow_Intercept")
    sheet.api.Names.Add(
        Name="Allow_Intercept",
        RefersTo=f"={sheet.name}!$D$2",
    )

    # Prediction input vector: col B, rows 2 to 2+k (k+1 entries: intercept + k predictors)
    # Row 2 holds the intercept term (=1 when Allow_Intercept); rows 3+ hold predictor values.
    # This aligns with Coefficients(...) so =SUMPRODUCT(Coefficients(...),pred_input) is valid.
    _drop_local_name(sheet, "pred_input")
    sheet.api.Names.Add(
        Name="pred_input",
        RefersTo=f"={sheet.name}!$B$2:$B${2 + k}",
    )


# ── Section writers ───────────────────────────────────────────────────────────

def _write_prediction_inputs(sheet: xw.Sheet) -> None:
    _v(sheet, 1, _C_A, "PREDICTION INPUTS")
    _bold(sheet, 1, _C_A)
    # Row 2: intercept term — auto-set to 1 when Allow_Intercept is TRUE
    _v(sheet, 2, _C_A, "Intercept")
    _f(sheet, 2, _C_B, "=IF(Allow_Intercept,1,0)")
    # Rows 3+: predictor name labels (static) and user-editable values
    for i, name in enumerate(PREDICTOR_NAMES):
        _v(sheet, 3 + i, _C_A, name)
        sheet.range(_rc(3 + i, _C_B)).value = 0.0


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    _v(sheet, 3, _C_C, "REGRESSION STATISTICS")
    _bold(sheet, 3, _C_C)
    for row, label, formula in [
        (4, "Multiple R",        "=Multiple_R(x_s,y,Allow_Intercept,fil)"),
        (5, "R Square",          "=R_squared(x_s,y,Allow_Intercept,fil)"),
        (6, "Adjusted R Square", "=Adjusted_R2(x_s,y,Allow_Intercept,fil)"),
        (7, "Standard Error",    "=SE_Regression(x_s,y,Allow_Intercept,fil)"),
        (8, "Observations",      "=Observations(y,fil)"),
    ]:
        _v(sheet, row, _C_C, label)
        _f(sheet, row, _C_D, formula)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    _v(sheet, 3, _C_F, "DIAGNOSTICS")
    _bold(sheet, 3, _C_F)
    for row, label, formula in [
        (4,  "PRESS",          "=PRESS(x_s,y,Allow_Intercept,fil)"),
        (5,  "PRESS R²",       "=1-PRESS(x_s,y,Allow_Intercept,fil)/SS_Total(y,Allow_Intercept,fil)"),
        (6,  "Mean Leverage",  "=(DF_Regression(x_s)+IF(Allow_Intercept,1,0))/Observations(y,fil)"),
        (7,  "AIC",            "=AIC(x_s,y,Allow_Intercept,fil)"),
        (8,  "BIC",            "=BIC(x_s,y,Allow_Intercept,fil)"),
        (9,  "AICc",           "=AICc(x_s,y,Allow_Intercept,fil)"),
        (10, "QQ Correlation", "=QQ_Correlation(x_s,y,Allow_Intercept,fil)"),
    ]:
        _v(sheet, row, _C_F, label)
        _f(sheet, row, _C_G, formula)


def _write_anova(sheet: xw.Sheet) -> None:
    _v(sheet, 10, _C_C, "ANOVA")
    _bold(sheet, 10, _C_C)

    for col, header in zip(
        [_C_C, _C_D, _C_E, _C_F, _C_G, _C_H],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        _v(sheet, 11, col, header)
    _bold_row(sheet, 11, _C_C, _C_H)

    # Regression row (12): MS = E12/D12, F = F12/F13, Sig F = F.DIST.RT(G12, D12, D13)
    _v(sheet, 12, _C_C, "Regression")
    _f(sheet, 12, _C_D, "=DF_Regression(x_s)")
    _f(sheet, 12, _C_E, "=SS_Regression(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 12, _C_F, f"={_a1(12,_C_E)}/{_a1(12,_C_D)}")
    _f(sheet, 12, _C_G, f"={_a1(12,_C_F)}/{_a1(13,_C_F)}")
    _f(sheet, 12, _C_H, f"=F.DIST.RT({_a1(12,_C_G)},{_a1(12,_C_D)},{_a1(13,_C_D)})")

    # Residual row (13): MS = E13/D13
    _v(sheet, 13, _C_C, "Residual")
    _f(sheet, 13, _C_D, "=DF_Residual(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 13, _C_E, "=SS_Residual(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 13, _C_F, f"={_a1(13,_C_E)}/{_a1(13,_C_D)}")

    # Total row (14)
    _v(sheet, 14, _C_C, "Total")
    _f(sheet, 14, _C_D, "=DF_Total(y,Allow_Intercept,fil)")
    _f(sheet, 14, _C_E, "=SS_Total(y,Allow_Intercept,fil)")


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    _v(sheet, 3, _C_I, "PREDICTION INTERVAL")
    _bold(sheet, 3, _C_I)
    for row, label, idx in [
        (4, "Point Estimate",   1),
        (5, "SE Prediction",    2),
        (6, "t Critical",       3),
        (7, "Lower 95%",        4),
        (8, "Upper 95%",        5),
        (9, "Confidence Level", 6),
    ]:
        _v(sheet, row, _C_I, label)
        _f(sheet, row, _C_J,
           f"=INDEX(Prediction_Interval(x_s,y,pred_input,Allow_Intercept,fil,0.05),{idx})")


def _write_coefficients(sheet: xw.Sheet) -> None:
    _v(sheet, 18, _C_C, "COEFFICIENTS")
    _bold(sheet, 18, _C_C)

    for col, header in zip(
        [_C_C, _C_D, _C_E, _C_F, _C_G, _C_H, _C_I],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"],
    ):
        _v(sheet, 19, col, header)
    _bold_row(sheet, 19, _C_C, _C_I)

    # Row labels: intercept then one row per predictor
    _v(sheet, 20, _C_C, "Intercept")
    for i, name in enumerate(PREDICTOR_NAMES):
        _v(sheet, 21 + i, _C_C, name)

    # Spill anchors — when intercepts are disabled, pad with a blank top row
    # so predictor values remain aligned with predictor labels.
    _f(sheet, 20, _C_D,
       '=IF(Allow_Intercept,Coefficients(x_s,y,Allow_Intercept,fil),VSTACK("",Coefficients(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 20, _C_E,
       '=IF(Allow_Intercept,SE_Coefficients(x_s,y,Allow_Intercept,fil),VSTACK("",SE_Coefficients(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 20, _C_F,
       '=IF(Allow_Intercept,T_Stats(x_s,y,Allow_Intercept,fil),VSTACK("",T_Stats(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 20, _C_G,
       '=IF(Allow_Intercept,P_Values(x_s,y,Allow_Intercept,fil),VSTACK("",P_Values(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 20, _C_H,
       '=IF(Allow_Intercept,CI_Lower(x_s,y,Allow_Intercept,fil),VSTACK("",CI_Lower(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 20, _C_I,
       '=IF(Allow_Intercept,CI_Upper(x_s,y,Allow_Intercept,fil),VSTACK("",CI_Upper(x_s,y,Allow_Intercept,fil)))')


def _write_residuals(sheet: xw.Sheet) -> None:
    _v(sheet, 1, _C_L, "RESIDUAL OUTPUT")
    _bold(sheet, 1, _C_L)

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q, _C_R, _C_S, _C_T, _C_U, _C_V, _C_W, _C_X],
        [
            "Observation", "Predicted Y", "Residuals", "LOOCV Prediction",
            "Rank Fraction", "Y Ranked", "Normal Scores",
            "Scaled Residuals", "Scaled Residuals Ranked",
            "Hat Diagonal", "Studentized Residuals", "Studentized Residuals Ranked",
            "Cook's Distance",
        ],
    ):
        _v(sheet, 2, col, header)
    _bold_row(sheet, 2, _C_L, _C_X)

    # Spill anchors — each spills n rows downward
    _f(sheet, 3, _C_L, "=SEQUENCE(Observations(y,fil))")
    _f(sheet, 3, _C_M, "=Predictions(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_N, "=Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_O, "=LOOCV_prediction(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_P, "=Rank_Fraction(y,fil)")
    _f(sheet, 3, _C_Q, "=Y_Ranked(y,fil)")
    _f(sheet, 3, _C_R, "=Normal_Scores(y,fil)")
    _f(sheet, 3, _C_S, "=Scaled_Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_T, "=Scaled_Residuals_Ranked(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_U, "=Hat_diagonal(x_s,Allow_Intercept,fil)")
    _f(sheet, 3, _C_V, "=Studentized_Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_W, "=Studentized_Residuals_Ranked(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_X, "=Cooks_Distance(x_s,y,Allow_Intercept,fil)")


# ── Diagnostic chart helpers ──────────────────────────────────────────────────

def _write_helper_columns(sheet: xw.Sheet) -> None:
    """Write filtered-Y (col Y) and top-predictor (col Z) helper columns."""
    # Col Y: actual Y values filtered to the same rows as the residual table
    _v(sheet, 2, _C_Y, "Actual Y")
    _f(sheet, 3, _C_Y, "=FILTER(y,fil)")

    # Col Z: values for the predictor with the largest |t-stat|
    # Row 2 header: dynamic predictor name (matches the coefficient table in col C)
    _f(sheet, 2, _C_Z,
       "=LET(t,DROP(T_Stats(x_s,y,Allow_Intercept,fil),IF(Allow_Intercept,1,0)),"
       "INDEX(C21:C38,MATCH(MAX(ABS(t)),ABS(t),0)))")
    # Row 3 spill: filtered values for that predictor column from the design matrix
    _f(sheet, 3, _C_Z,
       "=LET(t,DROP(T_Stats(x_s,y,Allow_Intercept,fil),IF(Allow_Intercept,1,0)),"
       "j,MATCH(MAX(ABS(t)),ABS(t),0),"
       "CHOOSECOLS(Design_Matrix(x_s,FALSE,fil),j))")


def _write_reference_line_helpers(sheet: xw.Sheet) -> None:
    """Write 2-row reference line endpoints for each chart into cols AA–AB.

    Each chart uses a consecutive pair of rows:
      rows 1–2   Actual vs Predicted   y = x line
      rows 3–4   Residuals vs Fitted   y = 0 line
      rows 5–6   Normal Q-Q            identity line
      rows 7–8   Predicted vs LOOCV    y = x line
      rows 9–10  Cook's Distance       4/n threshold line
      rows 11–12 Residuals vs Top Pred y = 0 line
    """
    # Chart 1 — Actual vs Predicted: y=x reference (AA/AB rows 1–2)
    _f(sheet, 1,  _C_AA, "=MIN(M3#,Y3#)")
    _f(sheet, 2,  _C_AA, "=MAX(M3#,Y3#)")
    _f(sheet, 1,  _C_AB, f"={_a1(1,  _C_AA)}")
    _f(sheet, 2,  _C_AB, f"={_a1(2,  _C_AA)}")

    # Chart 2 — Residuals vs Fitted: y=0 reference (rows 3–4)
    _f(sheet, 3,  _C_AA, "=MIN(M3#)")
    _f(sheet, 4,  _C_AA, "=MAX(M3#)")
    _v(sheet, 3,  _C_AB, 0)
    _v(sheet, 4,  _C_AB, 0)

    # Chart 3 — Normal Q-Q: identity reference (rows 5–6)
    _f(sheet, 5,  _C_AA, "=MIN(R3#)")
    _f(sheet, 6,  _C_AA, "=MAX(R3#)")
    _f(sheet, 5,  _C_AB, f"={_a1(5,  _C_AA)}")
    _f(sheet, 6,  _C_AB, f"={_a1(6,  _C_AA)}")

    # Chart 4 — Predicted vs LOOCV: y=x reference (rows 7–8)
    _f(sheet, 7,  _C_AA, "=MIN(MIN(M3#),MIN(O3#))")
    _f(sheet, 8,  _C_AA, "=MAX(MAX(M3#),MAX(O3#))")
    _f(sheet, 7,  _C_AB, f"={_a1(7,  _C_AA)}")
    _f(sheet, 8,  _C_AB, f"={_a1(8,  _C_AA)}")

    # Chart 5 — Cook's Distance: 4/n threshold (rows 9–10)
    _f(sheet, 9,  _C_AA, "=MIN(L3#)")
    _f(sheet, 10, _C_AA, "=MAX(L3#)")
    _f(sheet, 9,  _C_AB, "=4/Observations(y,fil)")
    _f(sheet, 10, _C_AB, "=4/Observations(y,fil)")

    # Chart 6 — Residuals vs Top Predictor: y=0 reference (rows 11–12)
    _f(sheet, 11, _C_AA, "=MIN(Z3#)")
    _f(sheet, 12, _C_AA, "=MAX(Z3#)")
    _v(sheet, 11, _C_AB, 0)
    _v(sheet, 12, _C_AB, 0)


def _add_chart(
    sheet: xw.Sheet,
    left: float,
    top: float,
    title: str,
    x_label: str,
    y_label: str,
    data_x_col: int,
    data_y_col: int,
    ref_row: int,
) -> None:
    """Add one XY scatter chart with a main series and a dashed reference line.

    Parameters
    ----------
    sheet : xw.Sheet
        Worksheet to embed the chart in.
    left, top : float
        Chart position in points from the sheet origin.
    title : str
        Chart title text.
    x_label, y_label : str
        Horizontal and vertical axis title text.
    data_x_col, data_y_col : int
        1-based column indices for the main scatter series X and Y data.
    ref_row : int
        First of the two helper rows in cols AA/AB that bound the reference line.
    """
    sname = sheet.name
    xl_c = _col_letter(data_x_col)
    yl_c = _col_letter(data_y_col)
    aa_c = _col_letter(_C_AA)
    ab_c = _col_letter(_C_AB)
    ref_row2 = ref_row + 1

    chart = sheet.charts.add(left=left, top=top, width=_CHART_W, height=_CHART_H)
    cobj = chart.api        # ChartObject (container)
    capi = cobj.Chart       # Chart COM object

    # Remove the auto-generated placeholder series Excel inserts on creation
    while capi.SeriesCollection().Count > 0:
        capi.SeriesCollection(1).Delete()

    capi.ChartType = _XL_XY_SCATTER
    capi.HasTitle = True
    capi.ChartTitle.Text = title
    capi.ChartTitle.Font.Size = 10
    capi.HasLegend = False

    # Main scatter series: rows 3 to _REF_MAX_ROW; empty trailing cells are ignored
    s1 = capi.SeriesCollection().NewSeries()
    s1.Formula = (
        f'=SERIES("",{sname}!${xl_c}$3:${xl_c}${_REF_MAX_ROW},'
        f'{sname}!${yl_c}$3:${yl_c}${_REF_MAX_ROW},1)'
    )
    s1.MarkerStyle = _XL_MARKER_CIRCLE
    s1.MarkerSize = 3
    s1.MarkerForegroundColor = _XL_COLOR_BLUE
    s1.MarkerBackgroundColor = _XL_COLOR_BLUE

    # Reference line series: 2 anchor points define a dashed line
    s2 = capi.SeriesCollection().NewSeries()
    s2.Formula = (
        f'=SERIES("",{sname}!${aa_c}${ref_row}:${aa_c}${ref_row2},'
        f'{sname}!${ab_c}${ref_row}:${ab_c}${ref_row2},2)'
    )
    s2.ChartType = _XL_XY_SCATTER_LINES  # line connecting the 2 anchor points
    s2.MarkerStyle = _XL_MARKER_NONE
    s2.Border.LineStyle = _XL_DASH
    s2.Border.Color = _XL_COLOR_GRAY
    s2.Border.Weight = 2  # xlThin

    # Axis labels
    x_ax = capi.Axes(1)
    x_ax.HasTitle = True
    x_ax.AxisTitle.Text = x_label
    x_ax.AxisTitle.Font.Size = 9

    y_ax = capi.Axes(2)
    y_ax.HasTitle = True
    y_ax.AxisTitle.Text = y_label
    y_ax.AxisTitle.Font.Size = 9


def _write_charts(sheet: xw.Sheet) -> None:
    """Add 6 diagnostic charts in a 2-column grid starting at cell AC2."""
    anchor = sheet.range("AC2")
    x0 = anchor.left
    y0 = anchor.top
    col_step = _CHART_W + _CHART_GAP
    row_step = _CHART_H + _CHART_GAP

    # (title, x_label, y_label, x_col, y_col, ref_row_in_AA_AB)
    specs = [
        ("Actual vs Predicted",        "Predicted Y",   "Actual Y",         _C_M, _C_Y,  1),
        ("Residuals vs Fitted",        "Fitted Values", "Residuals",         _C_M, _C_N,  3),
        ("Normal Q-Q",                 "Theoretical",   "Sample Quantiles",  _C_R, _C_T,  5),
        ("Predicted vs LOOCV",         "Predicted Y",   "LOOCV Predicted",   _C_M, _C_O,  7),
        ("Cook's Distance",            "Observation",   "Cook's Distance",   _C_L, _C_X,  9),
        ("Residuals vs Top Predictor", "Top Predictor", "Residuals",         _C_Z, _C_N, 11),
    ]

    for i, (title, x_label, y_label, x_col, y_col, ref_row) in enumerate(specs):
        _add_chart(
            sheet,
            left=x0 + (i % 2) * col_step,
            top=y0 + (i // 2) * row_step,
            title=title,
            x_label=x_label,
            y_label=y_label,
            data_x_col=x_col,
            data_y_col=y_col,
            ref_row=ref_row,
        )


# ── Public entry point ────────────────────────────────────────────────────────

def write_regression_output_sheet(workbook: xw.Book) -> None:
    """Create or refresh the ToolPak-style Regression sheet in workbook."""

    sheet = next(
        (s for s in workbook.sheets if s.name == REGRESSION_SHEET_NAME), None
    )
    if sheet is None:
        sheet = workbook.sheets.add(
            name=REGRESSION_SHEET_NAME, after=workbook.sheets[-1]
        )

    for idx in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(idx).Delete()
    for idx in range(sheet.api.ChartObjects().Count, 0, -1):
        sheet.api.ChartObjects(idx).Delete()
    sheet.api.Cells.Clear()
    sheet.activate()

    k = len(PREDICTOR_NAMES)
    _setup_local_names(sheet, k)

    # Allow_Intercept toggle: label in C2, value cell in D2
    _v(sheet, 2, _C_C, "Allow Intercept")
    _v(sheet, 2, _C_D, True)

    _write_prediction_inputs(sheet)
    _write_regression_statistics(sheet)
    _write_diagnostics(sheet)
    _write_anova(sheet)
    _write_prediction_interval(sheet)
    _write_coefficients(sheet)
    _write_residuals(sheet)
    _write_helper_columns(sheet)
    _write_reference_line_helpers(sheet)
    _write_charts(sheet)
    group_and_hide_columns(sheet, _C_Y, _C_AB)

    # Column widths (A–X only; Y–AB are hidden)
    for col_letter, width in {
        "A": 30, "B": 12, "C": 24, "D": 14, "E": 14,
        "F": 16, "G": 14, "H": 16, "I": 22, "J": 14,
        "K": 3,  "L": 14, "M": 14, "N": 12, "O": 18,
        "P": 12, "Q": 12, "R": 14, "S": 16, "T": 22,
        "U": 14, "V": 22, "W": 26, "X": 14,
    }.items():
        sheet.range(f"{col_letter}:{col_letter}").column_width = width

    # Freeze top 2 rows
    sheet.api.Application.ActiveWindow.SplitRow = 2
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

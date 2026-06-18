"""
write_sheet_regression.py
Writes the ToolPak-style Regression sheet into any target workbook.

Layout (five horizontal zones):
  Col A–B        — Model Inputs: independent variable labels + "In linear model?" toggles
                   B2 = Allow_Intercept toggle; B3:B16000 = per-independent variable on/off (orange)
  Col C          — thin gap (width 2)
  Col D–J        — Independent Variable Summary: names + Pearson R, Spearman R, Skewness, Kurtosis,
                   VIF, Tolerance (always using all independent variables in All_Xs)
  Col K          — thin gap (width 2)
  Col L–S        — Regression Outputs: Statistics (L–M rows 3–8), Diagnostics (O–P rows 3–10),
                   Sheet-scoped names: All_Xs, Ind_Var_Filter ($B$3:$B$16000), x_s (filtered),
                   Coefficient_Name_Col([Filter] optional), Allow_Intercept, alpha, pred_input
                   Alpha input (M12), ANOVA Table (rows 13–17),
                   Coefficients (rows 19+, cols L–R), Beta Weights (col S, rows 19+)
  Col T          — thin gap (width 2)
  Col U–V        — Prediction Outputs: Prediction Interval (U1:V8, boxed),
                   Prediction Inputs (U10+, no box — dynamic height)
  Col W          — thin gap (width 2)
  Col X–AH       — Residual Output: heading in X; 11 diagnostics columns (X–AH), spills downward from row 3
"""
from __future__ import annotations

from typing import Any

import xlwings as xw

# ── Conditional-formatting helpers ────────────────────────────────────────────

_XL_EXPRESSION = 2
_MAX_EXCEL_ROW = 1_048_576

# Excel built-in conditional-formatting colors
_CF_LIGHT_RED_FILL = (255, 199, 206)       # #FFC7CE
_CF_DARK_RED_TEXT = (156, 0, 6)            # #9C0006
_CF_LIGHT_YELLOW_FILL = (255, 235, 156)    # #FFEB9C
_CF_DARK_YELLOW_TEXT = (156, 101, 0)       # #9C6500

REGRESSION_SHEET_NAME = "Regression"

# ── 1-based column indices ─────────────────────────────────────────────────────

# Zone 1: Model Inputs
_C_A = 1    # Independent variable labels
_C_B = 2    # "In linear model?" toggles (orange user input); B2 = Allow_Intercept

# Zone 2: Independent Variable Summary
_C_C = 3    # thin gap
_C_D = 4    # Independent variable names
_C_E = 5    # Pearson R
_C_F = 6    # Spearman R
_C_G = 7    # Skewness
_C_H = 8    # Kurtosis
_C_I = 9    # VIF
_C_J = 10   # Tolerance

# Zone 3: Regression Outputs
_C_K = 11   # thin gap
_C_L = 12   # labels (stats / ANOVA / coefficients)
_C_M = 13   # stat values / ANOVA df / coefficient values
_C_N = 14   # ANOVA SS / coefficient SE
_C_O = 15   # diagnostics labels / ANOVA MS / coefficient t-stat
_C_P = 16   # diagnostics values / ANOVA F / coefficient p-value
_C_Q = 17   # ANOVA Sig F / coefficient CI lower
_C_R = 18   # coefficient CI upper
_C_S = 19   # Beta Weights

# Zone 4: Prediction Outputs
_C_U = 21   # prediction interval labels / prediction input labels
_C_V = 22   # prediction interval values / prediction input values

# Zone 5: Residual Output
_C_W = 23   # thin gap
_C_X = 24   # section heading anchor / Y
_C_Y = 25   # Predicted Y
_C_Z = 26   # Residuals
_C_AA = 27  # LOOCV residual
_C_AB = 28  # Hat Diagonal
_C_AC = 29  # Studentized Residuals
_C_AD = 30  # Cook's Distance
_C_AE = 31  # Normal Scores Ranked
_C_AF = 32  # Studentized Residuals Ranked
_C_AG = 33  # Scale-Location
_C_AH = 34  # PRESS Residual

# ── Chart constants ───────────────────────────────────────────────────────────
_XL_XY_SCATTER = -4169       # Excel xlXYScatter
_XL_COLUMN_CLUSTERED = 51    # Excel xlColumnClustered
_XL_CATEGORY = 1             # horizontal axis
_XL_VALUE = 2                # vertical axis
_CHART_WIDTH = 310.0         # points
_CHART_HEIGHT = 310.0        # points
_CHART_GAP = 10.0            # gap between charts in points

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


def _val(sheet: xw.Sheet, row: int, col: int, value: object) -> None:
    sheet.range(_rc(row, col)).value = value

def _named_range_column_count(sheet: xw.Sheet, name: str) -> int:
    """Return the number of columns in a sheet-scoped named range."""
    try:
        named_range = sheet.names[name].refers_to_range
        return named_range.columns.count
    except Exception as exc:
        raise RuntimeError(
            f"Could not determine the column count of "
            f"sheet-scoped name {name!r} on sheet {sheet.name!r}."
        ) from exc
    
def _f(sheet: xw.Sheet, row: int, col: int, formula: str) -> None:
    sheet.range(_rc(row, col)).api.Formula2 = formula


def _bold(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(_rc(row, col)).api.Font.Bold = True


def _bold_row(sheet: xw.Sheet, row: int, col1: int, col2: int) -> None:
    sheet.range(_rc(row, col1), _rc(row, col2)).api.Font.Bold = True


# ── Visual formatting helpers ─────────────────────────────────────────────────

_HEADER = (202, 237, 251)   # section headings
_INPUT = (251, 226, 213)   # user-editable input cells



def _section_heading(sheet: xw.Sheet, row: int, col: int, label: str) -> None:
    _val(sheet, row, col, label)
    sheet.range(_rc(row, col)).api.Font.Bold = True
    sheet.range(_rc(row, col)).color = _HEADER


def _format_input(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(_rc(row, col)).color = _INPUT


def _input_range(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    sheet.range(_rc(r1, c1), _rc(r2, c2)).color = _INPUT


def _border_box(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    rng = sheet.range(_rc(r1, c1), _rc(r2, c2)).api
    for edge in [7, 8, 9, 10]:   # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        rng.Borders(edge).LineStyle = 1   # xlContinuous
        rng.Borders(edge).Weight = 2      # xlThin


def _set_note(sheet: xw.Sheet, row: int, col: int, text: str) -> None:
    """Replace the cell's note/comment text with a plain-language explanation."""
    cell_api = sheet.range(_rc(row, col)).api
    try:
        cell_api.ClearComments()
    except Exception:
        pass
    cell_api.AddComment(text)
    cell_api.Comment.Visible = False


def _annotate_statistical_terms(sheet: xw.Sheet, sheet_notes: dict[str, str]) -> None:
    """Attach plain-language notes to key statistical labels on the sheet."""
    note_cells = [
        (2, _C_E, "Pearson R"),
        (2, _C_F, "Spearman R"),
        (2, _C_G, "Skewness"),
        (2, _C_H, "Kurtosis"),
        (2, _C_I, "VIF"),
        (2, _C_J, "Tolerance"),
        (4, _C_L, "Multiple R"),
        (5, _C_L, "R Square"),
        (6, _C_L, "Adjusted R Square"),
        (7, _C_L, "Standard Error"),
        (8, _C_L, "Observations"),
        (4, _C_O, "PRESS"),
        (5, _C_O, "PRESS R²"),
        (6, _C_O, "Mean Leverage"),
        (7, _C_O, "AIC"),
        (8, _C_O, "BIC"),
        (9, _C_O, "AICc"),
        (10, _C_O, "QQ Correlation"),
        (12, _C_L, "Alpha"),
        (14, _C_M, "df"),
        (14, _C_N, "SS"),
        (14, _C_O, "MS"),
        (14, _C_P, "F"),
        (14, _C_Q, "Significance F"),
        (15, _C_L, "Regression"),
        (16, _C_L, "Residual"),
        (17, _C_L, "Total"),
        (20, _C_M, "Coefficients"),
        (20, _C_N, "Std Error"),
        (20, _C_O, "t Stat"),
        (20, _C_P, "P-value"),
        (20, _C_Q, "Lower 95%"),
        (20, _C_R, "Upper 95%"),
        (20, _C_S, "Beta Weight"),
        (3, _C_U, "Point Estimate"),
        (4, _C_U, "SE Prediction"),
        (5, _C_U, "t Critical"),
        (6, _C_U, "Lower 95%"),
        (7, _C_U, "Upper 95%"),
        (8, _C_U, "Confidence Level"),
        (2, _C_X, "Y"),
        (2, _C_Y, "Predicted Y"),
        (2, _C_Z, "Residuals"),
        (2, _C_AA, "LOOCV Residual"),
        (2, _C_AB, "Hat Diagonal"),
        (2, _C_AC, "Studentized Residuals"),
        (2, _C_AD, "Cook's Distance"),
        (2, _C_AE, "Normal Scores Ranked"),
        (2, _C_AF, "Studentized Residuals Ranked"),
        (2, _C_AG, "Scale-Location"),
        (2, _C_AH, "PRESS Residual"),
    ]

    for row, col, key in note_cells:
        note_text = sheet_notes.get(key)
        if note_text is not None:
            _set_note(sheet, row, col, note_text)

# -- Conditional-formatting helpers ----------------------------------------------
def _excel_color(rgb: tuple[int, int, int]) -> int:
    """Convert an RGB tuple to the OLE color integer expected by Excel COM."""
    red, green, blue = rgb
    return red + green * 256 + blue * 65536

def _add_expression_format(
    sheet: xw.Sheet,
    address: str,
    formula: str,
    *,
    fill: tuple[int, int, int] | None = None,
    font_color: tuple[int, int, int] | None = None,
    bold: bool | None = None,
    strikethrough: bool | None = None,
    stop_if_true: bool = False,
):
    """Add a formula-based conditional-formatting rule to a range."""
    condition = sheet.range(address).api.FormatConditions.Add(
        Type=_XL_EXPRESSION,
        Formula1=formula,
    )

    if fill is not None:
        condition.Interior.Color = _excel_color(fill)

    if font_color is not None:
        condition.Font.Color = _excel_color(font_color)

    if bold is not None:
        condition.Font.Bold = bold

    if strikethrough is not None:
        condition.Font.Strikethrough = strikethrough

    condition.StopIfTrue = stop_if_true
    return condition

def _write_significance_conditional_formatting(sheet: xw.Sheet) -> None:
    """Flag nonsignificant coefficient and overall-model P-values."""

    coefficient_p_values = f"P21:P{_MAX_EXCEL_ROW}"
    significance_f = "Q15"

    sheet.range(coefficient_p_values).api.FormatConditions.Delete()
    sheet.range(significance_f).api.FormatConditions.Delete()

    # Individual coefficient P-values above alpha.
    _add_expression_format(
        sheet,
        coefficient_p_values,
        "=AND(ISNUMBER(P21),P21>$M$12)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

    # Overall regression P-value above alpha.
    _add_expression_format(
        sheet,
        significance_f,
        "=AND(ISNUMBER(Q15),Q15>$M$12)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

def _write_residual_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply diagnostic cutoffs to the residual-output columns."""

    addresses = {
        "hat":                f"AB3:AB{_MAX_EXCEL_ROW}",
        "studentized":        f"AC3:AC{_MAX_EXCEL_ROW}",
        "cooks":              f"AD3:AD{_MAX_EXCEL_ROW}",
        "studentized_ranked": f"AF3:AF{_MAX_EXCEL_ROW}",
        "scale_location":     f"AG3:AG{_MAX_EXCEL_ROW}",
        "press_residual":     f"AH3:AH{_MAX_EXCEL_ROW}",
    }

    # Remove existing rules so repeated builds do not duplicate them.
    for address in addresses.values():
        sheet.range(address).api.FormatConditions.Delete()

    # ── Hat diagonal ─────────────────────────────────────────────────────────
    # P6 contains mean leverage, p/n.
    # > 2p/n: light-red fill and dark-red text.
    _add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AB3),AB3>2*$P$6)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

    # > 3p/n: additionally bold.
    _add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AB3),AB3>3*$P$6)",
        bold=True,
    )

    # ── Studentized residuals ────────────────────────────────────────────────
    for column, address in [
        ("AC", addresses["studentized"]),
        ("AF", addresses["studentized_ranked"]),
    ]:
        # 2 < |r| < 3: light-yellow fill and dark-yellow text.
        _add_expression_format(
            sheet,
            address,
            (
                f"=AND("
                f"ISNUMBER({column}3),"
                f"ABS({column}3)>2,"
                f"ABS({column}3)<3"
                f")"
            ),
            fill=_CF_LIGHT_YELLOW_FILL,
            font_color=_CF_DARK_YELLOW_TEXT,
        )

        # |r| >= 3: light-red fill and dark-red text.
        _add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({column}3),ABS({column}3)>=3)",
            fill=_CF_LIGHT_RED_FILL,
            font_color=_CF_DARK_RED_TEXT,
        )

    # ── Cook's distance ──────────────────────────────────────────────────────
    # M8 contains the number of observations, n.
    # 4/n < D <= 0.9: light-yellow fill and dark-yellow text.
    _add_expression_format(
        sheet,
        addresses["cooks"],
        "=AND(ISNUMBER(AD3),AD3>4/$M$8,AD3<=0.9)",
        fill=_CF_LIGHT_YELLOW_FILL,
        font_color=_CF_DARK_YELLOW_TEXT,
    )

    # D > 0.9: light-red fill and dark-red text.
    _add_expression_format(
        sheet,
        addresses["cooks"],
        "=AND(ISNUMBER(AD3),AD3>0.9)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

    # ── Scale-Location: SQRT(|Studentized|) ─────────────────────────────────
    # SQRT(2) ≈ 1.414 corresponds to |Studentized| = 2.
    # SQRT(3) ≈ 1.732 corresponds to |Studentized| = 3.
    _add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AG3),AG3>1.414,AG3<=1.732)",
        fill=_CF_LIGHT_YELLOW_FILL,
        font_color=_CF_DARK_YELLOW_TEXT,
    )
    _add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AG3),AG3>1.732)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

    # ── PRESS Residual: e_i / (1 - h_i) ─────────────────────────────────────
    # M7 contains the Standard Error of the regression.
    # |PRESS| > 2*SE: mild concern; > 3*SE: strong concern.
    _add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AH3),ABS(AH3)>2*$M$7,ABS(AH3)<=3*$M$7)",
        fill=_CF_LIGHT_YELLOW_FILL,
        font_color=_CF_DARK_YELLOW_TEXT,
    )
    _add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AH3),ABS(AH3)>3*$M$7)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

def _write_prediction_inputs_strikethrough_cf(sheet: xw.Sheet) -> None:
    address = f"U13:U{_MAX_EXCEL_ROW}"
    sheet.range(address).api.FormatConditions.Delete()
    _add_expression_format(
        sheet,
        address,
        "=NOT(INDEX(TAKE(Ind_Var_Filter,COLUMNS(All_Xs)),ROW()-ROW($U$13)+1))",
        strikethrough=True,
    )

def _write_model_diagnostic_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply rule-of-thumb formatting to VIF, PRESS R², and QQ Correlation."""

    vif_address = f"I3:I{_MAX_EXCEL_ROW}"
    press_r2_address = "P5"
    qq_corr_address = "P10"

    # Prevent duplicate rules when rebuilding the sheet.
    sheet.range(vif_address).api.FormatConditions.Delete()
    sheet.range(press_r2_address).api.FormatConditions.Delete()
    sheet.range(qq_corr_address).api.FormatConditions.Delete()

    # ── VIF ─────────────────────────────────────────────────────────────────
    # 5 < VIF <= 10: possible multicollinearity; review.
    _add_expression_format(
        sheet,
        vif_address,
        "=AND(ISNUMBER(I3),I3>5,I3<=10)",
        fill=_CF_LIGHT_YELLOW_FILL,
        font_color=_CF_DARK_YELLOW_TEXT,
    )

    # VIF > 10: strong multicollinearity warning.
    _add_expression_format(
        sheet,
        vif_address,
        "=AND(ISNUMBER(I3),I3>10)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

    # ── PRESS R² ─────────────────────────────────────────────────────────────
    # Negative PRESS R² means cross-validated predictions perform worse than
    # predicting the outcome mean.
    _add_expression_format(
        sheet,
        press_r2_address,
        "=AND(ISNUMBER(P5),P5<0)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

    # ── QQ Correlation ────────────────────────────────────────────────────────
    # Pearson r of sorted scaled residuals vs. normal quantiles; near 1.0 = normal errors.
    # < 0.98: mild departure (yellow); < 0.95: stronger departure (red).
    _add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(P10),P10<0.98,P10>=0.95)",
        fill=_CF_LIGHT_YELLOW_FILL,
        font_color=_CF_DARK_YELLOW_TEXT,
    )
    _add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(P10),P10<0.95)",
        fill=_CF_LIGHT_RED_FILL,
        font_color=_CF_DARK_RED_TEXT,
    )

# ── Local name management ─────────────────────────────────────────────────────

def _drop_local_name(sheet: xw.Sheet, name: str) -> None:
    for idx in range(sheet.api.Names.Count, 0, -1):
        local = sheet.api.Names(idx).Name.split("!", 1)[-1]
        if local.lower() == name.lower():
            sheet.api.Names(idx).Delete()


def _setup_local_names(sheet: xw.Sheet) -> None:
    """Register sheet-scoped names used by every formula on this sheet."""
    sname = sheet.name

    # All_Xs: full 18-predictor range — predictor summary always uses this
    _drop_local_name(sheet, "All_Xs")
    sheet.api.Names.Add(
        Name="All_Xs",
        RefersTo="=LifeExpectancyData[[Adult Mortality]:[Schooling]]",
    )

    # Coefficient_Name_Col: optional Filter arg.
    #   (All_Xs)              → all n header names as a column vector
    #   (All_Xs, Ind_Var_Filter) → only headers whose toggle is TRUE
    # TAKE(Filter, n) trims the filter range to exactly n rows.
    # IFERROR fallback returns the first header when all toggles are off.
    _drop_local_name(sheet, "Coefficient_Name_Col")
    sheet.api.Names.Add(
        Name="Coefficient_Name_Col",
        RefersTo=(
            "=LAMBDA(All_Xs,[Filter],"
            "LET("
            "n,COLUMNS(All_Xs),"
            "headers,TRANSPOSE(OFFSET(All_Xs,-1,0,1,n)),"
            "IF(ISOMITTED(Filter),"
            "headers,"
            "IFERROR(FILTER(headers,TAKE(Filter,n)),TAKE(headers,1))"
            ")"
            "))"
        ),
    )

    # x_s: dynamic — only the predictors toggled TRUE in col B via Ind_Var_Filter.
    # Falls back to the first column if all toggles are off.
    _drop_local_name(sheet, "x_s")
    sheet.api.Names.Add(
        Name="x_s",
        RefersTo=(
            "=LET("
            "n,COLUMNS(All_Xs),"
            "sel,IFERROR(FILTER(SEQUENCE(n),TAKE(Ind_Var_Filter,n)),1),"
            "CHOOSECOLS(All_Xs,sel)"
            ")"
        ),
    )

    for name, ref in [
        ("y",   "=LifeExpectancyData[Life expectancy]"),
        ("fil", "=LifeExpectancyData[Full_Data]"),
    ]:
        _drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=ref)

    # Allow_Intercept toggle lives in B2
    _drop_local_name(sheet, "Allow_Intercept")
    sheet.api.Names.Add(
        Name="Allow_Intercept",
        RefersTo=f"={sname}!$B$2",
    )

    # Ind_Var_Filter: boolean range covering all predictor toggle cells.
    # Used by Coefficient_Name_Col and x_s to filter to selected predictors.
    _drop_local_name(sheet, "Ind_Var_Filter")
    sheet.api.Names.Add(
        Name="Ind_Var_Filter",
        RefersTo=f"={sname}!$B$3:$B$16000",
    )

    # alpha: confidence level input, lives in M12
    _drop_local_name(sheet, "alpha")
    sheet.api.Names.Add(
        Name="alpha",
        RefersTo=f"={sname}!$M$12",
    )

    # pred_input: intercept + predictor values for point prediction
    # prediction values column letter is derived from _C_V to stay in sync with the layout.
    _pred_val_letter = _col_letter(_C_V)
    _drop_local_name(sheet, "pred_input")
    sheet.api.Names.Add(
        Name="pred_input",
        RefersTo=(
            f"={sname}!${_pred_val_letter}$12:"
            f"OFFSET({sname}!${_pred_val_letter}$12,1,0,COLUMNS(x_s))"
        ),
    )

    # ── Chart data ranges (OFFSET-based, sized to n = $M$8 observations) ────────
    # These worksheet-scoped names feed chart SERIES formulas as
    # ='Regression'!<Name>, avoiding full-column references that degrade
    # performance and avoiding the unsupported # spill operator in chart formulas.
    for _name, _col_ltr in [
        ("QQPlotX", _col_letter(_C_AE)),        # Normal Scores Ranked
        ("QQPlotY", _col_letter(_C_AF)),        # Studentized Residuals Ranked
        ("FittedY", _col_letter(_C_Y)),         # Predicted Y (shared)
        ("ResidData", _col_letter(_C_Z)),       # Residuals
        ("ActualY", _col_letter(_C_X)),         # Actual Y
        ("ScaleLocData", _col_letter(_C_AG)),   # Scale-Location
        ("CooksDistData", _col_letter(_C_AD)),  # Cook's Distance
        ("LeverageData", _col_letter(_C_AB)),   # Hat Diagonal
        ("StudResidData", _col_letter(_C_AC)),  # Studentized Residuals
        ("PRESSResidData", _col_letter(_C_AH)), # PRESS Residual
    ]:
        _drop_local_name(sheet, _name)
        sheet.api.Names.Add(
            Name=_name,
            RefersTo=f"=OFFSET('{sname}'!${_col_ltr}$2,1,0,'{sname}'!$M$8,1)",
        )


# ── Section writers ───────────────────────────────────────────────────────────

def _write_model_selection(sheet: xw.Sheet, k: int) -> None:
    """Zone A–B: predictor labels + 'In linear model?' toggles."""
    _section_heading(sheet, 1, _C_A, "MODEL SELECTION")
    _val(sheet, 1, _C_B, "In linear model?")
    _bold(sheet, 1, _C_B)

    # Row 2: Allow Intercept toggle (the named Allow_Intercept cell)
    _val(sheet, 2, _C_A, "Allow Intercept")
    _val(sheet, 2, _C_B, True)

    # A3: spill formula — fills predictor names from the table headers
    _f(sheet, 3, _C_A, "=Coefficient_Name_Col(All_Xs)")

    # B3:B2+k contains one model-selection toggle per All_Xs column.
    for i in range(k):
        _val(sheet, 3 + i, _C_B, True)

    # Orange for all user-editable toggle cells
    _input_range(sheet, 2, _C_B, 2 + k, _C_B)


def _write_boolean_validation(sheet: xw.Sheet) -> None:
    """B2:B16000 — in-cell dropdown restricted to TRUE / FALSE."""
    rng = sheet.range(_rc(2, _C_B), _rc(16000, _C_B)).api
    rng.Validation.Delete()
    rng.Validation.Add(
        Type=3,        # xlValidateList
        AlertStyle=1,  # xlValidAlertStop
        Operator=1,    # xlBetween (required positional arg, unused for lists)
        Formula1="TRUE,FALSE",
    )
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True


def _write_predictor_summary(sheet: xw.Sheet, k: int) -> None:
    """Zone D–J: EDA stats for the predictors currently selected into x_s."""
    _section_heading(sheet, 1, _C_D, "PREDICTOR SUMMARY")

    for col, header in zip(
        [_C_D, _C_E, _C_F, _C_G, _C_H, _C_I, _C_J],
        ["", "Pearson R", "Spearman R", "Skewness", "Kurtosis", "VIF", "Tolerance"],
    ):
        _val(sheet, 2, col, header)
    _bold_row(sheet, 2, _C_D, _C_J)

    # Spill anchors at row 3 — each spills once per selected predictor
    _f(sheet, 3, _C_D, "=Coefficient_Name_Col(All_Xs,Ind_Var_Filter)")
    _f(sheet, 3, _C_E, "=Pearson_R(x_s,y,fil)")
    _f(sheet, 3, _C_F, "=Spearman_R(x_s,y,fil)")
    _f(sheet, 3, _C_G, "=Skewness(x_s,fil)")
    _f(sheet, 3, _C_H, "=Kurtosis(x_s,fil)")
    _f(sheet, 3, _C_I, "=VIF(x_s,Allow_Intercept,fil)")
    _f(sheet, 3, _C_J, "=Tolerance(x_s,Allow_Intercept,fil)")

    last = 2 + k
    sheet.range((_rc(3, _C_E)), (_rc(last, _C_J))).number_format = "0.00"


def _write_regression_outputs_header(sheet: xw.Sheet) -> None:
    _section_heading(sheet, 1, _C_L, "REGRESSION OUTPUTS")


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    """Cols L–M, rows 3–8."""
    _section_heading(sheet, 3, _C_L, "REGRESSION STATISTICS")
    for row, label, formula in [
        (4, "Multiple R",        "=Multiple_R(x_s,y,Allow_Intercept,fil)"),
        (5, "R Square",          "=R_squared(x_s,y,Allow_Intercept,fil)"),
        (6, "Adjusted R Square", "=Adjusted_R2(x_s,y,Allow_Intercept,fil)"),
        (7, "Standard Error",    "=SE_Regression(x_s,y,Allow_Intercept,fil)"),
        (8, "Observations",      "=Observations(y,fil)"),
    ]:
        _val(sheet, row, _C_L, label)
        _f(sheet, row, _C_M, formula)
    sheet.range(_rc(4, _C_M), _rc(7, _C_M)).number_format = "0.0000"
    sheet.range(_rc(8, _C_M), _rc(8, _C_M)).number_format = "0"
    _border_box(sheet, 3, _C_L, 8, _C_M)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    """Cols O–P, rows 3–10."""
    _section_heading(sheet, 3, _C_O, "DIAGNOSTICS")
    for row, label, formula in [
        (4,  "PRESS",          "=PRESS(x_s,y,Allow_Intercept,fil)"),
        (5,  "PRESS R²",  "=1-PRESS(x_s,y,Allow_Intercept,fil)/SS_Total(y,Allow_Intercept,fil)"),
        (6,  "Mean Leverage",  "=(DF_Regression(x_s)+IF(Allow_Intercept,1,0))/Observations(y,fil)"),
        (7,  "AIC",            "=AIC(x_s,y,Allow_Intercept,fil)"),
        (8,  "BIC",            "=BIC(x_s,y,Allow_Intercept,fil)"),
        (9,  "AICc",           "=AICc(x_s,y,Allow_Intercept,fil)"),
        (10, "QQ Correlation", "=QQ_Correlation(x_s,y,Allow_Intercept,fil)"),
    ]:
        _val(sheet, row, _C_O, label)
        _f(sheet, row, _C_P, formula)
    sheet.range(_rc(4, _C_P), _rc(10, _C_P)).number_format = "0.0000"
    _border_box(sheet, 3, _C_O, 10, _C_P)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at M12 — controls prediction interval confidence level."""
    _val(sheet, 12, _C_L, "Alpha")
    _bold(sheet, 12, _C_L)
    _val(sheet, 12, _C_M, 0.05)
    _format_input(sheet, 12, _C_M)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 13–17, cols L–Q."""
    _section_heading(sheet, 13, _C_L, "ANOVA TABLE")

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        _val(sheet, 14, col, header)
    _bold_row(sheet, 14, _C_L, _C_Q)

    _val(sheet, 15, _C_L, "Regression")
    _f(sheet, 15, _C_M, "=DF_Regression(x_s)")
    _f(sheet, 15, _C_N, "=SS_Regression(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 15, _C_O, "=MS_Regression(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 15, _C_P, "=F_Stat(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 15, _C_Q, "=P_Value_F(x_s,y,Allow_Intercept,fil)")

    _val(sheet, 16, _C_L, "Residual")
    _f(sheet, 16, _C_M, "=DF_Residual(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 16, _C_N, "=SS_Residual(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 16, _C_O, "=MS_Residual(x_s,y,Allow_Intercept,fil)")

    _val(sheet, 17, _C_L, "Total")
    _f(sheet, 17, _C_M, "=DF_Total(y,Allow_Intercept,fil)")
    _f(sheet, 17, _C_N, "=SS_Total(y,Allow_Intercept,fil)")

    sheet.range(_rc(15, _C_M), _rc(17, _C_M)).number_format = "0"
    sheet.range(_rc(15, _C_N), _rc(17, _C_N)).number_format = "0.0"
    sheet.range(_rc(15, _C_O), _rc(16, _C_O)).number_format = "0.0"
    sheet.range(_rc(15, _C_P), _rc(15, _C_P)).number_format = "0.0"
    sheet.range(_rc(15, _C_Q), _rc(15, _C_Q)).number_format = "0.0E+00"
    _border_box(sheet, 13, _C_L, 17, _C_Q)


def _write_coefficients(sheet: xw.Sheet, k: int) -> None:
    """Cols L–S, rows 19+. Spills downward — nothing placed below row 39 in these cols."""
    _section_heading(sheet, 19, _C_L, "COEFFICIENTS")

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q, _C_R, _C_S],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%", "Beta Weight"],
    ):
        _val(sheet, 20, col, header)
    _bold_row(sheet, 20, _C_L, _C_S)

    # Spill row labels aligned to selected predictors
    _f(
        sheet,
        21,
        _C_L,
        '=IF(Allow_Intercept,'
        'VSTACK("Intercept",Coefficient_Name_Col(All_Xs,Ind_Var_Filter)),'
        'VSTACK("",Coefficient_Name_Col(All_Xs,Ind_Var_Filter)))',
    )

    # Spill anchors at row 21 — pad with blank top row when intercept is disabled
    _f(sheet, 21, _C_M,
       '=IF(Allow_Intercept,Coefficients(x_s,y,Allow_Intercept,fil),VSTACK("",Coefficients(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 21, _C_N,
       '=IF(Allow_Intercept,SE_Coefficients(x_s,y,Allow_Intercept,fil),VSTACK("",SE_Coefficients(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 21, _C_O,
       '=IF(Allow_Intercept,T_Stats(x_s,y,Allow_Intercept,fil),VSTACK("",T_Stats(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 21, _C_P,
       '=IF(Allow_Intercept,P_Values(x_s,y,Allow_Intercept,fil),VSTACK("",P_Values(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 21, _C_Q,
       '=IF(Allow_Intercept,CI_Lower(x_s,y,Allow_Intercept,fil),VSTACK("",CI_Lower(x_s,y,Allow_Intercept,fil)))')
    _f(sheet, 21, _C_R,
       '=IF(Allow_Intercept,CI_Upper(x_s,y,Allow_Intercept,fil),VSTACK("",CI_Upper(x_s,y,Allow_Intercept,fil)))')
    # Beta Weights: k×1 (no intercept row); always prepend blank to align with other columns.
    _f(sheet, 21, _C_S, '=VSTACK("",Beta_Weights(x_s,y,Allow_Intercept,fil))')

    last_coef_row = 21 + k
    for col in [_C_M, _C_N, _C_O, _C_Q, _C_R, _C_S]:
        sheet.range(_rc(21, col), _rc(last_coef_row, col)).number_format = "0.0000"
    sheet.range(_rc(21, _C_P), _rc(last_coef_row, _C_P)).number_format = "0.0E+00"


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    """Zone U1:V8: boxed prediction interval output."""
    _section_heading(sheet, 1, _C_U, "PREDICTION OUTPUTS")
    _val(sheet, 2, _C_U, "PREDICTION INTERVAL")
    _bold(sheet, 2, _C_U)
    for row, label in [
        (3, "Point Estimate"),
        (4, "SE Prediction"),
        (5, "t Critical"),
        (6, "Lower 95%"),
        (7, "Upper 95%"),
        (8, "Confidence Level"),
    ]:
        _val(sheet, row, _C_U, label)
    _f(sheet, 3, _C_V, "=Prediction_Interval(x_s,y,pred_input,Allow_Intercept,fil,alpha)")
    sheet.range(_rc(3, _C_V), _rc(8, _C_V)).number_format = "0.0000"
    _border_box(sheet, 1, _C_U, 8, _C_V)


def _write_prediction_inputs(sheet: xw.Sheet, k: int) -> None:
    """Zone U10:V12+k: per-predictor values used for the point prediction."""
    _section_heading(sheet, 10, _C_U, "PREDICTION INPUTS")
    _val(sheet, 11, _C_U, "Predictor")
    _val(sheet, 11, _C_V, "Prediction Value")
    _bold_row(sheet, 11, _C_U, _C_V)

    # Row 12: intercept (auto-set, still orange to show it's a value)
    _val(sheet, 12, _C_U, "Intercept")
    _f(sheet, 12, _C_V, "=IF(Allow_Intercept,1,0)")
    _format_input(sheet, 12, _C_V)

    # T13: spill formula — fills predictor names from the table headers
    _f(sheet, 13, _C_U, "=Coefficient_Name_Col(All_Xs)")

    # U13:U12+k — mean of each predictor column (filtered), individually overridable
    for i in range(k):
        _f(sheet, 13 + i, _C_V, f"=AVERAGEIF(fil,TRUE,INDEX(All_Xs,,{i + 1}))")

    # Orange for all user-editable prediction value cells
    _input_range(sheet, 13, _C_V, 12 + k, _C_V)
    sheet.range(_rc(13, _C_V), _rc(12 + k, _C_V)).number_format = "0.0000"


def _write_residuals(sheet: xw.Sheet) -> None:
    """Residual diagnostic table — 11 columns starting at _C_X."""
    _section_heading(sheet, 1, _C_X, "RESIDUAL OUTPUT")

    for col, header in zip(
        [_C_X, _C_Y, _C_Z, _C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF, _C_AG, _C_AH],
        [
            "Y", "Predicted Y", "Residuals", "LOOCV Residual",
            "Hat Diagonal", "Studentized Residuals", "Cook's Distance",
            "Normal Scores Ranked", "Studentized Residuals Ranked",
            "Scale-Location", "PRESS Residual",
        ],
    ):
        _val(sheet, 2, col, header)
    _bold_row(sheet, 2, _C_X, _C_AH)

    # Spill anchors — each spills n rows downward
    _f(sheet, 3, _C_X,  "=Dependent_Var(y,fil)")
    _f(sheet, 3, _C_Y,  "=Predictions(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_Z,  "=Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AA,  "=Dependent_Var(y,fil)-LOOCV_prediction(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AB, "=Hat_diagonal(x_s,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AC, "=Studentized_Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AD, "=Cooks_Distance(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AE, "=SORT(Normal_Scores(y,fil))")
    _f(sheet, 3, _C_AF, "=Studentized_Residuals_Ranked(x_s,y,Allow_Intercept,fil)")
    # Scale-Location: SQRT(|Studentized_Residuals|) — horizontal spread should be flat.
    _f(sheet, 3, _C_AG, "=SQRT(ABS(Studentized_Residuals(x_s,y,Allow_Intercept,fil)))")
    # PRESS Residual: e_i / (1 - h_i) — large values flag high-influence observations.
    _f(sheet, 3, _C_AH, "=Residuals(x_s,y,Allow_Intercept,fil)/(1-Hat_diagonal(x_s,Allow_Intercept,fil))")
    sheet.range(f"{_col_letter(_C_Y)}:{_col_letter(_C_AH)}").number_format = "0.0000"


def _write_diagnostic_charts(sheet: xw.Sheet) -> None:
    """Create 7 pre-built diagnostic charts to the right of the Residual Output section."""
    start_left = sheet.range(_a1(1, _C_AH + 1)).left
    start_top = sheet.range("A3").top

    col_step = _CHART_WIDTH + _CHART_GAP
    row_step = _CHART_HEIGHT + _CHART_GAP

    def _pos(grid_row: int, grid_col: int) -> tuple[float, float]:
        return (
            start_left + (grid_col - 1) * col_step,
            start_top + (grid_row - 1) * row_step,
        )

    sname = REGRESSION_SHEET_NAME

    # Chart SERIES formulas require explicit references. The # spill operator is not
    # reliably supported in chart formulas, so use worksheet-scoped names sized to n.
    def _name_ref(local_name: str) -> str:
        return f"='{sname}'!{local_name}"

    chart_specs = [
        (
            "Residuals vs. Fitted", "scatter",
            _name_ref("FittedY"),
            _name_ref("ResidData"),
            "Fitted Values", "Residuals", 1, 1,
        ),
        (
            "Normal Q-Q", "scatter",
            _name_ref("QQPlotX"),
            _name_ref("QQPlotY"),
            "Theoretical Quantiles", "Studentized Residuals", 1, 2,
        ),
        (
            "Actual vs. Predicted", "scatter",
            _name_ref("FittedY"),
            _name_ref("ActualY"),
            "Predicted Y", "Actual Y", 2, 1,
        ),
        (
            "Scale-Location", "scatter",
            _name_ref("FittedY"),
            _name_ref("ScaleLocData"),
            "Fitted Values", "√|Studentized Residual|", 2, 2,
        ),
        (
            "Cook's Distance", "bar",
            None,
            _name_ref("CooksDistData"),
            "Observation", "Cook's Distance", 3, 1,
        ),
        (
            "Studentized Residuals vs. Leverage", "scatter",
            _name_ref("LeverageData"),
            _name_ref("StudResidData"),
            "Leverage (Hat Diagonal)", "Studentized Residuals", 3, 2,
        ),
        (
            "PRESS Residuals", "bar",
            None,
            _name_ref("PRESSResidData"),
            "Observation", "PRESS Residual", 4, 1,
        ),
    ]

    # Per-chart gridline strategy:
    # - Use Y major gridlines for residual magnitude judgment on residual and bar charts.
    # - Use both axes on comparative scatter plots where position relative to both axes matters.
    gridline_modes = {
        "Residuals vs. Fitted": "none",
        "Normal Q-Q": "both",
        "Actual vs. Predicted": "both",
        "Scale-Location": "y",
        "Cook's Distance": "y",
        "Studentized Residuals vs. Leverage": "both",
        "PRESS Residuals": "y",
    }

    def _set_equal_axis_scale_from_named_ranges(
        x_axis: Any,
        y_axis: Any,
        x_name: str,
        y_name: str,
    ) -> None:
        """Set equal min/max scales on both axes using two sheet-scoped named ranges."""
        common_min = float(sheet.api.Evaluate(f"=MIN('{sname}'!{x_name},'{sname}'!{y_name})"))
        common_max = float(sheet.api.Evaluate(f"=MAX('{sname}'!{x_name},'{sname}'!{y_name})"))

        if common_max <= common_min:
            return

        x_axis.MinimumScale = common_min
        x_axis.MaximumScale = common_max
        y_axis.MinimumScale = common_min
        y_axis.MaximumScale = common_max

    def _add_identity_line(chart: Any) -> None:
        """Draw a dotted y=x line inside the chart plot area without adding a data series."""
        plot_area = chart.PlotArea
        line = chart.Shapes.AddLine(
            plot_area.InsideLeft,
            plot_area.InsideTop + plot_area.InsideHeight,
            plot_area.InsideLeft + plot_area.InsideWidth,
            plot_area.InsideTop,
        )
        line.Line.ForeColor.RGB = _excel_color((120, 120, 120))
        line.Line.DashStyle = 3  # msoLineRoundDot
        line.Line.Weight = 1.25

    for title, chart_type, x_addr, y_addr, x_label, y_label, grid_row, grid_col in chart_specs:
        left, top = _pos(grid_row, grid_col)
        co = sheet.api.ChartObjects().Add(left, top, _CHART_WIDTH, _CHART_HEIGHT)
        chart = co.Chart

        chart.ChartType = _XL_XY_SCATTER if chart_type == "scatter" else _XL_COLUMN_CLUSTERED

        sc = chart.SeriesCollection()
        for i in range(sc.Count, 0, -1):
            sc.Item(i).Delete()

        series = chart.SeriesCollection().NewSeries()
        if x_addr is not None:
            series.XValues = x_addr
        series.Values = y_addr
        series.Name = title
        # Bar charts (Cook's Distance, PRESS Residuals) have no markers to resize.
        if chart_type == "scatter":
            series.MarkerSize = 4

        # All charts: Header-style title (bold, 14 pt, light-blue fill).
        chart.HasLegend = False
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        chart.ChartTitle.Font.Bold = True
        chart.ChartTitle.Font.Size = 14
        chart.ChartTitle.Format.Fill.Visible = True
        chart.ChartTitle.Format.Fill.Solid()
        chart.ChartTitle.Format.Fill.ForeColor.RGB = _excel_color(_HEADER)

        x_axis = chart.Axes(_XL_CATEGORY)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = x_label
        x_axis.TickLabels.NumberFormat = "0"

        y_axis = chart.Axes(_XL_VALUE)
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_label
        y_axis.TickLabels.NumberFormat = "0"

        gridline_mode = gridline_modes.get(title, "none")
        x_axis.HasMajorGridlines = gridline_mode == "both"
        y_axis.HasMajorGridlines = gridline_mode in {"y", "both"}

        if title == "Cook's Distance":
            y_axis.TickLabels.NumberFormat = "0.0E+00"
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone
        if title == "Studentized Residuals vs. Leverage":
            x_axis.TickLabels.NumberFormat = "0.00"
        if title == "PRESS Residuals":
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone

        if title == "Normal Q-Q":
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "QQPlotX", "QQPlotY")
            _add_identity_line(chart)
        if title == "Actual vs. Predicted":
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "FittedY", "ActualY")
            _add_identity_line(chart)


# ── Public entry point ────────────────────────────────────────────────────────

def write_regression_output_sheet(
    workbook: xw.Book,
    sheet_notes: dict[str, str] | None = None,
) -> None:
    """Create or refresh the ToolPak-style Regression sheet in workbook.

    Parameters
    ----------
    sheet_notes : dict[str, str] | None
        Mapping of sheet label → plain-language note text from the
        ``regression_sheet_notes`` key in lambda_functions.json.
        Pass ``None`` to skip annotation (useful for isolated tests).
    """

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

    _setup_local_names(sheet)

    # Derive the predictor count directly from the columns in All_Xs.
    k = _named_range_column_count(sheet, "All_Xs")

    _write_model_selection(sheet, k)
    _write_boolean_validation(sheet)
    _write_predictor_summary(sheet, k)
    _write_regression_outputs_header(sheet)
    _write_regression_statistics(sheet)
    _write_diagnostics(sheet)
    _write_model_diagnostic_conditional_formatting(sheet)
    _write_alpha(sheet)
    _write_anova(sheet)
    _write_coefficients(sheet, k)
    _write_significance_conditional_formatting(sheet)
    _write_prediction_interval(sheet)
    _write_prediction_inputs(sheet, k)
    _write_residuals(sheet)
    _annotate_statistical_terms(sheet, sheet_notes or {})
    _write_residual_conditional_formatting(sheet)
    _write_prediction_inputs_strikethrough_cf(sheet)

    sheet.range(_rc(2, _C_A), _rc(2, _C_AH)).api.WrapText = True

    # Column widths (U = prediction labels, V = prediction values; residuals start at X)
    for col_letter, width in {
        "A": 28, "B": 14,
        "C": 2,   # thin gap
        "D": 28, "E": 8, "F": 10, "G": 10, "H": 8, "I": 8, "J": 10,
        "K": 2,   # thin gap
        "L": 22, "M": 12, "N": 12, "O": 14, "P": 10, "Q": 13, "R": 10, "S": 12,
        "T": 2,   # thin gap
        "U": 20, "V": 14,  # prediction labels / values
        "W": 2,   # thin gap
        "X": 10, "Y": 9, "Z": 10, "AA": 9, "AB": 9, "AC": 12, "AD": 9, "AE": 14, "AF": 17,
        "AG": 14, "AH": 15,  # Scale-Location / PRESS Residual
    }.items():
        sheet.range(f"{col_letter}:{col_letter}").column_width = width

    # Charts must be positioned after column widths are set so that
    # sheet.range("AI1").left reflects the final column layout.
    _write_diagnostic_charts(sheet)

    # Freeze top 2 rows
    sheet.activate()
    sheet.range("A3").select()
    win = sheet.api.Application.ActiveWindow
    win.FreezePanes = False
    win.SplitRow = 2
    win.SplitColumn = 0
    win.FreezePanes = True

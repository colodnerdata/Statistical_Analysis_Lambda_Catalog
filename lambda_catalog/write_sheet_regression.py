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
  Col L–S        — Regression Outputs: Predicted Variable label (P2) + dependent-variable
                   header (Q2) at the top, then Statistics (L–M rows 3–8),
                   Diagnostics (O–P rows 3–10),
                   Sheet-scoped names: All_Xs, Ind_Var_Include ($B$3:$B$16000), x_s() (filtered),
                   Coefficient_Name_Col([Include] optional), Allow_Intercept, alpha
                   Alpha input (M12), ANOVA Table (rows 13–17),
                   Coefficients (rows 19+, cols L–R), Beta Weights (col S, rows 19+)
  Col T          — thin gap (width 2)
  Col U–V        — Prediction Outputs: Prediction Interval (U1:V8, boxed),
                   Prediction Inputs (U10+, no box — dynamic height)
  Col W          — thin gap (width 2)
  Col X–AI       — Residual Output: heading + row identifiers (data_identifiers) in X;
                   11 diagnostics columns (Y–AI), spills downward from row 3
"""
from __future__ import annotations

from typing import Any

import xlwings as xw

from .sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_DARK_YELLOW_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
    HEADER_COLOR as _HEADER,
    INPUT_COLOR as _INPUT,
)
from .workbook_helpers import (
    MAX_EXCEL_ROW, a1, add_expression_format, bold, bold_row, border_box,
    col_letter, drop_local_name, excel_color, f, format_input, rc,
    section_heading, val,
)

# ── Conditional-formatting helpers ────────────────────────────────────────────

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
_C_P = 16   # Predicted Variable label (P2) / diagnostics values / ANOVA F / coefficient p-value
_C_Q = 17   # predicted variable header (Q2) / ANOVA Sig F / coefficient CI lower
_C_R = 18   # coefficient CI upper
_C_S = 19   # Beta Weights

# Zone 4: Prediction Outputs
_C_U = 21   # prediction interval labels / prediction input labels
_C_V = 22   # prediction interval values / prediction input values

# Zone 5: Residual Output
_C_W = 23   # thin gap
_C_X = 24   # section heading anchor / data identifiers (row labels)
_C_Y = 25   # Y (actual dependent variable)
_C_Z = 26   # Predicted Y
_C_AA = 27  # Residuals
_C_AB = 28  # LOOCV residual
_C_AC = 29  # Hat Diagonal
_C_AD = 30  # Studentized Residuals
_C_AE = 31  # Cook's Distance
_C_AF = 32  # Normal Scores Ranked
_C_AG = 33  # Studentized Residuals Ranked
_C_AH = 34  # Scale-Location
_C_AI = 35  # PRESS Residual

# ── Chart constants ───────────────────────────────────────────────────────────
_XL_XY_SCATTER = -4169       # Excel xlXYScatter
_XL_XY_SCATTER_LINES_NO_MARKERS = 75  # Excel xlXYScatterLinesNoMarkers
_XL_COLUMN_CLUSTERED = 51    # Excel xlColumnClustered
_XL_CATEGORY = 1             # horizontal axis
_XL_VALUE = 2                # vertical axis
_CHART_WIDTH = 310.0         # points
_CHART_HEIGHT = 310.0        # points
_CHART_GAP = 10.0            # gap between charts in points

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
    
# ── Visual formatting helpers ─────────────────────────────────────────────────


def _input_range(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    sheet.range(rc(r1, c1), rc(r2, c2)).color = _INPUT


def _set_note(sheet: xw.Sheet, row: int, col: int, text: str) -> None:
    """Replace the cell's note/comment text with a plain-language explanation."""
    cell_api = sheet.range(rc(row, col)).api
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
        (2, _C_Y, "Y"),
        (2, _C_Z, "Predicted Y"),
        (2, _C_AA, "Residuals"),
        (2, _C_AB, "LOOCV Residual"),
        (2, _C_AC, "Hat Diagonal"),
        (2, _C_AD, "Studentized Residuals"),
        (2, _C_AE, "Cook's Distance"),
        (2, _C_AF, "Normal Scores Ranked"),
        (2, _C_AG, "Studentized Residuals Ranked"),
        (2, _C_AH, "Scale-Location"),
        (2, _C_AI, "PRESS Residual"),
    ]

    for row, col, key in note_cells:
        note_text = sheet_notes.get(key)
        if note_text is not None:
            _set_note(sheet, row, col, note_text)

def _write_significance_conditional_formatting(sheet: xw.Sheet) -> None:
    """Flag nonsignificant coefficient and overall-model P-values."""

    coefficient_p_values = f"P21:P{MAX_EXCEL_ROW}"
    significance_f = "Q15"

    sheet.range(coefficient_p_values).api.FormatConditions.Delete()
    sheet.range(significance_f).api.FormatConditions.Delete()

    # Individual coefficient P-values above alpha.
    add_expression_format(
        sheet,
        coefficient_p_values,
        "=AND(ISNUMBER(P21),P21>$M$12)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Overall regression P-value above alpha.
    add_expression_format(
        sheet,
        significance_f,
        "=AND(ISNUMBER(Q15),Q15>$M$12)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

def _write_residual_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply diagnostic cutoffs to the residual-output columns."""

    addresses = {
        "hat":                f"AC3:AC{MAX_EXCEL_ROW}",
        "studentized":        f"AD3:AD{MAX_EXCEL_ROW}",
        "cooks":              f"AE3:AE{MAX_EXCEL_ROW}",
        "studentized_ranked": f"AG3:AG{MAX_EXCEL_ROW}",
        "scale_location":     f"AH3:AH{MAX_EXCEL_ROW}",
        "press_residual":     f"AI3:AI{MAX_EXCEL_ROW}",
    }

    # Remove existing rules so repeated builds do not duplicate them.
    for address in addresses.values():
        sheet.range(address).api.FormatConditions.Delete()

    # ── Hat diagonal ─────────────────────────────────────────────────────────
    # P6 contains mean leverage, p/n.
    # > 2p/n: light-red fill and dark-red text.
    add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AC3),AC3>2*$P$6)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # > 3p/n: additionally bold.
    add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AC3),AC3>3*$P$6)",
        bold_font=True,
    )

    # ── Studentized residuals ────────────────────────────────────────────────
    for column, address in [
        ("AD", addresses["studentized"]),
        ("AG", addresses["studentized_ranked"]),
    ]:
        # 2 < |r| < 3: light-yellow fill and dark-yellow text.
        add_expression_format(
            sheet,
            address,
            (
                f"=AND("
                f"ISNUMBER({column}3),"
                f"ABS({column}3)>2,"
                f"ABS({column}3)<3"
                f")"
            ),
            fill=CF_YELLOW_FILL,
            font_color=CF_DARK_YELLOW_TEXT,
        )

        # |r| >= 3: light-red fill and dark-red text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({column}3),ABS({column}3)>=3)",
            fill=CF_LIGHT_RED_FILL,
            font_color=CF_DARK_RED_TEXT,
        )

    # ── Cook's distance ──────────────────────────────────────────────────────
    # M8 contains the number of observations, n.
    # 4/n < D <= 0.9: light-yellow fill and dark-yellow text.
    add_expression_format(
        sheet,
        addresses["cooks"],
        "=AND(ISNUMBER(AE3),AE3>4/$M$8,AE3<=0.9)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # D > 0.9: light-red fill and dark-red text.
    add_expression_format(
        sheet,
        addresses["cooks"],
        "=AND(ISNUMBER(AE3),AE3>0.9)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── Scale-Location: SQRT(|Studentized|) ─────────────────────────────────
    # SQRT(2) ≈ 1.414 corresponds to |Studentized| = 2.
    # SQRT(3) ≈ 1.732 corresponds to |Studentized| = 3.
    add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AH3),AH3>1.414,AH3<=1.732)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AH3),AH3>1.732)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS Residual: e_i / (1 - h_i) ─────────────────────────────────────
    # M7 contains the Standard Error of the regression.
    # |PRESS| > 2*SE: mild concern; > 3*SE: strong concern.
    add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AI3),ABS(AI3)>2*$M$7,ABS(AI3)<=3*$M$7)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AI3),ABS(AI3)>3*$M$7)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

def _write_prediction_inputs_strikethrough_cf(sheet: xw.Sheet) -> None:
    address = f"U13:U{MAX_EXCEL_ROW}"
    sheet.range(address).api.FormatConditions.Delete()
    add_expression_format(
        sheet,
        address,
        "=NOT(INDEX(TAKE(Ind_Var_Include,COLUMNS(All_Xs)),ROW()-ROW($U$13)+1))",
        strikethrough=True,
    )

def _write_model_diagnostic_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply rule-of-thumb formatting to VIF, PRESS R², and QQ Correlation."""

    vif_address = f"I3:I{MAX_EXCEL_ROW}"
    press_r2_address = "P5"
    qq_corr_address = "P10"

    # Prevent duplicate rules when rebuilding the sheet.
    sheet.range(vif_address).api.FormatConditions.Delete()
    sheet.range(press_r2_address).api.FormatConditions.Delete()
    sheet.range(qq_corr_address).api.FormatConditions.Delete()

    # ── VIF ─────────────────────────────────────────────────────────────────
    # 5 < VIF <= 10: possible multicollinearity; review.
    add_expression_format(
        sheet,
        vif_address,
        "=AND(ISNUMBER(I3),I3>5,I3<=10)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # VIF > 10: strong multicollinearity warning.
    add_expression_format(
        sheet,
        vif_address,
        "=AND(ISNUMBER(I3),I3>10)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS R² ─────────────────────────────────────────────────────────────
    # Negative PRESS R² means cross-validated predictions perform worse than
    # predicting the outcome mean.
    add_expression_format(
        sheet,
        press_r2_address,
        "=AND(ISNUMBER(P5),P5<0)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── QQ Correlation ────────────────────────────────────────────────────────
    # Pearson r of sorted scaled residuals vs. normal quantiles; near 1.0 = normal errors.
    # < 0.98: mild departure (yellow); < 0.95: stronger departure (red).
    add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(P10),P10<0.98,P10>=0.95)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(P10),P10<0.95)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

# ── Local name management ─────────────────────────────────────────────────────

def _setup_local_names(sheet: xw.Sheet) -> None:
    """Register sheet-scoped names used by every formula on this sheet."""
    sname = sheet.name

    # All_Xs: full 18-predictor range — predictor summary always uses this
    drop_local_name(sheet, "All_Xs")
    sheet.api.Names.Add(
        Name="All_Xs",
        RefersTo="=LifeExpectancyData[[Adult Mortality]:[Schooling]]",
    )

    # Coefficient_Name_Col: optional Include arg.
    #   (All_Xs)              → all n header names as a column vector
    #   (All_Xs, Ind_Var_Include) → only headers whose toggle is TRUE
    # TAKE(Include, n) trims the filter range to exactly n rows.
    # No IFERROR fallback: when every toggle is FALSE, FILTER propagates
    # Excel's natural #CALC! rather than fabricating a first-predictor result.
    drop_local_name(sheet, "Coefficient_Name_Col")
    sheet.api.Names.Add(
        Name="Coefficient_Name_Col",
        RefersTo=(
            "=LAMBDA(All_Xs,[Include],"
            "LET("
            "n,COLUMNS(All_Xs),"
            "headers,TRANSPOSE(OFFSET(All_Xs,-1,0,1,n)),"
            "IF(ISOMITTED(Include),"
            "headers,"
            "FILTER(headers,TAKE(Include,n))"
            ")"
            "))"
        ),
    )

    # Ind_Var_Include: boolean range covering all predictor toggle cells.
    # Define it before x_s() so Excel resolves the dependency as a local name.
    drop_local_name(sheet, "Ind_Var_Filter")   # remove legacy name if present
    drop_local_name(sheet, "Ind_Var_Include")
    sheet.api.Names.Add(
        Name="Ind_Var_Include",
        RefersTo=f"={sname}!$B$3:$B$16000",
    )

    # x_s(): dynamic — only predictors toggled TRUE in col B via Ind_Var_Include.
    # A zero-argument LAMBDA forces Excel to resolve the selected columns at call
    # time instead of caching the range when the worksheet name is created.
    # No IFERROR fallback: when zero predictors are selected, FILTER propagates
    # Excel's natural #CALC! instead of silently substituting the first predictor.
    drop_local_name(sheet, "x_s")
    sheet.api.Names.Add(
        Name="x_s",
        RefersTo=(
            "=LAMBDA("
            "TRANSPOSE(FILTER(TRANSPOSE(All_Xs),"
            "TAKE(Ind_Var_Include,COLUMNS(All_Xs)))))"
        ),
    )

    # Zero_Predictors_Selected(): TRUE when every predictor toggle in
    # Ind_Var_Include is FALSE. Shared condition so it isn't duplicated
    # ad hoc across every cell that needs the zero-predictor branch.
    drop_local_name(sheet, "Zero_Predictors_Selected")
    sheet.api.Names.Add(
        Name="Zero_Predictors_Selected",
        RefersTo="=LAMBDA(COUNTIF(TAKE(Ind_Var_Include,COLUMNS(All_Xs)),TRUE)=0)",
    )

    drop_local_name(sheet, "fil")   # remove legacy name if present
    for name, ref in [
        ("y",   "=LifeExpectancyData[Life expectancy]"),
        ("Regression_Sample_Include", "=LifeExpectancyData[Full_Data]"),
        ("data_identifiers", "=LifeExpectancyData[Country]"),
    ]:
        drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=ref)

    # Allow_Intercept toggle lives in B2
    drop_local_name(sheet, "Allow_Intercept")
    sheet.api.Names.Add(
        Name="Allow_Intercept",
        RefersTo=f"={sname}!$B$2",
    )

    # alpha: confidence level input, lives in M12
    drop_local_name(sheet, "alpha")
    sheet.api.Names.Add(
        Name="alpha",
        RefersTo=f"={sname}!$M$12",
    )

    # ── Intercept-only closed-form helpers ──────────────────────────────────
    # Used by _write_coefficients and _write_prediction_interval when
    # Zero_Predictors_Selected() is TRUE and Allow_Intercept is TRUE: an
    # intercept-only OLS model (Y = b0 + error) is still statistically
    # well-defined even though FILTER(...,TAKE(Ind_Var_Include,...)) has
    # nothing to select. Bypasses x_s()/Coefficients()/Prediction_Interval()
    # entirely since Excel cannot represent a valid zero-column array.
    drop_local_name(sheet, "Intercept_Only_N")
    sheet.api.Names.Add(
        Name="Intercept_Only_N",
        RefersTo="=LAMBDA(COUNT(FILTER(y,Regression_Sample_Include)))",
    )

    drop_local_name(sheet, "Intercept_Only_Point")
    sheet.api.Names.Add(
        Name="Intercept_Only_Point",
        RefersTo="=LAMBDA(AVERAGE(FILTER(y,Regression_Sample_Include)))",
    )

    drop_local_name(sheet, "Intercept_Only_S")
    sheet.api.Names.Add(
        Name="Intercept_Only_S",
        RefersTo="=LAMBDA(STDEV.S(FILTER(y,Regression_Sample_Include)))",
    )

    drop_local_name(sheet, "Intercept_Only_SE")
    sheet.api.Names.Add(
        Name="Intercept_Only_SE",
        RefersTo="=LAMBDA(Intercept_Only_S()/SQRT(Intercept_Only_N()))",
    )

    drop_local_name(sheet, "Intercept_Only_DF")
    sheet.api.Names.Add(
        Name="Intercept_Only_DF",
        RefersTo="=LAMBDA(Intercept_Only_N()-1)",
    )

    drop_local_name(sheet, "pred_input")  # remove the legacy cached-range name

    # ── Chart data ranges (OFFSET-based, sized to n = $M$8 observations) ────────
    # These worksheet-scoped names feed chart SERIES formulas as
    # ='Regression'!<Name>, avoiding full-column references that degrade
    # performance and avoiding the unsupported # spill operator in chart formulas.
    for _name, _col_ltr in [
        ("RegChartQQX", col_letter(_C_AF)),        # Normal Scores Ranked
        ("RegChartQQY", col_letter(_C_AG)),        # Studentized Residuals Ranked
        ("RegChartFitY", col_letter(_C_Z)),        # Predicted Y (shared)
        ("RegChartResid", col_letter(_C_AA)),      # Residuals
        ("RegChartActY", col_letter(_C_Y)),        # Actual Y
        ("RegChartScaleLoc", col_letter(_C_AH)),   # Scale-Location
        ("RegChartCookDist", col_letter(_C_AE)),   # Cook's Distance
        ("RegChartLeverage", col_letter(_C_AC)),   # Hat Diagonal
        ("RegChartStudResid", col_letter(_C_AD)),  # Studentized Residuals
        ("RegChartPRESSResid", col_letter(_C_AI)), # PRESS Residual
    ]:
        drop_local_name(sheet, _name)
        sheet.api.Names.Add(
            Name=_name,
            RefersTo=f"=OFFSET('{sname}'!${_col_ltr}$2,1,0,MAX(IFERROR('{sname}'!$M$8,1),1),1)",
        )


# ── Section writers ───────────────────────────────────────────────────────────

def _write_model_selection(sheet: xw.Sheet, k: int) -> None:
    """Zone A–B: predictor labels + 'In linear model?' toggles."""
    section_heading(sheet, 1, _C_A, "MODEL SELECTION")
    val(sheet, 1, _C_B, "In linear model?")
    bold(sheet, 1, _C_B)

    # Row 2: Allow Intercept toggle (the named Allow_Intercept cell)
    val(sheet, 2, _C_A, "Allow Intercept")
    val(sheet, 2, _C_B, True)

    # A3: spill formula — fills predictor names from the table headers
    f(sheet, 3, _C_A, "=Coefficient_Name_Col(All_Xs)")

    # B3:B2+k contains one model-selection toggle per All_Xs column.
    for i in range(k):
        val(sheet, 3 + i, _C_B, True)

    # Orange for all user-editable toggle cells
    _input_range(sheet, 2, _C_B, 2 + k, _C_B)


def _write_boolean_validation(sheet: xw.Sheet) -> None:
    """B2:B16000 — in-cell dropdown restricted to TRUE / FALSE."""
    rng = sheet.range(rc(2, _C_B), rc(16000, _C_B)).api
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
    """Zone D–J: EDA stats for the predictors currently selected by x_s()."""
    section_heading(sheet, 1, _C_D, "PREDICTOR SUMMARY")

    for col, header in zip(
        [_C_D, _C_E, _C_F, _C_G, _C_H, _C_I, _C_J],
        ["", "Pearson R", "Spearman R", "Skewness", "Kurtosis", "VIF", "Tolerance"],
    ):
        val(sheet, 2, col, header)
    bold_row(sheet, 2, _C_D, _C_J)

    # Spill anchors at row 3 — each spills once per selected predictor
    f(sheet, 3, _C_D, "=Coefficient_Name_Col(All_Xs,Ind_Var_Include)")
    f(sheet, 3, _C_E, "=Pearson_R(x_s(),y,Regression_Sample_Include)")
    f(sheet, 3, _C_F, "=Spearman_R(x_s(),y,Regression_Sample_Include)")
    f(sheet, 3, _C_G, "=Skewness(x_s(),Regression_Sample_Include)")
    f(sheet, 3, _C_H, "=Kurtosis(x_s(),Regression_Sample_Include)")
    f(sheet, 3, _C_I, "=VIF(x_s(),Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_J, "=Tolerance(x_s(),Allow_Intercept,Regression_Sample_Include)")

    last = 2 + k
    sheet.range((rc(3, _C_E)), (rc(last, _C_J))).number_format = "0.00"


def _write_regression_outputs_header(sheet: xw.Sheet) -> None:
    section_heading(sheet, 1, _C_L, "REGRESSION OUTPUTS")
    val(sheet, 2, _C_P, "Predicted Variable")
    bold(sheet, 2, _C_P)
    f(sheet, 2, _C_Q, "=OFFSET(y,-1,0,1,1)")


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    """Cols L–M, rows 3–8."""
    section_heading(sheet, 3, _C_L, "REGRESSION STATISTICS")
    for row, label, formula in [
        (4, "Multiple R",        "=Multiple_R(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (5, "R Square",          "=R_squared(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (6, "Adjusted R Square", "=Adjusted_R2(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (7, "Standard Error",    "=SE_Regression(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (8, "Observations",      "=Observations(y,Regression_Sample_Include)"),
    ]:
        val(sheet, row, _C_L, label)
        f(sheet, row, _C_M, formula)
    sheet.range(rc(4, _C_M), rc(7, _C_M)).number_format = "0.0000"
    sheet.range(rc(8, _C_M), rc(8, _C_M)).number_format = "0"
    border_box(sheet, 3, _C_L, 8, _C_M)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    """Cols O–P, rows 3–10."""
    section_heading(sheet, 3, _C_O, "DIAGNOSTICS")
    for row, label, formula in [
        (4,  "PRESS",          "=PRESS(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (5,  "PRESS R²",  "=1-PRESS(x_s(),y,Allow_Intercept,Regression_Sample_Include)/SS_Total(y,Allow_Intercept,Regression_Sample_Include)"),
        (6,  "Mean Leverage",  "=(DF_Regression(x_s())+IF(Allow_Intercept,1,0))/Observations(y,Regression_Sample_Include)"),
        (7,  "AIC",            "=AIC(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (8,  "BIC",            "=BIC(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (9,  "AICc",           "=AICc(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
        (10, "QQ Correlation", "=QQ_Correlation(x_s(),y,Allow_Intercept,Regression_Sample_Include)"),
    ]:
        val(sheet, row, _C_O, label)
        f(sheet, row, _C_P, formula)
    sheet.range(rc(4, _C_P), rc(10, _C_P)).number_format = "0.0000"
    border_box(sheet, 3, _C_O, 10, _C_P)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at M12 — controls prediction interval confidence level."""
    val(sheet, 12, _C_L, "Alpha")
    bold(sheet, 12, _C_L)
    val(sheet, 12, _C_M, 0.05)
    format_input(sheet, 12, _C_M)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 13–17, cols L–Q."""
    section_heading(sheet, 13, _C_L, "ANOVA TABLE")

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        val(sheet, 14, col, header)
    bold_row(sheet, 14, _C_L, _C_Q)

    val(sheet, 15, _C_L, "Regression")
    f(sheet, 15, _C_M, "=DF_Regression(x_s())")
    f(sheet, 15, _C_N, "=SS_Regression(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 15, _C_O, "=MS_Regression(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 15, _C_P, "=F_Stat(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 15, _C_Q, "=P_Value_F(x_s(),y,Allow_Intercept,Regression_Sample_Include)")

    val(sheet, 16, _C_L, "Residual")
    f(sheet, 16, _C_M, "=DF_Residual(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 16, _C_N, "=SS_Residual(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 16, _C_O, "=MS_Residual(x_s(),y,Allow_Intercept,Regression_Sample_Include)")

    val(sheet, 17, _C_L, "Total")
    f(sheet, 17, _C_M, "=DF_Total(y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 17, _C_N, "=SS_Total(y,Allow_Intercept,Regression_Sample_Include)")

    sheet.range(rc(15, _C_M), rc(17, _C_M)).number_format = "0"
    sheet.range(rc(15, _C_N), rc(17, _C_N)).number_format = "0.0"
    sheet.range(rc(15, _C_O), rc(16, _C_O)).number_format = "0.0"
    sheet.range(rc(15, _C_P), rc(15, _C_P)).number_format = "0.0"
    sheet.range(rc(15, _C_Q), rc(15, _C_Q)).number_format = "0.0E+00"
    border_box(sheet, 13, _C_L, 17, _C_Q)


def _write_coefficients(sheet: xw.Sheet, k: int) -> None:
    """Cols L–S, rows 19+. Spills downward — nothing placed below row 39 in these cols."""
    section_heading(sheet, 19, _C_L, "COEFFICIENTS")

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q, _C_R, _C_S],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%", "Beta Weight"],
    ):
        val(sheet, 20, col, header)
    bold_row(sheet, 20, _C_L, _C_S)

    # Spill row labels aligned to selected predictors. Zero_Predictors_Selected()
    # branch computes a real intercept-only model (label "Intercept") instead of
    # fabricating a result for an unselected variable; NA() when nothing is fit.
    f(
        sheet,
        21,
        _C_L,
        '=IF(Zero_Predictors_Selected(),'
        'IF(Allow_Intercept,"Intercept",NA()),'
        'IF(Allow_Intercept,'
        'VSTACK("Intercept",Coefficient_Name_Col(All_Xs,Ind_Var_Include)),'
        'VSTACK("",Coefficient_Name_Col(All_Xs,Ind_Var_Include))))',
    )

    # Spill anchors at row 21 — pad with blank top row when intercept is disabled;
    # zero-predictor branch uses the closed-form intercept-only statistic, or
    # NA() when there is nothing to fit (no intercept, no predictors).
    f(sheet, 21, _C_M,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,Intercept_Only_Point(),NA()),'
       'IF(Allow_Intercept,Coefficients(x_s(),y,Allow_Intercept,Regression_Sample_Include),VSTACK("",Coefficients(x_s(),y,Allow_Intercept,Regression_Sample_Include))))')
    f(sheet, 21, _C_N,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,SE_Coefficients(x_s(),y,Allow_Intercept,Regression_Sample_Include),VSTACK("",SE_Coefficients(x_s(),y,Allow_Intercept,Regression_Sample_Include))))')
    f(sheet, 21, _C_O,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,Intercept_Only_Point()/Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,T_Stats(x_s(),y,Allow_Intercept,Regression_Sample_Include),VSTACK("",T_Stats(x_s(),y,Allow_Intercept,Regression_Sample_Include))))')
    f(sheet, 21, _C_P,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,T.DIST.2T(ABS(Intercept_Only_Point()/Intercept_Only_SE()),Intercept_Only_DF()),NA()),'
       'IF(Allow_Intercept,P_Values(x_s(),y,Allow_Intercept,Regression_Sample_Include),VSTACK("",P_Values(x_s(),y,Allow_Intercept,Regression_Sample_Include))))')
    f(sheet, 21, _C_Q,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,Intercept_Only_Point()-T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,CI_Lower(x_s(),y,Allow_Intercept,Regression_Sample_Include),VSTACK("",CI_Lower(x_s(),y,Allow_Intercept,Regression_Sample_Include))))')
    f(sheet, 21, _C_R,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,Intercept_Only_Point()+T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,CI_Upper(x_s(),y,Allow_Intercept,Regression_Sample_Include),VSTACK("",CI_Upper(x_s(),y,Allow_Intercept,Regression_Sample_Include))))')
    # Beta Weights: k×1 (no intercept row); always prepend blank to align with other columns.
    # No predictor exists to standardize in the zero-predictor branch, so render
    # blank (not an error) when Allow_Intercept is TRUE; NA() when nothing is fit.
    f(sheet, 21, _C_S,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,"",NA()),'
       'VSTACK("",Beta_Weights(x_s(),y,Allow_Intercept,Regression_Sample_Include)))')

    last_coef_row = 21 + k
    for col in [_C_M, _C_N, _C_O, _C_Q, _C_R, _C_S]:
        sheet.range(rc(21, col), rc(last_coef_row, col)).number_format = "0.0000"
    sheet.range(rc(21, _C_P), rc(last_coef_row, _C_P)).number_format = "0.0E+00"


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    """Zone U1:V8: boxed prediction interval output."""
    section_heading(sheet, 1, _C_U, "PREDICTION OUTPUTS")
    val(sheet, 2, _C_U, "PREDICTION INTERVAL")
    bold(sheet, 2, _C_U)
    for row, label in [
        (3, "Point Estimate"),
        (4, "SE Prediction"),
        (5, "t Critical"),
        (6, "Lower 95%"),
        (7, "Upper 95%"),
        (8, "Confidence Level"),
    ]:
        val(sheet, row, _C_U, label)
    # Zero_Predictors_Selected() branch computes the closed-form single-mean
    # prediction interval instead of feeding a fabricated first-predictor
    # input into Prediction_Interval(); NA() when there is nothing to fit.
    f(
        sheet,
        3,
        _C_V,
        "=IF(Zero_Predictors_Selected(),"
        "IF(Allow_Intercept,"
        "LET(point,Intercept_Only_Point(),"
        "se_pred,Intercept_Only_S()*SQRT(1+1/Intercept_Only_N()),"
        "t_crit,T.INV.2T(alpha,Intercept_Only_DF()),"
        "VSTACK(point,se_pred,t_crit,point-t_crit*se_pred,point+t_crit*se_pred,1-alpha)),"
        "NA()),"
        "LET(pred_input,VSTACK($V$12,"
        "FILTER($V$13:$V$30,TAKE(Ind_Var_Include,COLUMNS(All_Xs)))),"
        "Prediction_Interval(x_s(),y,pred_input,Allow_Intercept,"
        "Regression_Sample_Include,alpha)))",
    )
    sheet.range(rc(3, _C_V), rc(8, _C_V)).number_format = "0.0000"
    border_box(sheet, 1, _C_U, 8, _C_V)


def _write_prediction_inputs(sheet: xw.Sheet, k: int) -> None:
    """Zone U10:V12+k: per-predictor values used for the point prediction."""
    section_heading(sheet, 10, _C_U, "PREDICTION INPUTS")
    val(sheet, 11, _C_U, "Predictor")
    val(sheet, 11, _C_V, "Prediction Value")
    bold_row(sheet, 11, _C_U, _C_V)

    # Row 12: intercept (auto-set, still orange to show it's a value)
    val(sheet, 12, _C_U, "Intercept")
    f(sheet, 12, _C_V, "=IF(Allow_Intercept,1,0)")
    format_input(sheet, 12, _C_V)

    # T13: spill formula — fills predictor names from the table headers
    f(sheet, 13, _C_U, "=Coefficient_Name_Col(All_Xs)")

    # U13:U12+k — mean of each predictor column (filtered), individually overridable
    for i in range(k):
        f(sheet, 13 + i, _C_V, f"=AVERAGEIF(Regression_Sample_Include,TRUE,INDEX(All_Xs,,{i + 1}))")

    # Orange for all user-editable prediction value cells
    _input_range(sheet, 13, _C_V, 12 + k, _C_V)
    sheet.range(rc(13, _C_V), rc(12 + k, _C_V)).number_format = "0.0000"


def _write_residuals(sheet: xw.Sheet) -> None:
    """Residual diagnostic table — row identifiers + 11 diagnostics columns starting at _C_X."""
    section_heading(sheet, 1, _C_X, "RESIDUAL OUTPUT")

    # X2: dynamic header pulled from the data_identifiers named range, falling
    # back to a generic label when the optional named range is unset/invalid.
    f(sheet, 2, _C_X, '=IFERROR(OFFSET(data_identifiers,-1,0,1,1),"Observation")')

    for col, header in zip(
        [_C_Y, _C_Z, _C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF, _C_AG, _C_AH, _C_AI],
        [
            "Y", "Predicted Y", "Residuals", "LOOCV Residual",
            "Hat Diagonal", "Studentized Residuals", "Cook's Distance",
            "Normal Scores Ranked", "Studentized Residuals Ranked",
            "Scale-Location", "PRESS Residual",
        ],
    ):
        val(sheet, 2, col, header)
    bold_row(sheet, 2, _C_X, _C_AI)

    # X3: row labels — actual identifiers filtered to the sample, falling back
    # to generic "Observation N" numbering when data_identifiers errors out.
    f(
        sheet, 3, _C_X,
        '=IFERROR(FILTER(data_identifiers,Regression_Sample_Include),'
        'LAMBDA(obs,"Observation "&obs)(SEQUENCE($M$8)))',
    )
    # Spill anchors — each spills n rows downward
    f(sheet, 3, _C_Y,  "=Dependent_Var(y,Regression_Sample_Include)")
    f(sheet, 3, _C_Z,  "=Predictions(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_AA, "=Residuals(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_AB, "=Dependent_Var(y,Regression_Sample_Include)-LOOCV_prediction(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_AC, "=Hat_diagonal(x_s(),Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_AD, "=Studentized_Residuals(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_AE, "=Cooks_Distance(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    f(sheet, 3, _C_AF, "=SORT(Normal_Scores(y,Regression_Sample_Include))")
    f(sheet, 3, _C_AG, "=Studentized_Residuals_Ranked(x_s(),y,Allow_Intercept,Regression_Sample_Include)")
    # Scale-Location: SQRT(|Studentized_Residuals|) — horizontal spread should be flat.
    f(sheet, 3, _C_AH, "=SQRT(ABS(Studentized_Residuals(x_s(),y,Allow_Intercept,Regression_Sample_Include)))")
    # PRESS Residual: e_i / (1 - h_i) — large values flag high-influence observations.
    f(sheet, 3, _C_AI, "=Residuals(x_s(),y,Allow_Intercept,Regression_Sample_Include)/(1-Hat_diagonal(x_s(),Allow_Intercept,Regression_Sample_Include))")
    sheet.range(f"{col_letter(_C_Z)}:{col_letter(_C_AI)}").number_format = "0.0000"


def _write_diagnostic_charts(sheet: xw.Sheet) -> None:
    """Create 7 pre-built diagnostic charts to the right of the Residual Output section."""
    start_left = sheet.range(a1(1, _C_AI + 1)).left
    start_top = sheet.range("A3").top

    col_step = _CHART_WIDTH + _CHART_GAP
    row_step = _CHART_HEIGHT + _CHART_GAP

    def _pos(grid_row: int, grid_col: int) -> tuple[float, float]:
        return (
            start_left + (grid_col - 1) * col_step,
            start_top + (grid_row - 1) * row_step,
        )

    sname = REGRESSION_SHEET_NAME

    def _name_ref(local_name: str) -> str:
        return f"='{sname}'!{local_name}"

    chart_specs = [
        (
            "Residuals vs. Fitted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartResid"),
            "Fitted Values", "Residuals", 1, 1,
        ),
        (
            "Normal Q-Q", "scatter",
            _name_ref("RegChartQQX"),
            _name_ref("RegChartQQY"),
            "Theoretical Quantiles", "Studentized Residuals", 1, 2,
        ),
        (
            "Actual vs. Predicted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartActY"),
            "Predicted Y", "Actual Y", 2, 1,
        ),
        (
            "Scale-Location", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartScaleLoc"),
            "Fitted Values", "√|Studentized Residual|", 2, 2,
        ),
        (
            "Cook's Distance", "bar",
            None,
            _name_ref("RegChartCookDist"),
            "Observation", "Cook's Distance", 3, 1,
        ),
        (
            "Studentized Residuals vs. Leverage", "scatter",
            _name_ref("RegChartLeverage"),
            _name_ref("RegChartStudResid"),
            "Leverage (Hat Diagonal)", "Studentized Residuals", 3, 2,
        ),
        (
            "PRESS Residuals", "bar",
            None,
            _name_ref("RegChartPRESSResid"),
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

    def _add_identity_line(chart: Any, name_ref: str) -> None:
        """Add a dotted y=x reference series using one column for both axes.

        Pointing XValues and Values at the same named range guarantees every
        plotted point sits exactly on the identity line — a real data series
        stays correct if the chart is resized, moved, or the axis scaling
        changes, unlike a shape drawn at fixed plot-area pixel coordinates.
        """
        series = chart.SeriesCollection().NewSeries()
        series.XValues = name_ref
        series.Values = name_ref
        series.Name = "Identity"
        series.ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS
        series.Format.Line.ForeColor.RGB = excel_color((120, 120, 120))
        series.Format.Line.DashStyle = 3  # msoLineRoundDot
        series.Format.Line.Weight = 1.25

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
        chart.ChartTitle.Format.Fill.ForeColor.RGB = excel_color(_HEADER)

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
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "RegChartQQX", "RegChartQQY")
            _add_identity_line(chart, x_addr)
        if title == "Actual vs. Predicted":
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "RegChartFitY", "RegChartActY")
            _add_identity_line(chart, x_addr)


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

    sheet.range(rc(2, _C_A), rc(2, _C_AI)).api.WrapText = True

    # Column widths (U = prediction labels, V = prediction values;
    # X = row identifiers, diagnostics start at Y)
    for col_letter, width in {
        "A": 28, "B": 14,
        "C": 2,   # thin gap
        "D": 28, "E": 8, "F": 10, "G": 10, "H": 8, "I": 8, "J": 10,
        "K": 2,   # thin gap
        "L": 22, "M": 12, "N": 12, "O": 14, "P": 10, "Q": 13, "R": 10, "S": 12,
        "T": 2,   # thin gap
        "U": 20, "V": 14,  # prediction labels / values
        "W": 2,   # thin gap
        "X": 16,  # data identifiers (row labels)
        "Y": 10, "Z": 9, "AA": 10, "AB": 9, "AC": 9, "AD": 12, "AE": 9, "AF": 14,
        "AG": 17, "AH": 14, "AI": 15,  # Scale-Location / PRESS Residual
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

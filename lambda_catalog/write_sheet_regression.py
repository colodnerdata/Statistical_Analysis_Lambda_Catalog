"""
write_sheet_regression.py
Writes the ToolPak-style Regression sheet into any target workbook.

Layout (five horizontal zones):
  Col A–B        — Prediction Inputs: predictor labels + "In linear model?" toggles
                   B2 = Allow_Intercept toggle; B3:B16000 = per-predictor on/off (orange)
  Col C          — thin gap (width 2)
  Col D–J        — Predictor Summary: names + Pearson R, Spearman R, Skewness, Kurtosis,
                   VIF, Tolerance (always using All_Xs — all 18 predictors)
  Col K          — thin gap (width 2)
  Col L–R        — Regression Outputs: Statistics (L–M rows 3–8), Diagnostics (O–P rows 3–10),
                   Sheet-scoped names: All_Xs, Ind_Var_Filter ($B$3:$B$16000), x_s (filtered),
                   Coefficient_Name_Col([Filter] optional), Allow_Intercept, alpha, pred_input
                   Alpha input (M12), ANOVA Table (rows 13–17),
                   Coefficients (rows 19+, spills downward)
  Col S          — thin gap (width 2)
  Col T–U        — Prediction Outputs: Prediction Interval (T1:U8, boxed),
                   Prediction Inputs (T10+, no box — dynamic height)
  Col V          — thin gap (width 2)
    Col W–AF       — Residual Output (9 columns, spills downward from row 3)
"""
from __future__ import annotations

import xlwings as xw


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

# ── 1-based column indices ─────────────────────────────────────────────────────

# Zone 1: Prediction Inputs
_C_A = 1    # predictor labels
_C_B = 2    # "In linear model?" toggles (orange user input); B2 = Allow_Intercept

# Zone 2: Predictor Summary
_C_C = 3    # thin gap
_C_D = 4    # predictor names
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

# Zone 4: Prediction Outputs
_C_S = 19   # thin gap
_C_T = 20   # prediction interval labels / prediction input labels
_C_U = 21   # prediction interval values / prediction input values (orange)

# Zone 5: Residual Output
_C_V = 22   # thin gap
_C_W = 23   # section heading anchor
_C_X = 24   # Y
_C_Y = 25   # Predicted Y
_C_Z = 26   # Residuals
_C_AA = 27  # LOOCV residual
_C_AB = 28  # Hat Diagonal
_C_AC = 29  # Studentized Residuals
_C_AD = 30  # Cook's Distance
_C_AE = 31  # Normal Scores Ranked
_C_AF = 32  # Studentized Residuals Ranked


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


# ── Visual formatting helpers ─────────────────────────────────────────────────

_YELLOW = (255, 255, 0)   # section headings
_ORANGE = (255, 192, 0)   # user-editable input cells


def _section_heading(sheet: xw.Sheet, row: int, col: int, label: str) -> None:
    _v(sheet, row, col, label)
    sheet.range(_rc(row, col)).api.Font.Bold = True
    sheet.range(_rc(row, col)).color = _YELLOW


def _orange_input(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(_rc(row, col)).color = _ORANGE


def _orange_range(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    sheet.range(_rc(r1, c1), _rc(r2, c2)).color = _ORANGE


def _border_box(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    rng = sheet.range(_rc(r1, c1), _rc(r2, c2)).api
    for edge in [7, 8, 9, 10]:   # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        rng.Borders(edge).LineStyle = 1   # xlContinuous
        rng.Borders(edge).Weight = 2      # xlThin


# ── Local name management ─────────────────────────────────────────────────────

def _drop_local_name(sheet: xw.Sheet, name: str) -> None:
    for idx in range(sheet.api.Names.Count, 0, -1):
        local = sheet.api.Names(idx).Name.split("!", 1)[-1]
        if local.lower() == name.lower():
            sheet.api.Names(idx).Delete()


def _setup_local_names(sheet: xw.Sheet, k: int) -> None:
    """Register sheet-scoped names used by every formula on this sheet."""
    sname = sheet.name

    # All_Xs: full 18-predictor range — predictor summary always uses this
    _drop_local_name(sheet, "All_Xs")
    sheet.api.Names.Add(
        Name="All_Xs",
        RefersTo="=LifeExpectancyData[[Adult Mortality]:[Schooling]]",
    )

    # Retire Selected_Col_Nums — clean it up from any existing workbook.
    _drop_local_name(sheet, "Selected_Col_Nums")

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
    # Row U12 = intercept (auto), U13:U{12+k} = per-predictor values
    _drop_local_name(sheet, "pred_input")
    sheet.api.Names.Add(
        Name="pred_input",
        RefersTo=f"={sname}!$U$12:INDEX({sname}!$U:$U,12+COLUMNS(x_s))",
    )


# ── Section writers ───────────────────────────────────────────────────────────

def _write_prediction_inputs(sheet: xw.Sheet) -> None:
    """Zone A–B: predictor labels + 'In linear model?' toggles."""
    _section_heading(sheet, 1, _C_A, "PREDICTION INPUTS")
    _v(sheet, 1, _C_B, "In linear model?")
    _bold(sheet, 1, _C_B)

    # Row 2: Allow Intercept toggle (the named Allow_Intercept cell)
    _v(sheet, 2, _C_A, "Allow Intercept")
    _v(sheet, 2, _C_B, True)

    # A3: spill formula — fills predictor names from the table headers
    _f(sheet, 3, _C_A, "=Coefficient_Name_Col(All_Xs)")

    # Rows 3+: write B toggles only (A names come from the spill formula above)
    for i in range(len(PREDICTOR_NAMES)):
        _v(sheet, 3 + i, _C_B, True)

    # Orange for all user-editable toggle cells
    k = len(PREDICTOR_NAMES)
    _orange_range(sheet, 2, _C_B, 2 + k, _C_B)


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


def _write_predictor_summary(sheet: xw.Sheet) -> None:
    """Zone D–J: EDA stats for the predictors currently selected into x_s."""
    _section_heading(sheet, 1, _C_D, "PREDICTOR SUMMARY")

    for col, header in zip(
        [_C_D, _C_E, _C_F, _C_G, _C_H, _C_I, _C_J],
        ["", "Pearson R", "Spearman R", "Skewness", "Kurtosis", "VIF", "Tolerance"],
    ):
        _v(sheet, 2, col, header)
    _bold_row(sheet, 2, _C_D, _C_J)

    # Spill anchors at row 3 — each spills once per selected predictor
    _f(sheet, 3, _C_D, "=Coefficient_Name_Col(All_Xs,Ind_Var_Filter)")
    _f(sheet, 3, _C_E, "=Pearson_R(x_s,y,fil)")
    _f(sheet, 3, _C_F, "=Spearman_R(x_s,y,fil)")
    _f(sheet, 3, _C_G, "=Skewness(x_s,fil)")
    _f(sheet, 3, _C_H, "=Kurtosis(x_s,fil)")
    _f(sheet, 3, _C_I, "=VIF(x_s,Allow_Intercept,fil)")
    _f(sheet, 3, _C_J, "=Tolerance(x_s,Allow_Intercept,fil)")

    last = 3 + len(PREDICTOR_NAMES) - 1
    sheet.range((_rc(3, _C_E)), (_rc(last, _C_F))).number_format = "0.000"
    sheet.range((_rc(3, _C_G)), (_rc(last, _C_H))).number_format = "0.00"
    sheet.range((_rc(3, _C_I)), (_rc(last, _C_I))).number_format = "0.00"
    sheet.range((_rc(3, _C_J)), (_rc(last, _C_J))).number_format = "0.000"


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
        _v(sheet, row, _C_L, label)
        _f(sheet, row, _C_M, formula)
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
        _v(sheet, row, _C_O, label)
        _f(sheet, row, _C_P, formula)
    _border_box(sheet, 3, _C_O, 10, _C_P)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at M12 — controls prediction interval confidence level."""
    _v(sheet, 12, _C_L, "Alpha")
    _bold(sheet, 12, _C_L)
    _v(sheet, 12, _C_M, 0.05)
    _orange_input(sheet, 12, _C_M)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 13–17, cols L–Q."""
    _section_heading(sheet, 13, _C_L, "ANOVA TABLE")

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        _v(sheet, 14, col, header)
    _bold_row(sheet, 14, _C_L, _C_Q)

    _v(sheet, 15, _C_L, "Regression")
    _f(sheet, 15, _C_M, "=DF_Regression(x_s)")
    _f(sheet, 15, _C_N, "=SS_Regression(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 15, _C_O, "=MS_Regression(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 15, _C_P, "=F_Stat(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 15, _C_Q, "=P_Value_F(x_s,y,Allow_Intercept,fil)")

    _v(sheet, 16, _C_L, "Residual")
    _f(sheet, 16, _C_M, "=DF_Residual(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 16, _C_N, "=SS_Residual(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 16, _C_O, "=MS_Residual(x_s,y,Allow_Intercept,fil)")

    _v(sheet, 17, _C_L, "Total")
    _f(sheet, 17, _C_M, "=DF_Total(y,Allow_Intercept,fil)")
    _f(sheet, 17, _C_N, "=SS_Total(y,Allow_Intercept,fil)")

    _border_box(sheet, 13, _C_L, 17, _C_Q)


def _write_coefficients(sheet: xw.Sheet) -> None:
    """Cols L–R, rows 19+. Spills downward — nothing placed below row 39 in these cols."""
    _section_heading(sheet, 19, _C_L, "COEFFICIENTS")

    for col, header in zip(
        [_C_L, _C_M, _C_N, _C_O, _C_P, _C_Q, _C_R],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"],
    ):
        _v(sheet, 20, col, header)
    _bold_row(sheet, 20, _C_L, _C_R)

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


def _write_prediction_outputs(sheet: xw.Sheet, k: int) -> None:
    """Zone T–U: prediction interval (boxed) + prediction inputs (no box, dynamic)."""

    # ── Prediction Interval (T1:U8, fixed height → box) ──────────────────────
    _section_heading(sheet, 1, _C_T, "PREDICTION OUTPUTS")
    _v(sheet, 2, _C_T, "PREDICTION INTERVAL")
    _bold(sheet, 2, _C_T)

    for row, label, idx in [
        (3, "Point Estimate",   1),
        (4, "SE Prediction",    2),
        (5, "t Critical",       3),
        (6, "Lower 95%",        4),
        (7, "Upper 95%",        5),
        (8, "Confidence Level", 6),
    ]:
        _v(sheet, row, _C_T, label)
        _f(sheet, row, _C_U,
           f"=INDEX(Prediction_Interval(x_s,y,pred_input,Allow_Intercept,fil,alpha),{idx})")

    _border_box(sheet, 1, _C_T, 8, _C_U)

    # ── Prediction Inputs (T10+, no box — height depends on k) ───────────────
    _section_heading(sheet, 10, _C_T, "PREDICTION INPUTS")
    _v(sheet, 11, _C_T, "Predictor")
    _v(sheet, 11, _C_U, "Prediction Value")
    _bold_row(sheet, 11, _C_T, _C_U)

    # Row 12: intercept (auto-set, still orange to show it's a value)
    _v(sheet, 12, _C_T, "Intercept")
    _f(sheet, 12, _C_U, "=IF(Allow_Intercept,1,0)")
    _orange_input(sheet, 12, _C_U)

    # T13: spill formula — fills predictor names from the table headers
    _f(sheet, 13, _C_T, "=Coefficient_Name_Col(All_Xs)")

    # Rows 13+: write U values only (T names come from the spill formula above)
    for i in range(len(PREDICTOR_NAMES)):
        _v(sheet, 13 + i, _C_U, 0.0)

    # Orange for all user-editable prediction value cells
    _orange_range(sheet, 13, _C_U, 12 + k, _C_U)


def _write_residuals(sheet: xw.Sheet) -> None:
    """Zone X–AF: residual diagnostic table."""
    _section_heading(sheet, 1, _C_W, "RESIDUAL OUTPUT")

    for col, header in zip(
        [_C_X, _C_Y, _C_Z, _C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF],
        [
            "Y", "Predicted Y", "Residuals", "LOOCV Residual",
            "Hat Diagonal", "Studentized Residuals", "Cook's Distance",
            "Normal Scores Ranked", "Studentized Residuals Ranked",
        ],
    ):
        _v(sheet, 2, col, header)
    _bold_row(sheet, 2, _C_X, _C_AF)

    # Spill anchors — each spills n rows downward
    _f(sheet, 3, _C_X,  "=Dependent_Var(y,fil)")
    _f(sheet, 3, _C_Y,  "=Predictions(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_Z,  "=Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AA, "=Dependent_Var(y,fil)-LOOCV_prediction(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AB, "=Hat_diagonal(x_s,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AC, "=Studentized_Residuals(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AD, "=Cooks_Distance(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_AE, "=SORT(Normal_Scores(y,fil))")
    _f(sheet, 3, _C_AF, "=Studentized_Residuals_Ranked(x_s,y,Allow_Intercept,fil)")


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

    _write_prediction_inputs(sheet)
    _write_boolean_validation(sheet)
    _write_predictor_summary(sheet)
    _write_regression_outputs_header(sheet)
    _write_regression_statistics(sheet)
    _write_diagnostics(sheet)
    _write_alpha(sheet)
    _write_anova(sheet)
    _write_coefficients(sheet)
    _write_prediction_outputs(sheet, k)
    _write_residuals(sheet)

    # Column widths
    for col_letter, width in {
        "A": 28, "B": 14, "C": 2,
        "D": 28, "E": 10, "F": 11, "G": 10, "H": 10, "I": 10, "J": 10,
        "K": 2,
        "L": 22, "M": 14, "N": 14, "O": 20, "P": 14, "Q": 16, "R": 14,
        "S": 2,
        "T": 20, "U": 14,
        "V": 2,
        "W": 4, "X": 12, "Y": 14, "Z": 12,
        "AA": 14, "AB": 14, "AC": 22, "AD": 14,
        "AE": 22, "AF": 26,
    }.items():
        sheet.range(f"{col_letter}:{col_letter}").column_width = width

    # Freeze top 2 rows
    sheet.activate()
    sheet.range("A3").select()
    win = sheet.api.Application.ActiveWindow
    win.FreezePanes = False
    win.SplitRow = 2
    win.SplitColumn = 0
    win.FreezePanes = True

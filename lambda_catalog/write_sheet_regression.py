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
  Cols L–O       — residual table (fixed columns, independent of k)
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
_C_O = 15   # residual: LOOCV residuals


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
    sheet.range(_rc(row, col)).formula = formula


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
        (4, "PRESS",          "=PRESS(x_s,y,Allow_Intercept,fil)"),
        (5, "PRESS R²",  "=1-PRESS(x_s,y,Allow_Intercept,fil)/SS_Total(y,Allow_Intercept,fil)"),
        (6, "Mean Leverage",  "=(DF_Regression(x_s)+IF(Allow_Intercept,1,0))/Observations(y,fil)"),
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
        [_C_L, _C_M, _C_N, _C_O],
        ["Observation", "Predicted Y", "Residuals", "LOOCV Residuals"],
    ):
        _v(sheet, 2, col, header)
    _bold_row(sheet, 2, _C_L, _C_O)

    # Spill anchors — each spills n rows downward
    _f(sheet, 3, _C_L, "=SEQUENCE(Observations(y,fil))")
    _f(sheet, 3, _C_M, "=Predictions(x_s,y,Allow_Intercept,fil)")
    _f(sheet, 3, _C_N, "=Residuals(x_s,y,Allow_Intercept,fil)")
    _v(sheet, 3, _C_O, "(paste LOOCV formula here)")


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

    # Column widths
    for col_letter, width in {
        "A": 30, "B": 12, "C": 24, "D": 14, "E": 14,
        "F": 16, "G": 14, "H": 16, "I": 22, "J": 14,
        "K": 3,  "L": 14, "M": 14, "N": 12, "O": 16,
    }.items():
        sheet.range(f"{col_letter}:{col_letter}").column_width = width

    # Freeze top 2 rows
    sheet.api.Application.ActiveWindow.SplitRow = 2
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

"""Build and write the MLR_Observation_Test worksheet for observation-level formula smoke tests."""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

from .analyze_life_expectancy import FEATURE_COLUMNS, RegressionObservationVectors, calculate_regression_observation_vectors
from .workbook_helpers import group_and_hide_columns, reset_column_groups
from .write_sheet_mlr_vector_outputs_test import _col_letter, _delete_sheet_scoped_name_if_present

_MLR_K_VALUES: list[int] = [1, 5, 10, 18]
_TESTS_COLS = 10
_CALC_START_COL = 11
_EXPECTED_START_COL = 19
_TOTAL_COLS = 26

_STATS: list[tuple[str, str, int, str]] = [
    ("Observation_Num", "Observation_Num", 0, "0"),
    ("Percentile", "Percentile", 4, "0.0000"),
    ("Y_Ranked", "Y_Ranked", 3, "0.000"),
    ("Normal_Scores", "Normal_Scores", 3, "0.000"),
    ("Predictions", "Predictions", 3, "0.000"),
    ("Residuals", "Residuals", 3, "0.000"),
    ("Scaled_Residuals", "Scaled_Residuals", 3, "0.000"),
    ("Scaled_Residuals_Ranked", "Scaled_Residuals_Ranked", 3, "0.000"),
]


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    for name, refers_to in (("y", "=LifeExpectancyData[Life expectancy]"), ("fil", "=LifeExpectancyData[Full_Data]")):
        _delete_sheet_scoped_name_if_present(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)


def _calc_formula(k: int, allow_intercept: bool, func_name: str) -> str:
    allow_arg = "TRUE" if allow_intercept else "FALSE"
    if func_name in {"Observation_Num", "Percentile", "Y_Ranked", "Normal_Scores"}:
        return f"={func_name}(y,fil)"
    return f"=LET(x_s,OFFSET(y,0,1,ROWS(y),{k}),{func_name}(x_s,y,{allow_arg},fil))"


def build_mlr_observation_row_configs(csv_path: Path) -> list[tuple[int, bool, RegressionObservationVectors]]:
    row_configs: list[tuple[int, bool, RegressionObservationVectors]] = []
    for k in _MLR_K_VALUES:
        for allow_intercept in (True, False):
            vectors = calculate_regression_observation_vectors(csv_path, allow_intercept, FEATURE_COLUMNS[:k])
            row_configs.append((k, allow_intercept, vectors))
    return row_configs


def write_mlr_observation_test_sheet(workbook: xw.Book, row_configs: list[tuple[int, bool, RegressionObservationVectors]]) -> None:
    sheet_name = "MLR_Observation_Test"
    sheet = workbook.sheets[sheet_name] if sheet_name in {s.name for s in workbook.sheets} else workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])
    sheet.api.Cells.Clear()
    reset_column_groups(sheet)
    _set_sheet_scoped_names(sheet)

    sheet.range((1, 1)).value = "Smoke Test"
    for i, (display_name, _, _, _) in enumerate(_STATS):
        sheet.range((1, 2 + i)).value = f"{display_name} Match"
        sheet.range((1, _CALC_START_COL + i)).value = display_name
        sheet.range((1, _EXPECTED_START_COL + i)).value = f"{display_name} (Exp.)"

    row = 2
    for k, allow_intercept, vectors in row_configs:
        n = len(vectors.observation_num)
        predictors = ", ".join(FEATURE_COLUMNS[:k])
        sheet.range((row, 1)).value = f"k={k}, Allow_Intercept={'TRUE' if allow_intercept else 'FALSE'} | {predictors}"
        row += 1

        exp_cols = [
            vectors.observation_num, vectors.percentile, vectors.y_ranked, vectors.normal_scores,
            vectors.predictions, vectors.residuals, vectors.scaled_residuals, vectors.scaled_residuals_ranked,
        ]
        for i, (_, fn_name, digits, num_fmt) in enumerate(_STATS):
            calc_col = _CALC_START_COL + i
            exp_col = _EXPECTED_START_COL + i
            match_col = 2 + i
            sheet.range((row, calc_col)).api.Formula2 = _calc_formula(k, allow_intercept, fn_name)
            for j, val in enumerate(exp_cols[i]):
                c = sheet.range((row + j, exp_col))
                c.api.NumberFormat = num_fmt
                c.value = val
            last = row + n - 1
            e_col = _col_letter(exp_col)
            c_col = _col_letter(calc_col)
            sheet.range((row, match_col)).api.Formula2 = f"=AND(ROUND({e_col}{row}:{e_col}{last},{digits*2})=ROUND({c_col}{row}#,{digits*2}))"
            sheet.range((row, calc_col), (last, calc_col)).api.NumberFormat = num_fmt

        refs = ",".join(f"{_col_letter(2+i)}{row}" for i in range(len(_STATS)))
        sheet.range((row, 1)).formula = f"=AND({refs})"
        row += n + 1

    sheet.range((1, 1), (1, _TOTAL_COLS)).api.WrapText = True
    sheet.range((1, 1), (row, _TOTAL_COLS)).columns.autofit()
    group_and_hide_columns(sheet, 2, _TESTS_COLS)
    group_and_hide_columns(sheet, _EXPECTED_START_COL, _TOTAL_COLS)
    sheet.activate()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

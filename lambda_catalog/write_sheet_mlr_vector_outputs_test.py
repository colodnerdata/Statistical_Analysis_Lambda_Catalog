"""Build and write the MLR_Vector_Outputs_Test worksheet for regression vector formula smoke tests."""
from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

import xlwings as xw

from .regression_shared import FEATURE_COLUMNS
from .workbook_helpers import _drop_local_name, reset_column_groups

if TYPE_CHECKING:
    from .regression_shared import RegressionVectors


_MLR_K_VALUES: list[int] = [1, 5, 10, 18]

_ALPHA = 0.05

_TERM_COL = 1
_CALC_START_COL = 2
_TOTAL_COLS = 12  # 1 (term) + 8 (_STATS) + 3 (_K_STATS)

_CI_FUNCS = frozenset({"CI_Lower", "CI_Upper"})

_D = 3
_STATS: list[tuple[str, str, int, str]] = [
    ("Coefficients",      "Coefficients",      _D, f"0.{'0' * _D}"),
    ("SE_Coefficients",   "SE_Coefficients",   _D, f"0.{'0' * _D}"),
    ("T_Stats",           "T_Stats",           _D, f"0.{'0' * _D}"),
    ("P_Values",          "P_Values",          _D, f"0.{'0' * _D}E+00"),
    ("CI_Lower",          "CI_Lower",          _D, f"0.{'0' * _D}"),
    ("CI_Upper",          "CI_Upper",          _D, f"0.{'0' * _D}"),
    ("Partial_R2",        "Partial_R2",        _D, f"0.{'0' * _D}"),
    ("Partial_Corr",      "Partial_Correlation", _D, f"0.{'0' * _D}"),
]

# k-only stats: return k rows (no intercept row); formula starts one row below intercept.
# needs_y=True means the function takes (X_s, Y, ...) rather than (X_s, ...).
_K_STATS: list[tuple[str, str, str, bool]] = [
    ("VIF",          "VIF",          "0.000", False),
    ("Tolerance",    "Tolerance",    "0.000", False),
    ("Beta_Weights", "Beta_Weights", "0.000", True),
]


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    _drop_local_name(sheet, "fil")  # remove legacy name if present
    for name, refers_to in (
        ("y", "=LifeExpectancyData[Life expectancy]"),
        ("Regression_Sample_Include", "=LifeExpectancyData[Full_Data]"),
    ):
        _drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)


def _calc_formula(k: int, allow_intercept: bool, func_name: str) -> str:
    allow_arg = "TRUE" if allow_intercept else "FALSE"
    extra = f",{_ALPHA}" if func_name in _CI_FUNCS else ""
    return f"=LET(x_s,OFFSET(y,0,1,ROWS(y),{k}),{func_name}(x_s,y,{allow_arg},Regression_Sample_Include{extra}))"


def _calc_k_formula(k: int, allow_intercept: bool, func_name: str, needs_y: bool) -> str:
    allow_arg = "TRUE" if allow_intercept else "FALSE"
    y_arg = ",y" if needs_y else ""
    return f"=LET(x_s,OFFSET(y,0,1,ROWS(y),{k}),{func_name}(x_s{y_arg},{allow_arg},Regression_Sample_Include))"


def build_mlr_vector_row_configs(
    csv_path: Path,
) -> list[tuple[int, bool, RegressionVectors]]:
    from .analyze_life_expectancy import calculate_regression_vectors

    row_configs: list[tuple[int, bool, RegressionVectors]] = []
    for k in _MLR_K_VALUES:
        for allow_intercept in (True, False):
            vectors = calculate_regression_vectors(
                input_csv_path=csv_path,
                include_intercept=allow_intercept,
                feature_columns=FEATURE_COLUMNS[:k],
            )
            row_configs.append((k, allow_intercept, vectors))
    return row_configs


def _write_section(
    sheet: xw.Sheet,
    k: int,
    allow_intercept: bool,
    vectors: RegressionVectors,
    first_data_row: int,
) -> None:
    for term_idx, term_name in enumerate(vectors.term_names):
        sheet.range((first_data_row + term_idx, _TERM_COL)).value = term_name
    for stat_idx, (_, func_name, _, num_fmt) in enumerate(_STATS):
        calc_col = _CALC_START_COL + stat_idx
        sheet.range((first_data_row, calc_col)).api.Formula2 = (
            _calc_formula(k, allow_intercept, func_name)
        )
        sheet.range((first_data_row, calc_col)).api.NumberFormat = num_fmt
    # k-only stats return k values (no intercept row), so start one row lower when
    # the model includes an intercept to align with the first predictor row.
    k_row = first_data_row + (1 if allow_intercept else 0)
    for k_idx, (_, func_name, num_fmt, needs_y) in enumerate(_K_STATS):
        calc_col = _CALC_START_COL + len(_STATS) + k_idx
        sheet.range((k_row, calc_col)).api.Formula2 = (
            _calc_k_formula(k, allow_intercept, func_name, needs_y)
        )
        sheet.range((k_row, calc_col)).api.NumberFormat = num_fmt


def write_mlr_vector_outputs_test_sheet(
    workbook: xw.Book,
    row_configs: list[tuple[int, bool, RegressionVectors]],
) -> None:
    sheet_name = "MLR_Vector_Outputs_Test"

    sheet_names = {sheet.name for sheet in workbook.sheets}
    if sheet_name in sheet_names:
        sheet = workbook.sheets[sheet_name]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = sheet_name
    else:
        sheet = workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])

    sheet.api.Cells.Clear()
    reset_column_groups(sheet)
    _set_sheet_scoped_names(sheet)

    header_row = 1
    sheet.range((header_row, _TERM_COL)).value = "Term"
    for i, (display_name, _, _, _) in enumerate(_STATS):
        sheet.range((header_row, _CALC_START_COL + i)).value = display_name
    for j, (display_name, _, _, _) in enumerate(_K_STATS):
        sheet.range((header_row, _CALC_START_COL + len(_STATS) + j)).value = display_name

    current_row = header_row + 1

    for k, allow_intercept, vectors in row_configs:
        n_terms = len(vectors.coefficients)
        sheet.range((current_row, _TERM_COL)).value = (
            f"k={k} | Intercept={'TRUE' if allow_intercept else 'FALSE'}"
        )
        current_row += 1
        _write_section(sheet, k, allow_intercept, vectors, current_row)
        current_row += n_terms + 1

    last_content_row = current_row - 1
    sheet.range((header_row, 1), (header_row, _TOTAL_COLS)).api.WrapText = True
    sheet.range((header_row, 1), (header_row, _TOTAL_COLS)).api.EntireRow.AutoFit()
    sheet.range((header_row, 1), (last_content_row, _TOTAL_COLS)).columns.autofit()

    sheet.activate()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

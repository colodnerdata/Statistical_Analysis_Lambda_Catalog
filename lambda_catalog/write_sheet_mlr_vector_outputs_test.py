"""Build and write the MLR_Vector_Outputs_Test worksheet for regression vector formula smoke tests."""
from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

import xlwings as xw

from .catalog_schema import catalog_argument_names
from .make_test_sheet import build_call, design_expression, predictor_expression
from .regression_shared import FEATURE_COLUMNS
from .workbook_helpers import (
    drop_local_name,
    reset_column_groups,
    safe_activate,
    safe_freeze_top_row,
)

if TYPE_CHECKING:
    from .regression_shared import RegressionVectors


_MLR_K_VALUES: list[int] = [1, 5, 10, 18]

_ALPHA = 0.05

_TERM_COL = 1
_CALC_START_COL = 2
_TOTAL_COLS = 12  # 1 (term) + 8 (_STATS) + 3 (_K_STATS)

_CI_FUNCS = frozenset({"Confidence_Interval_Lower", "Confidence_Interval_Upper"})

_D = 3
_STATS: list[tuple[str, str, int, str]] = [
    ("Coefficients",      "Coefficients",      _D, f"0.{'0' * _D}"),
    ("SE_Coefficients",   "SE_Coefficients",   _D, f"0.{'0' * _D}"),
    ("T_Statistics",           "T_Statistics",           _D, f"0.{'0' * _D}"),
    ("P_Values",          "P_Values",          _D, f"0.{'0' * _D}E+00"),
    ("Confidence_Interval_Lower",          "Confidence_Interval_Lower",          _D, f"0.{'0' * _D}"),
    ("Confidence_Interval_Upper",          "Confidence_Interval_Upper",          _D, f"0.{'0' * _D}"),
    ("Partial_R_Squared",        "Partial_R_Squared",        _D, f"0.{'0' * _D}"),
    ("Partial_Corr",      "Partial_Correlation", _D, f"0.{'0' * _D}"),
]

# k-only stats: return k rows (no intercept row); formula starts one row below intercept.
# needs_y=True means the function returns one row per predictor with a
# response argument; the argument list itself comes from the catalog.
_K_STATS: list[tuple[str, str, str, bool]] = [
    ("VIF",          "VIF",          "0.000", False),
    ("Tolerance",    "Tolerance",    "0.000", False),
    ("Beta_Weights", "Beta_Weights", "0.000", True),
]


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    drop_local_name(sheet, "fil")  # remove legacy name if present
    for name, refers_to in (
        ("y", "=LifeExpectancyData[Life expectancy]"),
        ("Regression_Sample_Include", "=LifeExpectancyData[Full_Data]"),
    ):
        drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)


def _reference_map(k: int, has_intercept: bool, func_name: str) -> dict[str, str | None]:
    """Formula text for each argument name a catalog function may declare.

    ``X`` and ``Predictors`` deliberately resolve differently: the collinearity
    diagnostics take pre-intercept columns, and handing them a design matrix
    would put a constant column into a correlation matrix.
    """
    return {
        "x": "X",
        "predictors": "predictors",
        "y": "y",
        "include": "Regression_Sample_Include",
        "has_intercept": "TRUE" if has_intercept else "FALSE",
        "alpha": str(_ALPHA) if func_name in _CI_FUNCS else None,
        "df_absorbed": None,
    }


def _let_bindings(k: int, has_intercept: bool, argument_names: tuple[str, ...]) -> str:
    keys = {name.lower() for name in argument_names}
    bindings = []
    if "x" in keys:
        bindings.append(f"X,{design_expression(k, has_intercept)}")
    if "predictors" in keys:
        bindings.append(f"predictors,{predictor_expression(k)}")
    return ",".join(bindings)


def _calc_formula(k: int, has_intercept: bool, func_name: str) -> str:
    argument_names = catalog_argument_names(func_name)
    call = build_call(func_name, argument_names, _reference_map(k, has_intercept, func_name))
    bindings = _let_bindings(k, has_intercept, argument_names)
    return f"=LET({bindings},{call})" if bindings else f"={call}"


def _calc_k_formula(k: int, has_intercept: bool, func_name: str, needs_y: bool) -> str:
    # needs_y is retained for the caller's column layout; the argument list
    # itself now comes from the catalog.
    del needs_y
    return _calc_formula(k, has_intercept, func_name)


def build_mlr_vector_row_configs(
    csv_path: Path,
) -> list[tuple[int, bool, RegressionVectors]]:
    """Build the per-section configs for the ``MLR_Vector_Outputs_Test`` sheet.

    For each ``k`` in ``_MLR_K_VALUES`` and each intercept setting, computes
    vector regression outputs against the Life Expectancy CSV
    (``FEATURE_COLUMNS[:k]``) and returns ``(k, has_intercept, vectors)``
    tuples consumed by :func:`write_mlr_vector_outputs_test_sheet`.
    """
    from .analyze_life_expectancy import calculate_regression_vectors

    row_configs: list[tuple[int, bool, RegressionVectors]] = []
    for k in _MLR_K_VALUES:
        for has_intercept in (True, False):
            vectors = calculate_regression_vectors(
                input_csv_path=csv_path,
                include_intercept=has_intercept,
                feature_columns=FEATURE_COLUMNS[:k],
            )
            row_configs.append((k, has_intercept, vectors))
    return row_configs


def _write_section(
    sheet: xw.Sheet,
    k: int,
    has_intercept: bool,
    vectors: RegressionVectors,
    first_data_row: int,
) -> None:
    for term_idx, term_name in enumerate(vectors.term_names):
        sheet.range((first_data_row + term_idx, _TERM_COL)).value = term_name
    for stat_idx, (_, func_name, _, num_fmt) in enumerate(_STATS):
        calc_col = _CALC_START_COL + stat_idx
        sheet.range((first_data_row, calc_col)).api.Formula2 = (
            _calc_formula(k, has_intercept, func_name)
        )
        sheet.range((first_data_row, calc_col)).api.NumberFormat = num_fmt
    # k-only stats return k values (no intercept row), so start one row lower when
    # the model includes an intercept to align with the first predictor row.
    k_row = first_data_row + (1 if has_intercept else 0)
    for k_idx, (_, func_name, num_fmt, needs_y) in enumerate(_K_STATS):
        calc_col = _CALC_START_COL + len(_STATS) + k_idx
        sheet.range((k_row, calc_col)).api.Formula2 = (
            _calc_k_formula(k, has_intercept, func_name, needs_y)
        )
        sheet.range((k_row, calc_col)).api.NumberFormat = num_fmt


def write_mlr_vector_outputs_test_sheet(
    workbook: xw.Book,
    row_configs: list[tuple[int, bool, RegressionVectors]],
) -> None:
    """Write the ``MLR_Vector_Outputs_Test`` sheet into ``workbook``.

    Lays out one section per ``(k, has_intercept)`` config in ``row_configs``;
    each section lists the regression terms as rows and writes the vector-output
    LAMBDA functions as Calc columns, with the expected vectors from
    ``row_configs`` available to the QC verification pass.
    """
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

    for k, has_intercept, vectors in row_configs:
        n_terms = len(vectors.coefficients)
        sheet.range((current_row, _TERM_COL)).value = (
            f"k={k} | Has_Intercept={'TRUE' if has_intercept else 'FALSE'}"
        )
        current_row += 1
        _write_section(sheet, k, has_intercept, vectors, current_row)
        current_row += n_terms + 1

    last_content_row = current_row - 1
    sheet.range((header_row, 1), (header_row, _TOTAL_COLS)).api.WrapText = True
    sheet.range((header_row, 1), (header_row, _TOTAL_COLS)).api.EntireRow.AutoFit()
    sheet.range((header_row, 1), (last_content_row, _TOTAL_COLS)).columns.autofit()

    safe_activate(sheet)
    safe_freeze_top_row(sheet)

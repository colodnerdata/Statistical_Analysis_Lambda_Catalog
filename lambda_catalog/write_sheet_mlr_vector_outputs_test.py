"""Build and write the MLR_Vector_Outputs_Test worksheet for regression vector formula smoke tests."""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

from .analyze_life_expectancy import (
    FEATURE_COLUMNS,
    RegressionVectors,
    calculate_regression_vectors,
)
from .make_test_sheet import _round_precision_for_format


_MLR_K_VALUES: list[int] = [1, 5, 10, 18]
_FIXED_COLS = 2  # Term + Smoke Test

# (function_name, number_format)
_STATS: list[tuple[str, str]] = [
    ("Coefficients", "0.000000"),
    ("SE_Coefficients", "0.000000"),
    ("T_Stats", "0.000"),
    ("P_Values", "0.00E+00"),
    ("CI_Lower", "0.000000"),
    ("CI_Upper", "0.000000"),
    ("CI_Excludes_Zero", "General"),
]


def _col_letter(col_idx: int) -> str:
    """Convert a 1-based column index to an Excel column letter string.

    Parameters
    ----------
    col_idx : int
        1-based column index (1 → 'A', 26 → 'Z', 27 → 'AA', …).

    Returns
    -------
    str
        Excel column letter string.
    """
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _delete_sheet_scoped_name_if_present(sheet: xw.Sheet, target_name: str) -> None:
    """Delete a local worksheet name if it already exists.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet whose Names collection will be searched.
    target_name : str
        The local (unqualified) name to remove, compared case-insensitively.
    """
    for index in range(sheet.api.Names.Count, 0, -1):
        existing_name = sheet.api.Names(index).Name
        local_name = existing_name.split("!", 1)[-1]
        if local_name.lower() == target_name.lower():
            sheet.api.Names(index).Delete()


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    """Create worksheet-scoped helper names used by the MLR vector formula tests.

    Parameters
    ----------
    sheet : xw.Sheet
        The MLR vector test worksheet to define names on.
    """
    for name, refers_to in (
        ("y", "=LifeExpectancyData[Life expectancy]"),
        ("fil", "=LifeExpectancyData[Full_Data]"),
    ):
        _delete_sheet_scoped_name_if_present(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)


def _calc_formula(k: int, allow_intercept: bool, func_name: str) -> str:
    """Build the spill formula for a vector stat column in a given test section.

    Parameters
    ----------
    k : int
        Number of predictor columns used in this section.
    allow_intercept : bool
        Whether the model includes an intercept.
    func_name : str
        Name of the workbook-scoped LAMBDA to call.

    Returns
    -------
    str
        Excel formula string that calls func_name via a LET-bound x_s.
    """
    allow_arg = "TRUE" if allow_intercept else "FALSE"
    return f"=LET(x_s,OFFSET(y,0,1,ROWS(y),{k}),{func_name}(x_s,y,{allow_arg},fil))"


def _match_formula_numeric(
    exp_col: str, calc_col: str, first_row: int, n_terms: int, num_fmt: str
) -> str:
    """Build the AND(ROUND…=ROUND…) match formula for a numeric stat column.

    Parameters
    ----------
    exp_col : str
        Excel column letter for the expected-value column.
    calc_col : str
        Excel column letter for the calculated-value (spill) column.
    first_row : int
        1-based Excel row of the first data row in this section.
    n_terms : int
        Number of terms (rows) in this section.
    num_fmt : str
        Excel number format string used to derive comparison precision.

    Returns
    -------
    str
        Excel formula string evaluating to TRUE when the two ranges match
        within the appropriate ROUND precision.
    """
    prec = _round_precision_for_format(num_fmt)
    last_row = first_row + n_terms - 1
    return (
        f"=AND(ROUND({exp_col}{first_row}:{exp_col}{last_row},{prec})"
        f"=ROUND({calc_col}{first_row}#,{prec}))"
    )


def _match_formula_bool(
    exp_col: str, calc_col: str, first_row: int, n_terms: int
) -> str:
    """Build the AND(…=…) match formula for a boolean stat column.

    Parameters
    ----------
    exp_col : str
        Excel column letter for the expected-value column.
    calc_col : str
        Excel column letter for the calculated-value (spill) column.
    first_row : int
        1-based Excel row of the first data row in this section.
    n_terms : int
        Number of terms (rows) in this section.

    Returns
    -------
    str
        Excel formula string evaluating to TRUE when every element of the
        expected range equals the corresponding spill value.
    """
    last_row = first_row + n_terms - 1
    return f"=AND({exp_col}{first_row}:{exp_col}{last_row}={calc_col}{first_row}#)"


def _smoke_test_formula(match_col_indices: list[int], first_data_row: int) -> str:
    """Build the AND(TRUE, …) smoke-test formula referencing each match column cell.

    Mirrors ``build_smoke_test_formula`` from make_test_sheet but uses absolute
    cell addresses instead of structured references since there is no ListObject.

    Parameters
    ----------
    match_col_indices : list[int]
        1-based column indices of every Match column.
    first_data_row : int
        1-based Excel row of the first data row in this section.

    Returns
    -------
    str
        Excel formula string that evaluates to TRUE when all match cells are TRUE.
    """
    refs = [f"{_col_letter(c)}{first_data_row}" for c in match_col_indices]
    return "=AND(TRUE," + ",".join(refs) + ")"


def build_mlr_vector_row_configs(
    csv_path: Path,
) -> list[tuple[int, bool, RegressionVectors]]:
    """Build the per-section vector configurations for all MLR k values.

    Parameters
    ----------
    csv_path : Path
        Path to the Life Expectancy CSV file used to compute expected values.

    Returns
    -------
    list[tuple[int, bool, RegressionVectors]]
        Two configs per k value (intercept=True then False), ordered by
        ascending k. Each tuple is (k, allow_intercept, vectors).
    """
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


def write_mlr_vector_outputs_test_sheet(
    workbook: xw.Book,
    row_configs: list[tuple[int, bool, RegressionVectors]],
) -> None:
    """Create or refresh the MLR_Vector_Outputs_Test sheet with vector comparison sections.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    row_configs : list[tuple[int, bool, RegressionVectors]]
        Per-section configurations produced by ``build_mlr_vector_row_configs``.
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
    _set_sheet_scoped_names(sheet)

    total_cols = _FIXED_COLS + len(_STATS) * 3  # 2 + 7×3 = 23

    # 1-based column indices of every Match column: 5, 8, 11, 14, 17, 20, 23
    match_col_indices = [_FIXED_COLS + 3 + i * 3 for i in range(len(_STATS))]

    HEADER_ROW = 1
    current_row = HEADER_ROW

    # Column headers
    sheet.range((current_row, 1)).value = "Term"
    sheet.range((current_row, 2)).value = "Smoke Test"
    for i, (stat_name, _) in enumerate(_STATS):
        base_col = _FIXED_COLS + 1 + i * 3
        sheet.range((current_row, base_col)).value = f"{stat_name}\n(Exp.)"
        sheet.range((current_row, base_col + 1)).value = f"{stat_name}\n(Calc.)"
        sheet.range((current_row, base_col + 2)).value = f"{stat_name}\n(Match)"
    current_row += 1

    for k, allow_intercept, vectors in row_configs:
        n_terms = len(vectors.coefficients)

        # Section divider row
        sheet.range((current_row, 1)).value = (
            f"k={k} | Intercept={'TRUE' if allow_intercept else 'FALSE'}"
        )
        current_row += 1

        first_data_row = current_row

        # Term labels — every data row
        for term_idx, term_name in enumerate(vectors.term_names):
            sheet.range((first_data_row + term_idx, 1)).value = term_name

        # Expected values — every data row
        exp_vectors: list[tuple] = [
            vectors.coefficients,
            vectors.std_errors,
            vectors.t_stats,
            vectors.p_values,
            vectors.ci_lower,
            vectors.ci_upper,
            vectors.ci_excludes_zero,
        ]
        for stat_idx, (_, num_fmt) in enumerate(_STATS):
            exp_col_idx = _FIXED_COLS + 1 + stat_idx * 3
            for term_idx, val in enumerate(exp_vectors[stat_idx]):
                cell = sheet.range((first_data_row + term_idx, exp_col_idx))
                if num_fmt != "General":
                    cell.api.NumberFormat = num_fmt
                cell.value = val

        # Calc. spill formulas — first data row only
        # .api.Formula2 is required for Excel 365 dynamic-array spill behavior;
        # .formula uses the legacy engine and collapses the result to a scalar.
        for stat_idx, (func_name, _) in enumerate(_STATS):
            calc_col_idx = _FIXED_COLS + 2 + stat_idx * 3
            sheet.range((first_data_row, calc_col_idx)).api.Formula2 = _calc_formula(
                k, allow_intercept, func_name
            )

        # Match formulas — first data row only
        # Also needs Formula2: these reference the spill range via `#` and compare
        # a multi-row expected range against the spilled calc range.
        for stat_idx, (stat_name, num_fmt) in enumerate(_STATS):
            exp_col_idx = _FIXED_COLS + 1 + stat_idx * 3
            calc_col_idx = _FIXED_COLS + 2 + stat_idx * 3
            match_col_idx = _FIXED_COLS + 3 + stat_idx * 3
            exp_col = _col_letter(exp_col_idx)
            calc_col = _col_letter(calc_col_idx)

            if stat_name == "CI_Excludes_Zero":
                mf = _match_formula_bool(exp_col, calc_col, first_data_row, n_terms)
            else:
                mf = _match_formula_numeric(
                    exp_col, calc_col, first_data_row, n_terms, num_fmt
                )
            sheet.range((first_data_row, match_col_idx)).api.Formula2 = mf

        # Smoke Test formula — first data row, col 2
        sheet.range((first_data_row, 2)).api.Formula2 = _smoke_test_formula(
            match_col_indices, first_data_row
        )

        current_row += n_terms + 1  # +1 for blank gap row

    last_content_row = current_row - 1

    # Wrap and autofit header row
    header_range = sheet.range((HEADER_ROW, 1), (HEADER_ROW, total_cols))
    header_range.api.WrapText = True
    header_range.api.EntireRow.AutoFit()

    # Autofit all columns
    sheet.range((HEADER_ROW, 1), (last_content_row, total_cols)).columns.autofit()

    # Freeze header row
    sheet.activate()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

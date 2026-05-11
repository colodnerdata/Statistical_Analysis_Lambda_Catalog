"""Build and write the MLR_Vector_Outputs_Test worksheet for regression vector formula smoke tests."""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

from .analyze_life_expectancy import (
    FEATURE_COLUMNS,
    RegressionVectors,
    calculate_regression_vectors,
)


_MLR_K_VALUES: list[int] = [1, 5, 10, 18]

_ALPHA = 0.05

_TESTS_COLS = 7    # col 1: Smoke Test; cols 2–7: one Match per stat
_FORMULA_COLS = 7  # col 8: Term names; cols 9–14: individual function calls
_EXPECTED_COLS = 6 # cols 15–20: expected values
_TOTAL_COLS = _TESTS_COLS + _FORMULA_COLS + _EXPECTED_COLS  # 20

_TERM_COL = _TESTS_COLS + 1                            # col 8
_CALC_START_COL = _TESTS_COLS + 2                      # col 9
_EXPECTED_START_COL = _TESTS_COLS + _FORMULA_COLS + 1  # col 15

_CI_FUNCS = frozenset({"CI_Lower", "CI_Upper"})

# (display name, Excel function name, display_digits, number format)
# display_digits: decimal places shown in cells; match precision = 2 × display_digits.
_D = 3  # display decimal digits for all numeric vector stats
_STATS: list[tuple[str, str, int, str]] = [
    ("Coefficients",    "Coefficients",    _D, f"0.{'0' * _D}"),
    ("SE_Coefficients", "SE_Coefficients", _D, f"0.{'0' * _D}"),
    ("T_Stats",         "T_Stats",         _D, f"0.{'0' * _D}"),
    ("P_Values",        "P_Values",        _D, f"0.{'0' * _D}E+00"),
    ("CI_Lower",        "CI_Lower",        _D, f"0.{'0' * _D}"),
    ("CI_Upper",        "CI_Upper",        _D, f"0.{'0' * _D}"),
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
    """Build an individual stat function call for a given k and intercept setting.

    CI functions receive the module-level significance level ``_ALPHA`` as an extra arg.

    Parameters
    ----------
    k : int
        Number of predictor columns.
    allow_intercept : bool
        Whether the model includes an intercept.
    func_name : str
        Excel function name (e.g. ``'Coefficients'``, ``'CI_Lower'``).

    Returns
    -------
    str
        Excel formula string starting with ``=``.
    """
    allow_arg = "TRUE" if allow_intercept else "FALSE"
    extra = f",{_ALPHA}" if func_name in _CI_FUNCS else ""
    return f"=LET(x_s,OFFSET(y,0,1,ROWS(y),{k}),{func_name}(x_s,y,{allow_arg},fil{extra}))"


def _match_formula_numeric(  # noqa: PLR0913
    exp_col_idx: int,
    calc_col_idx: int,
    first_row: int,
    n_terms: int,
    prec: int,
) -> str:
    """Build a ROUND-based match formula comparing an expected range to a spill column.

    Parameters
    ----------
    exp_col_idx : int
        1-based column index of the expected values.
    calc_col_idx : int
        1-based column index of the calc (spill anchor) column.
    first_row : int
        1-based Excel row of the first data row.
    n_terms : int
        Number of coefficient terms.
    prec : int
        ROUND precision derived from the number format.

    Returns
    -------
    str
        Excel formula evaluating to TRUE when the ranges match within precision.
    """
    exp_col = _col_letter(exp_col_idx)
    calc_col = _col_letter(calc_col_idx)
    last_row = first_row + n_terms - 1
    return (
        f"=AND(ROUND({exp_col}{first_row}:{exp_col}{last_row},{prec})"
        f"=ROUND({calc_col}{first_row}#,{prec}))"
    )


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


def _write_section(  # noqa: PLR0914
    sheet: xw.Sheet,
    k: int,
    allow_intercept: bool,
    vectors: RegressionVectors,
    first_data_row: int,
) -> None:
    """Write one section: term names, expected values, calc formulas, match formulas, smoke test.

    Parameters
    ----------
    sheet : xw.Sheet
        The target worksheet.
    k : int
        Number of predictor columns for this section.
    allow_intercept : bool
        Whether this section's model includes an intercept.
    vectors : RegressionVectors
        Computed regression statistics used as expected values.
    first_data_row : int
        1-based Excel row for the first data row of this section.
    """
    n_terms = len(vectors.coefficients)

    for term_idx, term_name in enumerate(vectors.term_names):
        sheet.range((first_data_row + term_idx, _TERM_COL)).value = term_name

    exp_vectors: list[tuple] = [
        vectors.coefficients,
        vectors.std_errors,
        vectors.t_stats,
        vectors.p_values,
        vectors.ci_lower,
        vectors.ci_upper,
    ]
    for stat_idx, (_, func_name, display_digits, num_fmt) in enumerate(_STATS):
        exp_col_idx = _EXPECTED_START_COL + stat_idx
        calc_col = _CALC_START_COL + stat_idx
        match_col = 2 + stat_idx

        for term_idx, val in enumerate(exp_vectors[stat_idx]):
            cell = sheet.range((first_data_row + term_idx, exp_col_idx))
            if num_fmt != "General":
                cell.api.NumberFormat = num_fmt
            cell.value = val

        sheet.range((first_data_row, calc_col)).api.Formula2 = (
            _calc_formula(k, allow_intercept, func_name)
        )

        mf = _match_formula_numeric(
            exp_col_idx, calc_col, first_data_row, n_terms, display_digits * 2,
        )
        sheet.range((first_data_row, match_col)).api.Formula2 = mf

    match_refs = ",".join(
        f"{_col_letter(2 + i)}{first_data_row}" for i in range(len(_STATS))
    )
    sheet.range((first_data_row, 1)).formula = f"=AND({match_refs})"


def write_mlr_vector_outputs_test_sheet(
    workbook: xw.Book,
    row_configs: list[tuple[int, bool, RegressionVectors]],
) -> None:
    """Create or refresh the MLR_Vector_Outputs_Test sheet.

    Layout (20 columns total):
      Cols 1–7   Tests: Smoke Test + one Match column per stat
      Col  8     Term names (Python-written)
      Cols 9–14  Individual function call per stat (each spills n_terms rows)
      Cols 15–20 Expected values

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

    header_row = 1
    sheet.range((header_row, 1)).value = "Smoke Test"
    for i, (display_name, _, _, _) in enumerate(_STATS):
        sheet.range((header_row, 2 + i)).value = f"{display_name}\n(Match)"
    sheet.range((header_row, _TERM_COL)).value = "Term"
    for i, (display_name, _, _, _) in enumerate(_STATS):
        sheet.range((header_row, _CALC_START_COL + i)).value = display_name
    for i, (display_name, _, _, _) in enumerate(_STATS):
        sheet.range((header_row, _EXPECTED_START_COL + i)).value = f"{display_name}\n(Exp.)"

    current_row = header_row + 1

    for k, allow_intercept, vectors in row_configs:
        n_terms = len(vectors.coefficients)
        sheet.range((current_row, 1)).value = (
            f"k={k} | Intercept={'TRUE' if allow_intercept else 'FALSE'}"
        )
        current_row += 1
        _write_section(sheet, k, allow_intercept, vectors, current_row)
        current_row += n_terms + 1  # data rows + blank gap

    last_content_row = current_row - 1
    sheet.range((header_row, 1), (header_row, _TOTAL_COLS)).api.WrapText = True
    sheet.range((header_row, 1), (header_row, _TOTAL_COLS)).api.EntireRow.AutoFit()
    sheet.range((header_row, 1), (last_content_row, _TOTAL_COLS)).columns.autofit()

    sheet.activate()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

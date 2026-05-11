"""Build and write the MLR_Scalar_Test worksheet for regression formula smoke tests."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

import xlwings as xw

from .analyze_life_expectancy import (
    FEATURE_COLUMNS,
    RegressionSummary,
    calculate_regression_summary,
)
from .make_test_sheet import (
    _ColumnSpec,
    _RowConfig,
    _match_column_formula,
    build_smoke_test_formula,
    write_test_table,
)


_MLR_K_VALUES: list[int] = [1, 5, 10, 18]
_MLR_TABLE_HEADER_ROW = 1

# Replaces x_s in every (Calc.) formula; resolves dynamically from each
# row's ind_vars value.
_MLR_X_S_OFFSET = "OFFSET(y,0,1,ROWS(y),[@[ind_vars]])"


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
    """Create worksheet-scoped helper names used by the MLR formula tests.

    Parameters
    ----------
    sheet : xw.Sheet
        The MLR test worksheet to define names on.
    """
    local_names = {
        "y": "=LifeExpectancyData[Life expectancy]",
        "fil": "=LifeExpectancyData[Full_Data]",
    }

    # Remove legacy x_s names from older workbook versions
    for legacy in ("x_s", "x_s_k1", "x_s_k5", "x_s_k10", "x_s_k18"):
        _delete_sheet_scoped_name_if_present(sheet, legacy)

    for name, refers_to in local_names.items():
        _delete_sheet_scoped_name_if_present(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)

    _delete_sheet_scoped_name_if_present(sheet, "Allow_Intercept")


def _actual_formula(
    function_name: str,
    argument_names: Iterable[str],
    x_s_name: str = "x_s",
) -> str:
    """Build an executable test formula that references local helper names.

    When ``x_s_name`` is not the literal ``'x_s'`` and the function takes an
    ``x_s`` argument, the formula is wrapped in
    ``LET(x_s, <x_s_name>, ...)`` so the LAMBDA is called with a plain
    ``'x_s'`` reference rather than an inlined expression.

    Parameters
    ----------
    function_name : str
        The Excel defined name of the LAMBDA to call.
    argument_names : Iterable[str]
        Ordered argument names accepted by the function.
    x_s_name : str, optional
        Expression to substitute for the ``x_s`` argument. Defaults to
        the literal name ``'x_s'``.

    Returns
    -------
    str
        An Excel formula string that calls the function with the
        appropriate helper-name references.
    """
    args = list(argument_names)
    uses_x_s = any(a.strip().lower() == "x_s" for a in args)
    use_let = uses_x_s and x_s_name != "x_s"

    reference_map = {
        "y": "y",
        "x_s": "x_s" if use_let else x_s_name,
        "filter": "fil",
        "allow_intercept": "[@[Allow_Intercept]]",
    }

    resolved_arguments = [
        reference_map.get(argument_name.strip().lower(), argument_name)
        for argument_name in args
    ]
    call = f"{function_name}({', '.join(resolved_arguments)})"
    return f"=LET(x_s, {x_s_name}, {call})" if use_let else f"={call}"


def _expected_values_map(summary: RegressionSummary) -> dict[str, float | int]:
    """Map workbook function names to expected values from a regression summary.

    Parameters
    ----------
    summary : RegressionSummary
        Regression metrics computed by ``calculate_regression_summary``.

    Returns
    -------
    dict[str, float | int]
        Mapping from Excel defined-name strings to their expected scalar
        values, keyed to match the ``(Exp.)`` column headers in the test
        table.
    """
    return {
        "Observations": summary.observations,
        "DF_Regression": summary.df_regression,
        "DF_Total": summary.df_total,
        "R_squared": summary.r_squared,
        "DF_Residual": summary.df_residual,
        "Multiple_R": summary.multiple_r,
        "Adjusted_R2": summary.adjusted_r2,
        "SS_Total": summary.ss_total,
        "SS_Residual": summary.ss_residual,
        "SS_Regression": summary.ss_regression,
        "SE_Regression": summary.se_regression,
    }


def build_test_columns(
    test_table: str,
    definitions: list,
) -> list[_ColumnSpec]:
    """Build the full column spec list for a test table from matching definitions.

    Parameters
    ----------
    test_table : str
        Tag value used to filter definitions by their ``test_table`` attribute.
    definitions : list
        Sequence of LambdaDefinition-like objects with ``test_table``,
        ``name``, ``argument_names``, and ``number_format`` attributes.

    Returns
    -------
    list[_ColumnSpec]
        Ordered column specs: fixed control columns (X_Variables, ind_vars,
        Allow_Intercept, Smoke Test), then one (Match) column per definition,
        then one (Calc.) column per definition, then one (Exp.) column per
        definition.
    """
    filtered = [d for d in definitions if d.test_table == test_table]
    exp_columns: list[_ColumnSpec] = []
    calc_columns: list[_ColumnSpec] = []
    match_columns: list[_ColumnSpec] = []
    match_headers: list[str] = []
    for d in filtered:
        exp_header = f"{d.name}\n(Exp.)"
        calc_header = f"{d.name}\n(Calc.)"
        match_header = f"{d.name}\n(Match)"
        exp_columns.append((exp_header, d.name, None, d.number_format))
        calc_columns.append(
            (
                calc_header,
                None,
                _actual_formula(d.name, d.argument_names, _MLR_X_S_OFFSET),
                d.number_format,
            )
        )
        match_columns.append(
            (
                match_header,
                None,
                _match_column_formula(exp_header, calc_header, d.number_format),
                "General",
            )
        )
        match_headers.append(match_header)

    smoke_formula = build_smoke_test_formula(test_table, match_headers)
    fixed_columns: list[_ColumnSpec] = [
        (
            "X_Variables",
            None,
            "=TEXTJOIN(\", \",TRUE,OFFSET(y,-1,1,1,[@[ind_vars]]))",
            "General",
        ),
        ("ind_vars", None, None, "0"),
        ("Allow_Intercept", None, None, "General"),
        ("Smoke Test", None, smoke_formula, "General"),
    ]
    return fixed_columns + match_columns + calc_columns + exp_columns


def build_mlr_row_configs(csv_path: Path) -> list[_RowConfig]:
    """Build the per-row test configurations for all MLR k values.

    Parameters
    ----------
    csv_path : Path
        Path to the Life Expectancy CSV file used to compute expected values.

    Returns
    -------
    list[_RowConfig]
        Two row configs per k value (intercept=True and intercept=False),
        ordered by ascending k.
    """
    row_configs: list[_RowConfig] = []
    for k in _MLR_K_VALUES:
        feature_cols = FEATURE_COLUMNS[:k]
        summary_true = calculate_regression_summary(
            input_csv_path=csv_path,
            include_intercept=True,
            feature_columns=feature_cols,
        )
        summary_false = calculate_regression_summary(
            input_csv_path=csv_path,
            include_intercept=False,
            feature_columns=feature_cols,
        )
        row_configs.append(
            ({"ind_vars": k, "Allow_Intercept": True}, _expected_values_map(summary_true))
        )
        row_configs.append(
            ({"ind_vars": k, "Allow_Intercept": False}, _expected_values_map(summary_false))
        )
    return row_configs


def write_mlr_scalar_test_sheet(
    workbook: xw.Book,
    definitions: list,
    row_configs: list[_RowConfig],
) -> None:
    """Create or refresh the MLR_Scalar_Test sheet with a metric comparison table.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    definitions : list
        Sequence of LambdaDefinition-like objects used to build column specs.
    row_configs : list[_RowConfig]
        Per-row configurations produced by ``build_mlr_row_configs``.
    """
    test_table = "MLR_Scalar_Test"

    sheet_names = {sheet.name for sheet in workbook.sheets}
    if test_table in sheet_names:
        sheet = workbook.sheets[test_table]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = test_table
    else:
        sheet = workbook.sheets.add(name=test_table, after=workbook.sheets[-1])

    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()

    _set_sheet_scoped_names(sheet)

    columns = build_test_columns(test_table, definitions)
    header_row = _MLR_TABLE_HEADER_ROW
    last_data_row = header_row + len(row_configs)
    col_count = len(columns)

    rows_data: list[tuple[int, dict, dict]] = [
        (header_row + 1 + i, row_vals, expected)
        for i, (row_vals, expected) in enumerate(row_configs)
    ]

    write_test_table(sheet, test_table, columns, header_row, rows_data)

    # Column A: fixed width, wrap text, auto-height rows
    sheet.range("A1").column_width = 100
    sheet.range((header_row, 1), (last_data_row, 1)).api.WrapText = True
    sheet.range((header_row, 1), (last_data_row, col_count)).api.EntireRow.AutoFit()

    # Header row: wrap text so Alt-Enter column names display on two lines
    sheet.range((header_row, 1), (header_row, col_count)).api.WrapText = True
    sheet.range((header_row, 1), (header_row, col_count)).api.EntireRow.AutoFit()

    # All columns except A: autofit width
    sheet.range((header_row, 2), (last_data_row, col_count)).columns.autofit()
    sheet.activate()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

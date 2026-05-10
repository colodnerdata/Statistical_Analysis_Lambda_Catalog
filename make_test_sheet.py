"""Low-level helpers for creating and populating Excel test tables."""
from __future__ import annotations

from typing import Any

import xlwings as xw

from workbook_helpers import XL_SRC_RANGE, XL_YES


_SMOKE_TEST_ROUND_PRECISION = 10

# (header, expected_key_or_None, formula_or_None, number_format)
_ColumnSpec = tuple[str, str | None, str | None, str]

# Per-row config: (per-row cell values, expected-values dict)
_RowConfig = tuple[dict[str, Any], dict[str, float | int]]


def _round_precision_for_format(num_fmt: str) -> int:
    """Return the ROUND precision for smoke-test comparisons given a number format.

    The comparison precision intentionally exceeds the display precision:
    decimal places are doubled (e.g. ``'0.000'`` → 6, ``'0.00'`` → 4) so
    the test is stricter than what the cell shows while still tolerating
    unavoidable floating-point rounding beyond those digits. Falls back to
    ``_SMOKE_TEST_ROUND_PRECISION`` for formats with no decimal point.

    Parameters
    ----------
    num_fmt : str
        Excel number format string (e.g. ``'0.000'``, ``'General'``).

    Returns
    -------
    int
        Number of decimal places to pass to Excel's ROUND function.
    """
    dot = num_fmt.find(".")
    if dot == -1:
        return _SMOKE_TEST_ROUND_PRECISION
    return (len(num_fmt) - dot - 1) * 2


def _match_column_formula(exp_header: str, calc_header: str, num_fmt: str) -> str:
    """Return the boolean comparison formula for a per-metric match column.

    Integer formats (``'0'``) compare exactly. All other formats compare
    via ROUND at a precision that exceeds the display precision.

    Parameters
    ----------
    exp_header : str
        Column header for the expected-value column.
    calc_header : str
        Column header for the calculated-value column.
    num_fmt : str
        Excel number format string used to determine comparison precision.

    Returns
    -------
    str
        An Excel formula string that evaluates to TRUE when the two columns
        match within the appropriate precision.
    """
    if num_fmt == "0":
        return f"=[@[{exp_header}]]=[@[{calc_header}]]"
    precision = _round_precision_for_format(num_fmt)
    return (
        f"=ROUND([@[{exp_header}]],{precision})"
        f"=ROUND([@[{calc_header}]],{precision})"
    )


def build_smoke_test_formula(
    test_table: str,
    match_headers: list[str],
) -> str:
    """Build the ``=AND(...)`` smoke-test formula referencing each match column.

    Parameters
    ----------
    test_table : str
        Name of the test table (unused; reserved for future scoping).
    match_headers : list[str]
        Column headers for the per-metric match columns to AND together.

    Returns
    -------
    str
        An Excel formula string that evaluates to TRUE when all match
        columns are TRUE.
    """
    _ = test_table
    return "=AND(" + ",".join(f"[@[{h}]]" for h in match_headers) + ")"


def write_test_table(
    sheet: xw.Sheet,
    table_name: str,
    columns: list[_ColumnSpec],
    header_row: int,
    rows_data: list[tuple[int, dict[str, Any], dict[str, float | int]]],
) -> None:
    """Create and populate an Excel ListObject for a test table.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet on which to create the table.
    table_name : str
        Name to assign to the Excel ListObject.
    columns : list[_ColumnSpec]
        Column specifications as (header, expected_key, formula, number_format)
        4-tuples.
    header_row : int
        1-based row index for the table header row.
    rows_data : list[tuple[int, dict, dict]]
        Per-row data as (row_index, row_values, expected_values) 3-tuples.
    """
    last_data_row = header_row + len(rows_data)
    col_count = len(columns)

    for col_idx, (header, _, _, _) in enumerate(columns, start=1):
        sheet.range((header_row, col_idx)).value = header

    table_range = sheet.range((header_row, 1), (last_data_row, col_count))
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = table_name
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = False
    table.ShowTableStyleColumnStripes = False

    for data_row, row_values, expected_values in rows_data:
        is_first_row = data_row == header_row + 1
        for col_idx, (header, exp_key, formula, num_fmt) in enumerate(columns, start=1):
            cell = sheet.range((data_row, col_idx))
            cell.api.NumberFormat = num_fmt
            if header in row_values:
                cell.value = row_values[header]
            elif formula is not None:
                # Formulas auto-propagate within a table; only write on the first row
                if is_first_row:
                    cell.formula = formula
            elif exp_key is not None:
                val = expected_values.get(exp_key)
                if val is not None:
                    cell.value = val

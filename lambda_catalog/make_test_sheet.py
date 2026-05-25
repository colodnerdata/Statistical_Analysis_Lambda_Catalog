"""Low-level helpers for creating and populating Excel test tables."""
from __future__ import annotations

from typing import Any

import xlwings as xw

from .workbook_helpers import XL_SRC_RANGE, XL_YES


# (header, formula_or_None, number_format)
_ColumnSpec = tuple[str, str | None, str]

# Per-row config: (per-row cell values, expected-values dict)
# The expected-values dict is used for Python-side verification only; not written to the sheet.
_RowConfig = tuple[dict[str, Any], dict[str, float | int]]


def write_test_table(
    sheet: xw.Sheet,
    table_name: str,
    columns: list[_ColumnSpec],
    header_row: int,
    rows_data: list[tuple[int, dict[str, Any]]],
) -> None:
    """Create and populate an Excel ListObject for a test table.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet on which to create the table.
    table_name : str
        Name to assign to the Excel ListObject.
    columns : list[_ColumnSpec]
        Column specifications as (header, formula_or_None, number_format) 3-tuples.
    header_row : int
        1-based row index for the table header row.
    rows_data : list[tuple[int, dict]]
        Per-row data as (row_index, row_values) 2-tuples. Values whose key matches
        a column header are written directly; columns with a formula receive it on
        the first data row and let the table auto-propagate it.
    """
    last_data_row = header_row + len(rows_data)
    col_count = len(columns)

    for col_idx, (header, _, _) in enumerate(columns, start=1):
        sheet.range((header_row, col_idx)).value = header

    table_range = sheet.range((header_row, 1), (last_data_row, col_count))
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = table_name
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False

    for data_row, row_values in rows_data:
        is_first_row = data_row == header_row + 1
        for col_idx, (header, formula, num_fmt) in enumerate(columns, start=1):
            cell = sheet.range((data_row, col_idx))
            cell.api.NumberFormat = num_fmt
            if header in row_values:
                cell.value = row_values[header]
            elif formula is not None and is_first_row:
                cell.formula = formula

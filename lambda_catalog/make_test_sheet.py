"""Low-level helpers for creating and populating Excel test tables."""
from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any

import xlwings as xw

from .workbook_helpers import XL_SRC_RANGE, XL_YES

# (header, formula_or_None, number_format)
_ColumnSpec = tuple[str, str | None, str]

# Per-row config: (per-row cell values, expected-values dict)
# The expected-values dict is used for Python-side verification only; not written to the sheet.
_RowConfig = tuple[dict[str, Any], dict[str, float | int]]


def predictor_expression(k: int, response_ref: str = "y") -> str:
    """The bare predictor block for a QC section: k columns right of the response."""
    return f"OFFSET({response_ref},0,1,ROWS({response_ref}),{k})"


def design_expression(k: int, has_intercept: bool, response_ref: str = "y") -> str:
    """The design matrix for a QC section — predictors, intercept column first.

    From v3.0 the engines no longer synthesize an intercept from a flag; the
    caller supplies the column. Each QC section knows its intercept setting at
    build time, so this resolves to a literal rather than a worksheet ``IF``.
    """
    predictors = predictor_expression(k, response_ref)
    if not has_intercept:
        return predictors
    return f"HSTACK(SEQUENCE(ROWS({response_ref}),1,1,0),{predictors})"


def build_call(
    function_name: str,
    argument_names: Iterable[str],
    reference_map: Mapping[str, str | None],
) -> str:
    """Render a call to ``function_name`` by resolving its declared arguments.

    Every QC test sheet builds its formulas this way rather than hard-coding a
    positional argument list, so a signature change in ``lambda_functions.json``
    reaches the harness automatically. That matters most for the arguments the
    v3.0 intercept relocation introduced: a function declaring ``X`` wants the
    design matrix (intercept column included), one declaring ``Predictors``
    wants the bare predictor block, and the two must never be confused.

    Parameters
    ----------
    function_name : str
        The catalog function being called.
    argument_names : Iterable[str]
        The function's declared argument names, in order.
    reference_map : Mapping[str, str or None]
        Lower-cased argument name to the formula text to pass. A value of
        ``None`` truncates the argument list at that point, which is how
        trailing optional arguments are left at their defaults.

    Returns
    -------
    str
        The rendered call, without a leading ``=``.

    Raises
    ------
    KeyError
        If a declared argument has no entry in ``reference_map``. Failing loudly
        is the point: a silently dropped argument would shift every argument
        after it by one position.
    """
    rendered: list[str] = []
    for name in argument_names:
        key = name.strip().lower()
        if key not in reference_map:
            raise KeyError(
                f"{function_name}: no reference for argument {name!r} "
                f"(known: {sorted(reference_map)})"
            )
        value = reference_map[key]
        if value is None:
            break
        rendered.append(value)
    return f"{function_name}({','.join(rendered)})"


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

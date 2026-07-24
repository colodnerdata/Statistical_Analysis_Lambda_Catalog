"""Write the bundled Auto MPG dataset into a workbook sheet/table."""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

from .workbook_helpers import (
    XL_SRC_RANGE,
    XL_YES,
    get_or_create_sheet,
    reset_generated_sheet,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_AUTO_MPG_XLSX_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.xlsx"
SHEET_NAME = "Auto_MPG"
TABLE_NAME = "Auto_MPG_Data"


def write_auto_mpg_sheet(
    workbook: xw.Book,
    source_workbook_path: Path = DEFAULT_AUTO_MPG_XLSX_PATH,
) -> None:
    """Copy the shipped Auto MPG sheet and recreate its table in ``workbook``."""
    source_workbook: xw.Book | None = None
    try:
        source_workbook = workbook.app.books.open(str(source_workbook_path.resolve()))
        source_sheet = source_workbook.sheets[SHEET_NAME]
        values = source_sheet.used_range.value
    finally:
        if source_workbook is not None:
            source_workbook.close()

    if values is None:
        raise ValueError(f"{source_workbook_path} has no used range on {SHEET_NAME!r}")

    if isinstance(values, list) and values and isinstance(values[0], list):
        rows = len(values)
        cols = len(values[0])
    elif isinstance(values, list):
        rows = 1
        cols = len(values)
    else:
        rows = 1
        cols = 1

    if rows < 2 or cols < 1:
        raise ValueError(f"{source_workbook_path} does not contain table-shaped Auto MPG data")

    target_sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(target_sheet)
    target_sheet.activate()
    target_sheet.range((1, 1), (rows, cols)).value = values

    table_range = target_sheet.range((1, 1), (rows, cols))
    table = target_sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = TABLE_NAME
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False

    target_sheet.used_range.columns.autofit()
    target_sheet.api.Application.ActiveWindow.SplitRow = 1
    target_sheet.api.Application.ActiveWindow.SplitColumn = 0
    target_sheet.api.Application.ActiveWindow.FreezePanes = True

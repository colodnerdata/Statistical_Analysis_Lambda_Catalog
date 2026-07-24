"""Write the shipped Auto MPG dataset sheet into a target workbook."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import xlwings as xw

from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    XL_SRC_RANGE,
    XL_YES,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_AUTO_MPG_XLSX_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.xlsx"
SHEET_NAME = "Auto_MPG"
TABLE_NAME = "Auto_MPG_Data"
_TABLE_STYLE = "TableStyleMedium2"


def _as_2d_rows(values: Any) -> list[list[Any]]:
    """Normalize xlwings UsedRange values into a rectangular 2D list."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [list(values)]
    return [list(row) for row in values]


def _load_source_rows(
    app: xw.App,
    source_workbook_path: Path,
) -> list[list[Any]]:
    """Read all visible data from the bundled Auto MPG source workbook."""
    source_workbook: xw.Book | None = None
    try:
        source_workbook = app.books.open(str(source_workbook_path))
        source_sheet = source_workbook.sheets[SHEET_NAME]
        rows = _as_2d_rows(source_sheet.used_range.value)
    finally:
        if source_workbook is not None:
            source_workbook.close()

    if len(rows) < 2 or len(rows[0]) == 0:
        raise ValueError(
            f"Auto MPG source workbook has no table-shaped data: {source_workbook_path}"
        )
    return rows


def write_auto_mpg_sheet(
    workbook: xw.Book,
    source_workbook_path: Path = DEFAULT_AUTO_MPG_XLSX_PATH,
) -> int:
    """Write Auto MPG headers/data and recreate the Auto_MPG_Data table."""
    rows = _load_source_rows(workbook.app, source_workbook_path.resolve())
    row_count = len(rows)
    col_count = len(rows[0])

    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)
    sheet.activate()
    sheet.range((1, 1), (row_count, col_count)).value = rows

    table_range = sheet.range((1, 1), (row_count, col_count))
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = TABLE_NAME
    table.TableStyle = _TABLE_STYLE
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False

    sheet.used_range.columns.autofit()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True
    return max(row_count - 1, 0)


def write_auto_mpg_data(
    workbook_path: Path,
    source_workbook_path: Path = DEFAULT_AUTO_MPG_XLSX_PATH,
) -> int:
    """Write the Auto MPG sheet to a workbook and return number of data rows."""
    workbook_path = workbook_path.resolve()
    source_workbook_path = source_workbook_path.resolve()
    row_count = 0

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook, workbook_exists = open_or_create_workbook(app, workbook_path)
            try:
                row_count = write_auto_mpg_sheet(
                    workbook,
                    source_workbook_path=source_workbook_path,
                )
                if workbook_exists:
                    workbook.save()
                else:
                    workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "write Auto MPG data in", exc)

    return row_count


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the Auto MPG sheet writer."""
    parser = argparse.ArgumentParser(
        description=(
            "Write auto_mpg_data.xlsx into a workbook sheet named "
            f"'{SHEET_NAME}' with table '{TABLE_NAME}'."
        )
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the workbook file to create or update.",
    )
    parser.add_argument(
        "--source-workbook",
        type=Path,
        default=DEFAULT_AUTO_MPG_XLSX_PATH,
        help="Path to the source auto_mpg_data.xlsx workbook.",
    )
    return parser.parse_args()


def main() -> None:
    """Write the Auto MPG sheet and print a short summary."""
    args = parse_args()
    row_count = write_auto_mpg_data(
        workbook_path=args.workbook,
        source_workbook_path=args.source_workbook,
    )
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet: {SHEET_NAME}")
    print(f"Table: {TABLE_NAME}")
    print(f"Rows written: {row_count}")


if __name__ == "__main__":
    main()

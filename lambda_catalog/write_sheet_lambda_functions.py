"""Write the LAMBDA function catalog to the LAMBDA_functions worksheet."""
from __future__ import annotations

import argparse
from collections.abc import Sequence
from pathlib import Path

import xlwings as xw

from .catalog_schema import CatalogFunction, load_catalog_document
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    XL_SRC_RANGE,
    XL_YES,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
    safe_activate,
    safe_freeze_top_row,
)

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
SHEET_NAME = "LAMBDA_functions"
TABLE_NAME = "LAMBDAFunctionsCatalog"
TABLE_HEADERS = [
    "Function Name",
    "Scope",
    "Definition",
    "Arguments",
    "Yields",
    "Plain-Language Summary",
    "Description",
]
COLUMN_WIDTHS = {"A": 18, "B": 18, "C": 72, "D": 30, "E": 15, "F": 34, "G": 66}
_LAST_COLUMN = "G"


def _scope_label(scope: str) -> str:
    """Render a function's scope for the catalog: 'Workbook' or the sheet name.

    Workbook-scoped functions are portable defined names; a sheet name means
    the function is a sheet-scoped closure that only resolves on that sheet
    (e.g. the Model Construction constructors), so browsers are not misled
    into calling it from elsewhere.
    """
    return "Workbook" if scope == "workbook" else scope


def write_catalog_sheet(workbook: xw.Book, entries: Sequence[CatalogFunction]) -> None:
    """Write catalog entries to the LAMBDA_functions sheet as a formatted table.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    entries : Sequence[CatalogFunction]
        Catalog functions to populate the table rows.
    """
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)
    safe_activate(sheet)

    for column_index, header in enumerate(TABLE_HEADERS, start=1):
        sheet.range((1, column_index)).value = header

    last_data_row = len(entries) + 1
    if entries:
        # Definition (column C) is forced to text so long LAMBDA strings are
        # never coerced or truncated by Excel's numeric parsing.
        sheet.range(f"C2:C{last_data_row}").api.NumberFormat = "@"

    for row_offset, entry in enumerate(entries, start=2):
        sheet.range((row_offset, 1)).value = entry.name
        sheet.range((row_offset, 2)).value = _scope_label(entry.scope)
        sheet.range((row_offset, 3)).value = entry.formula_display
        sheet.range((row_offset, 4)).value = entry.arguments_cell_text()
        sheet.range((row_offset, 5)).value = entry.yields
        sheet.range((row_offset, 6)).value = entry.plain_language_summary
        sheet.range((row_offset, 7)).value = entry.description

    table_range = sheet.range(f"A1:{_LAST_COLUMN}{last_data_row}")
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = TABLE_NAME
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False

    for col_letter, width in COLUMN_WIDTHS.items():
        sheet.range(f"{col_letter}:{col_letter}").column_width = width

    sheet.range(f"A1:{_LAST_COLUMN}{last_data_row}").api.WrapText = True
    sheet.range(f"A2:{_LAST_COLUMN}{last_data_row}").api.EntireRow.AutoFit()

    safe_freeze_top_row(sheet)


def write_lambda_catalog(
    workbook_path: Path,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
) -> int:
    """Write the catalog sheet to the workbook and return the number of entries written.

    Parameters
    ----------
    workbook_path : Path
        Path to the workbook file to create or update.
    definitions_path : Path, optional
        Path to the JSON file containing lambda catalog metadata.

    Returns
    -------
    int
        Number of catalog entries written to the sheet.
    """
    document = load_catalog_document(definitions_path)
    entries = document.functions
    workbook_path = workbook_path.resolve()

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook, workbook_exists = open_or_create_workbook(app, workbook_path)
            try:
                write_catalog_sheet(workbook, entries)
                if workbook_exists:
                    workbook.save()
                else:
                    workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "write catalog sheet in", exc)

    return len(entries)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the catalog sheet writer.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with workbook and definitions attributes.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Write the human-readable LAMBDA catalog into a workbook sheet "
            "named LAMBDA_functions."
        )
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the workbook file to create or update.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda catalog metadata.",
    )
    return parser.parse_args()


def main() -> None:
    """Write the catalog sheet and print a summary for interactive use."""
    args = parse_args()
    rows_written = write_lambda_catalog(
        workbook_path=args.workbook, definitions_path=args.definitions
    )
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet: {SHEET_NAME}")
    print(f"Rows written: {rows_written}")


if __name__ == "__main__":
    main()

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pywintypes  # type: ignore[import-untyped]
import xlwings as xw


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
SHEET_NAME = "LAMBDA_functions"
TABLE_NAME = "LAMBDAFunctionsCatalog"
TABLE_HEADERS = ["Function Name", "Definition", "Arguments", "Yields", "Description"]
COLUMN_WIDTHS = {"A": 22, "B": 72, "C": 55, "D": 36, "E": 64}
ROW_HEIGHT = 130
XL_SRC_RANGE = 1
XL_YES = 1
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)


@dataclass(frozen=True)
class Argument:
    name: str
    description: str
    optional: bool = False

    def display_name(self) -> str:
        return f"[{self.name}]" if self.optional else self.name


@dataclass(frozen=True)
class LambdaCatalogEntry:
    name: str
    formula_display: str
    arguments: list[Argument]
    yields: str
    description: str

    def arguments_cell_text(self) -> str:
        if not self.arguments:
            return ""

        max_len = max(len(argument.display_name()) for argument in self.arguments)
        lines = [
            f"{argument.display_name().ljust(max_len)}  {argument.description}"
            for argument in self.arguments
        ]
        return "\n".join(lines)


def _build_argument(argument_data: dict[str, Any]) -> Argument:
    return Argument(
        name=str(argument_data["name"]),
        description=str(argument_data["description"]),
        optional=bool(argument_data.get("optional", False)),
    )


def _build_catalog_entry(function_data: dict[str, Any]) -> LambdaCatalogEntry:
    return LambdaCatalogEntry(
        name=str(function_data["name"]),
        formula_display=str(function_data["formula_display"]),
        arguments=[
            _build_argument(argument_data)
            for argument_data in function_data["arguments"]
        ],
        yields=str(function_data["yields"]),
        description=str(function_data["description"]),
    )


def load_catalog_entries(definitions_path: Path) -> list[LambdaCatalogEntry]:
    with definitions_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    functions = payload.get("functions")
    if not isinstance(functions, list):
        raise ValueError("lambda_functions.json must contain a top-level 'functions' array.")

    return [_build_catalog_entry(function_data) for function_data in functions]


def _open_or_create_workbook(app: xw.App, workbook_path: Path) -> tuple[xw.Book, bool]:
    if workbook_path.exists():
        return app.books.open(str(workbook_path)), True
    return app.books.add(), False


def _get_or_create_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet:
    for sheet in workbook.sheets:
        if sheet.name == sheet_name:
            return sheet
    return workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])


def _reset_generated_sheet(sheet: xw.Sheet) -> None:
    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()


def write_catalog_sheet(workbook: xw.Book, entries: list[LambdaCatalogEntry]) -> None:
    sheet = _get_or_create_sheet(workbook, SHEET_NAME)
    _reset_generated_sheet(sheet)
    sheet.activate()

    for column_index, header in enumerate(TABLE_HEADERS, start=1):
        sheet.range((1, column_index)).value = header

    last_data_row = len(entries) + 1
    if entries:
        sheet.range(f"B2:B{last_data_row}").api.NumberFormat = "@"

    for row_offset, entry in enumerate(entries, start=2):
        sheet.range((row_offset, 1)).value = entry.name
        sheet.range((row_offset, 2)).value = entry.formula_display
        sheet.range((row_offset, 3)).value = entry.arguments_cell_text()
        sheet.range((row_offset, 4)).value = entry.yields
        sheet.range((row_offset, 5)).value = entry.description
        sheet.range(f"{row_offset}:{row_offset}").row_height = ROW_HEIGHT

    table_range = sheet.range(f"A1:E{last_data_row}")
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = TABLE_NAME
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False

    for column_letter, width in COLUMN_WIDTHS.items():
        sheet.range(f"{column_letter}:{column_letter}").column_width = width

    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True


def _excel_error_message(exc: BaseException) -> str:
    return str(exc).strip() or exc.__class__.__name__


def _raise_excel_access_error(workbook_path: Path, action: str, exc: BaseException) -> None:
    message = _excel_error_message(exc)
    normalized = message.lower()
    likely_locked = any(
        phrase in normalized
        for phrase in (
            "cannot access read-only document",
            "read-only",
            "currently in use",
            "sharing violation",
            "permission denied",
        )
    )

    if likely_locked:
        raise RuntimeError(
            f"Excel could not {action} {workbook_path.name!r}. "
            "The workbook is likely open in Excel or locked by another process. "
            f"Close Excel and retry. Original error: {message}"
        ) from exc

    raise RuntimeError(f"Excel could not {action} {workbook_path.name!r}: {message}") from exc


def write_lambda_catalog(
    workbook_path: Path,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
) -> int:
    """Write the catalog to the workbook and return the number of entries written."""
    entries = load_catalog_entries(definitions_path)
    workbook_path = workbook_path.resolve()

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook, workbook_exists = _open_or_create_workbook(app, workbook_path)
            try:
                write_catalog_sheet(workbook, entries)
                if workbook_exists:
                    workbook.save()
                else:
                    workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        _raise_excel_access_error(workbook_path, "write catalog sheet in", exc)

    return len(entries)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Write the human-readable LAMBDA catalog into a workbook sheet named LAMBDA_functions."
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
    args = parse_args()
    rows_written = write_lambda_catalog(workbook_path=args.workbook, definitions_path=args.definitions)
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet: {SHEET_NAME}")
    print(f"Rows written: {rows_written}")


if __name__ == "__main__":
    main()
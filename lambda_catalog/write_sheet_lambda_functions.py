"""Write the LAMBDA function catalog to the LAMBDA_functions worksheet."""
from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
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


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
SHEET_NAME = "LAMBDA_functions"
TABLE_NAME = "LAMBDAFunctionsCatalog"
TABLE_HEADERS = ["Function Name", "Definition", "Arguments", "Yields", "Description"]
COLUMN_WIDTHS = {"A": 120, "B": 500, "C": 210, "D": 100, "E": 500}


@dataclass(frozen=True)
class Argument:
    """A single LAMBDA function argument with its display metadata.

    Attributes
    ----------
    name : str
        The argument identifier as used in the LAMBDA formula.
    description : str
        Human-readable description of what the argument represents.
    optional : bool
        If True, the argument is optional and is displayed in brackets.
    """

    name: str
    description: str
    optional: bool = False

    def display_name(self) -> str:
        """Return the argument name formatted for display in the catalog.

        Returns
        -------
        str
            The name wrapped in square brackets when optional, otherwise
            the bare name.
        """
        return f"[{self.name}]" if self.optional else self.name


@dataclass(frozen=True)
class LambdaCatalogEntry:
    """A single row of data for the LAMBDA_functions catalog sheet.

    Attributes
    ----------
    name : str
        The Excel defined name for this LAMBDA.
    formula_display : str
        Human-readable LAMBDA formula shown in the Definition column.
    arguments : list[Argument]
        Ordered list of arguments accepted by the function.
    yields : str
        Short description of the value the function returns.
    description : str
        Full description shown in the Description column.
    """

    name: str
    formula_display: str
    arguments: list[Argument]
    yields: str
    description: str

    def arguments_cell_text(self) -> str:
        """Format all arguments as multi-line text for the Arguments cell.

        Returns
        -------
        str
            Each argument on its own line as ``name: description``, or an
            empty string when the function takes no arguments.
        """
        if not self.arguments:
            return ""

        lines = [
            f"{argument.display_name()}: {argument.description}"
            for argument in self.arguments
        ]
        return "\n".join(lines)


def _build_argument(argument_data: dict[str, Any]) -> Argument:
    """Construct an Argument from a raw JSON argument dict.

    Parameters
    ----------
    argument_data : dict[str, Any]
        Dictionary with ``name``, ``description``, and optional ``optional``
        keys, as found in lambda_functions.json.

    Returns
    -------
    Argument
        The constructed Argument instance.
    """
    return Argument(
        name=str(argument_data["name"]),
        description=str(argument_data["description"]),
        optional=bool(argument_data.get("optional", False)),
    )


def _build_catalog_entry(function_data: dict[str, Any]) -> LambdaCatalogEntry:
    """Construct a LambdaCatalogEntry from a raw JSON function dict.

    Parameters
    ----------
    function_data : dict[str, Any]
        Dictionary for one function entry from lambda_functions.json.

    Returns
    -------
    LambdaCatalogEntry
        The constructed catalog entry.
    """
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
    """Load and parse all catalog entries from the JSON definitions file.

    Parameters
    ----------
    definitions_path : Path
        Path to lambda_functions.json.

    Returns
    -------
    list[LambdaCatalogEntry]
        Ordered list of catalog entries, one per function definition.

    Raises
    ------
    ValueError
        If the JSON does not contain a top-level ``functions`` array.
    """
    with definitions_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    functions = payload.get("functions")
    if not isinstance(functions, list):
        raise ValueError("lambda_functions.json must contain a top-level 'functions' array.")

    return [_build_catalog_entry(function_data) for function_data in functions]


def write_catalog_sheet(workbook: xw.Book, entries: list[LambdaCatalogEntry]) -> None:
    """Write catalog entries to the LAMBDA_functions sheet as a formatted table.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    entries : list[LambdaCatalogEntry]
        Catalog entries to populate the table rows.
    """
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)
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

    sheet.range("A1").column_width = 18
    sheet.range("B1").column_width = 72
    sheet.range("C1").column_width = 30
    sheet.range("D1").column_width = 15
    sheet.range("E1").column_width = 66

    sheet.range(f"A1:E{last_data_row}").api.WrapText = True
    sheet.range(f"A2:E{last_data_row}").api.EntireRow.AutoFit()

    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True


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
    entries = load_catalog_entries(definitions_path)
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

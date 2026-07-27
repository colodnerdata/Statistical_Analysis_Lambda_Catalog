"""Write the Production Lots (learning-curve) dataset to the Production Lots worksheet.

A third sample dataset, alongside Life Expectancy Data and Mileage Data,
shipped specifically because it has a natural Fixed Effects grouping column
(Facility) and a natural Sequence column (Fiscal_Year) — a small unbalanced
panel (3 facilities, 51 lots) that exercises the Regression sheet's Fixed
Effects role end to end, unlike Auto MPG (no panel-unit variable).
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import xlwings as xw
from lxml import etree

from .workbook_helpers import XL_SRC_RANGE, XL_YES, get_or_create_sheet, reset_generated_sheet


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_XLSX_PATH = ROOT_DIR / "sample_data" / "production_lots.xlsx"
SOURCE_TABLE_NAME = "Production_Lots_Data"
SHEET_NAME = "Production Lots"
TABLE_NAME = "ProductionLotsData"
FULL_DATA_HEADER = "Full_Data"
# Every column except Lot_ID (text identifier) and Facility (categorical
# grouping column) is numeric, so completeness spans Fiscal_Year through the
# derived log columns — same reasoning Mileage/Life Expectancy use to exclude
# their own text/categorical columns from Data_Completeness.
FULL_DATA_FORMULA = (
    "=Data_Completeness(ProductionLotsData[@[Fiscal_Year]:[log Unit Cost]])"
)
INTEGER_PATTERN = re.compile(r"-?\d+")

_XML_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _parse_cell(raw_value: str) -> str | int | float | None:
    """Convert a raw cell string to the most specific Python type.

    Parameters
    ----------
    raw_value : str
        The raw cell string read from the workbook XML.

    Returns
    -------
    str or int or float or None
        None for empty strings, int for whole-number strings, float when the
        value parses as a float, and str otherwise.
    """
    value = raw_value.strip()
    if value == "":
        return None

    if INTEGER_PATTERN.fullmatch(value):
        return int(value)

    try:
        return float(value)
    except ValueError:
        return value


def _cell_row_col(cell_ref: str) -> tuple[int, int]:
    """Split a cell reference like "K407" into (row, 0-based column) ints."""
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if match is None:
        raise ValueError(f"Malformed cell reference: {cell_ref!r}")
    col_letters, row_str = match.groups()
    col = 0
    for char in col_letters:
        col = col * 26 + (ord(char) - ord("A") + 1)
    return int(row_str), col - 1


def load_production_lots_rows(
    xlsx_path: Path = DEFAULT_XLSX_PATH,
) -> tuple[list[str], list[list[str | int | float | None]]]:
    """Read the Production Lots table from the sample xlsx into headers and typed rows.

    Reads the raw OOXML parts directly (zipfile + lxml) rather than opening
    the file in Excel, matching ``load_mileage_rows``/``load_life_expectancy_rows``
    so this loader (and any QC oracle built on top of it) stays usable without
    a live Excel/COM instance.

    Parameters
    ----------
    xlsx_path : Path
        Path to the Production Lots sample xlsx file.

    Returns
    -------
    tuple[list[str], list[list[str | int | float | None]]]
        A 2-tuple of (headers, rows) where each row is a list of typed cell
        values aligned with the header list, in the source table's column
        order.

    Raises
    ------
    ValueError
        If the source table is missing or contains only a header row.
    """
    with zipfile.ZipFile(xlsx_path) as archive:
        shared_strings = _read_shared_strings(archive)
        table_ref, headers = _read_table_definition(archive)
        rows = _read_table_rows(archive, table_ref, len(headers), shared_strings)

    if not rows:
        raise ValueError(f"Source table has headers but no data rows: {xlsx_path}")

    return headers, rows


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    """Return the shared string table in index order.

    A workbook with no string cells at all omits xl/sharedStrings.xml
    entirely, so a missing part is a valid (empty) table, not an error.
    """
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = etree.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(text_el.text or "" for text_el in si.findall(".//m:t", _XML_NS))
        for si in root.findall("m:si", _XML_NS)
    ]


def _read_table_definition(archive: zipfile.ZipFile) -> tuple[str, list[str]]:
    """Locate the Production_Lots_Data table part and return (ref, column names)."""
    table_parts = sorted(
        name for name in archive.namelist() if name.startswith("xl/tables/table")
    )
    for part in table_parts:
        root = etree.fromstring(archive.read(part))
        if root.get("name") == SOURCE_TABLE_NAME:
            ref = root.get("ref")
            if ref is None:
                raise ValueError(f"Table {SOURCE_TABLE_NAME!r} in {part} has no ref attribute")
            columns: list[str] = []
            for column in root.findall(".//m:tableColumn", _XML_NS):
                name = column.get("name")
                if name is None:
                    raise ValueError(
                        f"Table {SOURCE_TABLE_NAME!r} in {part} has an unnamed tableColumn"
                    )
                columns.append(name)
            return ref, columns

    raise ValueError(
        f"Table {SOURCE_TABLE_NAME!r} not found in any of: {table_parts}"
    )


def _read_table_rows(
    archive: zipfile.ZipFile,
    table_ref: str,
    n_columns: int,
    shared_strings: list[str],
) -> list[list[str | int | float | None]]:
    """Read the data rows (excluding the header row) within table_ref."""
    first_ref, last_ref = table_ref.split(":")
    first_row, _ = _cell_row_col(first_ref)
    last_row, _ = _cell_row_col(last_ref)

    root = etree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    sheet_rows = root.findall(".//m:sheetData/m:row", _XML_NS)

    rows: list[list[str | int | float | None]] = []
    for sheet_row in sheet_rows:
        row_number = int(sheet_row.get("r"))
        if row_number <= first_row or row_number > last_row:
            continue  # skip the header row and anything outside the table

        row_values: list[str | int | float | None] = [None] * n_columns
        for cell in sheet_row.findall("m:c", _XML_NS):
            _, col = _cell_row_col(cell.get("r"))
            if col >= n_columns:
                continue
            value_el = cell.find("m:v", _XML_NS)
            if value_el is None or value_el.text is None:
                continue
            cell_type = cell.get("t")
            raw_text = shared_strings[int(value_el.text)] if cell_type == "s" else value_el.text
            row_values[col] = _parse_cell(raw_text)
        rows.append(row_values)

    return rows


def write_production_lots_sheet(
    workbook: xw.Book,
    headers: list[str],
    rows: list[list[str | int | float | None]],
) -> None:
    """Write headers and data rows to the Production Lots sheet as a table.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    headers : list[str]
        Column header strings (without the computed Full_Data column).
    rows : list[list[str | int | float | None]]
        Typed data rows aligned with headers.
    """
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)
    sheet.activate()

    all_headers = headers + [FULL_DATA_HEADER]
    last_data_row = len(rows) + 1
    last_column_index = len(all_headers)
    full_data_column_index = last_column_index

    sheet.range((1, 1), (1, last_column_index)).value = all_headers
    sheet.range((2, 1), (last_data_row, len(headers))).value = rows

    table_range = sheet.range((1, 1), (last_data_row, last_column_index))
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = TABLE_NAME
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False

    sheet.range(
        (2, full_data_column_index), (last_data_row, full_data_column_index)
    ).formula = FULL_DATA_FORMULA
    sheet.used_range.columns.autofit()
    full_data_col = sheet.range((1, full_data_column_index))
    full_data_col.column_width = max(full_data_col.column_width or 0, 12)
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True

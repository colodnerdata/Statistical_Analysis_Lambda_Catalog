"""Write the Auto MPG (Mileage) dataset to the Mileage Data worksheet.

This is a second sample dataset, alongside Life Expectancy Data, shipped so a
user can practice retargeting the Regression sheet's Source_Table defined
name (Formulas -> Name Manager) between LifeExpectancyData[#All] and
MileageData[#All] before pointing it at their own data.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import xlwings as xw
from lxml import etree

from .workbook_helpers import (
    XL_SRC_RANGE,
    XL_YES,
    get_or_create_sheet,
    reset_generated_sheet,
    safe_activate,
    safe_freeze_top_row,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_XLSX_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.xlsx"
SOURCE_TABLE_NAME = "Auto_MPG_Data"
SHEET_NAME = "Mileage Data"
TABLE_NAME = "MileageData"
FULL_DATA_HEADER = "Full_Data"
# The 6 contiguous continuous-measurement columns. Model Year/Origin are
# ordinal/categorical grouping columns (excluded, same reasoning Life
# Expectancy uses to exclude Status), and Car Name/Make/Model? are text
# (ISNUMBER would always be FALSE, making Full_Data permanently FALSE).
FULL_DATA_FORMULA = "=Data_Completeness(MileageData[@[MPG]:[Acceleration]])"
INTEGER_PATTERN = re.compile(r"-?\d+")

_XML_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
# Missing values in this xlsx are encoded as the literal text "NA" (the
# source conversion process normalized the canonical UCI dataset's "?"
# markers to "NA"), unlike Life Expectancy's CSV which uses blank cells.
_MISSING_MARKER = "NA"


def _verbose_checkpoint(verbose: bool, start_time: float, label: str) -> None:
    """Print a progress checkpoint for verbose mileage sheet writes."""
    if verbose:
        from time import monotonic

        print(f"  {label:<28} {monotonic() - start_time:6.1f}s", flush=True)


def _parse_cell(raw_value: str) -> str | int | float | None:
    """Convert a raw cell string to the most specific Python type.

    Parameters
    ----------
    raw_value : str
        The raw cell string read from the workbook XML.

    Returns
    -------
    str or int or float or None
        None for empty strings and the "NA" missing-value marker, int for
        whole-number strings, float when the value parses as a float, and
        str otherwise.
    """
    value = raw_value.strip()
    if value in ("", _MISSING_MARKER):
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


def load_mileage_rows(
    xlsx_path: Path = DEFAULT_XLSX_PATH,
) -> tuple[list[str], list[list[str | int | float | None]]]:
    """Read the Auto MPG table from the sample xlsx into headers and typed rows.

    Reads the raw OOXML parts directly (zipfile + lxml) rather than opening
    the file in Excel, so this loader (and the QC oracle built on top of it)
    stays usable without a live Excel/COM instance, matching how
    ``load_life_expectancy_rows`` works as a pure-Python, pre-Excel step.

    Parameters
    ----------
    xlsx_path : Path
        Path to the Auto MPG sample xlsx file.

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
    """Locate the Auto_MPG_Data table part and return (ref, column names)."""
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


def write_mileage_sheet(
    workbook: xw.Book,
    headers: list[str],
    rows: list[list[str | int | float | None]],
    verbose: bool = False,
) -> None:
    """Write headers and data rows to the Mileage Data sheet as a table.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    headers : list[str]
        Column header strings (without the computed Full_Data column).
    rows : list[list[str | int | float | None]]
        Typed data rows aligned with headers.
    """
    start_time = __import__("time").monotonic()
    _verbose_checkpoint(verbose, start_time, "Mileage: get/create start")
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    _verbose_checkpoint(verbose, start_time, "Mileage: get/create done")
    _verbose_checkpoint(verbose, start_time, "Mileage: reset start")
    reset_generated_sheet(sheet)
    _verbose_checkpoint(verbose, start_time, "Mileage: reset done")
    safe_activate(sheet)
    _verbose_checkpoint(verbose, start_time, "Mileage: activate done")

    all_headers = headers + [FULL_DATA_HEADER]
    last_data_row = len(rows) + 1
    last_column_index = len(all_headers)
    full_data_column_index = last_column_index

    _verbose_checkpoint(
        verbose,
        start_time,
        f"Mileage: headers start ({len(headers)} cols)",
    )
    sheet.range((1, 1), (1, last_column_index)).value = all_headers
    _verbose_checkpoint(verbose, start_time, "Mileage: headers done")
    _verbose_checkpoint(
        verbose,
        start_time,
        f"Mileage: row write start ({len(rows)} rows)",
    )
    sheet.range((2, 1), (last_data_row, len(headers))).value = rows
    _verbose_checkpoint(verbose, start_time, "Mileage: row write done")

    table_range = sheet.range((1, 1), (last_data_row, last_column_index))
    _verbose_checkpoint(verbose, start_time, "Mileage: table add start")
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    _verbose_checkpoint(verbose, start_time, "Mileage: table add done")
    table.Name = TABLE_NAME
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False
    _verbose_checkpoint(verbose, start_time, "Mileage: table style done")

    _verbose_checkpoint(verbose, start_time, "Mileage: formula start")
    sheet.range(
        (2, full_data_column_index), (last_data_row, full_data_column_index)
    ).formula = FULL_DATA_FORMULA
    _verbose_checkpoint(verbose, start_time, "Mileage: formula done")
    _verbose_checkpoint(verbose, start_time, "Mileage: autofit start")
    sheet.used_range.columns.autofit()
    _verbose_checkpoint(verbose, start_time, "Mileage: autofit done")
    full_data_col = sheet.range((1, full_data_column_index))
    full_data_col.column_width = max(full_data_col.column_width or 0, 12)
    _verbose_checkpoint(verbose, start_time, "Mileage: width done")
    _verbose_checkpoint(verbose, start_time, "Mileage: freeze start")
    safe_freeze_top_row(sheet)
    _verbose_checkpoint(verbose, start_time, "Mileage: freeze done")

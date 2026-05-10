"""Write the Life Expectancy CSV dataset to the Life Expectancy Data worksheet."""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).resolve().parent))

from workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    XL_SRC_RANGE,
    XL_YES,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CSV_PATH = ROOT_DIR / "Life Expectancy Data.csv"
SHEET_NAME = "Life Expectancy Data"
TABLE_NAME = "LifeExpectancyData"
FULL_DATA_HEADER = "Full_Data"
FULL_DATA_FORMULA = "=COUNT(LifeExpectancyData[@[Life expectancy]:[Schooling]])=19"
INTEGER_PATTERN = re.compile(r"-?\d+")


def _normalize_headers(headers: list[str]) -> list[str]:
    """Deduplicate and clean raw CSV header strings.

    Parameters
    ----------
    headers : list[str]
        Raw header strings from the CSV file.

    Returns
    -------
    list[str]
        Normalized headers with internal whitespace collapsed. Duplicate
        names are made unique by appending ``_2``, ``_3``, etc.
    """
    normalized: list[str] = []
    used_names: dict[str, int] = {}

    for index, header in enumerate(headers, start=1):
        base_name = " ".join(header.strip().split()) or f"Column_{index}"
        suffix = used_names.get(base_name, 0)
        used_names[base_name] = suffix + 1
        normalized.append(base_name if suffix == 0 else f"{base_name}_{suffix + 1}")

    return normalized


def _parse_cell(raw_value: str) -> str | int | float | None:
    """Convert a raw CSV cell string to the most specific Python type.

    Parameters
    ----------
    raw_value : str
        The raw cell string from the CSV reader.

    Returns
    -------
    str or int or float or None
        None for empty strings, int for whole-number strings, float when
        the value parses as a float, and str otherwise.
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


def load_life_expectancy_rows(
    csv_path: Path,
) -> tuple[list[str], list[list[str | int | float | None]]]:
    """Read the CSV and return normalized headers and typed data rows.

    Parameters
    ----------
    csv_path : Path
        Path to the Life Expectancy CSV file.

    Returns
    -------
    tuple[list[str], list[list[str | int | float | None]]]
        A 2-tuple of (headers, rows) where each row is a list of typed
        cell values aligned with the header list.

    Raises
    ------
    ValueError
        If the CSV file is empty or contains only a header row.
    """
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        try:
            raw_headers = next(reader)
        except StopIteration as exc:
            raise ValueError(f"CSV file is empty: {csv_path}") from exc

        headers = _normalize_headers(raw_headers)
        rows = [[_parse_cell(value) for value in row] for row in reader]

    if not rows:
        raise ValueError(f"CSV file has headers but no data rows: {csv_path}")

    return headers, rows


def write_life_expectancy_sheet(
    workbook: xw.Book,
    headers: list[str],
    rows: list[list[str | int | float | None]],
) -> None:
    """Write headers and data rows to the Life Expectancy Data sheet as a table.

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


def write_life_expectancy_data(
    workbook_path: Path,
    csv_path: Path = DEFAULT_CSV_PATH,
) -> int:
    """Write the life expectancy sheet to the workbook and return the row count.

    Parameters
    ----------
    workbook_path : Path
        Path to the workbook file to create or update.
    csv_path : Path, optional
        Path to the Life Expectancy CSV file.

    Returns
    -------
    int
        Number of data rows written (excluding the header row).
    """
    headers, rows = load_life_expectancy_rows(csv_path)
    workbook_path = workbook_path.resolve()

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook, workbook_exists = open_or_create_workbook(app, workbook_path)
            try:
                write_life_expectancy_sheet(workbook, headers, rows)
                if workbook_exists:
                    workbook.save()
                else:
                    workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "write life expectancy data in", exc)

    return len(rows)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the life expectancy sheet writer.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with workbook and csv attributes.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Write Life Expectancy Data.csv into a workbook sheet "
            "named 'Life Expectancy Data'."
        )
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the workbook file to create or update.",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_CSV_PATH,
        help="Path to the Life Expectancy CSV file.",
    )
    return parser.parse_args()


def main() -> None:
    """Write the life expectancy sheet and print a summary for interactive use."""
    args = parse_args()
    row_count = write_life_expectancy_data(workbook_path=args.workbook, csv_path=args.csv)
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet: {SHEET_NAME}")
    print(f"Table: {TABLE_NAME}")
    print(f"Rows written: {row_count}")


if __name__ == "__main__":
    main()

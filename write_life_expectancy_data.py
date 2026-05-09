from __future__ import annotations

import argparse
import csv
import re
from pathlib import Path

import pywintypes  # type: ignore[import-untyped]
import xlwings as xw


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_CSV_PATH = ROOT_DIR / "Life Expectancy Data.csv"
SHEET_NAME = "Life Expectancy Data"
TABLE_NAME = "LifeExpectancyData"
FULL_DATA_HEADER = "Full_Data"
FULL_DATA_FORMULA = "=COUNT(LifeExpectancyData[@[Life expectancy]:[Schooling]])=19"
XL_SRC_RANGE = 1
XL_YES = 1
INTEGER_PATTERN = re.compile(r"-?\d+")
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)


def _normalize_headers(headers: list[str]) -> list[str]:
    normalized: list[str] = []
    used_names: dict[str, int] = {}

    for index, header in enumerate(headers, start=1):
        base_name = " ".join(header.strip().split()) or f"Column_{index}"
        suffix = used_names.get(base_name, 0)
        used_names[base_name] = suffix + 1
        normalized.append(base_name if suffix == 0 else f"{base_name}_{suffix + 1}")

    return normalized


def _parse_cell(raw_value: str) -> str | int | float | None:
    value = raw_value.strip()
    if value == "":
        return None

    if INTEGER_PATTERN.fullmatch(value):
        return int(value)

    try:
        return float(value)
    except ValueError:
        return value


def load_life_expectancy_rows(csv_path: Path) -> tuple[list[str], list[list[str | int | float | None]]]:
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


def write_life_expectancy_sheet(
    workbook: xw.Book,
    headers: list[str],
    rows: list[list[str | int | float | None]],
) -> None:
    sheet = _get_or_create_sheet(workbook, SHEET_NAME)
    _reset_generated_sheet(sheet)
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

    sheet.range((2, full_data_column_index), (last_data_row, full_data_column_index)).formula = FULL_DATA_FORMULA
    sheet.used_range.columns.autofit()
    sheet.range((1, full_data_column_index)).column_width = max(sheet.range((1, full_data_column_index)).column_width, 12)
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


def write_life_expectancy_data(
    workbook_path: Path,
    csv_path: Path = DEFAULT_CSV_PATH,
) -> int:
    headers, rows = load_life_expectancy_rows(csv_path)
    workbook_path = workbook_path.resolve()

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook, workbook_exists = _open_or_create_workbook(app, workbook_path)
            try:
                write_life_expectancy_sheet(workbook, headers, rows)
                if workbook_exists:
                    workbook.save()
                else:
                    workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        _raise_excel_access_error(workbook_path, "write life expectancy data in", exc)

    return len(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Write Life Expectancy Data.csv into a workbook sheet named 'Life Expectancy Data'."
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
    args = parse_args()
    row_count = write_life_expectancy_data(workbook_path=args.workbook, csv_path=args.csv)
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet: {SHEET_NAME}")
    print(f"Table: {TABLE_NAME}")
    print(f"Rows written: {row_count}")


if __name__ == "__main__":
    main()
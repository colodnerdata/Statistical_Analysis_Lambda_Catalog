"""Write a CSV-backed sample dataset to its worksheet.

Unifies the three sample datasets shipped with this project — Life
Expectancy, Auto MPG ("Mileage"), and Production Lots — behind one
config-driven loader/writer/CLI instead of three near-duplicate modules.
Each dataset's quirks (missing-value markers, header normalization) are
captured in its ``CsvDatasetConfig`` rather than in bespoke parsing code.
"""
from __future__ import annotations

import argparse
import csv
import re
from dataclasses import dataclass
from pathlib import Path
from time import monotonic

import xlwings as xw

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
INTEGER_PATTERN = re.compile(r"-?\d+")
_EXCEL_ERROR_LITERALS = frozenset(
    {"#VALUE!", "#REF!", "#NAME?", "#NULL!", "#DIV/0!", "#N/A", "#NUM!"}
)


@dataclass(frozen=True)
class CsvDatasetConfig:
    """Everything a dataset needs from the shared load/write engine."""

    name: str
    default_csv_path: Path
    sheet_name: str
    table_name: str
    # An optional formula column appended after the CSV columns (written once
    # across the whole body range, evaluated per row via structured refs). None
    # for datasets that ship no appended column. Life Expectancy appends
    # "Developed Country after 2013" (=AND([@Status]="Developed",[@Year]>2013));
    # Mileage and Production Lots append nothing.
    derived_header: str | None = None
    derived_formula: str | None = None
    missing_values: frozenset[str] = frozenset({""})
    normalize_headers: bool = False


MILEAGE = CsvDatasetConfig(
    name="mileage",
    default_csv_path=ROOT_DIR / "sample_data" / "auto_mpg_data.csv",
    sheet_name="Mileage Data",
    table_name="MileageData",
    # Missing values in this CSV are encoded as the literal text "NA" (the
    # source conversion process normalized the canonical UCI dataset's "?"
    # markers to "NA"), unlike Life Expectancy's CSV which uses blank cells.
    missing_values=frozenset({"", "NA"}),
)

LIFE_EXPECTANCY = CsvDatasetConfig(
    name="life_expectancy",
    default_csv_path=ROOT_DIR / "sample_data" / "Life Expectancy Data.csv",
    sheet_name="Life Expectancy Data",
    table_name="LifeExpectancyData",
    # The one shipped derived column: a dormant filter flagging developed
    # countries from 2014 on. Ships Role=Omit (dormant) in the spec block --
    # flip Role to Filter to restrict the sample. AND([@Status]="Developed",
    # [@Year]>2013) is TRUE for developed-country rows with Year >= 2014.
    derived_header="Developed Country after 2013",
    derived_formula='=AND([@Status]="Developed",[@Year]>2013)',
    normalize_headers=True,
)

PRODUCTION_LOTS = CsvDatasetConfig(
    name="production_lots",
    default_csv_path=ROOT_DIR / "sample_data" / "production_lots.csv",
    sheet_name="Production Lots",
    table_name="ProductionLotsData",
)

_DATASETS = {config.name: config for config in (MILEAGE, LIFE_EXPECTANCY, PRODUCTION_LOTS)}


def _verbose_checkpoint(verbose: bool, start_time: float, label: str) -> None:
    """Print a progress checkpoint for verbose sheet writes."""
    if verbose:
        print(f"  {label:<28} {monotonic() - start_time:6.1f}s", flush=True)


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


def _parse_cell(raw_value: str, missing_values: frozenset[str]) -> str | int | float | None:
    """Convert a raw CSV cell string to the most specific Python type.

    Parameters
    ----------
    raw_value : str
        The raw cell string from the CSV reader.
    missing_values : frozenset[str]
        Stripped cell values that should be treated as missing.

    Returns
    -------
    str or int or float or None
        None when the stripped value is in ``missing_values``, int for
        whole-number strings, float when the value parses as a float, and
        str otherwise.
    """
    value = raw_value.strip()
    if value in missing_values:
        return None

    if INTEGER_PATTERN.fullmatch(value):
        return int(value)

    try:
        return float(value)
    except ValueError:
        return value


def _excel_safe_cell_value(value: str | int | float | None) -> str | int | float | None:
    """Return a value that Excel will keep as a literal cell value.

    Excel treats strings matching its error literals (for example ``"#VALUE!"``)
    as actual error cells when written through COM. Prefix them with an
    apostrophe so the displayed text stays the same but the stored cell type is
    text, not ``<c t="e">``.
    """
    if isinstance(value, str) and value in _EXCEL_ERROR_LITERALS:
        return f"'{value}"
    return value


def load_csv_rows(
    csv_path: Path,
    config: CsvDatasetConfig,
) -> tuple[list[str], list[list[str | int | float | None]]]:
    """Read a dataset's CSV and return typed headers and data rows.

    Parameters
    ----------
    csv_path : Path
        Path to the dataset's CSV file.
    config : CsvDatasetConfig
        The dataset config supplying header-normalization and
        missing-value rules.

    Returns
    -------
    tuple[list[str], list[list[str | int | float | None]]]
        A 2-tuple of (headers, rows) where each row is a list of typed
        cell values aligned with the header list.

    Raises
    ------
    ValueError
        If the CSV file is empty, contains only a header row, or contains a
        data row with more columns than the header (fully blank lines and
        short rows are tolerated: blank lines are skipped and short rows are
        padded with missing values).
    """
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        try:
            raw_headers = next(reader)
        except StopIteration as exc:
            raise ValueError(f"CSV file is empty: {csv_path}") from exc

        headers = _normalize_headers(raw_headers) if config.normalize_headers else list(raw_headers)
        n_columns = len(headers)
        rows: list[list[str | int | float | None]] = []
        for raw_row in reader:
            if not raw_row:
                continue  # skip fully blank lines
            if len(raw_row) > n_columns:
                raise ValueError(
                    f"CSV row at line {reader.line_num} has {len(raw_row)} columns, "
                    f"more than the {n_columns} header columns: {csv_path}"
                )
            padded_row = raw_row + [""] * (n_columns - len(raw_row))
            rows.append([_parse_cell(value, config.missing_values) for value in padded_row])

    if not rows:
        raise ValueError(f"CSV file has headers but no data rows: {csv_path}")

    return headers, rows


def write_csv_dataset_sheet(
    workbook: xw.Book,
    headers: list[str],
    rows: list[list[str | int | float | None]],
    config: CsvDatasetConfig,
    verbose: bool = False,
) -> None:
    """Write headers and data rows to the dataset's sheet as a table.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to write into.
    headers : list[str]
        Column header strings (without the optional derived column).
    rows : list[list[str | int | float | None]]
        Typed data rows aligned with headers.
    config : CsvDatasetConfig
        The dataset config supplying the sheet name, table name, and optional
        derived formula column.
    """
    start_time = monotonic()

    def checkpoint(label: str) -> None:
        _verbose_checkpoint(verbose, start_time, f"{config.name}: {label}")

    checkpoint("get/create start")
    sheet = get_or_create_sheet(workbook, config.sheet_name)
    checkpoint("get/create done")
    checkpoint("reset start")
    reset_generated_sheet(sheet)
    checkpoint("reset done")
    safe_activate(sheet)
    checkpoint("activate done")

    # The optional derived formula column (Life Expectancy's "Developed
    # Country after 2013"); None for Mileage/Production Lots, which ship no
    # appended column. When present it is the last column of the table.
    derived_header = config.derived_header
    if derived_header is not None:
        all_headers = headers + [derived_header]
    else:
        all_headers = headers
    last_data_row = len(rows) + 1
    last_column_index = len(all_headers)

    checkpoint(f"headers start ({len(headers)} cols)")
    sheet.range((1, 1), (1, last_column_index)).value = all_headers
    checkpoint("headers done")
    checkpoint(f"row write start ({len(rows)} rows)")
    safe_rows = [
        [_excel_safe_cell_value(value) for value in row]
        for row in rows
    ]
    sheet.range((2, 1), (last_data_row, len(headers))).value = safe_rows
    checkpoint("row write done")

    table_range = sheet.range((1, 1), (last_data_row, last_column_index))
    checkpoint("table add start")
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    checkpoint("table add done")
    table.Name = config.table_name
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    table.ShowTableStyleColumnStripes = False
    checkpoint("table style done")

    # Write the derived formula across its body range only when a derived
    # column is present. last_column_index is that column's index then.
    if config.derived_formula is not None:
        checkpoint("formula start")
        sheet.range(
            (2, last_column_index), (last_data_row, last_column_index)
        ).formula = config.derived_formula
        checkpoint("formula done")
    checkpoint("autofit start")
    sheet.used_range.columns.autofit()
    checkpoint("autofit done")
    if last_column_index > len(headers):
        derived_col = sheet.range((1, last_column_index))
        derived_col.column_width = max(derived_col.column_width or 0, 12)
    checkpoint("width done")
    checkpoint("freeze start")
    safe_freeze_top_row(sheet)
    checkpoint("freeze done")


def write_csv_dataset(
    workbook_path: Path,
    config: CsvDatasetConfig,
    csv_path: Path | None = None,
) -> int:
    """Write the dataset sheet to the workbook and return the row count.

    Parameters
    ----------
    workbook_path : Path
        Path to the workbook file to create or update.
    config : CsvDatasetConfig
        The dataset to write.
    csv_path : Path, optional
        Path to the dataset's CSV file. Defaults to ``config.default_csv_path``.

    Returns
    -------
    int
        Number of data rows written (excluding the header row).
    """
    csv_path = csv_path or config.default_csv_path
    headers, rows = load_csv_rows(csv_path, config)
    workbook_path = workbook_path.resolve()

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook, workbook_exists = open_or_create_workbook(app, workbook_path)
            try:
                write_csv_dataset_sheet(workbook, headers, rows, config)
                if workbook_exists:
                    workbook.save()
                else:
                    workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, f"write {config.name} data in", exc)

    return len(rows)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the CSV dataset sheet writer.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with dataset, workbook, and csv attributes.
    """
    parser = argparse.ArgumentParser(
        description="Write a sample CSV dataset into a workbook sheet.",
    )
    parser.add_argument(
        "dataset",
        choices=sorted(_DATASETS),
        help="Which sample dataset to write.",
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the workbook file to create or update.",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=None,
        help="Path to the dataset's CSV file (defaults to the dataset's sample CSV).",
    )
    return parser.parse_args()


def main() -> None:
    """Write the selected dataset sheet and print a summary for interactive use."""
    args = parse_args()
    config = _DATASETS[args.dataset]
    row_count = write_csv_dataset(workbook_path=args.workbook, config=config, csv_path=args.csv)
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet: {config.sheet_name}")
    print(f"Table: {config.table_name}")
    print(f"Rows written: {row_count}")


if __name__ == "__main__":
    main()

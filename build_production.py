"""Build Lambda_Library.xlsx: LAMBDA_functions, Life Expectancy Data, Univariate Analysis, and Regression."""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

import xlwings as xw

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
    _delete_sheet_if_present,
    _validate_workbook_reopen,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet
from lambda_catalog.write_sheet_life_expectancy_data import (
    DEFAULT_CSV_PATH,
    load_life_expectancy_rows,
    write_life_expectancy_sheet,
)
from lambda_catalog.write_sheet_diagnostic_guide import write_diagnostic_guide_sheet
from lambda_catalog.write_sheet_regression import write_regression_output_sheet
from lambda_catalog.write_sheet_regression_instructions import write_regression_instructions_sheet
from lambda_catalog.write_sheet_univariate import write_univariate_sheet
from lambda_catalog.write_sheet_version_history import write_version_history_sheet


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
_PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
_QC_SHEET_NAMES = ("MLR_Scalar_Test", "MLR_Vector_Outputs_Test", "MLR_Observation_Test")


def _recalculate_and_save(workbook_path: Path) -> None:
    """Fully calculate Data Tables after name sync, then save semiautomatic mode."""
    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
            try:
                app.api.Calculation = XL_CALCULATION_AUTOMATIC
                app.api.CalculateFullRebuild()
                app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "recalculate and save", exc)


def _backup_unopenable_workbook(workbook_path: Path) -> Path:
    """Move aside a workbook Excel cannot open so the build can recreate it."""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    backup_path = workbook_path.with_name(
        f"{workbook_path.stem}_unopenable_{timestamp}{workbook_path.suffix}"
    )
    workbook_path.replace(backup_path)
    return backup_path


def build_production_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = DEFAULT_CSV_PATH,
    validate_reopen: bool = False,
    verbose: bool = False,
) -> NameSyncResult:
    """Build the production sheets and sync the LAMBDA name manager.

    Parameters
    ----------
    workbook_path : Path, optional
        Path to the workbook to create or update.
    definitions_path : Path, optional
        Path to the JSON catalog file.
    csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    validate_reopen : bool, optional
        If True, reopens the workbook in Excel after patching to verify it.
    verbose : bool, optional
        If True, prints timing information for each build phase to stdout.

    Returns
    -------
    NameSyncResult
        Counts of created versus updated workbook names.
    """
    _t = time.monotonic()
    document = load_catalog_document(definitions_path)
    csv_headers, csv_rows = load_life_expectancy_rows(csv_path)
    if verbose:
        print(f"  Prep:           {time.monotonic() - _t:.1f}s", flush=True)

    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    _t = time.monotonic()
    try:
        with xw.App(visible=True, add_book=False) as app:
            workbook = None
            rebuilt_from_scratch = False
            if workbook_exists:
                try:
                    workbook = app.books.open(str(workbook_path))
                except OPEN_WORKBOOK_ERRORS as exc:
                    message = str(exc).lower()
                    if "open method of workbooks class failed" not in message:
                        raise
                    try:
                        _backup_unopenable_workbook(workbook_path)
                    except OSError as backup_exc:
                        raise_excel_access_error(workbook_path, "open", backup_exc)
                    workbook = app.books.add()
                    rebuilt_from_scratch = True
            else:
                workbook = app.books.add()
                rebuilt_from_scratch = True

            if rebuilt_from_scratch:
                for sheet in list(workbook.sheets)[1:]:
                    sheet.delete()

            try:
                app.api.Calculation = XL_CALCULATION_MANUAL
                _delete_sheet_if_present(workbook, _PREDICTIONS_SHEET_NAME)
                for qc_sheet in _QC_SHEET_NAMES:
                    _delete_sheet_if_present(workbook, qc_sheet)
                if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                    workbook.sheets["Sheet1"].name = "LAMBDA_functions"
                write_catalog_sheet(workbook, document.functions)
                write_life_expectancy_sheet(workbook, csv_headers, csv_rows)
                write_univariate_sheet(workbook)
                write_regression_instructions_sheet(workbook)
                write_diagnostic_guide_sheet(workbook)
                write_version_history_sheet(workbook)
                write_regression_output_sheet(workbook, document.regression_sheet_notes)
                app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)
    if verbose:
        print(f"  Write sheets:   {time.monotonic() - _t:.1f}s", flush=True)

    _t = time.monotonic()
    try:
        result = sync_workbook_names(workbook_path, document.functions)
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    except (PermissionError, OSError) as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    if verbose:
        print(f"  Sync names:     {time.monotonic() - _t:.1f}s", flush=True)

    _t = time.monotonic()
    _recalculate_and_save(workbook_path)
    if verbose:
        print(f"  Recalculate:    {time.monotonic() - _t:.1f}s", flush=True)

    if validate_reopen:
        _validate_workbook_reopen(workbook_path)

    return result


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for production workbook generation.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with workbook, definitions, csv, validate_reopen,
        and verbose attributes.
    """
    parser = argparse.ArgumentParser(
        description="Build Lambda_Library.xlsx from lambda_functions.json."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK_PATH,
        help="Path to the workbook to create or update.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda definitions.",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_CSV_PATH,
        help="Path to the Life Expectancy CSV data file.",
    )
    parser.add_argument(
        "--validate-reopen",
        action="store_true",
        help="Reopen the workbook after syncing names to verify Excel accepts the result.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print timing information for each build phase.",
    )
    return parser.parse_args()


def main() -> None:
    """Build the production workbook and print a short sync summary for interactive use."""
    args = parse_args()
    while True:
        try:
            result = build_production_workbook(
                workbook_path=args.workbook,
                definitions_path=args.definitions,
                csv_path=args.csv,
                validate_reopen=args.validate_reopen,
                verbose=args.verbose,
            )
            break
        except RuntimeError as exc:
            if "likely open in Excel" in str(exc):
                if not sys.stdin.isatty():
                    raise
                prompt = (
                    f"\n{args.workbook.name} is open in Excel — close it "
                    "and press Enter to retry (or Ctrl+C to cancel): "
                )
                input(prompt)
            else:
                raise

    print(f"Workbook: {args.workbook.resolve()}")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print("Sheet updated: Univariate Analysis")
    print("Sheet updated: Regression Instructions")
    print("Sheet updated: Diagnostic Guide")
    print("Sheet updated: Version History")
    print("Sheet updated: Regression")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    if args.validate_reopen:
        print("Reopen validation: passed")

    subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])


if __name__ == "__main__":
    main()

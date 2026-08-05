"""Build Lambda_Library_QC.xlsx and verify the QC workbook."""
from __future__ import annotations

import argparse
import io
import subprocess
import sys
import time
from pathlib import Path
from typing import TextIO

import xlwings as xw

from lambda_catalog.analyze_regression_spec import build_regression_spec_qc_configs
from lambda_catalog.analysis_cache import DEFAULT_CACHE_PATH
from lambda_catalog.build_common import print_name_sync_summary, warn_if_workbook_open
from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
    _delete_sheet_if_present,
    _validate_workbook_reopen,
    drop_workbook_names,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import (
    LOCK_HINT,
    OPEN_WORKBOOK_ERRORS,
    raise_excel_access_error,
)
from lambda_catalog.write_sheet_diagnostic_guide import write_diagnostic_guide_sheet
from lambda_catalog.write_sheet_dummy_test import write_dummy_test_sheet
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    MILEAGE,
    PRODUCTION_LOTS,
    load_csv_rows,
    write_csv_dataset_sheet,
)
from lambda_catalog.write_sheet_regression import write_regression_output_sheet
from lambda_catalog.write_sheet_regression_instructions import (
    write_regression_instructions_sheet,
)
from lambda_catalog.write_sheet_univariate import write_univariate_sheet
from lambda_catalog.write_sheet_version_history import write_version_history_sheet
from lambda_catalog.deep_verify import _verbose_checkpoint, verify_test_sheets


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library_QC.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
DEFAULT_LOG_PATH = ROOT_DIR / "logs" / "qc_log.log"
_PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
_QC_SHEET_NAMES = (
    "MLR_Scalar_Test",
    "MLR_Vector_Outputs_Test",
    "MLR_Observation_Test",
    "Dummy_Test",
)

def build_qc_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = LIFE_EXPECTANCY.default_csv_path,
    cache_path: Path = DEFAULT_CACHE_PATH,
    validate_reopen: bool = False,
    verbose: bool = False,
    *,
    mileage_csv_path: Path = MILEAGE.default_csv_path,
    production_lots_csv_path: Path = PRODUCTION_LOTS.default_csv_path,
    no_verify: bool = False,
    timings_out: dict[str, float | None] | None = None,
) -> NameSyncResult:
    """Build production/QC sheets and verify the workbook."""
    _ = cache_path
    phase_start = time.monotonic()
    document = load_catalog_document(definitions_path)
    _verbose_checkpoint(verbose, phase_start, "Prep: catalog loaded")
    regression_sheet_configs = build_regression_spec_qc_configs(mileage_csv_path)
    _verbose_checkpoint(verbose, phase_start, "Prep: regression QC loaded")
    csv_headers, csv_rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)
    mileage_headers, mileage_rows = load_csv_rows(mileage_csv_path, MILEAGE)
    production_lots_headers, production_lots_rows = load_csv_rows(
        production_lots_csv_path, PRODUCTION_LOTS
    )
    _verbose_checkpoint(verbose, phase_start, "Prep: csv loaded")
    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    if workbook_exists:
        try:
            _verbose_checkpoint(verbose, phase_start, "Prep: drop names start")
            drop_workbook_names(
                workbook_path,
                (definition.name for definition in document.workbook_functions),
            )
            _verbose_checkpoint(verbose, phase_start, "Prep: drop names done")
        except OPEN_WORKBOOK_ERRORS as exc:
            raise_excel_access_error(workbook_path, "update", exc)
        except (PermissionError, OSError) as exc:
            raise_excel_access_error(workbook_path, "update", exc)

    _verbose_checkpoint(verbose, phase_start, "Prep total")
    prep_elapsed = time.monotonic() - phase_start

    phase_start = time.monotonic()
    try:
        _verbose_checkpoint(verbose, phase_start, "Write: app start")
        with xw.App(visible=False, add_book=False) as app:
            _verbose_checkpoint(verbose, phase_start, "Write: app ready")
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            if workbook_exists:
                _verbose_checkpoint(verbose, phase_start, "Write: open workbook")
                workbook = app.books.open(str(workbook_path))
                _verbose_checkpoint(verbose, phase_start, "Write: workbook opened")
            else:
                _verbose_checkpoint(verbose, phase_start, "Write: add workbook")
                workbook = app.books.add()
                _verbose_checkpoint(verbose, phase_start, "Write: workbook added")
                for sheet in list(workbook.sheets)[1:]:
                    sheet.delete()

            try:
                app.api.Calculation = XL_CALCULATION_MANUAL
                _delete_sheet_if_present(workbook, _PREDICTIONS_SHEET_NAME)
                _delete_sheet_if_present(workbook, "Model Construction")
                for qc_sheet in _QC_SHEET_NAMES:
                    _delete_sheet_if_present(workbook, qc_sheet)
                if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                    workbook.sheets["Sheet1"].name = "LAMBDA_functions"

                _verbose_checkpoint(verbose, phase_start, "Write: catalog start")
                write_catalog_sheet(workbook, document.functions)
                _verbose_checkpoint(verbose, phase_start, "Write: catalog done")
                _verbose_checkpoint(verbose, phase_start, "Write: life exp start")
                write_csv_dataset_sheet(
                    workbook, csv_headers, csv_rows, LIFE_EXPECTANCY, verbose=verbose
                )
                _verbose_checkpoint(verbose, phase_start, "Write: life exp done")
                _verbose_checkpoint(verbose, phase_start, "Write: mileage start")
                write_csv_dataset_sheet(
                    workbook, mileage_headers, mileage_rows, MILEAGE, verbose=verbose
                )
                _verbose_checkpoint(verbose, phase_start, "Write: mileage done")
                _verbose_checkpoint(verbose, phase_start, "Write: production lots start")
                write_csv_dataset_sheet(
                    workbook, production_lots_headers, production_lots_rows, PRODUCTION_LOTS
                )
                _verbose_checkpoint(verbose, phase_start, "Write: production lots done")
                _verbose_checkpoint(verbose, phase_start, "Write: univariate start")
                write_univariate_sheet(workbook)
                _verbose_checkpoint(verbose, phase_start, "Write: univariate done")
                _verbose_checkpoint(verbose, phase_start, "Write: regression instr start")
                write_regression_instructions_sheet(workbook)
                _verbose_checkpoint(verbose, phase_start, "Write: regression instr done")
                _verbose_checkpoint(verbose, phase_start, "Write: diagnostic start")
                write_diagnostic_guide_sheet(workbook)
                _verbose_checkpoint(verbose, phase_start, "Write: diagnostic done")
                _verbose_checkpoint(verbose, phase_start, "Write: version history start")
                write_version_history_sheet(workbook)
                _verbose_checkpoint(verbose, phase_start, "Write: version history done")
                _verbose_checkpoint(verbose, phase_start, "Write: regression start")
                write_regression_output_sheet(
                    workbook,
                    document.regression_sheet_notes,
                    document.functions_for_sheet("Regression"),
                )
                _verbose_checkpoint(verbose, phase_start, "Write: regression done")
                _verbose_checkpoint(verbose, phase_start, "Write: dummy test start")
                write_dummy_test_sheet(workbook)
                _verbose_checkpoint(verbose, phase_start, "Write: dummy test done")
                app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
                _verbose_checkpoint(verbose, phase_start, "Write: save start")
                workbook.save(str(workbook_path))
                _verbose_checkpoint(verbose, phase_start, "Write: save done")
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)
    write_elapsed = time.monotonic() - phase_start
    if verbose:
        print(f"  Write sheets:   {write_elapsed:.1f}s", flush=True)

    phase_start = time.monotonic()
    try:
        _verbose_checkpoint(verbose, phase_start, "Sync: start")
        result = sync_workbook_names(workbook_path, document.workbook_functions)
        _verbose_checkpoint(verbose, phase_start, "Sync: done")
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    except (PermissionError, OSError) as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    sync_elapsed = time.monotonic() - phase_start
    if verbose:
        print(f"  Sync names:     {sync_elapsed:.1f}s", flush=True)

    phase_start = time.monotonic()
    verify_elapsed: float | None
    if no_verify:
        print("Verify: skipped (--no-verify)", flush=True)
        verify_elapsed = None
    else:
        try:
            _verbose_checkpoint(verbose, phase_start, "Verify: app start")
            with xw.App(visible=False, add_book=False) as app:
                _verbose_checkpoint(verbose, phase_start, "Verify: app ready")
                app.api.DisplayAlerts = False
                app.api.AskToUpdateLinks = False
                _verbose_checkpoint(verbose, phase_start, "Verify: open workbook")
                workbook = app.books.open(str(workbook_path))
                try:
                    _verbose_checkpoint(verbose, phase_start, "Verify: workbook opened")
                    verify_test_sheets(
                        workbook,
                        regression_sheet_configs,
                        csv_path,
                        mileage_path=mileage_csv_path,
                        production_lots_path=production_lots_csv_path,
                        verbose=verbose,
                    )
                finally:
                    workbook.close()
        except OPEN_WORKBOOK_ERRORS as exc:
            raise_excel_access_error(workbook_path, "verify", exc)
        verify_elapsed = time.monotonic() - phase_start
        if verbose:
            print(f"  Verify:         {verify_elapsed:.1f}s", flush=True)

    if validate_reopen:
        _validate_workbook_reopen(workbook_path)

    if timings_out is not None:
        timings_out.update(
            {
                "prep_seconds": prep_elapsed,
                "write_seconds": write_elapsed,
                "sync_seconds": sync_elapsed,
                "verify_seconds": verify_elapsed,
            }
        )

    return result


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for QC workbook generation."""
    parser = argparse.ArgumentParser(
        description="Build Lambda_Library_QC.xlsx with spec-driven Regression QC."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK_PATH,
        help="Path to the QC workbook to create or update.",
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
        default=LIFE_EXPECTANCY.default_csv_path,
        help="Path to the Life Expectancy CSV data file.",
    )
    parser.add_argument(
        "--mileage-csv",
        type=Path,
        default=MILEAGE.default_csv_path,
        help="Path to the Auto MPG (Mileage) sample CSV data file.",
    )
    parser.add_argument(
        "--production-lots-csv",
        type=Path,
        default=PRODUCTION_LOTS.default_csv_path,
        help="Path to the Production Lots sample CSV data file.",
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=DEFAULT_CACHE_PATH,
        help="Retained for compatibility; spec-driven QC computes on demand.",
    )
    parser.add_argument(
        "--validate-reopen",
        action="store_true",
        help="Reopen the workbook after syncing names to verify Excel accepts the result.",
    )
    parser.add_argument(
        "--no-verify",
        action="store_true",
        help="Skip the spec-driven verify pass. Escape hatch for iterating on a "
        "known-broken sheet; the skip is logged to qc_log.txt so the absence "
        "is visible.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print timing information for each build phase.",
    )
    return parser.parse_args()


class _Tee(io.TextIOBase):
    """Write to two streams simultaneously."""

    def __init__(self, primary: TextIO, secondary: TextIO) -> None:
        self._primary = primary
        self._secondary = secondary

    def write(self, s: str) -> int:
        self._secondary.write(s)
        return self._primary.write(s)

    def flush(self) -> None:
        self._primary.flush()
        self._secondary.flush()


def main() -> None:
    """Build the QC workbook and print a sync summary for interactive use."""
    args = parse_args()
    # logs/ is gitignored, so it does not exist in a fresh clone or after a
    # clean. Create it the same way build_common.run_log_path's writer does,
    # or this open() raises FileNotFoundError before the build even starts.
    DEFAULT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(DEFAULT_LOG_PATH, "w", encoding="utf-8") as log_file:
        real_stdout = sys.stdout
        sys.stdout = _Tee(sys.stdout, log_file)
        try:
            print(f"python {Path(__file__).name} " + " ".join(sys.argv[1:]))
            _run_main(args)
        finally:
            sys.stdout = real_stdout


def _run_main(args: argparse.Namespace) -> None:
    total_start = time.monotonic()
    timings: dict[str, float | None] = {}
    # Up front, because the retry loop below only learns of a lock when the
    # save fails at the end of the write: Excel opens a locked workbook
    # read-only without raising, and this loop retries the ENTIRE build.
    warn_if_workbook_open(
        args.workbook, action_label=f"{args.workbook.name} is open in Excel"
    )
    while True:
        try:
            result = build_qc_workbook(
                workbook_path=args.workbook,
                definitions_path=args.definitions,
                csv_path=args.csv,
                cache_path=args.cache,
                validate_reopen=args.validate_reopen,
                verbose=args.verbose,
                mileage_csv_path=args.mileage_csv,
                production_lots_csv_path=args.production_lots_csv,
                no_verify=args.no_verify,
                timings_out=timings,
            )
            break
        except RuntimeError as exc:
            if LOCK_HINT in str(exc):
                if not sys.stdin.isatty():
                    raise
                prompt = (
                    f"\n{args.workbook.name} is open in Excel - close it "
                    "and press Enter to retry (or Ctrl+C to cancel): "
                )
                input(prompt)
            else:
                raise

    print(f"Workbook: {args.workbook.resolve()}")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print("Sheet updated: Mileage Data")
    print("Sheet updated: Univariate")
    print("Sheet updated: Regression Instructions")
    print("Sheet updated: Diagnostic Guide")
    print("Sheet updated: Version History")
    print("Sheet updated: Regression")
    print("Sheet updated: Dummy_Test")
    if timings["verify_seconds"] is not None:
        print("Sheet verified: Regression")
        print("Sheet verified: Univariate")
        print("Sheet verified: Dummy_Test")
    print_name_sync_summary(result)
    if args.validate_reopen:
        print("Reopen validation: passed")
    print(f"Timing: prep          {timings['prep_seconds']:.1f}s")
    print(f"Timing: write sheets  {timings['write_seconds']:.1f}s")
    print(f"Timing: sync names    {timings['sync_seconds']:.1f}s")
    if timings["verify_seconds"] is None:
        print("Timing: verify        skipped")
    else:
        print(f"Timing: verify        {timings['verify_seconds']:.1f}s")
    print(f"Timing: total         {time.monotonic() - total_start:.1f}s")

    subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])
    print(f"Log: {DEFAULT_LOG_PATH}")


if __name__ == "__main__":
    main()

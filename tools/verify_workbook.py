"""Standalone CLI that runs the spec-driven verifier against a production workbook.

Wraps ``lambda_catalog.deep_verify.verify_test_sheets`` behind a small CLI that mirrors
``build_production --verify`` but takes only a workbook path (and optional
CSV). Useful for re-verifying a previously-built workbook without going
through a full ``build_production`` cycle — for example, after manually
editing a sheet to confirm the spec oracle still holds.

Usage:
    python tools/verify_workbook.py Lambda_Library.xlsx
    python tools/verify_workbook.py Lambda_Library.xlsx --csv path/to/data.csv
    python tools/verify_workbook.py Lambda_Library.xlsx --json
    python tools/verify_workbook.py Lambda_Library.xlsx --skip-regression
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_regression_spec import build_regression_spec_qc_configs
from lambda_catalog.deep_verify import verify_test_sheets
from lambda_catalog.verify_report import (
    VerifyReport,
    render_human,
    render_json,
    report_from_failures,
)
from lambda_catalog.workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    raise_excel_access_error,
)
from lambda_catalog.write_sheet_csv_dataset import LIFE_EXPECTANCY, MILEAGE

DEFAULT_CSV_PATH = LIFE_EXPECTANCY.default_csv_path
DEFAULT_MILEAGE_CSV_PATH = MILEAGE.default_csv_path


def verify_workbook(
    workbook_path: Path,
    csv_path: Path,
    *,
    mileage_path: Path = DEFAULT_MILEAGE_CSV_PATH,
    verbose: bool = False,
    skip_regression: bool = False,
) -> VerifyReport:
    """Run the spec-driven verifier against ``workbook_path`` and return a report.

    On success, the returned ``VerifyReport`` has ``passed=True`` and an empty
    ``failures`` tuple. On drift, the report carries the structured failure
    list and the function returns normally; the caller is responsible for
    deciding whether to ``sys.exit(1)``.
    """
    start = time.monotonic()
    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        try:
            app = xw.App(visible=False, add_book=False)
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
        except OPEN_WORKBOOK_ERRORS as exc:
            raise_excel_access_error(workbook_path, "open", exc)
        try:
            captured: list[str] = []
            try:
                regression_sheet_configs = (
                    None
                    if skip_regression
                    else build_regression_spec_qc_configs(mileage_path)
                )
                verify_test_sheets(
                    workbook,
                    regression_sheet_configs,
                    csv_path,
                    verbose=verbose,
                    skip_regression=skip_regression,
                    failures_out=captured,
                )
            except RuntimeError as exc:
                if "QC verification failed" not in str(exc):
                    raise
                elapsed = time.monotonic() - start
                return report_from_failures(
                    captured,
                    elapsed_seconds=elapsed,
                    mode="spec",
                    workbook=str(workbook_path),
                )
        finally:
            if workbook is not None:
                try:
                    workbook.close()
                except OPEN_WORKBOOK_ERRORS:
                    pass
    finally:
        if app is not None:
            try:
                app.quit()
            except OPEN_WORKBOOK_ERRORS:
                pass

    elapsed = time.monotonic() - start
    return VerifyReport(
        passed=True,
        categories={},
        failures=(),
        elapsed_seconds=elapsed,
        mode="spec",
        workbook=str(workbook_path),
    )


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the standalone verifier."""
    parser = argparse.ArgumentParser(
        description=(
            "Run the spec-driven verifier against a workbook and print a "
            "structured pass/fail report."
        )
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the Excel workbook to verify.",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_CSV_PATH,
        help=(
            "Path to the Life Expectancy CSV used for the 'Developed Country "
            "after 2013' derived-column comparison."
        ),
    )
    parser.add_argument(
        "--mileage",
        type=Path,
        default=DEFAULT_MILEAGE_CSV_PATH,
        help=(
            "Path to the Auto MPG sample CSV used for the Regression sheet's "
            "spec-driven QC oracle (the Regression sheet's Source_Table "
            "default)."
        ),
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="as_json",
        help="Emit the report as JSON (for agentic consumption) instead of the human-readable form.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print per-phase checkpoints from the spec-driven verifier.",
    )
    parser.add_argument(
        "--skip-regression",
        action="store_true",
        help=(
            "Skip every Regression-sheet check. Use this to verify only the "
            "Univariate side of the unified workbook; the Life Expectancy "
            "Data 'Developed Country after 2013' check and the Univariate "
            "check still run."
        ),
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Error: workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)
    report = verify_workbook(
        workbook_path=workbook_path,
        csv_path=args.csv,
        mileage_path=args.mileage,
        verbose=args.verbose,
        skip_regression=args.skip_regression,
    )
    if args.as_json:
        print(render_json(report))
    else:
        print(render_human(report))
    if not report.passed:
        sys.exit(1)


if __name__ == "__main__":
    main()

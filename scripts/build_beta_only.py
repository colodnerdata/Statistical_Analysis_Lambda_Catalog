"""Build a minimal workbook for Beta distribution grid search performance testing.

The "Beta-only" build is a bare-bones shim: catalog, CSV data, Univariate, save.
No tab colors, no version history, no sheet reordering — this exists only to
isolate the Beta two-stage Data Table for timing the alternative ``beta_grid_size``
choices. See ``scripts/test_beta_grid_performance.py`` for the consumer.
"""

from __future__ import annotations

import argparse
import time
from pathlib import Path
from typing import Callable

import xlwings as xw

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import sync_workbook_names
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    load_csv_rows,
    write_csv_dataset_sheet,
)
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_BETA_ONLY_WORKBOOK_PATH = ROOT_DIR / "dist" / "Lambda_Library_BetaOnly.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


def build_beta_only_workbook(
    workbook_path: Path = DEFAULT_BETA_ONLY_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = LIFE_EXPECTANCY.default_csv_path,
    beta_grid_size: int = 20,
    recalculate: bool = True,
    calculate: bool = True,
    write_univariate: Callable | None = None,
) -> None:
    """Build the Beta-only workbook for performance testing.

    Parameters
    ----------
    workbook_path : Path, optional
        Path to the workbook to create or update.
    definitions_path : Path, optional
        Path to the JSON catalog file.
    csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    beta_grid_size : int, optional
        Size of the Beta distribution grid search (default: 20).
    recalculate : bool, optional
        If True (default), recalculates and saves as the final build step.
    calculate : bool, optional
        If True (default), the workbook is switched to Automatic before save.
    write_univariate : Callable, optional
        Univariate writer to call. Defaults to the standard Data-Table writer
        ``write_univariate_sheet``. Pass ``write_univariate_sheet_cartesian``
        to compare the Cartesian product approach.
    """
    if write_univariate is None:
        from lambda_catalog.write_sheet_univariate import write_univariate_sheet
        write_univariate = write_univariate_sheet

    start_time = time.monotonic()

    document = load_catalog_document(definitions_path)
    csv_headers, csv_rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)
    workbook_path = workbook_path.resolve()

    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False

        # Always start from a fresh workbook so we don't try to delete the
        # last visible sheet (Excel refuses once only one remains).
        workbook = app.books.add()

        write_catalog_sheet(workbook, document.functions)
        write_csv_dataset_sheet(workbook, csv_headers, csv_rows, LIFE_EXPECTANCY)
        write_univariate(
            workbook,
            document.univariate_sheet_notes,
            beta_grid_size=beta_grid_size,
        )

        if recalculate and calculate:
            app.api.Calculation = xw.constants.Calculation.xlCalculationAutomatic
        elif not calculate:
            # Keep in Manual mode, don't calculate
            pass
        else:
            # Recalculate but don't set Automatic
            app.api.Calculation = xw.constants.Calculation.xlCalculationManual
            app.api.CalculateFullRebuild()
        workbook.save(str(workbook_path))
    finally:
        if workbook:
            workbook.close()
        if app:
            app.quit()

    # Register workbook-scoped catalog LAMBDAs (NLL_Beta, Grid_Argument_Minimum,
    # etc.) by patching workbook.xml. This is what makes the formulas written
    # by the Univariate sheet (which reference UV_Data / NLL_Beta) actually
    # resolve when Excel opens the workbook.
    sync_workbook_names(workbook_path, document.workbook_functions)

    elapsed = time.monotonic() - start_time
    print(f"Beta-only workbook built in {elapsed:.1f}s")
    print(f"Workbook saved to: {workbook_path}")
    print(f"Beta grid size: {beta_grid_size}")


def parse_args(
    default_workbook: Path = DEFAULT_BETA_ONLY_WORKBOOK_PATH,
    description: str = "Build a minimal workbook for Beta grid search performance testing.",
) -> argparse.Namespace:
    """Parse command-line arguments for Beta-only workbook generation."""
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=default_workbook,
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
        default=LIFE_EXPECTANCY.default_csv_path,
        help="Path to the Life Expectancy CSV data file.",
    )
    parser.add_argument(
        "--beta-grid-size",
        type=int,
        default=20,
        help="Size of the Beta distribution grid search (default: 20).",
    )
    parser.add_argument(
        "--skip-recalculation",
        action="store_true",
        help="Skip the final Excel recalculation step (for build-time-only measurement).",
    )
    parser.add_argument(
        "--no-calculation",
        action="store_true",
        help="Never calculate. Excel stays in Manual mode for the whole run.",
    )
    return parser.parse_args()


def main() -> None:
    """Build the Beta-only workbook."""
    args = parse_args()
    build_beta_only_workbook(
        workbook_path=args.workbook,
        definitions_path=args.definitions,
        csv_path=args.csv,
        beta_grid_size=args.beta_grid_size,
        recalculate=not args.skip_recalculation,
        calculate=not args.no_calculation,
    )


if __name__ == "__main__":
    main()

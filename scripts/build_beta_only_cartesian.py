"""Build a Beta-only workbook using the Cartesian product writer (no Data Table).

Thin wrapper around ``scripts/build_beta_only.py`` that swaps in
``write_univariate_sheet_cartesian`` for the standard Data-Table writer. The
Cartesian product variant writes Stage 1 only (Stage 2 is not yet implemented)
and exists for performance comparison.
"""

from __future__ import annotations

from pathlib import Path

from lambda_catalog.write_sheet_univariate_cartesian import write_univariate_sheet_cartesian
from scripts.build_beta_only import build_beta_only_workbook, parse_args

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CARTESIAN_WORKBOOK_PATH = ROOT_DIR / "dist" / "Lambda_Library_BetaOnly_Cartesian.xlsx"


def main() -> None:
    """Build the Beta-only workbook with the Cartesian product writer."""
    args = parse_args(
        default_workbook=DEFAULT_CARTESIAN_WORKBOOK_PATH,
        description=(
            "Build a minimal workbook for Beta grid search performance testing "
            "using the Cartesian product approach (no Data Tables)."
        ),
    )
    build_beta_only_workbook(
        workbook_path=args.workbook,
        definitions_path=args.definitions,
        csv_path=args.csv,
        beta_grid_size=args.beta_grid_size,
        recalculate=not args.skip_recalculation,
        calculate=not args.no_calculation,
        write_univariate=write_univariate_sheet_cartesian,
    )


if __name__ == "__main__":
    main()

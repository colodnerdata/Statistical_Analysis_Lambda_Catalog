"""Re-apply the workbook-scope name sync to a built artifact, without Excel.

``sync_workbook_names`` is pure zipfile + lxml, so the workbook-scope cleanup a
normal build performs can also be applied to an already-built .xlsx on a
machine with no Excel — a Linux CI box, or a checkout where rebuilding the
artifact is not worth a full Excel run. It rewrites workbook.xml only: the
catalog's workbook-scoped LAMBDA names are replaced from
``lambda_functions.json``, every other workbook-scoped entry is dropped as
residue, and orphaned external links go with them. Sheet-scoped names, cell
values, charts and formatting are untouched.

This is a maintenance tool, not a build step. The builds
(``build_production.py`` / ``build_univariate.py``) already call
``sync_workbook_names`` and are the normal way to produce a clean artifact;
reach for this when the artifact is otherwise current and only its name manager
is stale.

Usage:
    python tools/resync_workbook_names.py Lambda_Library.xlsx
    python tools/resync_workbook_names.py Lambda_Library.xlsx Lambda_Library_Univariate.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import sync_workbook_names


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


def parse_args() -> argparse.Namespace:
    """Parse the workbook paths and the catalog path to sync from."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "workbooks",
        type=Path,
        nargs="+",
        help="Built .xlsx artifacts to re-sync in place.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda definitions.",
    )
    return parser.parse_args()


def main() -> None:
    """Re-sync each named workbook and print what changed."""
    args = parse_args()
    document = load_catalog_document(args.definitions)

    missing = [path for path in args.workbooks if not path.exists()]
    if missing:
        print("Not found: " + ", ".join(str(path) for path in missing), file=sys.stderr)
        sys.exit(1)

    for workbook_path in args.workbooks:
        result = sync_workbook_names(workbook_path.resolve(), document.workbook_functions)
        print(f"{workbook_path.name}:")
        print(f"  Created names: {result.created}")
        print(f"  Updated names: {result.updated}")
        print(f"  Removed names: {result.removed}")
        if result.skipped:
            print(
                "  Skipped names: "
                + ", ".join(result.skipped)
                + " (reference a worksheet this workbook does not have)"
            )


if __name__ == "__main__":
    main()

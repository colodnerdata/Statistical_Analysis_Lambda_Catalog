"""Build Lambda_Library_TestModels.xlsx — one worksheet per test-model case.

A third build target, alongside ``build_production.py`` (the Regression
artifact) and ``build_univariate.py`` (the Univariate artifact). Unlike
those two this one is **not shipped**: it is a QC fixture, gitignored like
``Lambda_Library_QC.xlsx``, and its only audience is whoever is changing the
spec block or the engine.

What it produces is the regression test-model suite made physical. Every
``RegressionSpecCase`` in ``docs/MODEL_TESTING_ASSETS.md`` § 1.1–1.3 and
every ``GuardStateCase`` in § 1.4 gets its own Regression-shaped sheet, named
for the corner it covers (``M09 Cat x Cat Full Product``, ``G10 Symmetric
Product Pair``), pre-populated and fully calculated. Verification then reads
those sheets and writes nothing at all, which is the whole point: the legacy
harness has to mutate one shared sheet ~33 times, so a failure leaves nothing
behind to look at and no case can be inspected independently of the ones
before it.

Usage::

    python build_test_models.py                       # every non-heavy case
    python build_test_models.py --include-heavy        # + L07 / L08
    python build_test_models.py --cases M09,G10        # just those two
    python build_test_models.py --verify --no-launch   # build, check, exit 1 on drift
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_regression_guard_states import (
    build_guard_state_cases,
    calculate_guard_state_case,
)
from lambda_catalog.analyze_regression_spec import (
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.build_common import print_name_sync_summary
from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    raise_excel_access_error,
)
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    MILEAGE,
    PRODUCTION_LOTS,
    load_csv_rows,
    write_csv_dataset_sheet,
)
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet
from lambda_catalog.write_sheet_test_model import (
    write_guard_state_sheet,
    write_test_model_sheet,
)

ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library_TestModels.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"

# Fixture columns some cases declare as Filters (Is_USA for M15). The legacy
# verifier adds these to MileageData and deletes them again around each case,
# because it runs against a shipped artifact that must not keep them. This
# workbook is itself a fixture, so they are written once and left in place —
# no add/delete dance, and no case order dependence.
_FIXTURE_COLUMN_FORMULAS = {
    MILEAGE.table_name: {
        "Is_USA": '=--([@Origin]="US")',
    },
}


def _selected_cases(
    case_filter: set[str] | None, include_heavy: bool
) -> tuple[list, list]:
    """Return (fittable cases, guard cases) after --cases / --include-heavy.

    ``case_filter`` matches a plan ID (``"M09"``) or a case name
    (``"interaction_categorical_cross"``), case-insensitively, so either
    identifier a developer has to hand works. An explicit ``--cases``
    selection overrides ``heavy``: naming L07 means you want L07.
    """
    model_cases = build_regression_spec_cases()
    guard_cases = build_guard_state_cases()

    if case_filter is not None:
        wanted = {token.casefold() for token in case_filter}

        def _matches(case) -> bool:
            return (
                case.plan_id.casefold() in wanted
                or case.name.casefold() in wanted
            )

        model_cases = [case for case in model_cases if _matches(case)]
        guard_cases = [case for case in guard_cases if _matches(case)]
        unmatched = wanted - {
            token
            for case in (*model_cases, *guard_cases)
            for token in (case.plan_id.casefold(), case.name.casefold())
        }
        if unmatched:
            raise ValueError(
                "No test-model case matches: " + ", ".join(sorted(unmatched))
            )
    elif not include_heavy:
        model_cases = [case for case in model_cases if not case.heavy]

    return model_cases, guard_cases


def _write_fixture_columns(workbook: xw.Book) -> None:
    """Add the fixture columns some specs declare, once, permanently."""
    for config in (MILEAGE, LIFE_EXPECTANCY, PRODUCTION_LOTS):
        columns = _FIXTURE_COLUMN_FORMULAS.get(config.table_name)
        if not columns:
            continue
        sheet = workbook.sheets[config.sheet_name]
        table = sheet.api.ListObjects(config.table_name)
        existing = {
            table.ListColumns(index).Name
            for index in range(1, table.ListColumns.Count + 1)
        }
        for name, formula in columns.items():
            if name in existing:
                continue
            column = table.ListColumns.Add()
            column.Name = name
            column.DataBodyRange.Formula = formula


def build_test_models_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    *,
    case_filter: set[str] | None = None,
    include_heavy: bool = False,
    verbose: bool = False,
) -> tuple[NameSyncResult, list[str]]:
    """Write every selected case's sheet and sync the catalog names.

    Returns the name-sync result and the ordered list of sheet names built,
    so a caller can verify exactly what was produced rather than re-deriving
    the selection.
    """
    document = load_catalog_document(definitions_path)
    closures = document.functions_for_sheet("Regression")
    model_cases, guard_cases = _selected_cases(case_filter, include_heavy)

    if verbose:
        print(
            f"Building {len(model_cases)} model sheet(s) and "
            f"{len(guard_cases)} guard sheet(s)",
            flush=True,
        )

    # Compute every oracle BEFORE opening Excel. A spec error should fail in
    # seconds, not after a multi-minute build has already written 30 sheets.
    model_expectations = [
        calculate_regression_spec_case(case) for case in model_cases
    ]
    guard_expectations = [calculate_guard_state_case(case) for case in guard_cases]

    headers = {
        config.name: load_csv_rows(config.default_csv_path, config)
        for config in (LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS)
    }

    workbook_path = workbook_path.resolve()
    built: list[str] = []
    start = time.monotonic()
    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            # Always from scratch. Sheet names encode the case list, so an
            # incremental build over a previous run would silently keep
            # sheets for cases that have since been renamed or dropped —
            # and a stale sheet that still calculates is the worst possible
            # QC artifact.
            workbook = app.books.add()
            try:
                app.api.Calculation = XL_CALCULATION_MANUAL
                for sheet in list(workbook.sheets)[1:]:
                    sheet.delete()
                workbook.sheets[0].name = "LAMBDA_functions"

                write_catalog_sheet(workbook, document.functions)
                for config in (LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS):
                    config_headers, config_rows = headers[config.name]
                    write_csv_dataset_sheet(
                        workbook, config_headers, config_rows, config
                    )
                _write_fixture_columns(workbook)

                for expected in model_expectations:
                    if verbose:
                        print(f"  {expected.case.sheet_name}", flush=True)
                    write_test_model_sheet(
                        workbook, expected, document.regression_sheet_notes, closures
                    )
                    built.append(expected.case.sheet_name)
                for guard in guard_expectations:
                    if verbose:
                        print(f"  {guard.case.sheet_name}", flush=True)
                    write_guard_state_sheet(
                        workbook, guard, document.regression_sheet_notes, closures
                    )
                    built.append(guard.case.sheet_name)

                app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)

    if verbose:
        print(f"  Write sheets:   {time.monotonic() - start:.1f}s", flush=True)

    result = sync_workbook_names(workbook_path, document.workbook_functions)

    # The rebuild is not optional here, for the same reason CLAUDE.md gives
    # for the Regression artifact: the name sync rewrites the catalog's
    # workbook-scoped LAMBDAs, and a per-sheet Calculate() does not rebuild
    # the dependency tree behind them, so every engine value would read nan.
    # This artifact has no Data Tables, so the rebuild is cheap.
    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
            try:
                app.api.CalculateFullRebuild()
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "recalculate", exc)

    return result, built


def _run_verify(workbook_path: Path, built: list[str]) -> int:
    """Verify the built workbook; return the process exit code."""
    from tools.inspect_test_model_sheets import verify_test_model_workbook

    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
            try:
                failures = verify_test_model_workbook(workbook, built)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "verify", exc)

    if failures:
        for message in failures[:200]:
            print(f"ERROR {message}", flush=True)
        if len(failures) > 200:
            print(
                f"ERROR ... and {len(failures) - 200} more mismatch(es).",
                flush=True,
            )
        print(f"Verify: FAILED with {len(failures)} mismatch(es).", flush=True)
        return 1
    print("Verify: passed", flush=True)
    return 0


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the test-model workbook build."""
    parser = argparse.ArgumentParser(
        description="Build Lambda_Library_TestModels.xlsx — one sheet per test model."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK_PATH,
        help="Path to the test-model workbook to create.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda definitions.",
    )
    parser.add_argument(
        "--cases",
        type=str,
        default=None,
        help="Comma-separated plan IDs or case names to build (e.g. 'M09,G10'). "
        "Overrides --include-heavy for the cases named.",
    )
    parser.add_argument(
        "--include-heavy",
        action="store_true",
        help="Also build the cases marked heavy (L07's ~2900x205 design "
        "matrix, L08's 193 Fixed Effects groups). Their Python oracles run "
        "in the unit suite either way; only the sheets are gated.",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="After building, read every sheet back and compare against the "
        "oracle. Exits 1 on drift.",
    )
    parser.add_argument(
        "--no-launch",
        action="store_true",
        help="Do not open the workbook in Excel when the build finishes.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print each sheet as it is written, plus phase timings.",
    )
    return parser.parse_args()


def main() -> None:
    """Build the test-model workbook, optionally verifying it."""
    args = parse_args()
    case_filter = (
        {token.strip() for token in args.cases.split(",") if token.strip()}
        if args.cases
        else None
    )
    total_start = time.monotonic()

    result, built = build_test_models_workbook(
        workbook_path=args.workbook,
        definitions_path=args.definitions,
        case_filter=case_filter,
        include_heavy=args.include_heavy,
        verbose=args.verbose,
    )

    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheets built: {len(built)}")
    print_name_sync_summary(result)
    print(f"Timing: total         {time.monotonic() - total_start:.1f}s")

    exit_code = 0
    if args.verify:
        exit_code = _run_verify(args.workbook.resolve(), built)

    if not args.no_launch and exit_code == 0:
        subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])

    sys.exit(exit_code)


if __name__ == "__main__":
    main()

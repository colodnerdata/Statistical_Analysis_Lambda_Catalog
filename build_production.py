"""Build Lambda_Library.xlsx with all production sheets and LAMBDA name sync."""
from __future__ import annotations

import argparse
import importlib.util
import subprocess
import sys
import time
from collections.abc import Callable
from pathlib import Path

import xlwings as xw

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.verify_report import (
    VerifyReport,
    report_from_failures,
    render_human,
)
from lambda_catalog.workbook_builder import (
    NameSyncResult,
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
from lambda_catalog.write_sheet_mileage_data import (
    DEFAULT_XLSX_PATH as DEFAULT_MILEAGE_XLSX_PATH,
    load_mileage_rows,
    write_mileage_sheet,
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
_RPC_FAILURE_PHRASES = (
    "remote procedure call failed",
    "rpc server is unavailable",
)


def _is_rpc_failure(exc: BaseException) -> bool:
    """Return True when Excel's COM server disappeared mid-call."""
    message = str(exc).lower()
    return any(phrase in message for phrase in _RPC_FAILURE_PHRASES)


def _close_workbook_quietly(workbook: xw.Book | None) -> None:
    """Best-effort close; Excel may already be gone after an RPC failure."""
    if workbook is None:
        return
    try:
        workbook.close()
    except OPEN_WORKBOOK_ERRORS:
        pass


def _quit_app_quietly(app: xw.App | None) -> None:
    """Best-effort quit; cleanup must not mask the original Excel error."""
    if app is None:
        return
    try:
        app.quit()
    except OPEN_WORKBOOK_ERRORS:
        pass


def _recalculate_and_save(workbook_path: Path) -> None:
    """Fully calculate Data Tables after name sync, then save semiautomatic mode."""
    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False

        workbook = app.books.open(str(workbook_path))
        try:
            app.api.Calculation = XL_CALCULATION_MANUAL
            app.api.CalculateFullRebuild()
            app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
            workbook.save(str(workbook_path))
        finally:
            _close_workbook_quietly(workbook)
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "recalculate and save", exc)
    finally:
        _quit_app_quietly(app)


def _load_build_qc_module() -> object:
    """Import build_qc.py from the repo root without requiring it to be a package.

    build_qc.py is a top-level script in the repo (not under lambda_catalog/),
    so a normal ``import build_qc`` would rely on the consumer adding the
    repo root to ``sys.path``. Loading it explicitly here keeps the verify
    path self-contained when build_production is invoked as a script.
    """
    spec = importlib.util.spec_from_file_location("build_qc", ROOT_DIR / "build_qc.py")
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load build_qc.py from {ROOT_DIR}")
    module = importlib.util.module_from_spec(spec)
    sys.modules.setdefault("build_qc", module)
    spec.loader.exec_module(module)
    return module


def _run_deep_verify(
    workbook_path: Path,
    csv_path: Path,
    *,
    mileage_path: Path = DEFAULT_MILEAGE_XLSX_PATH,
    verbose: bool = False,
    skip_univariate: bool = False,
) -> VerifyReport:
    """Run the spec-driven verifier against the production workbook.

    Opens the workbook in a headless Excel instance, computes the per-config
    Python oracle, and calls ``build_qc.verify_test_sheets(..., skip_dummy=True,
    skip_univariate=skip_univariate, failures_out=...)``. On success, returns
    a passing ``VerifyReport``. On drift, ``verify_test_sheets`` raises
    ``RuntimeError("QC verification failed with N mismatch(es).")``;
    ``failures_out`` is populated before the raise so we can return a
    structured ``VerifyReport`` for the caller to render and ``sys.exit(1)``
    without unwinding the build pipeline. Other exceptions propagate.

    ``skip_univariate`` should mirror the ``--skip-univariate`` flag passed
    to the build phase: when the Univariate sheet was never written (or a
    prior build already omitted it), ``verify_test_sheets`` skips its
    Univariate checks with a warning instead of crashing.
    """
    start = time.monotonic()
    build_qc = _load_build_qc_module()
    if verbose:
        print("  Verify:         spec-driven verifier starting", flush=True)

    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False
        workbook = app.books.open(str(workbook_path))
        try:
            captured: list[str] = []
            try:
                build_qc.verify_test_sheets(
                    workbook,
                    build_qc.build_regression_spec_qc_configs(csv_path),
                    csv_path,
                    mileage_path=mileage_path,
                    verbose=verbose,
                    skip_dummy=True,
                    skip_univariate=skip_univariate,
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
            _close_workbook_quietly(workbook)
    finally:
        _quit_app_quietly(app)

    elapsed = time.monotonic() - start
    return VerifyReport(
        passed=True,
        categories={},
        failures=(),
        elapsed_seconds=elapsed,
        mode="spec",
        workbook=str(workbook_path),
    )


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
    mileage_xlsx_path: Path = DEFAULT_MILEAGE_XLSX_PATH,
    validate_reopen: bool = False,
    verbose: bool = False,
    recalculate: bool = True,
    skip_univariate: bool = False,
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
    mileage_xlsx_path : Path, optional
        Path to the Auto MPG sample xlsx file written to the Mileage Data
        sheet — a second sample dataset for practicing the Source_Table
        retarget workflow alongside Life Expectancy Data.
    validate_reopen : bool, optional
        If True, reopens the workbook in Excel after patching to verify it.
    verbose : bool, optional
        If True, prints timing information for each build phase to stdout.
    recalculate : bool, optional
        If True (default), recalculates and saves as the final build step.
        Pass False when the caller manages the recalculate step separately
        (e.g. to allow a targeted retry without re-running the full build).
    skip_univariate : bool, optional
        If True, skips writing the Univariate Analysis sheet. An existing
        Univariate Analysis sheet is left untouched; on a from-scratch build
        the sheet is simply absent.

    Returns
    -------
    NameSyncResult
        Counts of created versus updated workbook names.
    """
    _t = time.monotonic()
    document = load_catalog_document(definitions_path)
    csv_headers, csv_rows = load_life_expectancy_rows(csv_path)
    mileage_headers, mileage_rows = load_mileage_rows(mileage_xlsx_path)
    if verbose:
        print(f"  Prep:           {time.monotonic() - _t:.1f}s", flush=True)

    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    _t = time.monotonic()
    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        app = xw.App(visible=True, add_book=False)
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
            # v3.0 changeover: the spec block moved onto the Regression sheet,
            # so a carried-forward Model Construction sheet is stale and gets
            # dropped.
            _delete_sheet_if_present(workbook, "Model Construction")
            for qc_sheet in _QC_SHEET_NAMES:
                _delete_sheet_if_present(workbook, qc_sheet)
            if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                workbook.sheets["Sheet1"].name = "LAMBDA_functions"
            write_catalog_sheet(workbook, document.functions)
            write_life_expectancy_sheet(workbook, csv_headers, csv_rows)
            write_mileage_sheet(workbook, mileage_headers, mileage_rows)
            if not skip_univariate:
                write_univariate_sheet(workbook)
            write_regression_instructions_sheet(workbook)
            write_diagnostic_guide_sheet(workbook)
            write_version_history_sheet(workbook)
            write_regression_output_sheet(
                workbook,
                document.regression_sheet_notes,
                document.functions_for_sheet("Regression"),
            )
            app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
            workbook.save(str(workbook_path))
        finally:
            _close_workbook_quietly(workbook)
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)
    finally:
        _quit_app_quietly(app)
    if verbose:
        print(f"  Write sheets:   {time.monotonic() - _t:.1f}s", flush=True)

    _t = time.monotonic()
    try:
        result = sync_workbook_names(workbook_path, document.workbook_functions)
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    except (PermissionError, OSError) as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    if verbose:
        print(f"  Sync names:     {time.monotonic() - _t:.1f}s", flush=True)

    if recalculate:
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
        verbose, skip_univariate, and skip_data_table_calculations attributes.
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
    parser.add_argument(
        "--skip-univariate",
        action="store_true",
        help=(
            "Skip writing the Univariate Analysis sheet to speed up iteration "
            "on other sheets (e.g. Regression). An existing Univariate "
            "Analysis sheet is left as-is; a from-scratch build omits it."
        ),
    )
    parser.add_argument(
        "--skip-data-table-calculations",
        action="store_true",
        help=(
            "Skip the final Excel CalculateFullRebuild phase that evaluates "
            "Data Tables. The workbook is still written, names are synced, "
            "and formulas/Data Tables can calculate later in Excel."
        ),
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        default=False,
        help=(
            "After the build, run the spec-driven verifier (build_qc."
            "verify_test_sheets) against the production sheets. On any "
            "drift, print a structured VerifyReport and sys.exit(1). The "
            "post-build Excel handoff (cmd /c start) only fires when "
            "verify passes. Combine with --no-launch to run verify without "
            "opening Excel at all."
        ),
    )
    parser.add_argument(
        "--no-verify",
        action="store_false",
        dest="verify",
        help=(
            "Explicitly disable the spec-driven verifier pass (mainly useful "
            "for wrapper scripts that default to --verify)."
        ),
    )
    parser.add_argument(
        "--no-launch",
        action="store_true",
        help=(
            "Suppress the post-build 'cmd /c start <workbook>' Excel handoff. "
            "Useful for agentic loops where the human is not in the QA loop "
            "and Excel should not pop up."
        ),
    )
    return parser.parse_args()


def _retry_on_open(
    action_label: str,
    fn: Callable[[], None],
    *,
    retry_rpc: bool = False,
    max_rpc_retries: int = 1,
) -> None:
    """Call fn(); retry open-workbook locks and optionally dropped COM sessions."""
    rpc_retries = 0
    while True:
        try:
            fn()
            return
        except RuntimeError as exc:
            if retry_rpc and _is_rpc_failure(exc) and rpc_retries < max_rpc_retries:
                rpc_retries += 1
                print(
                    "\nExcel's COM session dropped; "
                    "retrying in a fresh Excel instance...",
                    file=sys.stderr,
                    flush=True,
                )
                if sys.stdin.isatty():
                    time.sleep(2)
                continue
            if "likely open in Excel" not in str(exc):
                raise
            if not sys.stdin.isatty():
                raise
            input(
                f"\n{action_label} — close it in Excel "
                "and press Enter to retry (or Ctrl+C to cancel): "
            )


def main() -> None:
    """Build the production workbook and print a short sync summary for interactive use."""
    args = parse_args()
    workbook_path = args.workbook.resolve()
    total_start = time.monotonic()
    verify_elapsed: float | None = None

    # Phase 1: write all sheets + sync names + inject charts.
    # If the workbook is open in Excel at this point the entire write must be retried.
    result: NameSyncResult | None = None

    def _run_build() -> None:
        nonlocal result
        result = build_production_workbook(
            workbook_path=workbook_path,
            definitions_path=args.definitions,
            csv_path=args.csv,
            validate_reopen=False,  # handled below after recalculate
            verbose=args.verbose,
            recalculate=False,      # handled separately so only this step retries
            skip_univariate=args.skip_univariate,
        )

    build_phase_start = time.monotonic()
    _retry_on_open(
        f"{args.workbook.name} is open in Excel",
        _run_build,
        retry_rpc=True,
    )
    build_elapsed = time.monotonic() - build_phase_start
    assert result is not None

    # Phase 2: recalculate Data Tables and save.
    # This is a quick step; if the workbook is open in Excel now (e.g. the user
    # opened it to inspect progress), only this step is retried — not the full build.
    if args.skip_data_table_calculations:
        if args.verbose:
            print("  Recalculate:    skipped", flush=True)
        recalc_elapsed = None
    else:
        _t = time.monotonic()
        _retry_on_open(
            f"{args.workbook.name} is open in Excel",
            lambda: _recalculate_and_save(workbook_path),
            retry_rpc=True,
        )
        recalc_elapsed = time.monotonic() - _t
        if args.verbose:
            print(f"  Recalculate:    {recalc_elapsed:.1f}s", flush=True)

    if args.validate_reopen:
        _validate_workbook_reopen(workbook_path)

    print(f"Workbook: {workbook_path}")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print("Sheet updated: Mileage Data")
    if args.skip_univariate:
        print("Sheet skipped: Univariate Analysis")
    else:
        print("Sheet updated: Univariate Analysis")
    print("Sheet updated: Regression Instructions")
    print("Sheet updated: Diagnostic Guide")
    print("Sheet updated: Version History")
    print("Sheet updated: Regression")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    if args.validate_reopen:
        print("Reopen validation: passed")

    if args.verify:
        verify_start = time.monotonic()
        # The verify-success branch is the only path that opens Excel. On
        # verify failure we sys.exit(1) and never shell out to `cmd /c start`
        # so a stale build cannot be launched in place of a fresh one.
        report = _run_deep_verify(
            workbook_path,
            args.csv,
            verbose=args.verbose,
            skip_univariate=args.skip_univariate,
        )
        verify_elapsed = time.monotonic() - verify_start
        print(render_human(report))
        print(f"Timing: build+sync    {build_elapsed:.1f}s")
        if recalc_elapsed is None:
            print("Timing: recalculate   skipped")
        else:
            print(f"Timing: recalculate   {recalc_elapsed:.1f}s")
        print(f"Timing: verify        {verify_elapsed:.1f}s")
        print(f"Timing: total         {time.monotonic() - total_start:.1f}s")
        if not report.passed:
            sys.exit(1)
    else:
        print(f"Timing: build+sync    {build_elapsed:.1f}s")
        if recalc_elapsed is None:
            print("Timing: recalculate   skipped")
        else:
            print(f"Timing: recalculate   {recalc_elapsed:.1f}s")
        print(f"Timing: total         {time.monotonic() - total_start:.1f}s")

    if not args.no_launch:
        subprocess.Popen(["cmd", "/c", "start", "", str(workbook_path)])


if __name__ == "__main__":
    main()

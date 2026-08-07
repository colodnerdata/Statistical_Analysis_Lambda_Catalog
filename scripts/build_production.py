"""Build Lambda_Library.xlsx — the Regression workbook — with its production sheets and LAMBDA name sync.

From v3.0 the build emits two artifacts: this script builds the Regression
workbook (``Lambda_Library.xlsx``), and ``build_univariate.py`` builds the
standalone Univariate workbook (``Lambda_Library_Univariate.xlsx``). The
Univariate Analysis sheet ships in its own workbook, so each artifact sets
its own calculation mode and the Regression workbook runs in full Automatic
(see DECISIONS.md § v3.0 "Univariate becomes its own workbook").
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_regression_spec import build_regression_spec_qc_configs
from lambda_catalog.build_common import (
    RUN_LOG_DIR_NAME,
    RUN_LOG_FILE_SUFFIX,
    _close_workbook_quietly,
    _quit_app_quietly,
    _recalculate_and_save,
    _retry_on_open,
    print_name_sync_summary,
    run_log_path,
    tee_run_log,
    warn_if_workbook_open,
)
from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.deep_verify import verify_test_sheets
from lambda_catalog.sheet_styles import SUBHDR_COLOR
from lambda_catalog.verify_report import (
    VerifyReport,
    render_human,
    report_from_failures,
)
from lambda_catalog.workbook_builder import (
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    NameSyncResult,
    _validate_workbook_reopen,
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
from lambda_catalog.write_sheet_diagnostic_guide import write_diagnostic_guide_sheet
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet
from lambda_catalog.write_spec_block import SPEC_DATASET_PROFILES
from lambda_catalog.write_sheet_regression import write_regression_output_sheet
from lambda_catalog.write_sheet_regression_instructions import (
    write_regression_instructions_sheet,
)
from lambda_catalog.write_sheet_version_history import write_version_history_sheet

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "dist" / "Lambda_Library.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
_TAB_COLOR_LIGHT_GRAY = (217, 217, 217)
_TAB_COLOR_DARK_GRAY = (128, 128, 128)

_SHEET_NAME_LAMBDA_FUNCTIONS = "LAMBDA_functions"
_SHEET_NAME_MILEAGE_DATA = "Mileage Data"
_SHEET_NAME_LIFE_EXPECTANCY_DATA = "Life Expectancy Data"
_SHEET_NAME_PRODUCTION_LOTS = "Production Lots"
_SHEET_NAME_VERSION_HISTORY = "Version History"
_SHEET_NAME_REGRESSION_INSTRUCTIONS = "Regression Instructions"
_SHEET_NAME_REGRESSION = "Regression"
_SHEET_NAME_DIAGNOSTIC_GUIDE = "Diagnostic Guide"

# Dataset profiles: maps the --regression-dataset choice to both the
# Source_Table RefersTo formula and the spec block's default Role/Include/
# Type/Sequence values — SPEC_DATASET_PROFILES (write_spec_block.py)
# is the single source of truth for both, so retargeting the dataset also
# retargets the shipped default model instead of leaving every column of
# the newly-targeted dataset an un-flagged Predictor. "auto_mpg" is the
# shipped default (MileageData table on the Mileage Data sheet).
# "life_expectancy" retargets to the Life Expectancy Data sheet.
# "production_lots" retargets to the Production Lots learning-curve panel —
# the only shipped dataset with a natural Fixed Effects grouping column
# (Facility) and Sequence column (Fiscal_Year), so it is what exercises the
# Fixed Effects role end to end (see analyze_regression_spec.py's
# production_lots_fixed_effects QC case).


def _run_deep_verify(
    workbook_path: Path,
    csv_path: Path,
    *,
    mileage_path: Path = MILEAGE.default_csv_path,
    production_lots_path: Path = PRODUCTION_LOTS.default_csv_path,
    verbose: bool = False,
) -> VerifyReport:
    """Run the spec-driven verifier against the Regression production workbook.

    Opens the workbook in a headless Excel instance, computes the per-config
    Python oracle, and calls ``deep_verify.verify_test_sheets(...,
    skip_univariate=True, failures_out=...)``. On success, returns
    a passing ``VerifyReport``. On drift, ``verify_test_sheets`` raises
    ``RuntimeError("QC verification failed with N mismatch(es).")``;
    ``failures_out`` is populated before the raise so we can return a
    structured ``VerifyReport`` for the caller to render and ``sys.exit(1)``
    without unwinding the build pipeline. Other exceptions propagate.

    ``skip_univariate=True`` is hardcoded: the Regression workbook ships no
    Univariate sheet (it moved to its own artifact in v3.0; see
    DECISIONS.md § v3.0 "Univariate becomes its own workbook"). The
    verifier therefore skips its Univariate checks silently — no warning
    is printed because the absence is by design, not a build regression.
    """
    start = time.monotonic()
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
                verify_test_sheets(
                    workbook,
                    build_regression_spec_qc_configs(mileage_path),
                    csv_path,
                    mileage_path=mileage_path,
                    production_lots_path=production_lots_path,
                    verbose=verbose,
                    skip_univariate=True,
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


def _rgb_to_excel_color(rgb: tuple[int, int, int]) -> int:
    """Convert an RGB tuple to Excel's BGR-packed tab color integer."""
    r, g, b = rgb
    return r + (g * 256) + (b * 65536)


def _sheet_names(workbook: xw.Book) -> set[str]:
    """Return the current workbook sheet names."""
    return {sheet.name for sheet in workbook.sheets}


def _move_sheet_before(sheet: xw.Sheet, anchor: xw.Sheet) -> None:
    """Move one sheet before another."""
    sheet.api.Move(Before=anchor.api)


def _set_tab_color(sheet: xw.Sheet, color: tuple[int, int, int]) -> None:
    """Set an Excel sheet tab color from an RGB tuple."""
    sheet.api.Tab.Color = _rgb_to_excel_color(color)


def _reorder_and_style_sheet_tabs(workbook: xw.Book) -> None:
    """Apply build-time tab order and tab colors for the Regression workbook's sheets."""
    ordered_front = [
        _SHEET_NAME_MILEAGE_DATA,
        _SHEET_NAME_LIFE_EXPECTANCY_DATA,
        _SHEET_NAME_PRODUCTION_LOTS,
        _SHEET_NAME_VERSION_HISTORY,
        _SHEET_NAME_REGRESSION_INSTRUCTIONS,
        _SHEET_NAME_REGRESSION,
        _SHEET_NAME_DIAGNOSTIC_GUIDE,
    ]
    ordered_front.append(_SHEET_NAME_LAMBDA_FUNCTIONS)

    present = _sheet_names(workbook)
    for sheet_name in reversed(ordered_front):
        if sheet_name not in present:
            continue
        sheet = workbook.sheets[sheet_name]
        first_sheet = workbook.sheets[0]
        if sheet.name != first_sheet.name:
            _move_sheet_before(sheet, first_sheet)

    tab_colors: dict[str, tuple[int, int, int]] = {
        _SHEET_NAME_MILEAGE_DATA: _TAB_COLOR_LIGHT_GRAY,
        _SHEET_NAME_LIFE_EXPECTANCY_DATA: _TAB_COLOR_LIGHT_GRAY,
        _SHEET_NAME_PRODUCTION_LOTS: _TAB_COLOR_LIGHT_GRAY,
        _SHEET_NAME_VERSION_HISTORY: _TAB_COLOR_DARK_GRAY,
        _SHEET_NAME_REGRESSION_INSTRUCTIONS: SUBHDR_COLOR,
        _SHEET_NAME_REGRESSION: SUBHDR_COLOR,
        _SHEET_NAME_DIAGNOSTIC_GUIDE: SUBHDR_COLOR,
    }

    present = _sheet_names(workbook)
    for sheet_name, color in tab_colors.items():
        if sheet_name in present:
            _set_tab_color(workbook.sheets[sheet_name], color)


def build_production_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = LIFE_EXPECTANCY.default_csv_path,
    mileage_csv_path: Path = MILEAGE.default_csv_path,
    production_lots_csv_path: Path = PRODUCTION_LOTS.default_csv_path,
    validate_reopen: bool = False,
    verbose: bool = False,
    recalculate: bool = True,
    regression_dataset: str = "auto_mpg",
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
    mileage_csv_path : Path, optional
        Path to the Auto MPG sample CSV file written to the Mileage Data
        sheet — the dataset the Regression sheet's Source_Table targets by
        default. Life Expectancy Data ships alongside it as a second
        sample dataset for practicing the Source_Table retarget workflow.
    production_lots_csv_path : Path, optional
        Path to the Production Lots sample CSV file written to the
        Production Lots sheet — a small unbalanced panel (3 facilities, 51
        lots) with a natural Fixed Effects grouping column (Facility) and
        Sequence column (Fiscal_Year), for retargeting Source_Table to
        exercise the Fixed Effects role.
    regression_dataset : str, optional
        Regression source-table profile: ``"auto_mpg"`` (default, targets
        ``MileageData[#All]``), ``"life_expectancy"`` (targets
        ``LifeExpectancyData[#All]``), or ``"production_lots"`` (targets
        ``ProductionLotsData[#All]``).
    validate_reopen : bool, optional
        If True, reopens the workbook in Excel after patching to verify it.
    verbose : bool, optional
        If True, prints timing information for each build phase to stdout.
    recalculate : bool, optional
        If True (default), recalculates and saves as the final build step.
        Pass False when the caller manages the recalculate step separately
        (e.g. to allow a targeted retry without re-running the full build).

    Returns
    -------
    NameSyncResult
        Counts of created versus updated workbook names.
    """
    if regression_dataset not in SPEC_DATASET_PROFILES:
        raise ValueError(
            "regression_dataset must be one of: "
            + ", ".join(sorted(SPEC_DATASET_PROFILES))
        )
    _t = time.monotonic()
    document = load_catalog_document(definitions_path)
    csv_headers, csv_rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)
    mileage_headers, mileage_rows = load_csv_rows(mileage_csv_path, MILEAGE)
    production_lots_headers, production_lots_rows = load_csv_rows(
        production_lots_csv_path, PRODUCTION_LOTS
    )
    regression_spec_profile = SPEC_DATASET_PROFILES[regression_dataset]
    source_table_ref = regression_spec_profile.source_table_ref
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
            if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                workbook.sheets["Sheet1"].name = _SHEET_NAME_LAMBDA_FUNCTIONS
            write_catalog_sheet(workbook, document.functions)
            write_csv_dataset_sheet(workbook, csv_headers, csv_rows, LIFE_EXPECTANCY)
            write_csv_dataset_sheet(workbook, mileage_headers, mileage_rows, MILEAGE)
            write_csv_dataset_sheet(
                workbook, production_lots_headers, production_lots_rows, PRODUCTION_LOTS
            )
            write_regression_instructions_sheet(workbook)
            write_diagnostic_guide_sheet(workbook)
            write_version_history_sheet(workbook)
            write_regression_output_sheet(
                workbook,
                document.regression_sheet_notes,
                document.functions_for_sheet("Regression"),
                source_table_ref=source_table_ref,
                spec_profile=regression_spec_profile,
            )
            _reorder_and_style_sheet_tabs(workbook)
            app.api.Calculation = XL_CALCULATION_AUTOMATIC
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
        verbose, verify, and no_launch attributes.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Build Lambda_Library.xlsx — the Regression workbook — from "
            "lambda_functions.json. The Univariate Analysis sheet ships in "
            "its own workbook (build_univariate.py)."
        )
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
        "--verify",
        action="store_true",
        default=False,
        help=(
            "After the build, run the spec-driven verifier (deep_verify."
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
    parser.add_argument(
        "--log",
        type=Path,
        default=None,
        help=(
            "Where to archive this run's transcript (stdout + stderr, "
            "including a traceback from an aborted run). Defaults to "
            f"'{RUN_LOG_DIR_NAME}/<script> <flags>{RUN_LOG_FILE_SUFFIX}' — the "
            "same naming build_test_models.py uses, so a directory listing "
            "reads as a list of what was actually run."
        ),
    )
    parser.add_argument(
        "--regression-dataset",
        choices=tuple(SPEC_DATASET_PROFILES),
        default="auto_mpg",
        help=(
            "Which dataset the Regression sheet's Source_Table points to, "
            "and which shipped default spec (Role/Include/Type/Sequence "
            "per column) pre-fills the MODEL SPECIFICATION block. "
            "'auto_mpg' (default) targets MileageData[#All]; "
            "'life_expectancy' targets LifeExpectancyData[#All] (Response: "
            "Life expectancy; Predictors: the 18-column FEATURE_COLUMNS "
            "model); 'production_lots' targets ProductionLotsData[#All] "
            "(Facility as Fixed Effects, Fiscal_Year as Sequence, the "
            "Crawford/Wright learning-curve model)."
        ),
    )
    return parser.parse_args()


def _build_and_verify(args: argparse.Namespace, workbook_path: Path) -> int:
    """Run the build, the recalculate phase and (optionally) the verifier.

    Returns the process exit code — 1 when the spec-driven verifier reported
    drift, 0 otherwise — rather than calling ``sys.exit`` itself. ``main``
    runs this inside ``tee_run_log``, and a ``SystemExit`` raised in there
    would be archived as a traceback, burying the ``VerifyReport`` that
    actually explains the failure under a stack trace that explains nothing.
    """
    total_start = time.monotonic()
    verify_elapsed: float | None = None

    # Said up front rather than after the multi-minute write: Excel opens a
    # locked workbook read-only without raising, so the first call that fails is
    # the save at the END of phase 1 — _retry_on_open below would then discard
    # every sheet write before telling the user to close Excel.
    warn_if_workbook_open(
        workbook_path, action_label=f"{args.workbook.name} is open in Excel"
    )

    # Phase 1: write all sheets + sync names + inject charts.
    # If the workbook is open in Excel at this point the entire write must be retried.
    result: NameSyncResult | None = None

    def _run_build() -> None:
        nonlocal result
        result = build_production_workbook(
            workbook_path=workbook_path,
            definitions_path=args.definitions,
            csv_path=args.csv,
            mileage_csv_path=args.mileage_csv,
            production_lots_csv_path=args.production_lots_csv,
            validate_reopen=False,  # handled below after recalculate
            verbose=args.verbose,
            recalculate=False,  # handled separately so only this step retries
            regression_dataset=args.regression_dataset,
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
    #
    # The Regression workbook has no Data Tables, so CalculateFullRebuild is
    # cheap. It is also required: the spec-driven verifier only does a per-sheet
    # Calculate(), which does NOT rebuild the dependency tree after 100+
    # workbook names are re-synced — skipping the rebuild would leave the
    # Regression engines uncomputed and every QC value reading nan. So the
    # Regression build always rebuilds.
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
    print("Sheet updated: Production Lots")
    print("Sheet updated: Regression Instructions")
    print("Sheet updated: Diagnostic Guide")
    print("Sheet updated: Version History")
    print("Sheet updated: Regression")
    print_name_sync_summary(result)
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
            mileage_path=args.mileage_csv,
            production_lots_path=args.production_lots_csv,
            verbose=args.verbose,
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
            return 1
    else:
        print(f"Timing: build+sync    {build_elapsed:.1f}s")
        if recalc_elapsed is None:
            print("Timing: recalculate   skipped")
        else:
            print(f"Timing: recalculate   {recalc_elapsed:.1f}s")
        print(f"Timing: total         {time.monotonic() - total_start:.1f}s")

    return 0


def main() -> None:
    """Build the production workbook, archiving the run's transcript.

    The whole run — the verifier's ``VerifyReport`` on drift, and the
    traceback of anything that aborts it, stderr included — is teed into
    ``excel-only-runs/``. This build needs Excel, so it cannot run on the
    GitHub-hosted Linux CI; that directory is the cross-tool substitute for
    the CI log a failed ``poe verify-deep`` would otherwise leave only in
    somebody's terminal scrollback. See ``excel-only-runs/README.md`` for
    which transcripts belong on a branch and when to retire one.
    """
    args = parse_args()
    workbook_path = args.workbook.resolve()
    log_path = (
        args.log
        if args.log is not None
        else run_log_path(ROOT_DIR, Path(__file__).name, sys.argv[1:])
    )
    command = f"python {Path(__file__).name} " + " ".join(sys.argv[1:])

    with tee_run_log(log_path, command):
        exit_code = _build_and_verify(args, workbook_path)
        print(f"Log: {log_path}")

    # The launch is gated on a clean run for the same reason the verify
    # branch was: a workbook the verifier rejected must not be opened in
    # place of a fresh one.
    if not args.no_launch and exit_code == 0:
        subprocess.Popen(["cmd", "/c", "start", "", str(workbook_path)])

    if exit_code:
        sys.exit(exit_code)


if __name__ == "__main__":
    main()

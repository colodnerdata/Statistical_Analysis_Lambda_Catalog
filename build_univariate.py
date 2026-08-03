"""Build Lambda_Library_Univariate.xlsx — the standalone Univariate workbook — with its sheet and LAMBDA name sync.

From v3.0 the build emits two artifacts: ``build_production.py`` builds the
Regression workbook (``Lambda_Library.xlsx``), and this script builds the
standalone Univariate workbook (``Lambda_Library_Univariate.xlsx``). Moving
Univariate Analysis into its own workbook lets each artifact set its own
calculation mode: the Univariate grid searches — two two-input Data Tables
(Beta, across two stages) plus Weibull/Gamma static formula grids, ~2,400 NLL
evaluations per recalc in total — now recalculate on edit in full Automatic
inside this file, instead of forcing the shared workbook into
``XL_CALCULATION_SEMIAUTOMATIC`` and leaving fit results stale until a manual
Ctrl+Alt+F9 (see DECISIONS.md § v3.0 "Univariate becomes its own workbook").
Weibull and Gamma keep their 2-D grids as plain formula cells (no Data Table
object) pending the planned grid-shrink migration to 1-D profile line charts.

The shipped sheet set is small — LAMBDA_functions, Life Expectancy Data,
Univariate, Version History — but the workbook carries the complete
126-function LAMBDA library (no subsetting): the Univariate sheet depends on
30 workbook-scoped LAMBDA names (Skewness, Kurtosis, the NLL_*/CDF_* families,
etc.) that ``sync_workbook_names`` writes into ``xl/workbook.xml``. The Life
Expectancy Data sheet is required because the Univariate data zone reads
``LifeExpectancyData[Life expectancy]`` directly.
"""
from __future__ import annotations

import argparse
import importlib.util
import subprocess
import sys
import time
from pathlib import Path

import xlwings as xw

from lambda_catalog.build_common import (
    _close_workbook_quietly,
    _quit_app_quietly,
    _recalculate_and_save,
    _retry_on_open,
    print_name_sync_summary,
)
from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.verify_report import (
    VerifyReport,
    report_from_failures,
    render_human,
)
from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    _validate_workbook_reopen,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    load_csv_rows,
    write_csv_dataset_sheet,
)
from lambda_catalog.write_sheet_univariate import UNIVARIATE_SHEET_NAME, write_univariate_sheet
from lambda_catalog.write_sheet_version_history import write_version_history_sheet
from lambda_catalog.sheet_styles import SUBHDR_COLOR


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_UNIVARIATE_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library_Univariate.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"

_TAB_COLOR_LIGHT_GRAY = (217, 217, 217)
_TAB_COLOR_DARK_GRAY = (128, 128, 128)

_SHEET_NAME_LAMBDA_FUNCTIONS = "LAMBDA_functions"
_SHEET_NAME_LIFE_EXPECTANCY_DATA = "Life Expectancy Data"
_SHEET_NAME_VERSION_HISTORY = "Version History"

# The complete sheet set this artifact ships. Used both to assert the build
# produced exactly these sheets and to drop any stray sheets left behind when
# an existing file is reused (e.g. a Lambda_Library.xlsx copied as a starting
# point would carry Regression-side sheets this artifact must not keep).
_TARGET_SHEET_NAMES = frozenset(
    {
        _SHEET_NAME_LAMBDA_FUNCTIONS,
        _SHEET_NAME_LIFE_EXPECTANCY_DATA,
        UNIVARIATE_SHEET_NAME,
        _SHEET_NAME_VERSION_HISTORY,
    }
)


def _load_build_qc_module() -> object:
    """Import build_qc.py from the repo root without requiring it to be a package.

    build_qc.py is a top-level script in the repo (not under lambda_catalog/),
    so a normal ``import build_qc`` would rely on the consumer adding the
    repo root to ``sys.path``. Loading it explicitly here keeps the verify
    path self-contained when build_univariate is invoked as a script.
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
    verbose: bool = False,
) -> VerifyReport:
    """Run the spec-driven verifier against the standalone Univariate workbook.

    Opens the workbook in a headless Excel instance and calls
    ``build_qc.verify_test_sheets(..., skip_dummy=True, skip_regression=True,
    failures_out=...)``. ``skip_regression=True`` drops every Regression /
    Mileage / Production Lots check (this artifact carries none of those
    sheets); the Life Expectancy Full_Data check and the Univariate sheet
    check still run. On success, returns a passing ``VerifyReport``. On drift,
    ``verify_test_sheets`` raises ``RuntimeError("QC verification failed
    with N mismatch(es).")``; ``failures_out`` is populated before the raise
    so we can return a structured ``VerifyReport`` for the caller to render
    and ``sys.exit(1)`` without unwinding the build pipeline. Other exceptions
    propagate.
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
                    None,  # regression_sheet_configs — none for this artifact
                    csv_path,
                    verbose=verbose,
                    skip_dummy=True,
                    skip_regression=True,
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
    """Apply build-time tab order and tab colors for the Univariate workbook's sheets.

    The Univariate workbench sits front-most (the reason the file exists),
    followed by its data source and the version history, with the
    LAMBDA_functions catalog last.
    """
    ordered_front = [
        UNIVARIATE_SHEET_NAME,
        _SHEET_NAME_LIFE_EXPECTANCY_DATA,
        _SHEET_NAME_VERSION_HISTORY,
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
        UNIVARIATE_SHEET_NAME: SUBHDR_COLOR,
        _SHEET_NAME_LIFE_EXPECTANCY_DATA: _TAB_COLOR_LIGHT_GRAY,
        _SHEET_NAME_VERSION_HISTORY: _TAB_COLOR_DARK_GRAY,
    }

    present = _sheet_names(workbook)
    for sheet_name, color in tab_colors.items():
        if sheet_name in present:
            _set_tab_color(workbook.sheets[sheet_name], color)


def build_univariate_workbook(
    workbook_path: Path = DEFAULT_UNIVARIATE_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = LIFE_EXPECTANCY.default_csv_path,
    validate_reopen: bool = False,
    verbose: bool = False,
    recalculate: bool = True,
) -> NameSyncResult:
    """Build the standalone Univariate workbook and sync the LAMBDA name manager.

    Parameters
    ----------
    workbook_path : Path, optional
        Path to the workbook to create or update.
    definitions_path : Path, optional
        Path to the JSON catalog file.
    csv_path : Path, optional
        Path to the Life Expectancy CSV file — the dataset the Univariate
        data zone reads (``LifeExpectancyData[Life expectancy]``).
    validate_reopen : bool, optional
        If True, reopens the workbook in Excel after patching to verify it.
    verbose : bool, optional
        If True, prints timing information for each build phase to stdout.
    recalculate : bool, optional
        If True (default), recalculates and saves as the final build step —
        this is what computes the Univariate grid searches (the two Beta Data
        Tables and the Weibull/Gamma formula grids) so the shipped artifact is
        not stale. Pass False (or use ``--skip-data-table-calculations``) when
        the caller manages the recalculate step separately, or for fast
        iteration where the ~2,400 NLL evaluations per recalc would dominate
        build time.

    Returns
    -------
    NameSyncResult
        Counts of created versus updated workbook names.
    """
    _t = time.monotonic()
    document = load_catalog_document(definitions_path)
    csv_headers, csv_rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)
    if verbose:
        print(f"  Prep:           {time.monotonic() - _t:.1f}s", flush=True)

    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    _t = time.monotonic()
    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        app = xw.App(visible=True, add_book=False)
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
        else:
            workbook = app.books.add()

        # If the file is a reused workbook (e.g. a copied Lambda_Library.xlsx),
        # drop every sheet that is not part of this artifact before writing.
        for sheet in list(workbook.sheets):
            if sheet.name not in _TARGET_SHEET_NAMES and sheet.name != "Sheet1":
                sheet.delete()

        try:
            app.api.Calculation = XL_CALCULATION_MANUAL
            if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                workbook.sheets["Sheet1"].name = _SHEET_NAME_LAMBDA_FUNCTIONS
            write_catalog_sheet(workbook, document.functions)
            write_csv_dataset_sheet(workbook, csv_headers, csv_rows, LIFE_EXPECTANCY)
            write_univariate_sheet(workbook, document.univariate_sheet_notes)
            write_version_history_sheet(workbook, artifact="univariate")
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
    """Parse command-line arguments for Univariate workbook generation.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with workbook, definitions, csv, validate_reopen,
        verbose, skip_data_table_calculations, verify, and no_launch
        attributes.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Build Lambda_Library_Univariate.xlsx — the standalone Univariate "
            "workbook — from lambda_functions.json. The Regression workbook "
            "ships separately (build_production.py)."
        )
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_UNIVARIATE_WORKBOOK_PATH,
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
        help="Path to the Life Expectancy CSV data file the Univariate sheet reads.",
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
        "--skip-data-table-calculations",
        action="store_true",
        help=(
            "Skip the final Excel CalculateFullRebuild phase. The Univariate "
            "grid searches make that rebuild slow (~2,400 NLL evaluations per "
            "recalc across the Beta Data Tables and the Weibull/Gamma formula "
            "grids), so this is the primary fast-iteration flag for this "
            "artifact. Note: the spec-driven verifier (build_qc) only does a "
            "per-sheet Calculate, which does not reliably resolve the Data "
            "Tables after a name sync — so combining this flag with --verify "
            "may report stale-fit mismatches that a real rebuild would not."
        ),
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        default=False,
        help=(
            "After the build, run the spec-driven verifier "
            "(build_qc.verify_test_sheets with skip_regression=True) against "
            "the Univariate sheet and the Life Expectancy Data sheet. On any "
            "drift, print a structured VerifyReport and sys.exit(1). The "
            "post-build Excel handoff (cmd /c start) only fires when verify "
            "passes. Combine with --no-launch to run verify without opening "
            "Excel at all."
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


def main() -> None:
    """Build the Univariate workbook and print a short sync summary for interactive use."""
    args = parse_args()
    workbook_path = args.workbook.resolve()
    total_start = time.monotonic()
    recalc_elapsed: float | None = None

    # Phase 1: write all sheets + sync names.
    # If the workbook is open in Excel at this point the entire write must be retried.
    result: NameSyncResult | None = None

    def _run_build() -> None:
        nonlocal result
        result = build_univariate_workbook(
            workbook_path=workbook_path,
            definitions_path=args.definitions,
            csv_path=args.csv,
            validate_reopen=False,  # handled below after recalculate
            verbose=args.verbose,
            recalculate=False,      # handled separately so only this step retries
        )

    build_phase_start = time.monotonic()
    _retry_on_open(
        f"{args.workbook.name} is open in Excel",
        _run_build,
        retry_rpc=True,
    )
    build_elapsed = time.monotonic() - build_phase_start
    assert result is not None

    # Phase 2: recalculate the Data Tables and save.
    # This is the slow step for this artifact (the six Univariate Data Tables),
    # and the one most likely to fail when the user opens the workbook to
    # inspect progress — so it gets its own retry phase, separate from the
    # multi-minute write. Skipping it (--skip-data-table-calculations) leaves
    # the Data Tables uncomputed; the shipped artifact is built without it.
    if not args.skip_data_table_calculations:
        _t = time.monotonic()
        _retry_on_open(
            f"{args.workbook.name} is open in Excel",
            lambda: _recalculate_and_save(workbook_path),
            retry_rpc=True,
        )
        recalc_elapsed = time.monotonic() - _t
        if args.verbose:
            print(f"  Recalculate:    {recalc_elapsed:.1f}s", flush=True)
    elif args.verbose:
        print("  Recalculate:    skipped (--skip-data-table-calculations)", flush=True)

    if args.validate_reopen:
        _validate_workbook_reopen(workbook_path)

    print(f"Workbook: {workbook_path}")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print("Sheet updated: Univariate")
    print("Sheet updated: Version History")
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
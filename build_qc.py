"""Build Lambda_Library_QC.xlsx: all sheets, analysis cache update, and test verification."""
from __future__ import annotations

import argparse
import importlib.util
import io
import subprocess
import sys
import time
from pathlib import Path

import xlwings as xw

from lambda_catalog.analysis_cache import DEFAULT_CACHE_PATH, get_analysis_results
from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_AUTOMATIC,
    _delete_sheet_if_present,
    _validate_workbook_reopen,
    load_lambda_definitions,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_lambda_functions import (
    load_catalog_entries,
    load_regression_sheet_notes,
    write_catalog_sheet,
)
from lambda_catalog.write_sheet_life_expectancy_data import (
    DEFAULT_CSV_PATH,
    load_life_expectancy_rows,
    write_life_expectancy_sheet,
)
from lambda_catalog.write_sheet_mlr_observation_test import write_mlr_observation_test_sheet
from lambda_catalog.write_sheet_mlr_scalar_test import write_mlr_scalar_test_sheet
from lambda_catalog.write_sheet_mlr_vector_outputs_test import write_mlr_vector_outputs_test_sheet
from lambda_catalog.write_sheet_diagnostic_guide import write_diagnostic_guide_sheet
from lambda_catalog.write_sheet_regression import write_regression_output_sheet
from lambda_catalog.write_sheet_regression_instructions import write_regression_instructions_sheet


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library_QC.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
_PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
_QC_SHEET_NAMES = ("MLR_Scalar_Test", "MLR_Vector_Outputs_Test", "MLR_Observation_Test")


def _verbose_checkpoint(verbose: bool, start_time: float, label: str) -> None:
    """Print a single verbose progress checkpoint with elapsed seconds."""
    if verbose:
        print(f"  {label:<28} {time.monotonic() - start_time:6.1f}s", flush=True)


def verify_test_sheets(
    workbook: xw.Book,
    scalar_row_configs: list,
    vector_row_configs: list,
    observation_row_configs: list,
    regression_sheet_configs: list,
    verbose: bool = False,
) -> None:
    """Compare Excel Calc columns against Python-computed expected values in all test sheets.

    Forces a full recalculation on the open workbook, then reads each sheet's
    Calc columns via inspect_test_sheets, compares them against expected values
    derived from the supplied analysis configs (not from the sheet's (Exp.) columns),
    and warns (without raising) for every value that diverges within the tolerance band.

    Parameters
    ----------
    workbook : xw.Book
        Open xlwings Book to inspect (must already be saved so formulas are
        linked to the workbook-scoped defined names).
    scalar_row_configs : list
        Per-row configs from ``build_mlr_row_configs``; provides the Python-computed
        expected scalar values used for comparison.
    vector_row_configs : list
        Per-section configs from ``build_mlr_vector_row_configs``; provides the
        Python-computed expected coefficient vectors used for comparison.
    observation_row_configs : list
        Per-section configs from ``build_mlr_observation_row_configs``; provides the
        Python-computed expected observation-level values used for comparison.
    regression_sheet_configs : list
        Per-config tuples from ``build_regression_sheet_qc_configs``; provides the
        Python-computed expected values for all Regression sheet output zones.
    """
    tool_path = ROOT_DIR / "tools" / "inspect_test_sheets.py"
    spec = importlib.util.spec_from_file_location("inspect_test_sheets", tool_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load inspect_test_sheets from {tool_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    phase_start = time.monotonic()
    _verbose_checkpoint(verbose, phase_start, "Verify: calculate start")
    workbook.app.calculate()
    _verbose_checkpoint(verbose, phase_start, "Verify: calculate done")

    _verbose_checkpoint(verbose, phase_start, "Verify: read scalar start")
    scalar_df = mod.read_scalar_df(workbook, scalar_row_configs)
    _verbose_checkpoint(verbose, phase_start, "Verify: read scalar done")
    _verbose_checkpoint(verbose, phase_start, "Verify: read vector start")
    vector_df = mod.read_vector_df(workbook, vector_row_configs)
    _verbose_checkpoint(verbose, phase_start, "Verify: read vector done")
    _verbose_checkpoint(verbose, phase_start, "Verify: read obs start")
    observation_df = mod.read_observation_df(workbook, observation_row_configs)
    _verbose_checkpoint(verbose, phase_start, "Verify: read obs done")
    tol = mod.TOLERANCE_DECIMALS

    for _, row in scalar_df.iterrows():
        fdd = row["first_digit_deviation"]
        if fdd is not None and fdd <= tol:
            k = row["k"]
            ai = row["allow_intercept"]
            stat = row["stat_name"]
            print(
                f"WARNING [MLR_Scalar_Test] k={k} intercept={'TRUE' if ai else 'FALSE'} "
                f"stat={stat!r}: Calc vs Python-expected mismatch — first digit of deviation at "
                f"decimal place {fdd} (tolerance={tol}). "
                f"expected={row['expected']!r}, excel_calc={row['excel_calc']!r}, "
                f"abs_diff={row['abs_diff']!r}",
                flush=True,
            )

    for _, row in vector_df.iterrows():
        fdd = row["first_digit_deviation"]
        if fdd is not None and fdd <= tol:
            k = row["k"]
            ai = row["allow_intercept"]
            stat = row["stat_name"]
            term = row["term_name"]
            print(
                f"WARNING [MLR_Vector_Outputs_Test] k={k} intercept={'TRUE' if ai else 'FALSE'} "
                f"stat={stat!r} term={term!r}: Calc vs Python-expected mismatch — first digit of deviation "
                f"at decimal place {fdd} (tolerance={tol}). "
                f"expected={row['expected']!r}, excel_calc={row['excel_calc']!r}, "
                f"abs_diff={row['abs_diff']!r}",
                flush=True,
            )

    for _, row in observation_df.iterrows():
        fdd = row["first_digit_deviation"]
        if fdd is not None and fdd <= tol:
            print(
                f"WARNING [MLR_Observation_Test] k={row['k']} intercept={'TRUE' if row['allow_intercept'] else 'FALSE'} "
                f"stat={row['stat_name']!r}: expected={row['expected']!r}, excel_calc={row['excel_calc']!r}",
                flush=True,
            )

    # Phase 4: Regression sheet verification
    reg_tool_path = ROOT_DIR / "tools" / "inspect_regression_sheet.py"
    reg_spec = importlib.util.spec_from_file_location("inspect_regression_sheet", reg_tool_path)
    if reg_spec is None or reg_spec.loader is None:
        raise RuntimeError(f"Could not load inspect_regression_sheet from {reg_tool_path}")
    reg_mod = importlib.util.module_from_spec(reg_spec)
    reg_spec.loader.exec_module(reg_mod)

    _verbose_checkpoint(verbose, phase_start, "Verify: regression start")
    reg_dfs = reg_mod.read_regression_df(workbook, regression_sheet_configs)
    _verbose_checkpoint(verbose, phase_start, "Verify: regression done")

    _SECTION_DF_KEYS = [
        ("scalars",             ["config_name", "allow_intercept", "stat_name"]),
        ("predictors",          ["config_name", "allow_intercept", "predictor_name", "stat_name"]),
        ("coefficients",        ["config_name", "allow_intercept", "term_name", "stat_name"]),
        ("prediction_interval", ["config_name", "allow_intercept", "stat_name"]),
        ("residuals",           ["config_name", "allow_intercept", "row_idx", "stat_name"]),
    ]
    for section_key, id_cols in _SECTION_DF_KEYS:
        df = reg_dfs[section_key]
        for _, row in df.iterrows():
            fdd = row["first_digit_deviation"]
            if fdd is not None and fdd <= tol:
                identity = " ".join(
                    f"{col}={row[col]!r}" for col in id_cols if col in row.index
                )
                print(
                    f"WARNING [Regression/{section_key}] {identity}: "
                    f"expected={row['expected']!r}, excel_calc={row['excel_calc']!r}, "
                    f"abs_diff={row['abs_diff']!r}, fdd={fdd}",
                    flush=True,
                )


def build_qc_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = DEFAULT_CSV_PATH,
    cache_path: Path = DEFAULT_CACHE_PATH,
    validate_reopen: bool = False,
    verbose: bool = False,
) -> NameSyncResult:
    """Build all 6 sheets, update the analysis cache, and verify test-sheet accuracy.

    Parameters
    ----------
    workbook_path : Path, optional
        Path to the QC workbook to create or update.
    definitions_path : Path, optional
        Path to the JSON catalog file.
    csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    cache_path : Path, optional
        Path to the analysis cache JSON file.
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
    definitions = load_lambda_definitions(definitions_path)
    _verbose_checkpoint(verbose, _t, "Prep: definitions loaded")
    catalog_entries = load_catalog_entries(definitions_path)
    _verbose_checkpoint(verbose, _t, "Prep: catalog loaded")
    sheet_notes = load_regression_sheet_notes(definitions_path)
    _verbose_checkpoint(verbose, _t, "Prep: sheet notes loaded")
    row_configs, vector_row_configs, observation_row_configs, regression_sheet_configs = get_analysis_results(
        csv_path, cache_path
    )
    _verbose_checkpoint(verbose, _t, "Prep: analysis loaded")
    csv_headers, csv_rows = load_life_expectancy_rows(csv_path)
    _verbose_checkpoint(verbose, _t, "Prep: csv loaded")
    _verbose_checkpoint(verbose, _t, "Prep total")

    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    # Phase 1: Write all sheets and save
    _t = time.monotonic()
    try:
        _verbose_checkpoint(verbose, _t, "Write: app start")
        with xw.App(visible=False, add_book=False) as app:
            _verbose_checkpoint(verbose, _t, "Write: app ready")
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            if workbook_exists:
                _verbose_checkpoint(verbose, _t, "Write: open workbook")
                workbook = app.books.open(str(workbook_path))
                _verbose_checkpoint(verbose, _t, "Write: workbook opened")
            else:
                _verbose_checkpoint(verbose, _t, "Write: add workbook")
                workbook = app.books.add()
                _verbose_checkpoint(verbose, _t, "Write: workbook added")
                for sheet in list(workbook.sheets)[1:]:
                    sheet.delete()

            try:
                _delete_sheet_if_present(workbook, _PREDICTIONS_SHEET_NAME)
                for qc_sheet in _QC_SHEET_NAMES:
                    _delete_sheet_if_present(workbook, qc_sheet)
                if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                    workbook.sheets["Sheet1"].name = "LAMBDA_functions"
                _verbose_checkpoint(verbose, _t, "Write: catalog start")
                write_catalog_sheet(workbook, catalog_entries)
                _verbose_checkpoint(verbose, _t, "Write: catalog done")
                _verbose_checkpoint(verbose, _t, "Write: life exp start")
                write_life_expectancy_sheet(workbook, csv_headers, csv_rows, verbose=verbose)
                _verbose_checkpoint(verbose, _t, "Write: life exp done")
                _verbose_checkpoint(verbose, _t, "Write: regression instr start")
                write_regression_instructions_sheet(workbook)
                _verbose_checkpoint(verbose, _t, "Write: regression instr done")
                _verbose_checkpoint(verbose, _t, "Write: diagnostic start")
                write_diagnostic_guide_sheet(workbook)
                _verbose_checkpoint(verbose, _t, "Write: diagnostic done")
                _verbose_checkpoint(verbose, _t, "Write: regression start")
                write_regression_output_sheet(workbook, sheet_notes)
                _verbose_checkpoint(verbose, _t, "Write: regression done")
                _verbose_checkpoint(verbose, _t, "Write: scalar start")
                write_mlr_scalar_test_sheet(workbook, definitions, row_configs)
                _verbose_checkpoint(verbose, _t, "Write: scalar done")
                _verbose_checkpoint(verbose, _t, "Write: vector start")
                write_mlr_vector_outputs_test_sheet(workbook, vector_row_configs)
                _verbose_checkpoint(verbose, _t, "Write: vector done")
                _verbose_checkpoint(verbose, _t, "Write: observation start")
                write_mlr_observation_test_sheet(workbook, observation_row_configs)
                _verbose_checkpoint(verbose, _t, "Write: observation done")
                app.api.Calculation = XL_CALCULATION_AUTOMATIC
                _verbose_checkpoint(verbose, _t, "Write: save start")
                workbook.save(str(workbook_path))
                _verbose_checkpoint(verbose, _t, "Write: save done")
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)
    if verbose:
        print(f"  Write sheets:   {time.monotonic() - _t:.1f}s", flush=True)

    # Phase 2: Sync LAMBDA definitions into the Name Manager via XML patch
    _t = time.monotonic()
    try:
        _verbose_checkpoint(verbose, _t, "Sync: start")
        result = sync_workbook_names(workbook_path, definitions)
        _verbose_checkpoint(verbose, _t, "Sync: done")
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    except (PermissionError, OSError) as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    if verbose:
        print(f"  Sync names:     {time.monotonic() - _t:.1f}s", flush=True)

    # Phase 3: Reopen and verify test sheets against the freshly synced definitions
    _t = time.monotonic()
    try:
        _verbose_checkpoint(verbose, _t, "Verify: app start")
        with xw.App(visible=False, add_book=False) as app:
            _verbose_checkpoint(verbose, _t, "Verify: app ready")
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            _verbose_checkpoint(verbose, _t, "Verify: open workbook")
            workbook = app.books.open(str(workbook_path))
            try:
                _verbose_checkpoint(verbose, _t, "Verify: workbook opened")
                verify_test_sheets(
                    workbook,
                    row_configs,
                    vector_row_configs,
                    observation_row_configs,
                    regression_sheet_configs,
                    verbose=verbose,
                )
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "verify", exc)
    if verbose:
        print(f"  Verify:         {time.monotonic() - _t:.1f}s", flush=True)

    if validate_reopen:
        _validate_workbook_reopen(workbook_path)

    return result


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for QC workbook generation.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with workbook, definitions, csv, cache,
        validate_reopen, and verbose attributes.
    """
    parser = argparse.ArgumentParser(
        description="Build Lambda_Library_QC.xlsx with all sheets, cache update, and test verification."
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
        default=DEFAULT_CSV_PATH,
        help="Path to the Life Expectancy CSV data file.",
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=DEFAULT_CACHE_PATH,
        help="Path to the analysis cache JSON file.",
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


class _Tee(io.TextIOBase):
    """Write to two streams simultaneously."""

    def __init__(self, primary: io.TextIOBase, secondary: io.TextIOBase) -> None:
        self._primary = primary
        self._secondary = secondary

    def write(self, s: str) -> int:
        self._secondary.write(s)
        return self._primary.write(s)

    def flush(self) -> None:
        self._primary.flush()
        self._secondary.flush()


DEFAULT_LOG_PATH = ROOT_DIR / "qc_log.txt"


def main() -> None:
    """Build the QC workbook and print a sync summary for interactive use."""
    args = parse_args()
    with open(DEFAULT_LOG_PATH, "w", encoding="utf-8") as _log_file:
        _real_stdout = sys.stdout
        sys.stdout = _Tee(sys.stdout, _log_file)  # type: ignore[assignment]
        try:
            print(f"python {Path(__file__).name} " + " ".join(sys.argv[1:]))
            _run_main(args)
        finally:
            sys.stdout = _real_stdout


def _run_main(args: argparse.Namespace) -> None:
    while True:
        try:
            result = build_qc_workbook(
                workbook_path=args.workbook,
                definitions_path=args.definitions,
                csv_path=args.csv,
                cache_path=args.cache,
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
    print("Sheet updated: Regression")
    print("Sheet updated: MLR_Scalar_Test")
    print("Sheet updated: MLR_Vector_Outputs_Test")
    print("Sheet updated: MLR_Observation_Test")
    print("Sheet verified: Regression")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    if args.validate_reopen:
        print("Reopen validation: passed")

    subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])
    print(f"Log: {DEFAULT_LOG_PATH}")


if __name__ == "__main__":
    main()

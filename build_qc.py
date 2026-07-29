"""Build Lambda_Library_QC.xlsx and verify the QC workbook."""
from __future__ import annotations

import argparse
from collections import Counter
import importlib.util
import io
import subprocess
import sys
import time
from pathlib import Path
from typing import TextIO

import xlwings as xw

from lambda_catalog.analyze_life_expectancy import calculate_data_completeness_flags
from lambda_catalog.analyze_mileage import calculate_mileage_completeness_flags
from lambda_catalog.analyze_production_lots import calculate_production_lots_completeness_flags
from lambda_catalog.analyze_regression_spec import build_regression_spec_qc_configs
from lambda_catalog.analyze_regression_spec_block import read_regression_spec_block_failures
from lambda_catalog.analysis_cache import DEFAULT_CACHE_PATH
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
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_diagnostic_guide import write_diagnostic_guide_sheet
from lambda_catalog.write_sheet_dummy_test import (
    read_dummy_check_failures,
    write_dummy_test_sheet,
)
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


ROOT_DIR = Path(__file__).resolve().parent
_MAX_REPORTED_QC_FAILURES = 200
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library_QC.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
DEFAULT_LOG_PATH = ROOT_DIR / "qc_log.txt"
_PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
_QC_SHEET_NAMES = (
    "MLR_Scalar_Test",
    "MLR_Vector_Outputs_Test",
    "MLR_Observation_Test",
    "Dummy_Test",
)
_VERIFY_CALC_SHEET_NAMES = (
    LIFE_EXPECTANCY.sheet_name,
    MILEAGE.sheet_name,
    PRODUCTION_LOTS.sheet_name,
    "Regression",
    "Univariate",
    "Dummy_Test",
)


def _verification_calc_sheet_names(
    *, skip_dummy: bool, skip_univariate: bool = False
) -> tuple[str, ...]:
    """Return workbook sheets that must be force-calculated before verify."""
    names = _VERIFY_CALC_SHEET_NAMES
    if skip_dummy:
        names = tuple(name for name in names if name != "Dummy_Test")
    if skip_univariate:
        names = tuple(name for name in names if name != "Univariate")
    return names


def _verbose_checkpoint(verbose: bool, start_time: float, label: str) -> None:
    """Print one verbose progress checkpoint with elapsed seconds."""
    if verbose:
        print(f"  {label:<28} {time.monotonic() - start_time:6.1f}s", flush=True)


def _normalize_excel_bool(value):
    if value in ("TRUE", "True", 1, 1.0):
        return True
    if value in ("FALSE", "False", 0, 0.0):
        return False
    return value


def _report_qc_failure(failures: list[str], message: str) -> None:
    failures.append(message)
    if len(failures) <= _MAX_REPORTED_QC_FAILURES:
        print(f"ERROR {message}", flush=True)
    elif len(failures) == _MAX_REPORTED_QC_FAILURES + 1:
        print("ERROR Additional QC mismatches suppressed.", flush=True)


_OPTIONAL_VERIFY_SHEET_NAMES = ("Univariate",)


def _calculate_verification_sheets(
    workbook: xw.Book,
    verbose: bool,
    phase_start: float,
    *,
    skip_dummy: bool,
    skip_univariate: bool = False,
) -> None:
    sheet_names = _verification_calc_sheet_names(
        skip_dummy=skip_dummy, skip_univariate=skip_univariate
    )
    workbook_sheet_names = {sheet.name for sheet in workbook.sheets}
    missing_sheets = [name for name in sheet_names if name not in workbook_sheet_names]

    # A sheet named here can be legitimately absent (e.g. a workbook built
    # with --skip-univariate never got a Univariate sheet at all) even when
    # the caller didn't pass the matching skip flag — warn and skip
    # calculating it rather than crashing. Any other missing sheet indicates
    # a real build problem and still hard-fails.
    required_missing = [
        name for name in missing_sheets if name not in _OPTIONAL_VERIFY_SHEET_NAMES
    ]
    if required_missing:
        raise RuntimeError(
            "Verification requires worksheet(s) that are missing from the workbook: "
            + ", ".join(required_missing)
        )
    for name in missing_sheets:
        print(
            f"WARNING Sheet '{name}' not found in workbook; skipping its verification.",
            flush=True,
        )
    sheet_names = [name for name in sheet_names if name not in missing_sheets]

    _verbose_checkpoint(verbose, phase_start, "Verify: calculate start")
    workbook.app.api.Calculation = XL_CALCULATION_MANUAL
    try:
        for sheet_name in sheet_names:
            _verbose_checkpoint(verbose, phase_start, f"Calc: {sheet_name[:20]} start")
            workbook.sheets[sheet_name].api.Calculate()
            _verbose_checkpoint(verbose, phase_start, f"Calc: {sheet_name[:20]} done")
    finally:
        workbook.app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
    _verbose_checkpoint(verbose, phase_start, "Verify: calculate done")


def _load_module(module_name: str, path: Path):
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load {module_name} from {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _verify_life_expectancy_full_data(
    workbook: xw.Book,
    csv_path: Path,
    failures: list[str],
) -> None:
    full_data_expected = calculate_data_completeness_flags(csv_path)
    life_expectancy_sheet = workbook.sheets[LIFE_EXPECTANCY.sheet_name]
    life_expectancy_data = life_expectancy_sheet.used_range.value
    if not life_expectancy_data:
        return

    if isinstance(life_expectancy_data[0], list):
        life_expectancy_rows = life_expectancy_data
    else:
        life_expectancy_rows = [life_expectancy_data]
    life_expectancy_headers = [
        str(header).strip() if header is not None else ""
        for header in life_expectancy_rows[0]
    ]
    full_data_col_idx = life_expectancy_headers.index(LIFE_EXPECTANCY.full_data_header)
    for row_offset, expected in enumerate(full_data_expected, start=1):
        row = life_expectancy_rows[row_offset] if row_offset < len(life_expectancy_rows) else []
        actual = _normalize_excel_bool(
            row[full_data_col_idx] if full_data_col_idx < len(row) else None
        )
        if actual is not expected:
            _report_qc_failure(
                failures,
                f"[Life Expectancy Data] row={row_offset + 1} stat='Full_Data': "
                f"expected={expected!r}, excel_calc={actual!r}",
            )


def _verify_mileage_full_data(
    workbook: xw.Book,
    mileage_path: Path,
    failures: list[str],
) -> None:
    full_data_expected = calculate_mileage_completeness_flags(mileage_path)
    mileage_sheet = workbook.sheets[MILEAGE.sheet_name]
    mileage_data = mileage_sheet.used_range.value
    if not mileage_data:
        return

    if isinstance(mileage_data[0], list):
        mileage_rows = mileage_data
    else:
        mileage_rows = [mileage_data]
    mileage_headers = [
        str(header).strip() if header is not None else ""
        for header in mileage_rows[0]
    ]
    full_data_col_idx = mileage_headers.index(MILEAGE.full_data_header)
    for row_offset, expected in enumerate(full_data_expected, start=1):
        row = mileage_rows[row_offset] if row_offset < len(mileage_rows) else []
        actual = _normalize_excel_bool(
            row[full_data_col_idx] if full_data_col_idx < len(row) else None
        )
        if actual is not expected:
            _report_qc_failure(
                failures,
                f"[Mileage Data] row={row_offset + 1} stat='Full_Data': "
                f"expected={expected!r}, excel_calc={actual!r}",
            )


def _verify_production_lots_full_data(
    workbook: xw.Book,
    production_lots_path: Path,
    failures: list[str],
) -> None:
    full_data_expected = calculate_production_lots_completeness_flags(production_lots_path)
    production_lots_sheet = workbook.sheets[PRODUCTION_LOTS.sheet_name]
    production_lots_data = production_lots_sheet.used_range.value
    if not production_lots_data:
        return

    if isinstance(production_lots_data[0], list):
        production_lots_rows = production_lots_data
    else:
        production_lots_rows = [production_lots_data]
    production_lots_headers = [
        str(header).strip() if header is not None else ""
        for header in production_lots_rows[0]
    ]
    full_data_col_idx = production_lots_headers.index(PRODUCTION_LOTS.full_data_header)
    for row_offset, expected in enumerate(full_data_expected, start=1):
        row = (
            production_lots_rows[row_offset]
            if row_offset < len(production_lots_rows)
            else []
        )
        actual = _normalize_excel_bool(
            row[full_data_col_idx] if full_data_col_idx < len(row) else None
        )
        if actual is not expected:
            _report_qc_failure(
                failures,
                f"[Production Lots] row={row_offset + 1} stat='Full_Data': "
                f"expected={expected!r}, excel_calc={actual!r}",
            )


def verify_test_sheets(
    workbook: xw.Book,
    regression_sheet_configs: list,
    csv_path: Path,
    verbose: bool = False,
    *,
    mileage_path: Path = MILEAGE.default_csv_path,
    production_lots_path: Path = PRODUCTION_LOTS.default_csv_path,
    skip_dummy: bool = False,
    skip_univariate: bool = False,
    failures_out: list[str] | None = None,
) -> None:
    """Compare live workbook output against Python-computed QC oracle values.

    Parameters
    ----------
    workbook : xw.Book
        The live xlwings workbook to verify.
    regression_sheet_configs : list
        Per-config Python oracle (from ``build_regression_spec_qc_configs``).
    csv_path : Path
        Path to the canonical CSV used for Full_Data comparison.
    verbose : bool
        Print per-phase checkpoints to stdout.
    mileage_path : Path
        Path to the Auto MPG sample CSV used for the Mileage Data sheet's
        Full_Data comparison. Defaults to the committed sample file.
    production_lots_path : Path
        Path to the Production Lots sample CSV used for the Production Lots
        sheet's Full_Data comparison. Defaults to the committed sample file.
    skip_dummy : bool
        When True, skip the Dummy_Test check (``read_dummy_check_failures``).
        Used by ``build_production.py --verify`` which produces a workbook
        without a ``Dummy_Test`` sheet. Defaults to False (legacy QC build
        behaviour).
    skip_univariate : bool
        When True, skip the Univariate sheet check
        (``read_univariate_failures``). Used by
        ``build_production.py --verify --skip-univariate`` which produces
        a workbook without a ``Univariate`` sheet. Defaults to False. Even
        when False, a workbook that is missing the ``Univariate`` sheet
        (e.g. built earlier with --skip-univariate) is handled the same
        way — the check is skipped with a warning rather than raising,
        since the sheet's absence is expected in that case, not a build
        failure.
    failures_out : list[str] | None
        Optional external list. When supplied, every captured failure message
        is appended to it before the function raises on drift. Used by
        ``build_production.py --verify`` to populate a ``VerifyReport`` for
        downstream structured rendering. The internal ``failures`` list is
        the source of truth; this is a write-through mirror.
    """
    failures: list[str] = []
    phase_start = time.monotonic()
    _calculate_verification_sheets(
        workbook,
        verbose,
        phase_start,
        skip_dummy=skip_dummy,
        skip_univariate=skip_univariate,
    )
    _verify_life_expectancy_full_data(workbook, csv_path, failures)
    _verify_mileage_full_data(workbook, mileage_path, failures)
    _verify_production_lots_full_data(workbook, production_lots_path, failures)

    _verbose_checkpoint(verbose, phase_start, "Verify: reg spec block start")
    for failure in read_regression_spec_block_failures(workbook, mileage_path):
        _report_qc_failure(failures, failure)
    _verbose_checkpoint(verbose, phase_start, "Verify: reg spec block done")

    reg_mod = _load_module(
        "inspect_regression_sheet",
        ROOT_DIR / "tools" / "inspect_regression_sheet.py",
    )

    _verbose_checkpoint(verbose, phase_start, "Verify: regression start")
    reg_dfs = reg_mod.read_regression_df(workbook, regression_sheet_configs)
    _verbose_checkpoint(verbose, phase_start, "Verify: regression done")

    section_df_keys = [
        ("scalars", ["config_name", "allow_intercept", "stat_name"]),
        ("predictors", ["config_name", "allow_intercept", "predictor_name", "stat_name"]),
        ("coefficients", ["config_name", "allow_intercept", "term_name", "stat_name"]),
        ("prediction_interval", ["config_name", "allow_intercept", "stat_name"]),
        ("residuals", ["config_name", "allow_intercept", "row_idx", "stat_name"]),
    ]
    for section_key, id_cols in section_df_keys:
        df = reg_dfs[section_key]
        for _, row in df.iterrows():
            fdd = row["first_digit_deviation"]
            if fdd is not None and fdd <= reg_mod.TOLERANCE_DECIMALS:
                identity = " ".join(
                    f"{col}={row[col]!r}" for col in id_cols if col in row.index
                )
                _report_qc_failure(
                    failures,
                    f"[Regression/{section_key}] {identity}: "
                    f"expected={row['expected']!r}, excel_calc={row['excel_calc']!r}, "
                    f"abs_diff={row['abs_diff']!r}, fdd={fdd}",
                )

    workbook_sheet_names = {sheet.name for sheet in workbook.sheets}
    if skip_univariate or "Univariate" not in workbook_sheet_names:
        print(
            "WARNING Univariate sheet not verified "
            "(missing from workbook or --skip-univariate was used).",
            flush=True,
        )
    else:
        uv_mod = _load_module(
            "inspect_univariate_sheet",
            ROOT_DIR / "tools" / "inspect_univariate_sheet.py",
        )
        _verbose_checkpoint(verbose, phase_start, "Verify: univariate start")
        for failure in uv_mod.read_univariate_failures(workbook, csv_path):
            _report_qc_failure(failures, failure)
        _verbose_checkpoint(verbose, phase_start, "Verify: univariate done")

    if not skip_dummy:
        _verbose_checkpoint(verbose, phase_start, "Verify: dummy test start")
        for failure in read_dummy_check_failures(workbook):
            _report_qc_failure(failures, failure)
        _verbose_checkpoint(verbose, phase_start, "Verify: dummy test done")

    if failures:
        if failures_out is not None:
            failures_out.extend(failures)
        category_counts = Counter(
            message.split("]", 1)[0].removeprefix("[") for message in failures
        )
        summary = ", ".join(
            f"{category}={count}" for category, count in sorted(category_counts.items())
        )
        print(f"ERROR QC mismatch totals: {summary}", flush=True)
        raise RuntimeError(f"QC verification failed with {len(failures)} mismatch(es).")


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
            if "likely open in Excel" in str(exc):
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
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
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

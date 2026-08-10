"""Spec-driven deep verification helpers for generated workbooks.

Importable from `lambda_catalog` so `build_production.py` and
`tools/verify_workbook.py` can call `verify_test_sheets` directly.

**Runs from a repository checkout, not from an installed wheel.** It reads the
`tools/` inspector scripts (`_load_module`) and the `sample_data/` CSVs (via
`write_sheet_csv_dataset`'s `default_csv_path`s), and it drives desktop Excel
through xlwings. The wheel packages `lambda_catalog` only, so none of those are
present in an installed environment — which is also why CI never runs this
layer (see CLAUDE.md).
"""

from __future__ import annotations

import importlib.util
import time
from collections import Counter
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_life_expectancy import (
    calculate_developed_country_flags,
)
from lambda_catalog.analyze_regression_spec_block import (
    read_regression_spec_block_failures,
)
from lambda_catalog.workbook_builder import (
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
)
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    MILEAGE,
    PRODUCTION_LOTS,
)
from lambda_catalog.write_sheet_univariate import UNIVARIATE_SHEET_NAME

ROOT_DIR = Path(__file__).resolve().parent.parent
_MAX_REPORTED_QC_FAILURES = 200
_VERIFY_CALC_SHEET_NAMES = (
    LIFE_EXPECTANCY.sheet_name,
    MILEAGE.sheet_name,
    PRODUCTION_LOTS.sheet_name,
    "Regression",
    "Univariate",
)


def _verbose_checkpoint(verbose: bool, start_time: float, label: str) -> None:
    """Print one verbose progress checkpoint with elapsed seconds."""
    if verbose:
        print(f"  {label:<28} {time.monotonic() - start_time:6.1f}s", flush=True)


def _verification_calc_sheet_names(
    *, skip_regression: bool = False
) -> tuple[str, ...]:
    """Return workbook sheets that must be force-calculated before verify."""
    names = _VERIFY_CALC_SHEET_NAMES
    if skip_regression:
        # A test-model artifact may have no Regression / Mileage / Production
        # Lots sheets — drop them from the force-calc set so their absence is
        # not treated as a missing required sheet.
        names = tuple(
            name
            for name in names
            if name
            not in ("Regression", MILEAGE.sheet_name, PRODUCTION_LOTS.sheet_name)
        )
    return names


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


def _univariate_verification_action(
    workbook_sheet_names: set[str],
) -> str:
    """Decide what the Univariate stage of verify should do.

    Two outcomes:

    * ``"warn"`` — the workbook does not have a Univariate sheet. Not a QC
      failure, and not announced a second time here:
      ``_calculate_verification_sheets`` already warned about the missing
      sheet, so the verify stage only has to decline to run the comparison.
    * ``"check"`` — run the Univariate comparison.

    Parameters
    ----------
    workbook_sheet_names : set[str]
        Sheet names present in the workbook under verification.

    Returns
    -------
    str
        One of ``"warn"``, ``"check"``.
    """
    if UNIVARIATE_SHEET_NAME not in workbook_sheet_names:
        return "warn"
    return "check"


def _calculate_verification_sheets(
    workbook: xw.Book,
    verbose: bool,
    phase_start: float,
    *,
    skip_regression: bool = False,
) -> None:
    sheet_names = _verification_calc_sheet_names(
        skip_regression=skip_regression,
    )
    workbook_sheet_names = {sheet.name for sheet in workbook.sheets}
    missing_sheets = [name for name in sheet_names if name not in workbook_sheet_names]

    # A sheet named here can be legitimately absent: the Univariate sheet
    # may not be present in a test-model artifact. Warn and skip calculating
    # it rather than crashing. Any other missing sheet indicates a real
    # build problem and still hard-fails.
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
    """Load one of the ``tools/`` inspector scripts by path.

    ``tools/`` is not a package and is not on ``sys.path``, so the two
    inspectors are loaded from their file. That makes this module — like
    ``write_sheet_csv_dataset``, which reaches ``sample_data/`` the same
    checkout-relative way — a repo-checkout facility, not something the
    installed wheel (``packages = ["lambda_catalog"]``) can run. Say so when
    the path is absent rather than surfacing a bare ``FileNotFoundError``
    from ``exec_module``.
    """
    if not path.exists():
        raise RuntimeError(
            f"Could not load {module_name}: {path} does not exist. "
            "The spec-driven verifier runs from a repository checkout — it "
            "needs the tools/ inspectors and the sample_data/ CSVs, neither "
            "of which ships in the wheel."
        )
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load {module_name} from {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _verify_life_expectancy_developed_country(
    workbook: xw.Book,
    csv_path: Path,
    failures: list[str],
) -> None:
    expected_flags = calculate_developed_country_flags(csv_path)
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
    derived_header = LIFE_EXPECTANCY.derived_header
    if derived_header is None:
        _report_qc_failure(
            failures,
            "[Life Expectancy Data] expected 'Developed Country after 2013' "
            "column but the dataset config ships no derived header.",
        )
        return
    col_idx = life_expectancy_headers.index(derived_header)
    for row_offset, expected in enumerate(expected_flags, start=1):
        row = (
            life_expectancy_rows[row_offset]
            if row_offset < len(life_expectancy_rows)
            else []
        )
        actual = _normalize_excel_bool(
            row[col_idx] if col_idx < len(row) else None
        )
        if actual is not expected:
            _report_qc_failure(
                failures,
                f"[Life Expectancy Data] row={row_offset + 1} "
                f"stat='Developed Country after 2013': "
                f"expected={expected!r}, excel_calc={actual!r}",
            )


def verify_test_sheets(
    workbook: xw.Book,
    regression_sheet_configs: list | None,
    csv_path: Path,
    verbose: bool = False,
    *,
    skip_regression: bool = False,
    failures_out: list[str] | None = None,
) -> None:
    """Compare live workbook output against Python-computed QC oracle values.

    Parameters
    ----------
    workbook : xw.Book
        The live xlwings workbook to verify.
    regression_sheet_configs : list | None
        Per-config Python oracle (from ``build_regression_spec_qc_configs``).
        May be ``None`` when ``skip_regression`` is set (a test-model
        artifact may have no Regression sheet to compare against).
    csv_path : Path
        Path to the Life Expectancy CSV used for the "Developed Country
        after 2013" derived-column comparison.
    verbose : bool
        Print per-phase checkpoints to stdout.
    skip_regression : bool
        When True, skip every Regression-side check (the Regression
        spec-block check and the Regression DF comparison against
        ``regression_sheet_configs``). Used when verifying a workbook with
        no Regression sheet. Defaults to False. The Life Expectancy
        "Developed Country after 2013" check and the Univariate sheet check
        still run.
    failures_out : list[str] | None
        Optional external list. When supplied, every captured failure message
        is appended to it before the function raises on drift. Used by the
        build scripts' ``--verify`` to populate a ``VerifyReport`` for
        downstream structured rendering. The internal ``failures`` list is
        the source of truth; this is a write-through mirror.
    """
    failures: list[str] = []
    phase_start = time.monotonic()
    _calculate_verification_sheets(
        workbook,
        verbose,
        phase_start,
        skip_regression=skip_regression,
    )
    _verify_life_expectancy_developed_country(workbook, csv_path, failures)
    if not skip_regression:
        _verbose_checkpoint(verbose, phase_start, "Verify: reg spec block start")
        # No csv_path: the spec-block verifier resolves the dataset from the
        # sheet's own Source_Table and loads that dataset's CSV. Passing
        # mileage_path here was the bug — it forced Auto MPG expectations onto
        # whatever the workbook actually shipped.
        for failure in read_regression_spec_block_failures(workbook):
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
            (
                "predictors",
                ["config_name", "allow_intercept", "predictor_name", "stat_name"],
            ),
            (
                "coefficients",
                ["config_name", "allow_intercept", "term_name", "stat_name"],
            ),
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
    univariate_action = _univariate_verification_action(
        workbook_sheet_names
    )
    # ``"warn"`` produces no output here: _calculate_verification_sheets()
    # already printed the missing-sheet warning, which is where this
    # function's docstring says that warning lives. Printing a second one
    # would duplicate it.
    if univariate_action == "check":
        uv_mod = _load_module(
            "inspect_univariate_sheet",
            ROOT_DIR / "tools" / "inspect_univariate_sheet.py",
        )
        _verbose_checkpoint(verbose, phase_start, "Verify: univariate start")
        for failure in uv_mod.read_univariate_failures(workbook, csv_path):
            _report_qc_failure(failures, failure)
        _verbose_checkpoint(verbose, phase_start, "Verify: univariate done")

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

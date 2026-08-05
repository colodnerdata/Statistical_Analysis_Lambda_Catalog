"""Inspect the Regression worksheet against Python-computed expected values.

Writes each QC case's spec block onto the single ``Regression`` sheet in
turn, forces a recalculation, reads every output zone, and compares against
the Python oracle.

This is the LEGACY single-sheet path, and it stays because it verifies the
sheet users actually get: the production ``Regression`` sheet, driven through
the same spec-block inputs a human would type. Its cost is that every case
mutates one shared sheet, so cases can only be checked serially and a
failure leaves nothing to open and look at. The one-sheet-per-model artifact
(``build_test_models.py`` / ``tools/inspect_test_model_sheets.py``) is the
complement: every case gets its own sheet, built once, read without writing.

Both use ``lambda_catalog/regression_spec_sheet_io.py`` for the writing and
the reading, so they cannot disagree about what a case is.

Usage:
    python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx
    python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx --mileage path/to/data.csv
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd  # type: ignore[import-untyped]  # pyright: ignore[reportMissingTypeStubs]
import xlwings as xw

from lambda_catalog.analyze_mileage import DEFAULT_INPUT_CSV
from lambda_catalog.analyze_regression_spec import (
    RegressionSpecExpected,
    build_regression_spec_qc_configs,
)
from lambda_catalog.regression_spec_sheet_io import (
    apply_spec_case,
    read_case_comparison_rows,
    set_prediction_inputs,
)
from lambda_catalog.workbook_builder import XL_CALCULATION_MANUAL
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_csv_dataset import MILEAGE
from lambda_catalog.write_sheet_regression import REGRESSION_SHEET_NAME

DATA_SHEET_NAME = MILEAGE.sheet_name
DATA_TABLE_NAME = MILEAGE.table_name

# ── Tolerance (shared numeric QC tolerance) ───────────────────────────
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6

# ── DataFrame column definitions ─────────────────────────────────────────────
_DF_BASE = ["config_name", "allow_intercept", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_PRED = ["config_name", "allow_intercept", "predictor_name", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_COEF = ["config_name", "allow_intercept", "term_name", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_RESID = ["config_name", "allow_intercept", "row_idx", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]

_SECTION_COLUMNS = {
    "scalars": _DF_BASE,
    "predictors": _DF_PRED,
    "coefficients": _DF_COEF,
    "prediction_interval": _DF_BASE,
    "residuals": _DF_RESID,
}


def _delete_table_column_if_present(data_sheet: xw.Sheet, column_name: str) -> None:
    table = data_sheet.api.ListObjects(DATA_TABLE_NAME)
    try:
        table.ListColumns(column_name).Delete()
    except Exception:  # pylint: disable=broad-exception-caught
        pass


def _apply_extra_columns(
    data_sheet: xw.Sheet,
    expected: RegressionSpecExpected,
    all_extra_names: set[str],
) -> None:
    for name in all_extra_names:
        _delete_table_column_if_present(data_sheet, name)
    table = data_sheet.api.ListObjects(DATA_TABLE_NAME)
    for extra in expected.case.extra_columns:
        column = table.ListColumns.Add()
        column.Name = extra.name
        column.DataBodyRange.Formula = extra.excel_formula
    data_sheet.api.Calculate()


def read_regression_df(
    workbook: xw.Book,
    regression_sheet_configs: list[RegressionSpecExpected],
) -> dict[str, pd.DataFrame]:
    """Set spec toggles, read all zones, and compare against expected values.

    Returns a dict with keys 'scalars', 'predictors', 'coefficients',
    'residuals', 'prediction_interval', each mapping to a comparison DataFrame.
    """
    sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    data_sheet = workbook.sheets[DATA_SHEET_NAME]
    all_extra_names = {
        extra.name
        for expected in regression_sheet_configs
        for extra in expected.case.extra_columns
    }
    sections: dict[str, list[dict]] = {key: [] for key in _SECTION_COLUMNS}

    # The per-config loop below makes ~150-250 individual cell writes (spec
    # clear/rewrite, prediction inputs, extra-column formulas) before each
    # deliberate sheet.api.Calculate(). Excel's automatic recalculation engine
    # runs synchronously on every single COM write whenever Calculation is not
    # Manual, and ScreenUpdating left on lets every write also dirty the
    # workbook's chart objects (wired to volatile OFFSET named ranges per
    # write_sheet_regression.py). Across every QC config that is thousands of
    # redundant full recalculations/redraws instead of the one-per-config the
    # code below intends — this is what turns the loop from seconds into
    # (observed) hours. Suspend both for the whole loop, matching the
    # Manual-during-writes convention build_qc.py/build_production.py already
    # use, and restore the caller's settings afterward.
    app = workbook.app
    previous_calculation = app.api.Calculation
    previous_screen_updating = app.api.ScreenUpdating
    app.api.Calculation = XL_CALCULATION_MANUAL
    app.api.ScreenUpdating = False
    try:
        for expected in regression_sheet_configs:
            # Set source-table fixture columns, full spec, prediction inputs.
            _apply_extra_columns(data_sheet, expected, all_extra_names)
            apply_spec_case(sheet, expected)
            set_prediction_inputs(
                sheet,
                expected.results.prediction_interval.pred_input_values,
                expected.design.constructed_column_transforms,
            )
            # Recalculate only the Regression sheet after changing the visible
            # inputs. A full dependency-graph rebuild here pulls in the entire
            # workbook for every QC config and can take several minutes.
            sheet.api.Calculate()

            for key, rows in read_case_comparison_rows(sheet, expected).items():
                sections[key].extend(rows)
    finally:
        try:
            app.api.ScreenUpdating = previous_screen_updating
        except OPEN_WORKBOOK_ERRORS:
            pass
        try:
            app.api.Calculation = previous_calculation
        except OPEN_WORKBOOK_ERRORS:
            pass

    return {
        key: pd.DataFrame(sections[key], columns=columns)
        for key, columns in _SECTION_COLUMNS.items()
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inspect the Regression sheet against Python-computed expected values."
    )
    parser.add_argument("workbook", type=Path, help="Path to the Excel workbook.")
    parser.add_argument("--mileage", type=Path, default=DEFAULT_INPUT_CSV)
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Error: workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)

    regression_sheet_configs = build_regression_spec_qc_configs(args.mileage)

    try:
        with xw.App(visible=False, add_book=False) as app:
            try:
                workbook = app.books.open(str(workbook_path))
            except OPEN_WORKBOOK_ERRORS as exc:
                raise_excel_access_error(workbook_path, "open", exc)
            try:
                dfs = read_regression_df(workbook, regression_sheet_configs)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "inspect", exc)

    for section, df in dfs.items():
        print(f"\n=== Regression / {section} ===")
        print(df.to_string(index=False))


if __name__ == "__main__":
    main()

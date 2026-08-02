"""Inspect the Regression worksheet against Python-computed expected values.

Sets the spec block's Include toggles (column C) and the Allow_Intercept
control (C2) for each QC configuration, forces recalculation, reads every
output zone, and compares against cached Python expected values.

Usage:
    python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx
    python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx --mileage path/to/data.csv
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path
from typing import Any

import pandas as pd  # type: ignore[import-untyped]  # pyright: ignore[reportMissingTypeStubs]
import xlwings as xw

from lambda_catalog.analyze_mileage import DEFAULT_INPUT_CSV
from lambda_catalog.analyze_regression_spec import (
    RegressionSpecExpected,
    build_regression_spec_qc_configs,
)
from lambda_catalog.inspection_compare import compare_values, to_float_or_none
from lambda_catalog.workbook_builder import XL_CALCULATION_MANUAL
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_csv_dataset import MILEAGE
from lambda_catalog.write_sheet_model_construction import (
    _C_INCLUDE as _C_SPEC_INCLUDE,
    _C_REFERENCE as _C_SPEC_REFERENCE,
    _C_ROLE as _C_SPEC_ROLE,
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
    _C_TRANSFORM as _C_SPEC_TRANSFORM,
    _C_TYPE as _C_SPEC_TYPE,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _INTERCEPT_ROW,
)
from lambda_catalog.write_sheet_regression import (  # noqa: F401  (re-exported layout)
    REGRESSION_SHEET_NAME,
    _C_AA,
    _C_AB,
    _C_AC,
    _C_AD,
    _C_AE,
    _C_AF,
    _C_AG,
    _C_AH,
    _C_AJ,
    _C_AK,
    _C_AL,
    _C_AN,
    _C_AO,
    _C_AP,
    _C_AQ,
    _C_AR,
    _C_AS,
    _C_AT,
    _C_AU,
    _C_AV,
    _C_AW,
    _C_AX,
    _C_S,
    _C_T,
    _C_U,
    _C_V,
    _C_W,
    _C_X,
    _C_Y,
)

DATA_SHEET_NAME = MILEAGE.sheet_name
DATA_TABLE_NAME = MILEAGE.table_name

# ── Tolerance (shared with inspect_test_sheets.py) ───────────────────────────
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6

# ── Column indices ────────────────────────────────────────────────────────────
# IMPORTED from write_sheet_regression, not restated. This file used to keep
# its own copy of the whole map "to match" the writer's — which meant the
# layout-break MAJOR silently pointed this inspector at the wrong columns
# until the copy was hand-updated. One source of truth, per the same rule the
# writer's own constants follow.

# ── Row positions (1-based) ───────────────────────────────────────────────────
_ROW_SUMMARY_FIRST = 3   # S3 spills constructed names; T3–Y3 spill the stats

_ROW_MULTIPLE_R = 4
_ROW_R_SQUARED = 5
_ROW_ADJ_R2 = 6
_ROW_SE_REG = 7
_ROW_OBS = 8

_ROW_PRESS = 4
_ROW_PRESS_R2 = 5
_ROW_MEAN_LEV = 6
_ROW_AIC = 7
_ROW_BIC = 8
_ROW_AICC = 9
_ROW_QQ_CORR = 10
_ROW_DURBIN_WATSON = 11

_ROW_ANOVA_REG = 15
_ROW_ANOVA_RES = 16
_ROW_ANOVA_TOT = 17

# v2.1 Fixed Effects Prediction Outputs shape (write_sheet_regression.py
# _write_prediction_interval/_write_prediction_inputs): a 9-row CI+PI box
# (rows 3-11: point/se_mean/se_new/t_crit/ci_lower/ci_upper/pi_lower/pi_upper/
# confidence) plus an FE Group selector and ybar_i/T_i readouts (rows 12-14),
# and PREDICTION INPUTS starting at row 16 (first predictor row 19, no more
# Intercept row). RegressionPredictionInterval / analyze_regression_sheet.py
# model this exact shape via Group_Prediction_Interval's group-mean-recovery
# formula — a no-FE case selects the constant "(all)" group, which collapses
# to the pre-v2.1 single-PI numbers exactly (see
# tests/test_group_prediction_interval.py).
_ROW_COEFF_DATA = 21   # AA21 spills coefficient labels (k+1 rows)
_ROW_PI_POINT = 3      # AK3 = point estimate
_ROW_PI_SE_MEAN = 4
_ROW_PI_SE_NEW = 5
_ROW_PI_T = 6
_ROW_PI_CI_LOWER = 7
_ROW_PI_CI_UPPER = 8
_ROW_PI_PI_LOWER = 9
_ROW_PI_PI_UPPER = 10
_ROW_PI_CONF = 11
_ROW_FE_GROUP = 12     # AK12 = FE Group selector (input — written per case)
_ROW_GROUP_MEAN = 13   # AK13 = Group Mean (y)
_ROW_GROUP_COUNT = 14  # AK14 = Group Count
_ROW_PRED_INPUT_FIRST = 19  # AK19 = first user-editable predictor value (was 13)
_ROW_PRED_INPUT_LAST = 62   # end of the guarded prefill band

_ROW_RESID_FIRST = 3   # residual output starts at row 3


# ── DataFrame column definitions ─────────────────────────────────────────────
_DF_BASE = ["config_name", "allow_intercept", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_PRED = ["config_name", "allow_intercept", "predictor_name", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_COEF = ["config_name", "allow_intercept", "term_name", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_RESID = ["config_name", "allow_intercept", "row_idx", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]


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


def _apply_spec_case(sheet: xw.Sheet, expected: RegressionSpecExpected) -> None:
    """Retarget Source_Table, then write C2 and the full visible spec block for one QC case."""
    # Source_Table is the ONE name that retargets which data sheet the spec
    # block/design matrix reads from (see
    # write_sheet_model_construction._set_sheet_scoped_names). Every case sets
    # this unconditionally (RegressionSpecCase.source_table_ref is not
    # Optional) so a case is never accidentally evaluated against whatever
    # dataset the PREVIOUS case in the loop left Source_Table pointing at.
    sheet.api.Names.Item("Source_Table").RefersTo = expected.case.source_table_ref
    sheet.range((_INTERCEPT_ROW, _C_SPEC_INCLUDE)).value = expected.case.allow_intercept

    # FE Group selector ($AK$12): always written explicitly to
    # expected.resolved_prediction_group (never left at whatever the
    # PREVIOUS case's write left behind, and never blank — typing a value
    # into this cell replaces its default-computing formula, so "leave it
    # alone" would silently carry the last case's group forward). This is
    # the fixed cell above the variable-size Prediction Inputs band (rows
    # 19+) that the group choice belongs in, not a new slot below it.
    sheet.range(_ROW_FE_GROUP, _C_AK).value = expected.resolved_prediction_group

    # Clear only the spec rows (plus one blank row) so we don't wipe the
    # Sequence Spacing block that lives under the spec on the Regression sheet.
    from lambda_catalog.write_sheet_model_construction import _LAST_DATA_ROW as _SPEC_LAST_DATA_ROW

    last_row = max(
        _SPEC_LAST_DATA_ROW + 1,
        _SPEC_FIRST_DATA_ROW + len(expected.case.spec) - 1,
    )
    for row in range(_SPEC_FIRST_DATA_ROW, last_row + 1):
        for col in (
            _C_SPEC_ROLE,
            _C_SPEC_INCLUDE,
            _C_SPEC_TYPE,
            _C_SPEC_REFERENCE,
            _C_SPEC_SEQUENCE,
            _C_SPEC_TRANSFORM,
        ):
            sheet.range(row, col).clear_contents()

    for offset, variable in enumerate(expected.case.spec):
        row = _SPEC_FIRST_DATA_ROW + offset
        sheet.range(row, _C_SPEC_ROLE).value = variable.role
        sheet.range(row, _C_SPEC_INCLUDE).value = variable.include
        sheet.range(row, _C_SPEC_TYPE).value = variable.var_type
        sheet.range(row, _C_SPEC_REFERENCE).value = variable.reference
        sheet.range(row, _C_SPEC_SEQUENCE).value = variable.sequence
        sheet.range(row, _C_SPEC_TRANSFORM).value = variable.transform


def _set_pred_inputs(
    sheet: xw.Sheet,
    pred_input_values: tuple[float, ...],
    constructed_column_transforms: tuple[str, ...],
) -> None:
    """Write prediction inputs to AH13.., one per constructed column.

    The constructed columns are the selected predictors in table order —
    identical to the config's column order — so values are written
    positionally. The rest of the guarded prefill band is cleared so a
    previous config's values cannot linger.

    ``pred_input_values`` is the mean of each x_features column, which for
    a Log-transformed column is already in log space (build_spec_design's
    transform math logs the source values before this mean is taken). The
    sheet's AH cells are raw-value inputs (the row-3 prediction formula
    applies Ln_Positive itself, matching the confirmed UX — a user types
    the real-world value, never ln(x)), so a logged column's mean must be
    EXP'd back to input space here — the harness-side mirror of the AI19
    Training Mean spill's own geometric-mean fix in write_sheet_regression.py.
    """
    sheet.range(
        (_ROW_PRED_INPUT_FIRST, _C_AK), (_ROW_PRED_INPUT_LAST, _C_AK)
    ).clear_contents()
    for i, value in enumerate(pred_input_values):
        if constructed_column_transforms[i] == "Log":
            value = math.exp(value)
        sheet.range(_ROW_PRED_INPUT_FIRST + i, _C_AK).value = value


def _read_cell(sheet: xw.Sheet, row: int, col: int) -> float | None:
    val = sheet.range(row, col).value
    return to_float_or_none(val)


def _read_col(sheet: xw.Sheet, start_row: int, col: int, n_rows: int) -> list[float | None]:
    if n_rows <= 0:
        return []
    rng = sheet.range((start_row, col), (start_row + n_rows - 1, col))
    raw: list[Any] = rng.value if n_rows > 1 else [rng.value]
    return [to_float_or_none(v) for v in raw]


def _read_block(
    sheet: xw.Sheet, start_row: int, col1: int, col2: int, n_rows: int
) -> list[list[float | None]]:
    """Read a rectangular block; returns list of rows."""
    if n_rows <= 0:
        return []
    rng = sheet.range((start_row, col1), (start_row + n_rows - 1, col2))
    raw = rng.value
    if n_rows == 1:
        raw = [raw]
    return [
        [to_float_or_none(v) for v in row]
        for row in raw
    ]


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
    scalar_rows: list[dict] = []
    predictor_rows: list[dict] = []
    coeff_rows: list[dict] = []
    resid_rows: list[dict] = []
    pi_rows: list[dict] = []

    # The per-config loop below makes ~150-250 individual cell writes (spec
    # clear/rewrite, prediction inputs, extra-column formulas) before each
    # deliberate sheet.api.Calculate(). Excel's automatic recalculation engine
    # runs synchronously on every single COM write whenever Calculation is not
    # Manual, and ScreenUpdating left on lets every write also dirty the
    # workbook's chart objects (wired to volatile OFFSET named ranges per
    # write_sheet_regression.py). Across ~11 QC configs that is thousands of
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
            config_name = expected.case.name
            allow_intercept = expected.case.allow_intercept
            results = expected.results
            summary = results.summary
            vectors = results.vectors
            ps = results.predictor_summary
            fr = results.full_residuals
            pi = results.prediction_interval
            k = len(ps.predictor_names)
            n = summary.observations

            # Set source-table fixture columns, full spec, and prediction inputs.
            _apply_extra_columns(data_sheet, expected, all_extra_names)
            _apply_spec_case(sheet, expected)
            _set_pred_inputs(
                sheet, pi.pred_input_values, expected.design.constructed_column_transforms
            )
            # Recalculate only the Regression sheet after changing the visible
            # inputs. A full dependency-graph rebuild here pulls in the entire
            # workbook for every QC config and can take several minutes.
            sheet.api.Calculate()

            # ── Scalars: Regression Statistics (column AB) ────────────────────
            scalar_specs: list[tuple[str, float, int, int]] = [
                ("Multiple_R",    summary.multiple_r,    _ROW_MULTIPLE_R, _C_AB),
                ("R_Squared",     summary.r_squared,     _ROW_R_SQUARED,  _C_AB),
                ("Adjusted_R_Squared",   summary.adjusted_r2,   _ROW_ADJ_R2,     _C_AB),
                ("SE_Regression", summary.se_regression, _ROW_SE_REG,     _C_AB),
                ("Observations",  float(summary.observations), _ROW_OBS,  _C_AB),
            ]

            # Diagnostics (column AE)
            press_r2 = 1.0 - summary.press / summary.ss_total
            mean_lev = (summary.df_regression + (1 if allow_intercept else 0)) / summary.observations
            ms_reg = summary.ss_regression / summary.df_regression
            ms_res = summary.ss_residual / summary.df_residual

            scalar_specs += [
                ("PRESS",         summary.press,     _ROW_PRESS,    _C_AE),
                ("PRESS_R2",      press_r2,          _ROW_PRESS_R2, _C_AE),
                ("Mean_Leverage", mean_lev,          _ROW_MEAN_LEV, _C_AE),
                ("AIC",           summary.aic,       _ROW_AIC,      _C_AE),
                ("BIC",           summary.bic,       _ROW_BIC,      _C_AE),
                ("AICc",          summary.aicc,      _ROW_AICC,     _C_AE),
                ("QQ_Correlation",summary.qq_correlation, _ROW_QQ_CORR, _C_AE),
                ("Durbin_Watson",  summary.durbin_watson, _ROW_DURBIN_WATSON, _C_AE),
            ]

            # ANOVA
            scalar_specs += [
                ("Regression_Degrees_Of_Freedom",  float(summary.df_regression), _ROW_ANOVA_REG, _C_AB),
                ("SS_Regression",  summary.ss_regression,        _ROW_ANOVA_REG, _C_AC),
                ("MS_Regression",  ms_reg,                       _ROW_ANOVA_REG, _C_AD),
                ("F_Statistic",         summary.f_stat,               _ROW_ANOVA_REG, _C_AE),
                ("F_Statistic_P_Value",      summary.p_value_f,            _ROW_ANOVA_REG, _C_AF),
                ("Residual_Degrees_Of_Freedom",    float(summary.df_residual),   _ROW_ANOVA_RES, _C_AB),
                ("SS_Residual",    summary.ss_residual,          _ROW_ANOVA_RES, _C_AC),
                ("MS_Residual",    ms_res,                       _ROW_ANOVA_RES, _C_AD),
                ("Total_Degrees_Of_Freedom",       float(summary.df_total),      _ROW_ANOVA_TOT, _C_AB),
                ("SS_Total",       summary.ss_total,             _ROW_ANOVA_TOT, _C_AC),
            ]

            for stat_name, exp_val, row, col in scalar_specs:
                xl_val = _read_cell(sheet, row, col)
                diff, fdd_val = compare_values(exp_val, xl_val)
                scalar_rows.append({
                    "config_name": config_name,
                    "allow_intercept": allow_intercept,
                    "stat_name": stat_name,
                    "expected": exp_val,
                    "excel_calc": xl_val,
                    "abs_diff": diff,
                    "first_digit_deviation": fdd_val,
                })

            # ── Predictor Summary (columns T–Y, rows 3 to 3+k-1) ─────────────
            pred_stat_names = ["Pearson_R", "Spearman_R", "Skewness", "Kurtosis", "GVIF", "Tolerance"]
            pred_exp_tuples = [ps.pearson_r, ps.spearman_r, ps.skewness, ps.kurtosis, ps.gvif, ps.tolerance]
            pred_col_indices = [_C_T, _C_U, _C_V, _C_W, _C_X, _C_Y]

            for stat_name, exp_tuple, col in zip(pred_stat_names, pred_exp_tuples, pred_col_indices):
                xl_vals = _read_col(sheet, _ROW_SUMMARY_FIRST, col, k)
                for j, (exp_val, xl_val) in enumerate(zip(exp_tuple, xl_vals)):
                    diff, fdd_val = compare_values(exp_val, xl_val)
                    predictor_rows.append({
                        "config_name": config_name,
                        "allow_intercept": allow_intercept,
                        "predictor_name": ps.predictor_names[j],
                        "stat_name": stat_name,
                        "expected": exp_val,
                        "excel_calc": xl_val,
                        "abs_diff": diff,
                        "first_digit_deviation": fdd_val,
                    })

            # ── Coefficients (columns Y–AD, rows 21 to 21+k) ──────────────────
            # Intercept models: Coefficients() spills k+1 rows (intercept first);
            # read all k+1 and compare directly against (Intercept, pred1..predk).
            # No-intercept models prepend one blank row so predictor rows align with
            # intercept models. Drop that one display row before comparison.
            n_coef_rows = k + 1
            coef_stat_names = ["Coefficients", "SE_Coefficients", "T_Statistics", "P_Values", "Confidence_Interval_Lower", "Confidence_Interval_Upper"]
            coef_col_indices = [_C_AB, _C_AC, _C_AD, _C_AE, _C_AF, _C_AG]
            coef_exp_tuples = [
                vectors.coefficients, vectors.std_errors, vectors.t_stats,
                vectors.p_values, vectors.ci_lower, vectors.ci_upper,
            ]

            for stat_name, exp_tuple, col in zip(coef_stat_names, coef_exp_tuples, coef_col_indices):
                xl_vals_all = _read_col(sheet, _ROW_COEFF_DATA, col, n_coef_rows)
                xl_vals = xl_vals_all if allow_intercept else xl_vals_all[1:]
                for i, (exp_val, xl_val) in enumerate(zip(exp_tuple, xl_vals)):
                    term = vectors.term_names[i]
                    diff, fdd_val = compare_values(exp_val, xl_val)
                    coeff_rows.append({
                        "config_name": config_name,
                        "allow_intercept": allow_intercept,
                        "term_name": term,
                        "stat_name": stat_name,
                        "expected": exp_val,
                        "excel_calc": xl_val,
                        "abs_diff": diff,
                        "first_digit_deviation": fdd_val,
                    })

            beta_vals_all = _read_col(sheet, _ROW_COEFF_DATA, _C_AH, n_coef_rows)
            beta_vals = beta_vals_all[1:]
            for i, (exp_val, xl_val) in enumerate(zip(vectors.beta_weights, beta_vals)):
                term = ps.predictor_names[i]
                diff, fdd_val = compare_values(exp_val, xl_val)
                coeff_rows.append({
                    "config_name": config_name,
                    "allow_intercept": allow_intercept,
                    "term_name": term,
                    "stat_name": "Beta_Weights",
                    "expected": exp_val,
                    "excel_calc": xl_val,
                    "abs_diff": diff,
                    "first_digit_deviation": fdd_val,
                })

            # ── Prediction Interval (column AK, rows 3–14) ────────────────────
            # Group_Prediction_Interval's 9-value CI+PI box plus the Group
            # Mean/Count readouts. $AK$12 (the FE Group selector) was already
            # written to expected.resolved_prediction_group in _apply_spec_case,
            # above the variable-size Prediction Inputs band written below.
            pi_specs: list[tuple[str, float]] = [
                ("Point_Estimate",   pi.point_estimate),
                ("SE_Mean",          pi.se_mean),
                ("SE_New",           pi.se_new),
                ("T_Critical",       pi.t_critical),
                ("CI_Lower",         pi.ci_lower),
                ("CI_Upper",         pi.ci_upper),
                ("PI_Lower",         pi.pi_lower),
                ("PI_Upper",         pi.pi_upper),
                ("Confidence_Level", pi.confidence_level),
            ]
            pi_rows_data = _read_col(sheet, _ROW_PI_POINT, _C_AK, 9)
            for (stat_name, exp_val), xl_val in zip(pi_specs, pi_rows_data):
                diff, fdd_val = compare_values(exp_val, xl_val)
                pi_rows.append({
                    "config_name": config_name,
                    "allow_intercept": allow_intercept,
                    "stat_name": stat_name,
                    "expected": exp_val,
                    "excel_calc": xl_val,
                    "abs_diff": diff,
                    "first_digit_deviation": fdd_val,
                })

            for stat_name, exp_val, row in [
                ("Group_Mean", pi.group_mean, _ROW_GROUP_MEAN),
                ("Group_Count", float(pi.group_count), _ROW_GROUP_COUNT),
            ]:
                xl_val = _read_cell(sheet, row, _C_AK)
                diff, fdd_val = compare_values(exp_val, xl_val)
                pi_rows.append({
                    "config_name": config_name,
                    "allow_intercept": allow_intercept,
                    "stat_name": stat_name,
                    "expected": exp_val,
                    "excel_calc": xl_val,
                    "abs_diff": diff,
                    "first_digit_deviation": fdd_val,
                })

            # ── Residual Output (columns AL–AU, rows 3 to 3+n-1) ──────────────
            resid_stat_names = [
                "Dependent_Variable", "Predictions", "Residuals",
                "Hat_Diagonal", "Studentized_Residuals", "Cooks_Distance",
                "Normal_Scores_Ranked", "Studentized_Residuals_Ranked",
                "Scale_Location", "PRESS_Residual",
            ]
            resid_exp_tuples = [
                fr.dependent_var, fr.predictions, fr.residuals,
                fr.hat_diagonal, fr.studentized_residuals, fr.cooks_distance,
                fr.normal_scores_ranked, fr.studentized_residuals_ranked,
                fr.scale_location, fr.loocv_residuals,  # loocv_residuals = e/(1-h) = PRESS
            ]
            # Block: columns AL(38) through AU(47) = 10 columns, n rows
            block = _read_block(sheet, _ROW_RESID_FIRST, _C_AO, _C_AX, n)
            for row_idx, xl_row in enumerate(block):
                for stat_name, exp_tuple, xl_val in zip(resid_stat_names, resid_exp_tuples, xl_row):
                    exp_val: float | None = float(exp_tuple[row_idx]) if row_idx < len(exp_tuple) else None
                    diff, fdd_val = compare_values(exp_val, xl_val)
                    resid_rows.append({
                        "config_name": config_name,
                        "allow_intercept": allow_intercept,
                        "row_idx": row_idx + 1,
                        "stat_name": stat_name,
                        "expected": exp_val,
                        "excel_calc": xl_val,
                        "abs_diff": diff,
                        "first_digit_deviation": fdd_val,
                    })
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
        "scalars": pd.DataFrame(scalar_rows, columns=_DF_BASE),
        "predictors": pd.DataFrame(predictor_rows, columns=_DF_PRED),
        "coefficients": pd.DataFrame(coeff_rows, columns=_DF_COEF),
        "prediction_interval": pd.DataFrame(pi_rows, columns=_DF_BASE),
        "residuals": pd.DataFrame(resid_rows, columns=_DF_RESID),
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

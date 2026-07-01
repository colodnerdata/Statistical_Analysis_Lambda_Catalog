"""Inspect the Regression worksheet against Python-computed expected values.

Sets B-column predictor toggles for each QC configuration, forces
recalculation, reads every output zone, and compares against cached
Python expected values.

Usage:
    python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx
    python tools/inspect_regression_sheet.py Lambda_Library_QC.xlsx --csv path/to/data.csv
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import pandas as pd  # type: ignore[import-untyped]  # pyright: ignore[reportMissingTypeStubs]
import xlwings as xw

from lambda_catalog.analyze_life_expectancy import DEFAULT_INPUT_CSV, FEATURE_COLUMNS
from lambda_catalog.analyze_regression_sheet import (
    RegressionSheetResults,
    build_regression_sheet_qc_configs,
)
from lambda_catalog.inspection_compare import compare_values, to_float_or_none
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_regression import REGRESSION_SHEET_NAME

# ── Tolerance (shared with inspect_test_sheets.py) ───────────────────────────
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6

# ── Column indices (1-based, must match write_sheet_regression.py constants) ─
_C_B = 2    # "In linear model?" toggles; B2 = Allow_Intercept
_C_D = 4    # predictor names (predictor summary)
_C_E = 5    # Pearson R
_C_F = 6    # Spearman R
_C_G = 7    # Skewness
_C_H = 8    # Kurtosis
_C_I = 9    # VIF
_C_J = 10   # Tolerance
_C_M = 13   # regression stat values / ANOVA df + coeff values
_C_N = 14   # ANOVA SS / coeff SE
_C_O = 15   # ANOVA MS / coeff t-stat
_C_P = 16   # diagnostics values / ANOVA F / coeff p-value
_C_Q = 17   # ANOVA Sig F / coeff CI lower
_C_R = 18   # coeff CI upper
_C_U = 21   # prediction interval values + prediction input values
_C_Y = 25   # Y (filtered dependent var)
_C_Z = 26   # Predicted Y
_C_AA = 27  # Residuals
_C_AB = 28  # LOOCV residual
_C_AC = 29  # Hat Diagonal
_C_AD = 30  # Studentized Residuals
_C_AE = 31  # Cook's Distance
_C_AF = 32  # Normal Scores Ranked
_C_AG = 33  # Studentized Residuals Ranked

# ── Row positions (1-based) ───────────────────────────────────────────────────
_ROW_ALLOW_INTERCEPT = 2
_ROW_PREDICTOR_FIRST = 3     # B3 = first predictor toggle; A3 spills predictor names
_ROW_PREDICTOR_LAST = 20     # B20 = last predictor toggle (18 predictors: rows 3–20)

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

_ROW_ANOVA_REG = 15
_ROW_ANOVA_RES = 16
_ROW_ANOVA_TOT = 17

_ROW_COEFF_DATA = 21   # L21 spills coefficients (k+1 rows)
_ROW_PI_POINT = 3      # U3 = point estimate
_ROW_PI_SE = 4
_ROW_PI_T = 5
_ROW_PI_LOWER = 6
_ROW_PI_UPPER = 7
_ROW_PI_CONF = 8
_ROW_PRED_INPUT_FIRST = 13  # U13 = first user-editable predictor value

_ROW_RESID_FIRST = 3   # residual output starts at row 3


# ── DataFrame column definitions ─────────────────────────────────────────────
_DF_BASE = ["config_name", "allow_intercept", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_PRED = ["config_name", "allow_intercept", "predictor_name", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_COEF = ["config_name", "allow_intercept", "term_name", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_DF_RESID = ["config_name", "allow_intercept", "row_idx", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]


def _set_toggles(
    sheet: xw.Sheet,
    feature_columns: list[str],
    allow_intercept: bool,
) -> None:
    """Write B2 (Allow_Intercept) and B3:B20 (per-predictor on/off)."""
    sheet.range(_ROW_ALLOW_INTERCEPT, _C_B).value = allow_intercept
    selected = set(feature_columns)
    for i, name in enumerate(FEATURE_COLUMNS):
        row = _ROW_PREDICTOR_FIRST + i
        sheet.range(row, _C_B).value = name in selected


def _set_pred_inputs(
    sheet: xw.Sheet,
    pred_input_values: tuple[float, ...],
) -> None:
    """Write training-data column means to U13:U(12+k) for the prediction interval."""
    k = len(pred_input_values)
    for i, val in enumerate(pred_input_values):
        sheet.range(_ROW_PRED_INPUT_FIRST + i, _C_U).value = val
    # Clear any leftover values from a longer prior config
    for row in range(_ROW_PRED_INPUT_FIRST + k, _ROW_PRED_INPUT_FIRST + len(FEATURE_COLUMNS)):
        sheet.range(row, _C_U).value = 0.0


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
    regression_sheet_configs: list[tuple[str, bool, RegressionSheetResults]],
) -> dict[str, pd.DataFrame]:
    """Set toggles, read all zones, and compare against expected values.

    Returns a dict with keys 'scalars', 'predictors', 'coefficients',
    'residuals', 'prediction_interval', each mapping to a comparison DataFrame.
    """
    sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    scalar_rows: list[dict] = []
    predictor_rows: list[dict] = []
    coeff_rows: list[dict] = []
    resid_rows: list[dict] = []
    pi_rows: list[dict] = []

    for config_name, allow_intercept, results in regression_sheet_configs:
        summary = results.summary
        vectors = results.vectors
        ps = results.predictor_summary
        fr = results.full_residuals
        pi = results.prediction_interval
        k = len(ps.predictor_names)
        n = summary.observations

        # Set toggles and prediction inputs, then recalculate
        _set_toggles(sheet, list(ps.predictor_names), allow_intercept)
        _set_pred_inputs(sheet, pi.pred_input_values)
        # Excel does not reliably track dependencies through the dynamic x_s
        # worksheet name after toggle changes, so rebuild the dependency graph.
        workbook.app.api.CalculateFullRebuild()

        # ── Scalars: Regression Statistics (column M) ─────────────────────
        scalar_specs: list[tuple[str, float, int, int]] = [
            ("Multiple_R",    summary.multiple_r,    _ROW_MULTIPLE_R, _C_M),
            ("R_squared",     summary.r_squared,     _ROW_R_SQUARED,  _C_M),
            ("Adjusted_R2",   summary.adjusted_r2,   _ROW_ADJ_R2,     _C_M),
            ("SE_Regression", summary.se_regression, _ROW_SE_REG,     _C_M),
            ("Observations",  float(summary.observations), _ROW_OBS,  _C_M),
        ]

        # Diagnostics (column P)
        press_r2 = 1.0 - summary.press / summary.ss_total
        mean_lev = (summary.df_regression + (1 if allow_intercept else 0)) / summary.observations
        ms_reg = summary.ss_regression / summary.df_regression
        ms_res = summary.ss_residual / summary.df_residual

        scalar_specs += [
            ("PRESS",         summary.press,     _ROW_PRESS,    _C_P),
            ("PRESS_R2",      press_r2,          _ROW_PRESS_R2, _C_P),
            ("Mean_Leverage", mean_lev,          _ROW_MEAN_LEV, _C_P),
            ("AIC",           summary.aic,       _ROW_AIC,      _C_P),
            ("BIC",           summary.bic,       _ROW_BIC,      _C_P),
            ("AICc",          summary.aicc,      _ROW_AICC,     _C_P),
            ("QQ_Correlation",summary.qq_correlation, _ROW_QQ_CORR, _C_P),
        ]

        # ANOVA
        scalar_specs += [
            ("DF_Regression",  float(summary.df_regression), _ROW_ANOVA_REG, _C_M),
            ("SS_Regression",  summary.ss_regression,        _ROW_ANOVA_REG, _C_N),
            ("MS_Regression",  ms_reg,                       _ROW_ANOVA_REG, _C_O),
            ("F_Stat",         summary.f_stat,               _ROW_ANOVA_REG, _C_P),
            ("P_Value_F",      summary.p_value_f,            _ROW_ANOVA_REG, _C_Q),
            ("DF_Residual",    float(summary.df_residual),   _ROW_ANOVA_RES, _C_M),
            ("SS_Residual",    summary.ss_residual,          _ROW_ANOVA_RES, _C_N),
            ("MS_Residual",    ms_res,                       _ROW_ANOVA_RES, _C_O),
            ("DF_Total",       float(summary.df_total),      _ROW_ANOVA_TOT, _C_M),
            ("SS_Total",       summary.ss_total,             _ROW_ANOVA_TOT, _C_N),
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

        # ── Predictor Summary (columns E–J, rows 3 to 3+k-1) ─────────────
        pred_stat_names = ["Pearson_R", "Spearman_R", "Skewness", "Kurtosis", "VIF", "Tolerance"]
        pred_exp_tuples = [ps.pearson_r, ps.spearman_r, ps.skewness, ps.kurtosis, ps.vif, ps.tolerance]
        pred_col_indices = [_C_E, _C_F, _C_G, _C_H, _C_I, _C_J]

        for stat_name, exp_tuple, col in zip(pred_stat_names, pred_exp_tuples, pred_col_indices):
            xl_vals = _read_col(sheet, _ROW_PREDICTOR_FIRST, col, k)
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

        # ── Coefficients (columns M–R, rows 21 to 21+k) ──────────────────
        # Intercept models: Coefficients() spills k+1 rows (intercept first);
        # read all k+1 and compare directly against (Intercept, pred1..predk).
        # No-intercept models prepend one blank row so predictor rows align with
        # intercept models. Drop that one display row before comparison.
        n_coef_rows = k + 1
        coef_stat_names = ["Coefficients", "SE_Coefficients", "T_Stats", "P_Values", "CI_Lower", "CI_Upper"]
        coef_col_indices = [_C_M, _C_N, _C_O, _C_P, _C_Q, _C_R]
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

        # ── Prediction Interval (column U, rows 3–8) ─────────────────────
        pi_specs: list[tuple[str, float]] = [
            ("Point_Estimate",  pi.point_estimate),
            ("SE_Prediction",   pi.se_prediction),
            ("T_Critical",      pi.t_critical),
            ("Lower",           pi.lower),
            ("Upper",           pi.upper),
            ("Confidence_Level",pi.confidence_level),
        ]
        pi_rows_data = _read_col(sheet, _ROW_PI_POINT, _C_U, 6)
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

        # ── Residual Output (columns Y–AG, rows 3 to 3+n-1) ──────────────
        resid_stat_names = [
            "Dependent_Var", "Predictions", "Residuals", "LOOCV_Residual",
            "Hat_Diagonal", "Studentized_Residuals", "Cooks_Distance",
            "Normal_Scores_Ranked", "Studentized_Residuals_Ranked",
        ]
        resid_exp_tuples = [
            fr.dependent_var, fr.predictions, fr.residuals, fr.loocv_residuals,
            fr.hat_diagonal, fr.studentized_residuals, fr.cooks_distance,
            fr.normal_scores_ranked, fr.studentized_residuals_ranked,
        ]
        # Block: columns Y(25) through AG(33) = 9 columns, n rows
        block = _read_block(sheet, _ROW_RESID_FIRST, _C_Y, _C_AG, n)
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
    parser.add_argument("--csv", type=Path, default=DEFAULT_INPUT_CSV)
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Error: workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)

    regression_sheet_configs = build_regression_sheet_qc_configs(args.csv)

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

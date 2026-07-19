"""Inspect the Regression worksheet against Python-computed expected values.

Sets the spec block's Include toggles (column C) and the Allow_Intercept
control (C2) for each QC configuration, forces recalculation, reads every
output zone, and compares against cached Python expected values.

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

from lambda_catalog.analyze_life_expectancy import DEFAULT_INPUT_CSV
from lambda_catalog.analyze_regression_spec import (
    RegressionSpecExpected,
    build_regression_spec_qc_configs,
)
from lambda_catalog.inspection_compare import compare_values, to_float_or_none
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_life_expectancy_data import (
    SHEET_NAME as DATA_SHEET_NAME,
    TABLE_NAME as DATA_TABLE_NAME,
)
from lambda_catalog.write_sheet_model_construction import (
    _C_INCLUDE as _C_SPEC_INCLUDE,
    _C_REFERENCE as _C_SPEC_REFERENCE,
    _C_ROLE as _C_SPEC_ROLE,
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
    _C_TYPE as _C_SPEC_TYPE,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _INTERCEPT_ROW,
)
from lambda_catalog.write_sheet_regression import REGRESSION_SHEET_NAME

# ── Tolerance (shared with inspect_test_sheets.py) ───────────────────────────
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6

# ── Column indices (1-based, must match write_sheet_regression.py constants) ─
# Layout: A–L spec, M/N spec feedback + I Verdict overlay, O gap, P–V
# predictor summary, W gap, X–AE regression outputs, AF gap, AG–AI
# prediction outputs, AJ gap, AK–AV residual output. The constants below
# mirror write_sheet_regression.py's column letters (i.e. _C_P is column P).
_C_P = 16   # constructed column names (predictor summary)
_C_Q = 17   # Pearson R
_C_R = 18   # Spearman R
_C_S = 19   # Skewness
_C_T = 20   # Kurtosis
_C_U = 21   # VIF
_C_V = 22   # Tolerance
_C_X = 24   # labels (stats / ANOVA / coefficients)
_C_Y = 25   # stat values / ANOVA df / coefficient values
_C_Z = 26   # ANOVA SS / coefficient SE
_C_AA = 27  # diagnostics labels / ANOVA MS / coefficient t-stat
_C_AB = 28  # Predicted Variable label (AB2) / diagnostics values / ANOVA F / coefficient p-value
_C_AC = 29  # predicted variable readout (AC2) / ANOVA Sig F / coefficient CI lower
_C_AD = 30  # coefficient CI upper
_C_AE = 31  # Beta Weights
_C_AG = 33  # prediction interval labels / prediction input labels
_C_AH = 34  # prediction interval values / prediction input values
_C_AI = 35  # Training Mean
_C_AK = 37  # row identifiers (Row_Labels)
_C_AL = 38  # Y (filtered dependent var)
_C_AM = 39  # Predicted Y
_C_AN = 40  # Residuals
_C_AO = 41  # Hat Diagonal
_C_AP = 42  # Studentized Residuals
_C_AQ = 43  # Cook's Distance
_C_AR = 44  # Normal Scores Ranked
_C_AS = 45  # Studentized Residuals Ranked
_C_AT = 46  # Scale-Location
_C_AU = 47  # PRESS Residual

# ── Row positions (1-based) ───────────────────────────────────────────────────
_ROW_SUMMARY_FIRST = 3   # P3 spills constructed names; Q3–V3 spill the stats

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

_ROW_COEFF_DATA = 21   # X21 spills coefficient labels (k+1 rows)
_ROW_PI_POINT = 3      # AH3 = point estimate
_ROW_PI_SE = 4
_ROW_PI_T = 5
_ROW_PI_LOWER = 6
_ROW_PI_UPPER = 7
_ROW_PI_CONF = 8
_ROW_PRED_INPUT_FIRST = 13  # AH13 = first user-editable predictor value
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
    """Write C2 and the full visible spec block for one QC case."""
    sheet.range(_INTERCEPT_ROW, _C_SPEC_INCLUDE).value = expected.case.allow_intercept

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
        ):
            sheet.range(row, col).clear_contents()

    for offset, variable in enumerate(expected.case.spec):
        row = _SPEC_FIRST_DATA_ROW + offset
        sheet.range(row, _C_SPEC_ROLE).value = variable.role
        sheet.range(row, _C_SPEC_INCLUDE).value = variable.include
        sheet.range(row, _C_SPEC_TYPE).value = variable.var_type
        sheet.range(row, _C_SPEC_REFERENCE).value = variable.reference
        sheet.range(row, _C_SPEC_SEQUENCE).value = variable.sequence


def _set_pred_inputs(
    sheet: xw.Sheet,
    pred_input_values: tuple[float, ...],
) -> None:
    """Write prediction inputs to AH13.., one per constructed column.

    The constructed columns are the selected predictors in table order —
    identical to the config's column order — so values are written
    positionally. The rest of the guarded prefill band is cleared so a
    previous config's values cannot linger.
    """
    sheet.range(
        (_ROW_PRED_INPUT_FIRST, _C_AH), (_ROW_PRED_INPUT_LAST, _C_AH)
    ).clear_contents()
    for i, value in enumerate(pred_input_values):
        sheet.range(_ROW_PRED_INPUT_FIRST + i, _C_AH).value = value


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
        _set_pred_inputs(sheet, pi.pred_input_values)
        # Recalculate only the Regression sheet after changing the visible
        # inputs. A full dependency-graph rebuild here pulls in the entire
        # workbook for every QC config and can take several minutes.
        sheet.api.Calculate()

        # ── Scalars: Regression Statistics (column Y) ─────────────────────
        scalar_specs: list[tuple[str, float, int, int]] = [
            ("Multiple_R",    summary.multiple_r,    _ROW_MULTIPLE_R, _C_Y),
            ("R_Squared",     summary.r_squared,     _ROW_R_SQUARED,  _C_Y),
            ("Adjusted_R_Squared",   summary.adjusted_r2,   _ROW_ADJ_R2,     _C_Y),
            ("SE_Regression", summary.se_regression, _ROW_SE_REG,     _C_Y),
            ("Observations",  float(summary.observations), _ROW_OBS,  _C_Y),
        ]

        # Diagnostics (column AB)
        press_r2 = 1.0 - summary.press / summary.ss_total
        mean_lev = (summary.df_regression + (1 if allow_intercept else 0)) / summary.observations
        ms_reg = summary.ss_regression / summary.df_regression
        ms_res = summary.ss_residual / summary.df_residual

        scalar_specs += [
            ("PRESS",         summary.press,     _ROW_PRESS,    _C_AB),
            ("PRESS_R2",      press_r2,          _ROW_PRESS_R2, _C_AB),
            ("Mean_Leverage", mean_lev,          _ROW_MEAN_LEV, _C_AB),
            ("AIC",           summary.aic,       _ROW_AIC,      _C_AB),
            ("BIC",           summary.bic,       _ROW_BIC,      _C_AB),
            ("AICc",          summary.aicc,      _ROW_AICC,     _C_AB),
            ("QQ_Correlation",summary.qq_correlation, _ROW_QQ_CORR, _C_AB),
            ("Durbin_Watson",  summary.durbin_watson, _ROW_DURBIN_WATSON, _C_AB),
        ]

        # ANOVA
        scalar_specs += [
            ("Regression_Degrees_Of_Freedom",  float(summary.df_regression), _ROW_ANOVA_REG, _C_Y),
            ("SS_Regression",  summary.ss_regression,        _ROW_ANOVA_REG, _C_Z),
            ("MS_Regression",  ms_reg,                       _ROW_ANOVA_REG, _C_AA),
            ("F_Statistic",         summary.f_stat,               _ROW_ANOVA_REG, _C_AB),
            ("F_Statistic_P_Value",      summary.p_value_f,            _ROW_ANOVA_REG, _C_AC),
            ("Residual_Degrees_Of_Freedom",    float(summary.df_residual),   _ROW_ANOVA_RES, _C_Y),
            ("SS_Residual",    summary.ss_residual,          _ROW_ANOVA_RES, _C_Z),
            ("MS_Residual",    ms_res,                       _ROW_ANOVA_RES, _C_AA),
            ("Total_Degrees_Of_Freedom",       float(summary.df_total),      _ROW_ANOVA_TOT, _C_Y),
            ("SS_Total",       summary.ss_total,             _ROW_ANOVA_TOT, _C_Z),
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

        # ── Predictor Summary (columns Q–V, rows 3 to 3+k-1) ─────────────
        pred_stat_names = ["Pearson_R", "Spearman_R", "Skewness", "Kurtosis", "VIF", "Tolerance"]
        pred_exp_tuples = [ps.pearson_r, ps.spearman_r, ps.skewness, ps.kurtosis, ps.vif, ps.tolerance]
        pred_col_indices = [_C_Q, _C_R, _C_S, _C_T, _C_U, _C_V]

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
        coef_col_indices = [_C_Y, _C_Z, _C_AA, _C_AB, _C_AC, _C_AD]
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

        beta_vals_all = _read_col(sheet, _ROW_COEFF_DATA, _C_AE, n_coef_rows)
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

        # ── Prediction Interval (column AH, rows 3–8) ─────────────────────
        pi_specs: list[tuple[str, float]] = [
            ("Point_Estimate",  pi.point_estimate),
            ("SE_Prediction",   pi.se_prediction),
            ("T_Critical",      pi.t_critical),
            ("Lower",           pi.lower),
            ("Upper",           pi.upper),
            ("Confidence_Level",pi.confidence_level),
        ]
        pi_rows_data = _read_col(sheet, _ROW_PI_POINT, _C_AH, 6)
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
        block = _read_block(sheet, _ROW_RESID_FIRST, _C_AL, _C_AU, n)
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

    regression_sheet_configs = build_regression_spec_qc_configs(args.csv)

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

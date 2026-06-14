"""Compute Python expected values for every output zone of the Regression worksheet."""
from __future__ import annotations

from dataclasses import dataclass
from math import sqrt
from pathlib import Path
from statistics import NormalDist

import numpy as np
import statsmodels.api as sm  # type: ignore[import-untyped]
from scipy import stats as _scipy_stats  # type: ignore[import-untyped]

from .analyze_life_expectancy import (
    FEATURE_COLUMNS,
    RegressionSummary,
    RegressionVectors,
    _build_training_arrays,
    _fit_ols_model,
    _load_normalized_rows,
    _validate_required_headers,
    calculate_regression_summary,
    calculate_regression_vectors,
    DEFAULT_INPUT_CSV,
)


# ---------------------------------------------------------------------------
# QC test configurations: (config_name, allow_intercept, predictor_columns)
# Predictor subsets are non-consecutive to exercise the B-column toggle
# mechanism beyond the "first k" pattern already covered by test sheets.
# ---------------------------------------------------------------------------

_SPARSE_PREDICTORS = [
    "Adult Mortality",
    "BMI",
    "HIV/AIDS",
    "Schooling",
]

_MEDIUM_PREDICTORS = [
    "Adult Mortality",
    "Alcohol",
    "Hepatitis B",
    "Polio",
    "HIV/AIDS",
    "GDP",
    "thinness 1-19 years",
    "Schooling",
]

REGRESSION_QC_CONFIGS: list[tuple[str, bool, list[str]]] = [
    ("sparse_intercept",    True,  _SPARSE_PREDICTORS),
    ("sparse_no_intercept", False, _SPARSE_PREDICTORS),
    ("medium_intercept",    True,  _MEDIUM_PREDICTORS),
    ("medium_no_intercept", False, _MEDIUM_PREDICTORS),
    ("full_intercept",      True,  FEATURE_COLUMNS),
    ("full_no_intercept",   False, FEATURE_COLUMNS),
]


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class RegressionPredictorSummary:
    """Per-predictor summary statistics for the Predictor Summary zone (D–J)."""

    predictor_names: tuple[str, ...]
    pearson_r: tuple[float, ...]
    spearman_r: tuple[float, ...]
    skewness: tuple[float, ...]
    kurtosis: tuple[float, ...]
    vif: tuple[float, ...]
    tolerance: tuple[float, ...]


@dataclass(frozen=True)
class RegressionFullResiduals:
    """Per-observation diagnostics for the Residual Output zone (W–AI)."""

    predictions: tuple[float, ...]
    residuals: tuple[float, ...]
    loocv_predictions: tuple[float, ...]
    rank_fraction: tuple[float, ...]
    y_ranked: tuple[float, ...]
    normal_scores: tuple[float, ...]
    scaled_residuals: tuple[float, ...]
    scaled_residuals_ranked: tuple[float, ...]
    hat_diagonal: tuple[float, ...]
    studentized_residuals: tuple[float, ...]
    studentized_residuals_ranked: tuple[float, ...]
    cooks_distance: tuple[float, ...]


@dataclass(frozen=True)
class RegressionPredictionInterval:
    """Prediction interval at training-data column means (T–U zone, U3:U8)."""

    pred_input_values: tuple[float, ...]  # means of selected predictors, for writing to U13:U(12+k)
    point_estimate: float
    se_prediction: float
    t_critical: float
    lower: float
    upper: float
    confidence_level: float


@dataclass(frozen=True)
class RegressionSheetResults:
    """All Python expected values for one (predictor_set, intercept) configuration."""

    summary: RegressionSummary
    vectors: RegressionVectors
    predictor_summary: RegressionPredictorSummary
    full_residuals: RegressionFullResiduals
    prediction_interval: RegressionPredictionInterval


# ---------------------------------------------------------------------------
# Calculation
# ---------------------------------------------------------------------------

def calculate_regression_sheet_results(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    include_intercept: bool = True,
    feature_columns: list[str] | None = None,
    alpha: float = 0.05,
) -> RegressionSheetResults:
    """Fit OLS and compute expected values for every Regression sheet output zone."""
    columns = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    k = len(columns)
    input_path = input_csv_path.resolve()

    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers, columns)

    x_train, y_train, _, _ = _build_training_arrays(
        normalized_rows, include_intercept, columns, filter_columns=FEATURE_COLUMNS
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)
    n = len(y_train)

    # Feature columns only (no intercept), same filtered rows
    x_features = x_train[:, 1:] if include_intercept else x_train

    # ── hat matrix (shared across groups) ────────────────────────────────────
    xtx = x_train.T @ x_train
    xtx_inv = np.linalg.inv(xtx)
    z = x_train @ xtx_inv
    h = np.sum(z * x_train, axis=1)
    e = np.asarray(model.resid, dtype=np.float64)

    # ── Group 1: Scalar summary ───────────────────────────────────────────────
    summary = calculate_regression_summary(input_csv_path, include_intercept, columns)

    # ── Group 2: Coefficient vectors ─────────────────────────────────────────
    vectors = calculate_regression_vectors(input_csv_path, include_intercept, columns, alpha)

    # ── Group 3: Predictor summary ────────────────────────────────────────────
    pearson_r_vals: list[float] = []
    spearman_r_vals: list[float] = []
    skewness_vals: list[float] = []
    kurtosis_vals: list[float] = []
    vif_vals: list[float] = []

    for j in range(k):
        x_col = x_features[:, j]

        pearson_r_vals.append(float(np.corrcoef(x_col, y_train)[0, 1]))
        spearman_r_vals.append(float(_scipy_stats.spearmanr(x_col, y_train).statistic))
        # bias=False matches Excel SKEW (adjusted sample skewness)
        skewness_vals.append(float(_scipy_stats.skew(x_col, bias=False)))
        # fisher=True (excess), bias=False matches Excel KURT
        kurtosis_vals.append(float(_scipy_stats.kurtosis(x_col, fisher=True, bias=False)))

        if k == 1:
            vif_vals.append(1.0)
        else:
            # VIF: regress predictor j on all others WITH intercept — matches LAMBDA
            others = np.delete(x_features, j, axis=1)
            others_with_const = np.column_stack([np.ones(n), others])
            r2_j = float(sm.OLS(x_col, others_with_const).fit().rsquared)
            vif_vals.append(1.0 / (1.0 - r2_j))

    tolerance_vals = [1.0 / v for v in vif_vals]
    predictor_summary = RegressionPredictorSummary(
        predictor_names=tuple(columns),
        pearson_r=tuple(pearson_r_vals),
        spearman_r=tuple(spearman_r_vals),
        skewness=tuple(skewness_vals),
        kurtosis=tuple(kurtosis_vals),
        vif=tuple(vif_vals),
        tolerance=tuple(tolerance_vals),
    )

    # ── Group 4: Full residuals ───────────────────────────────────────────────
    se_regression = float(sqrt(float(model.mse_resid)))
    predictions = np.asarray(model.fittedvalues, dtype=np.float64)
    residuals = y_train - predictions

    # LOOCV: yhat - h*e/(1-h)  [matches LAMBDA: yhat - h*e/(1-h)]
    loocv_predictions = predictions - h * e / (1.0 - h)

    # Q-Q plotting positions: rank_fraction and normal_scores
    sorted_y = np.sort(y_train)
    count_leq = np.searchsorted(sorted_y, y_train, side="right")
    count_less = np.searchsorted(sorted_y, y_train, side="left")
    rank_fraction_arr = count_leq / n
    nd = NormalDist()
    normal_scores_arr = np.array([nd.inv_cdf(float((cl + 0.5) / n)) for cl in count_less])

    scaled_residuals = residuals / se_regression
    scaled_residuals_ranked = np.sort(scaled_residuals)

    # Studentized residuals: e / (se * sqrt(1-h))  [matches LAMBDA exactly]
    studentized_residuals = e / (se_regression * np.sqrt(1.0 - h))
    studentized_residuals_ranked = np.sort(studentized_residuals)

    # Cook's distance: r² * h / ((1-h) * p)  [matches LAMBDA exactly]
    p_design = k + (1 if include_intercept else 0)
    cooks_distance = studentized_residuals**2 * h / ((1.0 - h) * p_design)

    full_residuals = RegressionFullResiduals(
        predictions=tuple(float(v) for v in predictions),
        residuals=tuple(float(v) for v in residuals),
        loocv_predictions=tuple(float(v) for v in loocv_predictions),
        rank_fraction=tuple(float(v) for v in rank_fraction_arr),
        y_ranked=tuple(float(v) for v in np.sort(y_train)),
        normal_scores=tuple(float(v) for v in normal_scores_arr),
        scaled_residuals=tuple(float(v) for v in scaled_residuals),
        scaled_residuals_ranked=tuple(float(v) for v in scaled_residuals_ranked),
        hat_diagonal=tuple(float(v) for v in h),
        studentized_residuals=tuple(float(v) for v in studentized_residuals),
        studentized_residuals_ranked=tuple(float(v) for v in studentized_residuals_ranked),
        cooks_distance=tuple(float(v) for v in cooks_distance),
    )

    # ── Group 5: Prediction interval at training-data column means ────────────
    x_means = np.mean(x_features, axis=0)  # shape (k,)
    if include_intercept:
        x_new_design = np.concatenate([[1.0], x_means])
    else:
        x_new_design = x_means

    h_new = float(x_new_design @ xtx_inv @ x_new_design)
    point_estimate = float(x_new_design @ model.params)
    se_pred = se_regression * sqrt(1.0 + h_new)
    t_crit = float(_scipy_stats.t.ppf(1.0 - alpha / 2.0, summary.df_residual))
    margin = t_crit * se_pred
    prediction_interval = RegressionPredictionInterval(
        pred_input_values=tuple(float(v) for v in x_means),
        point_estimate=point_estimate,
        se_prediction=se_pred,
        t_critical=t_crit,
        lower=point_estimate - margin,
        upper=point_estimate + margin,
        confidence_level=1.0 - alpha,
    )

    return RegressionSheetResults(
        summary=summary,
        vectors=vectors,
        predictor_summary=predictor_summary,
        full_residuals=full_residuals,
        prediction_interval=prediction_interval,
    )


def build_regression_sheet_qc_configs(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[tuple[str, bool, RegressionSheetResults]]:
    """Compute RegressionSheetResults for all REGRESSION_QC_CONFIGS entries."""
    results = []
    for name, allow_intercept, predictors in REGRESSION_QC_CONFIGS:
        result = calculate_regression_sheet_results(
            input_csv_path=csv_path,
            include_intercept=allow_intercept,
            feature_columns=predictors,
        )
        results.append((name, allow_intercept, result))
    return results

"""Compute Python expected values for every output zone of the Regression worksheet."""
from __future__ import annotations

from math import sqrt
from pathlib import Path
from statistics import NormalDist

import numpy as np
import statsmodels.api as sm  # type: ignore[import-untyped]
from scipy import stats as _scipy_stats  # type: ignore[import-untyped]

from .analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
    _build_training_arrays,
    _fit_ols_model,
    _load_normalized_rows,
    _validate_required_headers,
)
from .regression_shared import (
    FEATURE_COLUMNS,
    RegressionFullResiduals,
    RegressionPredictionInterval,
    RegressionPredictorSummary,
    RegressionSheetResults,
    RegressionSummary,
    RegressionVectors,
)


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
    ("sparse_intercept", True, _SPARSE_PREDICTORS),
    ("sparse_no_intercept", False, _SPARSE_PREDICTORS),
    ("medium_intercept", True, _MEDIUM_PREDICTORS),
    ("medium_no_intercept", False, _MEDIUM_PREDICTORS),
    ("full_intercept", True, FEATURE_COLUMNS),
    ("full_no_intercept", False, FEATURE_COLUMNS),
]


def calculate_regression_results_from_matrix(
    x_features: np.ndarray,
    y_train: np.ndarray,
    predictor_names: tuple[str, ...],
    include_intercept: bool = True,
    alpha: float = 0.05,
    sequence_values: np.ndarray | None = None,
    allow_singular: bool = False,
) -> RegressionSheetResults:
    """Fit OLS and compute expected values for the current Regression sheet.

    ``x_features`` is the constructed design matrix without an intercept
    column. It may include continuous predictors, dummy-coded categorical
    columns, or any future spec-derived numeric design columns.
    """
    x_features = np.asarray(x_features, dtype=np.float64)
    y_train = np.asarray(y_train, dtype=np.float64)
    if x_features.ndim != 2:
        raise ValueError("x_features must be a 2-D matrix")
    if x_features.shape[0] != len(y_train):
        raise ValueError("x_features and y_train row counts differ")
    if x_features.shape[1] != len(predictor_names):
        raise ValueError("predictor_names length must match x_features columns")

    n = len(y_train)
    k = x_features.shape[1]
    x_train = (
        np.column_stack([np.ones(n), x_features])
        if include_intercept
        else x_features
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)

    df_regression = k
    df_total = n - 1 if include_intercept else n
    df_residual = int(model.df_resid)
    r_squared = float(model.rsquared)
    ss_total = float(model.centered_tss if include_intercept else model.uncentered_tss)
    ss_residual = float(model.ssr)
    ss_regression = ss_total - ss_residual
    se_regression = float(np.sqrt(model.mse_resid))

    xtx = x_train.T @ x_train
    xtx_inv = np.linalg.pinv(xtx) if allow_singular else np.linalg.inv(xtx)
    z = x_train @ xtx_inv
    h = np.sum(z * x_train, axis=1)
    e = np.asarray(model.resid, dtype=np.float64)
    predictions = np.asarray(model.fittedvalues, dtype=np.float64)
    p_design = k + (1 if include_intercept else 0)

    press = float(np.sum((e / (1.0 - h)) ** 2))
    if sequence_values is None:
        dw_resid = e
    else:
        order = np.argsort(np.asarray(sequence_values, dtype=np.float64), kind="stable")
        dw_resid = e[order]
    durbin_watson = float(np.sum(np.diff(dw_resid) ** 2) / np.sum(dw_resid**2))

    f_stat = (
        float((ss_regression / df_regression) / (ss_residual / df_residual))
        if df_regression > 0 and df_residual > 0 and ss_residual != 0.0
        else float("nan")
    )
    log_term = float(n * np.log(ss_residual / n))
    aic = log_term + 2.0 * p_design
    bic = log_term + p_design * float(np.log(n))
    aicc = (
        aic + 2.0 * p_design * (p_design + 1) / (n - p_design - 1)
        if n > p_design + 1
        else float("nan")
    )
    scaled_residuals_sorted = np.sort(e / se_regression)
    qq_scores = _scipy_stats.norm.ppf((np.arange(1, n + 1) - 0.5) / n)
    qq_correlation, _ = _scipy_stats.pearsonr(scaled_residuals_sorted, qq_scores)

    summary = RegressionSummary(
        observations=n,
        df_regression=df_regression,
        df_total=df_total,
        r_squared=r_squared,
        df_residual=df_residual,
        multiple_r=float(np.sqrt(max(r_squared, 0.0))),
        adjusted_r2=float(model.rsquared_adj),
        ss_total=ss_total,
        ss_residual=ss_residual,
        ss_regression=ss_regression,
        se_regression=se_regression,
        press=press,
        durbin_watson=durbin_watson,
        f_stat=f_stat,
        p_value_f=float(model.f_pvalue),
        aic=aic,
        bic=bic,
        aicc=aicc,
        qq_correlation=float(qq_correlation),
    )

    ci = np.asarray(model.conf_int(alpha=alpha), dtype=np.float64)
    pred_coefs = np.asarray(model.params[1:] if include_intercept else model.params)
    vectors = RegressionVectors(
        term_names=(
            ("Intercept", *predictor_names)
            if include_intercept
            else predictor_names
        ),
        coefficients=tuple(float(v) for v in model.params),
        std_errors=tuple(float(v) for v in model.bse),
        t_stats=tuple(float(v) for v in model.tvalues),
        p_values=tuple(float(v) for v in model.pvalues),
        ci_lower=tuple(float(v) for v in ci[:, 0]),
        ci_upper=tuple(float(v) for v in ci[:, 1]),
        beta_weights=tuple(
            float(coef * np.std(x_features[:, j], ddof=1) / np.std(y_train, ddof=1))
            for j, coef in enumerate(pred_coefs)
        ),
    )

    pearson_r_vals: list[float] = []
    spearman_r_vals: list[float] = []
    skewness_vals: list[float] = []
    kurtosis_vals: list[float] = []
    vif_vals: list[float] = []
    for j in range(k):
        x_col = x_features[:, j]
        pearson_r_vals.append(float(np.corrcoef(x_col, y_train)[0, 1]))
        spearman_r_vals.append(float(_scipy_stats.spearmanr(x_col, y_train).statistic))
        skewness_vals.append(float(_scipy_stats.skew(x_col, bias=False)))
        kurtosis_vals.append(float(_scipy_stats.kurtosis(x_col, fisher=True, bias=False)))
        if k == 1:
            vif_vals.append(1.0)
        else:
            others = np.delete(x_features, j, axis=1)
            others_with_const = np.column_stack([np.ones(n), others])
            r2_j = float(sm.OLS(x_col, others_with_const).fit().rsquared)
            denom = 1.0 - r2_j
            vif_vals.append(float(np.inf) if abs(denom) <= 1e-12 else 1.0 / denom)

    predictor_summary = RegressionPredictorSummary(
        predictor_names=predictor_names,
        pearson_r=tuple(pearson_r_vals),
        spearman_r=tuple(spearman_r_vals),
        skewness=tuple(skewness_vals),
        kurtosis=tuple(kurtosis_vals),
        vif=tuple(vif_vals),
        tolerance=tuple(0.0 if np.isinf(v) else 1.0 / v for v in vif_vals),
    )

    loocv_predictions = predictions - h * e / (1.0 - h)
    sorted_y = np.sort(y_train)
    count_less = np.searchsorted(sorted_y, y_train, side="left")
    nd = NormalDist()
    normal_scores = np.array([nd.inv_cdf(float((cl + 0.5) / n)) for cl in count_less])
    studentized = e / (se_regression * np.sqrt(1.0 - h))
    cooks_distance = studentized**2 * h / ((1.0 - h) * p_design)
    scale_location = np.sqrt(np.abs(studentized))
    full_residuals = RegressionFullResiduals(
        dependent_var=tuple(float(v) for v in y_train),
        predictions=tuple(float(v) for v in predictions),
        residuals=tuple(float(v) for v in e),
        loocv_residuals=tuple(float(v) for v in y_train - loocv_predictions),
        hat_diagonal=tuple(float(v) for v in h),
        studentized_residuals=tuple(float(v) for v in studentized),
        cooks_distance=tuple(float(v) for v in cooks_distance),
        normal_scores_ranked=tuple(float(v) for v in np.sort(normal_scores)),
        studentized_residuals_ranked=tuple(float(v) for v in np.sort(studentized)),
        scale_location=tuple(float(v) for v in scale_location),
    )

    x_means = np.mean(x_features, axis=0)
    x_new_design = np.concatenate([[1.0], x_means]) if include_intercept else x_means
    h_new = float(x_new_design @ xtx_inv @ x_new_design)
    point_estimate = float(x_new_design @ model.params)
    se_pred = se_regression * sqrt(1.0 + h_new)
    t_crit = float(_scipy_stats.t.ppf(1.0 - alpha / 2.0, df_residual))
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


def calculate_regression_sheet_results(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    include_intercept: bool = True,
    feature_columns: list[str] | None = None,
    alpha: float = 0.05,
) -> RegressionSheetResults:
    """Fit OLS and return expected values for a continuous-predictor config."""
    columns = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    input_path = input_csv_path.resolve()

    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers, columns)
    x_train, y_train, _, _ = _build_training_arrays(
        normalized_rows, include_intercept, columns, filter_columns=columns
    )
    x_features = x_train[:, 1:] if include_intercept else x_train
    return calculate_regression_results_from_matrix(
        x_features=x_features,
        y_train=y_train,
        predictor_names=tuple(columns),
        include_intercept=include_intercept,
        alpha=alpha,
    )


def build_regression_sheet_qc_configs(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[tuple[str, bool, RegressionSheetResults]]:
    """Compute RegressionSheetResults for the legacy continuous QC configs."""
    results = []
    for name, allow_intercept, predictors in REGRESSION_QC_CONFIGS:
        result = calculate_regression_sheet_results(
            input_csv_path=csv_path,
            include_intercept=allow_intercept,
            feature_columns=predictors,
        )
        results.append((name, allow_intercept, result))
    return results

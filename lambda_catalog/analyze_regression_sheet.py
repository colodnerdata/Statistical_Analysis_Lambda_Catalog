"""Compute Python expected values for every output zone of the Regression worksheet."""
from __future__ import annotations

from math import sqrt
from pathlib import Path
from statistics import NormalDist

import numpy as np
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


def _predictor_groups(predictor_names: tuple[str, ...]) -> list[str]:
    """Group key per constructed column: the text before the first ``": "``.

    Mirrors the Constructed_Column_Names() convention (``"Header: level"``
    for a categorical predictor's dummy columns) and the GVIF LAMBDA's own
    grouping logic in Excel, so dummy columns from the same source variable
    share a group key.
    """
    return [name.split(": ", 1)[0] for name in predictor_names]


def _generalized_vif(x_features: np.ndarray, groups: list[str]) -> list[float]:
    """Fox & Monette (1992) Generalized VIF, one value per column of ``x_features``.

    Independent numpy mirror of the GVIF LAMBDA: for each group of columns
    (dummy columns from the same categorical predictor share a group),
    GVIF = det(R11) * det(R22) / det(R), where R is the correlation matrix of
    ``x_features``, R11 is restricted to the group's own columns, and R22 is
    restricted to every other column. Reduces exactly to ordinary VIF (and to
    1 when there is only one group total) when a column stands alone.

    det_full appears in every column's denominator, so exact multicollinearity
    anywhere in x_features (det_full == 0) yields inf/nan for every column, not
    only the collinear ones — mirroring the #DIV/0! the GVIF LAMBDA produces in
    Excel, and the same all-or-nothing failure mode ordinary VIF's own
    1/(1-R^2) already has at R^2 = 1. numpy's float division on the det_full
    scalar returns inf/nan with a RuntimeWarning rather than raising.
    """
    k = x_features.shape[1]
    distinct = sorted(set(groups))
    if len(distinct) <= 1:
        return [1.0] * k

    corr = np.corrcoef(x_features, rowvar=False)
    det_full = float(np.linalg.det(corr))

    result = [0.0] * k
    for group in distinct:
        own_idx = [i for i, g in enumerate(groups) if g == group]
        other_idx = [i for i in range(k) if i not in own_idx]
        r11 = corr[np.ix_(own_idx, own_idx)]
        r22 = corr[np.ix_(other_idx, other_idx)]
        value = float(np.linalg.det(r11) * np.linalg.det(r22) / det_full)
        for i in own_idx:
            result[i] = value
    return result


def _demean_within_groups(values: np.ndarray, groups: np.ndarray) -> np.ndarray:
    """One-way within-group demeaning — the Fixed Effects fit-time transform.

    Mirrors the within estimator validated independently in
    ``tests/test_within_estimator.py``/``tests/test_df_absorbed_threading.py``
    (Demean_By/``y_s()``/``X_s_Within()`` on the Excel side): subtract each
    group's own mean from every value in that group. Works for both a 1-D
    response and a 2-D design matrix (demeaned column-by-column via the
    ``axis=0`` mean).
    """
    values = np.asarray(values, dtype=np.float64)
    demeaned = values.copy()
    for group in np.unique(groups):
        mask = groups == group
        demeaned[mask] = values[mask] - values[mask].mean(axis=0)
    return demeaned


def calculate_regression_results_from_matrix(
    x_features: np.ndarray,
    y_train: np.ndarray,
    predictor_names: tuple[str, ...],
    include_intercept: bool = True,
    alpha: float = 0.05,
    sequence_values: np.ndarray | None = None,
    group_labels: np.ndarray | None = None,
) -> RegressionSheetResults:
    """Fit OLS and compute expected values for the current Regression sheet.

    ``x_features`` is the constructed design matrix without an intercept
    column. It may include continuous predictors, dummy-coded categorical
    columns, or any future spec-derived numeric design columns.

    ``group_labels``, when given, is the Fixed Effects grouping column
    (one label per row, aligned with ``x_features``/``y_train``). Every
    workbook function that fits on ``X_s_Within()``/``y_s()`` — Coefficients,
    Predictions, Residuals, Hat_Diagonal, SE_Regression and everything built
    on it (Adjusted R², SE/t/p/CI, AIC/BIC/AICc, F, Studentized Residuals,
    Cook's Distance, Scale-Location, QQ Correlation, PRESS, Prediction
    Interval), and Beta_Weights — fits on the one-way within-demeaned pair
    instead of the raw one, with the G-1 absorbed degrees of freedom (G =
    distinct group count) subtracted from every df-dependent statistic
    (Absorbed_Degrees_Of_Freedom()/DF_Absorbed threading, see
    ``tests/test_df_absorbed_threading.py``). Predictor Summary
    (Pearson/Spearman R, Skewness, Kurtosis, GVIF, Tolerance) stays on the
    RAW ``x_features``/``y_train`` in both cases — the sheet computes those
    off ``X_s()``/``Response_Column()`` regardless of Fixed Effects (see the
    comment above the Predictor Summary formulas in write_sheet_regression.py).
    Durbin-Watson has no valid FE-active reading (the sheet's AB11 cell shows
    "n/a — FE active" instead of a number whenever a Fixed Effects row is
    declared — see the DW/BFN trigger matrix in
    ``tests/test_bfn_panel_durbin_watson_verification.py``), so
    ``durbin_watson`` is NaN whenever ``group_labels`` is given; the
    ``compare_values`` QC comparison already treats NaN/None on both sides as
    "both missing", not a mismatch.
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

    if group_labels is not None:
        group_labels = np.asarray(group_labels)
        if len(group_labels) != n:
            raise ValueError("group_labels length must match x_features/y_train rows")
        x_fit = _demean_within_groups(x_features, group_labels)
        y_fit = _demean_within_groups(y_train, group_labels)
        df_absorbed = len(np.unique(group_labels)) - 1
    else:
        x_fit = x_features
        y_fit = y_train
        df_absorbed = 0

    x_train = (
        np.column_stack([np.ones(n), x_fit])
        if include_intercept
        else x_fit
    )
    model = _fit_ols_model(x_train, y_fit, include_intercept)

    df_regression = k
    df_total = n - 1 if include_intercept else n
    naive_df_residual = int(model.df_resid)
    df_residual = naive_df_residual - df_absorbed
    if df_residual <= 0:
        raise ValueError(
            "Fixed Effects absorbed too many degrees of freedom: "
            f"naive_df_residual={naive_df_residual}, df_absorbed={df_absorbed} "
            f"leaves df_residual={df_residual} <= 0 (n={n}, k={k}). "
            "Too many FE groups relative to the number of observations — "
            "reduce the group count or increase the sample size."
        )
    # SQRT(naive_df/true_df): the exact SE_Coefficients/SE_Regression
    # rescaling validated in test_df_absorbed_threading.py — LINEST's own SE
    # (computed at naive_df) times this factor reproduces an explicit LSDV
    # fit's SE without ever materializing the G-1 group dummies.
    df_absorbed_scale = (
        float(np.sqrt(naive_df_residual / df_residual)) if df_absorbed else 1.0
    )
    r_squared = float(model.rsquared)
    ss_total = float(model.centered_tss if include_intercept else model.uncentered_tss)
    ss_residual = float(model.ssr)
    ss_regression = ss_total - ss_residual
    se_regression = float(np.sqrt(ss_residual / df_residual))

    xtx = x_train.T @ x_train
    xtx_inv = np.linalg.inv(xtx)
    z = x_train @ xtx_inv
    h = np.sum(z * x_train, axis=1)
    e = np.asarray(model.resid, dtype=np.float64)
    predictions = np.asarray(model.fittedvalues, dtype=np.float64)
    p_design = k + (1 if include_intercept else 0)  # unaffected by DF_Absorbed
    # AIC/BIC/AICc's own parameter count DOES add the absorbed group
    # intercepts back in (an LSDV fit's own p is the ground truth) — a
    # separate count from p_design above (Cook's Distance/Hat_Diagonal never
    # see the absorbed adjustment).
    p_information_criteria = p_design + df_absorbed

    press = float(np.sum((e / (1.0 - h)) ** 2))
    if group_labels is not None:
        # No valid reading: AB11 shows "n/a — FE active" on the sheet
        # whenever a Fixed Effects row is declared, regardless of Sequence
        # state (BFN_Panel_Durbin_Watson takes over instead, which this QC
        # oracle does not model — see the module docstring above).
        durbin_watson = float("nan")
    elif sequence_values is None:
        durbin_watson = float(np.sum(np.diff(e) ** 2) / np.sum(e**2))
    else:
        order = np.argsort(np.asarray(sequence_values, dtype=np.float64), kind="stable")
        dw_resid = e[order]
        durbin_watson = float(np.sum(np.diff(dw_resid) ** 2) / np.sum(dw_resid**2))

    f_stat = (
        float((ss_regression / df_regression) / (ss_residual / df_residual))
        if df_regression > 0 and df_residual > 0 and ss_residual != 0.0
        else float("nan")
    )
    p_value_f = (
        float(_scipy_stats.f.sf(f_stat, df_regression, df_residual))
        if np.isfinite(f_stat)
        else float("nan")
    )
    log_term = float(n * np.log(ss_residual / n))
    aic = log_term + 2.0 * p_information_criteria
    bic = log_term + p_information_criteria * float(np.log(n))
    aicc = (
        aic
        + 2.0
        * p_information_criteria
        * (p_information_criteria + 1)
        / (n - p_information_criteria - 1)
        if n > p_information_criteria + 1
        else float("nan")
    )
    scaled_residuals_sorted = np.sort(e / se_regression)
    qq_scores = _scipy_stats.norm.ppf((np.arange(1, n + 1) - 0.5) / n)
    qq_correlation, _ = _scipy_stats.pearsonr(scaled_residuals_sorted, qq_scores)

    adjusted_r2 = 1.0 - (1.0 - r_squared) * df_total / df_residual

    summary = RegressionSummary(
        observations=n,
        df_regression=df_regression,
        df_total=df_total,
        r_squared=r_squared,
        df_residual=df_residual,
        multiple_r=float(np.sqrt(max(r_squared, 0.0))),
        adjusted_r2=float(adjusted_r2),
        ss_total=ss_total,
        ss_residual=ss_residual,
        ss_regression=ss_regression,
        se_regression=se_regression,
        press=press,
        durbin_watson=durbin_watson,
        f_stat=f_stat,
        p_value_f=p_value_f,
        aic=aic,
        bic=bic,
        aicc=aicc,
        qq_correlation=float(qq_correlation),
    )

    # SE/t/p/CI all rescale LINEST's own naive-df numbers by df_absorbed_scale
    # (or true_df in place of naive_df where a df appears directly), matching
    # SE_Coefficients/T_Statistics/P_Values/Confidence_Interval_* exactly —
    # see test_df_absorbed_threading.py's se_coefficients_mirror.
    std_errors = np.asarray(model.bse, dtype=np.float64) * df_absorbed_scale
    t_stats = np.asarray(model.params, dtype=np.float64) / std_errors
    p_values = 2.0 * _scipy_stats.t.sf(np.abs(t_stats), df_residual)
    t_crit_coef = float(_scipy_stats.t.ppf(1.0 - alpha / 2.0, df_residual))
    ci_lower = np.asarray(model.params, dtype=np.float64) - t_crit_coef * std_errors
    ci_upper = np.asarray(model.params, dtype=np.float64) + t_crit_coef * std_errors
    pred_coefs = np.asarray(model.params[1:] if include_intercept else model.params)
    vectors = RegressionVectors(
        term_names=(
            ("Intercept", *predictor_names)
            if include_intercept
            else predictor_names
        ),
        coefficients=tuple(float(v) for v in model.params),
        std_errors=tuple(float(v) for v in std_errors),
        t_stats=tuple(float(v) for v in t_stats),
        p_values=tuple(float(v) for v in p_values),
        ci_lower=tuple(float(v) for v in ci_lower),
        ci_upper=tuple(float(v) for v in ci_upper),
        beta_weights=tuple(
            # Beta_Weights(X_s_Within(),y_s(),...) standardizes by the
            # WITHIN-demeaned x/y, not the raw predictor-summary columns.
            float(coef * np.std(x_fit[:, j], ddof=1) / np.std(y_fit, ddof=1))
            for j, coef in enumerate(pred_coefs)
        ),
    )

    pearson_r_vals: list[float] = []
    spearman_r_vals: list[float] = []
    skewness_vals: list[float] = []
    kurtosis_vals: list[float] = []
    for j in range(k):
        x_col = x_features[:, j]
        pearson_r_vals.append(float(np.corrcoef(x_col, y_train)[0, 1]))
        spearman_r_vals.append(float(_scipy_stats.spearmanr(x_col, y_train).statistic))
        skewness_vals.append(float(_scipy_stats.skew(x_col, bias=False)))
        kurtosis_vals.append(float(_scipy_stats.kurtosis(x_col, fisher=True, bias=False)))

    gvif_vals = _generalized_vif(x_features, _predictor_groups(predictor_names))

    predictor_summary = RegressionPredictorSummary(
        predictor_names=predictor_names,
        pearson_r=tuple(pearson_r_vals),
        spearman_r=tuple(spearman_r_vals),
        skewness=tuple(skewness_vals),
        kurtosis=tuple(kurtosis_vals),
        gvif=tuple(gvif_vals),
        tolerance=tuple(1.0 / v for v in gvif_vals),
    )

    # Y (dependent_var), the Normal_Scores basis, and LOOCV_Residual all take
    # y_s() on the sheet — the within-demeaned response under FE, same as
    # every other fit-stage quantity above (see the docstring: the whole
    # Residual Output table has to read as one internally consistent block).
    loocv_predictions = predictions - h * e / (1.0 - h)
    sorted_y = np.sort(y_fit)
    count_less = np.searchsorted(sorted_y, y_fit, side="left")
    nd = NormalDist()
    normal_scores = np.array([nd.inv_cdf(float((cl + 0.5) / n)) for cl in count_less])
    studentized = e / (se_regression * np.sqrt(1.0 - h))
    cooks_distance = studentized**2 * h / ((1.0 - h) * p_design)
    scale_location = np.sqrt(np.abs(studentized))
    full_residuals = RegressionFullResiduals(
        dependent_var=tuple(float(v) for v in y_fit),
        predictions=tuple(float(v) for v in predictions),
        residuals=tuple(float(v) for v in e),
        loocv_residuals=tuple(float(v) for v in y_fit - loocv_predictions),
        hat_diagonal=tuple(float(v) for v in h),
        studentized_residuals=tuple(float(v) for v in studentized),
        cooks_distance=tuple(float(v) for v in cooks_distance),
        normal_scores_ranked=tuple(float(v) for v in np.sort(normal_scores)),
        studentized_residuals_ranked=tuple(float(v) for v in np.sort(studentized)),
        scale_location=tuple(float(v) for v in scale_location),
    )

    # NOTE: this box models the pre-v2.1 single mean-response CI, not the
    # shipped sheet's Group_Prediction_Interval/group-mean-recovery mechanism
    # (see the STALE comment in tools/inspect_regression_sheet.py) — accurate
    # for every existing no-FE case, but not a faithful FE-active oracle.
    # Callers building an FE spec case should exclude this section from
    # comparison rather than trust these numbers against the sheet.
    x_means = np.mean(x_fit, axis=0)
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

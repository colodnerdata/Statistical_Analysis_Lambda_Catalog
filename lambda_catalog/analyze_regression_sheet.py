"""Compute Python expected values for every output zone of the Regression worksheet."""
from __future__ import annotations

import math
from math import sqrt
from pathlib import Path
from statistics import NormalDist

import numpy as np
from scipy import stats as _scipy_stats  # type: ignore[import-untyped]

from .analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
    TARGET_COLUMN,
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
    RegressionUnitSpace,
    RegressionVectors,
)


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
    (Demean_By/``Design_Response()``/``Design_Columns()`` on the Excel side): subtract each
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


def _build_model_formula(
    response_display: str,
    predictor_names: tuple[str, ...],
    include_intercept: bool,
    fixed_effects_name: str | None,
) -> str:
    """Assemble the Model Formula readout's text from the spec's display names.

    A CHARACTER-EXACT mirror of the sheet-scoped ``Model_Formula()`` closure
    (``lambda_functions.json``, scope ``"Regression"``), which the readout
    cell in the §4b materialization band calls::

        =<response>&" ~ "&IF(Allow_Intercept,"1 + ","0 + ")
         &IFERROR(TEXTJOIN(" + ",TRUE,Constructed_Column_Names()),"")
         &IF(<fe count>>0," | "&<fe name>,"")

    Three consequences of mirroring it literally rather than assembling
    something equivalent-looking, each a way a hand-built equivalent can
    diverge from the cell (and the reason the assertion string-compares
    against the live cell rather than an equivalent):

    * The intercept term is ``"0 + "`` when the intercept is OFF, not
      omitted — the sheet always writes one or the other.
    * The intercept marker is a PREFIX, not the first element of the join,
      so an empty predictor list renders ``"MPG ~ 1 + "`` with a trailing
      separator. That is genuinely what the cell shows.
    * ``TEXTJOIN(..., TRUE, ...)`` skips empty strings, so a degenerate
      column contributing no name leaves no double separator behind.

    Built from the constructed-column names (``Constructed_Column_Names()``
    on the sheet side) — which already emit ``Ln(name)`` per logged
    predictor, level-qualified dummy names, and ``left × right`` interaction
    names — so the mixed Log/None predictor case renders correctly with no
    extra work. ``fixed_effects_name`` is the declared FE variable's own
    name from the spec (e.g. ``Facility``), identical to the spec feedback
    block's FE Variable cell — never a group level value.
    """
    intercept_term = "1 + " if include_intercept else "0 + "
    joined = " + ".join(name for name in predictor_names if name)
    suffix = f" | {fixed_effects_name}" if fixed_effects_name else ""
    return f"{response_display} ~ {intercept_term}{joined}{suffix}"


def _bfn_panel_durbin_watson(
    residuals: np.ndarray,
    group_labels: np.ndarray | None,
    sequence_values: np.ndarray | None,
    base_period_delta: float | None,
) -> float:
    """Mirror the ``BFN_Panel_Durbin_Watson`` LAMBDA, gating included.

        BFN = Σᵢ Σₜ (û(i,t) − û(i,t−Δ))² ÷ Σᵢ Σₜ û(i,t)²

    Returns NaN for every state in which the sheet's ``AE12`` shows text or
    ``#N/A`` rather than a number, so the oracle never offers a value for a
    cell that is not one. Those states, in the LAMBDA's own order:

    * no Sequence axis, or no Fixed Effects variable — the cell's own
      ``IF`` chain short-circuits to an ``n/a`` string;
    * ``Δ`` unresolved. ``Base_Period_Delta()`` reads the **typed**
      ``Sequence Period`` and returns ``#N/A`` when none is typed — the
      accessor is the override, never a silent 1 (see DECISIONS § *Sequence
      Period / Period In Use split*). With ``step`` at ``#N/A`` every
      ``Difference_By`` lookup misses;
    * zero computable differences — every group a singleton, say. The
      LAMBDA's ``n_terms`` guard returns ``#N/A`` here rather than letting
      ``IFERROR(d,0)`` mask an all-error column into an all-zero numerator,
      which would display BFN = 0: a fake strong-negative-autocorrelation
      reading. The same guard is why this returns NaN and not 0.0.

    The differencing is an exact ``(group, seq − Δ)`` match, not row
    arithmetic, so the statistic is invariant to physical row order and a
    panel gap contributes no fabricated term. The DENOMINATOR sums every
    residual's square — first periods and gap rows still count there, as
    the BFN definition requires.
    """
    if group_labels is None or sequence_values is None or base_period_delta is None:
        return float("nan")
    if not np.isfinite(base_period_delta):
        return float("nan")

    sequence = np.asarray(sequence_values, dtype=np.float64)
    # Exact-match lookup keyed on (group, seq), mirroring Difference_By's
    # XLOOKUP. First writer wins on a duplicate key, matching XLOOKUP's
    # default first-match search.
    position: dict[tuple[object, float], int] = {}
    for index, (group, seq) in enumerate(zip(group_labels, sequence)):
        position.setdefault((group, float(seq)), index)

    numerator = 0.0
    terms = 0
    for index, (group, seq) in enumerate(zip(group_labels, sequence)):
        prior = position.get((group, float(seq) - float(base_period_delta)))
        if prior is None:
            continue
        numerator += float(residuals[index] - residuals[prior]) ** 2
        terms += 1

    if terms == 0:
        return float("nan")
    return numerator / float(np.sum(np.asarray(residuals, dtype=np.float64) ** 2))


def calculate_regression_results_from_matrix(
    x_features: np.ndarray,
    y_train: np.ndarray,
    predictor_names: tuple[str, ...],
    include_intercept: bool = True,
    alpha: float = 0.05,
    sequence_values: np.ndarray | None = None,
    group_labels: np.ndarray | None = None,
    selected_group: str | None = None,
    response_transform: str = "None",
    predictor_transform: str = "None",
    response_name: str = "Response",
    fixed_effects_name: str | None = None,
    back_transform: str = "Duan",
    base_period_delta: float | None = None,
) -> RegressionSheetResults:
    """Fit OLS and compute expected values for the current Regression sheet.

    ``x_features`` is the constructed design matrix without an intercept
    column. It may include continuous predictors, dummy-coded categorical
    columns, or any future spec-derived numeric design columns.

    ``group_labels``, when given, is the Fixed Effects grouping column
    (one label per row, aligned with ``x_features``/``y_train``). Every
    workbook function that fits on ``Design_Columns()``/``Design_Response()`` — Coefficients,
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
    off ``Predictor_Columns()``/``Response_Column()`` regardless of Fixed Effects (see the
    comment above the Predictor Summary formulas in write_sheet_regression.py).
    Durbin-Watson has no valid FE-active reading (the sheet's AB11 cell shows
    "n/a — FE active" instead of a number whenever a Fixed Effects row is
    declared — see the DW/BFN trigger matrix in
    ``tests/test_bfn_panel_durbin_watson_verification.py``), so
    ``durbin_watson`` is NaN whenever ``group_labels`` is given; the
    ``compare_values`` QC comparison already treats NaN/None on both sides as
    "both missing", not a mismatch.

    ``bfn_panel_durbin_watson`` is the other half of that pair — the cell
    that DOES hold a number in the FE-active state — and is gated the
    opposite way, so at most one of the two is ever live. It needs
    ``base_period_delta``, the TYPED Sequence Period: ``Base_Period_Delta()``
    is the override accessor and returns ``#N/A`` when nothing is typed, so
    a Fixed Effects model with no declared period has no computable panel
    statistic and this stays NaN too. Both cells reading as text is a
    legitimate state, and saying so beats inventing a Δ of 1.

    ``back_transform`` mirrors the sheet's Back-Transform Method input
    (``$AH$4``, "Duan" or "Naive"), which the unit-space block's
    ``Unit_Space_R_Squared`` / ``Unit_Space_Adjusted_R_Squared`` /
    ``Unit_Space_RMSE`` calls all take as an argument, and which the AL
    prediction column and the AZ/BA original-units residual columns
    dispatch on. Under a Log response, "Duan" multiplies ``EXP(ŷ)`` by the
    smearing factor and "Naive" does not — so every statistic derived from
    the unit-space residuals differs between the two, which is exactly why
    both need an oracle. Two things do NOT dispatch on it: the CI/PI bounds
    (quantiles, back-transformed with ``EXP`` only under both settings —
    see the AL7:AL10 caveat row on the sheet), and the observed column
    (an observation is not a prediction and never carries the smearing
    factor). Under ``response_transform="None"`` the two methods coincide
    exactly, which keeps the reduction invariant holding either way.

    ``selected_group`` picks which group's mean/count the Prediction
    Interval box (AK3:AK14 on the sheet) is anchored to via
    ``Group_Prediction_Interval``'s group-mean-recovery form — the sheet's
    own ``$AK$12`` cell, which defaults to the alphabetically-first observed
    group when nothing is typed. ``None`` (the default here) mirrors that:
    picks the alphabetically-first group. With no ``group_labels``, every
    row is treated as one constant ``"(all)"`` group, which makes this
    collapse exactly to the pre-v2.1 single-interval numbers (see
    ``tests/test_group_prediction_interval.py``).
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
    # The two serial-correlation cells, AE11 (plain DW) and AE12 (BFN), are
    # mutually gated on the sheet and exactly one of them is ever a number.
    # The oracle mirrors that gating rather than always computing both: a
    # value here for a cell showing "n/a — FE active" would be comparing
    # against text.
    if group_labels is not None:
        # No valid reading: AE11 shows "n/a — FE active" on the sheet
        # whenever a Fixed Effects row is declared, regardless of Sequence
        # state. BFN_Panel_Durbin_Watson takes over — computed just below.
        durbin_watson = float("nan")
    elif sequence_values is None:
        # No Sequence axis declared, so AE11 reads "n/a — requires Sequence"
        # — the cell's FIRST gate, before the FE one above.
        #
        # The oracle returns NaN here, matching the sheet: differencing
        # residuals in row order is not a weaker reading of the statistic,
        # it is a different one. DW is only meaningful along a declared
        # ordering, which is exactly why the sheet refuses to show one
        # without a Sequence axis, so the oracle refuses too.
        durbin_watson = float("nan")
    else:
        order = np.argsort(np.asarray(sequence_values, dtype=np.float64), kind="stable")
        dw_resid = e[order]
        durbin_watson = float(np.sum(np.diff(dw_resid) ** 2) / np.sum(dw_resid**2))

    bfn_panel_durbin_watson = _bfn_panel_durbin_watson(
        e, group_labels, sequence_values, base_period_delta
    )

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
        bfn_panel_durbin_watson=bfn_panel_durbin_watson,
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
            # Beta_Weights(Design_Columns(),Design_Response(),...) standardizes by the
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
    # Design_Response() on the sheet — the within-demeaned response under FE, same as
    # every other fit-stage quantity above (see the docstring: the whole
    # Residual Output table has to read as one internally consistent block).
    loocv_predictions = predictions - h * e / (1.0 - h)
    # Normal_Scores ranks the within-demeaned response y_fit by its strict-less-than
    # count (Rankit: Φ⁻¹((rank − 0.5)/n)). The response is reported to few decimals, so
    # after country Fixed-Effects demeaning it is heavily tied (Life Expectancy: ~61% of
    # rows sit in exact-tie groups). Rank on values rounded to 9 decimals so tied rows
    # collapse to one rank — matching the Normal_Scores LAMBDA's ROUND(filtered, 9) —
    # instead of being split by sub-ULP float differences between NumPy and Excel, which
    # otherwise shift the Q-Q axis by one rank at hundreds of tied positions. The
    # dependent_var column below stays unrounded; only the ranking is rounded.
    y_for_rank = np.round(y_fit, 9)
    sorted_y = np.sort(y_for_rank)
    count_less = np.searchsorted(sorted_y, y_for_rank, side="left")
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

    # Group_Prediction_Interval's own group-mean recovery: y_hat = ybar_i +
    # (x_new - xbar_i)'beta, with mean-response and new-observation variance
    # terms (see tests/test_group_prediction_interval.py's docstring for the
    # derivation). This ALWAYS demeans internally — even a no-FE model
    # group-demeans by one constant "(all)" group covering every row.
    #
    # df_residual/t_crit_coef are reused as-is: Residual_Degrees_Of_Freedom is
    # a pure function of n/k/allow_arg/absorbed_arg, never of the data, so it
    # can't differ here. beta/sigma CANNOT be shortcut through pred_coefs/
    # se_regression, though — that reuse only holds with an intercept (adding
    # a constant shift to already-demeaned columns doesn't move an
    # intercept-included fit's slope coefficients or residuals by FWL), and
    # breaks for a no-intercept model, where centering by the group mean
    # genuinely changes the through-the-origin fit. So this always re-fits on
    # the group-demeaned pair, matching the formula exactly rather than
    # assuming an equivalence that only sometimes holds.
    if group_labels is not None:
        pi_group = np.asarray(group_labels)
    else:
        pi_group = np.full(n, "(all)", dtype=object)
    if selected_group is None:
        selected_group = sorted(np.unique(pi_group))[0]
    elif selected_group not in pi_group:
        raise ValueError(
            f"selected_group={selected_group!r} is not an observed group "
            f"(observed groups: {sorted(np.unique(pi_group).tolist())!r}) — "
            "group_count would be 0 and the 1/group_count variance terms "
            "below would divide by zero."
        )

    x_within_pi = _demean_within_groups(x_features, pi_group)
    y_within_pi = _demean_within_groups(y_train, pi_group)
    design_pi = (
        np.column_stack([np.ones(n), x_within_pi]) if include_intercept else x_within_pi
    )
    beta_full_pi, *_ = np.linalg.lstsq(design_pi, y_within_pi, rcond=None)
    beta_pi = beta_full_pi[1:] if include_intercept else beta_full_pi
    resid_pi = y_within_pi - design_pi @ beta_full_pi
    ssr_pi = float(resid_pi @ resid_pi)
    sigma_pi = sqrt(ssr_pi / df_residual)

    xtx_inv_pi = np.linalg.inv(x_within_pi.T @ x_within_pi)
    selected_mask = pi_group == selected_group
    group_count = int(np.sum(selected_mask))
    group_mean = float(np.mean(y_train[selected_mask]))
    xbar_i = np.mean(x_features[selected_mask], axis=0)

    # Training Mean prefill: AVERAGE of the RAW Predictor_Columns() columns — FE-independent,
    # unlike x_fit (which is demeaned once a Fixed Effects row is declared).
    x_new = np.mean(x_features, axis=0)
    deviation = x_new - xbar_i
    quad = float(deviation @ xtx_inv_pi @ deviation)
    point_estimate = group_mean + float(deviation @ beta_pi)
    se_mean = sigma_pi * sqrt(1.0 / group_count + quad)
    se_new = sigma_pi * sqrt(1.0 + 1.0 / group_count + quad)
    margin_mean = t_crit_coef * se_mean
    margin_new = t_crit_coef * se_new
    prediction_interval = RegressionPredictionInterval(
        pred_input_values=tuple(float(v) for v in x_new),
        point_estimate=point_estimate,
        se_mean=se_mean,
        se_new=se_new,
        t_critical=t_crit_coef,
        ci_lower=point_estimate - margin_mean,
        ci_upper=point_estimate + margin_mean,
        pi_lower=point_estimate - margin_new,
        pi_upper=point_estimate + margin_new,
        confidence_level=1.0 - alpha,
        group_mean=group_mean,
        group_count=group_count,
    )

    # ── v3.3 unit-space / back-transformation arithmetic ──────────────────
    # Mirrors the AG3:AH9 unit-space block, the AL Original Units prediction
    # column, and the AZ/BA Predicted Y (Original Units) / Residual
    # (Original Units) columns on the sheet.
    #
    # Two response columns are in play, and conflating them is the whole
    # hazard here:
    #
    #   y_train  the FIT-space response, un-demeaned — Response_Column() on
    #            the sheet, so ln(y) under a Log Response row. This is what
    #            the sheet passes as Unit_Space_*'s Y_Full argument.
    #   y_fit    the same column after within-demeaning — Design_Response(),
    #            what the model was actually fitted on.
    #
    # Their difference is the level the within transformation removed. It is
    # added back before exponentiating so that EXP() under Fixed Effects
    # exponentiates a predicted log response rather than a group deviation,
    # and it is GATED ON Log: nothing is exponentiated under None, and
    # shifting there would turn the within-flavoured statistics into total
    # ones and break the reduction invariant.
    level_shift = (y_train - y_fit) if response_transform == "Log" else np.zeros(n)

    # Smearing factor: 1 under None, AVERAGE(EXP(residuals)) under Log.
    # residuals are already fit-space (transformed + within-demeaned).
    if response_transform == "Log":
        smearing_factor = float(np.mean(np.exp(e)))
    elif response_transform == "None":
        smearing_factor = 1.0
    else:
        smearing_factor = float("nan")

    # Unit_Space_Predictions: Predictions(X, Y, Include) + shift, then
    # back-transformed — EXP(fitted) * smearing under Duan, EXP(fitted)
    # under Naive, fitted unchanged under None.
    fit_space_predictions = predictions + level_shift
    if response_transform == "Log":
        duan_predictions = np.exp(fit_space_predictions) * smearing_factor
        naive_predictions = np.exp(fit_space_predictions)
    elif response_transform == "None":
        duan_predictions = fit_space_predictions.copy()
        naive_predictions = fit_space_predictions.copy()
    else:
        duan_predictions = np.full(n, float("nan"))
        naive_predictions = np.full(n, float("nan"))

    # Unit_Space_Observed: the observed response read in the SAME space the
    # predictions come back in — y_fit + shift, back-transformed with the
    # Naive branch (an observation is not a prediction and never carries the
    # smearing factor). Under Log that is raw y; under None it stays the
    # within-demeaned column the ordinary statistics use, which is what
    # makes the reduction invariant hold under FE.
    y_level = y_fit + level_shift
    if response_transform == "Log":
        y_unit = np.exp(y_level)
    elif response_transform == "None":
        y_unit = y_level.copy()
    else:
        y_unit = np.full(n, float("nan"))

    # The Back-Transform Method input ($AH$4) selects between them. An
    # unrecognised method is not guessed at: the sheet's Unit_Space_*
    # functions return #N/A for anything outside the two-item validation
    # list, so the oracle refuses here rather than silently defaulting to
    # Duan and reporting agreement the workbook would not show.
    if back_transform == "Duan":
        predictions_unit = duan_predictions
    elif back_transform == "Naive":
        predictions_unit = naive_predictions
    else:
        raise ValueError(
            f"Unknown back-transform method: {back_transform!r} "
            "(expected 'Duan' or 'Naive')"
        )
    residuals_unit = y_unit - predictions_unit

    # SST_unit: centered when the model has an intercept, uncentered when
    # forced through the origin — same convention SS_Total uses.
    if include_intercept:
        sst_unit = float(np.sum((y_unit - np.mean(y_unit)) ** 2))
    else:
        sst_unit = float(np.sum(y_unit ** 2))
    sse_unit = float(np.sum(residuals_unit ** 2))
    r_squared_unit = 1.0 - sse_unit / sst_unit if sst_unit > 0 else float("nan")
    adjusted_r2_unit = (
        1.0 - (1.0 - r_squared_unit) * df_total / df_residual
        if df_residual > 0
        else float("nan")
    )
    rmse_unit = sqrt(sse_unit / df_residual) if df_residual > 0 else float("nan")

    # Original-units prediction column: same arithmetic, on the point estimate
    # from the Prediction Interval block. The point estimate is in fit space
    # (predicted log y for a Log response, predicted y for None). CI/PI bounds
    # are quantiles — back-transform with EXP only, never smeared.
    if response_transform == "Log":
        # The point estimate is a prediction, so it carries the smearing
        # factor under Duan and not under Naive. The four bounds below are
        # quantiles and stay EXP-only under both — the note on the
        # Back-Transform label at AG4 says exactly this.
        prediction_point_unit = math.exp(point_estimate) * (
            smearing_factor if back_transform == "Duan" else 1.0
        )
        prediction_ci_lower_unit = math.exp(prediction_interval.ci_lower)
        prediction_ci_upper_unit = math.exp(prediction_interval.ci_upper)
        prediction_pi_lower_unit = math.exp(prediction_interval.pi_lower)
        prediction_pi_upper_unit = math.exp(prediction_interval.pi_upper)
    elif response_transform == "None":
        prediction_point_unit = point_estimate
        prediction_ci_lower_unit = prediction_interval.ci_lower
        prediction_ci_upper_unit = prediction_interval.ci_upper
        prediction_pi_lower_unit = prediction_interval.pi_lower
        prediction_pi_upper_unit = prediction_interval.pi_upper
    else:
        nan_v = float("nan")
        prediction_point_unit = nan_v
        prediction_ci_lower_unit = nan_v
        prediction_ci_upper_unit = nan_v
        prediction_pi_lower_unit = nan_v
        prediction_pi_upper_unit = nan_v

    # Model formula: mirror the sheet's Model_Formula() readout. Built from the response
    # name, the Log/None transform flag on the response, the Allow_Intercept
    # value, the constructed column names (which already emit "Ln(name)" for
    # a logged predictor, level-qualified dummy names, and "left × right"
    # interaction names), and the Fixed Effects variable's spec name.
    response_display = response_name
    if response_transform == "Log":
        response_display = f"Ln({response_name})"
    model_formula = _build_model_formula(
        response_display=response_display,
        predictor_names=predictor_names,
        include_intercept=include_intercept,
        fixed_effects_name=fixed_effects_name,
    )

    unit_space = RegressionUnitSpace(
        smearing_factor=smearing_factor,
        r_squared_unit=r_squared_unit,
        adjusted_r2_unit=adjusted_r2_unit,
        rmse_unit=rmse_unit,
        prediction_point_unit=prediction_point_unit,
        prediction_ci_lower_unit=prediction_ci_lower_unit,
        prediction_ci_upper_unit=prediction_ci_upper_unit,
        prediction_pi_lower_unit=prediction_pi_lower_unit,
        prediction_pi_upper_unit=prediction_pi_upper_unit,
        predictions_unit=tuple(float(v) for v in predictions_unit),
        residuals_unit=tuple(float(v) for v in residuals_unit),
        model_formula=model_formula,
    )

    return RegressionSheetResults(
        summary=summary,
        vectors=vectors,
        predictor_summary=predictor_summary,
        full_residuals=full_residuals,
        prediction_interval=prediction_interval,
        unit_space=unit_space,
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
        response_name=TARGET_COLUMN,
    )

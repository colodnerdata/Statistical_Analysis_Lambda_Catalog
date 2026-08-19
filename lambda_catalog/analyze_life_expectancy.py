"""Fit OLS regression on the Life Expectancy dataset and compute QC reference values."""
from __future__ import annotations

import csv
from math import sqrt
from pathlib import Path
from statistics import NormalDist

import numpy as np
import statsmodels.api as sm  # type: ignore[import-untyped]
from scipy import stats as _scipy_stats  # type: ignore[import-untyped]
from statsmodels.regression.linear_model import (  # type: ignore[import-untyped]
    RegressionResultsWrapper,
)
from statsmodels.stats.stattools import (
    durbin_watson as _durbin_watson,  # type: ignore[import-untyped]
)

from .regression_shared import (
    FEATURE_COLUMNS,
    RegressionObservationVectors,
    RegressionSummary,
    RegressionVectors,
)
from .write_sheet_csv_dataset import LIFE_EXPECTANCY, load_csv_rows

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_CSV = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"
TARGET_COLUMN = "Life expectancy"


def calculate_developed_country_flags(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
) -> tuple[bool, ...]:
    """Return expected "Developed Country after 2013" flags for each row.

    The Life Expectancy Data sheet appends a ``Developed Country after 2013``
    column computed with ``=AND([@Status]="Developed",[@Year]>2013)`` — TRUE
    for developed-country rows with Year >= 2014. This is the Python-side QC
    mirror of that formula, used by the source-row loader and the deep
    verifier. ``Status`` is the developed/developing text column and ``Year``
    is an integer, matching the typed values ``load_csv_rows`` returns.
    """
    headers, rows = load_csv_rows(input_csv_path, LIFE_EXPECTANCY)
    status_idx = headers.index("Status")
    year_idx = headers.index("Year")
    return tuple(
        bool(row[status_idx] == "Developed" and _is_int_over_2013(row[year_idx]))
        for row in rows
    )


def _is_int_over_2013(value: object) -> bool:
    """Year > 2013 test that tolerates int/float/None (None/blank -> False)."""
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value > 2013


def load_life_expectancy_source_rows(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[dict[str, object]]:
    """Load the source CSV as row dicts matching the LifeExpectancyData table.

    Cell values are typed exactly as the data sheet writer types them
    (int/float/str/None), and the appended ``Developed Country after 2013``
    column is computed with the same ``AND([@Status]="Developed",
    [@Year]>2013)`` rule the sheet's derived formula applies. Mirrors
    ``analyze_production_lots.load_production_lots_source_rows`` and
    ``analyze_model_construction.load_source_rows``, which are hardcoded to
    the Production Lots and Mileage tables respectively (neither of which
    appends a derived column).

    Headers come back normalized (``LIFE_EXPECTANCY.normalize_headers`` is
    True), so the keys are the collapsed-whitespace names the spec block and
    ``SPEC_DATASET_PROFILES["life_expectancy"]`` use — ``"Life expectancy"``,
    not the CSV's ``"Life expectancy "`` with its trailing space.
    """
    headers, rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)
    flags = calculate_developed_country_flags(csv_path)
    derived_header = LIFE_EXPECTANCY.derived_header
    assert derived_header is not None  # LE ships "Developed Country after 2013"
    table_headers = [*headers, derived_header]
    return [
        dict(zip(table_headers, [*row, flag]))
        for row, flag in zip(rows, flags)
    ]


def _normalize_header(name: str) -> str:
    """Collapse internal whitespace in a CSV header name.

    Parameters
    ----------
    name : str
        Raw header string from the CSV file.

    Returns
    -------
    str
        Header with leading/trailing whitespace stripped and internal
        runs of whitespace collapsed to a single space.
    """
    return " ".join(name.strip().split())


def _parse_float(raw: str | None) -> float | None:
    """Convert a raw CSV cell value to float, returning None for blanks.

    Parameters
    ----------
    raw : str or None
        The raw cell value from the CSV reader.

    Returns
    -------
    float or None
        Parsed float, or None if the value is None or an empty string.
    """
    if raw is None:
        return None
    value = raw.strip()
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _load_normalized_rows(
    input_path: Path,
) -> tuple[list[str], list[dict[str, str]]]:
    """Load CSV rows and normalize column header whitespace.

    Parameters
    ----------
    input_path : Path
        Path to the input CSV file.

    Returns
    -------
    tuple[list[str], list[dict[str, str]]]
        A 2-tuple of (original_headers, normalized_rows) where each
        normalized row maps normalized header names to raw cell strings.

    Raises
    ------
    ValueError
        If the CSV file has no header row.
    """
    with input_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None:
            raise ValueError(f"CSV has no headers: {input_path}")

        original_headers = list(reader.fieldnames)
        normalized_headers = [_normalize_header(name) for name in original_headers]
        normalized_rows: list[dict[str, str]] = []

        for row in reader:
            normalized_row = {
                normalized_name: row.get(original_name, "")
                for original_name, normalized_name in zip(
                    original_headers, normalized_headers, strict=True
                )
            }
            normalized_rows.append(normalized_row)

    return original_headers, normalized_rows


def _validate_required_headers(
    original_headers: list[str],
    feature_columns: list[str] | None = None,
) -> None:
    """Raise ValueError if any required column is absent from the CSV.

    Parameters
    ----------
    original_headers : list[str]
        Raw header names read from the CSV file.
    feature_columns : list[str] or None, optional
        Predictor columns to validate. Defaults to FEATURE_COLUMNS.

    Raises
    ------
    ValueError
        If any required column (target or predictor) is missing.
    """
    columns_to_check = (
        feature_columns if feature_columns is not None else FEATURE_COLUMNS
    )
    normalized_headers = [_normalize_header(name) for name in original_headers]
    missing_headers = [
        header
        for header in [TARGET_COLUMN, *columns_to_check]
        if header not in normalized_headers
    ]
    if missing_headers:
        raise ValueError(f"Missing required columns: {', '.join(missing_headers)}")


def _build_training_arrays(
    normalized_rows: list[dict[str, str]],
    include_intercept: bool,
    feature_columns: list[str] | None = None,
    filter_columns: list[str] | None = None,
) -> tuple[np.ndarray, np.ndarray, list[list[float] | None], list[float | None]]:
    """Build NumPy training arrays and track per-row parsed values.

    ``filter_columns`` controls which columns must be non-null for a row to be
    included. It defaults to ``feature_columns``, but callers can supply the
    full ``FEATURE_COLUMNS`` list to mimic an all-columns completeness filter
    that always checks every column regardless of how many predictors are
    actually used.

    Parameters
    ----------
    normalized_rows : list[dict[str, str]]
        Rows with normalized header keys from ``_load_normalized_rows``.
    include_intercept : bool
        If True, prepend a column of ones to the design matrix.
    feature_columns : list[str] or None, optional
        Predictor columns to use. Defaults to FEATURE_COLUMNS.
    filter_columns : list[str] or None, optional
        Columns that must be non-null for a row to be included in training.
        Defaults to ``feature_columns``.

    Returns
    -------
    tuple
        A 4-tuple of (x_train, y_train, parsed_features_per_row,
        parsed_targets_per_row).

    Raises
    ------
    ValueError
        If no training rows remain after filtering.
    """
    predictor_cols = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    filter_cols = filter_columns if filter_columns is not None else predictor_cols
    x_train_list: list[list[float]] = []
    y_train_list: list[float] = []
    parsed_features_per_row: list[list[float] | None] = []
    parsed_targets_per_row: list[float | None] = []

    for row in normalized_rows:
        predictor_values = [_parse_float(row.get(column)) for column in predictor_cols]
        filter_values = [_parse_float(row.get(column)) for column in filter_cols]
        target_value = _parse_float(row.get(TARGET_COLUMN))

        row_passes_filter = all(value is not None for value in filter_values)
        predictors_complete = all(value is not None for value in predictor_values)
        parsed_targets_per_row.append(target_value)

        if row_passes_filter and predictors_complete:
            dense_features = [float(value) for value in predictor_values if value is not None]
            parsed_features_per_row.append(dense_features)
            if target_value is not None:
                design_row = dense_features[:]
                if include_intercept:
                    design_row.insert(0, 1.0)
                x_train_list.append(design_row)
                y_train_list.append(target_value)
        else:
            parsed_features_per_row.append(None)

    if not x_train_list:
        raise ValueError("No training rows available for regression.")

    x_train = np.array(x_train_list, dtype=np.float64)
    y_train = np.array(y_train_list, dtype=np.float64)
    return x_train, y_train, parsed_features_per_row, parsed_targets_per_row


def _fit_ols_model(
    x_train: np.ndarray, y_train: np.ndarray, _include_intercept: bool
) -> RegressionResultsWrapper:
    """Fit an OLS model using statsmodels.

    Parameters
    ----------
    x_train : np.ndarray
        Design matrix (n_samples, n_features). The intercept column, when
        required, is already embedded by ``_build_training_arrays``.
    y_train : np.ndarray
        Target vector (n_samples,).
    _include_intercept : bool
        Unused; retained for call-site symmetry with other helpers.

    Returns
    -------
    statsmodels.regression.linear_model.RegressionResultsWrapper
        Fitted OLS results object.

    Notes
    -----
    The solver is pinned to ``method="qr"`` rather than left on the
    statsmodels default of ``"pinv"``. Both are backward stable, but on an
    ill-conditioned design the pseudoinverse is the less accurate of the two
    by several orders of magnitude, and this project's oracles are compared
    against the workbook at a decimal-place tolerance, so the oracle should
    be the more accurate side by as wide a margin as the choice allows.

    It also matches the workbook, whose ``Coefficients`` is built on LINEST —
    itself QR-based. Oracle and sheet solving the same problem by the same
    factorization is one fewer difference to account for when the two
    disagree.

    """
    # intercept column already embedded by _build_training_arrays
    model = sm.OLS(y_train, x_train)
    return model.fit(method="qr")


def _predict_single_row(
    coefficients: np.ndarray, features: list[float] | None, include_intercept: bool
) -> float | None:
    """Predict life expectancy for one row using fitted OLS coefficients.

    Parameters
    ----------
    coefficients : np.ndarray
        Fitted coefficient vector from the OLS model.
    features : list[float] or None
        Predictor values for the row, or None if the row was filtered out.
    include_intercept : bool
        If True, prepend 1.0 to the feature vector before the dot product.

    Returns
    -------
    float or None
        Predicted value, or None when features is None.
    """
    if features is None:
        return None
    design_row = np.array(features, dtype=np.float64)
    if include_intercept:
        design_row = np.concatenate([[1.0], design_row])
    return float(np.dot(design_row, coefficients))


def calculate_regression_summary(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    include_intercept: bool = True,
    feature_columns: list[str] | None = None,
    filter_columns: list[str] | None = None,
) -> RegressionSummary:
    """Fit OLS and return regression metrics without writing any output files.

    Parameters
    ----------
    input_csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    include_intercept : bool, optional
        If True, fit a model with an intercept term.
    feature_columns : list[str] or None, optional
        Predictor columns to include. Defaults to FEATURE_COLUMNS.
    filter_columns : list[str] or None, optional
        Columns that must be non-null for a row to be included. Defaults to
        the full ``FEATURE_COLUMNS`` (the all-columns completeness rule the
        legacy MLR test sheets bind to). Pass the model's own
        ``feature_columns`` for the spec-driven completeness the Regression
        sheet now uses.

    Returns
    -------
    RegressionSummary
        Scalar regression metrics matching the workbook LAMBDA outputs.
    """
    columns = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    filter_cols = filter_columns if filter_columns is not None else FEATURE_COLUMNS
    input_path = input_csv_path.resolve()
    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers, columns)

    x_train, y_train, _, _ = _build_training_arrays(
        normalized_rows, include_intercept, columns, filter_columns=filter_cols
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)

    observations = int(model.nobs)
    df_regression = len(columns)
    df_total = observations - 1 if include_intercept else observations
    r_squared = float(model.rsquared)
    df_residual = int(model.df_resid)
    multiple_r = float(np.sqrt(max(r_squared, 0.0)))
    adjusted_r2 = float(model.rsquared_adj)
    ss_total = float(model.centered_tss if include_intercept else model.uncentered_tss)
    ss_residual = float(model.ssr)
    ss_regression = ss_total - ss_residual
    se_regression = float(np.sqrt(model.mse_resid))

    # PRESS via hat-matrix diagonal shortcut: h = row-sums of (X * (X'X)^{-1} * X')
    xtx = x_train.T @ x_train
    xtx_inv = np.linalg.inv(xtx)
    z = x_train @ xtx_inv
    h = np.sum(z * x_train, axis=1)
    e = np.asarray(model.resid, dtype=np.float64)
    press = float(np.sum((e / (1.0 - h)) ** 2))

    dw = float(_durbin_watson(e))

    f_stat = (
        float((ss_regression / df_regression) / (ss_residual / df_residual))
        if df_regression > 0 and df_residual > 0 and ss_residual != 0.0
        else float("nan")
    )
    p_value_f = float(model.f_pvalue)

    # AIC/BIC/AICc: statsmodels .aic/.bic count σ² as a free parameter; use
    # the classical regression form matching the Excel LAMBDAs instead.
    _p = df_regression + (1 if include_intercept else 0)
    _log_term = float(observations * np.log(ss_residual / observations))
    aic = _log_term + 2.0 * _p
    bic = _log_term + _p * float(np.log(observations))
    aicc = aic + 2.0 * _p * (_p + 1) / (observations - _p - 1)

    # QQ correlation: Pearson r of sorted scaled residuals vs. normal scores.
    _sr = np.sort(e / se_regression)
    _ns = _scipy_stats.norm.ppf((np.arange(1, observations + 1) - 0.5) / observations)
    qq_correlation, _ = _scipy_stats.pearsonr(_sr, _ns)
    qq_correlation = float(qq_correlation)

    return RegressionSummary(
        observations=observations,
        df_regression=df_regression,
        df_total=df_total,
        r_squared=r_squared,
        df_residual=df_residual,
        multiple_r=multiple_r,
        adjusted_r2=adjusted_r2,
        ss_total=ss_total,
        ss_residual=ss_residual,
        ss_regression=ss_regression,
        se_regression=se_regression,
        press=press,
        durbin_watson=dw,
        # This module fits the plain pooled model — no Fixed Effects, so the
        # panel form has nothing to group by and the sheet's AE12 shows
        # "n/a — no fixed effects". NaN is that state.
        bfn_panel_durbin_watson=float("nan"),
        f_stat=f_stat,
        p_value_f=p_value_f,
        aic=aic,
        bic=bic,
        aicc=aicc,
        qq_correlation=qq_correlation,
    )


def calculate_regression_vectors(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    include_intercept: bool = True,
    feature_columns: list[str] | None = None,
    alpha: float = 0.05,
    filter_columns: list[str] | None = None,
) -> RegressionVectors:
    """Fit OLS and return per-coefficient statistics without writing any output files.

    Parameters
    ----------
    input_csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    include_intercept : bool, optional
        If True, fit a model with an intercept term.
    feature_columns : list[str] or None, optional
        Predictor columns to include. Defaults to FEATURE_COLUMNS.
    alpha : float, optional
        Significance level for confidence intervals. Default 0.05 yields 95% CIs.
    filter_columns : list[str] or None, optional
        Columns that must be non-null for a row to be included. Defaults to the
        full ``FEATURE_COLUMNS`` (the all-columns completeness rule). Pass the
        model's own ``feature_columns`` for the Regression sheet's spec-driven
        completeness.

    Returns
    -------
    RegressionVectors
        Per-coefficient statistics matching the workbook vector LAMBDA outputs.
    """
    columns = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    filter_cols = filter_columns if filter_columns is not None else FEATURE_COLUMNS
    input_path = input_csv_path.resolve()
    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers, columns)

    x_train, y_train, _, _ = _build_training_arrays(
        normalized_rows, include_intercept, columns, filter_columns=filter_cols
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)
    ci = model.conf_int(alpha=alpha)

    if include_intercept:
        term_names: tuple[str, ...] = ("Intercept", *columns)
    else:
        term_names = tuple(columns)

    ci_array = np.asarray(ci, dtype=np.float64)
    ci_lower_vals = tuple(float(v) for v in ci_array[:, 0])
    ci_upper_vals = tuple(float(v) for v in ci_array[:, 1])

    # Beta weights: b_j * std(X_j) / std(Y), predictor rows only.
    k = len(columns)
    x_features = x_train[:, 1:] if include_intercept else x_train
    std_y = float(np.std(y_train, ddof=1))
    coefs_arr = np.asarray(model.params, dtype=np.float64)
    pred_coefs = coefs_arr[1:] if include_intercept else coefs_arr
    beta_weights_vals = tuple(
        float(pred_coefs[j] * np.std(x_features[:, j], ddof=1) / std_y)
        for j in range(k)
    )

    return RegressionVectors(
        term_names=term_names,
        coefficients=tuple(float(v) for v in model.params),
        std_errors=tuple(float(v) for v in model.bse),
        t_stats=tuple(float(v) for v in model.tvalues),
        p_values=tuple(float(v) for v in model.pvalues),
        ci_lower=ci_lower_vals,
        ci_upper=ci_upper_vals,
        beta_weights=beta_weights_vals,
    )


def calculate_regression_observation_vectors(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    include_intercept: bool = True,
    feature_columns: list[str] | None = None,
) -> RegressionObservationVectors:
    """Fit OLS and return observation-level diagnostics without writing files.

    Parameters
    ----------
    input_csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    include_intercept : bool, optional
        If True, fit a model with an intercept term.
    feature_columns : list[str] or None, optional
        Predictor columns to include. Defaults to FEATURE_COLUMNS.

    Returns
    -------
    RegressionObservationVectors
        Observation-level diagnostics matching workbook spill LAMBDA outputs.
    """
    columns = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    input_path = input_csv_path.resolve()
    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers, columns)

    x_train, y_train, _, _ = _build_training_arrays(
        normalized_rows, include_intercept, columns, filter_columns=FEATURE_COLUMNS
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)

    n = len(y_train)
    observation_num = tuple(range(1, n + 1))
    sorted_y = np.sort(y_train)
    count_leq = np.searchsorted(sorted_y, y_train, side='right')   # SUMPRODUCT(filtered<=v)
    # Normal_Scores ranks on values rounded to 9 decimals so tied rows collapse to one
    # rank identically in Excel and the Python oracle, instead of being split by sub-ULP
    # float differences. Mirrors the Normal_Scores LAMBDA's ROUND(filtered, 9) and
    # analyze_regression_sheet.calculate_regression_results_from_matrix. rank_fraction
    # (count_leq) is left unrounded: it is a different field and not part of the tie fix.
    y_for_rank = np.round(y_train, 9)
    sorted_y_ranked = np.sort(y_for_rank)
    count_less = np.searchsorted(sorted_y_ranked, y_for_rank, side='left')   # SUMPRODUCT(ROUND(filtered,9)<ROUND(v,9))
    rank_fraction_array = count_leq / n
    normal_dist = NormalDist()
    normal_scores_array = np.array([normal_dist.inv_cdf(float((cl + 0.5) / n)) for cl in count_less])
    predictions = np.asarray(model.fittedvalues, dtype=np.float64)
    residuals = y_train - predictions
    se_regression = sqrt(float(model.mse_resid))
    scaled_residuals = residuals / se_regression

    return RegressionObservationVectors(
        observation_num=observation_num,
        rank_fraction=tuple(float(v) for v in rank_fraction_array),
        y_ranked=tuple(float(v) for v in np.sort(y_train)),
        normal_scores=tuple(float(v) for v in normal_scores_array),
        predictions=tuple(float(v) for v in predictions),
        residuals=tuple(float(v) for v in residuals),
        scaled_residuals=tuple(float(v) for v in scaled_residuals),
        scaled_residuals_ranked=tuple(float(v) for v in np.sort(scaled_residuals)),
    )

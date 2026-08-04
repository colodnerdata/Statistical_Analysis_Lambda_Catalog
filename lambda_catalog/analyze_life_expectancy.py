"""Fit OLS regression on the Life Expectancy dataset and compute QC reference values."""
from __future__ import annotations

import argparse
import csv
from math import sqrt
from statistics import NormalDist
from pathlib import Path

import numpy as np
import statsmodels.api as sm  # type: ignore[import-untyped]
from scipy import stats as _scipy_stats  # type: ignore[import-untyped]
from statsmodels.regression.linear_model import (  # type: ignore[import-untyped]
    RegressionResultsWrapper,
)
from statsmodels.stats.stattools import durbin_watson as _durbin_watson  # type: ignore[import-untyped]

from .regression_shared import (
    FEATURE_COLUMNS,
    RegressionObservationVectors,
    RegressionSummary,
    RegressionVectors,
)
from .write_sheet_csv_dataset import LIFE_EXPECTANCY, load_csv_rows

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_CSV = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"
DEFAULT_OUTPUT_CSV = ROOT_DIR / "Life Expectancy Predictions.csv"
TARGET_COLUMN = "Life expectancy"


def calculate_data_completeness_flags(
    input_csv_path: Path,
) -> tuple[bool, ...]:
    """Return expected Full_Data flags matching the workbook Data_Completeness formula.

    The Life Expectancy Data sheet computes ``Full_Data`` with
    ``Data_Completeness(LifeExpectancyData[@[Life expectancy]:[Schooling]])``,
    so the Python-side QC expectation is that the target and all predictor
    columns parse as numeric on a given row.
    """
    original_headers, normalized_rows = _load_normalized_rows(input_csv_path)
    _validate_required_headers(original_headers)

    required_columns = [TARGET_COLUMN, *FEATURE_COLUMNS]
    return tuple(
        all(_parse_float(row.get(column)) is not None for column in required_columns)
        for row in normalized_rows
    )


def load_life_expectancy_source_rows(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[dict[str, object]]:
    """Load the source CSV as row dicts matching the LifeExpectancyData table.

    Cell values are typed exactly as the data sheet writer types them
    (int/float/str/None), and the computed ``Full_Data`` column is appended
    with the same completeness rule the sheet's ``Data_Completeness`` formula
    applies. Mirrors ``analyze_production_lots.load_production_lots_source_rows``
    and ``analyze_model_construction.load_source_rows``, which are hardcoded to
    the Production Lots and Mileage tables respectively.

    Headers come back normalized (``LIFE_EXPECTANCY.normalize_headers`` is
    True), so the keys are the collapsed-whitespace names the spec block and
    ``SPEC_DATASET_PROFILES["life_expectancy"]`` use — ``"Life expectancy"``,
    not the CSV's ``"Life expectancy "`` with its trailing space.
    """
    headers, rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)
    flags = calculate_data_completeness_flags(csv_path)
    table_headers = [*headers, LIFE_EXPECTANCY.full_data_header]
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
    full ``FEATURE_COLUMNS`` list to match an Excel Full_Data filter that always
    checks all columns regardless of how many predictors are actually used.

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
    """
    # intercept column already embedded by _build_training_arrays
    model = sm.OLS(y_train, x_train)
    return model.fit()


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
        the full ``FEATURE_COLUMNS`` (the ``Full_Data`` completeness rule the
        MLR test sheets bind to). Pass the model's own ``feature_columns`` for
        the spec-driven completeness the Regression sheet now uses.

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
        full ``FEATURE_COLUMNS`` (Full_Data). Pass the model's own
        ``feature_columns`` for the Regression sheet's spec-driven completeness.

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
    count_less = np.searchsorted(sorted_y, y_train, side='left')   # SUMPRODUCT(filtered<v)
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


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for life expectancy prediction.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with csv, output, and no_intercept attributes.
    """
    parser = argparse.ArgumentParser(
        description="Fit OLS on Life Expectancy data and output predicted values to CSV."
    )
    parser.add_argument("--csv", type=Path, default=DEFAULT_INPUT_CSV, help="Input CSV path.")
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_CSV,
        help="Output CSV path with predicted values.",
    )
    parser.add_argument(
        "--no-intercept",
        action="store_true",
        help="Fit model without an intercept term.",
    )
    return parser.parse_args()


def generate_life_expectancy_predictions(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    output_csv_path: Path = DEFAULT_OUTPUT_CSV,
    include_intercept: bool = True,
) -> tuple[int, int]:
    """Fit OLS predictions and write an output CSV.

    Parameters
    ----------
    input_csv_path : Path, optional
        Path to the input Life Expectancy CSV file.
    output_csv_path : Path, optional
        Path where the output CSV with predictions will be written.
    include_intercept : bool, optional
        If True, fit a model with an intercept term.

    Returns
    -------
    tuple[int, int]
        A 2-tuple of (training_rows_used, rows_with_predictions).
    """
    input_path = input_csv_path.resolve()
    output_path = output_csv_path.resolve()

    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers)

    x_train, y_train, parsed_features_per_row, parsed_targets_per_row = (
        _build_training_arrays(normalized_rows, include_intercept)
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)
    coefficients = model.params

    output_headers = original_headers + ["Predicted_Life_expectancy", "Residual"]
    predicted_count = 0

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=output_headers)
        writer.writeheader()

        for original_row, features, target in zip(
            normalized_rows,
            parsed_features_per_row,
            parsed_targets_per_row,
            strict=True,
        ):
            out_row = {
                header: original_row.get(_normalize_header(header), "")
                for header in original_headers
            }
            predicted_value: float | None = None
            residual: float | None = None

            if features is not None:
                predicted_value = _predict_single_row(coefficients, features, include_intercept)
                predicted_count += 1

                if target is not None and predicted_value is not None:
                    residual = target - predicted_value

            out_row["Predicted_Life_expectancy"] = (
                "" if predicted_value is None else f"{predicted_value:.6f}"
            )
            out_row["Residual"] = "" if residual is None else f"{residual:.6f}"
            writer.writerow(out_row)

    coef_names = FEATURE_COLUMNS[:]
    if include_intercept:
        coef_names.insert(0, "Intercept")

    print(f"Input CSV: {input_path}")
    print(f"Output CSV: {output_path}")
    print(f"Training rows used: {len(y_train)}")
    print(f"Rows with predictions: {predicted_count}")
    for name, coef in zip(coef_names, coefficients, strict=True):
        print(f"{name}: {coef:.10g}")

    return len(y_train), predicted_count


def main() -> None:
    """Fit predictions and write output CSV from command-line arguments."""
    args = parse_args()
    generate_life_expectancy_predictions(
        input_csv_path=args.csv,
        output_csv_path=args.output,
        include_intercept=not args.no_intercept,
    )


if __name__ == "__main__":
    main()

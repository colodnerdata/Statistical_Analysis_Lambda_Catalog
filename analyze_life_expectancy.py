from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import statsmodels.api as sm


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_CSV = ROOT_DIR / "Life Expectancy Data.csv"
DEFAULT_OUTPUT_CSV = ROOT_DIR / "Life Expectancy Predictions.csv"
TARGET_COLUMN = "Life expectancy"
FEATURE_COLUMNS = [
    "Adult Mortality",
    "infant deaths",
    "Alcohol",
    "percentage expenditure",
    "Hepatitis B",
    "Measles",
    "BMI",
    "under-five deaths",
    "Polio",
    "Total expenditure",
    "Diphtheria",
    "HIV/AIDS",
    "GDP",
    "Population",
    "thinness 1-19 years",
    "thinness 5-9 years",
    "Income composition of resources",
    "Schooling",
]


@dataclass(frozen=True)
class RegressionSummary:
    """Core regression metrics aligned with workbook LAMBDA outputs."""

    observations: int
    df_regression: int
    df_total: int
    r_squared: float
    df_residual: int
    multiple_r: float
    adjusted_r2: float
    ss_total: float
    ss_residual: float
    ss_regression: float
    se_regression: float


def _normalize_header(name: str) -> str:
    return " ".join(name.strip().split())


def _parse_float(raw: str | None) -> float | None:
    if raw is None:
        return None
    value = raw.strip()
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _load_normalized_rows(input_path: Path) -> tuple[list[str], list[dict[str, str]]]:
    """Load CSV and normalize headers."""
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
                for original_name, normalized_name in zip(original_headers, normalized_headers, strict=True)
            }
            normalized_rows.append(normalized_row)

    return original_headers, normalized_rows


def _validate_required_headers(
    original_headers: list[str],
    feature_columns: list[str] | None = None,
) -> None:
    """Check that all required columns are present."""
    columns_to_check = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    normalized_headers = [_normalize_header(name) for name in original_headers]
    missing_headers = [
        header for header in [TARGET_COLUMN, *columns_to_check] if header not in normalized_headers
    ]
    if missing_headers:
        raise ValueError(f"Missing required columns: {', '.join(missing_headers)}")


def _build_training_arrays(
    normalized_rows: list[dict[str, str]],
    include_intercept: bool,
    feature_columns: list[str] | None = None,
    filter_columns: list[str] | None = None,
) -> tuple[np.ndarray, np.ndarray, list[list[float] | None], list[float | None]]:
    """Build NumPy arrays for training and track per-row parsed values.

    filter_columns controls which columns must be non-null for a row to be included.
    It defaults to feature_columns, but callers can supply the full FEATURE_COLUMNS list
    to match an Excel filter (e.g. Full_Data) that always checks all columns regardless
    of how many predictors are actually used.
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
            dense_features = [float(value) for value in predictor_values]
            parsed_features_per_row.append(dense_features)
        else:
            parsed_features_per_row.append(None)

        if row_passes_filter and predictors_complete and target_value is not None:
            design_row = dense_features[:]
            if include_intercept:
                design_row.insert(0, 1.0)
            x_train_list.append(design_row)
            y_train_list.append(target_value)

    if not x_train_list:
        raise ValueError("No training rows available for regression.")

    x_train = np.array(x_train_list, dtype=np.float64)
    y_train = np.array(y_train_list, dtype=np.float64)
    return x_train, y_train, parsed_features_per_row, parsed_targets_per_row


def _fit_ols_model(
    x_train: np.ndarray, y_train: np.ndarray, include_intercept: bool
) -> sm.regression.linear_model.RegressionResults:
    """Fit OLS model using statsmodels."""
    if include_intercept:
        model = sm.OLS(y_train, x_train)
    else:
        model = sm.OLS(y_train, x_train)
    return model.fit()


def _predict_single_row(
    coefficients: np.ndarray, features: list[float] | None, include_intercept: bool
) -> float | None:
    """Predict a single row using fitted coefficients."""
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
) -> RegressionSummary:
    """Fit OLS and return regression metrics without writing output files."""

    columns = feature_columns if feature_columns is not None else FEATURE_COLUMNS
    input_path = input_csv_path.resolve()
    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers, columns)

    x_train, y_train, _, _ = _build_training_arrays(
        normalized_rows, include_intercept, columns, filter_columns=FEATURE_COLUMNS
    )
    model = _fit_ols_model(x_train, y_train, include_intercept)

    observations = int(model.nobs)
    df_regression = len(columns)
    df_total = observations - 1 if include_intercept else observations
    r_squared = model.rsquared
    df_residual = int(model.df_resid)
    multiple_r = np.sqrt(max(r_squared, 0.0))
    adjusted_r2 = model.rsquared_adj
    ss_total = float(model.centered_tss if include_intercept else model.uncentered_tss)
    ss_residual = float(model.ssr)
    ss_regression = ss_total - ss_residual
    se_regression = float(np.sqrt(model.mse_resid))

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
    )


def parse_args() -> argparse.Namespace:
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

    Returns a tuple of (training_rows_used, rows_with_predictions).
    """

    input_path = input_csv_path.resolve()
    output_path = output_csv_path.resolve()

    original_headers, normalized_rows = _load_normalized_rows(input_path)
    _validate_required_headers(original_headers)

    x_train, y_train, parsed_features_per_row, parsed_targets_per_row = _build_training_arrays(
        normalized_rows,
        include_intercept,
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
            out_row = {header: original_row.get(_normalize_header(header), "") for header in original_headers}
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
    args = parse_args()
    generate_life_expectancy_predictions(
        input_csv_path=args.csv,
        output_csv_path=args.output,
        include_intercept=not args.no_intercept,
    )


if __name__ == "__main__":
    main()

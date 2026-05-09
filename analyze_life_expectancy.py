from __future__ import annotations

import argparse
import csv
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence


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


def _dot(left: Sequence[float], right: Sequence[float]) -> float:
    return sum(a * b for a, b in zip(left, right, strict=True))


def _transpose(matrix: list[list[float]]) -> list[list[float]]:
    if not matrix:
        return []
    return [list(col) for col in zip(*matrix, strict=True)]


def _matmul(left: list[list[float]], right: list[list[float]]) -> list[list[float]]:
    if not left or not right:
        return []
    right_t = _transpose(right)
    return [[_dot(row, col) for col in right_t] for row in left]


def _solve_linear_system(a: list[list[float]], b: list[float]) -> list[float]:
    """Solve Ax=b with Gauss-Jordan elimination and partial pivoting."""

    n = len(a)
    if n == 0:
        raise ValueError("Cannot solve an empty linear system.")

    augmented = [row[:] + [b_val] for row, b_val in zip(a, b, strict=True)]

    for pivot_col in range(n):
        pivot_row = pivot_col
        pivot_abs = abs(augmented[pivot_col][pivot_col])
        for row in range(pivot_col + 1, n):
            current_abs = abs(augmented[row][pivot_col])
            if current_abs > pivot_abs:
                pivot_abs = current_abs
                pivot_row = row
        pivot_value = augmented[pivot_row][pivot_col]

        if math.isclose(pivot_value, 0.0, abs_tol=1e-12):
            raise ValueError("Design matrix is singular; regression cannot be solved with current inputs.")

        if pivot_row != pivot_col:
            augmented[pivot_col], augmented[pivot_row] = augmented[pivot_row], augmented[pivot_col]

        pivot_value = augmented[pivot_col][pivot_col]
        augmented[pivot_col] = [value / pivot_value for value in augmented[pivot_col]]

        for row in range(n):
            if row == pivot_col:
                continue
            factor = augmented[row][pivot_col]
            if math.isclose(factor, 0.0, abs_tol=1e-15):
                continue
            augmented[row] = [
                current - factor * pivot
                for current, pivot in zip(augmented[row], augmented[pivot_col], strict=True)
            ]

    return [row[-1] for row in augmented]


def _fit_ols(x_rows: list[list[float]], y_values: list[float]) -> list[float]:
    if not x_rows:
        raise ValueError("No training rows available for regression.")

    x_t = _transpose(x_rows)
    xtx = _matmul(x_t, x_rows)
    xty = [_dot(col, y_values) for col in x_t]
    return _solve_linear_system(xtx, xty)


def _build_training_matrices(
    normalized_rows: list[dict[str, str]],
    include_intercept: bool,
) -> tuple[list[list[float]], list[float], list[list[float] | None], list[float | None]]:
    """Build regression matrices and parsed row values from normalized CSV rows."""

    x_train: list[list[float]] = []
    y_train: list[float] = []
    parsed_features_per_row: list[list[float] | None] = []
    parsed_targets_per_row: list[float | None] = []

    for row in normalized_rows:
        feature_values = [_parse_float(row.get(column)) for column in FEATURE_COLUMNS]
        target_value = _parse_float(row.get(TARGET_COLUMN))

        features_complete = all(value is not None for value in feature_values)
        parsed_targets_per_row.append(target_value)

        if features_complete:
            dense_features = [float(value) for value in feature_values if value is not None]
            parsed_features_per_row.append(dense_features)
        else:
            parsed_features_per_row.append(None)

        if features_complete and target_value is not None:
            design_row = dense_features[:]
            if include_intercept:
                design_row.insert(0, 1.0)
            x_train.append(design_row)
            y_train.append(target_value)

    return x_train, y_train, parsed_features_per_row, parsed_targets_per_row


def _compute_regression_summary(
    x_train: list[list[float]],
    y_train: list[float],
    coefficients: list[float],
    include_intercept: bool,
) -> RegressionSummary:
    """Compute regression summary metrics from fitted data."""

    observations = len(y_train)
    df_regression = len(FEATURE_COLUMNS)
    df_total = observations - 1 if include_intercept else observations
    fitted_values = [_dot(row, coefficients) for row in x_train]
    residuals = [actual - fitted for actual, fitted in zip(y_train, fitted_values, strict=True)]
    ss_residual = sum(residual * residual for residual in residuals)

    if include_intercept:
        mean_y = sum(y_train) / observations
        ss_total = sum((actual - mean_y) ** 2 for actual in y_train)
    else:
        ss_total = sum(actual * actual for actual in y_train)

    if math.isclose(ss_total, 0.0, abs_tol=1e-15):
        r_squared = 1.0 if math.isclose(ss_residual, 0.0, abs_tol=1e-15) else 0.0
    else:
        r_squared = 1.0 - (ss_residual / ss_total)

    df_residual = df_total - df_regression
    multiple_r = math.sqrt(max(r_squared, 0.0))
    if df_residual <= 0:
        adjusted_r2 = float("nan")
    else:
        adjusted_r2 = 1.0 - (1.0 - r_squared) * (df_total / df_residual)

    return RegressionSummary(
        observations=observations,
        df_regression=df_regression,
        df_total=df_total,
        r_squared=r_squared,
        df_residual=df_residual,
        multiple_r=multiple_r,
        adjusted_r2=adjusted_r2,
    )


def calculate_regression_summary(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
    include_intercept: bool = True,
) -> RegressionSummary:
    """Fit OLS and return regression metrics without writing output files."""

    input_path = input_csv_path.resolve()

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

    missing_headers = [
        header for header in [TARGET_COLUMN, *FEATURE_COLUMNS] if header not in normalized_headers
    ]
    if missing_headers:
        raise ValueError(f"Missing required columns: {', '.join(missing_headers)}")

    x_train, y_train, _, _ = _build_training_matrices(normalized_rows, include_intercept)
    coefficients = _fit_ols(x_train, y_train)
    return _compute_regression_summary(x_train, y_train, coefficients, include_intercept)


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

    missing_headers = [
        header for header in [TARGET_COLUMN, *FEATURE_COLUMNS] if header not in normalized_headers
    ]
    if missing_headers:
        raise ValueError(f"Missing required columns: {', '.join(missing_headers)}")

    x_train, y_train, parsed_features_per_row, parsed_targets_per_row = _build_training_matrices(
        normalized_rows,
        include_intercept,
    )
    coefficients = _fit_ols(x_train, y_train)

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
                predict_row = features[:]
                if include_intercept:
                    predict_row.insert(0, 1.0)
                predicted_value = _dot(predict_row, coefficients)
                predicted_count += 1

                if target is not None:
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

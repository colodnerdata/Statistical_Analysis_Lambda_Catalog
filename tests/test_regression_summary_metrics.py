from __future__ import annotations

import csv
import math
import tempfile
import unittest
from pathlib import Path

import numpy as np

from lambda_catalog.analyze_life_expectancy import (
    FEATURE_COLUMNS,
    TARGET_COLUMN,
    _fit_ols_model,
    _load_normalized_rows,
    calculate_regression_summary,
)


def _write_csv(path: Path, rows: list[dict[str, float]]) -> None:
    headers = [TARGET_COLUMN, *FEATURE_COLUMNS]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({h: row[h] for h in headers})


def _row(y: float, x1: float, x2: float = 0.0) -> dict[str, float]:
    values = {name: 1.0 for name in FEATURE_COLUMNS}
    values[FEATURE_COLUMNS[0]] = x1
    values[FEATURE_COLUMNS[1]] = x2
    values[TARGET_COLUMN] = y
    return values


class RegressionSummaryMetricsTests(unittest.TestCase):
    def test_aic_bic_aicc_match_manual_formulas(self) -> None:
        rows = [
            _row(7.2, 1.0, 2.0),
            _row(8.0, 2.0, 0.5),
            _row(9.4, 3.0, 1.5),
            _row(10.1, 4.0, 1.0),
            _row(11.8, 5.0, 2.5),
            _row(12.3, 6.0, 2.0),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "metrics.csv"
            _write_csv(csv_path, rows)
            summary = calculate_regression_summary(
                input_csv_path=csv_path,
                include_intercept=True,
                feature_columns=FEATURE_COLUMNS[:2],
            )

        n = summary.observations
        p = summary.df_regression + 1
        log_term = n * math.log(summary.ss_residual / n)
        expected_aic = log_term + 2.0 * p
        expected_bic = log_term + p * math.log(n)
        expected_aicc = expected_aic + 2.0 * p * (p + 1) / (n - p - 1)

        self.assertAlmostEqual(summary.aic, expected_aic, places=12)
        self.assertAlmostEqual(summary.bic, expected_bic, places=12)
        self.assertAlmostEqual(summary.aicc, expected_aicc, places=12)
        self.assertTrue(math.isfinite(summary.qq_correlation))
        self.assertLessEqual(abs(summary.qq_correlation), 1.0)

    def test_aicc_denominator_zero_raises(self) -> None:
        # n=3 and p=2 (1 predictor + intercept) gives n-p-1 = 0.
        rows = [
            _row(3.0, 1.0),
            _row(5.0, 2.0),
            _row(7.0, 3.0),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "aicc_zero_denom.csv"
            _write_csv(csv_path, rows)
            with self.assertRaises(ZeroDivisionError):
                calculate_regression_summary(
                    input_csv_path=csv_path,
                    include_intercept=True,
                    feature_columns=FEATURE_COLUMNS[:1],
                )

    def test_near_perfect_fit_keeps_metrics_well_defined(self) -> None:
        # Near-perfect fit drives SSR very close to zero.
        rows = [
            _row(7.0, 1.0, 1.0),
            _row(9.0, 2.0, 1.0),
            _row(19.0, 3.0, 3.0),
            _row(17.0, 4.0, 2.0),
            _row(19.0, 5.0, 2.0),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "perfect_fit.csv"
            _write_csv(csv_path, rows)
            summary = calculate_regression_summary(
                input_csv_path=csv_path,
                include_intercept=True,
                feature_columns=FEATURE_COLUMNS[:2],
            )

            _, normalized_rows = _load_normalized_rows(csv_path)
            x_data = []
            y_data = []
            for source in normalized_rows:
                y_data.append(float(source[TARGET_COLUMN]))
                x_data.append([1.0, float(source[FEATURE_COLUMNS[0]]), float(source[FEATURE_COLUMNS[1]])])
            model = _fit_ols_model(np.array(x_data, dtype=np.float64), np.array(y_data, dtype=np.float64), True)
            self.assertLess(float(model.ssr), 1e-20)

        self.assertTrue(math.isfinite(summary.aic))
        self.assertTrue(math.isfinite(summary.bic))
        self.assertLess(summary.aic, 0.0)
        self.assertLess(summary.bic, 0.0)
        if math.isfinite(summary.qq_correlation):
            self.assertLessEqual(abs(summary.qq_correlation), 1.0)


if __name__ == "__main__":
    unittest.main()

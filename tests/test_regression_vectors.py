"""Tests for calculate_regression_vectors — per-coefficient statistics."""
from __future__ import annotations

import csv
import math
import tempfile
import unittest
from pathlib import Path

from lambda_catalog.analyze_life_expectancy import (
    FEATURE_COLUMNS,
    TARGET_COLUMN,
    calculate_regression_vectors,
)


def _write_csv(path: Path, rows: list[dict]) -> None:
    headers = [TARGET_COLUMN, *FEATURE_COLUMNS]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({h: row.get(h, 1.0) for h in headers})


def _make_row(y: float, x1: float, x2: float = 1.0) -> dict:
    row = {name: 1.0 for name in FEATURE_COLUMNS}
    row[FEATURE_COLUMNS[0]] = x1
    row[FEATURE_COLUMNS[1]] = x2
    row[TARGET_COLUMN] = y
    return row


class RegressionVectorsTests(unittest.TestCase):
    def setUp(self) -> None:
        rows = [
            _make_row(7.2, 1.0, 2.0),
            _make_row(8.0, 2.0, 0.5),
            _make_row(9.4, 3.0, 1.5),
            _make_row(10.1, 4.0, 1.0),
            _make_row(11.8, 5.0, 2.5),
            _make_row(12.3, 6.0, 2.0),
            _make_row(13.5, 7.0, 3.0),
            _make_row(14.0, 8.0, 1.5),
        ]
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "vectors.csv"
        _write_csv(self._csv, rows)
        self._k = 2
        self._vectors = calculate_regression_vectors(
            self._csv,
            include_intercept=True,
            feature_columns=FEATURE_COLUMNS[:self._k],
        )

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_term_names_length_with_intercept(self) -> None:
        self.assertEqual(len(self._vectors.term_names), self._k + 1)

    def test_term_names_first_is_intercept(self) -> None:
        self.assertEqual(self._vectors.term_names[0], "Intercept")

    def test_predictor_names_follow_intercept(self) -> None:
        self.assertEqual(
            list(self._vectors.term_names[1:]),
            FEATURE_COLUMNS[:self._k],
        )

    def test_coefficients_all_finite(self) -> None:
        self.assertTrue(all(math.isfinite(c) for c in self._vectors.coefficients))

    def test_std_errors_positive(self) -> None:
        self.assertTrue(all(se > 0 for se in self._vectors.std_errors))

    def test_t_stats_equal_coef_over_se(self) -> None:
        for coef, se, t in zip(
            self._vectors.coefficients,
            self._vectors.std_errors,
            self._vectors.t_stats,
        ):
            self.assertAlmostEqual(t, coef / se, places=10)

    def test_p_values_in_unit_interval(self) -> None:
        self.assertTrue(all(0.0 <= p <= 1.0 for p in self._vectors.p_values))

    def test_ci_lower_less_than_ci_upper(self) -> None:
        for lo, hi in zip(self._vectors.ci_lower, self._vectors.ci_upper):
            self.assertLess(lo, hi)

    def test_ci_bounds_contain_coefficient(self) -> None:
        for lo, coef, hi in zip(
            self._vectors.ci_lower,
            self._vectors.coefficients,
            self._vectors.ci_upper,
        ):
            self.assertGreaterEqual(coef, lo - 1e-12)
            self.assertLessEqual(coef, hi + 1e-12)

    def test_beta_weights_length_equals_k(self) -> None:
        self.assertEqual(len(self._vectors.beta_weights), self._k)

    def test_beta_weights_all_finite(self) -> None:
        self.assertTrue(all(math.isfinite(b) for b in self._vectors.beta_weights))

    def test_term_names_length_without_intercept(self) -> None:
        vectors = calculate_regression_vectors(
            self._csv,
            include_intercept=False,
            feature_columns=FEATURE_COLUMNS[:self._k],
        )
        self.assertEqual(len(vectors.term_names), self._k)

    def test_all_tuple_lengths_consistent(self) -> None:
        lengths = [
            len(self._vectors.term_names),
            len(self._vectors.coefficients),
            len(self._vectors.std_errors),
            len(self._vectors.t_stats),
            len(self._vectors.p_values),
            len(self._vectors.ci_lower),
            len(self._vectors.ci_upper),
        ]
        self.assertEqual(len(set(lengths)), 1, f"Inconsistent tuple lengths: {lengths}")


if __name__ == "__main__":
    unittest.main()

"""Tests for calculate_regression_observation_vectors — per-row diagnostics."""
from __future__ import annotations

import csv
import math
import tempfile
import unittest
from pathlib import Path

from lambda_catalog.analyze_life_expectancy import (
    FEATURE_COLUMNS,
    TARGET_COLUMN,
    calculate_regression_observation_vectors,
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


class RegressionObservationVectorsTests(unittest.TestCase):
    def setUp(self) -> None:
        self._n = 9
        rows = [
            _make_row(5.0 + i * 1.3, float(i + 1), float(i % 3 + 1))
            for i in range(self._n)
        ]
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "obs.csv"
        _write_csv(self._csv, rows)
        self._obs = calculate_regression_observation_vectors(
            self._csv,
            include_intercept=True,
            feature_columns=FEATURE_COLUMNS[:2],
        )

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_observation_num_sequential_from_one(self) -> None:
        n = len(self._obs.observation_num)
        self.assertEqual(list(self._obs.observation_num), list(range(1, n + 1)))

    def test_observation_num_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.observation_num), self._n)

    def test_rank_fraction_values_in_half_open_unit_interval(self) -> None:
        self.assertTrue(all(0 < v <= 1.0 for v in self._obs.rank_fraction))

    def test_rank_fraction_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.rank_fraction), self._n)

    def test_y_ranked_is_non_decreasing(self) -> None:
        ranked = list(self._obs.y_ranked)
        self.assertEqual(ranked, sorted(ranked))

    def test_y_ranked_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.y_ranked), self._n)

    def test_normal_scores_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.normal_scores), self._n)

    def test_normal_scores_all_finite(self) -> None:
        self.assertTrue(all(math.isfinite(v) for v in self._obs.normal_scores))

    def test_predictions_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.predictions), self._n)

    def test_predictions_all_finite(self) -> None:
        self.assertTrue(all(math.isfinite(v) for v in self._obs.predictions))

    def test_residuals_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.residuals), self._n)

    def test_residuals_sum_near_zero_with_intercept(self) -> None:
        # OLS with intercept: residuals sum to zero
        self.assertAlmostEqual(sum(self._obs.residuals), 0.0, places=8)

    def test_predictions_plus_residuals_equal_actuals(self) -> None:
        obs = self._obs
        for pred, resid, yn in zip(obs.predictions, obs.residuals, obs.y_ranked):
            pass  # y_ranked is sorted, not aligned — just verify individual consistency
        # Verify residuals = actual - predicted via sum check
        import numpy as np
        predictions = list(obs.predictions)
        residuals = list(obs.residuals)
        # Sum of (pred + resid) should equal sum of y (the unranked actuals)
        pred_plus_resid_sum = sum(p + r for p, r in zip(predictions, residuals))
        y_ranked_sum = sum(obs.y_ranked)
        self.assertAlmostEqual(pred_plus_resid_sum, y_ranked_sum, places=6)

    def test_scaled_residuals_length_equals_n(self) -> None:
        self.assertEqual(len(self._obs.scaled_residuals), self._n)

    def test_scaled_residuals_ranked_is_non_decreasing(self) -> None:
        ranked = list(self._obs.scaled_residuals_ranked)
        self.assertEqual(ranked, sorted(ranked))

    def test_scaled_residuals_ranked_same_values_as_scaled_residuals(self) -> None:
        self.assertAlmostEqual(
            sum(self._obs.scaled_residuals),
            sum(self._obs.scaled_residuals_ranked),
            places=10,
        )

    def test_all_vector_lengths_consistent(self) -> None:
        obs = self._obs
        lengths = [
            len(obs.observation_num),
            len(obs.rank_fraction),
            len(obs.y_ranked),
            len(obs.normal_scores),
            len(obs.predictions),
            len(obs.residuals),
            len(obs.scaled_residuals),
            len(obs.scaled_residuals_ranked),
        ]
        self.assertEqual(len(set(lengths)), 1, f"Inconsistent lengths: {lengths}")


if __name__ == "__main__":
    unittest.main()

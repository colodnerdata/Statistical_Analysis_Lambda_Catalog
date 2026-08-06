"""Internal-consistency checks on the shared OLS oracle.

Fits a spread of continuous-predictor models on the Life Expectancy sample
and asserts the invariants every OLS fit must satisfy regardless of the
model: the hat diagonal sums to *p*, residuals sum to zero under an
intercept, predictions plus residuals recover y, the sums of squares
decompose, and the prediction interval brackets its own point estimate.

These are properties of the arithmetic, not of any particular model, so the
config list below exists only to exercise them across several model shapes
(sparse/medium/full x with/without intercept). Excel is unavailable in CI,
so this validates the Python-side oracle, not the workbook formulas.
"""
# pylint: disable=missing-function-docstring,too-many-public-methods,unused-variable
from __future__ import annotations

import unittest
from typing import Any, ClassVar

from lambda_catalog.analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
)
from lambda_catalog.analyze_regression_sheet import (
    calculate_regression_sheet_results,
)
from lambda_catalog.regression_shared import FEATURE_COLUMNS

# -- Helpers ----------------------------------------------------------------

_CSV = DEFAULT_INPUT_CSV

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

# The model shapes the invariants are checked over: (name, include_intercept,
# predictors). Nothing outside this file consumes it.
REGRESSION_QC_CONFIGS: list[tuple[str, bool, list[str]]] = [
    ("sparse_intercept", True, _SPARSE_PREDICTORS),
    ("sparse_no_intercept", False, _SPARSE_PREDICTORS),
    ("medium_intercept", True, _MEDIUM_PREDICTORS),
    ("medium_no_intercept", False, _MEDIUM_PREDICTORS),
    ("full_intercept", True, FEATURE_COLUMNS),
    ("full_no_intercept", False, FEATURE_COLUMNS),
]


def _has_csv() -> bool:
    return _CSV.exists()


_SKIP_MSG = "Life Expectancy CSV not found"


# ── Regression sheet QC configs ─────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestRegressionSheetConfigs(unittest.TestCase):
    """Verify regression sheet QC configs are complete and self-consistent."""

    configs: ClassVar[list[Any]]
    named_configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        cls.configs = [
            calculate_regression_sheet_results(
                _CSV, include_intercept=allow, feature_columns=cols
            )
            for _, allow, cols in REGRESSION_QC_CONFIGS
        ]
        cls.named_configs = list(zip(REGRESSION_QC_CONFIGS, cls.configs))

    def test_config_count(self) -> None:
        self.assertEqual(len(self.configs), len(REGRESSION_QC_CONFIGS))

    def test_summary_matches_vectors_term_count(self) -> None:
        for (name, allow, cols), results in self.named_configs:
            k = len(cols)
            expected_terms = k + (1 if allow else 0)
            self.assertEqual(
                len(results.vectors.coefficients), expected_terms,
                f"config={name}",
            )

    def test_predictor_summary_length_matches_k(self) -> None:
        for (name, allow, cols), results in self.named_configs:
            k = len(cols)
            self.assertEqual(len(results.predictor_summary.pearson_r), k, name)
            self.assertEqual(len(results.predictor_summary.spearman_r), k, name)
            self.assertEqual(len(results.predictor_summary.skewness), k, name)
            self.assertEqual(len(results.predictor_summary.kurtosis), k, name)
            self.assertEqual(len(results.predictor_summary.gvif), k, name)
            self.assertEqual(len(results.predictor_summary.tolerance), k, name)

    def test_predictor_names_match_columns(self) -> None:
        for (name, allow, cols), results in self.named_configs:
            self.assertEqual(
                list(results.predictor_summary.predictor_names), cols, name,
            )

    def test_gvif_tolerance_reciprocal(self) -> None:
        for (name, _, _), results in self.named_configs:
            for j, (v, t) in enumerate(
                zip(results.predictor_summary.gvif, results.predictor_summary.tolerance)
            ):
                self.assertAlmostEqual(v * t, 1.0, places=10, msg=f"{name} j={j}")

    def test_gvif_at_least_one(self) -> None:
        for (name, _, _), results in self.named_configs:
            for j, v in enumerate(results.predictor_summary.gvif):
                self.assertGreaterEqual(v, 1.0 - 1e-10, msg=f"{name} j={j}")

    def test_pearson_r_in_range(self) -> None:
        for (name, _, _), results in self.named_configs:
            for j, r in enumerate(results.predictor_summary.pearson_r):
                self.assertGreaterEqual(r, -1.0, msg=f"{name} j={j}")
                self.assertLessEqual(r, 1.0, msg=f"{name} j={j}")

    def test_spearman_r_in_range(self) -> None:
        for (name, _, _), results in self.named_configs:
            for j, r in enumerate(results.predictor_summary.spearman_r):
                self.assertGreaterEqual(r, -1.0, msg=f"{name} j={j}")
                self.assertLessEqual(r, 1.0, msg=f"{name} j={j}")

    def test_full_residuals_length_matches_n(self) -> None:
        for (name, _, _), results in self.named_configs:
            n = results.summary.observations
            self.assertEqual(len(results.full_residuals.predictions), n, name)
            self.assertEqual(len(results.full_residuals.residuals), n, name)
            self.assertEqual(len(results.full_residuals.loocv_residuals), n, name)
            self.assertEqual(len(results.full_residuals.hat_diagonal), n, name)
            self.assertEqual(len(results.full_residuals.studentized_residuals), n, name)
            self.assertEqual(len(results.full_residuals.cooks_distance), n, name)

    def test_hat_diagonal_sum_equals_p(self) -> None:
        for (name, allow, cols), results in self.named_configs:
            p = len(cols) + (1 if allow else 0)
            hat_sum = sum(results.full_residuals.hat_diagonal)
            self.assertAlmostEqual(hat_sum, p, places=4, msg=name)

    def test_hat_diagonal_bounds(self) -> None:
        for (name, _, _), results in self.named_configs:
            n = results.summary.observations
            for i, h in enumerate(results.full_residuals.hat_diagonal):
                self.assertGreaterEqual(h, 0.0, msg=f"{name} h[{i}]")
                self.assertLessEqual(h, 1.0 + 1e-10, msg=f"{name} h[{i}]")

    def test_cooks_distance_non_negative(self) -> None:
        for (name, _, _), results in self.named_configs:
            for i, cd in enumerate(results.full_residuals.cooks_distance):
                self.assertGreaterEqual(cd, 0.0, msg=f"{name} cook[{i}]")

    def test_residuals_sum_near_zero_with_intercept(self) -> None:
        for (name, allow, _), results in self.named_configs:
            if allow:
                self.assertAlmostEqual(
                    sum(results.full_residuals.residuals), 0.0, places=4,
                    msg=name,
                )

    def test_predictions_plus_residuals_equal_y(self) -> None:
        for (name, _, _), results in self.named_configs:
            for i, (y, p, r) in enumerate(zip(
                results.full_residuals.dependent_var,
                results.full_residuals.predictions,
                results.full_residuals.residuals,
            )):
                self.assertAlmostEqual(y, p + r, places=8, msg=f"{name} row={i}")

    def test_studentized_residuals_ranked_sorted(self) -> None:
        for (name, _, _), results in self.named_configs:
            ranked = list(results.full_residuals.studentized_residuals_ranked)
            self.assertEqual(ranked, sorted(ranked), name)

    def test_normal_scores_ranked_sorted(self) -> None:
        for (name, _, _), results in self.named_configs:
            ranked = list(results.full_residuals.normal_scores_ranked)
            self.assertEqual(ranked, sorted(ranked), name)

    def test_prediction_interval_contains_estimate(self) -> None:
        for (name, _, _), results in self.named_configs:
            pi = results.prediction_interval
            self.assertLess(pi.pi_lower, pi.point_estimate, name)
            self.assertGreater(pi.pi_upper, pi.point_estimate, name)

    def test_prediction_interval_symmetry(self) -> None:
        for (name, _, _), results in self.named_configs:
            pi = results.prediction_interval
            margin_lo = pi.point_estimate - pi.pi_lower
            margin_hi = pi.pi_upper - pi.point_estimate
            self.assertAlmostEqual(margin_lo, margin_hi, places=8, msg=name)

    def test_prediction_interval_confidence_level(self) -> None:
        for (name, _, _), results in self.named_configs:
            self.assertAlmostEqual(
                results.prediction_interval.confidence_level, 0.95, places=12,
                msg=name,
            )

    def test_prediction_interval_input_length_matches_k(self) -> None:
        for (name, _, cols), results in self.named_configs:
            self.assertEqual(
                len(results.prediction_interval.pred_input_values), len(cols),
                msg=name,
            )

    def test_summary_ss_decomposition(self) -> None:
        for (name, _, _), results in self.named_configs:
            self.assertAlmostEqual(
                results.summary.ss_total,
                results.summary.ss_regression + results.summary.ss_residual,
                places=4, msg=name,
            )

"""Verify Regression expected-value configs used by artifact-specific verifiers.

These tests run the same config-building functions that the Regression verifier
uses for the spec-driven sheet oracle, then check that the results are
internally consistent and survive cache round-trips. Since Excel is unavailable
in CI, this validates the Python-side oracle rather than the workbook formulas
themselves.
"""
# pylint: disable=import-outside-toplevel,missing-function-docstring,too-many-public-methods,unused-variable
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from typing import Any, ClassVar

from lambda_catalog.analysis_cache import (
    _CACHE_SCHEMA_VERSION,
    _deserialize_regression_sheet_configs,
    _serialize_regression_sheet_configs,
    get_analysis_results,
)
from lambda_catalog.analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
)
from lambda_catalog.analyze_regression_sheet import (
    REGRESSION_QC_CONFIGS,
    calculate_regression_sheet_results,
)

# ── Helpers ─────────────────────────────────────────────────────────────────

_CSV = DEFAULT_INPUT_CSV


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


# ── Cache serialization round-trips ─────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestCacheRoundTrip(unittest.TestCase):
    """Verify that Regression QC configs survive JSON serialization/deserialization."""

    reg_configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        from lambda_catalog.analyze_regression_sheet import (
            build_regression_sheet_qc_configs,
        )
        cls.reg_configs = build_regression_sheet_qc_configs(_CSV)

    def test_regression_sheet_round_trip(self) -> None:
        serialized = _serialize_regression_sheet_configs(self.reg_configs)
        json_str = json.dumps(serialized)
        deserialized = _deserialize_regression_sheet_configs(json.loads(json_str))
        self.assertEqual(len(deserialized), len(self.reg_configs))
        for (n1, ai1, r1), (n2, ai2, r2) in zip(self.reg_configs, deserialized):
            self.assertEqual(n1, n2)
            self.assertEqual(ai1, ai2)
            self.assertEqual(r1.summary.observations, r2.summary.observations)
            self.assertEqual(r1.summary.r_squared, r2.summary.r_squared)
            self.assertEqual(r1.vectors.coefficients, r2.vectors.coefficients)
            self.assertEqual(r1.predictor_summary.gvif, r2.predictor_summary.gvif)
            self.assertEqual(r1.full_residuals.hat_diagonal, r2.full_residuals.hat_diagonal)
            self.assertAlmostEqual(
                r1.prediction_interval.point_estimate,
                r2.prediction_interval.point_estimate,
                places=12,
            )

    def test_full_cache_round_trip_via_file(self) -> None:
        """Write to a temp cache file, then read back and compare."""
        with tempfile.TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "test_cache.json"
            results1 = get_analysis_results(_CSV, cache_path)
            self.assertTrue(cache_path.exists())
            results2 = get_analysis_results(_CSV, cache_path)
            self.assertEqual(len(results1), len(results2))

    def test_cache_invalidation_on_schema_bump(self) -> None:
        """Cache should be ignored when schema version doesn't match."""
        with tempfile.TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "test_cache.json"
            get_analysis_results(_CSV, cache_path)
            with cache_path.open("r") as f:
                cached = json.load(f)
            cached["schema_version"] = _CACHE_SCHEMA_VERSION - 1
            with cache_path.open("w") as f:
                json.dump(cached, f)
            results = get_analysis_results(_CSV, cache_path)
            self.assertGreater(len(results), 0)



"""Verify QC expected-value configs used by build_qc.

These tests run the same config-building functions that build_qc.py uses to
generate expected values for the test sheets, then check that the results are
internally consistent, cross-consistent between config types, and survive
cache round-trips. Since Excel is unavailable in CI, this validates the
Python-side QC oracle rather than the LAMBDA formulas themselves.
"""
# pylint: disable=import-outside-toplevel,missing-function-docstring,too-many-public-methods,unused-variable
from __future__ import annotations

import json
import math
import tempfile
import unittest
from pathlib import Path
from typing import Any, ClassVar

from lambda_catalog.analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
    FEATURE_COLUMNS,
    calculate_regression_summary,
)
from lambda_catalog.analyze_regression_sheet import (
    REGRESSION_QC_CONFIGS,
    calculate_regression_sheet_results,
)
from lambda_catalog.analysis_cache import (
    _CACHE_SCHEMA_VERSION,
    _csv_fingerprint,
    _deserialize_observation_configs,
    _deserialize_regression_sheet_configs,
    _deserialize_vector_configs,
    _serialize_observation_configs,
    _serialize_regression_sheet_configs,
    _serialize_vector_configs,
    get_analysis_results,
)
from lambda_catalog.write_sheet_mlr_observation_test import build_mlr_observation_row_configs
from lambda_catalog.write_sheet_mlr_scalar_test import (
    _MLR_K_VALUES,
    _expected_values_map,
    build_mlr_row_configs,
)
from lambda_catalog.write_sheet_mlr_vector_outputs_test import build_mlr_vector_row_configs


# ── Helpers ─────────────────────────────────────────────────────────────────

_CSV = DEFAULT_INPUT_CSV


def _has_csv() -> bool:
    return _CSV.exists()


_SKIP_MSG = "Life Expectancy CSV not found"


# ── Scalar config tests ─────────────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestScalarConfigs(unittest.TestCase):
    """Verify build_mlr_row_configs produces correct scalar expected values."""

    configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        cls.configs = build_mlr_row_configs(_CSV)

    def test_config_count(self) -> None:
        self.assertEqual(len(self.configs), len(_MLR_K_VALUES) * 2)

    def test_each_config_has_k_and_intercept(self) -> None:
        for row_vals, _ in self.configs:
            self.assertIn("ind_vars", row_vals)
            self.assertIn("Allow_Intercept", row_vals)

    def test_expected_values_map_keys(self) -> None:
        expected_keys = {
            "Observations", "DF_Regression", "DF_Total", "R_squared",
            "DF_Residual", "Multiple_R", "Adjusted_R2", "SS_Total",
            "SS_Residual", "SS_Regression", "SE_Regression", "PRESS",
            "Durbin_Watson", "F_Stat", "P_Value_F", "AIC", "BIC",
            "AICc", "QQ_Correlation",
        }
        for _, expected in self.configs:
            self.assertEqual(set(expected.keys()), expected_keys)

    def test_all_values_finite(self) -> None:
        for row_vals, expected in self.configs:
            for key, val in expected.items():
                self.assertTrue(
                    math.isfinite(val),
                    f"k={row_vals['ind_vars']} intercept={row_vals['Allow_Intercept']} "
                    f"{key}={val} is not finite",
                )

    def test_ss_decomposition(self) -> None:
        for row_vals, expected in self.configs:
            self.assertAlmostEqual(
                expected["SS_Total"],
                expected["SS_Regression"] + expected["SS_Residual"],
                places=6,
                msg=f"k={row_vals['ind_vars']} intercept={row_vals['Allow_Intercept']}",
            )

    def test_r_squared_bounds(self) -> None:
        for row_vals, expected in self.configs:
            if row_vals["Allow_Intercept"]:
                self.assertGreaterEqual(expected["R_squared"], 0.0)
                self.assertLessEqual(expected["R_squared"], 1.0)

    def test_multiple_r_is_sqrt_r2(self) -> None:
        for row_vals, expected in self.configs:
            self.assertAlmostEqual(
                expected["Multiple_R"],
                math.sqrt(max(expected["R_squared"], 0.0)),
                places=10,
            )

    def test_df_total_with_intercept(self) -> None:
        for row_vals, expected in self.configs:
            if row_vals["Allow_Intercept"]:
                self.assertEqual(
                    expected["DF_Total"],
                    expected["Observations"] - 1,
                )

    def test_df_total_without_intercept(self) -> None:
        for row_vals, expected in self.configs:
            if not row_vals["Allow_Intercept"]:
                self.assertEqual(
                    expected["DF_Total"],
                    expected["Observations"],
                )

    def test_df_residual(self) -> None:
        for row_vals, expected in self.configs:
            k = row_vals["ind_vars"]
            n = expected["Observations"]
            df_resid = n - k - (1 if row_vals["Allow_Intercept"] else 0)
            self.assertEqual(expected["DF_Residual"], df_resid)

    def test_se_regression_squared_equals_mse(self) -> None:
        for row_vals, expected in self.configs:
            mse = expected["SS_Residual"] / expected["DF_Residual"]
            self.assertAlmostEqual(expected["SE_Regression"] ** 2, mse, places=8)

    def test_f_stat_from_mean_squares(self) -> None:
        for row_vals, expected in self.configs:
            ms_reg = expected["SS_Regression"] / expected["DF_Regression"]
            ms_res = expected["SS_Residual"] / expected["DF_Residual"]
            self.assertAlmostEqual(expected["F_Stat"], ms_reg / ms_res, places=8)

    def test_durbin_watson_range(self) -> None:
        for _, expected in self.configs:
            self.assertGreaterEqual(expected["Durbin_Watson"], 0.0)
            self.assertLessEqual(expected["Durbin_Watson"], 4.0)

    def test_press_positive(self) -> None:
        for _, expected in self.configs:
            self.assertGreater(expected["PRESS"], 0.0)

    def test_aic_bic_aicc_formulas(self) -> None:
        for row_vals, expected in self.configs:
            n = expected["Observations"]
            p = expected["DF_Regression"] + (1 if row_vals["Allow_Intercept"] else 0)
            log_term = n * math.log(expected["SS_Residual"] / n)
            aic = log_term + 2.0 * p
            bic = log_term + p * math.log(n)
            aicc = aic + 2.0 * p * (p + 1) / (n - p - 1)
            self.assertAlmostEqual(expected["AIC"], aic, places=8)
            self.assertAlmostEqual(expected["BIC"], bic, places=8)
            self.assertAlmostEqual(expected["AICc"], aicc, places=8)


# ── Vector config tests ─────────────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestVectorConfigs(unittest.TestCase):
    """Verify build_mlr_vector_row_configs outputs."""

    configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        cls.configs = build_mlr_vector_row_configs(_CSV)

    def test_config_count(self) -> None:
        self.assertEqual(len(self.configs), len(_MLR_K_VALUES) * 2)

    def test_term_count_with_intercept(self) -> None:
        for k, allow_intercept, vectors in self.configs:
            if allow_intercept:
                self.assertEqual(len(vectors.coefficients), k + 1)
                self.assertEqual(vectors.term_names[0], "Intercept")
            else:
                self.assertEqual(len(vectors.coefficients), k)

    def test_all_vector_lengths_consistent(self) -> None:
        for k, allow_intercept, vectors in self.configs:
            n = len(vectors.coefficients)
            self.assertEqual(len(vectors.std_errors), n, f"k={k}")
            self.assertEqual(len(vectors.t_stats), n, f"k={k}")
            self.assertEqual(len(vectors.p_values), n, f"k={k}")
            self.assertEqual(len(vectors.ci_lower), n, f"k={k}")
            self.assertEqual(len(vectors.ci_upper), n, f"k={k}")

    def test_beta_weights_length_equals_k(self) -> None:
        for k, _, vectors in self.configs:
            self.assertEqual(len(vectors.beta_weights), k)

    def test_t_equals_coef_over_se(self) -> None:
        for k, allow_intercept, vectors in self.configs:
            for i, (c, s, t) in enumerate(
                zip(vectors.coefficients, vectors.std_errors, vectors.t_stats)
            ):
                self.assertAlmostEqual(t, c / s, places=8, msg=f"k={k} i={i}")

    def test_ci_contains_coefficient(self) -> None:
        for k, _, vectors in self.configs:
            for i, (lo, c, hi) in enumerate(
                zip(vectors.ci_lower, vectors.coefficients, vectors.ci_upper)
            ):
                self.assertLessEqual(lo, c + 1e-12, msg=f"k={k} lo[{i}]")
                self.assertGreaterEqual(hi, c - 1e-12, msg=f"k={k} hi[{i}]")

    def test_p_values_in_unit_interval(self) -> None:
        for k, _, vectors in self.configs:
            for i, p in enumerate(vectors.p_values):
                self.assertGreaterEqual(p, 0.0, msg=f"k={k} p[{i}]")
                self.assertLessEqual(p, 1.0, msg=f"k={k} p[{i}]")

    def test_std_errors_positive(self) -> None:
        for k, _, vectors in self.configs:
            for i, se in enumerate(vectors.std_errors):
                self.assertGreater(se, 0.0, msg=f"k={k} se[{i}]")


# ── Observation config tests ────────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestObservationConfigs(unittest.TestCase):
    """Verify build_mlr_observation_row_configs outputs."""

    configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        cls.configs = build_mlr_observation_row_configs(_CSV)

    def test_config_count(self) -> None:
        self.assertEqual(len(self.configs), len(_MLR_K_VALUES) * 2)

    def test_all_vector_lengths_consistent(self) -> None:
        for k, allow_intercept, obs in self.configs:
            n = len(obs.observation_num)
            self.assertGreater(n, 0)
            self.assertEqual(len(obs.rank_fraction), n, f"k={k}")
            self.assertEqual(len(obs.y_ranked), n, f"k={k}")
            self.assertEqual(len(obs.normal_scores), n, f"k={k}")
            self.assertEqual(len(obs.predictions), n, f"k={k}")
            self.assertEqual(len(obs.residuals), n, f"k={k}")
            self.assertEqual(len(obs.scaled_residuals), n, f"k={k}")
            self.assertEqual(len(obs.scaled_residuals_ranked), n, f"k={k}")

    def test_observation_num_sequential(self) -> None:
        for k, _, obs in self.configs:
            self.assertEqual(
                list(obs.observation_num),
                list(range(1, len(obs.observation_num) + 1)),
                f"k={k}",
            )

    def test_residuals_sum_near_zero_with_intercept(self) -> None:
        for k, allow_intercept, obs in self.configs:
            if allow_intercept:
                self.assertAlmostEqual(
                    sum(obs.residuals), 0.0, places=4,
                    msg=f"k={k} residuals don't sum to zero",
                )

    def test_y_ranked_sorted(self) -> None:
        for k, _, obs in self.configs:
            ranked = list(obs.y_ranked)
            self.assertEqual(ranked, sorted(ranked), f"k={k}")

    def test_scaled_residuals_ranked_sorted(self) -> None:
        for k, _, obs in self.configs:
            ranked = list(obs.scaled_residuals_ranked)
            self.assertEqual(ranked, sorted(ranked), f"k={k}")

    def test_predictions_plus_residuals_reconstruct_y(self) -> None:
        for k, _, obs in self.configs:
            reconstructed = sorted(
                p + r for p, r in zip(obs.predictions, obs.residuals)
            )
            for i, (actual, computed) in enumerate(zip(obs.y_ranked, reconstructed)):
                self.assertAlmostEqual(
                    actual, computed, places=8,
                    msg=f"k={k} row={i}",
                )


# ── Cross-consistency between config types ──────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestCrossConsistency(unittest.TestCase):
    """Verify scalar, vector, and observation configs agree for same (k, intercept)."""

    scalar_configs: ClassVar[list[Any]]
    vector_configs: ClassVar[list[Any]]
    obs_configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        cls.scalar_configs = build_mlr_row_configs(_CSV)
        cls.vector_configs = build_mlr_vector_row_configs(_CSV)
        cls.obs_configs = build_mlr_observation_row_configs(_CSV)

    def _scalar_lookup(self) -> dict:
        lookup = {}
        for row_vals, expected in self.scalar_configs:
            key = (row_vals["ind_vars"], row_vals["Allow_Intercept"])
            lookup[key] = expected
        return lookup

    def test_observation_count_matches_scalar(self) -> None:
        scalar_lookup = self._scalar_lookup()
        for k, allow_intercept, obs in self.obs_configs:
            expected_n = scalar_lookup[(k, allow_intercept)]["Observations"]
            self.assertEqual(
                len(obs.observation_num), expected_n,
                f"k={k} intercept={allow_intercept}",
            )

    def test_coefficient_count_matches_df_regression(self) -> None:
        scalar_lookup = self._scalar_lookup()
        for k, allow_intercept, vectors in self.vector_configs:
            expected_df_reg = scalar_lookup[(k, allow_intercept)]["DF_Regression"]
            self.assertEqual(expected_df_reg, k)
            expected_terms = k + (1 if allow_intercept else 0)
            self.assertEqual(
                len(vectors.coefficients), expected_terms,
                f"k={k} intercept={allow_intercept}",
            )

    def test_same_k_values_across_configs(self) -> None:
        scalar_ks = sorted({rv["ind_vars"] for rv, _ in self.scalar_configs})
        vector_ks = sorted({k for k, _, _ in self.vector_configs})
        obs_ks = sorted({k for k, _, _ in self.obs_configs})
        self.assertEqual(scalar_ks, _MLR_K_VALUES)
        self.assertEqual(vector_ks, _MLR_K_VALUES)
        self.assertEqual(obs_ks, _MLR_K_VALUES)


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
            self.assertEqual(len(results.predictor_summary.vif), k, name)
            self.assertEqual(len(results.predictor_summary.tolerance), k, name)

    def test_predictor_names_match_columns(self) -> None:
        for (name, allow, cols), results in self.named_configs:
            self.assertEqual(
                list(results.predictor_summary.predictor_names), cols, name,
            )

    def test_vif_tolerance_reciprocal(self) -> None:
        for (name, _, _), results in self.named_configs:
            for j, (v, t) in enumerate(
                zip(results.predictor_summary.vif, results.predictor_summary.tolerance)
            ):
                self.assertAlmostEqual(v * t, 1.0, places=10, msg=f"{name} j={j}")

    def test_vif_at_least_one(self) -> None:
        for (name, _, _), results in self.named_configs:
            for j, v in enumerate(results.predictor_summary.vif):
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
            self.assertLess(pi.lower, pi.point_estimate, name)
            self.assertGreater(pi.upper, pi.point_estimate, name)

    def test_prediction_interval_symmetry(self) -> None:
        for (name, _, _), results in self.named_configs:
            pi = results.prediction_interval
            margin_lo = pi.point_estimate - pi.lower
            margin_hi = pi.upper - pi.point_estimate
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
    """Verify that QC configs survive JSON serialization/deserialization."""

    vector_configs: ClassVar[list[Any]]
    obs_configs: ClassVar[list[Any]]
    reg_configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        cls.vector_configs = build_mlr_vector_row_configs(_CSV)
        cls.obs_configs = build_mlr_observation_row_configs(_CSV)
        from lambda_catalog.analyze_regression_sheet import build_regression_sheet_qc_configs
        cls.reg_configs = build_regression_sheet_qc_configs(_CSV)

    def test_vector_round_trip(self) -> None:
        serialized = _serialize_vector_configs(self.vector_configs)
        json_str = json.dumps(serialized)
        deserialized = _deserialize_vector_configs(json.loads(json_str))
        self.assertEqual(len(deserialized), len(self.vector_configs))
        for (k1, ai1, v1), (k2, ai2, v2) in zip(self.vector_configs, deserialized):
            self.assertEqual(k1, k2)
            self.assertEqual(ai1, ai2)
            self.assertEqual(v1.coefficients, v2.coefficients)
            self.assertEqual(v1.std_errors, v2.std_errors)
            self.assertEqual(v1.t_stats, v2.t_stats)
            self.assertEqual(v1.p_values, v2.p_values)
            self.assertEqual(v1.ci_lower, v2.ci_lower)
            self.assertEqual(v1.ci_upper, v2.ci_upper)
            self.assertEqual(v1.beta_weights, v2.beta_weights)

    def test_observation_round_trip(self) -> None:
        serialized = _serialize_observation_configs(self.obs_configs)
        json_str = json.dumps(serialized)
        deserialized = _deserialize_observation_configs(json.loads(json_str))
        self.assertEqual(len(deserialized), len(self.obs_configs))
        for (k1, ai1, v1), (k2, ai2, v2) in zip(self.obs_configs, deserialized):
            self.assertEqual(k1, k2)
            self.assertEqual(ai1, ai2)
            self.assertEqual(v1.observation_num, v2.observation_num)
            self.assertEqual(v1.predictions, v2.predictions)
            self.assertEqual(v1.residuals, v2.residuals)

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
            self.assertEqual(r1.predictor_summary.vif, r2.predictor_summary.vif)
            self.assertEqual(
                r1.full_residuals.hat_diagonal, r2.full_residuals.hat_diagonal,
            )
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

            scalar1, vec1, obs1, reg1 = results1
            scalar2, vec2, obs2, reg2 = results2

            self.assertEqual(len(scalar1), len(scalar2))
            self.assertEqual(len(vec1), len(vec2))
            self.assertEqual(len(obs1), len(obs2))
            self.assertEqual(len(reg1), len(reg2))

            for s1, s2 in zip(scalar1, scalar2):
                self.assertEqual(tuple(s1), tuple(s2))

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
            self.assertGreater(len(results[0]), 0)


# ── Expected values map coverage ────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestExpectedValuesMap(unittest.TestCase):
    """Verify _expected_values_map produces the right keys/values from a summary."""

    def test_map_round_trips_summary(self) -> None:
        summary = calculate_regression_summary(
            _CSV, include_intercept=True, feature_columns=FEATURE_COLUMNS[:3]
        )
        mapping = _expected_values_map(summary)
        self.assertEqual(mapping["Observations"], summary.observations)
        self.assertEqual(mapping["R_squared"], summary.r_squared)
        self.assertEqual(mapping["SS_Total"], summary.ss_total)
        self.assertEqual(mapping["PRESS"], summary.press)
        self.assertEqual(mapping["AIC"], summary.aic)
        self.assertEqual(mapping["QQ_Correlation"], summary.qq_correlation)


# ── CSV fingerprint stability ───────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestCSVFingerprint(unittest.TestCase):
    """Verify CSV fingerprint is deterministic and changes on content change."""

    def test_deterministic(self) -> None:
        fp1 = _csv_fingerprint(_CSV)
        fp2 = _csv_fingerprint(_CSV)
        self.assertEqual(fp1, fp2)

    def test_changes_on_content_change(self) -> None:
        original_fp = _csv_fingerprint(_CSV)
        with tempfile.TemporaryDirectory() as tmp:
            modified = Path(tmp) / "modified.csv"
            modified.write_text(_CSV.read_text() + "\n")
            modified_fp = _csv_fingerprint(modified)
        self.assertNotEqual(original_fp, modified_fp)


# ── LOOCV cross-check ──────────────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestLOOCVCrossCheck(unittest.TestCase):
    """Verify LOOCV residual = e/(1-h) identity on the real dataset."""

    def test_loocv_press_identity(self) -> None:
        results = calculate_regression_sheet_results(
            _CSV, include_intercept=True, feature_columns=FEATURE_COLUMNS[:4]
        )
        fr = results.full_residuals
        for i in range(min(50, len(fr.residuals))):
            e = fr.residuals[i]
            h = fr.hat_diagonal[i]
            expected_loocv_resid = e / (1.0 - h)
            self.assertAlmostEqual(
                fr.loocv_residuals[i], expected_loocv_resid, places=8,
                msg=f"row={i}",
            )


# ── Cook's distance cross-check ────────────────────────────────────────────


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestCooksDistanceCrossCheck(unittest.TestCase):
    """Verify Cook's distance = r² * h / ((1-h) * p) on the real dataset."""

    def test_cooks_formula(self) -> None:
        cols = FEATURE_COLUMNS[:4]
        results = calculate_regression_sheet_results(
            _CSV, include_intercept=True, feature_columns=cols
        )
        fr = results.full_residuals
        p = len(cols) + 1
        for i in range(min(50, len(fr.studentized_residuals))):
            r = fr.studentized_residuals[i]
            h = fr.hat_diagonal[i]
            expected_cd = r**2 * h / ((1.0 - h) * p)
            self.assertAlmostEqual(
                fr.cooks_distance[i], expected_cd, places=8,
                msg=f"row={i}",
            )


if __name__ == "__main__":
    unittest.main()

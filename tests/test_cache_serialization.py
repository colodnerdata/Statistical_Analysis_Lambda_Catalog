"""Tests for analysis_cache serialization round-trips (pure Python, no xlwings needed)."""
from __future__ import annotations

import unittest

from lambda_catalog.regression_shared import (
    RegressionObservationVectors,
    RegressionVectors,
)
from lambda_catalog.analysis_cache import (
    _deserialize_observation_configs,
    _deserialize_vector_configs,
    _serialize_observation_configs,
    _serialize_vector_configs,
)


def _make_vectors(n_terms: int = 3) -> RegressionVectors:
    vals = tuple(float(i) for i in range(n_terms))
    return RegressionVectors(
        term_names=tuple(f"term_{i}" for i in range(n_terms)),
        coefficients=vals,
        std_errors=tuple(float(i) + 0.1 for i in range(n_terms)),
        t_stats=tuple(float(i) + 0.2 for i in range(n_terms)),
        p_values=tuple(0.05 for _ in range(n_terms)),
        ci_lower=tuple(float(i) - 1.0 for i in range(n_terms)),
        ci_upper=tuple(float(i) + 1.0 for i in range(n_terms)),
        beta_weights=tuple(float(i) * 0.1 for i in range(n_terms - 1)),
    )


def _make_observation_vectors(n: int = 4) -> RegressionObservationVectors:
    vals = tuple(float(i) for i in range(n))
    return RegressionObservationVectors(
        observation_num=tuple(range(1, n + 1)),
        rank_fraction=tuple(float(i + 1) / n for i in range(n)),
        y_ranked=vals,
        normal_scores=vals,
        predictions=vals,
        residuals=vals,
        scaled_residuals=vals,
        scaled_residuals_ranked=tuple(sorted(vals)),
    )


class VectorConfigRoundTripTests(unittest.TestCase):
    def _roundtrip(self, configs):
        return _deserialize_vector_configs(_serialize_vector_configs(configs))

    def test_term_names_preserved(self) -> None:
        vectors = _make_vectors(3)
        restored = self._roundtrip([(3, True, vectors)])[0][2]
        self.assertEqual(restored.term_names, vectors.term_names)

    def test_coefficients_preserved(self) -> None:
        vectors = _make_vectors(3)
        restored = self._roundtrip([(3, True, vectors)])[0][2]
        for orig, rt in zip(vectors.coefficients, restored.coefficients):
            self.assertAlmostEqual(orig, rt)

    def test_ci_bounds_preserved(self) -> None:
        vectors = _make_vectors(4)
        restored = self._roundtrip([(4, False, vectors)])[0][2]
        for orig, rt in zip(vectors.ci_lower, restored.ci_lower):
            self.assertAlmostEqual(orig, rt)
        for orig, rt in zip(vectors.ci_upper, restored.ci_upper):
            self.assertAlmostEqual(orig, rt)

    def test_k_and_intercept_flag_preserved(self) -> None:
        vectors = _make_vectors(2)
        k, allow_intercept, _ = self._roundtrip([(5, False, vectors)])[0]
        self.assertEqual(k, 5)
        self.assertFalse(allow_intercept)

    def test_empty_list_round_trips(self) -> None:
        self.assertEqual(self._roundtrip([]), [])

    def test_multiple_configs_preserved(self) -> None:
        configs = [(i, True, _make_vectors(i + 1)) for i in range(1, 4)]
        restored = self._roundtrip(configs)
        self.assertEqual(len(restored), 3)
        for (k_orig, flag_orig, v_orig), (k_rt, flag_rt, v_rt) in zip(configs, restored):
            self.assertEqual(k_orig, k_rt)
            self.assertEqual(flag_orig, flag_rt)
            self.assertEqual(v_orig.term_names, v_rt.term_names)


class ObservationConfigRoundTripTests(unittest.TestCase):
    def _roundtrip(self, configs):
        return _deserialize_observation_configs(_serialize_observation_configs(configs))

    def test_observation_num_preserved(self) -> None:
        obs = _make_observation_vectors(4)
        restored = self._roundtrip([(2, True, obs)])[0][2]
        self.assertEqual(restored.observation_num, obs.observation_num)

    def test_residuals_preserved(self) -> None:
        obs = _make_observation_vectors(5)
        restored = self._roundtrip([(2, True, obs)])[0][2]
        for orig, rt in zip(obs.residuals, restored.residuals):
            self.assertAlmostEqual(orig, rt)

    def test_rank_fraction_preserved(self) -> None:
        obs = _make_observation_vectors(6)
        restored = self._roundtrip([(3, False, obs)])[0][2]
        for orig, rt in zip(obs.rank_fraction, restored.rank_fraction):
            self.assertAlmostEqual(orig, rt)

    def test_k_and_intercept_flag_preserved(self) -> None:
        obs = _make_observation_vectors(3)
        k, allow_intercept, _ = self._roundtrip([(7, False, obs)])[0]
        self.assertEqual(k, 7)
        self.assertFalse(allow_intercept)

    def test_empty_list_round_trips(self) -> None:
        self.assertEqual(self._roundtrip([]), [])


if __name__ == "__main__":
    unittest.main()

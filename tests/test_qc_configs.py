"""Verify Python-side QC config/cache helpers for spec-driven Regression checks."""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from typing import Any, ClassVar

from lambda_catalog.analyze_life_expectancy import DEFAULT_INPUT_CSV
from lambda_catalog.analyze_regression_sheet import calculate_regression_sheet_results
from lambda_catalog.analysis_cache import (
    _CACHE_SCHEMA_VERSION,
    _csv_fingerprint,
    _deserialize_regression_sheet_configs,
    _serialize_regression_sheet_configs,
    get_analysis_results,
)

_CSV = DEFAULT_INPUT_CSV
_SKIP_MSG = "Life Expectancy CSV not found"


def _has_csv() -> bool:
    return _CSV.exists()


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestRegressionSheetConfigs(unittest.TestCase):
    """Verify the active Regression QC oracle produces internally consistent results."""

    named_configs: ClassVar[list[tuple[tuple[str, bool, tuple[str, ...]], Any]]]

    @classmethod
    def setUpClass(cls) -> None:
        from lambda_catalog.analyze_regression_sheet import REGRESSION_QC_CONFIGS

        cls.named_configs = [
            (config, calculate_regression_sheet_results(_CSV, include_intercept=allow, feature_columns=cols))
            for config in REGRESSION_QC_CONFIGS
            for name, allow, cols in [config]
        ]

    def test_configs_exist(self) -> None:
        self.assertGreater(len(self.named_configs), 0)

    def test_prediction_interval_confidence_level(self) -> None:
        for (name, _, _), results in self.named_configs:
            self.assertAlmostEqual(results.prediction_interval.confidence_level, 0.95, places=12, msg=name)

    def test_prediction_interval_input_length_matches_k(self) -> None:
        for (name, _, cols), results in self.named_configs:
            self.assertEqual(len(results.prediction_interval.pred_input_values), len(cols), msg=name)

    def test_summary_ss_decomposition(self) -> None:
        for (name, _, _), results in self.named_configs:
            self.assertAlmostEqual(
                results.summary.ss_total,
                results.summary.ss_regression + results.summary.ss_residual,
                places=4,
                msg=name,
            )


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestCacheRoundTrip(unittest.TestCase):
    """Verify that active Regression QC configs survive JSON/cache round-trips."""

    reg_configs: ClassVar[list[Any]]

    @classmethod
    def setUpClass(cls) -> None:
        from lambda_catalog.analyze_regression_sheet import build_regression_sheet_qc_configs

        cls.reg_configs = build_regression_sheet_qc_configs(_CSV)

    def test_regression_sheet_round_trip(self) -> None:
        serialized = _serialize_regression_sheet_configs(self.reg_configs)
        deserialized = _deserialize_regression_sheet_configs(json.loads(json.dumps(serialized)))
        self.assertEqual(len(deserialized), len(self.reg_configs))
        for (n1, ai1, r1), (n2, ai2, r2) in zip(self.reg_configs, deserialized):
            self.assertEqual(n1, n2)
            self.assertEqual(ai1, ai2)
            self.assertEqual(r1.summary.observations, r2.summary.observations)
            self.assertEqual(r1.summary.r_squared, r2.summary.r_squared)
            self.assertEqual(r1.vectors.coefficients, r2.vectors.coefficients)
            self.assertEqual(r1.predictor_summary.gvif, r2.predictor_summary.gvif)
            self.assertEqual(r1.full_residuals.hat_diagonal, r2.full_residuals.hat_diagonal)
            self.assertAlmostEqual(r1.prediction_interval.point_estimate, r2.prediction_interval.point_estimate, places=12)

    def test_full_cache_round_trip_via_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "test_cache.json"
            results1 = get_analysis_results(_CSV, cache_path)
            self.assertTrue(cache_path.exists())
            results2 = get_analysis_results(_CSV, cache_path)
            self.assertEqual(len(results1), len(results2))
            self.assertGreater(len(results1), 0)

    def test_cache_invalidation_on_schema_bump(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "test_cache.json"
            get_analysis_results(_CSV, cache_path)
            cached = json.loads(cache_path.read_text())
            cached["schema_version"] = _CACHE_SCHEMA_VERSION - 1
            cache_path.write_text(json.dumps(cached))
            self.assertGreater(len(get_analysis_results(_CSV, cache_path)), 0)


@unittest.skipUnless(_has_csv(), _SKIP_MSG)
class TestCSVFingerprint(unittest.TestCase):
    """Verify CSV fingerprint is deterministic and changes on content change."""

    def test_deterministic(self) -> None:
        self.assertEqual(_csv_fingerprint(_CSV), _csv_fingerprint(_CSV))

    def test_changes_on_content_change(self) -> None:
        original_fp = _csv_fingerprint(_CSV)
        with tempfile.TemporaryDirectory() as tmp:
            modified = Path(tmp) / "copy.csv"
            modified.write_bytes(_CSV.read_bytes() + b"\n")
            self.assertNotEqual(original_fp, _csv_fingerprint(modified))

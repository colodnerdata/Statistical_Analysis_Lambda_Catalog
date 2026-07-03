"""Disk cache for OLS analysis results, keyed on CSV SHA-256 hash and schema version."""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .regression_shared import (
    RegressionFullResiduals,
    RegressionObservationVectors,
    RegressionPredictionInterval,
    RegressionPredictorSummary,
    RegressionSheetResults,
    RegressionSummary,
    RegressionVectors,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CACHE_PATH = ROOT_DIR / ".analysis_cache.json"
_CACHE_SCHEMA_VERSION = 12


def _csv_fingerprint(csv_path: Path) -> str:
    sha = hashlib.sha256()
    with csv_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            sha.update(chunk)
    return sha.hexdigest()


def _serialize_vector_configs(
    configs: list[tuple[int, bool, RegressionVectors]],
) -> list[dict[str, Any]]:
    result = []
    for k, allow_intercept, vectors in configs:
        result.append({
            "k": k,
            "allow_intercept": allow_intercept,
            "term_names": list(vectors.term_names),
            "coefficients": list(vectors.coefficients),
            "std_errors": list(vectors.std_errors),
            "t_stats": list(vectors.t_stats),
            "p_values": list(vectors.p_values),
            "ci_lower": list(vectors.ci_lower),
            "ci_upper": list(vectors.ci_upper),
            "beta_weights": list(vectors.beta_weights),
        })
    return result


def _deserialize_vector_configs(
    data: list[dict[str, Any]],
) -> list[tuple[int, bool, RegressionVectors]]:
    result = []
    for item in data:
        vectors = RegressionVectors(
            term_names=tuple(item["term_names"]),
            coefficients=tuple(item["coefficients"]),
            std_errors=tuple(item["std_errors"]),
            t_stats=tuple(item["t_stats"]),
            p_values=tuple(item["p_values"]),
            ci_lower=tuple(item["ci_lower"]),
            ci_upper=tuple(item["ci_upper"]),
            beta_weights=tuple(item["beta_weights"]),
        )
        result.append((item["k"], item["allow_intercept"], vectors))
    return result


def _serialize_observation_configs(
    configs: list[tuple[int, bool, RegressionObservationVectors]],
) -> list[dict[str, Any]]:
    return [{"k": k, "allow_intercept": allow_intercept, **vectors.__dict__} for k, allow_intercept, vectors in configs]


def _deserialize_observation_configs(
    data: list[dict[str, Any]],
) -> list[tuple[int, bool, RegressionObservationVectors]]:
    result = []
    for item in data:
        vector_fields = {
            k: tuple(v)
            for k, v in item.items()
            if k not in {"k", "allow_intercept"}
        }
        vectors = RegressionObservationVectors(**vector_fields)
        result.append((item["k"], item["allow_intercept"], vectors))
    return result


def _serialize_regression_sheet_configs(
    configs: list[tuple[str, bool, RegressionSheetResults]],
) -> list[dict[str, Any]]:
    result = []
    for name, allow_intercept, r in configs:
        result.append({
            "name": name,
            "allow_intercept": allow_intercept,
            "summary": asdict(r.summary),
            "vectors": {
                "term_names": list(r.vectors.term_names),
                "coefficients": list(r.vectors.coefficients),
                "std_errors": list(r.vectors.std_errors),
                "t_stats": list(r.vectors.t_stats),
                "p_values": list(r.vectors.p_values),
                "ci_lower": list(r.vectors.ci_lower),
                "ci_upper": list(r.vectors.ci_upper),
                "beta_weights": list(r.vectors.beta_weights),
            },
            "predictor_summary": {
                "predictor_names": list(r.predictor_summary.predictor_names),
                "pearson_r": list(r.predictor_summary.pearson_r),
                "spearman_r": list(r.predictor_summary.spearman_r),
                "skewness": list(r.predictor_summary.skewness),
                "kurtosis": list(r.predictor_summary.kurtosis),
                "vif": list(r.predictor_summary.vif),
                "tolerance": list(r.predictor_summary.tolerance),
            },
            "full_residuals": {
                "dependent_var": list(r.full_residuals.dependent_var),
                "predictions": list(r.full_residuals.predictions),
                "residuals": list(r.full_residuals.residuals),
                "loocv_residuals": list(r.full_residuals.loocv_residuals),
                "hat_diagonal": list(r.full_residuals.hat_diagonal),
                "studentized_residuals": list(r.full_residuals.studentized_residuals),
                "cooks_distance": list(r.full_residuals.cooks_distance),
                "normal_scores_ranked": list(r.full_residuals.normal_scores_ranked),
                "studentized_residuals_ranked": list(r.full_residuals.studentized_residuals_ranked),
            },
            "prediction_interval": {
                "pred_input_values": list(r.prediction_interval.pred_input_values),
                "point_estimate": r.prediction_interval.point_estimate,
                "se_prediction": r.prediction_interval.se_prediction,
                "t_critical": r.prediction_interval.t_critical,
                "lower": r.prediction_interval.lower,
                "upper": r.prediction_interval.upper,
                "confidence_level": r.prediction_interval.confidence_level,
            },
        })
    return result


def _deserialize_regression_sheet_configs(
    data: list[dict[str, Any]],
) -> list[tuple[str, bool, RegressionSheetResults]]:
    result = []
    for item in data:
        s = item["summary"]
        summary = RegressionSummary(**s)
        v = item["vectors"]
        vectors = RegressionVectors(
            term_names=tuple(v["term_names"]),
            coefficients=tuple(v["coefficients"]),
            std_errors=tuple(v["std_errors"]),
            t_stats=tuple(v["t_stats"]),
            p_values=tuple(v["p_values"]),
            ci_lower=tuple(v["ci_lower"]),
            ci_upper=tuple(v["ci_upper"]),
            beta_weights=tuple(v["beta_weights"]),
        )
        ps = item["predictor_summary"]
        predictor_summary = RegressionPredictorSummary(
            predictor_names=tuple(ps["predictor_names"]),
            pearson_r=tuple(ps["pearson_r"]),
            spearman_r=tuple(ps["spearman_r"]),
            skewness=tuple(ps["skewness"]),
            kurtosis=tuple(ps["kurtosis"]),
            vif=tuple(ps["vif"]),
            tolerance=tuple(ps["tolerance"]),
        )
        fr = item["full_residuals"]
        full_residuals = RegressionFullResiduals(
            dependent_var=tuple(fr["dependent_var"]),
            predictions=tuple(fr["predictions"]),
            residuals=tuple(fr["residuals"]),
            loocv_residuals=tuple(fr["loocv_residuals"]),
            hat_diagonal=tuple(fr["hat_diagonal"]),
            studentized_residuals=tuple(fr["studentized_residuals"]),
            cooks_distance=tuple(fr["cooks_distance"]),
            normal_scores_ranked=tuple(fr["normal_scores_ranked"]),
            studentized_residuals_ranked=tuple(fr["studentized_residuals_ranked"]),
        )
        pi = item["prediction_interval"]
        prediction_interval = RegressionPredictionInterval(
            pred_input_values=tuple(pi["pred_input_values"]),
            point_estimate=pi["point_estimate"],
            se_prediction=pi["se_prediction"],
            t_critical=pi["t_critical"],
            lower=pi["lower"],
            upper=pi["upper"],
            confidence_level=pi["confidence_level"],
        )
        result.append((
            item["name"],
            item["allow_intercept"],
            RegressionSheetResults(
                summary=summary,
                vectors=vectors,
                predictor_summary=predictor_summary,
                full_residuals=full_residuals,
                prediction_interval=prediction_interval,
            ),
        ))
    return result


def get_analysis_results(
    csv_path: Path,
    cache_path: Path = DEFAULT_CACHE_PATH,
) -> tuple[
    list,
    list[tuple[int, bool, RegressionVectors]],
    list[tuple[int, bool, RegressionObservationVectors]],
    list[tuple[str, bool, RegressionSheetResults]],
]:
    """Return (scalar, vector, observation, regression_sheet configs), from cache or fresh.

    The cache is invalidated when the CSV content changes (SHA-256 hash).
    Delete .analysis_cache.json manually after code or schema changes.
    """
    csv_path = csv_path.resolve()
    fingerprint = _csv_fingerprint(csv_path)

    if cache_path.exists():
        try:
            with cache_path.open("r", encoding="utf-8") as handle:
                cached = json.load(handle)
            if (
                cached.get("schema_version") == _CACHE_SCHEMA_VERSION
                and cached.get("csv_fingerprint") == fingerprint
            ):
                scalar_configs = [tuple(item) for item in cached["scalar_row_configs"]]
                vector_configs = _deserialize_vector_configs(cached["vector_row_configs"])
                observation_configs = _deserialize_observation_configs(cached["observation_row_configs"])
                regression_sheet_configs = _deserialize_regression_sheet_configs(cached["regression_sheet_configs"])
                return scalar_configs, vector_configs, observation_configs, regression_sheet_configs
        except (json.JSONDecodeError, KeyError, TypeError, ValueError, OSError):
            pass

    from .write_sheet_mlr_scalar_test import build_mlr_row_configs
    from .write_sheet_mlr_observation_test import build_mlr_observation_row_configs
    from .write_sheet_mlr_vector_outputs_test import build_mlr_vector_row_configs
    from .analyze_regression_sheet import build_regression_sheet_qc_configs

    scalar_configs = build_mlr_row_configs(csv_path)
    vector_configs = build_mlr_vector_row_configs(csv_path)
    observation_configs = build_mlr_observation_row_configs(csv_path)

    regression_sheet_configs = build_regression_sheet_qc_configs(csv_path)

    try:
        payload = {
            "schema_version": _CACHE_SCHEMA_VERSION,
            "csv_fingerprint": fingerprint,
            "scalar_row_configs": [list(item) for item in scalar_configs],
            "vector_row_configs": _serialize_vector_configs(vector_configs),
            "observation_row_configs": _serialize_observation_configs(observation_configs),
            "regression_sheet_configs": _serialize_regression_sheet_configs(regression_sheet_configs),
        }
        with cache_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
    except (OSError, TypeError):
        pass

    return scalar_configs, vector_configs, observation_configs, regression_sheet_configs

"""Disk cache for OLS analysis results, keyed on CSV SHA-256 hash and schema version."""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .regression_shared import (
    RegressionFullResiduals,
    RegressionPredictionInterval,
    RegressionPredictorSummary,
    RegressionSheetResults,
    RegressionSummary,
    RegressionUnitSpace,
    RegressionVectors,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CACHE_PATH = ROOT_DIR / ".analysis_cache.json"
# Bump on any change to the analysis oracle so a stale .analysis_cache.json is
# invalidated. v13: Regression sheet oracle switched to spec-driven completeness
# (filter on the model's own predictors, not all FEATURE_COLUMNS) when Full_Data
# was demoted from the default Filter — sparse/medium configs now compute on
# more rows, so cached v12 expected values are wrong.
# v14: RegressionFullResiduals gained scale_location (sqrt(|studentized|)), so
# a cached v13 entry has no key for it.
# v15: RegressionPredictorSummary.vif renamed to .gvif (Generalized VIF, one
# shared value per source variable) — the JSON key changed, so a cached v14
# entry has no "gvif" key.
# v16: RegressionPredictionInterval rebuilt from the pre-v2.1 6-value shape
# (se_prediction/lower/upper) to the v2.1 group-mean-recovery 9-value shape
# (se_mean/se_new/ci_lower/ci_upper/pi_lower/pi_upper/group_mean/group_count)
# — a cached v15 entry has none of the new keys.
# v17: RegressionSheetResults gained unit_space (v3.3 unit-space / back-
# transformation outputs) — a cached v16 entry has no "unit_space" key.
# v19: RegressionSummary gained bfn_panel_durbin_watson (the panel
# Durbin-Watson, the diagnostic that carries the number whenever Fixed
# Effects make the plain DW cell read "n/a — FE active"). The field has no
# default, so a cached v18 entry raises TypeError in RegressionSummary(**s)
# — caught by get_analysis_results' own except and recomputed, which is the
# intended behaviour. The bump makes it a version check rather than an
# exception.
# v20: retired the legacy MLR scalar/vector/observation QC sheets and stopped
# caching their row configs; cache payloads now hold only spec-driven
# Regression sheet configs.
_CACHE_SCHEMA_VERSION = 20


def _csv_fingerprint(csv_path: Path) -> str:
    sha = hashlib.sha256()
    with csv_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            sha.update(chunk)
    return sha.hexdigest()


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
                "gvif": list(r.predictor_summary.gvif),
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
                "scale_location": list(r.full_residuals.scale_location),
            },
            "prediction_interval": {
                "pred_input_values": list(r.prediction_interval.pred_input_values),
                "point_estimate": r.prediction_interval.point_estimate,
                "se_mean": r.prediction_interval.se_mean,
                "se_new": r.prediction_interval.se_new,
                "t_critical": r.prediction_interval.t_critical,
                "ci_lower": r.prediction_interval.ci_lower,
                "ci_upper": r.prediction_interval.ci_upper,
                "pi_lower": r.prediction_interval.pi_lower,
                "pi_upper": r.prediction_interval.pi_upper,
                "confidence_level": r.prediction_interval.confidence_level,
                "group_mean": r.prediction_interval.group_mean,
                "group_count": r.prediction_interval.group_count,
            },
            "unit_space": {
                "smearing_factor": r.unit_space.smearing_factor,
                "r_squared_unit": r.unit_space.r_squared_unit,
                "adjusted_r2_unit": r.unit_space.adjusted_r2_unit,
                "rmse_unit": r.unit_space.rmse_unit,
                "prediction_point_unit": r.unit_space.prediction_point_unit,
                "prediction_ci_lower_unit": r.unit_space.prediction_ci_lower_unit,
                "prediction_ci_upper_unit": r.unit_space.prediction_ci_upper_unit,
                "prediction_pi_lower_unit": r.unit_space.prediction_pi_lower_unit,
                "prediction_pi_upper_unit": r.unit_space.prediction_pi_upper_unit,
                "predictions_unit": list(r.unit_space.predictions_unit),
                "residuals_unit": list(r.unit_space.residuals_unit),
                "model_formula": r.unit_space.model_formula,
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
            gvif=tuple(ps["gvif"]),
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
            scale_location=tuple(fr["scale_location"]),
        )
        pi = item["prediction_interval"]
        prediction_interval = RegressionPredictionInterval(
            pred_input_values=tuple(pi["pred_input_values"]),
            point_estimate=pi["point_estimate"],
            se_mean=pi["se_mean"],
            se_new=pi["se_new"],
            t_critical=pi["t_critical"],
            ci_lower=pi["ci_lower"],
            ci_upper=pi["ci_upper"],
            pi_lower=pi["pi_lower"],
            pi_upper=pi["pi_upper"],
            confidence_level=pi["confidence_level"],
            group_mean=pi["group_mean"],
            group_count=pi["group_count"],
        )
        us = item["unit_space"]
        unit_space = RegressionUnitSpace(
            smearing_factor=us["smearing_factor"],
            r_squared_unit=us["r_squared_unit"],
            adjusted_r2_unit=us["adjusted_r2_unit"],
            rmse_unit=us["rmse_unit"],
            prediction_point_unit=us["prediction_point_unit"],
            prediction_ci_lower_unit=us["prediction_ci_lower_unit"],
            prediction_ci_upper_unit=us["prediction_ci_upper_unit"],
            prediction_pi_lower_unit=us["prediction_pi_lower_unit"],
            prediction_pi_upper_unit=us["prediction_pi_upper_unit"],
            predictions_unit=tuple(us["predictions_unit"]),
            residuals_unit=tuple(us["residuals_unit"]),
            model_formula=us["model_formula"],
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
                unit_space=unit_space,
            ),
        ))
    return result


def get_analysis_results(
    csv_path: Path,
    cache_path: Path = DEFAULT_CACHE_PATH,
) -> list[tuple[str, bool, RegressionSheetResults]]:
    """Return regression sheet configs from cache or fresh.

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
                return _deserialize_regression_sheet_configs(cached["regression_sheet_configs"])
        except (json.JSONDecodeError, KeyError, TypeError, ValueError, OSError):
            pass

    from .analyze_regression_sheet import build_regression_sheet_qc_configs

    regression_sheet_configs = build_regression_sheet_qc_configs(csv_path)

    try:
        payload = {
            "schema_version": _CACHE_SCHEMA_VERSION,
            "csv_fingerprint": fingerprint,
            "regression_sheet_configs": _serialize_regression_sheet_configs(regression_sheet_configs),
        }
        with cache_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
    except (OSError, TypeError):
        pass

    return regression_sheet_configs

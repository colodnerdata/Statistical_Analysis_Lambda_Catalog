"""Disk cache for OLS analysis results, keyed on CSV file content hash."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

from .analyze_life_expectancy import RegressionObservationVectors, RegressionVectors
from .write_sheet_mlr_scalar_test import build_mlr_row_configs
from .write_sheet_mlr_observation_test import build_mlr_observation_row_configs
from .write_sheet_mlr_vector_outputs_test import build_mlr_vector_row_configs


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CACHE_PATH = ROOT_DIR / ".analysis_cache.json"
_CACHE_SCHEMA_VERSION = 4


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
        vectors = RegressionObservationVectors(**{k: tuple(v) for k, v in item.items() if k not in {"k", "allow_intercept"}})
        result.append((item["k"], item["allow_intercept"], vectors))
    return result


def get_analysis_results(
    csv_path: Path,
    cache_path: Path = DEFAULT_CACHE_PATH,
) -> tuple[list, list[tuple[int, bool, RegressionVectors]], list[tuple[int, bool, RegressionObservationVectors]]]:
    """Return (scalar, vector, observation configs), from cache or computed fresh.

    The cache is invalidated when the CSV content changes (SHA-256 hash).
    Delete .analysis_cache.json manually after code or schema changes.
    """
    csv_path = csv_path.resolve()
    fingerprint = _csv_fingerprint(csv_path)

    if cache_path.exists():
        try:
            with cache_path.open("r", encoding="utf-8") as handle:
                cached = json.load(handle)
            if cached.get("schema_version") == _CACHE_SCHEMA_VERSION and cached.get("csv_fingerprint") == fingerprint:
                scalar_configs = [tuple(item) for item in cached["scalar_row_configs"]]
                vector_configs = _deserialize_vector_configs(cached["vector_row_configs"])
                observation_configs = _deserialize_observation_configs(cached["observation_row_configs"])
                return scalar_configs, vector_configs, observation_configs
        except (json.JSONDecodeError, KeyError, TypeError, ValueError, OSError):
            pass

    scalar_configs = build_mlr_row_configs(csv_path)
    vector_configs = build_mlr_vector_row_configs(csv_path)
    observation_configs = build_mlr_observation_row_configs(csv_path)

    try:
        payload = {
            "schema_version": _CACHE_SCHEMA_VERSION,
            "csv_fingerprint": fingerprint,
            "scalar_row_configs": [list(item) for item in scalar_configs],
            "vector_row_configs": _serialize_vector_configs(vector_configs),
            "observation_row_configs": _serialize_observation_configs(observation_configs),
        }
        with cache_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
    except (OSError, TypeError):
        pass

    return scalar_configs, vector_configs, observation_configs

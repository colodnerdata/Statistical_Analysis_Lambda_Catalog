"""Shared regression dataclasses used by analysis engines, cache, and sheet writers."""
from __future__ import annotations

from dataclasses import dataclass


FEATURE_COLUMNS = [
    "Adult Mortality",
    "infant deaths",
    "Alcohol",
    "percentage expenditure",
    "Hepatitis B",
    "Measles",
    "BMI",
    "under-five deaths",
    "Polio",
    "Total expenditure",
    "Diphtheria",
    "HIV/AIDS",
    "GDP",
    "Population",
    "thinness 1-19 years",
    "thinness 5-9 years",
    "Income composition of resources",
    "Schooling",
]


@dataclass(frozen=True)
class RegressionVectors:
    """Per-coefficient regression statistics aligned with workbook vector outputs."""

    term_names: tuple[str, ...]
    coefficients: tuple[float, ...]
    std_errors: tuple[float, ...]
    t_stats: tuple[float, ...]
    p_values: tuple[float, ...]
    ci_lower: tuple[float, ...]
    ci_upper: tuple[float, ...]
    beta_weights: tuple[float, ...]


@dataclass(frozen=True)
class RegressionSummary:
    """Core scalar regression metrics aligned with workbook outputs."""

    observations: int
    df_regression: int
    df_total: int
    r_squared: float
    df_residual: int
    multiple_r: float
    adjusted_r2: float
    ss_total: float
    ss_residual: float
    ss_regression: float
    se_regression: float
    press: float
    durbin_watson: float
    f_stat: float
    p_value_f: float
    aic: float
    bic: float
    aicc: float
    qq_correlation: float


@dataclass(frozen=True)
class RegressionObservationVectors:
    """Observation-level diagnostics aligned with workbook spill outputs."""

    observation_num: tuple[int, ...]
    rank_fraction: tuple[float, ...]
    y_ranked: tuple[float, ...]
    normal_scores: tuple[float, ...]
    predictions: tuple[float, ...]
    residuals: tuple[float, ...]
    scaled_residuals: tuple[float, ...]
    scaled_residuals_ranked: tuple[float, ...]


@dataclass(frozen=True)
class RegressionPredictorSummary:
    """Per-predictor summary statistics for the Regression sheet."""

    predictor_names: tuple[str, ...]
    pearson_r: tuple[float, ...]
    spearman_r: tuple[float, ...]
    skewness: tuple[float, ...]
    kurtosis: tuple[float, ...]
    gvif: tuple[float, ...]
    tolerance: tuple[float, ...]


@dataclass(frozen=True)
class RegressionFullResiduals:
    """Per-observation diagnostics for the Regression sheet residual zone."""

    dependent_var: tuple[float, ...]
    predictions: tuple[float, ...]
    residuals: tuple[float, ...]
    loocv_residuals: tuple[float, ...]
    hat_diagonal: tuple[float, ...]
    studentized_residuals: tuple[float, ...]
    cooks_distance: tuple[float, ...]
    normal_scores_ranked: tuple[float, ...]
    studentized_residuals_ranked: tuple[float, ...]
    scale_location: tuple[float, ...]


@dataclass(frozen=True)
class RegressionPredictionInterval:
    """Prediction interval outputs for the Regression sheet.

    The v2.1 group-mean-recovery shape (``Group_Prediction_Interval``):
    both a mean-response CI and a wider new-observation PI, plus the
    selected group's own mean/count (AH13/AH14). A no-FE case selects the
    constant ``"(all)"`` group, which collapses this exactly to the
    pre-v2.1 single-PI numbers (se_new/pi_lower/pi_upper matching the old
    se_prediction/lower/upper) — see
    ``tests/test_group_prediction_interval.py``.
    """

    pred_input_values: tuple[float, ...]
    point_estimate: float
    se_mean: float
    se_new: float
    t_critical: float
    ci_lower: float
    ci_upper: float
    pi_lower: float
    pi_upper: float
    confidence_level: float
    group_mean: float
    group_count: int


@dataclass(frozen=True)
class RegressionSheetResults:
    """All expected values for one Regression sheet QC configuration."""

    summary: RegressionSummary
    vectors: RegressionVectors
    predictor_summary: RegressionPredictorSummary
    full_residuals: RegressionFullResiduals
    prediction_interval: RegressionPredictionInterval

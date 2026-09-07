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
    # The panel form, for the state plain DW cannot read. The two are
    # mutually exclusive by construction and exactly one is a number at a
    # time: no Sequence axis leaves both NaN; Sequence without Fixed Effects
    # gives DW a value and NaN here; Sequence WITH Fixed Effects inverts
    # that, because within-demeaned residuals in a panel have no single
    # ordering and row-adjacent differencing manufactures correlation at
    # every group seam. No default, deliberately: constructing a summary
    # should fail loudly when a caller has not computed this rather than
    # quietly admit one whose panel diagnostic is silently absent.
    bfn_panel_durbin_watson: float
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
class RegressionUnitSpace:
    """v3.3 unit-space / back-transformation outputs for the Regression sheet.

    Mirrors the AG4:AH10 unit-space block, the Original Units prediction column
    (AL), the two v3.3 residual columns (AZ/BA), and the v3.4 leave-one-out
    family: the LOOCV residual column (BB) plus the LOOCV RMSE / MAE scalars
    (AH12/AH13) and the named smearing treatment (AH14). The smearing factor
    is the scalar that lifts EXP(ŷ) from a median predictor to a mean predictor
    under a Log response; under ``None`` it is exactly 1 so the reduction
    invariant (Unit_Space_* ≡ ordinary statistic) holds.
    """

    smearing_factor: float
    r_squared_unit: float
    adjusted_r2_unit: float
    rmse_unit: float
    prediction_point_unit: float
    prediction_ci_lower_unit: float
    prediction_ci_upper_unit: float
    prediction_pi_lower_unit: float
    prediction_pi_upper_unit: float
    predictions_unit: tuple[float, ...]
    residuals_unit: tuple[float, ...]
    loocv_residuals_unit: tuple[float, ...]
    loocv_rmse_unit: float
    loocv_mae_unit: float
    smearing_treatment: str
    model_formula: str


@dataclass(frozen=True)
class RegressionSheetResults:
    """All expected values for one Regression sheet QC configuration."""

    summary: RegressionSummary
    vectors: RegressionVectors
    predictor_summary: RegressionPredictorSummary
    full_residuals: RegressionFullResiduals
    prediction_interval: RegressionPredictionInterval
    unit_space: RegressionUnitSpace

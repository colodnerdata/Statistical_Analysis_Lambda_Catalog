"""Push a spec case onto a Regression-shaped sheet, and read one back.

Two halves of one contract, extracted from ``tools/inspect_regression_sheet.py``
so that the three things that now need them cannot drift apart:

* ``tools/inspect_regression_sheet.py`` — the legacy single-sheet verifier,
  which writes each case onto the one ``Regression`` sheet in turn and reads
  it back.
* ``lambda_catalog/write_sheet_test_model.py`` — the test-model sheet
  builder, which writes each case onto its OWN sheet once, at build time.
* ``tools/inspect_test_model_sheets.py`` — the test-model verifier, which
  reads those sheets back and writes nothing at all.

The write half and the read half used to be private functions inside the
inspector. A second copy of either is exactly the drift this repo's
one-source-of-truth rule exists to prevent: a builder that writes the spec
one way and a verifier that reads it another would disagree about what a
case *is*, and the disagreement would surface as a QC failure blamed on the
workbook.

Every cell address here is imported from the sheet writers, never spelled
out — the same rule ``tools/inspect_regression_sheet.py`` and
``lambda_catalog/analyze_regression_spec_block.py`` already follow.
"""
from __future__ import annotations

import math
from typing import Any

import xlwings as xw

from .analyze_regression_spec import RegressionSpecExpected
from .inspection_compare import compare_values, to_float_or_none
from .write_spec_block import (
    _C_INCLUDE as _C_SPEC_INCLUDE,
)
from .write_spec_block import (
    _C_INTERACTION_OPERATION as _C_SPEC_INTERACTION_OPERATION,
)
from .write_spec_block import (
    _C_INTERACTION_TERM as _C_SPEC_INTERACTION_TERM,
)
from .write_spec_block import (
    _C_REFERENCE as _C_SPEC_REFERENCE,
)
from .write_spec_block import (
    _C_ROLE as _C_SPEC_ROLE,
)
from .write_spec_block import (
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
)
from .write_spec_block import (
    _C_SEQUENCE_PERIOD as _C_SPEC_SEQUENCE_PERIOD,
)
from .write_spec_block import (
    _C_TRANSFORM as _C_SPEC_TRANSFORM,
)
from .write_spec_block import (
    _C_TYPE as _C_SPEC_TYPE,
)
from .write_spec_block import (
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
)
from .write_spec_block import (
    _INTERCEPT_ROW,
)
from .write_spec_block import (
    _LAST_DATA_ROW as _SPEC_LAST_DATA_ROW,
)
from .write_sheet_regression import (
    _C_AA,
    _C_AB,
    _C_AC,
    _C_AD,
    _C_AE,
    _C_AF,
    _C_AG,
    _C_AH,
    _C_AK,
    _C_AO,
    _C_AX,
    _C_MODEL_FORMULA,
    _C_T,
    _C_U,
    _C_V,
    _C_W,
    _C_X,
    _C_Y,
    _ROW_MODEL_FORMULA,
)

# ── Row positions (1-based) ──────────────────────────────────────────────
# The writers state these only as literals inside their own formula loops,
# so there is nothing to import. They are instead pinned against the writers'
# actual output by `test_row_constants_match_the_writers_own_layout` in
# tests/test_test_model_sheets.py, which runs each zone writer against a
# RecordingSheet and asserts the labelled cells land on the rows named here.
# Without that pin a zone could move and this reader would silently compare
# the wrong cells — reporting a wrong NUMBER rather than an error, which is
# the failure mode the whole layout-constant discipline exists to prevent.
ROW_SUMMARY_FIRST = 3      # S3 spills constructed names; T3–Y3 the stats
ROW_MULTIPLE_R = 4
ROW_R_SQUARED = 5
ROW_ADJ_R2 = 6
ROW_SE_REG = 7
ROW_OBS = 8

ROW_PRESS = 4
ROW_PRESS_R2 = 5
ROW_MEAN_LEV = 6
ROW_AIC = 7
ROW_BIC = 8
ROW_AICC = 9
ROW_QQ_CORR = 10
ROW_DURBIN_WATSON = 11
ROW_BFN_PANEL_DW = 12

ROW_ANOVA_REG = 15
ROW_ANOVA_RES = 16
ROW_ANOVA_TOT = 17

ROW_COEFF_DATA = 21        # AA21 spills coefficient labels (k+1 rows)
ROW_PI_POINT = 3           # AK3 = point estimate
ROW_FE_GROUP = 12          # AK12 = FE Group selector (an input)
ROW_GROUP_MEAN = 13
ROW_GROUP_COUNT = 14
ROW_PRED_INPUT_FIRST = 19  # AK19 = first user-editable predictor value
ROW_PRED_INPUT_LAST = 62   # end of the guarded prefill band
ROW_RESID_FIRST = 3

# The v3.3 Back-Transform Method input ($AH$4).
ROW_BACK_TRANSFORM = 4

# Statistics that must compare scale-free — as 3 SIGNIFICANT digits rather
# than 3 decimal places — because their accuracy is bounded by the data's own
# magnitude rather than by the formula.
#
# Two families, arriving for the same underlying reason:
#
# * Sums of squares and PRESS scale with the response column. Raw production
#   cost data runs into the 1e10 range, where IEEE-754 leaves ~6 decimal
#   digits — already above a 3-decimal tolerance.
# * t-statistics and p-values inherit the CONDITIONING of the normal
#   equations. L05 (the shipped life_expectancy profile) mixes Population,
#   which spans 34 to 1.3e9, with predictors of order 1-100; Excel's LINEST
#   and statsmodels then disagree in the 6th significant digit on the one
#   coefficient that is statistically indistinguishable from zero
#   (t = -0.367, p = 0.71). That is a floating-point property of the design,
#   not a disagreement about the model.
#
# Both sides are divided by the same factor, so a genuinely wrong number
# still fails — this widens the unit, never the tolerance.
SCALE_FREE_STATS = frozenset({
    "SS_Regression", "SS_Residual", "SS_Total",
    "MS_Regression", "MS_Residual",
    "PRESS",
    "T_Statistics", "P_Values",
})


# ── Write half ───────────────────────────────────────────────────────────


def apply_spec_case(sheet: xw.Sheet, expected: RegressionSpecExpected) -> None:
    """Write one case's full visible spec onto ``sheet``.

    Retargets ``Source_Table``, sets the intercept toggle, the FE group
    selector and the Back-Transform Method, then clears and rewrites every
    spec row.

    Four details are load-bearing and were all learned the hard way:

    * ``Source_Table`` is the ONE name that retargets which data sheet the
      spec block reads from, and every case sets it unconditionally — a case
      must never be evaluated against whatever dataset the previous write
      left behind.
    * ``$AK$12`` is always written explicitly. Typing a value into that cell
      REPLACES its default-computing formula, so "leave it alone" silently
      carries the previous group forward.
    * Only the spec rows are cleared, not the whole column: the Sequence
      Spacing block lives below the spec on the Regression sheet.
    * A row with no interaction gets its M/N cells genuinely BLANK rather
      than ``""``. Both satisfy ``mate()``'s ``LEN(t&"")=0`` gate, but a
      written ``""`` defeats the dropdown's own blank default.
    """
    sheet.api.Names.Item("Source_Table").RefersTo = expected.case.source_table_ref
    sheet.range((_INTERCEPT_ROW, _C_SPEC_INCLUDE)).value = (
        expected.case.allow_intercept
    )
    sheet.range(ROW_FE_GROUP, _C_AK).value = expected.resolved_prediction_group
    sheet.range(ROW_BACK_TRANSFORM, _C_AH).value = expected.case.back_transform

    last_row = max(
        _SPEC_LAST_DATA_ROW + 1,
        _SPEC_FIRST_DATA_ROW + len(expected.case.spec) - 1,
    )
    for row in range(_SPEC_FIRST_DATA_ROW, last_row + 1):
        for col in (
            _C_SPEC_ROLE,
            _C_SPEC_INCLUDE,
            _C_SPEC_TYPE,
            _C_SPEC_REFERENCE,
            _C_SPEC_SEQUENCE,
            _C_SPEC_TRANSFORM,
            _C_SPEC_INTERACTION_TERM,
            _C_SPEC_INTERACTION_OPERATION,
        ):
            sheet.range(row, col).clear_contents()

    for offset, variable in enumerate(expected.case.spec):
        row = _SPEC_FIRST_DATA_ROW + offset
        sheet.range(row, _C_SPEC_ROLE).value = variable.role
        sheet.range(row, _C_SPEC_INCLUDE).value = variable.include
        sheet.range(row, _C_SPEC_TYPE).value = variable.var_type
        sheet.range(row, _C_SPEC_REFERENCE).value = variable.reference
        sheet.range(row, _C_SPEC_SEQUENCE).value = variable.sequence
        sheet.range(row, _C_SPEC_TRANSFORM).value = variable.transform
        if variable.interaction_term:
            sheet.range(row, _C_SPEC_INTERACTION_TERM).value = (
                variable.interaction_term
            )
        if variable.interaction_operation:
            sheet.range(row, _C_SPEC_INTERACTION_OPERATION).value = (
                variable.interaction_operation
            )


def apply_sequence_period_overrides(
    sheet: xw.Sheet,
    spec: tuple,
    overrides: dict[str, float],
) -> None:
    """Type Sequence Period values into spec column I.

    Column I is an INPUT: a typed number there replaces the Period In Use
    cell's ``Base_Period_Delta_Candidate()`` formula, which is the whole
    mechanism M16 and P07 exist to test. Rows with no override are left
    alone so their J cell keeps computing the candidate.
    """
    for offset, variable in enumerate(spec):
        if variable.name in overrides:
            sheet.range(
                _SPEC_FIRST_DATA_ROW + offset, _C_SPEC_SEQUENCE_PERIOD
            ).value = overrides[variable.name]


def set_prediction_inputs(
    sheet: xw.Sheet,
    pred_input_values: tuple[float, ...],
    constructed_column_transforms: tuple[str, ...],
) -> None:
    """Write prediction inputs to the AK band, one per constructed column.

    The constructed columns are the selected predictors in table order —
    identical to the case's column order — so values are written
    positionally, and the rest of the guarded prefill band is cleared so a
    previous case's values cannot linger.

    ``pred_input_values`` is each design column's mean, which for a
    Log-transformed column is already in LOG space (the transform is applied
    before the mean is taken). The sheet's AK cells are raw-value inputs —
    the row-3 prediction formula applies ``Ln_Positive`` itself, matching
    the confirmed UX where a user types a real-world value, never ln(x) — so
    a logged column's mean is EXP'd back to input space here. That mirrors
    the AI19 Training Mean spill's own geometric-mean handling in
    write_sheet_regression.py.
    """
    sheet.range(
        (ROW_PRED_INPUT_FIRST, _C_AK), (ROW_PRED_INPUT_LAST, _C_AK)
    ).clear_contents()
    for index, value in enumerate(pred_input_values):
        if constructed_column_transforms[index] == "Log":
            value = math.exp(value)
        sheet.range(ROW_PRED_INPUT_FIRST + index, _C_AK).value = value


def apply_case_inputs(sheet: xw.Sheet, expected: RegressionSpecExpected) -> None:
    """Apply a fittable case's visible inputs: spec, typed Sequence Period, prediction inputs.

    The three writes every spec-driven case needs on top of a sheet whose spec
    block already exists: ``apply_spec_case`` (Source_Table retarget, intercept,
    FE group, back-transform, the spec rows), the typed Sequence Period into
    column I, and ``set_prediction_inputs``. Centralizing the sequence is what
    stops a second call site forgetting the middle step — which is how the BFN
    panel Durbin-Watson cell (AE12) sat at ``nan`` for every verify run on
    record while the oracle held a real number: the inspector duplicated this
    sequence and dropped the override, and the only symptom was a multi-minute
    Excel run that read like a broken diagnostic.

    Padding the spec to the source table's width is the CALLER's concern, not
    this helper's: the builder pads (it writes a fresh sheet sized to the
    dataset and wants every column shown), the inspector does not (it writes
    the fixed production sheet). Pass the padded ``expected`` if you want
    padding — the Sequence-Period dict, row offsets, and prediction inputs are
    all unchanged by appended ``Omit`` padding (``pad_spec_to_source_table``
    appends, it does not interleave), so the helper reads the same values
    either way.

    Only cases that declare a period are touched by the override, so non-panel
    configs are unaffected.
    """
    apply_spec_case(sheet, expected)
    case = expected.case

    # Clear any previously-typed Sequence Period values so one case cannot
    # affect the next when reusing the same sheet (e.g., the verify inspector).
    last_row = max(_SPEC_LAST_DATA_ROW + 1, _SPEC_FIRST_DATA_ROW + len(case.spec) - 1)
    sheet.range(
        (_SPEC_FIRST_DATA_ROW, _C_SPEC_SEQUENCE_PERIOD),
        (last_row, _C_SPEC_SEQUENCE_PERIOD),
    ).clear_contents()

    if case.sequence_period is not None:
        apply_sequence_period_overrides(
            sheet,
            case.spec,
            {item.name: case.sequence_period for item in case.spec if item.sequence},
        )
        sheet,
        expected.results.prediction_interval.pred_input_values,
        expected.design.constructed_column_transforms,
    )


# ── Read half ────────────────────────────────────────────────────────────


def read_cell(sheet: xw.Sheet, row: int, col: int) -> float | None:
    """One numeric cell, or None for blank/error/text."""
    return to_float_or_none(sheet.range(row, col).value)


def read_col(
    sheet: xw.Sheet, start_row: int, col: int, n_rows: int
) -> list[float | None]:
    """``n_rows`` numeric cells down one column."""
    if n_rows <= 0:
        return []
    rng = sheet.range((start_row, col), (start_row + n_rows - 1, col))
    raw: list[Any] = rng.value if n_rows > 1 else [rng.value]
    return [to_float_or_none(value) for value in raw]


def read_block(
    sheet: xw.Sheet, start_row: int, col1: int, col2: int, n_rows: int
) -> list[list[float | None]]:
    """A rectangular numeric block, as a list of rows."""
    if n_rows <= 0:
        return []
    rng = sheet.range((start_row, col1), (start_row + n_rows - 1, col2))
    raw = rng.value
    if n_rows == 1:
        raw = [raw]
    return [[to_float_or_none(value) for value in row] for row in raw]


def read_case_comparison_rows(
    sheet: xw.Sheet, expected: RegressionSpecExpected
) -> dict[str, list[dict]]:
    """Read every output zone on ``sheet`` and compare against the oracle.

    Returns the five comparison sections — ``scalars``, ``predictors``,
    ``coefficients``, ``prediction_interval``, ``residuals`` — each a list
    of plain dicts carrying the expected value, the Excel value, and the
    two difference measures. Assembling those into DataFrames, filtering
    them by tolerance and rendering failures is the caller's job; this
    function neither writes to the sheet nor recalculates it.
    """
    case = expected.case
    results = expected.results
    summary = results.summary
    vectors = results.vectors
    predictor_summary = results.predictor_summary
    residuals = results.full_residuals
    interval = results.prediction_interval
    k = len(predictor_summary.predictor_names)
    n = summary.observations

    identity = {
        "config_name": case.name,
        "allow_intercept": case.allow_intercept,
    }

    def _row(extra: dict, expected_value, excel_value, scale_free=False) -> dict:
        diff, fdd = compare_values(
            expected_value, excel_value, scale_free=scale_free
        )
        return {
            **identity,
            **extra,
            "expected": expected_value,
            "excel_calc": excel_value,
            "abs_diff": diff,
            "first_digit_deviation": fdd,
        }

    # ── Scalars: Regression Statistics, diagnostics, ANOVA ───────────────
    press_r2 = 1.0 - summary.press / summary.ss_total
    mean_leverage = (
        summary.df_regression + (1 if case.allow_intercept else 0)
    ) / summary.observations
    scalar_specs: list[tuple[str, float, int, int]] = [
        ("Multiple_R", summary.multiple_r, ROW_MULTIPLE_R, _C_AB),
        ("R_Squared", summary.r_squared, ROW_R_SQUARED, _C_AB),
        ("Adjusted_R_Squared", summary.adjusted_r2, ROW_ADJ_R2, _C_AB),
        ("SE_Regression", summary.se_regression, ROW_SE_REG, _C_AB),
        ("Observations", float(summary.observations), ROW_OBS, _C_AB),
        ("PRESS", summary.press, ROW_PRESS, _C_AE),
        ("PRESS_R2", press_r2, ROW_PRESS_R2, _C_AE),
        ("Mean_Leverage", mean_leverage, ROW_MEAN_LEV, _C_AE),
        ("AIC", summary.aic, ROW_AIC, _C_AE),
        ("BIC", summary.bic, ROW_BIC, _C_AE),
        ("AICc", summary.aicc, ROW_AICC, _C_AE),
        ("QQ_Correlation", summary.qq_correlation, ROW_QQ_CORR, _C_AE),
        ("Durbin_Watson", summary.durbin_watson, ROW_DURBIN_WATSON, _C_AE),
        # The panel form at AE12, mutually gated with AE11 above: the
        # oracle NaNs whichever of the two the sheet shows as text, so
        # exactly one of this pair is ever a live comparison. Before
        # this row existed, a Fixed Effects sheet had NO verified
        # serial-correlation diagnostic at all — DW is NaN by design
        # there, and nothing read the cell that holds the number.
        (
            "BFN_Panel_Durbin_Watson",
            summary.bfn_panel_durbin_watson,
            ROW_BFN_PANEL_DW, _C_AE,
        ),
        (
            "Regression_Degrees_Of_Freedom",
            float(summary.df_regression), ROW_ANOVA_REG, _C_AB,
        ),
        ("SS_Regression", summary.ss_regression, ROW_ANOVA_REG, _C_AC),
        (
            "MS_Regression",
            summary.ss_regression / summary.df_regression,
            ROW_ANOVA_REG, _C_AD,
        ),
        ("F_Statistic", summary.f_stat, ROW_ANOVA_REG, _C_AE),
        ("F_Statistic_P_Value", summary.p_value_f, ROW_ANOVA_REG, _C_AF),
        (
            "Residual_Degrees_Of_Freedom",
            float(summary.df_residual), ROW_ANOVA_RES, _C_AB,
        ),
        ("SS_Residual", summary.ss_residual, ROW_ANOVA_RES, _C_AC),
        (
            "MS_Residual",
            summary.ss_residual / summary.df_residual,
            ROW_ANOVA_RES, _C_AD,
        ),
        ("Total_Degrees_Of_Freedom", float(summary.df_total), ROW_ANOVA_TOT, _C_AB),
        ("SS_Total", summary.ss_total, ROW_ANOVA_TOT, _C_AC),
    ]
    scalar_rows = [
        _row(
            {"stat_name": stat_name},
            expected_value,
            read_cell(sheet, row, col),
            scale_free=stat_name in SCALE_FREE_STATS,
        )
        for stat_name, expected_value, row, col in scalar_specs
    ]

    # ── v3.3 unit-space block (AH5:AH8) ─────────────────────────────────
    # Never previously compared cell-for-cell by this harness even though
    # the oracle has carried the numbers since v3.3; the Back-Transform
    # toggle (L02 vs L03) is only meaningful if these four are read.
    unit = results.unit_space
    for stat_name, expected_value, row in (
        ("Smearing_Factor", unit.smearing_factor, 5),
        ("Unit_Space_R_Squared", unit.r_squared_unit, 6),
        ("Unit_Space_Adjusted_R_Squared", unit.adjusted_r2_unit, 7),
        ("Unit_Space_RMSE", unit.rmse_unit, 8),
    ):
        scalar_rows.append(
            _row(
                {"stat_name": stat_name},
                expected_value,
                read_cell(sheet, row, _C_AH),
            )
        )

    # ── Predictor Summary (T–Y, k rows from row 3) ──────────────────────
    predictor_rows = [
        _row(
            {
                "predictor_name": predictor_summary.predictor_names[j],
                "stat_name": stat_name,
            },
            expected_value,
            excel_value,
        )
        for stat_name, expected_tuple, col in (
            ("Pearson_R", predictor_summary.pearson_r, _C_T),
            ("Spearman_R", predictor_summary.spearman_r, _C_U),
            ("Skewness", predictor_summary.skewness, _C_V),
            ("Kurtosis", predictor_summary.kurtosis, _C_W),
            ("GVIF", predictor_summary.gvif, _C_X),
            ("Tolerance", predictor_summary.tolerance, _C_Y),
        )
        for j, (expected_value, excel_value) in enumerate(
            zip(expected_tuple, read_col(sheet, ROW_SUMMARY_FIRST, col, k))
        )
    ]

    # ── Coefficients (AB–AG, k+1 rows from row 21) ──────────────────────
    # Intercept models spill k+1 rows, intercept first. No-intercept models
    # prepend one blank display row so predictor rows stay aligned across
    # both; drop it before comparing.
    n_coefficient_rows = k + 1
    coefficient_rows: list[dict] = []
    for stat_name, expected_tuple, col in (
        ("Coefficients", vectors.coefficients, _C_AB),
        ("SE_Coefficients", vectors.std_errors, _C_AC),
        ("T_Statistics", vectors.t_stats, _C_AD),
        ("P_Values", vectors.p_values, _C_AE),
        ("Confidence_Interval_Lower", vectors.ci_lower, _C_AF),
        ("Confidence_Interval_Upper", vectors.ci_upper, _C_AG),
    ):
        excel_values = read_col(sheet, ROW_COEFF_DATA, col, n_coefficient_rows)
        if not case.allow_intercept:
            excel_values = excel_values[1:]
        coefficient_rows.extend(
            _row(
                {"term_name": vectors.term_names[i], "stat_name": stat_name},
                expected_value,
                excel_value,
                scale_free=stat_name in SCALE_FREE_STATS,
            )
            for i, (expected_value, excel_value) in enumerate(
                zip(expected_tuple, excel_values)
            )
        )
    beta_values = read_col(sheet, ROW_COEFF_DATA, _C_AH, n_coefficient_rows)[1:]
    coefficient_rows.extend(
        _row(
            {
                "term_name": predictor_summary.predictor_names[i],
                "stat_name": "Beta_Weights",
            },
            expected_value,
            excel_value,
        )
        for i, (expected_value, excel_value) in enumerate(
            zip(vectors.beta_weights, beta_values)
        )
    )

    # ── Prediction Interval (AK3:AK14) ──────────────────────────────────
    interval_specs = (
        ("Point_Estimate", interval.point_estimate),
        ("SE_Mean", interval.se_mean),
        ("SE_New", interval.se_new),
        ("T_Critical", interval.t_critical),
        ("CI_Lower", interval.ci_lower),
        ("CI_Upper", interval.ci_upper),
        ("PI_Lower", interval.pi_lower),
        ("PI_Upper", interval.pi_upper),
        ("Confidence_Level", interval.confidence_level),
    )
    interval_values = read_col(sheet, ROW_PI_POINT, _C_AK, 9)
    interval_rows = [
        _row({"stat_name": stat_name}, expected_value, excel_value)
        for (stat_name, expected_value), excel_value in zip(
            interval_specs, interval_values
        )
    ]
    interval_rows.extend(
        _row({"stat_name": stat_name}, expected_value, read_cell(sheet, row, _C_AK))
        for stat_name, expected_value, row in (
            ("Group_Mean", interval.group_mean, ROW_GROUP_MEAN),
            ("Group_Count", float(interval.group_count), ROW_GROUP_COUNT),
        )
    )

    # ── Residual Output (AO–AX, n rows from row 3) ──────────────────────
    residual_stats = (
        ("Dependent_Variable", residuals.dependent_var),
        ("Predictions", residuals.predictions),
        ("Residuals", residuals.residuals),
        ("Hat_Diagonal", residuals.hat_diagonal),
        ("Studentized_Residuals", residuals.studentized_residuals),
        ("Cooks_Distance", residuals.cooks_distance),
        ("Normal_Scores_Ranked", residuals.normal_scores_ranked),
        ("Studentized_Residuals_Ranked", residuals.studentized_residuals_ranked),
        ("Scale_Location", residuals.scale_location),
        ("PRESS_Residual", residuals.loocv_residuals),
    )
    residual_rows = [
        _row(
            {"row_idx": row_index + 1, "stat_name": stat_name},
            float(expected_tuple[row_index])
            if row_index < len(expected_tuple)
            else None,
            excel_value,
        )
        for row_index, excel_row in enumerate(
            read_block(sheet, ROW_RESID_FIRST, _C_AO, _C_AX, n)
        )
        for (stat_name, expected_tuple), excel_value in zip(
            residual_stats, excel_row
        )
    ]

    return {
        "scalars": scalar_rows,
        "predictors": predictor_rows,
        "coefficients": coefficient_rows,
        "prediction_interval": interval_rows,
        "residuals": residual_rows,
    }


def read_model_formula(sheet: xw.Sheet) -> object:
    """The Model Formula readout, as written text.

    It sits on row 1 of the §4b band's terminal Constructed Design Matrix
    zone, right of that zone's heading (it used to be AB2); the coordinates
    come from the layout constants, so a future move needs no edit here.
    """
    return sheet.range(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA).value


def read_response_readout(sheet: xw.Sheet) -> object:
    """The AF2 Predicted Variable readout."""
    return sheet.range(2, _C_AF).value


def read_regression_outputs_label(sheet: xw.Sheet) -> object:
    """The AA1 zone heading — a cheap sanity check that a sheet is a
    Regression-shaped sheet before anything else is read off it."""
    return sheet.range(1, _C_AA).value

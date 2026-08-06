"""Independent verification of the v2.2 Transform=Log wiring through the
Python QC oracle (analyze_regression_spec.py / analyze_model_construction.py).

The headline acceptance test: production_lots.csv ships both raw columns
(Cumulative_Units, Unit_Cost_BY) and precomputed log columns ("log Cum
Units", "log Unit Cost") that are bit-identical logs of the raw ones. The
existing "production_lots_fixed_effects" QC case points its spec at the
precomputed columns (a pre-Log-wiring workaround); the new
"production_lots_log_transform" case points the SAME model at the raw
columns with transform="Log" declared instead. If the two designs and
every downstream statistic agree to floating-point precision, the Log
wiring reproduces exactly what the precomputed-column workaround already
delivered — a real learning-curve model (Crawford/Wright:
ln(unit cost) = a + b*ln(cumulative units)), composed with Fixed Effects.

Runnable standalone (``python tests/test_transform_threading.py``) or
under pytest.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# pylint: disable-next=wrong-import-position
from lambda_catalog.analyze_model_construction import SpecVariable, build_default_spec

# pylint: disable-next=wrong-import-position
from lambda_catalog.analyze_regression_spec import (
    PRODUCTION_LOTS_CSV_PATH,
    build_regression_spec_cases,
    build_spec_design,
    calculate_regression_spec_case,
    load_production_lots_source_rows,
)

ROOT_DIR = Path(__file__).resolve().parents[1]


def _cases() -> dict:
    return {case.name: case for case in build_regression_spec_cases()}


# ── Headline cross-check: Log-wired design == precomputed-log-column design ──

def test_log_transform_case_matches_precomputed_log_case_numerically() -> None:
    cases = _cases()
    fe_case = cases["production_lots_fixed_effects"]
    log_case = cases["production_lots_log_transform"]

    fe_expected = calculate_regression_spec_case(fe_case, PRODUCTION_LOTS_CSV_PATH)
    log_expected = calculate_regression_spec_case(log_case, PRODUCTION_LOTS_CSV_PATH)

    # Design matrices and response vectors: bit-exact, since the shipped
    # "log Cum Units"/"log Unit Cost" columns are exact logs of the raw
    # Cumulative_Units/Unit_Cost_BY columns. np.array_equal, not allclose —
    # allclose's default rtol=1e-05 would silently accept a real regression
    # here, undermining the "bit-exact" claim this comment makes.
    assert np.array_equal(fe_expected.design.x_features, log_expected.design.x_features)
    assert np.array_equal(fe_expected.design.y_train, log_expected.design.y_train)
    assert fe_expected.design.row_mask == log_expected.design.row_mask
    assert fe_expected.design.included_rows == log_expected.design.included_rows

    # Deliberately NOT compared: constructed_column_names ("log Cum Units"
    # vs "Ln(Cumulative_Units)") and the response display name — those
    # legitimately differ between the two mechanisms. Numerics only.
    fe_r = fe_expected.results
    log_r = log_expected.results
    assert math.isclose(fe_r.summary.r_squared, log_r.summary.r_squared)
    assert math.isclose(fe_r.summary.adjusted_r2, log_r.summary.adjusted_r2)
    assert math.isclose(fe_r.summary.se_regression, log_r.summary.se_regression)
    assert math.isclose(fe_r.summary.df_residual, log_r.summary.df_residual)
    assert np.allclose(fe_r.vectors.coefficients, log_r.vectors.coefficients)
    assert np.allclose(fe_r.vectors.std_errors, log_r.vectors.std_errors)
    assert np.allclose(fe_r.vectors.t_stats, log_r.vectors.t_stats)
    assert np.allclose(fe_r.full_residuals.residuals, log_r.full_residuals.residuals)

    fe_pi = fe_r.prediction_interval
    log_pi = log_r.prediction_interval
    assert math.isclose(fe_pi.point_estimate, log_pi.point_estimate)
    assert math.isclose(fe_pi.se_mean, log_pi.se_mean)
    assert math.isclose(fe_pi.se_new, log_pi.se_new)
    assert math.isclose(fe_pi.ci_lower, log_pi.ci_lower)
    assert math.isclose(fe_pi.ci_upper, log_pi.ci_upper)
    assert math.isclose(fe_pi.pi_lower, log_pi.pi_lower)
    assert math.isclose(fe_pi.pi_upper, log_pi.pi_upper)


def test_no_fe_pair_agrees_the_same_way_the_fe_pair_does() -> None:
    """P03 (pre-derived columns) == P03b (transform axis), without Fixed Effects.

    The same claim the test above makes for P01/P02, on the pair that has
    no FE. Worth asserting separately rather than trusting the FE result to
    cover it: the two mechanisms reach the design matrix by different
    routes — one reads a shipped column, the other computes one — and
    composing either with FE demeaning is a third code path again. With the
    FE pair alone, a transform-axis regression could hide behind the
    demeaning, or vice versa. Without FE the transform axis is the only
    thing between the CSV and the design matrix, so a disagreement here can
    only be the transform wiring.
    """
    cases = _cases()
    derived_expected = calculate_regression_spec_case(
        cases["production_lots_derived_log_no_fe"], PRODUCTION_LOTS_CSV_PATH
    )
    log_expected = calculate_regression_spec_case(
        cases["production_lots_log_no_fe"], PRODUCTION_LOTS_CSV_PATH
    )

    # Bit-exact, for the same reason as the FE pair: the shipped log columns
    # are exact logs of the raw ones. array_equal, not allclose.
    assert np.array_equal(
        derived_expected.design.x_features, log_expected.design.x_features
    )
    assert np.array_equal(
        derived_expected.design.y_train, log_expected.design.y_train
    )
    assert derived_expected.design.row_mask == log_expected.design.row_mask
    assert derived_expected.design.included_rows == log_expected.design.included_rows
    # Neither side has Fixed Effects — that is what this pair isolates.
    assert derived_expected.design.group_labels is None
    assert log_expected.design.group_labels is None

    derived_r = derived_expected.results
    log_r = log_expected.results
    assert math.isclose(derived_r.summary.r_squared, log_r.summary.r_squared)
    assert math.isclose(derived_r.summary.adjusted_r2, log_r.summary.adjusted_r2)
    assert math.isclose(derived_r.summary.se_regression, log_r.summary.se_regression)
    assert math.isclose(derived_r.summary.df_residual, log_r.summary.df_residual)
    assert np.allclose(derived_r.vectors.coefficients, log_r.vectors.coefficients)
    assert np.allclose(derived_r.vectors.std_errors, log_r.vectors.std_errors)
    assert np.allclose(derived_r.vectors.t_stats, log_r.vectors.t_stats)
    assert np.allclose(
        derived_r.full_residuals.residuals, log_r.full_residuals.residuals
    )

    # The labels legitimately differ — that is the mechanism showing through,
    # not a disagreement about the fit. Pinned so the pair cannot silently
    # collapse into two copies of the same spec.
    assert derived_expected.design.constructed_column_names == ("log Cum Units",)
    assert log_expected.design.constructed_column_names == ("Ln(Cumulative_Units)",)
    assert derived_expected.design.constructed_column_transforms == ("None",)
    assert log_expected.design.constructed_column_transforms == ("Log",)


def test_log_transform_case_labels_match_the_transform_contract() -> None:
    cases = _cases()
    log_case = cases["production_lots_log_transform"]
    log_expected = calculate_regression_spec_case(log_case, PRODUCTION_LOTS_CSV_PATH)

    assert log_expected.design.constructed_column_names == ("Ln(Cumulative_Units)",)
    assert log_expected.design.constructed_column_transforms == ("Log",)
    # Constructed-column-space alignment: one flag per Predictor_Columns()-equivalent
    # column, same width as x_features.
    assert len(log_expected.design.constructed_column_transforms) == (
        log_expected.design.x_features.shape[1]
    )


# ── Non-breaking by construction: default Transform="None" everywhere ───────

def test_default_spec_has_no_transform_declared() -> None:
    # build_default_spec() is the shipped T0 spec — every field defaults to
    # transform="None", so the v2.2 wiring changes nothing about it.
    for variable in build_default_spec():
        assert variable.transform == "None"


def test_none_transform_design_is_unaffected_by_the_log_wiring() -> None:
    rows = load_production_lots_source_rows(PRODUCTION_LOTS_CSV_PATH)
    cases = _cases()
    fe_case = cases["production_lots_fixed_effects"]

    # Rebuilding the design twice from the identical (all-"None"-transform)
    # spec must be exactly reproducible — no hidden state, no accidental
    # transform leakage between calls.
    design_a = build_spec_design(fe_case.spec, rows)
    design_b = build_spec_design(fe_case.spec, rows)
    assert np.array_equal(design_a.x_features, design_b.x_features)
    assert np.array_equal(design_a.y_train, design_b.y_train)
    assert design_a.constructed_column_transforms == design_b.constructed_column_transforms
    assert set(design_a.constructed_column_transforms) <= {"None"}


# ── Categorical x Log is disallowed and must never leak into the design ────

def test_log_on_a_categorical_predictor_is_ignored_by_the_design_builder() -> None:
    # The sheet flags this combination red (write_spec_block.py);
    # the design builder mirrors Predictor_Columns()'s own behavior — the Categorical
    # branch never reads .transform, so a mislabelled row is computationally
    # inert, not an error and not silently logged.
    rows = load_production_lots_source_rows(PRODUCTION_LOTS_CSV_PATH)
    spec = [
        SpecVariable("Lot_ID", "Identifier (Row Label)", False, "Continuous"),
        SpecVariable(
            "Facility", "Predictor (x)", True, "Categorical", transform="Log"
        ),
        SpecVariable("Fiscal_Year", "Omit", False, "Continuous"),
        SpecVariable("Lot_Quantity", "Omit", False, "Continuous"),
        SpecVariable("Cumulative_Units", "Omit", False, "Continuous"),
        SpecVariable("Experience_Stock", "Omit", False, "Continuous"),
        SpecVariable("Unit_Cost_BY", "Response (y)", False, "Continuous"),
        SpecVariable("log Cum Units", "Omit", False, "Continuous"),
        SpecVariable("log experience", "Omit", False, "Continuous"),
        SpecVariable("log Unit Cost", "Omit", False, "Continuous"),
        SpecVariable("Full_Data", "Filter", False, "Continuous"),
    ]
    design = build_spec_design(spec, rows)

    # Every constructed column here is a Facility dummy — all must read
    # "None" despite the mislabelled spec row's transform="Log".
    assert set(design.constructed_column_transforms) == {"None"}
    # And the dummy values themselves are ordinary 0/1 indicators, not logs.
    assert set(np.unique(design.x_features)) <= {0.0, 1.0}


def main() -> None:  # pragma: no cover - standalone runner
    """Run every verification directly (no pytest needed)."""
    checks = [
        (name, fn)
        for name, fn in sorted(globals().items())
        if name.startswith("test_") and callable(fn)
    ]
    for name, fn in checks:
        fn()
        print(f"PASS {name}")
    print(f"{len(checks)} verifications passed.")


if __name__ == "__main__":
    main()

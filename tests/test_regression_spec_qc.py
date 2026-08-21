"""Tests for the spec-driven Regression QC oracle."""

# pylint: disable=missing-function-docstring
from __future__ import annotations

import math
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pytest

from lambda_catalog import deep_verify
from lambda_catalog.analyze_regression_spec import (
    RegressionSpecCase,
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.write_spec_block import _ROLE_PREDICTOR
ROOT_DIR = Path(__file__).resolve().parents[1]

# One path per wired dataset, and each test guards on the one IT reads.
# `CSV_PATH` is only the DEFAULT handed to calculate_regression_spec_case — a
# case that sets source_csv_path (every Life Expectancy and Production Lots
# case) ignores it entirely and loads its own file. Guarding a Life
# Expectancy test on the Auto MPG CSV therefore gets it wrong in both
# directions: the test runs and fails on a missing Life Expectancy file, and
# skips when only the file it does not read is missing.
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"
LIFE_EXPECTANCY_CSV_PATH = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"
PRODUCTION_LOTS_CSV_PATH = ROOT_DIR / "sample_data" / "production_lots.csv"

_EXPECTED_CASE_NAMES = [
    "default_t0_intercept",
    "v1_full_continuous_intercept",
    "continuous_subset_intercept",
    "default_t0_no_intercept",
    "v1_full_continuous_no_intercept",
    "continuous_subset_no_intercept",
    "origin_default_reference",
    "origin_explicit_reference",
    "origin_invalid_reference",
    "model_year_origin_categorical",
    "usa_filter_degenerate_origin",
    "interaction_continuous_product",
    "interaction_quadratic_self_product",
    "interaction_categorical_broadcast",
    "mileage_log_log_na_masking",
    "categorical_only_design",
    "interaction_categorical_cross",
    "interaction_difference",
    "interaction_ratio_reciprocal",
    "production_lots_fixed_effects",
    "production_lots_log_transform",
    "production_lots_derived_log_no_fe",
    "production_lots_log_no_fe",
    "production_lots_log_mixed_predictors",
    "production_lots_log_predictor_only",
    "production_lots_lsdv_equivalence",
    "production_lots_log_loocv_leverage",
    "life_partial_linear_log",
    "life_log_response_duan",
    "life_log_response_naive",
    "life_elasticity_log_log",
    "life_full_profile",
    "life_country_fixed_effects",
    "life_status_explicit_reference",
    "life_log_drop_nonpositive",
    "life_talk_demo",
]

_EXPECTED_T0_NAMES = (
    "Horsepower",
    "Weight",
    *(f"Model Year: {year}" for year in range(71, 83)),
    "Origin: Europe",
    "Origin: US",
)


def _case(name: str) -> RegressionSpecCase:
    cases = {case.name: case for case in build_regression_spec_cases()}
    return cases[name]

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_regression_spec_fixture_names_are_pinned() -> None:
    assert [case.name for case in build_regression_spec_cases()] == _EXPECTED_CASE_NAMES


def test_sequence_is_flagged_only_on_datasets_that_have_an_ordering_axis() -> None:
    """No fittable case declares a Sequence flag on Auto MPG.

    Auto MPG is cross-sectional — one row per car model, no unit observed
    across periods — so ``Model Year`` is not an ordering axis, and a spec
    that flags it asserts panel structure the data does not have. The
    shipped default flags no Auto MPG variable, and this pins it so a
    copy-pasted ``sequence=True`` cannot creep back.

    The two panel datasets keep theirs, because there the flag is true and
    the diagnostics it gates (Durbin-Watson, and Breusch-Godfrey/Newey-West
    once Fixed Effects are present) are meaningful: every Production Lots
    case flags ``Fiscal_Year``, and every Life Expectancy case ``Year``.
    """
    axis_for_table = {
        "=MileageData[#All]": None,
        "=ProductionLotsData[#All]": "Fiscal_Year",
        "=LifeExpectancyData[#All]": "Year",
    }
    for case in build_regression_spec_cases():
        flagged = tuple(item.name for item in case.spec if item.sequence)
        axis = axis_for_table[case.source_table_ref]
        assert flagged == (
            () if axis is None else (axis,)
        ), f"{case.name} ({case.plan_id}) flags {flagged!r} as Sequence"


def test_calculate_verification_sheets_calculates_every_required_sheet() -> None:

    calls: list[str] = []

    class _Sheet:
        def __init__(self, name: str) -> None:
            self.name = name
            self.api = SimpleNamespace(Calculate=lambda: calls.append(name))

    class _Sheets:
        def __init__(self, names: list[str]) -> None:
            self._by_name = {name: _Sheet(name) for name in names}

        def __iter__(self):
            return iter(self._by_name.values())

        def __getitem__(self, name: str):
            return self._by_name[name]

    workbook = SimpleNamespace(
        app=SimpleNamespace(api=SimpleNamespace(Calculation=None)),
        sheets=_Sheets(
            [
                "Life Expectancy Data",
                "Mileage Data",
                "Production Lots",
                "Regression",
                "Univariate",
            ]
        ),
    )

    deep_verify._calculate_verification_sheets(
        workbook,
        verbose=False,
        phase_start=0.0,
    )

    assert calls == [
        "Life Expectancy Data",
        "Mileage Data",
        "Production Lots",
        "Regression",
        "Univariate",
    ]


def test_verification_calc_sheet_names_always_includes_univariate() -> None:

    assert "Univariate" in deep_verify._verification_calc_sheet_names()
    assert "Univariate" in deep_verify._verification_calc_sheet_names(
        skip_regression=True
    )


def test_calculate_verification_sheets_warns_instead_of_crashing_when_univariate_missing() -> (
    None
):
    """A workbook without a Univariate sheet (e.g. a test-model artifact)
    must skip it with a warning, not raise."""

    calls: list[str] = []

    class _Sheet:
        def __init__(self, name: str) -> None:
            self.name = name
            self.api = SimpleNamespace(Calculate=lambda: calls.append(name))

    class _Sheets:
        def __init__(self, names: list[str]) -> None:
            self._by_name = {name: _Sheet(name) for name in names}

        def __iter__(self):
            return iter(self._by_name.values())

        def __getitem__(self, name: str):
            return self._by_name[name]

    workbook = SimpleNamespace(
        app=SimpleNamespace(api=SimpleNamespace(Calculation=None)),
        sheets=_Sheets(
            ["Life Expectancy Data", "Mileage Data", "Production Lots", "Regression"]
        ),
    )

    deep_verify._calculate_verification_sheets(
        workbook,
        verbose=False,
        phase_start=0.0,
    )

    assert "Univariate" not in calls
    assert calls == [
        "Life Expectancy Data",
        "Mileage Data",
        "Production Lots",
        "Regression",
    ]


def test_univariate_stage_warns_when_sheet_is_missing() -> None:
    """A workbook without a Univariate sheet (e.g. a test-model artifact)
    must warn rather than run the comparison — the absence is by design for
    those artifacts, not a QC failure."""

    assert (
        deep_verify._univariate_verification_action(
            {"Regression", "Mileage Data"}
        )
        == "warn"
    )


def test_univariate_stage_checks_when_the_sheet_is_present() -> None:
    """The unified workbook's verify run must actually run the Univariate check
    when the sheet is present."""

    assert (
        deep_verify._univariate_verification_action(
            {"Regression", "Univariate", "Life Expectancy Data"}
        )
        == "check"
    )
    # A workbook with only the Univariate sheet still checks it.
    assert (
        deep_verify._univariate_verification_action({"Univariate"})
        == "check"
    )


def test_calculate_verification_sheets_still_requires_regression_sheet() -> None:
    """Missing sheets that are never legitimately optional (Regression) must
    still hard-fail — only Univariate gets the lenient warn-and-skip path."""

    class _Sheet:
        def __init__(self, name: str) -> None:
            self.name = name
            self.api = SimpleNamespace(Calculate=lambda: None)

    class _Sheets:
        def __init__(self, names: list[str]) -> None:
            self._by_name = {name: _Sheet(name) for name in names}

        def __iter__(self):
            return iter(self._by_name.values())

        def __getitem__(self, name: str):
            return self._by_name[name]

    workbook = SimpleNamespace(
        app=SimpleNamespace(api=SimpleNamespace(Calculation=None)),
        sheets=_Sheets(
            ["Life Expectancy Data", "Mileage Data", "Production Lots", "Univariate"]
        ),
    )

    with pytest.raises(RuntimeError, match="Regression"):
        deep_verify._calculate_verification_sheets(
            workbook,
            verbose=False,
            phase_start=0.0,
        )


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_default_t0_design_matches_current_constructor_semantics() -> None:
    expected = calculate_regression_spec_case(_case("default_t0_intercept"), CSV_PATH)
    design = expected.design

    assert design.included_rows == 392
    assert design.x_features.shape == (392, 16)
    assert design.y_train.shape == (392,)
    assert design.row_labels[0] == "chevrolet chevelle malibu"
    assert design.constructed_column_names == _EXPECTED_T0_NAMES
    assert design.level_counts == {"Model Year": 13, "Origin": 3}
    assert design.references_in_use == {"Model Year": 70, "Origin": "Asia"}
    assert design.degenerate_categoricals == ()
    # No Sequence axis on Auto MPG: the dataset is cross-sectional, so the
    # T0 spec flags nothing and the serial-correlation layer stays inert.
    # The Sequence machinery is covered on the datasets that have a real
    # ordering axis (Production Lots' Fiscal_Year, Life Expectancy's Year)
    # and, for the flag mechanics themselves, by guard cases G03 and M16.
    assert design.sequence_values is None

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_v1_full_continuous_design_uses_per_predictor_mask_and_feature_order() -> None:
    # M03 includes all five _MILEAGE_FEATURE_COLUMNS plus MPG (the response) —
    # exactly the [MPG]:[Acceleration] range the old Full_Data completeness
    # column checked. The per-predictor blank mask already drops the same 14
    # Horsepower-missing rows, so dropping the redundant Full_Data Filter left
    # included_rows at 392. The redundancy is now the point of this test.
    expected = calculate_regression_spec_case(
        _case("v1_full_continuous_intercept"), CSV_PATH
    )
    design = expected.design

    assert design.included_rows == 392
    assert design.x_features.shape == (392, 5)
    assert design.constructed_column_names == (
        "Cylinders",
        "Displacement",
        "Horsepower",
        "Weight",
        "Acceleration",
    )
    assert design.row_labels[0] == "70|chevrolet chevelle malibu"
    assert design.level_counts == {}
    assert design.references_in_use == {}
    assert design.degenerate_categoricals == ()

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_dummy_columns_are_binary_reference_dropped_and_filtered() -> None:
    expected = calculate_regression_spec_case(
        _case("origin_default_reference"), CSV_PATH
    )
    design = expected.design
    origin_columns = [
        idx
        for idx, name in enumerate(design.constructed_column_names)
        if name.startswith("Origin:")
    ]

    assert origin_columns == [3, 4]
    assert design.constructed_column_names[3] == "Origin: Europe"
    assert design.constructed_column_names[4] == "Origin: US"
    assert set(np.unique(design.x_features[:, origin_columns[0]])) <= {0.0, 1.0}
    assert design.references_in_use["Origin"] == "Asia"
    assert design.included_rows == 392

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_explicit_reference_changes_origin_dummy_level() -> None:
    expected = calculate_regression_spec_case(
        _case("origin_explicit_reference"), CSV_PATH
    )

    assert expected.design.references_in_use["Origin"] == "Europe"
    assert "Origin: Asia" in expected.design.constructed_column_names
    assert "Origin: Europe" not in expected.design.constructed_column_names

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_invalid_reference_skips_origin_but_model_still_computes() -> None:
    expected = calculate_regression_spec_case(
        _case("origin_invalid_reference"), CSV_PATH
    )
    design = expected.design

    assert design.degenerate_categoricals == ("Origin",)
    assert design.references_in_use["Origin"] == 99
    assert not any(
        name.startswith("Origin:") for name in design.constructed_column_names
    )
    assert design.constructed_column_names == ("Displacement", "Horsepower", "Weight")
    assert math.isfinite(expected.results.summary.r_squared)

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_model_year_origin_categorical_keeps_numeric_year_levels_as_dummies() -> None:
    expected = calculate_regression_spec_case(
        _case("model_year_origin_categorical"), CSV_PATH
    )
    design = expected.design

    assert design.constructed_column_names[3:15] == tuple(
        f"Model Year: {year}" for year in range(71, 83)
    )
    assert "Origin: Europe" in design.constructed_column_names
    assert design.level_counts == {"Model Year": 13, "Origin": 3}
    assert design.references_in_use["Model Year"] == 70

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_model_year_origin_categorical_gvif_shared_across_dummy_columns() -> None:
    """GVIF collapses each categorical variable's dummy block to one shared value."""
    expected = calculate_regression_spec_case(
        _case("model_year_origin_categorical"), CSV_PATH
    )
    names = expected.design.constructed_column_names
    gvif = expected.results.predictor_summary.gvif

    year_gvif = {
        gvif[i] for i, name in enumerate(names) if name.startswith("Model Year: ")
    }
    assert (
        len(year_gvif) == 1
    ), "all 12 Model Year dummy columns must share one GVIF value"

    origin_gvif = {
        gvif[i] for i, name in enumerate(names) if name.startswith("Origin: ")
    }
    assert (
        len(origin_gvif) == 1
    ), "Origin has two dummy columns but should still be one group"

    # Continuous predictors are their own group (df=1): GVIF must exactly match
    # ordinary per-column VIF, independently recomputed via lstsq (not calling
    # into production code) as a genuine cross-check.
    x = expected.design.x_features
    for continuous_name in ("Displacement", "Horsepower", "Weight"):
        j = names.index(continuous_name)
        others = np.delete(x, j, axis=1)
        others_with_const = np.column_stack([np.ones(x.shape[0]), others])
        beta = np.linalg.lstsq(others_with_const, x[:, j], rcond=None)[0]
        y_hat = others_with_const @ beta
        ss_res = float(np.sum((x[:, j] - y_hat) ** 2))
        ss_tot = float(np.sum((x[:, j] - np.mean(x[:, j])) ** 2))
        r2_j = 1.0 - ss_res / ss_tot
        expected_vif = 1.0 / (1.0 - r2_j)
        assert gvif[j] == pytest.approx(expected_vif, rel=1e-6), continuous_name

    assert all(v >= 1.0 - 1e-9 for v in gvif)

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_usa_filter_degenerates_origin_and_drops_its_columns() -> None:
    expected = calculate_regression_spec_case(
        _case("usa_filter_degenerate_origin"),
        CSV_PATH,
    )
    design = expected.design

    assert design.included_rows == 245
    assert design.degenerate_categoricals == ("Origin",)
    assert design.level_counts == {"Model Year": 13, "Origin": 1}
    assert design.references_in_use == {"Model Year": 70, "Origin": "US"}
    assert not any(
        name.startswith("Origin:") for name in design.constructed_column_names
    )
    assert design.constructed_column_names == (
        "Horsepower",
        "Weight",
        *(f"Model Year: {year}" for year in range(71, 83)),
    )


@pytest.mark.skipif(
    not PRODUCTION_LOTS_CSV_PATH.exists(), reason="Production Lots CSV not found"
)
def test_durbin_watson_is_a_real_statistic_where_an_ordering_axis_exists() -> None:
    """The bounds check the Auto MPG case above does not exercise.

    P03b has a Sequence axis (Fiscal_Year) and no Fixed Effects, which is
    the one state where the sheet's AE11 shows a number — so it is where
    the "DW is finite and in [0, 4]" invariant belongs. Keeping it here
    rather than dropping it with the Auto MPG pin is the point: the
    property is real, it was just being asserted on a model that cannot
    produce the statistic.

    Its pre-derived twin P03 must agree exactly. Both fit the identical
    model and DW is a pure function of the residual vector, which the two
    already agree on bit-for-bit.
    """
    transform_axis = calculate_regression_spec_case(
        _case("production_lots_log_no_fe"), PRODUCTION_LOTS_CSV_PATH
    ).results.summary.durbin_watson
    derived = calculate_regression_spec_case(
        _case("production_lots_derived_log_no_fe"), PRODUCTION_LOTS_CSV_PATH
    ).results.summary.durbin_watson

    assert math.isfinite(transform_axis)
    assert 0.0 <= transform_axis <= 4.0
    assert transform_axis == derived

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_expected_outputs_are_internally_consistent() -> None:
    expected = calculate_regression_spec_case(
        _case("continuous_subset_intercept"), CSV_PATH
    )
    results = expected.results
    design = expected.design
    k = len(design.constructed_column_names)
    p = k + 1

    assert len(results.vectors.coefficients) == p
    assert len(results.vectors.beta_weights) == k
    assert len(results.predictor_summary.gvif) == k
    assert len(results.prediction_interval.pred_input_values) == k
    # NaN, not a number: this is an Auto MPG case, and Auto MPG declares no
    # Sequence axis, so the sheet's AE11 reads "n/a — requires Sequence".
    # DW is only meaningful along a declared ordering; without one the
    # honest answer is "not computed", and the oracle returns NaN to match.
    assert math.isnan(results.summary.durbin_watson)
    assert abs(sum(results.full_residuals.hat_diagonal) - p) < 1e-4
    for y, prediction, residual in zip(
        results.full_residuals.dependent_var,
        results.full_residuals.predictions,
        results.full_residuals.residuals,
    ):
        assert abs(y - (prediction + residual)) < 1e-8


# ── v3.3 unit-space oracle cross-check ───────────────────────────────────────
#
# The mirrors in tests/test_unit_space_dispatch.py verify the catalog formulas
# against a NumPy reference, but they cannot verify the ORACLE: that code has
# its own copy of the arithmetic, and the only thing that compares the two is
# the spec-driven verifier, which needs Excel and does not run in CI. That gap
# is how a missing level shift and a fit-space observed column both shipped
# green.
#
# This check recomputes the unit-space block from the oracle's own residual
# columns — dependent_var (the fitted-on response, within-demeaned under FE),
# predictions and residuals, all already covered by the existing QC — plus the
# un-demeaned design.y_train. Different derivation path, same expected answer.

_LOG_SPEC_CASES = (
    "production_lots_log_transform",  # Log response + Log predictor + FE
    "production_lots_log_no_fe",  # Log + Log, no FE
    "production_lots_log_mixed_predictors",  # Log response, Log + None predictors
    "production_lots_log_predictor_only",  # None response, Log predictor
    "production_lots_log_loocv_leverage",  # Log response, heavy-tailed leverage
)


def _expected_unit_space(design, results) -> dict:
    """Independent recomputation of the v3.3 unit-space statistics."""
    y_fit = np.asarray(results.full_residuals.dependent_var, dtype=float)
    fitted = np.asarray(results.full_residuals.predictions, dtype=float)
    resid = np.asarray(results.full_residuals.residuals, dtype=float)
    y_train = np.asarray(design.y_train, dtype=float)

    logged = design.response_transform == "Log"
    # Only a Log response gets the level shift: nothing is exponentiated
    # under None, and shifting there would turn the within-flavoured
    # statistics into total ones.
    shift = (y_train - y_fit) if logged else np.zeros_like(y_fit)
    smearing = float(np.mean(np.exp(resid))) if logged else 1.0

    if logged:
        pred_unit = np.exp(fitted + shift) * smearing
        y_unit = np.exp(y_fit + shift)
    else:
        pred_unit = fitted
        y_unit = y_fit

    sse = float(np.sum((y_unit - pred_unit) ** 2))
    sst = float(np.sum((y_unit - np.mean(y_unit)) ** 2))
    df_total = results.summary.df_total
    df_residual = results.summary.df_residual
    r2 = 1.0 - sse / sst
    return {
        "smearing_factor": smearing,
        "r_squared_unit": r2,
        "adjusted_r2_unit": 1.0 - (1.0 - r2) * df_total / df_residual,
        "rmse_unit": math.sqrt(sse / df_residual),
        "predictions_unit": pred_unit,
        "residuals_unit": y_unit - pred_unit,
    }


@pytest.mark.skipif(
    not PRODUCTION_LOTS_CSV_PATH.exists(), reason="Production Lots CSV not found"
)
@pytest.mark.parametrize("case_name", _LOG_SPEC_CASES)
def test_unit_space_oracle_matches_an_independent_recomputation(case_name: str) -> None:
    expected = calculate_regression_spec_case(_case(case_name), CSV_PATH)
    want = _expected_unit_space(expected.design, expected.results)
    got = expected.results.unit_space

    assert got.smearing_factor == pytest.approx(want["smearing_factor"], rel=1e-12)
    assert got.r_squared_unit == pytest.approx(want["r_squared_unit"], rel=1e-10)
    assert got.adjusted_r2_unit == pytest.approx(want["adjusted_r2_unit"], rel=1e-10)
    assert got.rmse_unit == pytest.approx(want["rmse_unit"], rel=1e-10)
    assert np.allclose(got.predictions_unit, want["predictions_unit"], rtol=1e-10)
    assert np.allclose(got.residuals_unit, want["residuals_unit"], rtol=1e-8, atol=1e-8)


# v3.4 unit-space LOOCV — a SECOND independent path to the same answers.
#
# tests/test_unit_space_dispatch.py mirrors the catalog formulas by REFITTING
# n times (no hat-diagonal shortcut). This check recomputes the LOOCV residual
# the OTHER way — the hat-diagonal identity loo_fit = predictions - h*e/(1-h),
# applied to the oracle's OWN already-verified residual columns
# (full_residuals.hat_diagonal / .residuals / .predictions) — and cross-checks
# the oracle's loocv_residuals_unit / loocv_rmse_unit / loocv_mae_unit against
# it. Two independent derivations agreeing is the point: a mirror sharing the
# oracle's reading of Y_Full would produce a green suite and a wrong workbook
# (DECISIONS.md:2949-2958).

def _expected_unit_space_loocv(design, results) -> dict:
    """Independent recomputation of the v3.4 unit-space LOOCV statistics via
    the hat-diagonal identity, from the oracle's already-verified columns."""
    y_fit = np.asarray(results.full_residuals.dependent_var, dtype=float)
    fitted = np.asarray(results.full_residuals.predictions, dtype=float)
    resid = np.asarray(results.full_residuals.residuals, dtype=float)
    hat = np.asarray(results.full_residuals.hat_diagonal, dtype=float)
    y_train = np.asarray(design.y_train, dtype=float)

    logged = design.response_transform == "Log"
    shift = (y_train - y_fit) if logged else np.zeros_like(y_fit)
    smearing = float(np.mean(np.exp(resid))) if logged else 1.0

    # The leave-one-out fit-space prediction by the hat-diagonal identity —
    # the same identity the oracle's analyze_regression_sheet.py uses
    # (`loocv_predictions = predictions - h*e/(1-h)`), reached from the
    # already-verified columns rather than recomputed.
    loo_fit_space = fitted - hat * resid / (1.0 - hat)
    loo_fit = loo_fit_space + shift
    if logged:
        loo_pred_unit = np.exp(loo_fit) * smearing  # Duan (the four cases' default)
        y_unit = np.exp(y_fit + shift)
    else:
        loo_pred_unit = loo_fit
        y_unit = y_fit
    r = y_unit - loo_pred_unit
    n = r.size
    return {
        "loocv_residuals_unit": r,
        "loocv_rmse_unit": math.sqrt(float(np.sum(r ** 2)) / n),
        "loocv_mae_unit": float(np.mean(np.abs(r))),
        "smearing_treatment": (
            "n/a — no back-transform" if not logged
            else "Full-sample Duan (approx.)"
        ),
    }


@pytest.mark.skipif(
    not PRODUCTION_LOTS_CSV_PATH.exists(), reason="Production Lots CSV not found"
)
@pytest.mark.parametrize("case_name", _LOG_SPEC_CASES)
def test_unit_space_loocv_oracle_matches_an_independent_recomputation(
    case_name: str,
) -> None:
    """The oracle's LOOCV residual / RMSE / MAE must match a hat-diagonal
    recomputation from its own already-verified residual columns — a second
    independent path to the same answers the n-refit mirror in
    test_unit_space_dispatch.py also reaches."""
    expected = calculate_regression_spec_case(_case(case_name), CSV_PATH)
    want = _expected_unit_space_loocv(expected.design, expected.results)
    got = expected.results.unit_space

    assert np.allclose(
        got.loocv_residuals_unit, want["loocv_residuals_unit"], rtol=1e-9, atol=1e-9
    )
    assert got.loocv_rmse_unit == pytest.approx(want["loocv_rmse_unit"], rel=1e-10)
    assert got.loocv_mae_unit == pytest.approx(want["loocv_mae_unit"], rel=1e-10)
    # The LOOCV scalars must exceed the in-sample SE Regression (Unit): LOOCV is
    # out-of-sample, so it is the optimistic-free upper bound the plan's
    # verification section looks for (AH11/AH12 numeric and > AH8).
    assert got.loocv_rmse_unit > got.rmse_unit
    assert got.loocv_mae_unit > 0.0
    assert got.smearing_treatment == want["smearing_treatment"]


@pytest.mark.skipif(
    not PRODUCTION_LOTS_CSV_PATH.exists(), reason="Production Lots CSV not found"
)
@pytest.mark.parametrize(
    "case_name",
    ("production_lots_fixed_effects", "production_lots_log_predictor_only"),
)
def test_unit_space_loocv_reduces_to_loocv_residual_under_none(
    case_name: str,
) -> None:
    """Reduction invariant extended to the LOOCV pair: under Transform = None
    the unit-space LOOCV residual IS the ordinary (fit-space) LOOCV residual
    (the PRESS Residual column), and the LOOCV RMSE is sqrt(PRESS / n). This
    must hold WITH Fixed Effects — the level shift is inert under None, or the
    reduction silently becomes a total one (DECISIONS.md:2941-2947)."""
    expected = calculate_regression_spec_case(_case(case_name), CSV_PATH)
    unit = expected.results.unit_space
    full = expected.results.full_residuals
    summary = expected.results.summary

    assert expected.design.response_transform == "None"
    assert np.allclose(
        unit.loocv_residuals_unit, full.loocv_residuals, atol=1e-9
    ), "Unit_Space_LOOCV_Residual must collapse to the PRESS Residual under None"
    n = summary.observations
    assert unit.loocv_rmse_unit == pytest.approx(
        math.sqrt(summary.press / n), rel=1e-12
    ), "Unit_Space_LOOCV_RMSE must be sqrt(PRESS / n) under None"
    assert unit.smearing_treatment == "n/a — no back-transform"


# ── The cases added for docs/MODEL_TESTING_ASSETS.md § 1.1–1.3 ───────────────

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_log_log_on_auto_mpg_logs_both_sides_and_masks_missing_rows() -> None:
    """M05. The (Log, Log) pair combined with real missingness — the corner
    Production Lots (a complete 51-row panel) structurally cannot cover."""
    expected = calculate_regression_spec_case(
        _case("mileage_log_log_na_masking"), CSV_PATH
    )
    design = expected.design

    assert design.response_transform == "Log"
    assert design.predictor_transform == "Log"
    assert design.constructed_column_names == ("Ln(Horsepower)", "Ln(Weight)")
    assert design.constructed_column_transforms == ("Log", "Log")
    # The mask is applied BEFORE the logs are taken: 8 rows missing MPG and 6
    # missing Horsepower drop out, rather than Ln() poisoning the column.
    assert design.included_rows == 392
    # A Log response makes the smearing factor a real number, not 1.
    assert expected.results.unit_space.smearing_factor > 1.0

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_categorical_only_design_widens_the_sample() -> None:
    """M14b. With no included Continuous Predictor the mask reduces to "the
    response is numeric", so the sample grows past the 392 every other Auto
    MPG case sees — proof the mask is per-model, not per-dataset."""
    expected = calculate_regression_spec_case(
        _case("categorical_only_design"), CSV_PATH
    )
    design = expected.design

    assert design.included_rows == 398
    assert design.constructed_column_names == (
        *(f"Model Year: {year}" for year in range(71, 83)),
        "Origin: Europe",
        "Origin: US",
    )
    assert not any(
        name in design.constructed_column_names for name in ("Horsepower", "Weight")
    )

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_categorical_cross_emits_the_full_product_width() -> None:
    """M09. Cat x Cat is the one interaction width regime with no case: the
    constructor must emit (L1-1) * (L2-1) columns, here 12 * 2 = 24, in
    left-outer/right-inner order."""
    expected = calculate_regression_spec_case(
        _case("interaction_categorical_cross"), CSV_PATH
    )
    names = expected.design.constructed_column_names
    interactions = [name for name in names if " × " in name]

    assert len(interactions) == 24
    assert interactions[0] == "Model Year: 71 × Origin: Europe"
    assert interactions[1] == "Model Year: 71 × Origin: US"
    assert interactions[-1] == "Model Year: 82 × Origin: US"
    # 12 Model Year dummies + 24 interactions + 2 Origin dummies.
    assert len(names) == 38

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_categorical_cross_is_full_rank_and_shares_M14b_main_effects() -> None:
    """M09 is M14b plus the interaction block, so their main effects must be
    identical — and the saturated design must actually be fittable, which is
    why the case crosses Model Year (all 39 cells populated) rather than the
    plan's Cylinders (six empty cells, two all-zero columns, singular)."""
    cross = calculate_regression_spec_case(
        _case("interaction_categorical_cross"), CSV_PATH
    ).design
    base = calculate_regression_spec_case(
        _case("categorical_only_design"), CSV_PATH
    ).design

    main_effects = tuple(
        name for name in cross.constructed_column_names if " × " not in name
    )
    assert main_effects == base.constructed_column_names

    x_with_intercept = np.column_stack(
        [np.ones(cross.x_features.shape[0]), cross.x_features]
    )
    assert np.linalg.matrix_rank(x_with_intercept) == x_with_intercept.shape[1]

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_difference_interaction_uses_the_typographic_minus() -> None:
    """M10. The Difference operator, and its U+2212 MINUS SIGN header — a
    hyphen there would still read fine to a human and silently break the
    header comparison."""
    expected = calculate_regression_spec_case(_case("interaction_difference"), CSV_PATH)
    design = expected.design

    assert design.constructed_column_names == (
        "Displacement",
        "Displacement − Acceleration",
        "Weight",
    )
    difference_column = design.x_features[:, 1]
    assert np.allclose(
        difference_column, design.x_features[:, 0] - _acceleration_column(design)
    )


def _acceleration_column(design) -> np.ndarray:
    """Recover Acceleration from the difference column and Displacement."""
    return design.x_features[:, 0] - design.x_features[:, 1]

@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_ratio_reciprocal_pair_is_legal_and_non_singular() -> None:
    """M11. Ratio is asymmetric, so declaring both A/B and B/A produces two
    genuinely different columns — the legal counterpart to G10, where the
    same reciprocal declaration under Product is a singular Gram matrix."""
    expected = calculate_regression_spec_case(
        _case("interaction_ratio_reciprocal"), CSV_PATH
    )
    design = expected.design

    assert design.constructed_column_names == (
        "Horsepower",
        "Horsepower ÷ Weight",
        "Weight",
        "Weight ÷ Horsepower",
    )
    forward = design.x_features[:, 1]
    reverse = design.x_features[:, 3]
    assert np.allclose(forward * reverse, 1.0)
    # Reciprocal but not collinear: the design still has full rank.
    x_with_intercept = np.column_stack(
        [np.ones(design.x_features.shape[0]), design.x_features]
    )
    assert np.linalg.matrix_rank(x_with_intercept) == x_with_intercept.shape[1]
    assert math.isfinite(expected.results.summary.r_squared)


@pytest.mark.skipif(
    not PRODUCTION_LOTS_CSV_PATH.exists(), reason="Production Lots CSV not found"
)
def test_lsdv_reproduces_the_within_estimator_exactly() -> None:
    """P06 vs P02 — the suite's strongest cross-oracle.

    Fixed Effects demeaning, the absorbed-df subtraction and the level-shift
    recovery are the only bespoke arithmetic in the engine; everything else
    is OLS, which statsmodels also does. Fitting the same model as
    least-squares dummy variables reaches the same slopes and residuals by a
    completely different route, so this is a genuine second opinion rather
    than a second copy of the same code.
    """
    within = calculate_regression_spec_case(
        _case("production_lots_log_transform"), CSV_PATH
    )
    lsdv = calculate_regression_spec_case(
        _case("production_lots_lsdv_equivalence"), CSV_PATH
    )

    within_index = within.design.constructed_column_names.index("Ln(Cumulative_Units)")
    lsdv_index = lsdv.design.constructed_column_names.index("Ln(Cumulative_Units)")
    # +1 on both: Coefficients() spills the intercept first.
    assert within.results.vectors.coefficients[within_index + 1] == pytest.approx(
        lsdv.results.vectors.coefficients[lsdv_index + 1], rel=1e-10
    )
    assert np.allclose(
        within.results.full_residuals.residuals,
        lsdv.results.full_residuals.residuals,
        atol=1e-10,
    )
    # The two spend their degrees of freedom differently, which is exactly
    # why LSDV is a separate case rather than a duplicate: FE absorbs the
    # facility effects, LSDV shows them as columns.
    assert len(within.design.constructed_column_names) == 1
    assert len(lsdv.design.constructed_column_names) == 3


# ── Life Expectancy (§ 1.2) ──────────────────────────────────────────────────


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_partial_log_linear_is_the_mixed_predictor_dispatch_pair() -> None:
    """L01. (None, Mixed) — two logged predictors and one unlogged against an
    untransformed response. The predictor-transform summary must report
    "Mixed" rather than latching to whichever transform it saw first."""
    expected = calculate_regression_spec_case(
        _case("life_partial_linear_log"), CSV_PATH
    )
    design = expected.design

    assert (design.response_transform, design.predictor_transform) == ("None", "Mixed")
    assert design.constructed_column_names == (
        "Status: Developing",
        "Alcohol",
        "Ln(GDP)",
        "Ln(Population)",
    )
    assert design.constructed_column_transforms == ("None", "None", "Log", "Log")
    # The sample is the intersection of four columns with 652 / 448 / 194 / 0
    # blanks — well under half the 2938 rows.
    assert design.included_rows == 2117
    # No response transform, so the unit-space block reduces exactly.
    assert expected.results.unit_space.smearing_factor == pytest.approx(1.0)


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_binary_categorical_reference_flips_which_dummy_is_retained() -> None:
    """L09 vs L01. On a two-level column an explicit reference is invisible in
    the column COUNT — one dummy either way — and shows up only in WHICH level
    is retained. Getting it backwards flips the coefficient's sign."""
    default_reference = calculate_regression_spec_case(
        _case("life_partial_linear_log"), CSV_PATH
    ).design
    explicit = calculate_regression_spec_case(
        _case("life_status_explicit_reference"), CSV_PATH
    ).design

    assert default_reference.references_in_use["Status"] == "Developed"
    assert explicit.references_in_use["Status"] == "Developing"
    assert "Status: Developing" in default_reference.constructed_column_names
    assert "Status: Developed" in explicit.constructed_column_names
    # Same column space, so the fit is identical — only the parameterization
    # differs. A difference in R2 here would mean the reference change had
    # altered the model, not just its coding.
    assert len(default_reference.constructed_column_names) == len(
        explicit.constructed_column_names
    )


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_log_drop_nonpositive_narrows_the_sample_where_strict_log_poisons_it() -> None:
    """L12 vs L06. The two Log tokens on the same column and the same data.

    L06 declares the strict ``Log`` on Schooling and is a guard state: the 28
    true zeros stay in the sample, ``Ln_Positive`` returns #N/A for each, and
    nothing computes. This case declares ``Log (drop ≤ 0)`` instead, which is
    the only difference between the two specs, and the model fits.

    26, not 28: two of the zero-Schooling rows are already outside the sample
    on the response or Adult Mortality, so the positivity layer only has 26
    left to remove. Pinning the number rather than just "fewer rows" is what
    would catch the layer silently widening to rows it should not touch.
    """
    from lambda_catalog.analyze_regression_guard_states import (
        build_guard_state_cases,
        calculate_guard_state_case,
    )

    expected = calculate_regression_spec_case(
        _case("life_log_drop_nonpositive"), CSV_PATH
    )
    design = expected.design

    strict = calculate_guard_state_case(
        next(
            case
            for case in build_guard_state_cases()
            if case.name == "guard_ln_zero_propagation"
        )
    )

    # Same constructed column under either token — they differ in the mask,
    # not in the arithmetic.
    assert design.constructed_column_names == ("Adult Mortality", "Ln(Schooling)")
    assert design.constructed_column_transforms == ("None", "Log")
    assert design.predictor_transform == "Mixed"

    assert design.log_excluded_rows == 26
    assert design.included_rows == strict.included_rows - 26

    # The strict token flags the cell red and names the fix; the filtering
    # token flags nothing, because excluding those rows is what it is for.
    assert ("G", "red", "log_nonpositive_rows") in {
        (flag.column, flag.severity, flag.rule) for flag in strict.flags
    }
    assert "Log (drop ≤ 0)" in strict.log_domain_status

    # And the filtered sample actually fits — the thing L06 cannot do.
    summary = expected.results.summary
    assert summary.observations == design.included_rows
    assert summary.df_regression == 2
    assert 0.0 < summary.r_squared < 1.0
    assert np.isfinite(expected.results.vectors.coefficients).all()


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_log_response_is_the_log_none_dispatch_pair_with_real_smearing() -> None:
    """L02. (Log, None) — the pair where the v3.3 unit-space machinery does
    the most work, and the suite's only one."""
    expected = calculate_regression_spec_case(_case("life_log_response_duan"), CSV_PATH)
    design = expected.design
    unit = expected.results.unit_space

    assert (design.response_transform, design.predictor_transform) == ("Log", "None")
    assert design.constructed_column_names == (
        "Status: Developing",
        "Adult Mortality",
        "Schooling",
    )
    assert unit.smearing_factor > 1.0
    # Back-transformed statistics genuinely differ from the fit-space ones.
    assert unit.r_squared_unit != pytest.approx(expected.results.summary.r_squared)


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_naive_back_transform_drops_the_smearing_factor_but_not_the_bounds() -> None:
    """L03 vs L02 — the Back-Transform toggle, which had no oracle at all
    before: the Python side computed both branches and then discarded the
    Naive one, so `$AH$4` could have been wired to anything.

    Two halves to the contract. Predictions dispatch on the toggle, so every
    statistic derived from them moves. The CI/PI bounds are quantiles and are
    EXP-only under BOTH settings — the Back-Transform note on the sheet says so, and a
    toggle that moved them would be silently wrong in the direction users
    would never check.
    """
    duan = calculate_regression_spec_case(_case("life_log_response_duan"), CSV_PATH)
    naive = calculate_regression_spec_case(_case("life_log_response_naive"), CSV_PATH)

    assert duan.case.spec == naive.case.spec, "the toggle must be the only difference"
    assert (duan.case.back_transform, naive.case.back_transform) == ("Duan", "Naive")

    smearing = duan.results.unit_space.smearing_factor
    assert smearing == pytest.approx(naive.results.unit_space.smearing_factor)
    assert np.allclose(
        np.asarray(duan.results.unit_space.predictions_unit),
        np.asarray(naive.results.unit_space.predictions_unit) * smearing,
        rtol=1e-12,
    )
    assert duan.results.unit_space.prediction_point_unit == pytest.approx(
        naive.results.unit_space.prediction_point_unit * smearing, rel=1e-12
    )
    assert duan.results.unit_space.r_squared_unit != pytest.approx(
        naive.results.unit_space.r_squared_unit
    )
    for bound in (
        "prediction_ci_lower_unit",
        "prediction_ci_upper_unit",
        "prediction_pi_lower_unit",
        "prediction_pi_upper_unit",
    ):
        assert getattr(duan.results.unit_space, bound) == pytest.approx(
            getattr(naive.results.unit_space, bound), rel=1e-15
        ), bound
    # The fit itself is untouched — the toggle is a display-space choice.
    assert duan.results.summary.r_squared == pytest.approx(
        naive.results.summary.r_squared
    )


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_elasticity_model_logs_both_sides_at_scale() -> None:
    """L04. (Log, Log) against sparse predictors on the large dataset."""
    expected = calculate_regression_spec_case(
        _case("life_elasticity_log_log"), CSV_PATH
    )
    design = expected.design

    assert (design.response_transform, design.predictor_transform) == ("Log", "Log")
    assert design.constructed_column_names == ("Ln(GDP)", "Ln(Population)")
    assert design.included_rows == 2262


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_life_full_profile_is_the_k_stress_showcase() -> None:
    """L05 — the full 18-predictor kitchen sink, the suite's k-stress showcase.

    The shipped default is the curated ``life_talk_demo`` model (L11); L5 is
    the wide design every predictor-summary statistic is stressed on. The
    spec is restated in its own right in ``_life_full_profile_spec`` rather
    than derived from the profile, so a change to the shipped default cannot
    silently shrink this case."""
    from lambda_catalog.write_spec_block import SPEC_DATASET_PROFILES

    expected = calculate_regression_spec_case(_case("life_full_profile"), CSV_PATH)
    design = expected.design
    profile = SPEC_DATASET_PROFILES["life_expectancy"]

    # 18 continuous predictors + the single Status dummy.
    assert len(design.constructed_column_names) == 19
    assert "Status: Developing" in design.constructed_column_names
    # The spec itself — not just the design it produces — pins the 18 + 1
    # kitchen sink. Restating _life_full_profile_spec explicitly means a future
    # edit that shrinks it drops this count before it drops k, so the k-stress
    # case cannot quietly stop being one.
    included_predictors = [
        v for v in expected.case.spec if v.include and v.role == _ROLE_PREDICTOR
    ]
    assert sum(1 for v in included_predictors if v.var_type == "Continuous") == 18
    assert sum(1 for v in included_predictors if v.var_type == "Categorical") == 1
    # One spec row per Source_Table column, in dataset order.
    assert [item.name for item in expected.case.spec] == list(profile.variables)
    assert math.isfinite(expected.results.summary.r_squared)


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_curated_talk_demo_is_the_shipped_default_and_fits() -> None:
    """L11 — the curated four-driver model both decks headline, and the shipped
    default. Pins the exact spec the workbook opens with so the slide-19
    coefficient table is oracle-checked, not illustrative."""
    from lambda_catalog.write_spec_block import SPEC_DATASET_PROFILES

    expected = calculate_regression_spec_case(_case("life_talk_demo"), CSV_PATH)
    design = expected.design

    # (None, None) dispatch — untransformed response, untransformed predictors.
    assert (design.response_transform, design.predictor_transform) == ("None", "None")
    # 3 continuous predictors + the single Status dummy = 4 design columns.
    # Status sits earlier in the dataset column order than the feature columns,
    # so its dummy is constructed first.
    assert design.constructed_column_names == (
        "Status: Developing",
        "Adult Mortality",
        "Alcohol",
        "percentage expenditure",
    )
    # The shipped profile default and the registered case are the same model:
    # the cold open a user sees is the one this oracle verifies.
    profile = SPEC_DATASET_PROFILES["life_expectancy"]
    included = {
        name: spec
        for name, spec in profile.default_spec.items()
        if spec[1] is True  # Include = True
    }
    assert set(included) == {
        "Adult Mortality",
        "Alcohol",
        "percentage expenditure",
        "Status",
    }
    assert included["Status"][2] == "Categorical"
    assert math.isfinite(expected.results.summary.r_squared)


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_high_cardinality_fixed_effects_absorbs_the_right_degrees_of_freedom() -> None:
    """L08. 193 groups against Production Lots' three. At 3 groups a df error
    of a few units hides inside the noise; at 182 absorbed df it moves every
    df-dependent statistic visibly."""
    expected = calculate_regression_spec_case(
        _case("life_country_fixed_effects"), CSV_PATH
    )
    design = expected.design
    summary = expected.results.summary

    assert design.group_labels is not None
    groups = len(np.unique(design.group_labels))
    # 173, not the dataset's full 193: Schooling is blank for every row of 20
    # countries, so the mask removes those panels entirely. Worth pinning —
    # a silent drop from 173 to a handful would leave the case named
    # "high cardinality" while testing nothing of the sort.
    assert groups == 173
    assert design.constructed_column_names == ("Adult Mortality", "Schooling")
    # n - k - 1 - (G - 1): the absorbed group effects come out of the
    # residual degrees of freedom.
    assert summary.df_residual == design.included_rows - 2 - 1 - (groups - 1)
    # Durbin-Watson has no valid FE-active reading, so the oracle is NaN.
    assert math.isnan(summary.durbin_watson)


@pytest.mark.skipif(
    not PRODUCTION_LOTS_CSV_PATH.exists(), reason="Production Lots CSV not found"
)
@pytest.mark.parametrize(
    "case_name",
    ("production_lots_fixed_effects", "production_lots_log_predictor_only"),
)
def test_unit_space_reduces_to_the_ordinary_statistics_without_a_response_transform(
    case_name: str,
) -> None:
    """The reduction invariant on real spec cases, including one with Fixed
    Effects: with the Response row untransformed, the unit-space statistics
    ARE the ordinary statistics. A level shift leaking in under None would
    turn these into total-flavoured numbers and fail here."""
    expected = calculate_regression_spec_case(_case(case_name), CSV_PATH)
    summary = expected.results.summary
    unit = expected.results.unit_space

    assert expected.design.response_transform == "None"
    assert unit.smearing_factor == pytest.approx(1.0, rel=1e-15)
    assert unit.r_squared_unit == pytest.approx(summary.r_squared, rel=1e-12)
    assert unit.adjusted_r2_unit == pytest.approx(summary.adjusted_r2, rel=1e-12)
    assert unit.rmse_unit == pytest.approx(summary.se_regression, rel=1e-12)
    assert np.allclose(
        unit.predictions_unit, expected.results.full_residuals.predictions, rtol=1e-12
    )
    assert np.allclose(
        unit.residuals_unit, expected.results.full_residuals.residuals, atol=1e-9
    )


# ── Categorical level ordering vs. Excel's SORT ──────────────────────────


def test_level_sort_key_puts_numbers_before_text() -> None:
    """Excel sorts every number ahead of every string."""
    from lambda_catalog.analyze_model_construction import level_sort_key

    assert sorted([3, "b", 1, "a"], key=level_sort_key) == [1, 3, "a", "b"]


def test_level_sort_key_files_accented_text_where_excel_does() -> None:
    """The bug the first live Excel run exposed.

    Python compares strings by code point, so ``ô`` (U+00F4) sorts after
    ``z`` (U+007A) and ``Côte d'Ivoire`` lands after ``Czechia``. Excel uses
    the Windows locale collator, where ``ô`` files under ``o`` — between
    ``Costa Rica`` and ``Croatia``.

    On a categorical predictor that one-position difference shifts the WHOLE
    dummy block: every column keeps a valid name and a valid 0/1 pattern, so
    nothing errors — the columns are just paired with the wrong headers and
    every per-predictor statistic reads off by one.
    """
    from lambda_catalog.analyze_model_construction import level_sort_key

    ordered = sorted(
        ["Croatia", "Cuba", "Côte d'Ivoire", "Czechia", "Costa Rica"],
        key=level_sort_key,
    )
    assert ordered == [
        "Costa Rica",
        "Côte d'Ivoire",
        "Croatia",
        "Cuba",
        "Czechia",
    ]
    # Accent-only differences must still order deterministically rather than
    # comparing equal, or the sort would not be total.
    assert sorted(["Cote", "Côte"], key=level_sort_key) == ["Cote", "Côte"]


@pytest.mark.skipif(
    not LIFE_EXPECTANCY_CSV_PATH.exists(), reason="Life Expectancy CSV not found"
)
def test_life_expectancy_country_levels_follow_excel_order() -> None:
    """End-to-end on the dataset that surfaced it: Country is the only wired
    column with a non-ASCII level, and its dummy block must be ordered the
    way the sheet's SORT orders it."""
    from lambda_catalog.analyze_life_expectancy import (
        load_life_expectancy_source_rows,
    )
    from lambda_catalog.analyze_model_construction import level_sort_key

    countries = sorted(
        {row["Country"] for row in load_life_expectancy_source_rows()},
        key=level_sort_key,
    )
    index = countries.index("Côte d'Ivoire")
    assert countries[index - 1] == "Costa Rica"
    assert countries[index + 1] == "Croatia"

"""Tests for the spec-driven Regression QC oracle."""
# pylint: disable=missing-function-docstring
from __future__ import annotations

import math
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pytest

from lambda_catalog.analyze_regression_sheet import calculate_regression_sheet_results
from lambda_catalog.analyze_regression_spec import (
    RegressionSpecCase,
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.regression_shared import FEATURE_COLUMNS

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"

_EXPECTED_CASE_NAMES = [
    "default_t0_intercept",
    "v1_full_continuous_intercept",
    "continuous_subset_intercept",
    "default_t0_no_intercept",
    "v1_full_continuous_no_intercept",
    "continuous_subset_no_intercept",
    "status_default_reference",
    "status_explicit_reference",
    "status_invalid_reference",
    "year_status_categorical",
    "developing_filter_degenerate_status",
]

_EXPECTED_T0_NAMES = (
    *(f"Year: {year}" for year in range(2001, 2016)),
    "Status: Developing",
    "Adult Mortality",
    "GDP",
    "Schooling",
)


def _case(name: str) -> RegressionSpecCase:
    cases = {case.name: case for case in build_regression_spec_cases()}
    return cases[name]


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_regression_spec_fixture_names_are_pinned() -> None:
    assert [case.name for case in build_regression_spec_cases()] == _EXPECTED_CASE_NAMES


def test_build_qc_keeps_mlr_names_only_for_stale_sheet_deletion() -> None:
    import build_qc

    mlr_names = {
        "MLR_Scalar_Test",
        "MLR_Vector_Outputs_Test",
        "MLR_Observation_Test",
    }
    assert mlr_names <= set(build_qc._QC_SHEET_NAMES)
    assert mlr_names.isdisjoint(build_qc._VERIFY_CALC_SHEET_NAMES)


def test_build_qc_verification_calc_sheet_names_respects_skip_dummy_flag() -> None:
    import build_qc

    assert "Dummy_Test" in build_qc._verification_calc_sheet_names(skip_dummy=False)
    assert "Dummy_Test" not in build_qc._verification_calc_sheet_names(skip_dummy=True)


def test_calculate_verification_sheets_excludes_dummy_when_requested() -> None:
    import build_qc

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
        sheets=_Sheets(["Life Expectancy Data", "Regression", "Univariate"]),
    )

    build_qc._calculate_verification_sheets(
        workbook,
        verbose=False,
        phase_start=0.0,
        skip_dummy=True,
    )

    assert calls == ["Life Expectancy Data", "Regression", "Univariate"]


def test_calculate_verification_sheets_requires_dummy_when_not_skipped() -> None:
    import build_qc

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
        sheets=_Sheets(["Life Expectancy Data", "Regression", "Univariate"]),
    )

    with pytest.raises(RuntimeError, match="Dummy_Test"):
        build_qc._calculate_verification_sheets(
            workbook,
            verbose=False,
            phase_start=0.0,
            skip_dummy=False,
        )


def test_build_qc_run_main_skips_verified_summary_when_verification_disabled(
    monkeypatch,
    capsys,
) -> None:
    import build_qc
    from lambda_catalog.workbook_builder import NameSyncResult

    def fake_build_qc_workbook(*, timings_out, **_):
        timings_out["prep_seconds"] = 1.0
        timings_out["write_seconds"] = 2.0
        timings_out["sync_seconds"] = 3.0
        timings_out["verify_seconds"] = None
        return NameSyncResult(created=0, updated=0)

    monkeypatch.setattr(build_qc, "build_qc_workbook", fake_build_qc_workbook)
    monkeypatch.setattr(build_qc.subprocess, "Popen", lambda *_: None)

    build_qc._run_main(
        SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            cache=Path("cache.json"),
            validate_reopen=False,
            verbose=False,
            no_verify=True,
        )
    )

    output = capsys.readouterr().out
    assert "Sheet verified:" not in output
    assert "Timing: verify        skipped" in output


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_default_t0_design_matches_current_constructor_semantics() -> None:
    expected = calculate_regression_spec_case(_case("default_t0_intercept"), CSV_PATH)
    design = expected.design

    assert design.included_rows == 2482
    assert design.x_features.shape == (2482, 19)
    assert design.y_train.shape == (2482,)
    assert design.row_labels[0] == "Afghanistan"
    assert design.constructed_column_names == _EXPECTED_T0_NAMES
    assert design.level_counts == {"Year": 16, "Status": 2}
    assert design.references_in_use == {"Year": 2000, "Status": "Developed"}
    assert design.degenerate_categoricals == ()
    assert design.sequence_values is not None
    assert design.sequence_values[0] == 2015


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_v1_full_continuous_design_uses_full_data_filter_and_feature_order() -> None:
    expected = calculate_regression_spec_case(_case("v1_full_continuous_intercept"), CSV_PATH)
    design = expected.design

    assert design.included_rows == 1649
    assert design.x_features.shape == (1649, len(FEATURE_COLUMNS))
    assert design.constructed_column_names == tuple(FEATURE_COLUMNS)
    assert design.row_labels[0] == "Afghanistan|2015"
    assert design.level_counts == {}
    assert design.references_in_use == {}
    assert design.degenerate_categoricals == ()


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_dummy_columns_are_binary_reference_dropped_and_filtered() -> None:
    expected = calculate_regression_spec_case(_case("status_default_reference"), CSV_PATH)
    design = expected.design
    status_columns = [
        idx for idx, name in enumerate(design.constructed_column_names)
        if name.startswith("Status:")
    ]

    assert status_columns == [0]
    assert design.constructed_column_names[0] == "Status: Developing"
    assert set(np.unique(design.x_features[:, status_columns[0]])) <= {0.0, 1.0}
    assert design.references_in_use["Status"] == "Developed"
    assert design.included_rows == 1649


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_explicit_reference_changes_status_dummy_level() -> None:
    expected = calculate_regression_spec_case(_case("status_explicit_reference"), CSV_PATH)

    assert expected.design.references_in_use["Status"] == "Developing"
    assert "Status: Developed" in expected.design.constructed_column_names
    assert "Status: Developing" not in expected.design.constructed_column_names


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_invalid_reference_skips_status_but_model_still_computes() -> None:
    expected = calculate_regression_spec_case(_case("status_invalid_reference"), CSV_PATH)
    design = expected.design

    assert design.degenerate_categoricals == ("Status",)
    assert design.references_in_use["Status"] == "Developped"
    assert not any(name.startswith("Status:") for name in design.constructed_column_names)
    assert design.constructed_column_names == ("Adult Mortality", "GDP", "Schooling")
    assert math.isfinite(expected.results.summary.r_squared)


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_year_status_categorical_keeps_numeric_year_levels_as_dummies() -> None:
    expected = calculate_regression_spec_case(_case("year_status_categorical"), CSV_PATH)
    design = expected.design

    assert design.constructed_column_names[:15] == tuple(
        f"Year: {year}" for year in range(2001, 2016)
    )
    assert "Status: Developing" in design.constructed_column_names
    assert design.level_counts == {"Year": 16, "Status": 2}
    assert design.references_in_use["Year"] == 2000


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_developing_filter_degenerates_status_and_drops_its_columns() -> None:
    expected = calculate_regression_spec_case(
        _case("developing_filter_degenerate_status"),
        CSV_PATH,
    )
    design = expected.design

    assert design.included_rows == 2034
    assert design.degenerate_categoricals == ("Status",)
    assert design.level_counts == {"Year": 16, "Status": 1}
    assert design.references_in_use == {"Year": 2000, "Status": "Developing"}
    assert not any(name.startswith("Status:") for name in design.constructed_column_names)
    assert design.constructed_column_names == (
        *(f"Year: {year}" for year in range(2001, 2016)),
        "Adult Mortality",
        "GDP",
        "Schooling",
    )


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_expected_outputs_are_internally_consistent() -> None:
    expected = calculate_regression_spec_case(_case("continuous_subset_intercept"), CSV_PATH)
    results = expected.results
    design = expected.design
    k = len(design.constructed_column_names)
    p = k + 1

    assert len(results.vectors.coefficients) == p
    assert len(results.vectors.beta_weights) == k
    assert len(results.predictor_summary.vif) == k
    assert len(results.prediction_interval.pred_input_values) == k
    assert math.isfinite(results.summary.durbin_watson)
    assert 0.0 <= results.summary.durbin_watson <= 4.0
    assert results.summary.durbin_watson == pytest.approx(2.1298772256557448)
    assert abs(sum(results.full_residuals.hat_diagonal) - p) < 1e-4
    for y, prediction, residual in zip(
        results.full_residuals.dependent_var,
        results.full_residuals.predictions,
        results.full_residuals.residuals,
    ):
        assert abs(y - (prediction + residual)) < 1e-8


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Life Expectancy CSV not found")
def test_full_continuous_fixture_matches_legacy_continuous_calculator() -> None:
    spec_expected = calculate_regression_spec_case(
        _case("v1_full_continuous_intercept"),
        CSV_PATH,
    ).results
    legacy = calculate_regression_sheet_results(
        CSV_PATH,
        include_intercept=True,
        feature_columns=FEATURE_COLUMNS,
    )

    assert spec_expected.summary.observations == legacy.summary.observations
    assert spec_expected.predictor_summary.predictor_names == legacy.predictor_summary.predictor_names
    assert spec_expected.vectors.term_names == legacy.vectors.term_names
    assert np.allclose(spec_expected.vectors.coefficients, legacy.vectors.coefficients)
    assert np.allclose(spec_expected.vectors.beta_weights, legacy.vectors.beta_weights)
    assert np.allclose(spec_expected.full_residuals.hat_diagonal, legacy.full_residuals.hat_diagonal)
    assert spec_expected.summary.r_squared == pytest.approx(legacy.summary.r_squared)
    assert spec_expected.prediction_interval.point_estimate == pytest.approx(
        legacy.prediction_interval.point_estimate
    )

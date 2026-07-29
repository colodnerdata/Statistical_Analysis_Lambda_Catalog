"""Tests for the spec-driven Regression QC oracle."""
# pylint: disable=missing-function-docstring
from __future__ import annotations

import math
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pytest

from lambda_catalog.analyze_regression_spec import (
    RegressionSpecCase,
    build_regression_spec_cases,
    calculate_regression_spec_case,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"

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
    "production_lots_fixed_effects",
    "production_lots_log_transform",
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
        sheets=_Sheets(
            ["Life Expectancy Data", "Mileage Data", "Production Lots", "Regression", "Univariate"]
        ),
    )

    build_qc._calculate_verification_sheets(
        workbook,
        verbose=False,
        phase_start=0.0,
        skip_dummy=True,
    )

    assert calls == [
        "Life Expectancy Data", "Mileage Data", "Production Lots", "Regression", "Univariate",
    ]


def test_build_qc_verification_calc_sheet_names_respects_skip_univariate_flag() -> None:
    import build_qc

    assert "Univariate" in build_qc._verification_calc_sheet_names(
        skip_dummy=True, skip_univariate=False
    )
    assert "Univariate" not in build_qc._verification_calc_sheet_names(
        skip_dummy=True, skip_univariate=True
    )


def test_calculate_verification_sheets_warns_instead_of_crashing_when_univariate_missing() -> None:
    """A workbook built with --skip-univariate never gets a Univariate sheet.
    Verification must skip it with a warning, not raise, regardless of
    whether skip_univariate was explicitly passed for this call."""
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
        sheets=_Sheets(["Life Expectancy Data", "Mileage Data", "Production Lots", "Regression"]),
    )

    # skip_univariate not passed (defaults False) — the sheet is simply
    # absent from this workbook, which must still be handled gracefully.
    build_qc._calculate_verification_sheets(
        workbook,
        verbose=False,
        phase_start=0.0,
        skip_dummy=True,
    )

    assert "Univariate" not in calls
    assert calls == ["Life Expectancy Data", "Mileage Data", "Production Lots", "Regression"]


def test_calculate_verification_sheets_still_requires_regression_sheet() -> None:
    """Missing sheets that are never legitimately optional (Regression) must
    still hard-fail — only Univariate gets the lenient warn-and-skip path."""
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
        sheets=_Sheets(["Life Expectancy Data", "Mileage Data", "Production Lots", "Univariate"]),
    )

    with pytest.raises(RuntimeError, match="Regression"):
        build_qc._calculate_verification_sheets(
            workbook,
            verbose=False,
            phase_start=0.0,
            skip_dummy=True,
        )


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
        sheets=_Sheets(
            ["Life Expectancy Data", "Mileage Data", "Production Lots", "Regression", "Univariate"]
        ),
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
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            cache=Path("cache.json"),
            validate_reopen=False,
            verbose=False,
            no_verify=True,
        )
    )

    output = capsys.readouterr().out
    assert "Sheet verified:" not in output
    assert "Timing: verify        skipped" in output


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
    assert design.sequence_values is not None
    assert design.sequence_values[0] == 70


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_v1_full_continuous_design_uses_full_data_filter_and_feature_order() -> None:
    expected = calculate_regression_spec_case(_case("v1_full_continuous_intercept"), CSV_PATH)
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
    expected = calculate_regression_spec_case(_case("origin_default_reference"), CSV_PATH)
    design = expected.design
    origin_columns = [
        idx for idx, name in enumerate(design.constructed_column_names)
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
    expected = calculate_regression_spec_case(_case("origin_explicit_reference"), CSV_PATH)

    assert expected.design.references_in_use["Origin"] == "Europe"
    assert "Origin: Asia" in expected.design.constructed_column_names
    assert "Origin: Europe" not in expected.design.constructed_column_names


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_invalid_reference_skips_origin_but_model_still_computes() -> None:
    expected = calculate_regression_spec_case(_case("origin_invalid_reference"), CSV_PATH)
    design = expected.design

    assert design.degenerate_categoricals == ("Origin",)
    assert design.references_in_use["Origin"] == 99
    assert not any(name.startswith("Origin:") for name in design.constructed_column_names)
    assert design.constructed_column_names == ("Displacement", "Horsepower", "Weight")
    assert math.isfinite(expected.results.summary.r_squared)


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_model_year_origin_categorical_keeps_numeric_year_levels_as_dummies() -> None:
    expected = calculate_regression_spec_case(_case("model_year_origin_categorical"), CSV_PATH)
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
    expected = calculate_regression_spec_case(_case("model_year_origin_categorical"), CSV_PATH)
    names = expected.design.constructed_column_names
    gvif = expected.results.predictor_summary.gvif

    year_gvif = {gvif[i] for i, name in enumerate(names) if name.startswith("Model Year: ")}
    assert len(year_gvif) == 1, "all 12 Model Year dummy columns must share one GVIF value"

    origin_gvif = {gvif[i] for i, name in enumerate(names) if name.startswith("Origin: ")}
    assert len(origin_gvif) == 1, "Origin has two dummy columns but should still be one group"

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
    assert not any(name.startswith("Origin:") for name in design.constructed_column_names)
    assert design.constructed_column_names == (
        "Horsepower",
        "Weight",
        *(f"Model Year: {year}" for year in range(71, 83)),
    )


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_expected_outputs_are_internally_consistent() -> None:
    expected = calculate_regression_spec_case(_case("continuous_subset_intercept"), CSV_PATH)
    results = expected.results
    design = expected.design
    k = len(design.constructed_column_names)
    p = k + 1

    assert len(results.vectors.coefficients) == p
    assert len(results.vectors.beta_weights) == k
    assert len(results.predictor_summary.gvif) == k
    assert len(results.prediction_interval.pred_input_values) == k
    assert math.isfinite(results.summary.durbin_watson)
    assert 0.0 <= results.summary.durbin_watson <= 4.0
    assert results.summary.durbin_watson == pytest.approx(0.8587513374458717)
    assert abs(sum(results.full_residuals.hat_diagonal) - p) < 1e-4
    for y, prediction, residual in zip(
        results.full_residuals.dependent_var,
        results.full_residuals.predictions,
        results.full_residuals.residuals,
    ):
        assert abs(y - (prediction + residual)) < 1e-8

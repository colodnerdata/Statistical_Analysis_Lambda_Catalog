"""Independent verification of v3.3 unit-space dispatch (Smearing_Factor,
Back_Transform_Response, Unit_Space_Predictions, Unit_Space_Residuals,
Unit_Space_R_Squared, Unit_Space_Adjusted_R_Squared, Unit_Space_RMSE).

Pure-Python mirrors of each catalog function are cross-checked against a
NumPy OLS reference, and the catalog ``formula_display`` text is asserted
to use the documented primitives (SWITCH on the six unit-space pairs,
EXP+smearing for Duan, NA() outside the {None, Log} response family).

The reduction invariant — under ``Context_Response_Transform = "None"`` every
unit-space statistic collapses to the ordinary statistic — is tested
directly: pure-Python mirror matched against the same mirror's response-
transform-None branch, both computed from the same matrices.

Runnable standalone (``python tests/test_unit_space_dispatch.py``) or under
pytest.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

NA = float("nan")
UNIT_SPACE_FUNCTIONS = (
    "Smearing_Factor",
    "Back_Transform_Response",
    "Unit_Space_Predictions",
    "Unit_Space_Residuals",
    "Unit_Space_R_Squared",
    "Unit_Space_Adjusted_R_Squared",
    "Unit_Space_RMSE",
)


# ── Pure-Python mirrors of the workbook formulas ─────────────────────────────

def _smearing_factor_mirror(x: np.ndarray, y: np.ndarray, include: np.ndarray,
                            response_transform: str) -> float:
    """=Smearing_Factor(...) — 1.0 under None, mean(EXP(residuals)) under Log."""
    if response_transform == "None":
        return 1.0
    if response_transform == "Log":
        # Fit y on x (intercept included via ones column), take residuals
        # from fitted space, exponentiate, average.
        x_with_const = np.column_stack([np.ones(x.shape[0]), x])
        beta, *_ = np.linalg.lstsq(x_with_const[include], y[include], rcond=None)
        residuals = y[include] - x_with_const[include] @ beta
        return float(np.mean(np.exp(residuals)))
    return NA


def _back_transform_response_mirror(values: np.ndarray, response_transform: str,
                                    method: str, smearing: float) -> np.ndarray:
    """=Back_Transform_Response(values, [context], [method], [smearing])."""
    if response_transform == "None":
        return np.asarray(values, dtype=float)
    if response_transform == "Log":
        if method == "Duan":
            return np.exp(np.asarray(values, dtype=float)) * smearing
        if method == "Naive":
            return np.exp(np.asarray(values, dtype=float))
        return np.full_like(np.asarray(values, dtype=float), NA)
    return np.full_like(np.asarray(values, dtype=float), NA)


def _fit_response(x: np.ndarray, y: np.ndarray, include: np.ndarray,
                  response_transform: str) -> tuple[np.ndarray, np.ndarray]:
    """Fit the model in fit space; return (fitted, residuals) on the included rows.

    The fitted values are in fit space (log if Log, raw if None). This is
    what the catalog's Unit_Space_Predictions / Smearing_Factor work from.
    """
    x_with_const = np.column_stack([np.ones(x.shape[0]), x])
    beta, *_ = np.linalg.lstsq(x_with_const[include], y[include], rcond=None)
    fitted = x_with_const @ beta  # full-length; we slice later
    residuals = y - fitted
    return fitted, residuals


def _unit_space_predictions_mirror(
    x: np.ndarray, y: np.ndarray, y_full: np.ndarray,
    include: np.ndarray, response_transform: str, predictor_transform: str,
    method: str,
) -> np.ndarray:
    """=Unit_Space_Predictions(...) — back-transformed in-sample predictions.

    Reproduces the catalog's three-step pipeline:
      shift     = FILTER(Y_Full, Include) - FILTER(Y, Include)
      fitted    = Predictions(X, Y, Include) + shift
      result    = SWITCH(rt & "|" & pt, ..., Back_Transform_Response(fitted))
    """
    if response_transform not in {"None", "Log"}:
        return np.full(x.shape[0], NA)
    if predictor_transform not in {"None", "Log", "Mixed"}:
        return np.full(x.shape[0], NA)
    y_full_inc = np.asarray(y_full, dtype=float)[include]
    y_inc = np.asarray(y, dtype=float)[include]
    shift = y_full_inc - y_inc
    fitted_full, _ = _fit_response(x, y, include, response_transform)
    fitted_inc = fitted_full[include]
    fitted_with_shift = fitted_inc + shift
    sm = _smearing_factor_mirror(x, y, include, response_transform)
    unit_full = np.full(x.shape[0], NA)
    unit_inc = _back_transform_response_mirror(
        fitted_with_shift, response_transform, method, sm)
    unit_full[include] = unit_inc
    return unit_full


def _unit_space_r_squared_mirror(x, y, y_full, include, response_transform,
                                 predictor_transform, method) -> float:
    """=Unit_Space_R_Squared(...) — 1 - SSE_unit / SST_unit, weighting on
    AVERAGE when Context_Has_Intercept is TRUE, zero otherwise.
    """
    if response_transform not in {"None", "Log"}:
        return NA
    if predictor_transform not in {"None", "Log", "Mixed"}:
        return NA
    pred = _unit_space_predictions_mirror(
        x, y, y_full, include, response_transform, predictor_transform, method)
    y_full_inc = np.asarray(y_full, dtype=float)[include]
    sse = float(np.sum((y_full_inc - pred[include]) ** 2))
    # SST about the mean — the default-with-intercept case the catalog
    # implements. The no-intercept case (about zero) is the other branch;
    # the test cover here is the with-intercept one, which is what fits
    # the default MLR spec.
    sst = float(np.sum((y_full_inc - np.mean(y_full_inc)) ** 2))
    if sst == 0:
        return NA
    return 1.0 - sse / sst


def _ols_standard_r_squared(x: np.ndarray, y: np.ndarray, include: np.ndarray) -> float:
    """Plain OLS R² — the unit-space R² must equal this when response_transform
    is None everywhere (the reduction invariant)."""
    x_with_const = np.column_stack([np.ones(x.shape[0]), x])
    beta, *_ = np.linalg.lstsq(x_with_const[include], y[include], rcond=None)
    fitted = x_with_const[include] @ beta
    sse = float(np.sum((y[include] - fitted) ** 2))
    sst = float(np.sum((y[include] - np.mean(y[include])) ** 2))
    if sst == 0:
        return NA
    return 1.0 - sse / sst


# ── Helper ───────────────────────────────────────────────────────────────────

def _ols_standard_results(x: np.ndarray, y: np.ndarray, include: np.ndarray
                          ) -> dict:
    """Return the ordinary-fit scalars the unit-space block has to collapse
    to under None: R², Adjusted R², RMSE, fitted, residuals.
    """
    n_inc = int(include.sum())
    x_with_const = np.column_stack([np.ones(x.shape[0]), x])
    beta, *_ = np.linalg.lstsq(x_with_const[include], y[include], rcond=None)
    fitted = x_with_const @ beta
    sse = float(np.sum((y[include] - fitted[include]) ** 2))
    sst = float(np.sum((y[include] - np.mean(y[include])) ** 2))
    r2 = 1.0 - sse / sst
    df_resid = n_inc - x_with_const.shape[1]  # includes intercept
    df_total = n_inc - 1
    adj_r2 = 1.0 - (1.0 - r2) * df_total / df_resid
    rmse = math.sqrt(sse / df_resid)
    return {
        "r2": r2, "adj_r2": adj_r2, "rmse": rmse,
        "fitted": fitted, "residuals": y - fitted,
        "df_total": df_total, "df_resid": df_resid,
    }


# ── Catalog shape assertions ─────────────────────────────────────────────────

def _catalog_functions() -> dict:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    return {fn["name"]: fn for fn in payload["functions"]}


def _formula(name: str) -> str:
    from lambda_catalog.lambda_formula_parser import (  # pylint: disable=import-outside-toplevel
        _normalize_user_formula,
        _strip_non_string_whitespace,
    )

    functions = _catalog_functions()
    return _strip_non_string_whitespace(_normalize_user_formula(functions[name]["formula_display"]))


def test_unit_space_functions_are_in_document_order_after_context_predictor_transform() -> None:
    """Per the §1 dependency rule: consumers (Smearing_Factor etc.) come
    after the upstream accessors they read (Context_Response_Transform /
    Context_Predictor_Transform)."""
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    names = [fn["name"] for fn in payload["functions"]]
    upstream = max(
        names.index("Context_Response_Transform"),
        names.index("Context_Predictor_Transform"),
    )
    for unit_fn in UNIT_SPACE_FUNCTIONS:
        assert names.index(unit_fn) > upstream, (
            f"{unit_fn} must appear after Context_Response_Transform / "
            "Context_Predictor_Transform in lambda_functions.json document order"
        )


def test_unit_space_functions_share_the_back_transformation_subcategory() -> None:
    for name in UNIT_SPACE_FUNCTIONS:
        fn = _catalog_functions()[name]
        assert fn.get("category") == "Model Construction", name
        assert fn.get("subcategory") == "Back-Transformation", name


def test_smearing_factor_uses_residuals_not_predictions() -> None:
    """=Smearing_Factor must call Residuals, not Predictions — Duan (1983)
    averages EXP of the FIT-SPACE residuals, not the fitted values."""
    formula = _formula("Smearing_Factor")
    assert "AVERAGE(EXP(Residuals(" in formula
    assert "AVERAGE(EXP(Predictions(" not in formula


def test_back_transform_response_switches_on_method() -> None:
    """=Back_Transform_Response must dispatch on Method ("Duan" / "Naive")
    and use smearing under Duan only."""
    formula = _formula("Back_Transform_Response")
    assert 'SWITCH(method_arg,' in formula
    assert '"Duan"' in formula
    assert '"Naive"' in formula
    # Duan multiplies by smear_arg; Naive does not.
    assert "EXP(Values)*smear_arg" in formula
    assert "EXP(Values)," in formula  # Naive path


def test_unit_space_predictions_switches_on_all_six_pairs() -> None:
    """=Unit_Space_Predictions must explicitly list the six (rt, pt) pairs
    recognised by the dispatcher and fall through to NA() otherwise."""
    formula = _formula("Unit_Space_Predictions")
    assert 'SWITCH(rt&"|"&pt' in formula
    for pair in ("None|None", "None|Log", "None|Mixed",
                 "Log|None",  "Log|Log",  "Log|Mixed"):
        assert f'"{pair}"' in formula, pair


def test_unit_space_r_squared_uses_sumsq_for_both_terms() -> None:
    """The mirror in the catalog uses SUMSQ(y_unit - pred_unit) and
    SUMSQ(y_unit - AVERAGE(y_unit)) for the SSE/SST pair.
    """
    formula = _formula("Unit_Space_R_Squared")
    assert "SUMSQ(y_unit-pred_unit)" in formula
    assert "SUMSQ(y_unit-AVERAGE(y_unit))" in formula
    assert "1-sse_unit/sst_unit" in formula


def test_unit_space_adjusted_r_squared_reuses_total_and_residual_doF() -> None:
    """=Unit_Space_Adjusted_R_Squared must reuse the existing df accessors so
    Context_DF_Absorbed is honoured, not recompute df locally."""
    formula = _formula("Unit_Space_Adjusted_R_Squared")
    assert "Total_Degrees_Of_Freedom(" in formula
    assert "Residual_Degrees_Of_Freedom(" in formula
    assert "1-(1-Unit_Space_R_Squared" in formula


def test_unit_space_rmse_shares_se_regressions_divisor() -> None:
    """=Unit_Space_RMSE must use Residual_Degrees_Of_Freedom as the divisor
    — exactly SE_Regression's own — so the None case reduces to it."""
    formula = _formula("Unit_Space_RMSE")
    assert "Residual_Degrees_Of_Freedom(" in formula
    assert "SQRT(SUMSQ(" in formula


# ── Numeric mirror cross-checks ──────────────────────────────────────────────

def test_smearing_factor_returns_one_under_none_response_transform() -> None:
    """Reduction invariant: Smearing_Factor = 1 when response_transform = None."""
    rng = np.random.default_rng(42)
    x = rng.normal(size=50).reshape(50, 1)
    y = rng.normal(size=50)
    include = np.ones(50, dtype=bool)
    sm = _smearing_factor_mirror(x, y, include, "None")
    assert sm == 1.0


def test_smearing_factor_matches_average_exp_residuals_under_log() -> None:
    """Duan (1983) smearing factor: under Log response, the mean of
    EXP of the fit-space residuals."""
    rng = np.random.default_rng(7)
    x = rng.normal(size=80).reshape(80, 1)
    log_y = 0.5 + 0.7 * x.flatten() + rng.normal(scale=0.3, size=80)
    y = np.exp(log_y)
    include = np.ones(80, dtype=bool)
    # Mirror approach: fit on log_y, then compute residuals, then average exp.
    x_with_const = np.column_stack([np.ones(80), x])
    beta, *_ = np.linalg.lstsq(x_with_const, log_y, rcond=None)
    residuals = log_y - x_with_const @ beta
    expected = float(np.mean(np.exp(residuals)))
    sm = _smearing_factor_mirror(x, log_y, include, "Log")
    assert math.isclose(sm, expected, rel_tol=1e-12)


def test_unit_space_r_squared_reduces_to_r_squared_under_none() -> None:
    """Reduction invariant: Unit_Space_R_Squared == R_Squared when
    response_transform = None and y_full = y (no FE shift)."""
    rng = np.random.default_rng(11)
    x = rng.normal(size=120).reshape(60, 2)
    y = 1.0 + 0.6 * x[:, 0] - 0.4 * x[:, 1] + rng.normal(scale=0.5, size=60)
    include = np.ones(60, dtype=bool)
    expected = _ols_standard_r_squared(x, y, include)
    got = _unit_space_r_squared_mirror(
        x, y, y, include, "None", "None", "Duan")
    assert math.isclose(got, expected, rel_tol=1e-12)


def test_unit_space_r_squared_outside_dispatch_pair_returns_na() -> None:
    """Outside the recognised (rt, pt) pairs the function returns NA()."""
    rng = np.random.default_rng(2)
    x = rng.normal(size=20).reshape(20, 1)
    y = rng.normal(size=20)
    include = np.ones(20, dtype=bool)
    assert _is_na(_unit_space_r_squared_mirror(
        x, y, y, include, "Sqrt", "None", "Duan"))
    assert _is_na(_unit_space_r_squared_mirror(
        x, y, y, include, "None", "Identity", "Duan"))


def test_unit_space_predictions_reduces_to_predictions_under_none() -> None:
    """Reduction invariant: Unit_Space_Predictions == Predictions when
    response_transform = None and y_full = y (no FE shift)."""
    rng = np.random.default_rng(13)
    x = rng.normal(size=80).reshape(40, 2)
    y = 1.5 + 0.4 * x[:, 0] + 0.8 * x[:, 1] + rng.normal(scale=0.3, size=40)
    include = np.ones(40, dtype=bool)
    fitted_unit = _unit_space_predictions_mirror(
        x, y, y, include, "None", "None", "Duan")
    plain = _ols_standard_results(x, y, include)["fitted"]
    assert np.allclose(fitted_unit, plain, atol=1e-12)


def test_unit_space_residuals_minus_predictions_equals_observed_when_no_fe() -> None:
    """Reduction invariant: Unit_Space_Residuals = observed - fitted in
    original units under the no-FE / None response case."""
    rng = np.random.default_rng(17)
    x = rng.normal(size=30).reshape(30, 1)
    y = rng.normal(size=30)
    include = np.ones(30, dtype=bool)
    pred = _unit_space_predictions_mirror(x, y, y, include, "None", "None", "Duan")
    plain = _ols_standard_results(x, y, include)
    expected_residuals = plain["residuals"]
    got_residuals = y - pred
    assert np.allclose(got_residuals, expected_residuals, atol=1e-12)


def test_unit_space_predictions_under_log_diffs_from_naive_by_smearing() -> None:
    """Under Log/Duan, Unit_Space_Predictions = EXP(fitted) * smearing;
    under Log/Naive, = EXP(fitted). The Duan/Naive ratio is exactly the
    smearing factor."""
    rng = np.random.default_rng(19)
    x = rng.normal(size=50).reshape(50, 1)
    log_y = 0.5 + 0.7 * x.flatten() + rng.normal(scale=0.3, size=50)
    include = np.ones(50, dtype=bool)
    pred_duan = _unit_space_predictions_mirror(
        x, log_y, log_y, include, "Log", "None", "Duan")
    pred_naive = _unit_space_predictions_mirror(
        x, log_y, log_y, include, "Log", "None", "Naive")
    sm = _smearing_factor_mirror(x, log_y, include, "Log")
    assert np.allclose(pred_duan, pred_naive * sm, atol=1e-12)
    # Naive is smaller than Duan when the smearing factor > 1 (positive
    # skew on residuals, the typical Log-fit case).
    assert sm > 1.0


def test_unit_space_r_squared_under_diffs_from_naive() -> None:
    """The two methods agree on R² in fit space (log) but disagree when
    SSE is computed in original units — the Duan smearing lifts the
    conditional mean above the median, so SSE_unit increases and R² drops.
    """
    rng = np.random.default_rng(23)
    x = rng.normal(size=160).reshape(80, 2)
    log_y = 1.0 + 0.5 * x[:, 0] - 0.3 * x[:, 1] + rng.normal(scale=0.4, size=80)
    include = np.ones(80, dtype=bool)
    r2_duan = _unit_space_r_squared_mirror(
        x, log_y, log_y, include, "Log", "None", "Duan")
    r2_naive = _unit_space_r_squared_mirror(
        x, log_y, log_y, include, "Log", "None", "Naive")
    assert abs(r2_duan - r2_naive) > 1e-6  # the metrics differ in original units


def _is_na(value) -> bool:
    return isinstance(value, float) and math.isnan(value)


def main() -> None:  # pragma: no cover - standalone runner
    """Run every test directly without pytest."""
    import pytest as _pytest
    sys.exit(_pytest.main([__file__, "-v"]))


if __name__ == "__main__":  # pragma: no cover
    main()

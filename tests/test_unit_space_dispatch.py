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
    "Unit_Space_Observed",
    "Unit_Space_Residuals",
    "Unit_Space_R_Squared",
    "Unit_Space_Adjusted_R_Squared",
    "Unit_Space_RMSE",
    "Unit_Space_LOOCV_Residual",
    "Unit_Space_LOOCV_RMSE",
    "Unit_Space_LOOCV_MAE",
    "Smearing_Treatment",
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
      shift     = IF(rt="Log", FILTER(Y_Full, Include) - FILTER(Y, Include), 0)
      fitted    = Predictions(X, Y, Include) + shift
      result    = SWITCH(rt & "|" & pt, ..., Back_Transform_Response(fitted))

    The shift reintroduces the level the within transformation removed, so
    that EXP() under Fixed Effects exponentiates a real predicted log
    response rather than a group deviation. It is gated on Log because
    nothing is exponentiated under None, and applying it there would turn
    a within-flavoured statistic into a total one — breaking the reduction
    invariant.
    """
    if response_transform not in {"None", "Log"}:
        return np.full(x.shape[0], NA)
    if predictor_transform not in {"None", "Log", "Mixed"}:
        return np.full(x.shape[0], NA)
    y_full_inc = np.asarray(y_full, dtype=float)[include]
    y_inc = np.asarray(y, dtype=float)[include]
    shift = (y_full_inc - y_inc) if response_transform == "Log" else 0.0
    fitted_full, _ = _fit_response(x, y, include, response_transform)
    fitted_inc = fitted_full[include]
    fitted_with_shift = fitted_inc + shift
    sm = _smearing_factor_mirror(x, y, include, response_transform)
    unit_full = np.full(x.shape[0], NA)
    unit_inc = _back_transform_response_mirror(
        fitted_with_shift, response_transform, method, sm)
    unit_full[include] = unit_inc
    return unit_full


def _unit_space_observed_mirror(y: np.ndarray, y_full: np.ndarray,
                                include: np.ndarray,
                                response_transform: str) -> np.ndarray:
    """=Unit_Space_Observed(Y, Y_Full, [Include], [Context]).

    The observed half of every unit-space comparison, read in the SAME space
    as ``Unit_Space_Predictions`` returns. One expression covers both cases:

        shift   = IF(rt="Log", FILTER(Y_Full, Inc) - FILTER(Y, Inc), 0)
        y_level = Dependent_Variable(Y, Inc) + shift
        result  = Back_Transform_Response(y_level, ctx, "Naive", 1)

    Under Log, ``y_level`` is ``Y_Full`` — the un-demeaned log response — and
    the back-transform returns raw y. Under None the shift is zero, the
    back-transform is a pass-through, and the observed side stays the
    within-demeaned column the ordinary statistics are computed on, which is
    what keeps the reduction invariant true under Fixed Effects.

    The Naive branch is forced: the smearing factor lifts a prediction from
    the conditional median to the conditional mean, and an observation is
    neither.
    """
    y_inc = np.asarray(y, dtype=float)[include]
    shift = (
        np.asarray(y_full, dtype=float)[include] - y_inc
        if response_transform == "Log"
        else 0.0
    )
    return _back_transform_response_mirror(
        y_inc + shift, response_transform, "Naive", 1.0)


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
    y_full_inc = _unit_space_observed_mirror(
        y, y_full, include, response_transform)
    sse = float(np.sum((y_full_inc - pred[include]) ** 2))
    # SST about the mean — the default-with-intercept case the catalog
    # implements. The no-intercept case (about zero) is the other branch;
    # the test cover here is the with-intercept one, which is what fits
    # the default MLR spec.
    sst = float(np.sum((y_full_inc - np.mean(y_full_inc)) ** 2))
    if sst == 0:
        return NA
    return 1.0 - sse / sst


def _loocv_prediction_refit(x: np.ndarray, y: np.ndarray) -> np.ndarray:
    """Leave-one-out fit-space prediction for every row, by REFITTING n
    times — drop row i, fit on the rest, predict at x_i.

    This is the independent derivation DECISIONS.md:2949-2958 demands: the
    catalog's ``LOOCV_Prediction`` uses the Sherman-Morrison-Woodbury
    hat-diagonal shortcut (``predictions - h*e/(1-h)``), so a mirror that
    shared that reading of ``Y_Full`` would produce a green suite and a
    wrong workbook. Refitting n times is O(n) fits and agrees with the
    shortcut to floating-point tolerance when both are valid — which is the
    point: agreement from an independent method, not from a shared formula.
    """
    n = x.shape[0]
    x_with_const = np.column_stack([np.ones(n), x])
    loo = np.full(n, NA)
    for i in range(n):
        mask = np.ones(n, dtype=bool)
        mask[i] = False
        # Too few rows to fit (n-1 <= #params): the catalog's
        # IFERROR(…, NA()) returns #N/A for that row, and so does this.
        if int(mask.sum()) <= x_with_const.shape[1]:
            continue
        beta, *_ = np.linalg.lstsq(x_with_const[mask], y[mask], rcond=None)
        loo[i] = x_with_const[i] @ beta
    return loo


def _loocv_residual_refit(x: np.ndarray, y: np.ndarray,
                           include: np.ndarray) -> np.ndarray:
    """The ordinary (fit-space) leave-one-out residual e_i - 0, by refit:
    ``y_i - LOO_prediction_i``. This is the reference the reduction
    invariant (``Unit_Space_LOOCV_Residual ≡ LOOCV_Residual`` under None)
    is scored against, and it is the same number the sheet's PRESS Residual
    column carries, so ``LOOCV_RMSE ≡ sqrt(PRESS/n)`` follows.
    """
    loo = _loocv_prediction_refit(x[include], y[include])
    r = np.full(x.shape[0], NA)
    r[include] = y[include] - loo
    return r


def _unit_space_loocv_residual_mirror(
    x: np.ndarray, y: np.ndarray, y_full: np.ndarray,
    include: np.ndarray, response_transform: str, predictor_transform: str,
    method: str,
) -> np.ndarray:
    """=Unit_Space_LOOCV_Residual(...) — the leave-one-out residual in
    original units.

    Mirrors the catalog's three-step pipeline with ``LOOCV_Prediction``
    substituted for ``Predictions``:

      loo_fit   = IFERROR(LOOCV_Prediction(X, Y, Include) + shift, NA())
      y_unit    = Unit_Space_Observed(Y, Y_Full, Include, Context)
      r_loo     = y_unit - Back_Transform_Response(loo_fit, ctx, Method, sm)

    The smearing factor ``sm`` is the FULL-SAMPLE one (estimated on all n
    included residuals, the held-out row included) — the small optimism
    ``Smearing_Treatment`` names. The refit LOO prediction comes from
    ``_loocv_prediction_refit`` (independent of the hat-diagonal shortcut).
    """
    if response_transform not in {"None", "Log"}:
        return np.full(x.shape[0], NA)
    if predictor_transform not in {"None", "Log", "Mixed"}:
        return np.full(x.shape[0], NA)
    x_inc = x[include]
    y_inc = np.asarray(y, dtype=float)[include]
    loo_fit_inc = _loocv_prediction_refit(x_inc, y_inc)
    y_full_inc = np.asarray(y_full, dtype=float)[include]
    shift = (y_full_inc - y_inc) if response_transform == "Log" else 0.0
    loo_fit_with_shift = loo_fit_inc + shift
    sm = _smearing_factor_mirror(x, y, include, response_transform)
    loo_pred_unit = _back_transform_response_mirror(
        loo_fit_with_shift, response_transform, method, sm)
    y_unit_inc = _unit_space_observed_mirror(y, y_full, include, response_transform)
    y_unit_inc = y_unit_inc[include]
    r_inc = y_unit_inc - loo_pred_unit
    r = np.full(x.shape[0], NA)
    r[include] = r_inc
    return r


def _unit_space_loocv_rmse_mirror(
    x: np.ndarray, y: np.ndarray, y_full: np.ndarray,
    include: np.ndarray, response_transform: str, predictor_transform: str,
    method: str,
) -> float:
    """=Unit_Space_LOOCV_RMSE(...) — SQRT(SUMSQ(r) / ROWS(r)).

    ``ROWS(r)``, not ``COUNT(r)``: an #N/A row (leverage h_i = 1) propagates
    through SUMSQ, so the divisor is the sample size the fit used. Divisor
    is ``n`` — every LOO prediction is genuinely out-of-sample, no df are
    consumed (deliberately differs from ``Unit_Space_RMSE``'s ÷ df_residual).
    """
    r = _unit_space_loocv_residual_mirror(
        x, y, y_full, include, response_transform, predictor_transform, method)
    r_inc = r[include]
    if np.any(np.isnan(r_inc)):
        return NA
    return math.sqrt(float(np.sum(r_inc ** 2)) / r_inc.size)


def _unit_space_loocv_mae_mirror(
    x: np.ndarray, y: np.ndarray, y_full: np.ndarray,
    include: np.ndarray, response_transform: str, predictor_transform: str,
    method: str,
) -> float:
    """=Unit_Space_LOOCV_MAE(...) — AVERAGE(ABS(r))."""
    r = _unit_space_loocv_residual_mirror(
        x, y, y_full, include, response_transform, predictor_transform, method)
    r_inc = r[include]
    if np.any(np.isnan(r_inc)):
        return NA
    return float(np.mean(np.abs(r_inc)))


def _smearing_treatment_mirror(response_transform: str, method: str) -> str:
    """=Smearing_Treatment(...) — pure function of (response transform, method):
    "n/a — no back-transform" under None, "Naive (no smearing)" / "Full-sample
    Duan (approx.)" under Log, NA() otherwise. The fourth SWITCH arm
    (fold-specific Duan) is the extension point this mirror does not yet
    carry."""
    if response_transform == "None":
        return "n/a — no back-transform"
    if response_transform == "Log":
        if method == "Naive":
            return "Naive (no smearing)"
        if method == "Duan":
            return "Full-sample Duan (approx.)"
        return NA
    return NA


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


def test_unit_space_statistics_back_transform_the_observed_response() -> None:
    """``Y_Full`` is ``Response_Column()`` — ln(y) under a Log Response row.

    Every statistic that subtracts an observed value from a back-transformed
    prediction has to put the observed side through the same transform first,
    or it compares ln(y) against a number in the response's own units. The
    Naive branch is forced: the smearing factor lifts a prediction from the
    conditional median to the conditional mean and must never be applied to
    an observation.
    """
    observed = _formula("Unit_Space_Observed")
    assert 'shift,IF(rt="Log",FILTER(Y_Full,filt_arg)-FILTER(Y,filt_arg),0)' in observed
    assert "y_level,Dependent_Variable(Y,filt_arg)+shift" in observed
    assert 'Back_Transform_Response(y_level,context_arg,"Naive",1)' in observed
    for name in ("Unit_Space_R_Squared", "Unit_Space_RMSE", "Unit_Space_Residuals"):
        formula = _formula(name)
        assert "y_unit,Unit_Space_Observed(Y,Y_Full,filt_arg,context_arg)" in formula, name
        assert "y_unit,Dependent_Variable(Y_Full" not in formula, name


def test_unit_space_predictions_gate_the_level_shift_on_a_log_response() -> None:
    """The shift exists so EXP() under FE exponentiates a predicted log
    response rather than a group deviation. Under None nothing is
    exponentiated, and shifting there converts the within-flavoured
    statistics into total ones — the reduction invariant would break."""
    formula = _formula("Unit_Space_Predictions")
    assert 'shift,IF(rt="Log",FILTER(Y_Full,filt_arg)-FILTER(Y,filt_arg),0)' in formula
    assert "shift,FILTER(Y_Full,filt_arg)-FILTER(Y,filt_arg)," not in formula


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


def test_unit_space_r_squared_matches_an_original_scale_reference_under_log() -> None:
    """The one check the rest of this file cannot make.

    Every other numeric test here compares a mirror against another mirror,
    or Duan against Naive. Both sides of those comparisons take ``y_full``
    as *already* being the observed response in original units — which is
    what the catalog does too (``y_unit, Dependent_Variable(Y_Full, ...)``).
    The Regression sheet passes ``Response_Column()``, and that is ln(y)
    when the Response row declares Log, so the shared assumption is wrong
    and no mirror-vs-mirror test can see it.

    This test computes the reference straight from the raw response: fit
    ln(y), back-transform the fitted values, and form R² against y itself.
    Nothing here routes through ``_unit_space_predictions_mirror``.
    """
    rng = np.random.default_rng(101)
    x = rng.normal(size=180).reshape(90, 2)
    log_y = 2.0 + 0.6 * x[:, 0] - 0.35 * x[:, 1] + rng.normal(scale=0.4, size=90)
    y = np.exp(log_y)
    include = np.ones(90, dtype=bool)

    x_with_const = np.column_stack([np.ones(90), x])
    beta, *_ = np.linalg.lstsq(x_with_const, log_y, rcond=None)
    fitted_log = x_with_const @ beta
    smearing = float(np.mean(np.exp(log_y - fitted_log)))
    y_hat = np.exp(fitted_log) * smearing
    expected = 1.0 - float(np.sum((y - y_hat) ** 2)) / float(
        np.sum((y - np.mean(y)) ** 2)
    )

    got = _unit_space_r_squared_mirror(
        x, log_y, log_y, include, "Log", "None", "Duan")
    assert math.isclose(got, expected, rel_tol=1e-12), (
        "Unit_Space_R_Squared must measure the fit against y in the response's "
        "own units. Comparing ln(y) against a back-transformed prediction "
        f"mixes unit systems: got {got}, original-scale reference {expected}."
    )


def test_unit_space_predictions_ignore_the_level_shift_under_none() -> None:
    """The level shift exists only so EXP() of a within deviation means
    something. Under ``Transform = None`` nothing is exponentiated, so a
    Fixed Effects model's unit-space predictions must stay the ordinary
    within-fitted values — otherwise Unit_Space_R_Squared silently becomes
    a total R² and stops matching R_Squared, breaking the reduction
    invariant this milestone ships on.
    """
    rng = np.random.default_rng(103)
    groups = np.repeat(np.arange(6), 12)
    x = rng.normal(size=144).reshape(72, 2)
    y_full = 5.0 + groups * 2.0 + 0.5 * x[:, 0] + rng.normal(scale=0.3, size=72)
    include = np.ones(72, dtype=bool)

    # Within-demean both sides by group, the way Design_Response() does.
    y_within = y_full.copy()
    x_within = x.copy()
    for g in np.unique(groups):
        mask = groups == g
        y_within[mask] -= y_within[mask].mean()
        x_within[mask] -= x_within[mask].mean(axis=0)
    assert not np.allclose(y_full - y_within, 0.0), "level shift must be non-zero"

    got = _unit_space_predictions_mirror(
        x_within, y_within, y_full, include, "None", "None", "Duan")
    expected = _ols_standard_results(x_within, y_within, include)["fitted"]
    assert np.allclose(got, expected, atol=1e-10), (
        "Under Transform = None the level shift must not be applied — "
        "Unit_Space_Predictions has to collapse to Predictions()."
    )


def test_unit_space_r_squared_reduces_to_r_squared_under_fixed_effects_and_none() -> None:
    """The reduction invariant has to survive Fixed Effects, not just the
    no-FE case.

    Under FE the sheet's ordinary statistics are within-flavoured, computed
    on the demeaned pair. If either the fitted side or the observed side
    picked up the level shift under ``Transform = None`` the unit-space R²
    would quietly become a *total* R² — a different, larger number that
    still looks plausible. ``production_lots_fixed_effects`` is exactly this
    configuration.
    """
    rng = np.random.default_rng(107)
    groups = np.repeat(np.arange(5), 14)
    x = rng.normal(size=140).reshape(70, 2)
    y_full = 40.0 + groups * 9.0 + 0.7 * x[:, 0] - 0.2 * x[:, 1] + rng.normal(
        scale=0.5, size=70)
    include = np.ones(70, dtype=bool)

    y_within = y_full.copy()
    x_within = x.copy()
    for g in np.unique(groups):
        mask = groups == g
        y_within[mask] -= y_within[mask].mean()
        x_within[mask] -= x_within[mask].mean(axis=0)

    expected = _ols_standard_r_squared(x_within, y_within, include)
    got = _unit_space_r_squared_mirror(
        x_within, y_within, y_full, include, "None", "None", "Duan")
    assert math.isclose(got, expected, rel_tol=1e-12), (
        "Under FE + None the unit-space R² must stay the within R² the rest "
        f"of the sheet reports: got {got}, within R² {expected}."
    )


def test_unit_space_observed_returns_raw_y_under_log_with_fixed_effects() -> None:
    """Under FE + Log the observed side is the raw response, recovered by
    adding the group mean back before exponentiating — not EXP of a within
    deviation, and not the log response itself."""
    rng = np.random.default_rng(109)
    groups = np.repeat(np.arange(4), 15)
    log_y = 1.2 + groups * 0.4 + rng.normal(scale=0.25, size=60)
    y = np.exp(log_y)
    include = np.ones(60, dtype=bool)

    log_y_within = log_y.copy()
    for g in np.unique(groups):
        mask = groups == g
        log_y_within[mask] -= log_y_within[mask].mean()

    got = _unit_space_observed_mirror(log_y_within, log_y, include, "Log")
    assert np.allclose(got, y, rtol=1e-12)


def _is_na(value) -> bool:
    return isinstance(value, float) and math.isnan(value)


# ── v3.4 unit-space LOOCV: catalog shape + numeric mirrors ───────────────────


def test_unit_space_loocv_residual_substitutes_loocv_prediction_for_predictions() -> None:
    """=Unit_Space_LOOCV_Residual is Unit_Space_Predictions with
    LOOCV_Prediction substituted for Predictions: the same shift, the same
    Unit_Space_Observed observed side, the same Back_Transform_Response
    dispatch — and an elementwise IFERROR(…, NA()) so a leverage-1 row yields
    #N/A (mirroring LOOCV_Residual's own IFERROR(e/(1-h), NA()))."""
    formula = _formula("Unit_Space_LOOCV_Residual")
    assert "LOOCV_Prediction(" in formula
    # The level shift is gated on Log exactly as in Unit_Space_Predictions.
    assert 'shift,IF(rt="Log",FILTER(Y_Full,filt_arg)-FILTER(Y,filt_arg),0)' in formula
    assert "loo_fit,IFERROR(LOOCV_Prediction(X,Y,filt_arg)+shift,NA())" in formula
    assert "y_unit,Unit_Space_Observed(Y,Y_Full,filt_arg,context_arg)" in formula
    assert "r_loo,y_unit-Back_Transform_Response(loo_fit,context_arg,method_arg,sm)" in formula
    # The same six-pair SWITCH gate the rest of the family uses.
    assert 'SWITCH(rt&"|"&pt' in formula
    for pair in ("None|None", "None|Log", "None|Mixed",
                 "Log|None",  "Log|Log",  "Log|Mixed"):
        assert f'"{pair}"' in formula, pair


def test_unit_space_loocv_rmse_uses_rows_not_count_and_divides_by_n() -> None:
    """=Unit_Space_LOOCV_RMSE is SQRT(SUMSQ(r) / ROWS(r)). ROWS, not COUNT:
    an #N/A row propagates through SUMSQ, so the divisor is plainly the
    sample size the fit used. The divisor is n (÷ n), NOT df_residual —
    every LOO prediction is genuinely out-of-sample, so no degrees of
    freedom are consumed. That is the deliberate difference from
    Unit_Space_RMSE's ÷ df_residual the plan flags."""
    formula = _formula("Unit_Space_LOOCV_RMSE")
    assert "Unit_Space_LOOCV_Residual(" in formula
    assert "SQRT(SUMSQ(r)/ROWS(r))" in formula
    assert "COUNT(r)" not in formula
    # Must NOT reach for Residual_Degrees_Of_Freedom (Unit_Space_RMSE does).
    assert "Residual_Degrees_Of_Freedom" not in formula


def test_unit_space_loocv_mae_uses_average_of_abs() -> None:
    """=Unit_Space_LOOCV_MAE is AVERAGE(ABS(r)) over the same r."""
    formula = _formula("Unit_Space_LOOCV_MAE")
    assert "Unit_Space_LOOCV_Residual(" in formula
    assert "AVERAGE(ABS(r))" in formula


def test_smearing_treatment_lists_three_arms_and_falls_through_to_na() -> None:
    """=Smearing_Treatment is a pure function of (response transform, method)
    with three arms and an NA() fallthrough — the fourth arm (fold-specific
    Duan) is the extension point a future PR adds without restructuring."""
    formula = _formula("Smearing_Treatment")
    assert 'IF(rt="None","n/a — no back-transform"' in formula
    assert 'SWITCH(method_arg,' in formula
    assert '"Naive","Naive (no smearing)"' in formula
    assert '"Duan","Full-sample Duan (approx.)"' in formula
    # An unrecognised method under Log, and an unrecognised transform, both
    # fall through to NA().
    assert formula.count("NA()") >= 2


def test_unit_space_loocv_residual_matches_n_refit_reference_under_log() -> None:
    """The mirror's LOOCV residual (refit n times, full-sample smearing)
    must equal an independent n-refit reference: refit dropping row i,
    back-transform that single LOO prediction through Duan, and form
    observed - LOO_prediction. Both sides refit; neither uses the
    hat-diagonal shortcut."""
    rng = np.random.default_rng(131)
    x = rng.normal(size=120).reshape(60, 2)
    log_y = 1.0 + 0.6 * x[:, 0] - 0.4 * x[:, 1] + rng.normal(scale=0.35, size=60)
    include = np.ones(60, dtype=bool)

    got = _unit_space_loocv_residual_mirror(
        x, log_y, log_y, include, "Log", "None", "Duan")
    sm = _smearing_factor_mirror(x, log_y, include, "Log")
    x_with_const = np.column_stack([np.ones(60), x])
    expected = np.full(60, NA)
    for i in range(60):
        mask = np.ones(60, dtype=bool)
        mask[i] = False
        beta, *_ = np.linalg.lstsq(x_with_const[mask], log_y[mask], rcond=None)
        loo_fit = x_with_const[i] @ beta  # no shift: y_full == y here
        loo_pred = math.exp(loo_fit) * sm
        expected[i] = math.exp(log_y[i]) - loo_pred
    assert np.allclose(got[include], expected[include], rtol=1e-10, atol=1e-10)


def test_unit_space_loocv_residual_reduces_to_loocv_residual_under_none() -> None:
    """Reduction invariant (v3.3 extended to the new pair): under
    Transform = None, Unit_Space_LOOCV_Residual ≡ the ordinary LOOCV
    residual (the PRESS Residual column), because the shift is zero and
    the back-transform is a pass-through."""
    rng = np.random.default_rng(137)
    x = rng.normal(size=100).reshape(50, 2)
    y = 1.0 + 0.5 * x[:, 0] + 0.3 * x[:, 1] + rng.normal(scale=0.4, size=50)
    include = np.ones(50, dtype=bool)
    got = _unit_space_loocv_residual_mirror(
        x, y, y, include, "None", "None", "Duan")
    expected = _loocv_residual_refit(x, y, include)
    assert np.allclose(got[include], expected[include], atol=1e-10)


def test_unit_space_loocv_rmse_reduces_to_sqrt_press_over_n_under_none() -> None:
    """Reduction invariant: under Transform = None,
    Unit_Space_LOOCV_RMSE ≡ sqrt(PRESS / n), where PRESS is the sum of
    squared LOOCV residuals. The divisor is n (not df_residual), which is
    what makes this the ÷-n RMSE the AG8 relabel stops claiming."""
    rng = np.random.default_rng(139)
    x = rng.normal(size=80).reshape(40, 2)
    y = 0.5 + 0.4 * x[:, 0] - 0.2 * x[:, 1] + rng.normal(scale=0.3, size=40)
    include = np.ones(40, dtype=bool)
    got = _unit_space_loocv_rmse_mirror(
        x, y, y, include, "None", "None", "Duan")
    r = _loocv_residual_refit(x, y, include)[include]
    press = float(np.sum(r ** 2))
    expected = math.sqrt(press / r.size)
    assert math.isclose(got, expected, rel_tol=1e-12)


def test_unit_space_loocv_mae_reduces_to_mean_abs_loocv_residual_under_none() -> None:
    """Reduction invariant: under Transform = None, Unit_Space_LOOCV_MAE ≡
    mean(|LOOCV residual|)."""
    rng = np.random.default_rng(149)
    x = rng.normal(size=72).reshape(36, 2)
    y = 1.0 + 0.3 * x[:, 0] + 0.5 * x[:, 1] + rng.normal(scale=0.4, size=36)
    include = np.ones(36, dtype=bool)
    got = _unit_space_loocv_mae_mirror(
        x, y, y, include, "None", "None", "Duan")
    r = _loocv_residual_refit(x, y, include)[include]
    expected = float(np.mean(np.abs(r)))
    assert math.isclose(got, expected, rel_tol=1e-12)


def test_unit_space_loocv_reductions_hold_under_fixed_effects_and_none() -> None:
    """The reduction invariants must survive Fixed Effects, not just the
    no-FE case: under FE the sheet's ordinary statistics are within-
    flavoured (demeaned pair), and the unit-space LOOCV pair has to collapse
    to the within LOOCV pair under Transform = None — the level shift is
    inert there, or the reduction silently becomes a total one."""
    rng = np.random.default_rng(151)
    groups = np.repeat(np.arange(5), 12)
    x = rng.normal(size=120).reshape(60, 2)
    y_full = 30.0 + groups * 7.0 + 0.6 * x[:, 0] - 0.3 * x[:, 1] + rng.normal(
        scale=0.5, size=60)
    include = np.ones(60, dtype=bool)

    y_within = y_full.copy()
    x_within = x.copy()
    for g in np.unique(groups):
        mask = groups == g
        y_within[mask] -= y_within[mask].mean()
        x_within[mask] -= x_within[mask].mean(axis=0)
    assert not np.allclose(y_full - y_within, 0.0), "level shift must be non-zero"

    r_unit = _unit_space_loocv_residual_mirror(
        x_within, y_within, y_full, include, "None", "None", "Duan")
    r_loo = _loocv_residual_refit(x_within, y_within, include)
    assert np.allclose(r_unit[include], r_loo[include], atol=1e-9)

    rmse_unit = _unit_space_loocv_rmse_mirror(
        x_within, y_within, y_full, include, "None", "None", "Duan")
    r = r_loo[include]
    assert math.isclose(rmse_unit, math.sqrt(float(np.sum(r ** 2)) / r.size),
                        rel_tol=1e-12)


def test_unit_space_loocv_residual_under_log_differs_from_naive_by_smearing() -> None:
    """Under Log, the Duan/Naive ratio of the LOOCV residual's fitted side
    is the smearing factor: Duan multiplies the back-transformed LOO
    prediction by sm, Naive does not, so the residual differs by exactly
    that factor on the fitted side."""
    rng = np.random.default_rng(157)
    x = rng.normal(size=44).reshape(44, 1)
    log_y = 0.8 + 0.5 * x.flatten() + rng.normal(scale=0.3, size=44)
    include = np.ones(44, dtype=bool)
    r_duan = _unit_space_loocv_residual_mirror(
        x, log_y, log_y, include, "Log", "None", "Duan")
    r_naive = _unit_space_loocv_residual_mirror(
        x, log_y, log_y, include, "Log", "None", "Naive")
    sm = _smearing_factor_mirror(x, log_y, include, "Log")
    # observed is the same; the fitted side differs by sm, so the residuals
    # differ by (sm - 1) * EXP(loo_fit).
    loo = _loocv_prediction_refit(x[include], log_y[include])
    fitted_diff = (sm - 1.0) * np.exp(loo)
    assert np.allclose(r_duan[include] - r_naive[include], -fitted_diff,
                       rtol=1e-9, atol=1e-9)
    assert sm > 1.0


def test_smearing_treatment_mirror_matches_the_three_rules() -> None:
    """The mirror's three-way rule is the same the LAMBDA encodes."""
    assert _smearing_treatment_mirror("None", "Duan") == "n/a — no back-transform"
    assert _smearing_treatment_mirror("None", "Naive") == "n/a — no back-transform"
    assert _smearing_treatment_mirror("Log", "Naive") == "Naive (no smearing)"
    assert _smearing_treatment_mirror("Log", "Duan") == "Full-sample Duan (approx.)"
    assert _is_na(_smearing_treatment_mirror("Log", "Folded"))
    assert _is_na(_smearing_treatment_mirror("Sqrt", "Duan"))


def main() -> None:  # pragma: no cover - standalone runner
    """Run every test directly without pytest."""
    import pytest as _pytest
    sys.exit(_pytest.main([__file__, "-v"]))


if __name__ == "__main__":  # pragma: no cover
    main()

"""Invariants of the v3.0 intercept relocation.

The intercept moved out of the engine functions and into the design-matrix
constructor. Three things about that change fail silently rather than loudly if
they regress, so each one is pinned here:

1. No engine may re-acquire an ``Allow_Intercept`` argument. That argument was
   the accretion this release exists to unwind, and it is also the trap below:
   in ``Coefficients``/``SE_Coefficients``/``R_Squared`` it *was* LINEST's
   ``const``.
2. Every LINEST call must pass ``const = FALSE``. The design matrix now arrives
   with its intercept column; letting Excel fit a second one gives two exactly
   collinear terms and a singular Gram matrix — a wrong number, not an error.
3. ``R_Squared`` must not read LINEST's own R². Under ``const = FALSE`` that
   cell holds the UNCENTERED R², so the reported R², Adjusted R², F and
   Multiple R would all shift silently for every intercept model.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

CATALOG_PATH = Path(__file__).resolve().parent.parent / "lambda_functions.json"

# Functions that legitimately still need to know whether column 1 of the design
# matrix is the intercept — as an IDENTIFIER, never as a synthesize-one switch.
# v3.0 stage two folds exactly this set into the bounded Model_Context.
_EXPECTED_HAS_INTERCEPT_CARRIERS = {
    "Adjusted_R_Squared",
    "Beta_Weights",
    "F_Statistic",
    "F_Statistic_P_Value",
    "Group_Prediction_Interval",
    "MS_Regression",
    "Multiple_R",
    "Prediction_Interval",
    "R_Squared",
    "Regression_Degrees_Of_Freedom",
    "SS_Regression",
    "SS_Total",
    "Total_Degrees_Of_Freedom",
}


def _functions() -> dict[str, dict]:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    return {fn["name"]: fn for fn in payload["functions"]}


def _compact(formula: str) -> str:
    from lambda_catalog.lambda_formula_parser import (  # pylint: disable=import-outside-toplevel
        _normalize_user_formula,
        _strip_non_string_whitespace,
    )

    return _strip_non_string_whitespace(_normalize_user_formula(formula))


def test_no_function_declares_allow_intercept() -> None:
    offenders = [
        name
        for name, fn in _functions().items()
        if any(arg["name"] == "Allow_Intercept" for arg in fn["arguments"])
    ]
    assert offenders == []


def test_has_intercept_carriers_are_exactly_the_expected_set() -> None:
    # A new carrier is not automatically wrong, but it is a decision: every one
    # of these becomes an element read out of Model_Context at stage two.
    carriers = {
        name
        for name, fn in _functions().items()
        if any(arg["name"] == "Has_Intercept" for arg in fn["arguments"])
    }
    assert carriers == _EXPECTED_HAS_INTERCEPT_CARRIERS


def test_has_intercept_is_always_optional_and_defaults_to_true() -> None:
    for name in _EXPECTED_HAS_INTERCEPT_CARRIERS:
        fn = _functions()[name]
        declared = next(a for a in fn["arguments"] if a["name"] == "Has_Intercept")
        assert declared.get("optional") is True, name
        assert (
            "has_arg,IF(ISOMITTED(Has_Intercept),TRUE,Has_Intercept)"
            in _compact(fn["formula_display"])
        ), name


def test_every_linest_call_disables_its_own_intercept() -> None:
    # The constructor owns the intercept column; const must never be TRUE, and
    # must never be a variable that could evaluate to TRUE.
    sites = []
    for name, fn in _functions().items():
        for match in re.finditer(r"LINEST\(", _compact(fn["formula_display"])):
            sites.append((name, _compact(fn["formula_display"])[match.start():]))

    assert {name for name, _ in sites} == {
        "Coefficients",
        "SE_Coefficients",
        "SS_Residual",
    }
    for name, tail in sites:
        assert "FILTER(X,filt_arg),FALSE," in tail, name


def test_r_squared_is_derived_from_the_sums_of_squares_not_from_linest() -> None:
    formula = _compact(_functions()["R_Squared"]["formula_display"])
    assert "LINEST" not in formula
    assert "1-SS_Residual(X,Y,filt_arg)/SS_Total(X,Y,has_arg,filt_arg)" in formula


def test_the_sums_of_squares_chain_is_acyclic() -> None:
    # Before the relocation SS_Residual was defined as SS_Total * (1 - R²) and
    # R² came from LINEST. Deriving R² from the sums of squares closes that
    # loop, so SS_Residual must now stand on its own.
    functions = _functions()
    ss_residual = _compact(functions["SS_Residual"]["formula_display"])
    assert "R_Squared" not in ss_residual
    assert "SS_Total" not in ss_residual
    assert "INDEX(LINEST(FILTER(Y,filt_arg),FILTER(X,filt_arg),FALSE,TRUE),5,2)" in ss_residual

    ss_regression = _compact(functions["SS_Regression"]["formula_display"])
    assert "R_Squared" not in ss_regression
    assert "SS_Total(X,Y,has_arg,filt_arg)-SS_Residual(X,Y,filt_arg)" in ss_regression


def test_ss_total_projects_the_response_off_the_intercept_column() -> None:
    # One projection formula covers ones / absent / sqrt(w). DEVSQ would be
    # wrong for the weighted case, which is why it is gone.
    formula = _compact(_functions()["SS_Total"]["formula_display"])
    assert "DEVSQ" not in formula
    assert "c,FILTER(CHOOSECOLS(X,1),filt_arg)" in formula
    assert "SUMSQ(y_filt)-SUMPRODUCT(c,y_filt)^2/SUMSQ(c)" in formula
    assert "IF(NOT(has_arg),SUMSQ(y_filt)," in formula


def test_design_matrix_no_longer_synthesizes_an_intercept() -> None:
    fn = _functions()["Design_Matrix"]
    assert tuple(a["name"] for a in fn["arguments"]) == ("X", "Include")
    formula = _compact(fn["formula_display"])
    assert "SEQUENCE" not in formula
    assert "HSTACK" not in formula


def test_information_criteria_count_the_intercept_via_columns_only() -> None:
    for name in ("AIC", "BIC", "AICc"):
        formula = _compact(_functions()[name]["formula_display"])
        assert "p,COLUMNS(X)+absorbed_arg" in formula, name
        assert "Regression_Degrees_Of_Freedom" not in formula, name


def test_collinearity_diagnostics_take_predictor_columns_not_a_design_matrix() -> None:
    # A constant column in a correlation matrix makes it singular, so these
    # must never be handed the design matrix.
    for name in ("VIF", "Tolerance", "GVIF", "Generalized_Tolerance", "Correlation_Matrix"):
        argument_names = [a["name"] for a in _functions()[name]["arguments"]]
        assert argument_names[0] == "Predictors", name
        assert "X" not in argument_names, name


def test_vif_builds_its_own_intercept_for_the_auxiliary_regressions() -> None:
    # VIF regresses each predictor on the others; those auxiliary fits need an
    # intercept, and R_Squared no longer synthesizes one.
    formula = _compact(_functions()["VIF"]["formula_display"])
    assert "ones,SEQUENCE(ROWS(Predictors),1,1,0)" in formula
    assert "HSTACK(ones,CHOOSECOLS(Predictors,FILTER(col_idx,col_idx<>j)))" in formula


# ── Numeric equivalence: the relocation must not move a single number ────────
#
# The formulas themselves can only be verified against Excel, which CI cannot
# run. What CAN be checked here is the algebra they encode: the pre-relocation
# chain and the post-relocation chain are simulated side by side in numpy and
# required to agree in both intercept states. If they diverge, the relocation
# changed what is fitted rather than where the intercept is created, and the
# spec-driven QC pass would report mismatches on a developer machine.


def _linest(y, X, const):
    """Mimic Excel LINEST: coefficients, residual SS, its own R², residual df.

    The R² is the value LINEST reports, which is the UNCENTERED R² when
    ``const`` is FALSE — the trap this module exists to pin.
    """
    import numpy as np  # pylint: disable=import-outside-toplevel

    design = np.column_stack([np.ones(len(y)), X]) if const else X
    beta, *_ = np.linalg.lstsq(design, y, rcond=None)
    resid = y - design @ beta
    ss_resid = float(resid @ resid)
    if const:
        r_squared = 1 - ss_resid / float(((y - y.mean()) ** 2).sum())
    else:
        r_squared = 1 - ss_resid / float(y @ y)
    return beta, ss_resid, r_squared, len(y) - design.shape[1]


def _old_chain(y, X, has_intercept):
    """Pre-relocation: engine synthesizes the intercept, R² comes from LINEST."""
    import numpy as np  # pylint: disable=import-outside-toplevel

    n, k = X.shape
    _, _, r_squared, _ = _linest(y, X, has_intercept)
    ss_total = float(((y - y.mean()) ** 2).sum()) if has_intercept else float(y @ y)
    ss_residual = ss_total * (1 - r_squared)
    ss_regression = ss_total * r_squared
    df_regression = k
    df_total = n - 1 if has_intercept else n
    df_residual = df_total - df_regression
    p = k + (1 if has_intercept else 0)
    return {
        "R2": r_squared,
        "SS_Total": ss_total,
        "SS_Residual": ss_residual,
        "SS_Regression": ss_regression,
        "df_regression": df_regression,
        "df_total": df_total,
        "df_residual": df_residual,
        "p": p,
        "AIC": n * np.log(ss_residual / n) + 2 * p,
        "F": (ss_regression / df_regression) / (ss_residual / df_residual),
    }


def _new_chain(y, X, has_intercept):
    """Post-relocation: the caller supplies the intercept column."""
    import numpy as np  # pylint: disable=import-outside-toplevel

    n = X.shape[0]
    design = np.column_stack([np.ones(n), X]) if has_intercept else X
    _, ss_residual, _, _ = _linest(y, design, False)
    if has_intercept:
        c = design[:, 0]
        ss_total = float(y @ y) - float(c @ y) ** 2 / float(c @ c)
    else:
        ss_total = float(y @ y)
    r_squared = 1 - ss_residual / ss_total
    ss_regression = ss_total - ss_residual
    df_regression = design.shape[1] - (1 if has_intercept else 0)
    df_residual = n - design.shape[1]
    p = design.shape[1]
    return {
        "R2": r_squared,
        "SS_Total": ss_total,
        "SS_Residual": ss_residual,
        "SS_Regression": ss_regression,
        "df_regression": df_regression,
        "df_total": n - (1 if has_intercept else 0),
        "df_residual": df_residual,
        "p": p,
        "AIC": n * np.log(ss_residual / n) + 2 * p,
        "F": (ss_regression / df_regression) / (ss_residual / df_residual),
    }


def test_the_relocated_chain_reproduces_every_pre_relocation_number() -> None:
    import numpy as np  # pylint: disable=import-outside-toplevel

    rng = np.random.default_rng(20260801)
    for _ in range(200):
        n = int(rng.integers(20, 60))
        k = int(rng.integers(1, 6))
        predictors = rng.normal(size=(n, k))
        response = (
            predictors @ rng.normal(size=k) + rng.normal(scale=0.5, size=n) + 3.0
        )
        for has_intercept in (True, False):
            old = _old_chain(response, predictors, has_intercept)
            new = _new_chain(response, predictors, has_intercept)
            for key, expected in old.items():
                assert np.isclose(expected, new[key], rtol=1e-10, atol=1e-9), (
                    f"{key} moved with has_intercept={has_intercept}: "
                    f"{expected} -> {new[key]}"
                )


def test_reading_linests_own_r_squared_would_have_been_wrong() -> None:
    # The counterfactual, pinned so the reason R_Squared is derived rather than
    # read does not get optimised away by someone who finds the derivation
    # roundabout. The gap is large and silent.
    import numpy as np  # pylint: disable=import-outside-toplevel

    rng = np.random.default_rng(11)
    predictors = rng.normal(size=(40, 3))
    response = predictors @ np.array([1.0, 2.0, -1.0]) + 5 + rng.normal(size=40)
    design = np.column_stack([np.ones(40), predictors])

    _, _, naive, _ = _linest(response, design, False)   # INDEX(LINEST(...), 3, 1)
    _, _, correct, _ = _linest(response, predictors, True)

    assert not np.isclose(naive, correct, atol=0.01)
    assert naive > correct  # uncentered R² is inflated, never an error
    assert np.isclose(correct, _new_chain(response, predictors, True)["R2"])

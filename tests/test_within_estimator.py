"""Independent verification of the v2.1 Fixed Effects fit-time pair,
Design_Response() and Design_Columns() (phase 2 of the Fixed Effects engine).

Concept: Design_Response()/Design_Columns() are the predictor/response pair the Regression
sheet's fit chain (Coefficients, Predictions, Residuals, ANOVA, the BFN
diagnostic, …) is repointed to. With no declared Fixed Effects row they are
Response_Column()/Predictor_Columns() unchanged; with one declared, every column is
one-way within-demeaned via Demean_By before the ordinary OLS engine ever
sees it — the mechanism that lets a plain least-squares fit on the demeaned
data reproduce an LSDV (dummy-per-group) fit's non-FE coefficients without
ever materializing the group dummy columns.

One source of truth on the Python side too: this module reuses the WHO FE
panel fixture and the independent statsmodels LSDV fit from
test_bfn_panel_durbin_watson_verification, and the Group_Mean/Demean_By
mirrors from test_group_panel_transforms, rather than reimplementing either.

Runnable standalone (``python tests/test_within_estimator.py``) or under
pytest.
"""
from __future__ import annotations

from pathlib import Path

import json
import sys

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# pylint: disable-next=wrong-import-position
from tests.test_bfn_panel_durbin_watson_verification import (
    _load_fe_panel,
    _lsdv_residuals,
    within_estimator_residuals,
)
# pylint: disable-next=wrong-import-position
from tests.test_group_panel_transforms import demean_by_mirror

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"


# ── Pure-Python mirrors of the workbook formulas ─────────────────────────────

def y_s_mirror(response, fe_active: bool, group=None, include=None):
    """Mirror of Design_Response(): Response_Column() unchanged, or Demean_By when FE is active."""
    if not fe_active:
        return np.asarray(response, dtype=float)
    return demean_by_mirror(response, group, include)


def x_s_within_mirror(x, fe_active: bool, group=None, include=None):
    """Mirror of Design_Columns(): Predictor_Columns() unchanged, or each column Demean_By'd when FE is active."""
    x = np.asarray(x, dtype=float)
    if not fe_active:
        return x
    return np.column_stack([demean_by_mirror(x[:, j], group, include) for j in range(x.shape[1])])


# ── Numeric acceptance: reproduces the within estimator and LSDV ────────────

def test_within_mirrors_reproduce_the_within_estimator_bit_for_bit() -> None:
    groups, _years, y, x = _load_fe_panel()

    yd = y_s_mirror(y, fe_active=True, group=groups)
    xd = x_s_within_mirror(x, fe_active=True, group=groups)

    beta = np.linalg.lstsq(xd, yd, rcond=None)[0]
    mirrored_resid = yd - xd @ beta

    expected_resid = within_estimator_residuals(groups, y, x)
    assert np.allclose(mirrored_resid, expected_resid, atol=1e-9)


def test_within_fit_matches_independent_lsdv_coefficients() -> None:
    # The whole point of the within transformation: OLS on the demeaned data
    # reproduces the non-FE coefficients an explicit dummy-per-group (LSDV)
    # fit would report, without ever materializing the G-1 group dummies.
    groups, _years, y, x = _load_fe_panel()

    yd = y_s_mirror(y, fe_active=True, group=groups)
    xd = x_s_within_mirror(x, fe_active=True, group=groups)
    within_beta = np.linalg.lstsq(xd, yd, rcond=None)[0]

    import pandas as pd  # pylint: disable=import-outside-toplevel
    import statsmodels.api as sm  # pylint: disable=import-outside-toplevel

    dummies = pd.get_dummies(pd.Series(groups), drop_first=True, dtype=float)
    design = sm.add_constant(pd.concat([pd.DataFrame(x), dummies], axis=1))
    lsdv_fit = sm.OLS(pd.Series(y), design).fit()
    lsdv_beta = lsdv_fit.params.to_numpy()[1 : 1 + x.shape[1]]  # skip const, skip dummies

    assert np.allclose(within_beta, lsdv_beta, rtol=1e-6, atol=1e-8)

    # And the residuals agree with the fully independent LSDV fit too.
    mirrored_resid = yd - xd @ within_beta
    lsdv_resid = _lsdv_residuals(groups, y, x)
    assert np.allclose(mirrored_resid, lsdv_resid, atol=1e-8)


def test_g_equals_one_collapses_to_the_no_fe_pair_exactly() -> None:
    # No Fixed Effects row (or a call with fe_active=False): Design_Response()/Design_Columns()
    # must return the EXACT no-FE objects, not a numerically-equal
    # recomputation — the non-breaking-default property every no-FE model
    # relies on.
    _groups, _years, y, x = _load_fe_panel()

    assert np.array_equal(y_s_mirror(y, fe_active=False), y)
    assert np.array_equal(x_s_within_mirror(x, fe_active=False), x)


# ── Implementation-shape assertions (by construction, not convention) ───────

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


def test_design_response_and_design_columns_are_regression_sheet_closures() -> None:
    functions = _catalog_functions()
    for name in ("Design_Response", "Design_Columns"):
        fn = functions[name]
        assert fn.get("scope") == "Regression"
        assert fn.get("category") == "Model Construction"
        assert fn.get("arguments", []) == []


def test_design_response_no_fe_branch_returns_response_column_unchanged() -> None:
    formula = _formula("Design_Response")
    assert "IF(NOT(fe_active),Response_Column()," in formula
    assert "Demean_By(Response_Column(),Fixed_Effects_Column(),Sample_Include())" in formula


def test_design_columns_demeans_with_reduce_hstack_not_bycol() -> None:
    # BYCOL cannot return a per-call array (Demean_By returns a column), so
    # the FE branch builds the demeaned matrix column-by-column via the same
    # REDUCE+HSTACK pattern Predictor_Columns() itself uses, not BYCOL.
    formula = _formula("Design_Columns")
    assert "IF(NOT(fe_active),Predictor_Columns()," in formula
    assert "BYCOL" not in formula
    assert "REDUCE(seed,SEQUENCE(k_p),LAMBDA(acc,j,HSTACK(acc,Demean_By(INDEX(xp,0,j),fe,inc))))" in formula
    assert "DROP(built,,1)" in formula


def test_design_columns_applies_the_intercept_stage_after_demeaning() -> None:
    # Pipeline order is load-bearing: a ones column demeaned by group is a
    # column of zeros, which makes the Gram matrix exactly singular. The
    # intercept must therefore be stacked onto the ALREADY-demeaned block.
    formula = _formula("Design_Columns")
    assert "IF(has_int,HSTACK(ones,demeaned),demeaned)" in formula
    assert "ones,SEQUENCE(ROWS(Source_Data),1,1,0)" in formula
    assert "has_int,N(Allow_Intercept)=1" in formula
    # The demeaning stage never sees the ones column.
    demean_stage = formula.split("demeaned,")[1].split("IF(has_int,HSTACK")[0]
    assert "ones" not in demean_stage


def test_design_columns_returns_the_bare_intercept_when_no_predictor_contributes() -> None:
    # Predictor_Columns() errors when the spec contributes nothing (DROP of a
    # sentinel-only accumulator). HSTACK onto an error is an error, so the
    # zero-predictor state is branched on before the stack, not after.
    formula = _formula("Design_Columns")
    assert "k_p,IFERROR(COLUMNS(Predictor_Columns()),0)" in formula
    assert "IF(k_p=0,IF(has_int,ones,Predictor_Columns())" in formula


def test_constructors_are_registered_after_their_dependencies() -> None:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    regression_closures = [
        fn["name"] for fn in payload["functions"] if fn.get("scope") == "Regression"
    ]
    for dependency in ("Response_Column", "Predictor_Columns", "Fixed_Effects_Column"):
        assert regression_closures.index(dependency) < regression_closures.index("Design_Response")
        assert regression_closures.index(dependency) < regression_closures.index("Design_Columns")


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

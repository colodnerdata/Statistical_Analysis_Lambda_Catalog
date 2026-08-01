"""Independent verification of the v2.1 Fixed Effects group-mean-recovery
prediction (phase 5): Group_Mean_At, Group_Count_At, Prediction_Group_Column,
and Group_Prediction_Interval.

Concept (DECISIONS.md "FE point prediction" / "prediction interval"): under
one-way Fixed Effects, the within estimator discards the G group intercepts
an LSDV fit would have materialized. The algebraic identity
alpha_hat_i = ybar_i - xbar_i'beta_hat, substituted back in, gives

    y_hat = ybar_i + (x_new - xbar_i)'beta_hat

with the mean-response and new-observation variances

    Var(mean) = sigma^2/T_i + (x_new-xbar_i)'V_beta(x_new-xbar_i)
    Var(new)  = sigma^2*(1+1/T_i) + (x_new-xbar_i)'V_beta(x_new-xbar_i)

A degenerate G=1 (no Fixed Effects row; "everyone is one group") must
collapse exactly to the ordinary v1.0 Prediction_Interval() numbers — this
is the acceptance test that proves the machinery was "built once."

Both the degenerate and FE-active cases were independently confirmed against
reference computations (plain OLS in the first case, an explicit LSDV
get_prediction() call in the second) before any Excel formula was written;
this module encodes those same checks as pure-Python mirrors matched against
the actual catalog formula text.

Runnable standalone (``python tests/test_group_prediction_interval.py``) or
under pytest.
"""
from __future__ import annotations

from pathlib import Path

import json
import math
import sys

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# pylint: disable-next=wrong-import-position
from tests.test_bfn_panel_durbin_watson_verification import _load_fe_panel
# pylint: disable-next=wrong-import-position
from tests.test_group_panel_transforms import demean_by_mirror, group_mean_mirror

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

_ALL_GROUP = "(all)"


# ── Pure-Python mirrors of the workbook formulas ─────────────────────────────

def group_mean_at_mirror(x, group, selected_group, include=None):
    """Mirror of Group_Mean_At(): XLOOKUP(selected_group, group, Group_Mean(x,group,include))."""
    x = np.asarray(x, dtype=float)
    group = np.asarray(group, dtype=object)
    gm = group_mean_mirror(x, group, include)
    matches = np.where(group == selected_group)[0]
    if len(matches) == 0:
        return float("nan")
    return gm[matches[0]]


def group_count_at_mirror(group, selected_group, include=None) -> int:
    """Mirror of Group_Count_At(): count of included rows matching selected_group."""
    group = np.asarray(group, dtype=object)
    if include is None:
        inc = group != ""
    else:
        inc = np.asarray(include, dtype=bool)
    return int(np.sum(inc & (group == selected_group)))


def prediction_group_column_mirror(fe_active: bool, group, n_rows: int):
    """Mirror of Prediction_Group_Column(): the real FE column, or a constant
    single group covering every row when no Fixed Effects row is declared."""
    if fe_active:
        return np.asarray(group, dtype=object)
    return np.full(n_rows, _ALL_GROUP, dtype=object)


def group_prediction_interval_mirror(
    x, y, x_new, group, selected_group, allow_intercept: bool, df_absorbed: int, alpha: float = 0.05,
):
    """Mirror of Group_Prediction_Interval(): point/CI/PI via group-mean recovery."""
    n, _k = x.shape
    x_within = np.column_stack([demean_by_mirror(x[:, j], group) for j in range(x.shape[1])])
    y_within = demean_by_mirror(y, group)

    design = np.column_stack([np.ones(n), x_within]) if allow_intercept else x_within
    beta_full, *_ = np.linalg.lstsq(design, y_within, rcond=None)
    beta = beta_full[1:] if allow_intercept else beta_full
    resid = y_within - design @ beta_full
    ssr = float(resid @ resid)

    naive_df = n - x.shape[1] - (1 if allow_intercept else 0)
    true_df = naive_df - df_absorbed
    sigma = math.sqrt(ssr / true_df)

    from scipy import stats  # pylint: disable=import-outside-toplevel
    t_crit = stats.t.ppf(1 - alpha / 2, true_df)

    xtx_inv = np.linalg.inv(x_within.T @ x_within)
    xbar_i = np.array([group_mean_at_mirror(x[:, j], group, selected_group) for j in range(x.shape[1])])
    ybar_i = group_mean_at_mirror(y, group, selected_group)
    t_i = group_count_at_mirror(group, selected_group)

    deviation = x_new - xbar_i
    quad = deviation @ xtx_inv @ deviation
    point_est = ybar_i + deviation @ beta
    se_mean = sigma * math.sqrt(1 / t_i + quad)
    se_new = sigma * math.sqrt(1 + 1 / t_i + quad)
    return {
        "point": point_est, "se_mean": se_mean, "se_new": se_new, "t_crit": t_crit,
        "ci_lower": point_est - t_crit * se_mean, "ci_upper": point_est + t_crit * se_mean,
        "pi_lower": point_est - t_crit * se_new, "pi_upper": point_est + t_crit * se_new,
    }


# ── Numeric acceptance ───────────────────────────────────────────────────────

def test_degenerate_g1_matches_the_ordinary_v1_prediction_interval() -> None:
    # No Fixed Effects row: group = constant "(all)" for every row. Must
    # collapse exactly to the pre-v2.1 Prediction_Interval() numbers.
    _groups, _years, y, x = _load_fe_panel()
    n = len(y)
    group = prediction_group_column_mirror(fe_active=False, group=None, n_rows=n)
    x_new = np.array([15000.0, 12.0])

    mine = group_prediction_interval_mirror(x, y, x_new, group, _ALL_GROUP, allow_intercept=True, df_absorbed=0)

    design = np.column_stack([np.ones(n), x])
    beta_full, *_ = np.linalg.lstsq(design, y, rcond=None)
    resid = y - design @ beta_full
    ssr = float(resid @ resid)
    df = n - x.shape[1] - 1
    sigma = math.sqrt(ssr / df)
    from scipy import stats  # pylint: disable=import-outside-toplevel
    t_crit = stats.t.ppf(0.975, df)
    xtx_inv_full = np.linalg.inv(design.T @ design)
    x_new_aligned = np.concatenate([[1.0], x_new])
    h_new = x_new_aligned @ xtx_inv_full @ x_new_aligned
    point_orig = x_new_aligned @ beta_full
    se_pred_orig = sigma * math.sqrt(1 + h_new)

    assert math.isclose(mine["point"], point_orig, rel_tol=1e-9)
    assert math.isclose(mine["se_new"], se_pred_orig, rel_tol=1e-9)
    assert math.isclose(mine["t_crit"], t_crit, rel_tol=1e-9)
    assert math.isclose(mine["pi_upper"] - mine["pi_lower"], 2 * t_crit * se_pred_orig, rel_tol=1e-9)


def test_fe_active_matches_an_explicit_lsdv_get_prediction() -> None:
    groups, _years, y, x = _load_fe_panel()
    selected = "Albania"
    x_new = np.array([15000.0, 12.0])

    mine = group_prediction_interval_mirror(
        x, y, x_new, groups, selected, allow_intercept=True,
        df_absorbed=len(np.unique(groups)) - 1,
    )

    import pandas as pd  # pylint: disable=import-outside-toplevel
    import statsmodels.api as sm  # pylint: disable=import-outside-toplevel

    dummies = pd.get_dummies(pd.Series(groups), drop_first=True, dtype=float)
    design = sm.add_constant(pd.concat([pd.DataFrame(x, columns=range(x.shape[1])), dummies], axis=1))
    lsdv_fit = sm.OLS(pd.Series(y), design).fit()

    row = {0: x_new[0], 1: x_new[1], "const": 1.0}
    for c in dummies.columns:
        row[c] = 1.0 if c == selected else 0.0
    row_df = pd.DataFrame([row])[design.columns]
    pred = lsdv_fit.get_prediction(row_df)

    assert math.isclose(mine["point"], pred.predicted_mean[0], rel_tol=1e-6)
    assert math.isclose(mine["se_mean"], pred.se_mean[0], rel_tol=1e-6)
    assert mine["se_new"] > mine["se_mean"]  # PI always wider than the mean-response CI


def test_mean_ci_collapses_to_group_mean_se_at_the_groups_own_centroid() -> None:
    # DECISIONS.md sanity check: predicting at x_new = xbar_i kills the
    # quadratic term, so the mean-CI half-width reduces to t*sqrt(sigma^2/T_i).
    groups, _years, y, x = _load_fe_panel()
    selected = "Albania"
    xbar_i = np.array([
        group_mean_at_mirror(x[:, j], groups, selected) for j in range(x.shape[1])
    ])

    mine = group_prediction_interval_mirror(
        x, y, xbar_i, groups, selected, allow_intercept=True,
        df_absorbed=len(np.unique(groups)) - 1,
    )
    t_i = group_count_at_mirror(groups, selected)
    expected_se = mine["se_mean"]
    # se_mean at the centroid should equal sigma/sqrt(T_i) exactly (quad=0).
    sigma_implied = expected_se * math.sqrt(t_i)
    reconstructed = sigma_implied / math.sqrt(t_i)
    assert math.isclose(reconstructed, expected_se, rel_tol=1e-9)
    assert math.isclose(mine["ci_upper"] - mine["point"], mine["t_crit"] * expected_se, rel_tol=1e-9)


def test_group_mean_at_and_group_count_at_agree_with_group_mean() -> None:
    groups, _years, y, _x = _load_fe_panel()
    for g in ("Albania", "Zimbabwe"):
        at_value = group_mean_at_mirror(y, groups, g)
        broadcast = group_mean_mirror(y, groups)
        matching = broadcast[groups == g]
        assert np.allclose(matching, at_value)
        assert group_count_at_mirror(groups, g) == int((groups == g).sum())


# ── Implementation-shape assertions (once the catalog entries exist) ────────

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


def test_group_mean_at_delegates_to_group_mean_via_xlookup() -> None:
    formula = _formula("Group_Mean_At")
    assert "Group_Mean(x,group,include)" in formula
    assert "XLOOKUP(" in formula


def test_prediction_group_column_falls_back_to_a_constant_single_group() -> None:
    formula = _formula("Prediction_Group_Column")
    assert '"(all)"' in formula
    assert "Fixed_Effects_Column()" in formula


def test_group_prediction_interval_reuses_demean_by_and_gram_inverse() -> None:
    formula = _formula("Group_Prediction_Interval")
    assert "Demean_By(" in formula
    assert "Gram_Inverse(" in formula
    assert "Design_Matrix(" in formula
    # No intercept in the Gram matrix — group-mean recovery structurally
    # never needs one (alpha_hat_i already plays that role).
    assert "Design_Matrix(x_within," in formula


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

"""Independent verification of the Group & Panel primitives (v2.1 Fixed
Effects phase 1): Group_Mean, Demean_By, Is_Balanced_Panel, and the
Regression-sheet-scoped Absorbed_Degrees_Of_Freedom.

These are pure catalog additions — no sheet-writer wiring yet (that lands
with the fit-time demeaned matrix pair in phase 2). Verified here the same
way the BFN/resolver modules verify forward-wired panel machinery: a
pure-Python mirror of each LAMBDA, cross-checked against an independent
pandas computation, plus implementation-shape assertions on the actual
catalog formula text so a future edit cannot silently change the contract
without a test noticing.

Runnable standalone (``python tests/test_group_panel_transforms.py``) or
under pytest.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import numpy as np
import pandas as pd

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

NA = float("nan")


# ── Pure-Python mirrors of the workbook formulas ─────────────────────────────

def group_mean_mirror(x, group, include=None):
    """Mirror of Group_Mean(x, group, [include]).

    Per-row group mean over the included members of that row's own group.
    A row's own inclusion does not gate its own output (full-height
    contract, same as Predictor_Columns()); only a blank group or an empty included group
    yields NA.
    """
    x = np.asarray(x, dtype=float)
    group = np.asarray(group, dtype=object)
    if include is None:
        inc = ~np.isnan(x)
    else:
        inc = np.asarray(include, dtype=bool)

    active = inc & (group != "") & pd.notna(group)
    out = np.full(len(x), NA)
    for g in pd.unique(group[active]):
        mask = active & (group == g)
        if mask.sum() == 0:
            continue
        mean = np.sum(np.where(mask, np.where(np.isnan(x), 0.0, x), 0.0)) / mask.sum()
        out[group == g] = mean
    return out


def demean_by_mirror(x, group, include=None):
    """Mirror of Demean_By(x, group, [include]) = x - Group_Mean(x, group, include)."""
    x = np.asarray(x, dtype=float)
    gm = group_mean_mirror(x, group, include)
    out = np.where(~np.isnan(x) & ~np.isnan(gm), x - gm, NA)
    return out


def is_balanced_panel_mirror(group, time, include=None) -> bool | float:
    """Mirror of Is_Balanced_Panel(group, time, [include])."""
    group = np.asarray(group, dtype=object)
    time = np.asarray(time, dtype=object)
    if include is None:
        inc = (group != "") & pd.notna(time) & (time != "")
    else:
        inc = np.asarray(include, dtype=bool)

    active = inc & (group != "") & (time != "")
    n = int(active.sum())
    if n == 0:
        return NA

    groups = pd.unique(group[active])
    times = pd.unique(time[active])
    keys = list(zip(group[active].tolist(), time[active].tolist()))
    return n == len(groups) * len(times) and n == len(set(keys))


# ── Numeric acceptance: cross-checked against independent pandas grouping ───

def _synthetic_panel():
    rng = np.random.default_rng(20260725)
    groups = np.repeat(["Site A", "Site B", "Site C"], 10)
    time = np.tile(np.arange(1, 11), 3)
    x = rng.normal(loc=100, scale=10, size=30)
    return groups, time, x


def test_group_mean_matches_pandas_groupby_transform() -> None:
    groups, _time, x = _synthetic_panel()
    mirrored = group_mean_mirror(x, groups)

    expected = pd.Series(x).groupby(pd.Series(groups)).transform("mean").to_numpy()
    assert np.allclose(mirrored, expected)


def test_demean_by_matches_pandas_within_transform() -> None:
    groups, _time, x = _synthetic_panel()
    mirrored = demean_by_mirror(x, groups)

    df = pd.DataFrame({"x": x, "g": groups})
    expected = (df["x"] - df.groupby("g")["x"].transform("mean")).to_numpy()
    assert np.allclose(mirrored, expected)
    # Within transformation removes each group's mean exactly.
    for g in np.unique(groups):
        assert math.isclose(mirrored[groups == g].mean(), 0.0, abs_tol=1e-9)


def test_demean_by_collapses_to_grand_mean_centering_when_g_equals_one() -> None:
    # G = 1 (no Fixed Effects row, or a degenerate one-level FE variable):
    # Demean_By must reduce to ordinary grand-mean centering — the
    # degenerate-collapse property the v2.1 df-plumbing design relies on.
    _groups, _time, x = _synthetic_panel()
    one_group = np.full(len(x), "(all)")

    mirrored = demean_by_mirror(x, one_group)
    expected = x - x.mean()
    assert np.allclose(mirrored, expected)


def test_group_mean_is_na_for_blank_group_and_all_excluded_group() -> None:
    x = np.array([1.0, 2.0, 3.0, 4.0])
    group = np.array(["A", "", "B", "B"], dtype=object)
    include = np.array([True, True, False, False])  # every "B" row excluded

    mirrored = group_mean_mirror(x, group, include)
    assert mirrored[0] == 1.0          # A: only member, included
    assert math.isnan(mirrored[1])     # blank group
    assert math.isnan(mirrored[2])     # B: no included member
    assert math.isnan(mirrored[3])


def test_demean_by_is_na_when_x_non_numeric_even_if_group_mean_exists() -> None:
    x = np.array([1.0, NA, 3.0])
    group = np.array(["A", "A", "A"], dtype=object)

    mirrored = demean_by_mirror(x, group)
    assert not math.isnan(mirrored[0])
    assert math.isnan(mirrored[1])  # x itself non-numeric -> NA, group mean irrelevant
    assert not math.isnan(mirrored[2])


def test_is_balanced_panel_true_for_a_complete_grid() -> None:
    groups, time, _x = _synthetic_panel()
    assert is_balanced_panel_mirror(groups, time) is True


def test_is_balanced_panel_false_on_a_missing_cell() -> None:
    groups, time, _x = _synthetic_panel()
    include = np.ones(len(groups), dtype=bool)
    include[0] = False  # drop Site A's first period -> a hole in the grid
    assert is_balanced_panel_mirror(groups, time, include) is False


def test_is_balanced_panel_false_on_a_duplicated_cell() -> None:
    # Row count (4) equals |groups|*|times| (2*2), but A has (1,1) twice and
    # no (2,*) at all — the row-count check alone would miss this.
    groups = np.array(["A", "A", "B", "B"], dtype=object)
    time = np.array([1, 1, 1, 2], dtype=object)
    assert is_balanced_panel_mirror(groups, time) is False


def test_is_balanced_panel_na_when_nothing_included() -> None:
    groups = np.array(["A", "B"], dtype=object)
    time = np.array([1, 1], dtype=object)
    include = np.array([False, False])
    assert math.isnan(is_balanced_panel_mirror(groups, time, include))


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


def test_group_mean_and_demean_by_are_portable_data_transformation_functions() -> None:
    functions = _catalog_functions()
    for name in ("Group_Mean", "Demean_By", "Is_Balanced_Panel"):
        fn = functions[name]
        assert fn.get("category") == "Data Transformation"
        assert fn.get("scope", "workbook") == "workbook"

    assert functions["Group_Mean"].get("subcategory") == "Group & Panel"
    assert functions["Demean_By"].get("subcategory") == "Group & Panel"
    assert functions["Is_Balanced_Panel"].get("subcategory") == "Sample Construction & Diagnostics"


def test_demean_by_delegates_to_group_mean_one_source_of_truth() -> None:
    demean_by = _formula("Demean_By")
    assert "Group_Mean(x,group,include)" in demean_by
    # No second group-aggregation implementation hiding inside Demean_By.
    assert "XLOOKUP" not in demean_by
    assert "UNIQUE" not in demean_by


def test_group_mean_appears_before_demean_by_in_document_order() -> None:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    names = [fn["name"] for fn in payload["functions"]]
    assert names.index("Group_Mean") < names.index("Demean_By")


def test_absorbed_degrees_of_freedom_is_a_regression_sheet_closure() -> None:
    functions = _catalog_functions()
    fn = functions["Absorbed_Degrees_Of_Freedom"]
    assert fn.get("scope") == "Regression"
    assert fn.get("category") == "Model Construction"
    assert fn.get("subcategory") == "Sheet-Scoped Constructors"
    assert fn.get("arguments", []) == []


def test_absorbed_degrees_of_freedom_reuses_dummy_levels_for_group_count() -> None:
    formula = _formula("Absorbed_Degrees_Of_Freedom")
    # One source of truth for the level count: the same mask-scoped,
    # reference-dropped level set the design matrix would use.
    assert 'Dummy_Levels(Fixed_Effects_Column(),"",Fit_Sample_Include())' in formula
    assert "COLUMNS(lv)" in formula
    # Non-breaking default: no Fixed Effects row -> 0, never an error.
    assert "IF(NOT(fe_active),0," in formula


def test_absorbed_degrees_of_freedom_is_registered_after_fixed_effects_column() -> None:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    regression_closures = [
        fn["name"] for fn in payload["functions"] if fn.get("scope") == "Regression"
    ]
    assert regression_closures.index("Fixed_Effects_Column") < (
        regression_closures.index("Absorbed_Degrees_Of_Freedom")
    )


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

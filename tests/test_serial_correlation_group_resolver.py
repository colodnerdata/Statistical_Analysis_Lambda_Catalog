"""Independent verification of the serial-correlation grouping-key resolver
(Serial_Correlation_Group).

Concept (restated as working notes, per the v2.1 design record §3): the
serial-correlation diagnostics need a grouping key — which dimension
partitions the residuals for within-group differencing. Today that is the
Fixed Effects column; at v3.5 a Cluster role will supply it for
pooled-panel diagnostics without absorption. To make that future additive,
the resolver ships NOW with a dormant, unreachable Cluster branch — the
switch that currently does nothing, mirroring the reserved-spec-column
pattern (Order/Transform named at grid-freeze, read by nothing). The
resolver is the single reference point for the grouping key: the BFN
wiring consumes Serial_Correlation_Group(), never Fixed_Effects_Column()
directly, so retargeting the key at v3.5 is a one-name change with zero
edits at the consumer.

Resolver contract, verified here:

  Role="Fixed Effects" declared  → the FE group column (via
                                   Fixed_Effects_Column() — the resolver
                                   dispatches, the accessor looks up)
  Role="Cluster" declared, no FE → the not-applicable token
                                   "n/a — Cluster grouping RESERVED"
                                   (dormant branch: present in the SWITCH,
                                   returns a token, never an error)
  neither role declared          → the none-sentinel "none"
                                   (→ ordinary DW path)

Acceptance is a PURE REFACTOR of the PR-105 BFN wiring: this module asserts
BIT-IDENTICAL results — the resolver hands back the very FE column object
the PR-105 path referenced directly, so every PR-105 numeric acceptance
(the pinned WHO BFN, the seam invariance, the gap-skip term counts) is
reproduced exactly, `==` not merely isclose. The PR-105 suite itself
(test_bfn_panel_durbin_watson_verification) continues to run unchanged.

Runnable standalone (``python tests/test_serial_correlation_group_resolver.py``)
or under pytest.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# One source of truth on the Python side too: the PR-105 fixtures and the
# BFN workbook mirror are imported, never re-implemented here — bit-identity
# against PR 105 is only meaningful if both paths share the same mirrors.
# pylint: disable-next=wrong-import-position
from tests.test_bfn_panel_durbin_watson_verification import (
    _COMPUTES_DW,
    _DELTA,
    _EXPECTED_BFN,
    _EXPECTED_NUMERATOR_TERMS,
    _diagnostic_cell_formulas,
    _load_fe_panel,
    bfn_cell_mirror,
    bfn_workbook_mirror,
    dw_cell_mirror,
    within_estimator_residuals,
)

# pylint: disable-next=wrong-import-position
from tests.test_difference_by_verification import NA

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

# The exact spec tokens (write_spec_block role constants).
_ROLE_FIXED_EFFECTS = "Fixed Effects"
_ROLE_CLUSTER = "Cluster"
_ROLE_RESPONSE = "Response (y)"
_ROLE_PREDICTOR = "Predictor (x)"

# The resolver's two non-column results, exactly as the LAMBDA returns them.
NONE_SENTINEL = "none"
CLUSTER_RESERVED_TOKEN = "n/a — Cluster grouping RESERVED"


# ── Pure-Python mirrors of the workbook formulas ────────────────────────────

def fixed_effects_column_mirror(roles, columns):
    """Mirror of Fixed_Effects_Column(): the first Role="Fixed Effects" column.

    XMATCH exact-match semantics — first match wins; NA when no row carries
    the role (the callers' count gates fire before that is ever consumed).
    """
    for role, column in zip(roles, columns):
        if role == _ROLE_FIXED_EFFECTS:
            return column
    return NA


def serial_correlation_group_mirror(roles, columns):
    """Mirror of the Serial_Correlation_Group() LAMBDA — dispatch only.

    key = IFS(FE count > 0 → "Fixed Effects", Cluster count > 0 → "Cluster",
    TRUE → "none"), then SWITCH(key): the FE branch routes through the
    Fixed_Effects_Column mirror (one lookup implementation, here as in the
    catalog), the Cluster branch returns the RESERVED token, and the
    default is the none-sentinel.
    """
    fe_vars = sum(1 for role in roles if role == _ROLE_FIXED_EFFECTS)
    cluster_vars = sum(1 for role in roles if role == _ROLE_CLUSTER)
    if fe_vars > 0:
        key = _ROLE_FIXED_EFFECTS
    elif cluster_vars > 0:
        key = _ROLE_CLUSTER
    else:
        key = NONE_SENTINEL

    if key == _ROLE_FIXED_EFFECTS:
        return fixed_effects_column_mirror(roles, columns)
    if key == _ROLE_CLUSTER:
        return CLUSTER_RESERVED_TOKEN  # dormant branch — RESERVED
    return NONE_SENTINEL


def _who_fe_spec(groups, years, y, x):
    """The small WHO FE spec as aligned (roles, columns) vectors.

    Life expectancy ~ GDP + Schooling, FE = Country, Sequence = Year — the
    exact model the PR-105 acceptance pinned.
    """
    roles = [
        _ROLE_FIXED_EFFECTS,   # Country
        _ROLE_PREDICTOR,       # Year (also the Sequence axis — role-independent)
        _ROLE_RESPONSE,        # Life expectancy
        _ROLE_PREDICTOR,       # GDP
        _ROLE_PREDICTOR,       # Schooling
    ]
    columns = [groups, years, y, x[:, 0], x[:, 1]]
    return roles, columns


# ── Acceptance 1: pure refactor — bit-identical to the PR-105 path ──────────

def test_resolver_hands_back_the_fe_column_itself() -> None:
    groups, years, y, x = _load_fe_panel()
    roles, columns = _who_fe_spec(groups, years, y, x)

    resolved = serial_correlation_group_mirror(roles, columns)

    # Not a copy, not a recomputation: THE Fixed Effects column, identically
    # what Fixed_Effects_Column() returned on the PR-105 path. Every
    # downstream PR-105 numeric (pinned BFN, seam invariance, gap-skip term
    # counts) is therefore reproduced exactly, by construction.
    assert resolved is groups
    assert resolved is fixed_effects_column_mirror(roles, columns)


def test_bfn_through_the_resolver_is_bit_identical_to_pr105() -> None:
    groups, years, y, x = _load_fe_panel()
    e = within_estimator_residuals(groups, y, x)

    # PR-105 wiring: the group argument is the FE column, referenced directly.
    bfn_direct, terms_direct = bfn_workbook_mirror(e, groups, years, _DELTA)

    # Refactored wiring: the group argument is the resolver's output.
    roles, columns = _who_fe_spec(groups, years, y, x)
    resolved = serial_correlation_group_mirror(roles, columns)
    bfn_resolved, terms_resolved = bfn_workbook_mirror(e, resolved, years, _DELTA)

    # Bit-identical, not merely close — the refactor moved a reference, not
    # a computation.
    assert bfn_resolved == bfn_direct
    assert terms_resolved == terms_direct == _EXPECTED_NUMERATOR_TERMS
    # And both sit on the PR-105 release pin.
    assert math.isclose(bfn_resolved, _EXPECTED_BFN, rel_tol=1e-9)


def test_gapped_panel_through_the_resolver_matches_pr105_term_counts() -> None:
    # The PR-105 punched-out-year acceptance, re-run through the resolver
    # path: Albania minus 2007 loses exactly two numerator terms, never a
    # bridged fake one-year change.
    groups, years, y, x = _load_fe_panel(punch_albania_2007=True)
    e = within_estimator_residuals(groups, y, x)
    roles, columns = _who_fe_spec(groups, years, y, x)

    resolved = serial_correlation_group_mirror(roles, columns)
    bfn_resolved, term_count = bfn_workbook_mirror(e, resolved, years, _DELTA)
    bfn_direct, terms_direct = bfn_workbook_mirror(e, groups, years, _DELTA)

    assert bfn_resolved == bfn_direct
    assert term_count == terms_direct == _EXPECTED_NUMERATOR_TERMS - 2


# ── Acceptance 2: the dormant Cluster branch — token path, verified inert ───

def test_cluster_declared_spec_takes_the_token_path_not_an_error() -> None:
    groups, years, y, x = _load_fe_panel()
    # Same spec, but the grouping role is Cluster (hand-typed: the dropdown
    # does not offer it) and no Fixed Effects row exists.
    roles = [_ROLE_CLUSTER, _ROLE_PREDICTOR, _ROLE_RESPONSE,
             _ROLE_PREDICTOR, _ROLE_PREDICTOR]
    columns = [groups, years, y, x[:, 0], x[:, 1]]

    resolved = serial_correlation_group_mirror(roles, columns)

    # The reserved branch answers with the documented token — a string a
    # human (or a future Model-Comparison XLOOKUP) can read — never #N/A or
    # a column of the wrong dimension.
    assert resolved == CLUSTER_RESERVED_TOKEN
    assert resolved is not NA
    assert "RESERVED" in resolved


def test_cluster_branch_is_unreachable_from_the_shipped_cells() -> None:
    # Dormant AND unreachable: a Cluster row contributes zero to the
    # fe_vars count the cells gate on, so the trigger matrix sits in exactly
    # the PR-105 no-FE state — BFN shows its own token before the group
    # argument (the resolver) is ever evaluated, and DW still computes.
    assert bfn_cell_mirror(1, 0) == "n/a — no fixed effects"
    assert dw_cell_mirror(1, 0) == _COMPUTES_DW
    # No Sequence axis: both cells token, exactly as in PR 105.
    assert bfn_cell_mirror(0, 0) == "n/a — requires Sequence"
    assert dw_cell_mirror(0, 0) == "n/a — requires Sequence"


def test_fixed_effects_wins_when_both_grouping_roles_are_declared() -> None:
    # Precedence mirrors the shipped gating: FE presence drives the BFN cell
    # (fe_vars > 0), so a stray Cluster row must not steal the key from a
    # declared FE variable.
    groups, years, y, x = _load_fe_panel()
    roles = [_ROLE_FIXED_EFFECTS, _ROLE_CLUSTER, _ROLE_RESPONSE,
             _ROLE_PREDICTOR, _ROLE_PREDICTOR]
    columns = [groups, years, y, x[:, 0], x[:, 1]]

    assert serial_correlation_group_mirror(roles, columns) is groups


def test_no_grouping_role_returns_the_none_sentinel() -> None:
    # The shipped workbook's state (no FE in the dropdown yet, no Cluster):
    # the resolver signals "no grouping key" — the ordinary DW path — with
    # an explicit sentinel, never an error.
    groups, years, y, x = _load_fe_panel()
    roles = ["Identifier (Row Label)", _ROLE_PREDICTOR, _ROLE_RESPONSE,
             _ROLE_PREDICTOR, _ROLE_PREDICTOR]
    columns = [groups, years, y, x[:, 0], x[:, 1]]

    assert serial_correlation_group_mirror(roles, columns) == NONE_SENTINEL


# ── Implementation-shape assertions (by construction, not convention) ───────

def _catalog_formula(name: str) -> str:
    """A catalog formula exactly as the writer installs it as a defined name
    (whitespace compacted OUTSIDE string literals, so the tokens survive)."""
    # pylint: disable-next=import-outside-toplevel
    from lambda_catalog.lambda_formula_parser import (
        _normalize_user_formula,
        _strip_non_string_whitespace,
    )

    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    formulas = {fn["name"]: fn["formula_display"] for fn in payload["functions"]}
    return _strip_non_string_whitespace(_normalize_user_formula(formulas[name]))


def test_resolver_lambda_is_a_switch_with_the_dormant_cluster_branch() -> None:
    resolver = _catalog_formula("Serial_Correlation_Group")

    # Dispatch is a SWITCH on the declared grouping role; the FE branch
    # routes through Fixed_Effects_Column() (the resolver dispatches, the
    # accessor looks up — one lookup implementation), the Cluster branch is
    # present but returns the reserved token, and the default is the
    # none-sentinel.
    assert (
        'SWITCH(key,'
        '"Fixed Effects",Fixed_Effects_Column(),'
        '"Cluster","n/a — Cluster grouping RESERVED",'
        '"none")'
    ) in resolver
    # In-formula documentation of the dormant branch.
    assert "RESERVED" in resolver
    # No second lookup implementation hiding inside the resolver.
    assert "XMATCH" not in resolver
    assert "INDEX(" not in resolver
    # Key precedence: FE tested before Cluster, sentinel as the IFS fallback.
    assert (
        'IFS('
        'SUMPRODUCT(N(roles="Fixed Effects"))>0,"Fixed Effects",'
        'SUMPRODUCT(N(roles="Cluster"))>0,"Cluster",'
        'TRUE,"none")'
    ) in resolver
    # Same TAKE-trimmed spec-band idiom as every other sheet closure.
    assert "TAKE(Spec_Role,COLUMNS(Source_Data))" in resolver


def test_resolver_is_registered_after_the_accessor_it_dispatches_to() -> None:
    # Sheet-scoped closures install in document (= dependency) order, so the
    # resolver must follow Fixed_Effects_Column in the catalog.
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    regression_closures = [
        fn["name"]
        for fn in payload["functions"]
        if fn.get("scope") == "Regression"
    ]
    assert regression_closures.index("Fixed_Effects_Column") < (
        regression_closures.index("Serial_Correlation_Group")
    )


def test_bfn_cell_consumes_the_resolver_as_the_single_retarget_point() -> None:
    dw, bfn = _diagnostic_cell_formulas()

    # The group argument is the resolver — the ONE reference point for the
    # grouping key; the FE accessor no longer appears at the consumer.
    assert (
        "BFN_Panel_Durbin_Watson(Design_Columns(),Design_Response(),"
        "Serial_Correlation_Group(),Sequence_Column(),Base_Period_Delta(),"
        "Sample_Include())"
    ) in bfn
    assert "Fixed_Effects_Column(" not in bfn
    # The ordinary DW path takes no grouping key: the resolver does not leak
    # into the single-series cell.
    assert "Serial_Correlation_Group" not in dw
    assert "Fixed_Effects_Column(" not in dw
    # The PR-105 guard structure is untouched by the refactor (the PR-105
    # suite asserts the full precedence order; this pins the counts still
    # gate BEFORE the resolver can be evaluated).
    for guard in ('IF(seq_flags=0,', 'IF(seq_flags>1,',
                  'IF(fe_vars=0,', 'IF(fe_vars>1,'):
        assert guard in bfn


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

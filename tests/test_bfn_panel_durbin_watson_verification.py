"""Independent verification of the BFN panel Durbin-Watson (BFN_Panel_Durbin_Watson).

Concept (restated as working notes, per the v2.1 design record §2):
under fixed effects, ordinary Durbin-Watson is invalid twice over — the
residuals are within-demeaned, and a panel has no single ordering, so
naive row-adjacent differencing manufactures correlation at every group
seam (unit A's last period sits next to unit B's first). The
Bhargava–Franzini–Narendranathan (1982) panel form restricts the
numerator's differencing to within-group pairs:

    BFN = Σᵢ Σₜ₌₂ (û(i,t) − û(i,t−Δ))² ÷ Σᵢ Σₜ û(i,t)²

Same "near 2 ⇒ no first-order autocorrelation" reading; its own
N,T-dependent critical values (standard DW bounds do NOT apply — surfacing
BFN bounds is a recorded open item, the shipped cell carries an
interpretation-caveat note instead).

This module verifies the release along two independent paths that share no
implementation:

  (A) an independent fit-and-difference: statsmodels LSDV (per-country
      dummies) residuals, BFN by an explicit per-group year-keyed loop;
  (B) a pure-Python mirror of the workbook formula: manual within-estimator
      (group-demeaning) residuals fed through the Difference_By mirror from
      test_difference_by_verification (one source of truth for the
      differencing semantics there as in the catalog), NA→0 masked in the
      numerator exactly as the LAMBDA's SUMPRODUCT(IFERROR(d,0)^2) does.

It asserts (A) == (B) to ≥ 10 significant figures, pins the release value,
asserts the seam property directly (BFN is invariant to permuting whole
country blocks, while naive row-adjacent DW is not), asserts the
punched-out-year gap-skip inherited from Difference_By (13 numerator terms
for the gapped group, not 14), and exercises the diagnostic trigger matrix
on all four states against the shipped cell formulas.

Runnable standalone (``python tests/test_bfn_panel_durbin_watson_verification.py``)
or under pytest.
"""
from __future__ import annotations

import csv
import json
import math
import sys
from pathlib import Path
from typing import cast

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# One source of truth for the gap-aware differencing semantics on the Python
# side too: the Difference_By mirror is imported, never re-implemented here.
# pylint: disable-next=wrong-import-position
from tests.test_difference_by_verification import EXCLUDED, NA, difference_by

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

# The WHO file's exact headers (note the trailing space on the response).
_COUNTRY = "Country"
_YEAR = "Year"
_LIFE_EXPECTANCY = "Life expectancy "
_PREDICTORS = ("GDP", "Schooling")

_DELTA = 1  # the WHO panel's base period (MODE of within-country spacings)

# Release numbers for the small FE model (Life expectancy ~ GDP + Schooling,
# FE = Country, Sequence = Year) under role-aware completeness — rows numeric
# on the response and both continuous predictors. 157 countries keep at least
# one complete row (the 193-country panel loses 36 entirely to GDP/Schooling
# sparsity); their masked year runs happen to be internally gap-free, so the
# numerator has exactly one skipped (first-period) term per country.
_EXPECTED_ROWS = 2482
_EXPECTED_COUNTRIES = 157
_EXPECTED_NUMERATOR_TERMS = _EXPECTED_ROWS - _EXPECTED_COUNTRIES  # 2325
_EXPECTED_BFN = 0.6362023311147436


def _as_number(value):
    text = (value or "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


# ── Fixture loading ─────────────────────────────────────────────────────────

def _load_fe_panel(punch_albania_2007: bool = False):
    """Masked (group, year, y, X) arrays for the small WHO FE model.

    Role-aware completeness: a row enters iff the response and every included
    continuous predictor are numeric (Country and Year are always present in
    the WHO file). ``punch_albania_2007`` deletes Albania's 2007 row to create
    a genuine within-panel gap on the Δ = 1 grid.
    """
    groups: list[str] = []
    years: list[float] = []
    y: list[float] = []
    x: list[list[float]] = []
    with CSV_PATH.open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            values = [_as_number(row[_LIFE_EXPECTANCY])]
            values += [_as_number(row[p]) for p in _PREDICTORS]
            year = _as_number(row[_YEAR])
            if any(v is None for v in values) or year is None:
                continue
            if (
                punch_albania_2007
                and row[_COUNTRY] == "Albania"
                and year == 2007
            ):
                continue
            groups.append(row[_COUNTRY])
            years.append(year)
            y.append(cast(float, values[0]))
            x.append(cast(list[float], values[1:]))
    return (
        np.array(groups),
        np.array(years, dtype=np.float64),
        np.array(y, dtype=np.float64),
        np.array(x, dtype=np.float64),
    )


# ── Path B: pure-Python mirror of the workbook formula ──────────────────────

def within_estimator_residuals(groups, y, x):
    """One-way FE residuals by the within (group-demeaning) estimator.

    Demean y and every predictor within each group over the masked rows, then
    OLS through the origin on the demeaned data. Numerically identical to
    LSDV residuals — the equivalence path A leans on from the other side.
    """
    yd = y.astype(np.float64).copy()
    xd = x.astype(np.float64).copy()
    for group in np.unique(groups):
        mask = groups == group
        yd[mask] -= y[mask].mean()
        xd[mask] -= x[mask].mean(axis=0)
    beta = np.linalg.lstsq(xd, yd, rcond=None)[0]
    return yd - xd @ beta


def bfn_workbook_mirror(residuals, groups, years, delta):
    """Mirror of the BFN_Panel_Durbin_Watson LAMBDA on given residuals.

    numerator   = SUMPRODUCT(IFERROR(Difference_By(e, group, seq, Δ), 0)^2)
    denominator = SUMPRODUCT(e^2)

    The NA→0 masking happens HERE, locally and visibly (the consumer-owns-
    masking rule): Difference_By returns #N/A at each group's first period
    and across gaps, and exactly those terms are zeroed out of the numerator.
    The denominator keeps every residual — first periods and gap rows still
    count there, as the BFN definition requires. The mask never turns an
    error STATE into a number: with zero computable differences (all groups
    singletons, or Δ unresolved) the LAMBDA's n_terms guard returns #N/A
    rather than a fake BFN = 0.

    Returns (bfn, numerator_term_count); bfn is NA when no term is computable.
    """
    diffs = difference_by(list(residuals), list(groups), list(years), delta)
    # An excluded ("") row cannot occur on the sheet's filtered inputs and
    # would surface as a visible #VALUE! there, never a silent drop.
    assert all(d is not EXCLUDED for d in diffs)
    term_count = sum(1 for d in diffs if isinstance(d, float))
    if term_count == 0:  # IF(n_terms = 0, NA(), ...)
        return NA, 0
    masked = [d if isinstance(d, float) else 0.0 for d in diffs]  # IFERROR(d,0)
    numerator = sum(d * d for d in masked)
    denominator = float(np.sum(np.asarray(residuals) ** 2))
    return numerator / denominator, term_count


# ── Path A: independent fit and differencing (no shared code) ───────────────

def _lsdv_residuals(groups, y, x):
    """LSDV residuals via statsmodels: constant + drop-first country dummies."""
    import pandas as pd  # pylint: disable=import-outside-toplevel
    import statsmodels.api as sm  # pylint: disable=import-outside-toplevel

    dummies = pd.get_dummies(pd.Series(groups), drop_first=True, dtype=float)
    design = sm.add_constant(
        pd.concat([pd.DataFrame(x, columns=list(_PREDICTORS)), dummies], axis=1)
    )
    fit = sm.OLS(pd.Series(y), design).fit()
    return fit.resid.to_numpy()


def bfn_independent(residuals, groups, years, delta):
    """BFN by an explicit per-(group, year) pair walk — not Difference_By.

    Keys residuals on (group, year) and pairs each with (group, year − Δ)
    when that exact period exists; a gap or first period simply contributes
    no numerator term. Returns (bfn, numerator_term_count).
    """
    keyed = {
        (g, t): e for g, t, e in zip(groups, years, residuals)
    }
    numerator = 0.0
    term_count = 0
    for (g, t), e in keyed.items():
        prior = keyed.get((g, t - delta))
        if prior is not None:
            numerator += (e - prior) ** 2
            term_count += 1
    return numerator / float(np.sum(residuals**2)), term_count


# ── Acceptance 1: Python-vs-workbook-formula agreement (≥ 10 sig figs) ──────

def test_who_fe_model_masks_the_documented_panel() -> None:
    groups, years, y, _ = _load_fe_panel()
    assert len(y) == _EXPECTED_ROWS
    assert len(np.unique(groups)) == _EXPECTED_COUNTRIES
    # Δ = 1 is the panel's base period: every within-country spacing on the
    # masked rows is exactly one year (the masked runs are gap-free), so the
    # numerator drops exactly one first-period term per country.
    keyed = set(zip(groups, years))
    contiguous = sum(1 for (g, t) in keyed if (g, t - _DELTA) in keyed)
    assert contiguous == _EXPECTED_NUMERATOR_TERMS


def test_bfn_workbook_mirror_matches_independent_fit_to_ten_sig_figs() -> None:
    groups, years, y, x = _load_fe_panel()

    # Path B: the workbook formula's semantics, end to end.
    e_within = within_estimator_residuals(groups, y, x)
    bfn_mirror, terms_mirror = bfn_workbook_mirror(e_within, groups, years, _DELTA)

    # Path A: independent estimator (LSDV) and independent differencing.
    e_lsdv = _lsdv_residuals(groups, y, x)
    bfn_indep, terms_indep = bfn_independent(e_lsdv, groups, years, _DELTA)

    # The two fits are algebraically identical (within ≡ LSDV residuals).
    assert float(np.max(np.abs(e_within - e_lsdv))) < 1e-8

    # ≥ 10 significant figures of agreement between the workbook formula's
    # value and the independently computed statistic.
    assert math.isclose(bfn_mirror, bfn_indep, rel_tol=1e-10)
    assert terms_mirror == terms_indep == _EXPECTED_NUMERATOR_TERMS

    # Pin the release number (strong positive within-autocorrelation — level
    # series; well below 2, as expected for this panel).
    assert math.isclose(bfn_mirror, _EXPECTED_BFN, rel_tol=1e-9)


# ── Acceptance 2: the seam property (group-block permutation invariance) ────

def test_bfn_is_invariant_to_permuting_country_blocks() -> None:
    groups, years, y, x = _load_fe_panel()
    e = within_estimator_residuals(groups, y, x)

    # Stack the panel country-by-country (each block year-ascending), then
    # permute the ORDER OF THE BLOCKS deterministically.
    order = np.lexsort((years, groups))
    g_s, t_s, e_s = groups[order], years[order], e[order]
    rng = np.random.default_rng(19820601)  # Bhargava et al., Rev. Econ. Stud. 1982
    permuted_countries = rng.permutation(np.unique(g_s))
    perm = np.concatenate([np.where(g_s == c)[0] for c in permuted_countries])

    bfn_stacked, _ = bfn_workbook_mirror(e_s, g_s, t_s, _DELTA)
    bfn_permuted, _ = bfn_workbook_mirror(e_s[perm], g_s[perm], t_s[perm], _DELTA)

    # Within-group differencing makes the statistic blind to block order
    # (identical up to float summation order).
    assert math.isclose(bfn_stacked, bfn_permuted, rel_tol=1e-12)

    # The contrast that motivates BFN: NAIVE row-adjacent differencing reads
    # a seam pair (country A's 2015 next to country B's 2000) at every block
    # boundary, so the same permutation MOVES the naive statistic.
    def naive_row_adjacent_dw(resid):
        return float(np.sum(np.diff(resid) ** 2) / np.sum(resid**2))

    assert abs(
        naive_row_adjacent_dw(e_s) - naive_row_adjacent_dw(e_s[perm])
    ) > 1e-3


# ── Acceptance 3: punched-out-year gap-skip (inherited from Difference_By) ──

def test_punched_out_year_contributes_13_numerator_terms_not_14() -> None:
    groups, years, y, x = _load_fe_panel(punch_albania_2007=True)
    assert len(y) == _EXPECTED_ROWS - 1
    albania_rows = int(np.sum(groups == "Albania"))
    assert albania_rows == 15  # 16 complete years minus the punched 2007

    e = within_estimator_residuals(groups, y, x)
    diffs = difference_by(list(e), list(groups), list(years), _DELTA)

    albania_terms = sum(
        1
        for g, d in zip(groups, diffs)
        if g == "Albania" and isinstance(d, float)
    )
    albania_nas = sorted(
        t for g, t, d in zip(groups, years, diffs) if g == "Albania" and d is NA
    )
    # 15 rows: 2000 is the first period and 2008's prior year is punched out —
    # BFN must skip the gap difference, never bridge 2008 − 2006 as a fake
    # one-year change. 13 terms, NOT the 14 a previous-available-row fallback
    # would fabricate.
    assert albania_terms == 13
    assert albania_terms != 14
    assert albania_nas == [2000.0, 2008.0]

    # And the whole-panel count drops by exactly two versus the unpunched
    # panel's rows-minus-countries: Albania loses the 2007→2008 and 2006→2007
    # pairs while its denominator loses one row.
    _, term_count = bfn_workbook_mirror(e, groups, years, _DELTA)
    assert term_count == _EXPECTED_NUMERATOR_TERMS - 2


def test_zero_computable_differences_is_na_never_a_fake_zero() -> None:
    # Every group a singleton (the WHO file's ten single-observation
    # countries are this shape): the difference column is all #N/A, and an
    # unguarded IFERROR mask would present numerator 0 / denominator > 0 as
    # BFN = 0.000 — a fake strong-negative-autocorrelation reading. The
    # LAMBDA's n_terms guard makes the state a genuine #N/A instead.
    groups = np.array(["A", "B", "C"])
    years = np.array([2000.0, 2000.0, 2000.0])
    residuals = np.array([0.5, -0.2, -0.3])

    bfn, term_count = bfn_workbook_mirror(residuals, groups, years, _DELTA)
    assert term_count == 0
    assert bfn is NA


# ── Acceptance 4: the trigger matrix on all four states ─────────────────────
#
# Two fixed diagnostic cells, each self-guarding on the same two LET-bound
# counts (seq_flags = declared Sequence axes, fe_vars = Role="Fixed Effects"
# spec rows). Mirrored here and asserted against the shipped cell formulas.

_COMPUTES_DW = "<computes Durbin_Watson_By>"
_COMPUTES_BFN = "<computes BFN_Panel_Durbin_Watson>"


def dw_cell_mirror(seq_flags: int, fe_vars: int) -> str:
    if seq_flags == 0:
        return "n/a — requires Sequence"
    if seq_flags > 1:
        return "n/a — multiple Sequence flags"
    if fe_vars > 0:
        return "n/a — FE active"
    return _COMPUTES_DW


def bfn_cell_mirror(seq_flags: int, fe_vars: int) -> str:
    if seq_flags == 0:
        return "n/a — requires Sequence"
    if seq_flags > 1:
        return "n/a — multiple Sequence flags"
    if fe_vars == 0:
        return "n/a — no fixed effects"
    if fe_vars > 1:
        return "n/a — multiple FE variables"
    return _COMPUTES_BFN


def test_trigger_matrix_states() -> None:
    # no Sequence → both tokens (with or without an FE declaration)
    assert dw_cell_mirror(0, 0) == "n/a — requires Sequence"
    assert bfn_cell_mirror(0, 0) == "n/a — requires Sequence"
    assert dw_cell_mirror(0, 1) == "n/a — requires Sequence"
    assert bfn_cell_mirror(0, 1) == "n/a — requires Sequence"
    # Sequence + no FE → DW active, BFN says why it is not
    assert dw_cell_mirror(1, 0) == _COMPUTES_DW
    assert bfn_cell_mirror(1, 0) == "n/a — no fixed effects"
    # Sequence + FE → BFN active, DW says why it is not
    assert dw_cell_mirror(1, 1) == "n/a — FE active"
    assert bfn_cell_mirror(1, 1) == _COMPUTES_BFN
    # spec-error states never compute in either cell
    assert dw_cell_mirror(2, 1) == "n/a — multiple Sequence flags"
    assert bfn_cell_mirror(2, 1) == "n/a — multiple Sequence flags"
    assert bfn_cell_mirror(1, 2) == "n/a — multiple FE variables"


def _diagnostic_cell_formulas() -> tuple[str, str]:
    """The shipped DW (AB11) and BFN (AB12) formulas, via the sheet writer."""
    # pylint: disable=import-outside-toplevel
    import xlwings as xw

    from lambda_catalog.write_sheet_regression import _C_AE, _write_diagnostics
    from tests.recording_sheet import RecordingSheet

    sheet = RecordingSheet(name="Regression")
    _write_diagnostics(cast(xw.Sheet, sheet))
    dw = cast(str, sheet.cell(11, _C_AE).api.Formula2)
    bfn = cast(str, sheet.cell(12, _C_AE).api.Formula2)
    return dw, bfn


def test_shipped_cell_formulas_guard_in_the_mirrored_order() -> None:
    # The mirrors above are only evidence if the shipped formulas branch the
    # same way: guards in the same precedence order, each cell self-guarding
    # (every one of its states visible in its own formula), and compute calls
    # only in the final branch.
    dw, bfn = _diagnostic_cell_formulas()

    for formula in (dw, bfn):
        assert formula.index('seq_flags=0,"n/a — requires Sequence"') < (
            formula.index('seq_flags>1,"n/a — multiple Sequence flags"')
        )
    assert 'fe_vars>0,"n/a — FE active"' in dw
    assert dw.index("seq_flags>1,") < dw.index("fe_vars>0,")
    assert "Durbin_Watson_By(" in dw and "BFN_Panel_Durbin_Watson(" not in dw

    assert 'fe_vars=0,"n/a — no fixed effects"' in bfn
    assert 'fe_vars>1,"n/a — multiple FE variables"' in bfn
    assert bfn.index("fe_vars=0,") < bfn.index("fe_vars>1,")
    assert "BFN_Panel_Durbin_Watson(" in bfn and "Durbin_Watson_By(" not in bfn


# ── Implementation-shape assertions (by construction, not convention) ───────

def _catalog_formula(name: str) -> str:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    formulas = {fn["name"]: fn["formula_display"] for fn in payload["functions"]}
    return "".join(formulas[name].split())


def test_bfn_lambda_differences_through_difference_by_only() -> None:
    bfn = _catalog_formula("BFN_Panel_Durbin_Watson")

    # One source of truth for within-group differencing: Difference_By. No
    # re-implemented lookup, sort, or row arithmetic inside BFN itself.
    assert "Difference_By(e,g_f,t_f,step)" in bfn
    for forbidden in ("XLOOKUP", "OFFSET", "SORTBY", "DROP("):
        assert forbidden not in bfn, forbidden

    # Consumer-owns-masking: the NA→0 zeroing of first-period/gap terms is
    # local and visible in the numerator; the denominator keeps every
    # residual's square.
    assert "SUMPRODUCT(IFERROR(d,0)^2)/SUMPRODUCT(e^2)" in bfn
    # ...but the mask never turns an error STATE into a number: zero
    # computable differences (all-singleton groups, unresolved Δ) is a
    # genuine #N/A, never a fake BFN = 0.
    assert "n_terms,SUMPRODUCT(N(ISNUMBER(d)))" in bfn
    assert "IF(n_terms=0,NA()," in bfn

    # Δ routes through the visible spec cell when omitted — never a silent 1.
    assert "IF(ISOMITTED(delta),Base_Period_Delta(),delta)" in bfn
    # Residuals are computed inside the LAMBDA (scalar out, no spill exposure),
    # with group and seq filtered by the same mask so rows stay aligned.
    assert "Residuals(X,Y,filt_arg)" in bfn
    assert "FILTER(group,filt_arg)" in bfn
    assert "FILTER(seq,filt_arg)" in bfn


def test_fixed_effects_column_is_an_exact_match_role_accessor() -> None:
    accessor = _catalog_formula("Fixed_Effects_Column")
    # Exact-match XMATCH on the TAKE-trimmed role vector — the Role-axis
    # sibling of Response_Column, with the trigger cells gating on the count
    # before this is ever evaluated.
    assert 'XMATCH("FixedEffects",TAKE(Spec_Role,n_c),0)' in accessor
    assert "INDEX(Source_Data,0," in accessor


# ── The spec-case oracle's own BFN, against the value verified above ────


def _bfn_crosscheck_case():
    """The module's model, expressed as a RegressionSpecCase.

    Life expectancy ~ GDP + Schooling, FE = Country, Sequence = Year, Δ = 1
    — the same fit the two independent paths above agree on. Built here
    rather than registered in the suite: it exists to check the ORACLE, not
    to add a worksheet, and it duplicates L08's high-cardinality-FE coverage.
    """
    from lambda_catalog.analyze_regression_spec import (
        LIFE_EXPECTANCY_CSV_PATH,
        RegressionSpecCase,
        _life_spec,
        _spec_var,
        load_life_expectancy_source_rows,
    )
    from lambda_catalog.write_sheet_model_construction import (
        _ROLE_FIXED_EFFECTS,
        _ROLE_OMIT,
        _ROLE_PREDICTOR,
        _ROLE_RESPONSE,
    )

    return RegressionSpecCase(
        name="bfn_crosscheck",
        spec=tuple(
            _life_spec(
                country=_spec_var("Country", _ROLE_FIXED_EFFECTS),
                year=_spec_var("Year", _ROLE_OMIT, sequence=True),
                response=_spec_var("Life expectancy", _ROLE_RESPONSE),
                gdp=_spec_var("GDP", _ROLE_PREDICTOR, True, "Continuous"),
                schooling=_spec_var("Schooling", _ROLE_PREDICTOR, True, "Continuous"),
            )
        ),
        allow_intercept=True,
        source_csv_path=LIFE_EXPECTANCY_CSV_PATH,
        row_loader=load_life_expectancy_source_rows,
        source_table_ref="=LifeExpectancyData[#All]",
        sequence_period=1.0,
    )


def test_spec_oracle_bfn_matches_the_independently_verified_value() -> None:
    """The QC oracle's BFN == the value two other implementations agree on.

    Paths (A) and (B) above share no code with each other; this adds a
    third, `analyze_regression_sheet._bfn_panel_durbin_watson`, which is the
    one the workbook is actually compared against. Without this assertion
    the oracle could drift from the statistic it claims to compute and the
    only symptom would be a sheet failing QC for the wrong reason.
    """
    from lambda_catalog.analyze_regression_spec import calculate_regression_spec_case

    expected = calculate_regression_spec_case(_bfn_crosscheck_case())
    # The mask has to land on the same rows first, or agreement downstream
    # would be a coincidence rather than a check.
    assert expected.design.included_rows == _EXPECTED_ROWS
    assert len(set(expected.design.group_labels)) == _EXPECTED_COUNTRIES
    # Close, NOT equal. The two sides reach this statistic through different
    # fits — statsmodels LSDV with per-country dummies over there, the
    # within-estimator here — so the residuals they sum differ in the last
    # bits, and by how much depends on the BLAS the runner happens to link.
    # An earlier revision asserted exact equality; it held on one machine
    # and failed in CI at 3 ULP (…433 vs …436), which is the difference
    # between a real invariant and a coincidence of build.
    #
    # rel_tol=1e-12 is still four orders tighter than the 6-decimal
    # first-differing-digit rule the workbook comparison uses, so any drift
    # that could matter to a QC result fails here first.
    assert math.isclose(
        expected.results.summary.bfn_panel_durbin_watson,
        _EXPECTED_BFN,
        rel_tol=1e-12,
    )


def test_a_fixed_effects_case_without_a_typed_period_yields_no_bfn() -> None:
    """Δ unresolved ⇒ #N/A, not a silent 1.

    `Base_Period_Delta()` reads the TYPED Sequence Period and returns #N/A
    when none is present, so the BFN cell's delta is unresolved, every
    Difference_By lookup misses, and the LAMBDA's n_terms guard returns
    #N/A. The oracle must produce NaN there rather than quietly assuming a
    period of 1 — assuming one would compare a fabricated number against a
    cell showing an error.
    """
    from dataclasses import replace

    from lambda_catalog.analyze_regression_spec import calculate_regression_spec_case

    untyped = replace(_bfn_crosscheck_case(), sequence_period=None)
    summary = calculate_regression_spec_case(untyped).results.summary
    assert math.isnan(summary.bfn_panel_durbin_watson)
    # And plain DW stays NaN too: FE is declared, so AE11 reads
    # "n/a — FE active" regardless. Both cells being text is a legitimate
    # state, and the oracle says so rather than inventing a reading.
    assert math.isnan(summary.durbin_watson)


def test_dw_and_bfn_are_never_both_live_in_the_registry() -> None:
    """The two serial-correlation cells are mutually gated on the sheet.

    Asserted across every fittable case, because the failure this guards
    against is a spec edit quietly putting a sheet in a state where the
    oracle offers a number for a cell displaying text — which reads as a QC
    failure in the workbook rather than as the oracle bug it is.
    """
    from lambda_catalog.analyze_regression_spec import (
        build_regression_spec_cases,
        calculate_regression_spec_case,
    )
    from lambda_catalog.write_sheet_model_construction import _ROLE_FIXED_EFFECTS

    live = []
    for case in build_regression_spec_cases():
        summary = calculate_regression_spec_case(case).results.summary
        dw_live = not math.isnan(summary.durbin_watson)
        bfn_live = not math.isnan(summary.bfn_panel_durbin_watson)
        assert not (dw_live and bfn_live), (
            f"{case.plan_id} has both DW and BFN live; the sheet shows one "
            "of them as text"
        )
        # AE11's gate in full, in the formula's own order: no Sequence axis
        # -> "n/a — requires Sequence"; more than one -> "n/a — multiple
        # Sequence flags"; any Fixed Effects row -> "n/a — FE active".
        # Only all three passing leaves a number.
        #
        # An earlier revision of this test asserted `dw_live == (not has_fe)`
        # and passed — because the ORACLE had the same gap, computing DW over
        # physical row order when no Sequence axis was declared. Dropping the
        # flag from the Auto MPG cases put all nineteen into that state and
        # the live verifier caught it: a real number against a text cell.
        # State the sheet's whole condition, not the half that was noticed.
        sequence_flags = sum(1 for item in case.spec if item.sequence)
        has_fe = any(item.role == _ROLE_FIXED_EFFECTS for item in case.spec)
        assert dw_live == (sequence_flags == 1 and not has_fe), (
            f"{case.plan_id}: DW gate disagrees with the sheet "
            f"(sequence_flags={sequence_flags}, has_fe={has_fe})"
        )
        if bfn_live:
            live.append(case.plan_id)

    # At least one case must actually exercise the statistic. An all-NaN
    # column would make every BFN comparison vacuously pass.
    assert live, "no fittable case makes the BFN cell live"


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

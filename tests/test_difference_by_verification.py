"""Independent verification of the gap-aware Difference_By / Lag_By semantics.

Concept (restated as working notes, per the base-period design record):
within-group differencing is Δx(i,t) = x(i,t) − x(i, t−Δ), where Δ is the
base period. Three semantic commitments:

(A) each group's FIRST period returns NA() — never a fabricated 0, because
    a 0 would enter a design matrix silently as a real observation;
(B) "prior period" means the literal time value t−Δ — if t−Δ is absent (a
    gap), return NA(); never fall back to "the previous available row",
    which injects a spurious multi-period jump;
(C) Δ is computed-and-displayed-with-override, never silently assumed 1 —
    the same pattern as the categorical reference level.

This module computes every expected value directly from
``sample_data/Life Expectancy Data.csv`` (and from two constructed
synthetic datasets the WHO data cannot exercise) using a pure-Python
mirror of the LAMBDA semantics written here, never delegating to the
project's oracle code — then hard-asserts the exact release numbers.
It also pins the implementation SHAPE of the shipped formulas: the prior
value must come from an exact-match lookup of (group, seq−Δ) pairs, by
construction — not from OFFSET/row arithmetic, which would be
sort-order-dependent and gap-blind.

Runnable standalone (``python tests/test_difference_by_verification.py``)
or under pytest.
"""
from __future__ import annotations

import csv
import json
from collections import Counter
from datetime import date
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

# The WHO file's exact headers (note the trailing space on the response).
_COUNTRY = "Country"
_YEAR = "Year"
_LIFE_EXPECTANCY = "Life expectancy "


# ── Sentinels mirroring the worksheet value domain ──────────────────────────

class _NAType:
    """Mirror of Excel's #N/A: the value of an incomputable difference."""

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "#N/A"


NA = _NAType()
EXCLUDED = ""  # rows failing the include mask return "" (row alignment)


def _as_number(value):
    """Mirror the sheet's ISNUMBER view of a CSV field: float or ''."""
    if isinstance(value, (int, float)):
        return value
    text = (value or "").strip()
    if not text:
        return ""
    try:
        return float(text)
    except ValueError:
        return ""


# ── Pure-Python mirror of the LAMBDA semantics ──────────────────────────────

def lag_by(x, group, seq, delta, include=None):
    """Mirror of Lag_By: exact-match lookup of (group, seq − Δ) pairs.

    Included rows yield the x value at the literal time seq − Δ within the
    same group, NA when that exact period is absent (gap or first period)
    or its x is blank; excluded rows yield "". Lookup sources are the
    included rows with numeric seq and non-blank x; the first occurrence
    of a duplicate (group, seq) key wins, as XLOOKUP's does.
    """
    n = len(x)
    if include is None:
        include = [g != "" and isinstance(t, (int, float)) for g, t in zip(group, seq)]
    sources: dict[tuple, object] = {}
    for g, t, v, inc in zip(group, seq, x, include):
        if inc and isinstance(t, (int, float)) and v != "":
            sources.setdefault((g, t), v)
    out: list[object] = []
    for i in range(n):
        if not include[i]:
            out.append(EXCLUDED)
        elif not isinstance(seq[i], (int, float)):
            out.append(NA)
        else:
            out.append(sources.get((group[i], seq[i] - delta), NA))
    return out


def difference_by(x, group, seq, delta, include=None):
    """Mirror of Difference_By: x(t) − Lag_By(...)(t) with NA propagation.

    An included row whose own x, or whose t−Δ partner's x, is not numeric
    yields NA — a visible incomputable observation, never a dropped row
    and never a fabricated 0.
    """
    if include is None:
        include = [g != "" and isinstance(t, (int, float)) for g, t in zip(group, seq)]
    prior = lag_by(x, group, seq, delta, include)
    out: list[object] = []
    for v, p, inc in zip(x, prior, include):
        if not inc:
            out.append(EXCLUDED)
        elif isinstance(v, (int, float)) and isinstance(p, (int, float)):
            out.append(v - p)
        else:
            out.append(NA)
    return out


_SEP = chr(31)  # CHAR(31), the multi-identifier group-key join separator


def _group_is_blank(g) -> bool:
    """Mirror of the emptiness check: a key that is only CHAR(31) separators
    (every Identifier field blank on the row) counts as no group, so
    unidentified rows never form a spurious group."""
    return isinstance(g, str) and g.replace(_SEP, "") == ""


def sequence_deltas(group, seq):
    """Mirror of Sequence_Deltas: seam-safe within-group consecutive spacings.

    Sort by (group, seq); keep diffs only where the sorted neighbors share
    a group (seams dropped by construction) and the diff is positive
    (duplicates dropped). Pooled column-wise diffs over raw row order are
    exactly what this exists to forbid — on Country/Year panel data they
    manufacture large negative seam deltas at every country boundary. Rows
    whose composite group key is delimiter-only (all identifiers blank) are
    excluded, matching the SUBSTITUTE guard in the formula.
    """
    rows = sorted(
        (g, t)
        for g, t in zip(group, seq)
        if not _group_is_blank(g) and isinstance(t, (int, float))
    )
    deltas = []
    for (g_prev, t_prev), (g_cur, t_cur) in zip(rows, rows[1:]):
        if g_cur == g_prev and t_cur - t_prev > 0:
            deltas.append(t_cur - t_prev)
    return deltas


def delta_spectrum(deltas):
    """Mirror of Sequence_Delta_Spectrum: {distinct delta: count}."""
    return dict(Counter(deltas))


def base_period_delta_candidate(deltas):
    """Mirror of Base_Period_Delta_Candidate: MODE, MIN fallback, NA.

    Returns (candidate, mode_is_defined). MODE.SNGL is defined only when
    some spacing repeats; with all spacings distinct the candidate falls
    back to MIN and the sheet surfaces the "no natural base period"
    override prompt.
    """
    if not deltas:
        return NA, False
    counts = Counter(deltas)
    most_common, top_count = counts.most_common(1)[0]
    if top_count < 2:
        return min(deltas), False
    return most_common, True


def regularity_flag_fires(deltas, delta_in_use):
    """Mirror of the Regularity verdict: any spacing differs from Δ."""
    return bool(deltas) and any(d != delta_in_use for d in deltas)


def off_grid_flag_fires(deltas, delta_in_use):
    """Mirror of the Off-grid verdict: a spacing not a whole multiple of Δ."""
    if not deltas or not isinstance(delta_in_use, (int, float)) or delta_in_use <= 0:
        return False
    return any(d % delta_in_use != 0 for d in deltas)


def calendar_signature_fires(deltas):
    """Mirror of the calendar-signature verdict: a MAJORITY of spacings sit
    in the calendar bands ~28–31 (monthly), ~90–92 (quarterly), ~365–366
    (yearly). Surfacing is the fix — the tool never quantizes calendar
    spacings to a scalar Δ."""
    if not deltas:
        return False
    in_band = sum(
        1 for d in deltas if 28 <= d <= 31 or 90 <= d <= 92 or 365 <= d <= 366
    )
    return in_band * 2 > len(deltas)


def _excel_serial(day: date) -> int:
    """Excel 1900-system serial number (days since 1899-12-30)."""
    return (day - date(1899, 12, 30)).days


# ── Fixture loading ─────────────────────────────────────────────────────────

def _load_who_rows() -> list[dict]:
    with CSV_PATH.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def _who_columns(rows: list[dict]):
    group = [r[_COUNTRY] for r in rows]
    seq = [_as_number(r[_YEAR]) for r in rows]
    x = [_as_number(r[_LIFE_EXPECTANCY]) for r in rows]
    return x, group, seq


# ── WHO expected values (Sequence = Year, groups = Country) ─────────────────

def test_who_delta_candidate_spectrum_and_quiet_verdicts() -> None:
    rows = _load_who_rows()
    _, group, seq = _who_columns(rows)

    deltas = sequence_deltas(group, seq)
    assert delta_spectrum(deltas) == {1: 2745}

    candidate, mode_defined = base_period_delta_candidate(deltas)
    assert candidate == 1
    assert mode_defined is True  # no "no natural base period" prompt

    assert regularity_flag_fires(deltas, candidate) is False
    assert off_grid_flag_fires(deltas, candidate) is False
    assert calendar_signature_fires(deltas) is False


def test_who_difference_by_na_and_value_counts() -> None:
    rows = _load_who_rows()
    x, group, seq = _who_columns(rows)
    assert len(rows) == 2938

    diff = difference_by(x, group, seq, 1)

    na_count = sum(1 for v in diff if v is NA)
    value_count = sum(1 for v in diff if isinstance(v, float))
    excluded_count = sum(1 for v in diff if v == EXCLUDED)

    # Exactly one NA per country: 183 sixteen-year panels contribute their
    # first period (A), and all 10 single-observation countries contribute
    # their only row. Every row has a Country and a numeric Year, so
    # nothing is mask-excluded — 193 + 2745 = 2938.
    assert na_count == 193
    assert value_count == 2745
    assert excluded_count == 0
    assert na_count + value_count == 2938

    per_country_na = Counter(
        g for g, v in zip(group, diff) if v is NA
    )
    assert len(per_country_na) == 193
    assert set(per_country_na.values()) == {1}


def test_who_spot_values_afghanistan_and_cook_islands() -> None:
    rows = _load_who_rows()
    x, group, seq = _who_columns(rows)
    diff = difference_by(x, group, seq, 1)

    by_key = {(g, t): v for g, t, v in zip(group, seq, diff)}
    life = {(g, t): v for g, t, v in zip(group, seq, x)}

    # Afghanistan Δ(2015) = LifeExp(2015) − LifeExp(2014) — verified from
    # the CSV (65.0 and 59.9), then hard-asserted.
    assert life[("Afghanistan", 2015)] == 65.0
    assert life[("Afghanistan", 2014)] == 59.9
    assert by_key[("Afghanistan", 2015)] == 65.0 - 59.9

    # Cook Islands is one of the 10 single-observation countries (T = 1):
    # its single row is a first period — NA, never a fabricated 0 (A), and
    # never "" (its Country and Year are present, so the row is included).
    cook_rows = [v for g, v in zip(group, diff) if g == "Cook Islands"]
    assert len(cook_rows) == 1
    assert cook_rows[0] is NA


# ── Synthetic case 1: punched-out year (gap on the Δ grid) ──────────────────

def _albania_punched():
    """WHO Albania with Year 2007 deleted: a genuine within-panel gap."""
    rows = [
        r for r in _load_who_rows()
        if r[_COUNTRY] == "Albania" and _as_number(r[_YEAR]) != 2007
    ]
    group = [r[_COUNTRY] for r in rows]
    seq = [_as_number(r[_YEAR]) for r in rows]
    x = [_as_number(r[_LIFE_EXPECTANCY]) for r in rows]
    return x, group, seq


def test_punched_out_year_gap_returns_na_not_a_two_year_jump() -> None:
    x, group, seq = _albania_punched()
    assert len(x) == 15
    # Precondition for the expected counts: Albania's response is complete.
    assert all(isinstance(v, float) for v in x)

    diff = difference_by(x, group, seq, 1)
    by_year = dict(zip(seq, diff))

    # (A) first period NA; (B) 2008 NA because 2007 is absent — the 2008
    # value must NOT become x(2008) − x(2006), the previous-available-row
    # fallback this design forbids.
    assert by_year[2000] is NA
    assert by_year[2008] is NA
    na_years = sorted(t for t, v in by_year.items() if v is NA)
    assert na_years == [2000, 2008]
    assert sum(1 for v in diff if isinstance(v, float)) == 13


def test_punched_out_year_spectrum_and_verdicts() -> None:
    _, group, seq = _albania_punched()
    deltas = sequence_deltas(group, seq)

    assert delta_spectrum(deltas) == {1: 13, 2: 1}

    candidate, mode_defined = base_period_delta_candidate(deltas)
    assert candidate == 1
    assert mode_defined is True

    # Regularity fires (one spacing ≠ 1); Off-grid stays quiet (2 IS a
    # whole multiple of 1 — the gap sits on the Δ grid); no calendar
    # signature.
    assert regularity_flag_fires(deltas, candidate) is True
    assert off_grid_flag_fires(deltas, candidate) is False
    assert calendar_signature_fires(deltas) is False


# ── Synthetic case 2: calendar dates (spacing that must NOT be quantized) ───

def _calendar_series():
    """First-of-month Excel serial dates, Jan 2020 – Dec 2021 (24 dates;
    2020 is a leap year). x is an arbitrary numeric series."""
    dates = [
        _excel_serial(date(year, month, 1))
        for year in (2020, 2021)
        for month in range(1, 13)
    ]
    x = [float(100 + i) for i in range(len(dates))]
    group = ["S1"] * len(dates)
    return x, group, dates


def test_calendar_dates_spectrum_mode_and_all_three_verdicts() -> None:
    _, group, seq = _calendar_series()
    assert len(seq) == 24

    deltas = sequence_deltas(group, seq)
    # 23 spacings = days in each preceding month: thirteen 31s, eight 30s,
    # one 29 (Feb 2020, leap) and one 28 (Feb 2021).
    assert len(deltas) == 23
    assert delta_spectrum(deltas) == {31: 13, 30: 8, 29: 1, 28: 1}

    candidate, mode_defined = base_period_delta_candidate(deltas)
    assert candidate == 31  # MODE — and a WRONG Δ for 11 months of the year
    assert mode_defined is True

    # Regularity fires broadly; Off-grid fires (28/29/30 are not whole
    # multiples of 31); the calendar signature fires (23 of 23 spacings in
    # the ~28–31 band) and recommends an integer period index upstream
    # (YEAR, or YEAR*12+MONTH) — the tool must surface this, never
    # quantize day counts to a scalar Δ.
    assert regularity_flag_fires(deltas, candidate) is True
    assert off_grid_flag_fires(deltas, candidate) is True
    assert calendar_signature_fires(deltas) is True


def test_calendar_dates_difference_by_only_matches_exact_31_day_lags() -> None:
    x, group, seq = _calendar_series()
    diff = difference_by(x, group, seq, 31)

    # Only months preceded by a 31-day month sit exactly 31 days after an
    # observed date: the 13 spacings of 31 in the spectrum. Every other
    # row — including the first — is NA. This is commitment (B) doing its
    # job on off-grid data: no silent nearest-row fallback.
    assert sum(1 for v in diff if isinstance(v, float)) == 13
    assert sum(1 for v in diff if v is NA) == 11


# ── Implementation-shape assertions (by construction, not convention) ───────

def _catalog_formulas() -> dict[str, str]:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    return {fn["name"]: fn["formula_display"] for fn in payload["functions"]}


def _compact(formula: str) -> str:
    return "".join(formula.split())


def test_lag_by_is_an_exact_match_lookup_of_group_seq_minus_delta_pairs() -> None:
    formulas = _catalog_formulas()
    lag = _compact(formulas["Lag_By"])

    # Exact-match XLOOKUP of composite (group, seq − Δ) keys...
    assert "XLOOKUP(g_v&sep&(t_v-step),source_keys,x_v,NA(),0)" in lag
    # ...and never OFFSET/row arithmetic, which is sort-order-dependent
    # and gap-blind.
    assert "OFFSET" not in lag
    assert "ROW(" not in lag


def test_difference_by_delegates_the_pair_lookup_and_signals_by_na() -> None:
    formulas = _catalog_formulas()
    diff = _compact(formulas["Difference_By"])

    # One source of truth for the (group, seq − Δ) lookup: Lag_By.
    assert "Lag_By(x,group,seq,step,inc)" in diff
    assert "XLOOKUP" not in diff and "OFFSET" not in diff
    # (A): the incomputable cases return a real error, never a 0.
    assert "NA()" in diff
    assert 'IF(ISNUMBER(x_v)*ISNUMBER(prior),x_v-prior,NA())' in diff


def test_omitted_delta_routes_through_the_visible_spec_cell() -> None:
    # (C): Δ defaults to the sheet's Base Period Δ cell (spec column I on
    # the Sequence-flagged row) via Base_Period_Delta() — never a silent 1.
    formulas = _catalog_formulas()
    for name in ("Lag_By", "Difference_By"):
        compact = _compact(formulas[name])
        assert "IF(ISOMITTED(delta),Base_Period_Delta(),delta)" in compact, name
        assert ",1," not in compact.split("ISOMITTED(delta)")[1][:40], name
    # The accessor's spec references are UNQUALIFIED, and that is the point:
    # a sheet-scoped name resolves against the sheet the calling formula lives
    # on, so each Regression-shaped sheet reads its own Δ. Sheet-qualifying
    # them (the old 'Regression'!Spec_Sequence form) made every sheet in a
    # multi-sheet workbook read one shared spec block, and made the function
    # unwritable — hence skipped, hence #NAME? — in a workbook with no sheet
    # by that name.
    accessor = _compact(formulas["Base_Period_Delta"])
    assert "XMATCH(TRUE,TAKE(Spec_Sequence" in accessor
    assert "TAKE(Spec_Sequence_Period" in accessor
    assert "IF(ISNUMBER(value),value,NA())" in accessor
    assert "'Regression'!" not in accessor


def test_spectrum_formula_is_seam_safe_not_a_pooled_columnwise_diff() -> None:
    formulas = _catalog_formulas()
    gaps = _compact(formulas["Sequence_Deltas"])

    # Sorted by (group, seq), diffs kept only where sorted neighbors share
    # a group: seams cannot contaminate the spacings.
    assert "SORTBY(g_f,g_f,1,t_f,1)" in gaps
    assert "SORTBY(t_f,g_f,1,t_f,1)" in gaps
    assert "same,DROP(g_s,1)=DROP(g_s,-1)" in gaps
    assert "FILTER(d,same*(d>0),NA())" in gaps
    # The emptiness check strips the CHAR(31) join separators, so a row whose
    # identifier fields are ALL blank (a delimiter-only composite key) is
    # excluded rather than forming a spurious group.
    assert 'ISNUMBER(t_v)*(SUBSTITUTE(g_v,CHAR(31),"")<>"")' in gaps

    candidate = _compact(formulas["Base_Period_Delta_Candidate"])
    assert "IFERROR(MODE.SNGL(d),MIN(d))" in candidate
    assert "IF(COUNT(d)=0,NA()" in candidate


def test_all_blank_identifier_rows_do_not_form_a_spurious_group() -> None:
    # Multi-identifier panel: two rows fully identified as ("A","X") and two
    # rows with BOTH identifiers blank, which TEXTJOIN(...,FALSE,...) joins
    # into a delimiter-only key (CHAR(31)). Without the SUBSTITUTE guard the
    # delimiter-only key reads as non-empty and the two unidentified rows
    # would contribute a spurious spacing; with it they are excluded, so
    # only the ("A","X") group's single spacing survives.
    group = [f"A{_SEP}X", f"A{_SEP}X", _SEP, _SEP]
    seq = [2000, 2001, 2005, 2006]
    deltas = sequence_deltas(group, seq)
    assert delta_spectrum(deltas) == {1: 1}

    # A partially-identified row keeps its identity (only the separator is
    # stripped for the emptiness test, never the real key), so it still
    # groups distinctly from a fully-identified one.
    group_partial = [f"A{_SEP}", f"A{_SEP}", f"A{_SEP}X"]
    seq_partial = [2000, 2001, 2002]
    assert delta_spectrum(sequence_deltas(group_partial, seq_partial)) == {1: 1}


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

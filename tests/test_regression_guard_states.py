"""Tests for the spec block's guard-rail oracles.

These assert what a guard configuration is SUPPOSED to make the sheet show —
status text, the per-row Design Columns audit, the model formula, and which
conditional-formatting rules fire. Everything here is pure Python; the live
Excel comparison is tools/inspect_test_model_sheets.py, which needs Excel and
so does not run in CI.
"""
# pylint: disable=missing-function-docstring
from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from lambda_catalog.analyze_regression_guard_states import (
    EMPTY_MODEL,
    SEVERITY_AMBER,
    SEVERITY_RED,
    base_period_delta_candidate,
    build_guard_state_cases,
    calculate_guard_state_case,
    delta_spectrum,
    sequence_verdict,
)
from lambda_catalog.write_spec_block import (
    _MSG_CALENDAR,
    _MSG_NO_NATURAL,
    _MSG_OFF_GRID,
    _MSG_REGULARITY,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"

# Ordered and pinned, exactly like _EXPECTED_CASE_NAMES for the fittable
# cases: adding, renaming, reordering or dropping a guard case must be a
# deliberate edit here in the same commit.
_EXPECTED_GUARD_NAMES = [
    "guard_no_response",
    "guard_empty_model",
    "guard_two_responses",
    "guard_two_sequence_flags",
    "guard_two_fixed_effects",
    "guard_fixed_effects_with_intercept",
    "guard_intercept_off_with_categorical",
    "guard_log_on_categorical",
    "guard_interaction_bad_operand",
    "guard_reciprocal_product",
    "guard_excluded_operand",
    "guard_unknown_interaction_operation",
    "guard_ln_zero_propagation",
    "guard_width_guard_warning",
    "guard_spec_block_retarget_widens",
    "guard_sequence_period_override",
    "guard_irregular_panel_spacing",
]


def _expected(name: str):
    case = next(c for c in build_guard_state_cases() if c.name == name)
    return calculate_guard_state_case(case)


def _flag_rules(expected) -> set[tuple[str, str, str]]:
    return {(f.column, f.severity, f.rule) for f in expected.flags}


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_guard_case_names_are_pinned() -> None:
    assert [case.name for case in build_guard_state_cases()] == _EXPECTED_GUARD_NAMES


# ── The sequence layer's pure functions ──────────────────────────────────


def test_base_period_delta_candidate_prefers_the_mode_then_the_minimum() -> None:
    assert base_period_delta_candidate([1, 1, 2, 3]) == 1
    # Nothing repeats: MODE.SNGL errors and the catalog's IFERROR takes MIN.
    assert base_period_delta_candidate([5, 7, 9]) == 5
    # No spacings at all is the #N/A branch.
    assert base_period_delta_candidate([]) is None


def test_delta_spectrum_is_sorted_unique_values_with_counts() -> None:
    assert delta_spectrum([2, 1, 1, 3, 1]) == ((1, 3), (2, 1), (3, 1))
    assert delta_spectrum([]) == ()


def test_sequence_verdict_priority_is_most_severe_first() -> None:
    # Off-grid outranks regularity: a non-multiple gap is BOTH, and the
    # sheet's nested IF reports only the first match.
    assert sequence_verdict([2, 2, 3], 2.0) == _MSG_OFF_GRID
    # Whole multiples but not all equal to Δ — regularity only.
    assert sequence_verdict([1, 1, 2], 1.0) == _MSG_REGULARITY
    # Every spacing equals Δ: quiet.
    assert sequence_verdict([3, 3, 3], 3.0) == ""
    # A day-count signature. Note how narrow the reachable window is: the
    # three branches above all outrank it, so the calendar message appears
    # only when the spacings are PERFECTLY uniform, that uniform value is
    # itself calendar-like, and Δ equals it. A mixed 30/31 monthly series —
    # the obvious real-world case — reports regularity instead.
    assert sequence_verdict([30, 30, 30], 30.0) == _MSG_CALENDAR
    assert sequence_verdict([30, 31, 30, 1], 1.0) == _MSG_REGULARITY


def test_sequence_verdict_is_blank_without_an_axis_or_a_period() -> None:
    assert sequence_verdict([], 1.0) == ""
    assert sequence_verdict([1, 1], None) == ""


def test_no_natural_period_fires_only_when_nothing_repeats() -> None:
    """MODE.SNGL errors on an all-distinct spacing set. Δ must still divide
    every gap and equal them all, or an earlier branch wins — so the only way
    to reach this message is a single spacing."""
    assert sequence_verdict([4.0], 4.0) == _MSG_NO_NATURAL


# ── The guard configurations ─────────────────────────────────────────────


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_no_response_degrades_the_readout_and_widens_the_sample() -> None:
    """Dropping the Response removes its completeness term from the mask, so
    the sample GROWS — the clearest demonstration that the mask is derived
    from the spec rather than fixed per dataset."""
    expected = _expected("guard_no_response")

    assert expected.responses_count == 0
    assert expected.response_name == "(none)"
    assert expected.model_formula.startswith("(none) ~ 1 + ")
    # 392 rows survive the shipped T0 spec; without the MPG completeness
    # term the 8 MPG-missing rows rejoin.
    assert expected.included_rows == 400
    assert expected.flags == ()


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_empty_model_is_the_state_that_produces_the_documented_string() -> None:
    expected = _expected("guard_empty_model")

    assert expected.audit_k == EMPTY_MODEL
    # The intercept is the only design column left.
    assert expected.design_columns_total == 1
    # The readout writes the intercept marker as a PREFIX, so an empty
    # predictor list genuinely renders a trailing separator.
    assert expected.model_formula == "MPG ~ 1 + "


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_two_responses_is_counted_but_still_resolves_the_first() -> None:
    expected = _expected("guard_two_responses")

    assert expected.responses_count == 2
    # Response_Column's XMATCH takes the first match, so MPG still wins even
    # though Weight also claims the role.
    assert expected.response_name == "MPG"


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_two_sequence_flags_errors_and_reddens_every_flagged_cell() -> None:
    expected = _expected("guard_two_sequence_flags")

    assert expected.sequence_flag_count == 2
    assert expected.sequence_status.startswith("ERROR: multiple Sequence rows")
    flagged_rows = sorted(f.row for f in expected.flags if f.column == "H")
    assert len(flagged_rows) == 2
    assert all(f.severity == SEVERITY_RED for f in expected.flags)


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_two_sequence_flags_still_computes_spacings_from_the_first_row() -> None:
    """The spacing layer must NOT go quiet in the error state: the sheet's
    XMATCH resolves the first flagged row regardless, and an oracle that
    blanked here would fail a sheet that is behaving correctly."""
    expected = _expected("guard_two_sequence_flags")

    assert expected.period_in_use is not None
    assert expected.delta_spectrum


def test_only_the_two_mechanics_cases_flag_sequence_on_auto_mpg() -> None:
    """G03 and M16 are the whole list, and both declare the flag explicitly.

    Auto MPG has no ordering axis (see the companion assertion in
    tests/test_regression_spec_qc.py), so nothing inherits a Sequence flag
    there any more. Two guard cases still wire one up, and neither is making
    a claim about the data: G03 tests the H2 cardinality rule, which counts
    flags, and M16 tests the typed-period override, which reads the flagged
    row positionally. Both are dataset-independent mechanisms that need a
    flag present to be reachable at all.

    Pinning the list is what stops the flag being reintroduced elsewhere on
    this dataset by a copy-paste.
    """
    flagged = {
        case.name: tuple(item.name for item in case.spec if item.sequence)
        for case in build_guard_state_cases()
        if case.source_table_ref == "=MileageData[#All]"
    }
    assert {name: names for name, names in flagged.items() if names} == {
        "guard_two_sequence_flags": ("Acceleration", "Model Year"),
        "guard_sequence_period_override": ("Model Year",),
    }


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_two_fixed_effects_rows_raise_the_cardinality_error() -> None:
    expected = _expected("guard_two_fixed_effects")

    assert expected.fixed_effects_count == 2
    # The Role column owns one status cell (B2), which picks between the
    # Response-cardinality and Fixed-Effects-cardinality errors by severity.
    # This spec has a legal Response, so the FE error is what surfaces.
    assert expected.role_status.startswith("ERROR: multiple Fixed Effects rows")


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_fixed_effects_with_intercept_flags_the_toggle() -> None:
    expected = _expected("guard_fixed_effects_with_intercept")

    assert ("C", SEVERITY_RED, "intercept_with_fixed_effects") in _flag_rules(expected)
    assert expected.role_status == ""


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_intercept_off_with_a_categorical_flags_the_toggle() -> None:
    expected = _expected("guard_intercept_off_with_categorical")

    assert _flag_rules(expected) == {
        ("C", SEVERITY_RED, "intercept_off_with_categorical")
    }
    # Advisory only — the model still builds, and the "0 + " marker is what
    # distinguishes the rendered formula from the intercept-on case.
    assert expected.model_formula.startswith("MPG ~ 0 + ")


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_log_on_a_categorical_is_flagged_but_never_applied() -> None:
    expected = _expected("guard_log_on_categorical")

    assert ("G", SEVERITY_RED, "log_on_categorical") in _flag_rules(expected)
    # The dummies still encode normally: no "Ln(" appears in the formula.
    assert "Ln(" not in expected.model_formula


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_interaction_naming_a_non_predictor_is_red_and_builds_no_columns() -> None:
    expected = _expected("guard_interaction_bad_operand")

    assert ("M", SEVERITY_RED, "operand_not_a_predictor") in _flag_rules(expected)
    # Car Name is the Identifier, so mate() returns 0 and Weight contributes
    # its main effect only — no interaction column in the formula.
    assert "×" not in expected.model_formula


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_reciprocal_product_reddens_both_rows_operation_cells() -> None:
    expected = _expected("guard_reciprocal_product")

    reciprocal = [f for f in expected.flags if f.rule == "reciprocal_declaration"]
    assert len(reciprocal) == 2
    assert {f.column for f in reciprocal} == {"N"}
    assert {f.severity for f in reciprocal} == {SEVERITY_RED}


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_excluded_operand_is_amber_not_red_and_still_builds() -> None:
    """A marginality violation is a modelling choice, so the library flags it
    and allows it. Turning this red would be the library deciding a
    statistical question on the user's behalf."""
    expected = _expected("guard_excluded_operand")

    assert _flag_rules(expected) == {("M", SEVERITY_AMBER, "operand_excluded")}
    assert "Weight × Acceleration" in expected.model_formula


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_unknown_interaction_operation_renders_the_refusal_marker() -> None:
    """No CF rule fires — the refusal is visible in the header itself, which
    is the constructor saying it will not guess."""
    expected = _expected("guard_unknown_interaction_operation")

    assert "Weight ? Displacement" in expected.model_formula
    assert not any(f.column in ("M", "N") for f in expected.flags)


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_ln_zero_guard_does_not_narrow_the_sample() -> None:
    """Strict ``Log`` does not filter, and now says so.

    The mask has no Log-positivity term for this token, so Schooling's 28 zero
    rows stay in and the #N/A propagates instead of the rows dropping out —
    unchanged, and it is the half of the pair that must stay this way. What is
    new is that the sheet no longer leaves the user to discover it from a wall
    of #N/A: the Transform cell goes red and G2 names the column, the count and
    the token that would exclude them. L12 is the other half. See
    analyze_regression_guard_states._LN_ZERO_GUARD_NOTE."""
    from lambda_catalog.analyze_life_expectancy import (
        load_life_expectancy_source_rows,
    )

    expected = _expected("guard_ln_zero_propagation")
    rows = load_life_expectancy_source_rows()

    def _numeric(value: object) -> bool:
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    unmasked = sum(
        1
        for row in rows
        if _numeric(row["Life expectancy"])
        and _numeric(row["Adult Mortality"])
        and _numeric(row["Schooling"])
    )
    zero_schooling = sum(
        1
        for row in rows
        if _numeric(row["Schooling"]) and row["Schooling"] == 0
    )

    assert zero_schooling == 28
    # The zero rows are INSIDE the sample, not excluded from it.
    assert expected.included_rows == unmasked
    assert "Ln(Schooling)" in expected.model_formula

    # Red on the Transform cell, pointing at the row the user has to change.
    assert ("G", SEVERITY_RED, "log_nonpositive_rows") in _flag_rules(expected)
    # G2 names the variable, the count of offending rows IN THE SAMPLE (26 of
    # the 28 zeros — two are already out on the response or Adult Mortality),
    # and the token that would exclude them.
    assert expected.log_domain_status == (
        "ERROR: Schooling has 26 values ≤ 0 under Log — "
        "the fit is #N/A. Use Log (drop ≤ 0)."
    )


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_typed_sequence_period_overrides_the_candidate_and_moves_the_verdict() -> None:
    expected = _expected("guard_sequence_period_override")
    # The baseline is M16's own case with the typed override removed, built
    # here rather than borrowed from another registered case. No other Auto
    # MPG case declares a Sequence flag at all — the dataset is
    # cross-sectional, so only the case that tests the override machinery
    # wires one up — and a baseline on a different dataset would change the
    # candidate as well as the override, which is the one thing this test
    # needs held fixed.
    override_case = next(
        c for c in build_guard_state_cases()
        if c.name == "guard_sequence_period_override"
    )
    baseline = calculate_guard_state_case(
        replace(override_case, sequence_period_override={})
    )

    # Same dataset and axis, so the candidate is the same; only the typed
    # override differs.
    assert baseline.period_in_use == 1.0
    assert expected.period_in_use == 2.0
    # Odd gaps are not whole multiples of 2, so the verdict escalates from
    # the candidate's "not evenly spaced" to "off-grid".
    assert expected.verdict == _MSG_OFF_GRID
    assert baseline.verdict == _MSG_REGULARITY


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_irregular_panel_spacing_fires_the_regularity_verdict() -> None:
    """Requires the panel unit to be the IDENTIFIER: Sequence_Deltas groups by
    Identifier columns, so P01/P02's unique Lot_ID leaves every group a
    singleton and the verdict unconditionally blank."""
    expected = _expected("guard_irregular_panel_spacing")

    assert expected.period_in_use == 1.0
    assert expected.verdict == _MSG_REGULARITY
    # More than one distinct spacing is exactly what "irregular" means.
    assert len(expected.delta_spectrum) > 1


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_design_columns_audit_is_blank_on_non_predictor_rows() -> None:
    """Column O mirrors the constructor's own predicate: "" when the row is
    not a Predictor candidate at all, 0 when it is a candidate that is
    currently out, and a count otherwise. Conflating the first two would
    make an excluded predictor indistinguishable from an Identifier."""
    expected = _expected("guard_empty_model")

    assert expected.design_columns[0] == ""      # MPG — the Response
    assert expected.design_columns[1] == 0       # Cylinders — excluded candidate
    assert all(
        value == "" or isinstance(value, int) for value in expected.design_columns
    )


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_only_l07_trips_the_width_guard() -> None:
    """Exactly one case reaches the 200-column threshold, and it is the one
    that exists for it. Any other case tripping the guard would mean a spec
    grew a design matrix nobody intended."""
    tripping = {
        case.plan_id
        for case in build_guard_state_cases()
        if calculate_guard_state_case(case).width_guard_status
    }
    assert tripping == {"L07"}


@pytest.mark.skipif(not CSV_PATH.exists(), reason="Auto MPG CSV not found")
def test_width_guard_case_crosses_the_threshold_and_degrades_visibly() -> None:
    """L07. The hard guard errors at 200 design columns, and this is the only
    case that reaches it.

    It is a GUARD state rather than a fittable model because of what the
    first live Excel run showed: at k = 205 the workbook cannot invert the
    Gram matrix and every engine cell reads nan. That is the condition the
    guard exists to error about, so the case asserts the error and the
    visible degradation instead of numbers the sheet cannot produce.

    The k assertion is exact on purpose. The plan's own arithmetic (193
    countries -> 192 dummies + 8 predictors = 200) does not survive contact
    with the data — the response's blanks cap the sample at 183 countries —
    so the case adds C(Year) to clear the threshold with margin. If a future
    data or spec change drops k back under 200, the guard silently stops
    being exercised at all, which is what this pins."""
    from lambda_catalog.write_sheet_regression import _DESIGN_MATRIX_MAX_COLUMNS

    expected = _expected("guard_width_guard_warning")

    assert expected.audit_k == 205
    assert expected.audit_k > _DESIGN_MATRIX_MAX_COLUMNS
    assert expected.width_guard_status == "ERROR"
    # Country vacated the Identifier role, so labels fall back to positional.
    assert expected.included_rows == 2909


def test_retarget_case_puts_its_evidence_past_the_narrow_shells_last_row() -> None:
    """L10. The spec block is built for Auto MPG (12 columns) and retargeted
    to Life Expectancy (23), which is what a user does by editing
    Source_Table in the Name Manager.

    The case is only worth its sheet if its evidence lies BEYOND the row the
    narrow shell used to end at. Before the block was made table-free, the
    Spec_* bands were structured references into a ListObject sized at build
    time; a band 12 rows tall under a 23-column table meant TAKE returned 12
    rows (it does not pad) and INDEX(rl, 23) ran off the end, so the whole
    engine read #REF!.

    So this pins the two things that make the case a detector rather than
    just another Life Expectancy fit: the audit is 23 long, and a predictor
    that actually contributes design columns sits past the old bottom edge.
    A future edit that moves the model's predictors up into the first 12
    rows would leave the case passing while testing nothing.
    """
    from lambda_catalog.write_spec_block import (
        _FIRST_DATA_ROW,
        _LAST_DATA_ROW,
        _N_VARIABLES,
    )

    case = next(
        c for c in build_guard_state_cases()
        if c.name == "guard_spec_block_retarget_widens"
    )
    expected = _expected("guard_spec_block_retarget_widens")

    # The shell is built NARROWER than the table it reads. Every other case
    # leaves this None, which is why none of them exercise the retarget.
    assert case.shell_profile_key == "auto_mpg"
    assert len(case.spec) == 23
    assert len(expected.design_columns) == len(case.spec)
    assert len(case.spec) > _N_VARIABLES

    contributing = [
        index
        for index, count in enumerate(expected.design_columns)
        if count != ""
    ]
    assert contributing, expected.design_columns
    # At least one contributing row lands past the narrow shell's last row.
    deepest_row = _FIRST_DATA_ROW + max(contributing)
    assert deepest_row > _LAST_DATA_ROW, (deepest_row, _LAST_DATA_ROW)

    # And the fit itself has to be right, not merely present.
    assert expected.response_name == "Life expectancy"
    assert expected.audit_k == 2
    assert expected.design_columns_total == 3  # + the intercept
    assert expected.width_guard_status == ""

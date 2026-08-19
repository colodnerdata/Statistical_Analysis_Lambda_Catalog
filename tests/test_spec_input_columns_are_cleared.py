"""``apply_spec_case`` must clear every writable spec column.

The spec block's A-O columns split in two: the cells a case DECLARES (the B-I
input band plus the appended M/N interaction pair) and the cells the sheet
DERIVES (the A label and the J/K/L/O computed displays). ``apply_spec_case``
clears the first set before rewriting it, because a case must never be
evaluated against whatever the previous write left behind -- the same
invariant the function already enforces unconditionally for ``Source_Table``
and ``$AK$12``.

The clear has to cover every input the spec can declare, including the ones a
LATER step writes. Sequence Period (column I) is the example: it is written by
``apply_sequence_period_overrides`` and only for cases that declare a period,
so a caller that reuses one sheet across cases can carry a typed period into a
case that declared none, and ``Period In Use`` prefers a typed period over its
computed candidate -- a stale value reads as a plausible number rather than an
error.

These tests pin the PARTITION rather than the tuple's contents, so a newly
appended input column that nobody adds to the clear list fails here instead of
on the next Excel run.
"""
from __future__ import annotations

from lambda_catalog.regression_spec_sheet_io import _SPEC_INPUT_COLUMNS
from lambda_catalog.spec_layout import (
    _C_DESIGN_COLUMNS,
    _C_INCLUDE,
    _C_INTERACTION_OPERATION,
    _C_INTERACTION_TERM,
    _C_LABEL,
    _C_LEVELS,
    _C_ORDER,
    _C_PERIOD_IN_USE,
    _C_REF_IN_USE,
    _C_REFERENCE,
    _C_ROLE,
    _C_SEQUENCE,
    _C_SEQUENCE_PERIOD,
    _C_SPEC_LAST,
    _C_TRANSFORM,
    _C_TYPE,
)

# The A–O columns a case never writes: the row label and the four computed
# displays that "derive, never feed". Everything else in the block is an input.
_DERIVED_COLUMNS = (
    _C_LABEL,
    _C_PERIOD_IN_USE,
    _C_LEVELS,
    _C_REF_IN_USE,
    _C_DESIGN_COLUMNS,
)


def test_sequence_period_is_cleared_between_cases() -> None:
    """Column I is in the clear set — the regression this module exists for.

    Without it the verify inspector carries P01/P02's typed Δ=1 into every
    later case whose Sequence-flagged row shares a row offset, and the symptom
    is a plausible number rather than an error.
    """
    assert _C_SEQUENCE_PERIOD in _SPEC_INPUT_COLUMNS, (
        "apply_spec_case must clear the Sequence Period (column I) or a typed "
        "period leaks into the next case on any caller that reuses one sheet"
    )


def test_every_writable_spec_column_is_cleared() -> None:
    """The clear set is exactly 'the block minus label minus computed displays'.

    Stated as a partition so appending a new INPUT column to the spec block
    fails here unless it is also added to the clear list. ``_C_ORDER`` (F) is
    the one deliberate exclusion: it is reserved-and-unwired, so no case writes
    it and clearing it would touch a column the spec does not yet own.
    """
    all_columns = set(range(_C_LABEL, _C_SPEC_LAST + 1))
    expected_inputs = all_columns - set(_DERIVED_COLUMNS) - {_C_ORDER}

    assert set(_SPEC_INPUT_COLUMNS) == expected_inputs, (
        "every writable spec column must be cleared before the case is "
        "rewritten; unexpected/missing: "
        f"{set(_SPEC_INPUT_COLUMNS) ^ expected_inputs}"
    )


def test_no_computed_display_is_ever_cleared() -> None:
    """J/K/L/O hold spills and formulas — clearing one would blank the display.

    The computed columns are single ``MAP`` spills anchored at the first data
    row (CLAUDE.md: 'the four computed columns ... are each ONE spill'), so a
    per-row ``clear_contents`` over them would destroy the spill rather than
    reset a value.
    """
    overlap = set(_SPEC_INPUT_COLUMNS) & set(_DERIVED_COLUMNS)
    assert not overlap, f"computed displays must never be cleared: {overlap}"


def test_clear_set_has_no_duplicates() -> None:
    """A duplicated column would mean two constants resolved to one address."""
    assert len(_SPEC_INPUT_COLUMNS) == len(set(_SPEC_INPUT_COLUMNS))


def test_clear_set_covers_the_declared_input_band() -> None:
    """Spot-pin the individual inputs, so a silent renumbering is caught too."""
    for column in (
        _C_ROLE,
        _C_INCLUDE,
        _C_TYPE,
        _C_REFERENCE,
        _C_TRANSFORM,
        _C_SEQUENCE,
        _C_SEQUENCE_PERIOD,
        _C_INTERACTION_TERM,
        _C_INTERACTION_OPERATION,
    ):
        assert column in _SPEC_INPUT_COLUMNS

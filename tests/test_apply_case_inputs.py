"""The one helper that applies a fittable case's visible spec inputs.

``apply_case_inputs`` composes the three writes every spec-driven case needs on
top of a sheet whose spec block already exists — ``apply_spec_case`` (Source_Table
retarget, intercept, FE group, back-transform, the spec rows), the typed Sequence
Period into column I, and ``set_prediction_inputs``. It exists so a second call
site cannot forget the middle step, which is how the BFN panel Durbin-Watson cell
sat at ``nan`` for every verify run on record: the inspector duplicated this
sequence and dropped the override, and the only symptom was a multi-minute Excel
run that read like a broken statistic.

These tests exercise the full composition headlessly via ``RecordingSheet`` (the
same mock the spec-block writer tests use) — strengthened here with
``RecordingRange.clear_contents`` and ``RecordingNames.Item`` so ``apply_spec_case``,
which had never been driven under the mock before, runs in full. A behavioral
test is what the BFN gap always lacked: it would have gone red the moment the
override was dropped, instead of surfacing as plausible-looking noise in an
Excel run nobody re-checked.
"""
# pylint: disable=missing-function-docstring
from __future__ import annotations

import math
from pathlib import Path

from lambda_catalog.analyze_regression_spec import (
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.regression_layout import _C_AK
from lambda_catalog.regression_spec_sheet_io import (
    ROW_PRED_INPUT_FIRST,
    apply_case_inputs,
)
from lambda_catalog.write_sheet_test_model import _padded
from lambda_catalog.write_spec_block import (
    _C_ROLE as _C_SPEC_ROLE,
    _C_SEQUENCE_PERIOD as _C_SPEC_SEQUENCE_PERIOD,
    _FIRST_DATA_ROW,
)
from tests.recording_sheet import RecordingSheet

ROOT_DIR = Path(__file__).resolve().parent.parent


def _as_xw_sheet(sheet: RecordingSheet) -> RecordingSheet:
    return sheet  # type: ignore[return-value]


def _p01_expected() -> object:
    """P01 (production_lots_fixed_effects) — the case that makes the BFN cell live.

    It is one of only two cases that declare a Sequence Period (1.0), so it is
    the case the typed-column-I wiring exists for. Built (not padded here — the
    builder pads before calling the helper, so the test mirrors the builder and
    pads too) into a full ``RegressionSpecExpected`` with ``.results`` and
    ``.design``, the two halves ``set_prediction_inputs`` reads.
    """
    case = next(
        c for c in build_regression_spec_cases()
        if c.name == "production_lots_fixed_effects"
    )
    assert case.sequence_period == 1.0, "P01 is the case that makes BFN live"
    return _padded(calculate_regression_spec_case(case))


def test_apply_case_inputs_types_column_i_and_writes_spec_and_prediction_inputs() -> None:
    """The helper does all three writes, in the order the builder always did.

    Pins the composition end to end on the padded P01 case: the Sequence Period
    lands on the SEQUENCE-flagGED row only (the row Base_Period_Delta's XMATCH
    resolves), the spec rows are written for the declared width, and the
    prediction-input band receives each constructed column's mean (EXP'd back
    to input space for a Log column, which is the sheet's raw-value input
    convention). If any of the three is dropped — the regression that hid for
    every verify run on record — this fails headlessly.
    """
    expected = _p01_expected()
    case = expected.case  # type: ignore[attr-defined]
    sheet = RecordingSheet(name=case.sheet_name)
    # apply_spec_case retargets the Source_Table name, which the spec block
    # creates on a real sheet — pre-add it so the mock's Names.Item finds it.
    sheet.api.Names.Add(Name="Source_Table", RefersTo="")

    apply_case_inputs(_as_xw_sheet(sheet), expected)

    spec = case.spec  # type: ignore[attr-defined]

    # (a) typed Sequence Period: exactly the flagged row, value 1.0, nothing else.
    written_period = {
        offset: sheet.cell(_FIRST_DATA_ROW + offset, _C_SPEC_SEQUENCE_PERIOD).value
        for offset in range(len(spec))
        if sheet.cell(_FIRST_DATA_ROW + offset, _C_SPEC_SEQUENCE_PERIOD).value
        is not None
    }
    flagged = [offset for offset, item in enumerate(spec) if item.sequence]
    assert len(flagged) == 1, "exactly one Sequence axis, or H2 errors"
    assert written_period == {flagged[0]: 1.0}
    assert spec[flagged[0]].name == "Fiscal_Year"

    # (b) spec rows written: every declared row carries a Role, no more no less.
    roles = [
        sheet.cell(_FIRST_DATA_ROW + offset, _C_SPEC_ROLE).value
        for offset in range(len(spec))
    ]
    assert all(role is not None for role in roles)
    # The clear loop runs to max(_LAST_DATA_ROW+1, first+len-1); rows past the
    # declared width are blanked, not written, so they must stay None.
    assert sheet.cell(_FIRST_DATA_ROW + len(spec), _C_SPEC_ROLE).value is None

    # (c) prediction inputs: each constructed column's mean lands at AK(19+i),
    # EXP'd for a Log column (the sheet applies Ln_Positive itself, so inputs
    # are real-world values), and the band beyond the written cells is cleared.
    pred_values = expected.results.prediction_interval.pred_input_values  # type: ignore[attr-defined]
    transforms = expected.design.constructed_column_transforms  # type: ignore[attr-defined]
    expected_pred = [
        (math.exp(v) if transforms[i] == "Log" else v)
        for i, v in enumerate(pred_values)
    ]
    written_pred = [
        sheet.cell(ROW_PRED_INPUT_FIRST + i, _C_AK).value
        for i in range(len(pred_values))
    ]
    assert written_pred == expected_pred
    assert sheet.cell(
        ROW_PRED_INPUT_FIRST + len(pred_values), _C_AK
    ).value is None


def test_apply_case_inputs_skips_the_override_when_no_period_declared() -> None:
    """A case with no Sequence Period leaves column I untouched.

    The override is gated on ``case.sequence_period is not None``; only the
    BFN cases (P01/P02) declare one. A non-panel config must not get a typed
    period, and must still receive its spec and prediction inputs — the guard
    is what keeps the helper from writing a period a case never declared.
    """
    case = next(
        c for c in build_regression_spec_cases() if c.sequence_period is None
    )
    expected = _padded(calculate_regression_spec_case(case))
    spec = expected.case.spec  # type: ignore[attr-defined]
    sheet = RecordingSheet(name=case.sheet_name)
    sheet.api.Names.Add(Name="Source_Table", RefersTo="")

    apply_case_inputs(_as_xw_sheet(sheet), expected)

    # No row in the spec block received a typed Sequence Period.
    for offset in range(len(spec)):
        assert sheet.cell(
            _FIRST_DATA_ROW + offset, _C_SPEC_SEQUENCE_PERIOD
        ).value is None

    # The spec and prediction inputs still landed — the guard skips ONLY the
    # override, not the surrounding writes.
    assert sheet.cell(_FIRST_DATA_ROW, _C_SPEC_ROLE).value is not None
    pred_values = expected.results.prediction_interval.pred_input_values  # type: ignore[attr-defined]
    if pred_values:
        assert sheet.cell(ROW_PRED_INPUT_FIRST, _C_AK).value is not None


def test_inspector_routes_through_apply_case_inputs_not_the_primitives() -> None:
    """The verify inspector must delegate to the helper, not re-inline the steps.

    Re-inlining the three writes is exactly how the BFN cell lost its typed
    period: an inlined copy drifted and dropped the override, and the only
    symptom was a multi-minute Excel run. Pin that the inspector calls the helper
    and does not name the primitives it composes — so re-inlining fails the
    suite instead of failing a verify run nobody re-checks.
    """
    source = (ROOT_DIR / "tools" / "inspect_regression_sheet.py").read_text(
        encoding="utf-8"
    )

    assert "apply_case_inputs(" in source
    for forbidden in (
        "apply_spec_case",
        "set_prediction_inputs",
        "apply_sequence_period_overrides",
    ):
        assert forbidden not in source, (
            f"the inspector must route through apply_case_inputs, not re-inline {forbidden} "
            f"— re-inlining is how the typed Sequence Period was dropped"
        )
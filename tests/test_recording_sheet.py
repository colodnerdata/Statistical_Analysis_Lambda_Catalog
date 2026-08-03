"""Unit coverage for the ``RecordingSheet`` COM double's ``Formula2`` semantics.

The Weibull/Gamma profile stages write a whole 20x1 body in one
``Range.Formula2`` COM call (see ``_write_profile_stage`` in
``write_sheet_univariate.py``), so the recorder has to expand that write the
way Excel does: per constituent cell, with the per-cell state holding a
formula string rather than the 2D list.
"""
from tests.recording_sheet import RecordingSheet


def test_single_cell_formula2_round_trips_a_string() -> None:
    sheet = RecordingSheet()

    sheet.cell(3, 4).api.Formula2 = "=SUM(A1:A9)"

    assert sheet.cell(3, 4).api.Formula2 == "=SUM(A1:A9)"
    assert sheet.cell(3, 4).state.formula2 == "=SUM(A1:A9)"


def test_block_formula2_distributes_to_constituent_cells() -> None:
    sheet = RecordingSheet()
    block = sheet.range((2, 5), (3, 6))

    block.api.Formula2 = [["=A1", "=B1"], ["=A2", "=B2"]]

    assert sheet.cell(2, 5).api.Formula2 == "=A1"
    assert sheet.cell(2, 6).api.Formula2 == "=B1"
    assert sheet.cell(3, 5).api.Formula2 == "=A2"
    assert sheet.cell(3, 6).api.Formula2 == "=B2"


def test_block_formula2_state_stays_a_per_cell_string() -> None:
    """The 2D list is never stored in ``RecordingRangeState.formula2``.

    The block getter mirrors Excel by reassembling a 2D tuple from the cells,
    so the state field remains ``str | None`` on every range.
    """
    sheet = RecordingSheet()
    block = sheet.range((2, 5), (3, 6))

    block.api.Formula2 = [["=A1", "=B1"], ["=A2", "=B2"]]

    assert block.state.formula2 is None
    assert block.api.Formula2 == (("=A1", "=B1"), ("=A2", "=B2"))


def test_unwritten_cells_and_ranges_report_no_formula() -> None:
    sheet = RecordingSheet()

    assert sheet.cell(1, 1).api.Formula2 is None
    assert sheet.range((1, 1), (2, 2)).api.Formula2 is None

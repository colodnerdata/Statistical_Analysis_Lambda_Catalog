"""Headless structural guard for the v3.4 Model Comparison sheet.

**What this file is.** The part of the Model Comparison sheet's guard that needs
no Excel. It drives the sheet's REAL entry point — ``write_comparison_sheet``,
not a hand-rebuilt copy of its body — against ``RecordingSheet``, and pins the
structure the sheet is made of: which cell reads which ``Comparison_Field``
index, which cells each gate compares, which columns each red rule shades, which
cells the flag bands cover, and the number format of every mirrored statistic
against the Regression sheet that supplies it.

**What it deliberately cannot answer, and where that is answered instead.**

* Does ``OFFSET``, handed a cross-sheet reference through a sheet-scoped name,
  actually read the OTHER sheet? — ``tests/test_comparison_excel.py``. Only Excel
  can say, and the whole anchor mechanism rests on the answer.
* Do the two readers' numeric offsets still match the layout? —
  ``tests/test_comparison_offsets.py`` (JSON plus the layout constants, no Excel).
* **No formula pinned here is EVALUATED.** This file pins the formula the design
  calls for, not that it calculates. Nothing here is a claim that the workbook
  computes correctly.

**Why the structure deserves its own guard.** A reader written into the wrong
column still parses, still calculates, and returns a plausible number from the
neighbouring statistic — the class ``AGENTS.md`` treats as worse than a build
error. The same is true of a gate that compares a row against ITSELF (every
verdict reads "same", for every row, on every build) and of a red rule shading
the zone whose gate it does not read. None of those fail a build; all of them are
visible in the recorded structure, which is what this file reads.

**Why the entry point, not the private writers.** Driving ``write_comparison_sheet``
means the sheet's own sequence — names before cells, formats before rules — plus
the width and outline loops that only it performs are exercised here too. A test
that called the private writers in an order of its own invention would pin a
sequence this module does not ship.
"""
from __future__ import annotations

import re
from types import SimpleNamespace
from typing import cast

import xlwings as xw

from lambda_catalog import write_sheet_regression as regression_writer
from lambda_catalog.regression_layout import (
    _A_DESIGN_COLUMNS_TOTAL,
    _A_OBSERVATIONS,
    _A_RESPONSE_READOUT,
    _C_AF,
    _C_AK,
    _FMT_COUNT,
    _FMT_F_STATISTIC,
    _FMT_SIGNIFICANCE_F,
    _FMT_STAT,
    _PRED_INPUT_FIRST_ROW,
    _PRED_INPUT_LAST_ROW,
    _ROW_RESPONSE_READOUT,
)
from lambda_catalog.workbook_helpers import col_letter
from lambda_catalog.write_sheet_comparison import (
    SHEET_NAME,
    _abs_range,
    _anchor_name,
    _BAND_LAST_ROW,
    _C_ANCHOR,
    _C_BACK_TRANSFORM,
    _C_DESIGN_COLUMNS,
    _C_MODEL_SHEET,
    _C_OBSERVATIONS,
    _C_PRED_INPUTS,
    _C_PRED_MATCH,
    _C_RESPONSE,
    _C_RESPONSE_DECLARED,
    _C_RESPONSE_SPACE,
    _C_SAME_METHOD,
    _C_SAME_SET,
    _C_SAME_SPACE,
    _COLUMN_WIDTHS,
    _CONTENT_COLUMNS,
    _FIELD_COLUMNS,
    _FLAG_BANDS,
    _GAP_COLUMNS,
    _GATE_INPUTS,
    _GATE_METHOD,
    _GATE_SET,
    _GATE_SPACE,
    _INDEX_RESPONSE,
    _INPUTS_COLS_OFFSET,
    _INPUTS_HEIGHT,
    _INPUTS_ROWS_OFFSET,
    _LAST_COL,
    _NUMBER_FORMATS,
    _ROW_FIRST,
    _ROW_HEADER,
    _ROW_LAST,
    _ROW_STATUS,
    _ZONES,
    write_comparison_sheet,
)
from tests.recording_sheet import RecordingSheet
from tests.test_comparison_offsets import _FIELD_LAYOUT, _row_col

# The registry every test here builds with: one registered row plus the module's
# own default of three template rows. The name carries a space on purpose — the
# anchor RefersTo has to quote it, and a name without one would not notice a
# quoter that stopped working.
_REGISTRY = ("Mileage Data",)
_TEMPLATE_ROWS = 3
_ROW_LAST_DRIVEN = _ROW_FIRST + len(_REGISTRY) + _TEMPLATE_ROWS - 1

# ``$E5`` / ``$E$4`` — an absolute column, a relative or absolute row. Every
# comparison a gate makes is one ``=`` between two references of this shape.
_PAIR_RE = re.compile(r"(\$[A-Z]{1,3}\$?\d+)=(\$[A-Z]{1,3}\$?\d+)")
_READER_CALL_RE = re.compile(r"Comparison_Field\([^,]+,\s*(\d+)\)")

# Each gate: the noun its verdict carries, the column its flag lives in, and the
# cells it compares — self first, the reference row second. Data, so the bodies
# below are loops rather than four near-identical blocks; derived from the gate
# constants, so a gate wired to another axis fails.
_GATES: tuple[tuple[str, int, tuple[int, ...]], ...] = (
    (_GATE_SET, _C_SAME_SET, (_C_RESPONSE_DECLARED, _C_OBSERVATIONS)),
    (_GATE_SPACE, _C_SAME_SPACE, (_C_RESPONSE_SPACE,)),
    (_GATE_METHOD, _C_SAME_METHOD, (_C_BACK_TRANSFORM,)),
    (_GATE_INPUTS, _C_PRED_MATCH, (_C_DESIGN_COLUMNS, _C_PRED_INPUTS)),
)

# The statistic zones in layout order (``_ZONES[1:]`` — the registry is first and
# nothing invalidates an identity fact), each with the gates that must hold for it
# to be readable. This is the design claim the red rules encode, restated from the
# gate constants rather than from the expressions the writer builds: the unit-space
# zone survives a change of response space, the fit-space zone does not, both are
# moved by the back-transform method, and the prediction zone additionally needs
# the two rows evaluated at the same point.
_ZONE_GATES: tuple[tuple[str, ...], ...] = (
    (_GATE_SET, _GATE_METHOD),
    (_GATE_SET, _GATE_SPACE),
    (_GATE_SET, _GATE_METHOD, _GATE_INPUTS),
)

# The Regression writers that own the format of every mirrored statistic, one per
# zone carrying one. Driven together so the comparison sheet's formats can be read
# against what the Regression sheet actually emits.
_REGRESSION_FORMAT_WRITERS = (
    regression_writer._write_regression_statistics,
    regression_writer._write_diagnostics,
    regression_writer._write_anova,
    regression_writer._write_unit_space_block,
    regression_writer._write_prediction_interval,
)

# The one comparison column that is NOT a mirrored Regression statistic. k reads
# the spec block's Σ Design Columns total, which that block leaves General — there
# is no Regression number format for it to agree with.
_NOT_MIRRORED = frozenset({_C_DESIGN_COLUMNS})

_GATE_COLUMN: dict[str, int] = {gate: col for gate, col, _ in _GATES}


# ── Harness ───────────────────────────────────────────────────────────────────

def _drive() -> RecordingSheet:
    """Build the sheet through its REAL entry point, on the recording mock.

    ``get_or_create_sheet`` looks the mock up by name, so the workbook stand-in
    only has to be the collection it searches — nothing else in the entry point
    reaches the workbook, which is why a list of one sheet is enough.
    """
    sheet = RecordingSheet(name=SHEET_NAME)
    write_comparison_sheet(
        cast(xw.Book, SimpleNamespace(sheets=[sheet])),
        registry=_REGISTRY,
        template_rows=_TEMPLATE_ROWS,
    )
    return sheet


def _letter(col: int) -> str:
    return col_letter(col)


def _formula(sheet: RecordingSheet, row: int, col: int) -> str | None:
    """The cell's formula, or ``None`` where nothing was written.

    ``None`` is the common case when scanning a whole row: only the columns the
    writer fills have one, and the gap and unformatted columns are blank.
    """
    return cast("str | None", sheet.cell(row, col).api.Formula2)


def _pairs(formula: str) -> list[tuple[str, str]]:
    return _PAIR_RE.findall(formula)


def _flag(col: int, row: int = _ROW_FIRST) -> str:
    """``$J4`` — the flag reference a rule reads, at the first data row."""
    return f"${_letter(col)}{row}"


def _reader_calls(formula: str | None) -> list[str]:
    return [] if not formula else _READER_CALL_RE.findall(formula)


def _column_format(sheet: RecordingSheet, col: int) -> str | None:
    """The format this sheet applies to a whole column band.

    Read through the address the writer used (``H4:H103``) rather than per cell:
    the recorder keys a range by the address token it was handed, and the
    comparison sheet's formats are applied one band at a time. Same cells in
    Excel either way, and the assertion is about what the writer emitted.
    """
    letter = _letter(col)
    return cast(
        "str | None",
        sheet.range(f"{letter}{_ROW_FIRST}:{letter}{_BAND_LAST_ROW}").number_format,
    )


# ── The readers ───────────────────────────────────────────────────────────────

def test_every_reader_cell_holds_its_own_index_in_its_own_column() -> None:
    """The position table, cell by cell, for every row the sheet ships.

    The index is HALF a reader: index 6 is F and displays in a fit-space column,
    index 24 is the declared response and displays in the registry zone. A swap
    between two entries reads a real statistic under the wrong header.
    """
    sheet = _drive()
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        name = _anchor_name(row)
        for index, col in _FIELD_COLUMNS:
            assert _formula(sheet, row, col) == (
                f'=IFERROR(Comparison_Field({name},{index}),"")'
            ), f"row {row} column {col} is not reader index {index}"


def test_the_only_other_reader_call_is_the_guard_in_the_sheet_name_cell() -> None:
    """No strays, and no reader written twice.

    The sheet-name cell is the ONE other ``Comparison_Field`` call site: it reads
    field 1 to ask whether the anchor points at a Regression-shaped sheet at all
    (the reader's A1 probe), which is what makes a mis-pointed row read blank
    rather than plausible. Any other cell holding a reader call is a duplicate of
    a cell already pinned above, or a reader in a column the table does not name.
    """
    sheet = _drive()
    readers = {
        (row, col): index
        for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1)
        for index, col in _FIELD_COLUMNS
    }
    strays: list[str] = []
    for (row, col), index in readers.items():
        calls = _reader_calls(_formula(sheet, row, col))
        if calls != [str(index)]:
            strays.append(f"reader cell ({row},{col}) calls {calls}, expected [{index}]")
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        for col in range(1, _LAST_COL + 1):
            if (row, col) in readers:
                continue
            calls = _reader_calls(_formula(sheet, row, col))
            if not calls:
                continue
            if col != _C_MODEL_SHEET or calls != [str(_INDEX_RESPONSE)]:
                strays.append(f"unexpected reader call in ({row},{col}): {calls}")
    assert not strays, "\n".join(strays)


# ── The gates ─────────────────────────────────────────────────────────────────

def test_every_gate_compares_a_row_against_the_reference_row_not_itself() -> None:
    """The bug this catches: a gate wired to the row's OWN cells.

    Every verdict would then read "same", for every row, on every build — a sheet
    reporting perfect agreement between models that disagree, which is the most
    confidently wrong thing this sheet could do. Asserted on rows past the first
    deliberately: row 4 IS the reference row, so there its own cells and the
    reference cells are the same cells and a self-comparing gate is invisible.
    """
    sheet = _drive()
    for row in range(_ROW_FIRST + 1, _ROW_LAST_DRIVEN + 1):
        for gate, col, _columns in _GATES:
            pairs = _pairs(_formula(sheet, row, col))
            assert pairs, f"row {row} gate {gate} compares nothing"
            for left, right in pairs:
                left_row, _ = _row_col(left)
                right_row, _ = _row_col(right)
                assert left_row == row, (
                    f"row {row} gate {gate} compares {right} against {left}, "
                    "which is not this row — one side must be the row itself"
                )
                assert right_row == _ROW_FIRST, (
                    f"row {row} gate {gate} compares against row {right_row}, not "
                    f"the reference row {_ROW_FIRST} — every row would be "
                    "compared with itself"
                )


def test_each_gate_compares_the_two_cells_it_names() -> None:
    """Which cells, not just which rows: F against F, declared response against
    declared response.

    The expectations are built from the column constants and the absolute-row
    address form the writer uses (`self_ref`/`ref_ref` are both `$E$5`-shaped —
    the gate cells are single cells, not a spill), never from the writer's own
    strings, so swapping two gates — or wiring the method gate to the
    response-space column — fails here rather than shipping a verdict about
    another axis.
    """
    sheet = _drive()
    row = _ROW_FIRST + 1
    for gate, col, columns in _GATES:
        expected = [
            (f"${_letter(column)}${row}", f"${_letter(column)}${_ROW_FIRST}")
            for column in columns
        ]
        assert _pairs(_formula(sheet, row, col)) == expected, (
            f"gate {gate} (column {col}) compares "
            f"{_pairs(_formula(sheet, row, col))}, expected {expected}"
        )


def test_the_reference_row_compares_against_itself_by_construction() -> None:
    """Row 4 is both sides, which is correct — and is what the test above needs.

    Stated so the difference is explicit rather than looking like an oversight:
    the reference row is the comparison baseline, so its gates are trivially
    satisfied. That is why every self-comparison guard here uses rows 5+.
    """
    sheet = _drive()
    for gate, col, columns in _GATES:
        for left, right in _pairs(_formula(sheet, _ROW_FIRST, col)):
            assert left == right, f"gate {gate} on the reference row: {left} != {right}"
        assert len(columns) == len(_pairs(_formula(sheet, _ROW_FIRST, col)))


def test_every_gate_blanks_an_unregistered_row_instead_of_failing_it() -> None:
    """A template row's flags are ``""``, not FALSE.

    Both the row-2 verdicts and the red rules depend on that: the rules test
    ``$J4<>""`` first precisely because an unregistered row's flags are empty, and
    a gate that returned FALSE would shade a blank row as a failed comparison
    instead of as an empty one.
    """
    sheet = _drive()
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        for gate, col, _columns in _GATES:
            formula = _formula(sheet, row, col)
            assert formula.startswith("=IF("), f"row {row} gate {gate}: {formula}"
            assert '=""' in formula, (
                f"row {row} gate {gate} has no empty-row branch: {formula}"
            )


# ── The flag bands and the row-2 verdicts ─────────────────────────────────────

def test_the_flag_bands_cover_the_shipped_rows_and_run_past_them() -> None:
    """The bands are what the row-2 verdicts count over.

    They start at the first data row and run PAST the last shipped row: someone
    who copies a template row downward is adding a model, and a verdict that
    stopped counting at the last row the build wrote would silently ignore it. A
    band starting late would miss the reference row itself.
    """
    sheet = _drive()
    assert _BAND_LAST_ROW > _ROW_LAST, (
        "the band must run past the rows the build ships, or a row copied "
        "downward falls outside the verdict"
    )
    assert _BAND_LAST_ROW > _ROW_LAST_DRIVEN
    for name, col, _gate in _FLAG_BANDS:
        recorded = sheet.api.Names.by_short_name(name)
        assert recorded.RefersTo == f"='{SHEET_NAME}'!{_abs_range(col, col)}", (
            f"band {name} covers {recorded.RefersTo}, expected its own flag "
            f"column {col} over rows {_ROW_FIRST}-{_BAND_LAST_ROW}"
        )
        assert recorded.Comment, f"band {name} carries no Name Manager comment"


def test_each_row_two_verdict_counts_its_own_band_and_names_its_own_axis() -> None:
    """The verdict sits in its flag's column and passes that gate's noun.

    ``Comparison_Flag_Status`` counts TRUE/FALSE over a band and returns a
    sentence naming the axis. Pairing a verdict with the wrong band reports one
    axis's count in another axis's column — a number that is right about
    something else, which reads as correct.
    """
    sheet = _drive()
    assert len(_FLAG_BANDS) == 4, "the sheet ships four gates"
    for name, col, gate in _FLAG_BANDS:
        assert _formula(sheet, _ROW_STATUS, col) == (
            f'=Comparison_Flag_Status({name},"{gate}")'
        ), (
            f"the {gate} verdict in column {col} is "
            f"{_formula(sheet, _ROW_STATUS, col)}, expected it to count {name} "
            f"and name its own axis"
        )
        assert sheet.cell(_ROW_STATUS, col).api.WrapText is True, (
            "row 2 owns a column with little runway; every verdict wraps"
        )
    assert {gate for _, _, gate in _FLAG_BANDS} == set(_GATE_COLUMN), (
        "the four band nouns must be the four gates this sheet is built from"
    )
    assert {col for _, col, _ in _FLAG_BANDS} == set(_GATE_COLUMN.values()), (
        "each band covers the column its gate lives in, and no other"
    )


def test_the_bands_are_the_four_distinct_columns_the_gates_use() -> None:
    """Four gates, four columns, four bands — no two sharing a column.

    Two bands over one column would make one verdict count the other's rows.
    """
    assert len({col for _, col, _ in _FLAG_BANDS}) == len(_FLAG_BANDS)
    assert len({name for name, _, _ in _FLAG_BANDS}) == len(_FLAG_BANDS)
    assert len({gate for _, _, gate in _FLAG_BANDS}) == len(_FLAG_BANDS)


# ── Conditional formatting ────────────────────────────────────────────────────

def test_each_statistic_zone_is_shaded_by_the_gates_that_gate_it() -> None:
    """One red rule per statistic zone, reading that zone's gates.

    The rule addresses are the zone spans themselves, and the expressions are
    rebuilt here from the gate constants. A rule shading the fit-space zone with
    the method gate would leave a log fit's R² looking comparable.
    """
    sheet = _drive()
    zones = _ZONES[1:]
    assert len(zones) == len(_ZONE_GATES), (
        f"{len(_ZONE_GATES)} zone rules are written for {len(zones)} statistic "
        "zones — the writer's rule list and the zone list have diverged"
    )
    for (first_col, last_col), gates in zip(zones, _ZONE_GATES):
        conditions = sheet.range(_abs_range(first_col, last_col)).api.FormatConditions.items
        assert len(conditions) == 1, f"zone {first_col}-{last_col}: {len(conditions)} rules"
        flags = [_flag(_GATE_COLUMN[gate]) for gate in gates]
        expected = f'=AND({flags[0]}<>"",NOT(AND({",".join(flags)})))'
        assert conditions[0].Formula1 == expected, (
            f"zone {first_col}-{last_col} is shaded by "
            f"{conditions[0].Formula1}, expected {expected}"
        )


def test_each_flag_column_shades_its_own_false() -> None:
    """A FALSE shades itself, so it is visible without reading across the sheet.

    The ``<>""`` half is load-bearing: without it an unregistered row — flags all
    empty — satisfies ``=FALSE``, and a blank template row reads as a failed
    comparison.
    """
    sheet = _drive()
    for _name, col, gate in _FLAG_BANDS:
        conditions = sheet.range(_abs_range(col, col)).api.FormatConditions.items
        assert len(conditions) == 1, f"{gate}: {len(conditions)} rules on its flag column"
        flag = _flag(col)
        assert conditions[0].Formula1 == f'=AND({flag}<>"",{flag}=FALSE)'


def test_the_red_rules_stop_where_the_flag_bands_stop() -> None:
    """One ceiling for the bands, the number formats and the rules.

    A rule whose range ended before the band it reads would leave the rows past
    its end unshaded while the row-2 verdict still counted them — the two
    disagreeing about how far the sheet reaches.
    """
    sheet = _drive()
    for first_col, last_col in _ZONES[1:]:
        assert _abs_range(first_col, last_col).endswith(f"${_BAND_LAST_ROW}")
    for _name, col, _gate in _FLAG_BANDS:
        assert _abs_range(col, col).endswith(f"${_BAND_LAST_ROW}")
    for col, _fmt in _NUMBER_FORMATS:
        letter = _letter(col)
        assert sheet.range(f"{letter}{_ROW_FIRST}:{letter}{_BAND_LAST_ROW}").number_format


# ── Registry wiring ───────────────────────────────────────────────────────────

def test_registry_rows_reference_a_sheet_and_template_rows_point_at_a1() -> None:
    """The anchor IS the registry: one sheet-scoped name per row.

    A registered row's name is a reference to the target's response-label cell,
    quoted (a sheet name with a space would be an invalid formula unquoted). An
    unregistered row's name must point at this sheet's own ``A1``, whose text is
    not the guard literal, so the row reads blank rather than a plausible number —
    and pointing one name is the whole of the add-a-model story.
    """
    sheet = _drive()
    names = [item.Name.split("!", 1)[-1] for item in sheet.api.Names.items]
    registered_target = f"='{_REGISTRY[0]}'!{_A_RESPONSE_READOUT}"
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        name = _anchor_name(row)
        assert names.count(name) == 1, f"{name} is registered {names.count(name)} times"
        recorded = sheet.api.Names.by_short_name(name)
        assert recorded.Comment, f"{name} carries no Name Manager comment"
        if row == _ROW_FIRST:
            assert recorded.RefersTo == registered_target
        else:
            assert recorded.RefersTo == f"='{SHEET_NAME}'!$A$1", (
                f"{name} points at {recorded.RefersTo}; a template row must point "
                "at this sheet's own A1 so its guard fails loudly"
            )


def test_row_three_heads_every_zone_and_every_content_column_is_named_once() -> None:
    """The three-row header grammar, and one header per content column.

    Row 1 is zone labels, row 2 verdicts, row 3 headers, data from row 4. A header
    on the wrong row puts a label inside the data band, where it reads as a value
    on a sheet whose rows are all formulas; a header in a gap column breaks the
    gap's job of separating two zones.
    """
    sheet = _drive()
    for col in _CONTENT_COLUMNS:
        header = sheet.cell(_ROW_HEADER, col).value
        assert isinstance(header, str) and header, f"column {col} has no header"
    for col in _GAP_COLUMNS:
        assert sheet.cell(_ROW_HEADER, col).value in (None, ""), (
            f"gap column {col} must stay empty"
        )
        assert sheet.cell(_ROW_STATUS, col).api.Formula2 is None, (
            f"no verdict belongs in gap column {col}"
        )
    for first_col, _last in _ZONES:
        label = sheet.cell(1, first_col).value
        assert isinstance(label, str) and label, f"zone at column {first_col} has no label"


# ── Layout: widths and outline groups ─────────────────────────────────────────

def test_every_content_column_is_sized_once_and_every_gap_stays_narrow() -> None:
    """Widths come from the table; the gaps are the only columns set by policy.

    The gaps are what let neighbouring zones collapse independently, so a gap
    sized like a content column would read as belonging to a zone.
    """
    sheet = _drive()
    sized = sorted(col for col, _ in _COLUMN_WIDTHS)
    assert sized == sorted(_CONTENT_COLUMNS), (
        "every content column must be sized exactly once and no others"
    )
    assert not set(_GAP_COLUMNS) & set(sized)
    for col, width in _COLUMN_WIDTHS:
        letter = _letter(col)
        assert sheet.range(f"{letter}:{letter}").column_width == width, (
            f"column {letter} is not at its declared width {width}"
        )
    gap_widths = {sheet.range(f"{_letter(col)}:{_letter(col)}").column_width
                  for col in _GAP_COLUMNS}
    assert gap_widths == {2}, (
        f"every gap column shares one narrow width, found {gap_widths}"
    )


def test_the_gaps_are_exactly_the_columns_between_consecutive_zones() -> None:
    """A gap is one column and sits BETWEEN two zones, never inside one.

    The same principle as the Regression sheet's: the ungrouped column is what
    stops Excel fusing two zones into a single collapse control.
    """
    derived = tuple(_ZONES[i][1] + 1 for i in range(len(_ZONES) - 1))
    assert _GAP_COLUMNS == derived, (
        f"gaps are {_GAP_COLUMNS}, expected one column after each zone: {derived}"
    )
    for gap in _GAP_COLUMNS:
        assert all(not (first <= gap <= last) for first, last in _ZONES), (
            f"column {gap} is a gap AND inside a zone"
        )
    assert max(last for _first, last in _ZONES) == _LAST_COL, (
        "the last zone must end at the sheet's last content column"
    )


def test_each_zone_is_its_own_outline_group_and_no_gap_is_grouped() -> None:
    """Groups are the zones; the gaps are left out of every one."""
    sheet = _drive()
    expected = [f"{_letter(first)}:{_letter(last)}" for first, last in _ZONES]
    assert sheet.column_groups == expected, (
        f"outline groups are {sheet.column_groups}, expected one per zone: {expected}"
    )
    for gap in _GAP_COLUMNS:
        assert str(gap) not in {part for group in sheet.column_groups for part in group}


# ── The prediction-inputs offset ──────────────────────────────────────────────

def test_the_prediction_input_offset_is_recomputed_from_the_layout() -> None:
    """The one offset written from a cell on THIS sheet is derived, not spelled.

    The inputs cell reads the target's prediction-input block by offsetting from
    the anchor, so its three numbers are a claim about where that block sits on
    every Regression sheet. Recomputed here from the same constants the block is
    written from and then matched against the emitted formula — a zone shift moves
    the block and fails this, rather than reading a plausible number from whatever
    now occupies the old cell.
    """
    assert _INPUTS_ROWS_OFFSET == _PRED_INPUT_FIRST_ROW - _ROW_RESPONSE_READOUT
    assert _INPUTS_COLS_OFFSET == _C_AK - _C_AF
    assert _INPUTS_HEIGHT == _PRED_INPUT_LAST_ROW - _PRED_INPUT_FIRST_ROW + 1
    assert _INPUTS_HEIGHT > 1, "a one-row block cannot hold several predictors"

    sheet = _drive()
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        formula = _formula(sheet, row, _C_PRED_INPUTS)
        expected = (
            f"OFFSET({_anchor_name(row)},{_INPUTS_ROWS_OFFSET},"
            f"{_INPUTS_COLS_OFFSET},{_INPUTS_HEIGHT},1)"
        )
        assert expected in formula, (
            f"row {row} reads its inputs as {formula}, expected it to contain "
            f"{expected}"
        )


def test_every_derived_cell_blanks_an_unregistered_row() -> None:
    """A template row must show empty cells, not errors.

    Its anchor points at this sheet's A1, so every reader fails; "ready to be
    pointed at a model" looks like blank cells. An error there would read as a
    broken sheet rather than as an unregistered row.

    The branch is an empty-string result, not necessarily a `=""` comparison: the
    anchor readout blanks through its inner IFERROR fallback, because for a
    reference on THIS sheet there is no `]` to split on and the CELL call itself
    is the second-best answer.
    """
    sheet = _drive()
    for row in range(_ROW_FIRST + 1, _ROW_LAST_DRIVEN + 1):
        for col in (_C_PRED_INPUTS, _C_PRED_MATCH, _C_SAME_SET, _C_ANCHOR):
            formula = _formula(sheet, row, col)
            assert formula is not None and '""' in formula, (
                f"({row},{col}) has no empty-string branch: {formula}"
            )


# ── The shared display formats ────────────────────────────────────────────────

def test_every_mirrored_statistic_is_formatted_as_the_regression_sheet_formats_it() -> None:
    """One number, one format — checked by driving BOTH writers.

    The comparison sheet's whole claim is that its cells are the target sheet's own
    values, and a mirrored statistic drawn at a different precision invites the
    reader to see a difference that is not there (976.1166 beside 976.1). Both
    writers are driven here and their emitted formats compared per statistic, so a
    drift on EITHER side fails — the Regression sheet moving to a new precision
    breaks the pair as much as this sheet choosing one.
    """
    comparison = _drive()
    regression = RecordingSheet(name="Regression")
    for writer in _REGRESSION_FORMAT_WRITERS:
        writer(cast(xw.Sheet, regression))

    # Field index → the Regression cell it mirrors, from the two tables that
    # already own those halves: the comparison column table, and the readers'
    # position table (which pins each index to a regression_layout anchor). The
    # position table covers the 23 single cells; index 24 is the declared
    # response variable — read as a ranged lookup over the spec block, and text,
    # so it has no number format to compare.
    mirrored = {
        col: _FIELD_LAYOUT[index - 1]
        for index, col in _FIELD_COLUMNS
        if index <= len(_FIELD_LAYOUT)
    }
    assert mirrored, "the position table and the column table share no entries"
    checked = set()
    for col, fmt in _NUMBER_FORMATS:
        if col in _NOT_MIRRORED:
            continue
        assert col in mirrored, (
            f"column {col} carries a number format but mirrors no single "
            "Regression cell, so nothing here can check it"
        )
        address = mirrored[col]
        source_row, source_col = _row_col(address)
        source_format = regression.cell(source_row, source_col).number_format
        assert source_format is not None, (
            f"the Regression sheet applies no number format to {address}, the cell "
            f"this sheet's column {col} mirrors — the pair has no anchor"
        )
        emitted = _column_format(comparison, col)
        assert emitted == fmt == source_format, (
            f"column {col} is drawn at {emitted} where its source {address} on the "
            f"Regression sheet is drawn at {source_format}, and the shared table "
            f"declares {fmt}: one number, two precisions"
        )
        checked.add(col)
    assert checked == {col for col, _ in _NUMBER_FORMATS} - _NOT_MIRRORED, (
        "every mirrored column in the format table must be checked"
    )


def test_every_number_format_comes_from_the_shared_table() -> None:
    """No format literal is left inline on this sheet.

    The formats are declared once in ``regression_layout`` (``_FMT_*``) so this
    sheet and the Regression sheet cannot disagree about a mirrored number. A
    literal typed here would be a third copy — and the one nothing checks.
    """
    assert {fmt for _, fmt in _NUMBER_FORMATS} <= {
        _FMT_COUNT, _FMT_STAT, _FMT_F_STATISTIC, _FMT_SIGNIFICANCE_F,
    }, (
        "every format on this sheet must come from the shared _FMT_* table in "
        "regression_layout, not from a literal: "
        f"{sorted({fmt for _, fmt in _NUMBER_FORMATS})}"
    )
    assert _FMT_SIGNIFICANCE_F != _FMT_STAT, (
        "a p-value and a goodness-of-fit statistic cannot share a format: four "
        "fixed decimals renders both 3E-08 and 4E-05 as 0.0000"
    )


def test_the_counts_read_as_counts_and_only_reader_columns_are_formatted() -> None:
    """k is the documented exception; nothing text-shaped carries a format."""
    formats = {col: fmt for col, fmt in _NUMBER_FORMATS}
    assert formats[_C_OBSERVATIONS] == _FMT_COUNT
    assert formats[_C_DESIGN_COLUMNS] == _FMT_COUNT
    assert _NOT_MIRRORED == {_C_DESIGN_COLUMNS}, (
        "the exception list must name exactly the column the table documents"
    )
    reader_columns = {col for _, col in _FIELD_COLUMNS}
    for col, fmt in _NUMBER_FORMATS:
        assert fmt, f"column {col} was given an empty number format"
        assert col in reader_columns, (
            f"column {col} carries a number format but is not a reader column"
        )
    for col in (_C_RESPONSE, _C_ANCHOR, _C_MODEL_SHEET, _C_RESPONSE_DECLARED):
        assert col not in formats, f"column {col} is text and must not be formatted"


# ── The sheet-name cell ───────────────────────────────────────────────────────

def test_the_sheet_name_link_escapes_its_link_target_and_not_its_label() -> None:
    """A sheet name may contain an apostrophe; the link must survive it.

    ``"#'" & nm & "'!$A$1"`` terminates the quoted sheet name at the first
    apostrophe inside it, so a target called ``Bob's Data`` produces a link to a
    sheet that does not exist. The apostrophes must be doubled — the Excel-side
    mirror of ``workbook_helpers.quoted_sheet_name`` — and only in the LINK
    LOCATION: the label shows the real name, undoubled.
    """
    sheet = _drive()
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        formula = _formula(sheet, row, _C_MODEL_SHEET)
        assert 'SUBSTITUTE(nm,"\'","\'\'")' in formula, (
            f"row {row}'s hyperlink does not escape apostrophes in the target "
            f"sheet name: {formula}"
        )
        assert formula.endswith(',nm)),""))'), (
            f"row {row}'s hyperlink must display the bare sheet name: {formula}"
        )


def test_the_anchor_cell_shows_the_address_the_row_actually_reads() -> None:
    """The anchor address is a debugging aid: it names the cell being read.

    For a mis-pointed anchor this is the cell the row is really reading, which is
    what someone diagnosing a blank row needs. It falls back for a reference on
    this sheet, where CELL returns a bare ``$A$1`` with no ``[book]Sheet!`` prefix
    to split on — which is exactly a template row's state.
    """
    sheet = _drive()
    for row in range(_ROW_FIRST, _ROW_LAST_DRIVEN + 1):
        formula = _formula(sheet, row, _C_ANCHOR)
        assert "CELL(\"address\"" in formula and _anchor_name(row) in formula
        assert "IFERROR(" in formula, (
            "the anchor readout must survive a reference on this sheet, where "
            f"there is no ] to split on: {formula}"
        )

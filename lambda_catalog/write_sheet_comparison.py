"""Write the Model Comparison sheet into the target workbook.

**What this sheet is.** Every other sheet in this workbook fits ONE model. This
one reads the *other* sheets: a row per fitted model, its goodness-of-fit
statistics side by side, and its prediction — so an analyst comparing candidate
specifications reads one table instead of copying numbers out of N sheets.

**Why the columns are grouped the way they are.** The organising axis is
comparability, and the sheet enforces it rather than asserting it in prose. The
grouping is about what each statistic is MEASURED AGAINST, which decides whether
two rows' numbers can be read beside each other at all:

* a different **response variable** or a different **sample** (n) makes
  everything incomparable — the two rows are answers to different questions;
* a different **response space** (a Log fit beside an untransformed one) puts the
  fit-space statistics on different scales. This is not a matter of degree: the
  log-Jacobian alone moves this workbook's Life Expectancy AIC by tens of
  thousands, and F is a test of the fitted response, so a Log fit's F tests a
  different hypothesis from a level fit's;
* a different **back-transform method** (Duan beside Naive) leaves the fit
  identical and moves only the original-units numbers. Comparing two rows that
  differ that way compares a display choice, not two models.

The unit-space triplet and the prediction block sit in response units, so a
difference between two rows there is a real difference in how each model
predicts the actual response — which is what makes them readable across a change
of response transform, and the reason the Regression sheet exposes a Response
Space readout at all. They are not *invariant* to that change (the fitted
function itself changes, so the numbers move) and they are not immune to a change
of method, which is why the method is its own gate rather than folded into the
space.

So the columns fall into three statistic zones with three different gates, each
carried as conditional formatting on its zone: UNIT-SPACE STATISTICS (gated on
the comparison set *and* the method), FIT-SPACE STATISTICS (gated on the set and
the response space), PREDICTION COMPARISON (gated on the set, the method, and
the two rows predicting the same point). Four flag columns in the registry zone
state the four gates per row, and the row-2 verdicts summarise them. This is the
dynamic counterpart of the Model Comparison Guide's comparability section.

**The comparison set keys on the DECLARED response variable, not the fitted
one.** The Regression sheet's ``AF3`` readout shows the response as *fitted* — it
reads ``Ln(MPG)`` for a logged fit and ``MPG`` for a level fit of the same data.
Keying the set on that label would call one response two and shade exactly the
Log-versus-level pair the unit-space columns exist to compare. The reader
therefore exposes the declared name as its own field (``Comparison_Field`` index
24), read out of the target's spec block, and that is what the set test compares.

**The anchor, and why it is a defined name.** Each row reads its target sheet
through ``Comparison_Field``/``Model_Formula_String``, which take a **reference**
to that sheet's response-label cell and offset from it. Excel's ``OFFSET`` reads
its reference argument's own location and does not dereference it, so the
reference cannot be held in a cell and passed along — a cell holding
``='Other'!$AF$3`` offsets from *that cell*, silently. It has to be a reference
written into the formula, and the one editable, single-source place to put one
is a **defined name**. So each row owns ``Comp_Anchor_<n>``, whose ``RefersTo``
is ``'<target sheet>'!$AF$3``; every cell in the row derives from it. Adding a
model is therefore one Name Manager edit, not a formula rewrite.

**A mis-pointed anchor is loud, not silent.** ``Comparison_Field`` probes
``OFFSET(anchor,-2,-31)`` for the literal ``"MODEL SPECIFICATION"`` that every
Regression-shaped sheet carries at ``A1``, and returns ``#N/A`` when the probe
fails. Every read here wraps in ``IFERROR(..., "")``, so a row pointed at a
non-model sheet is **blank**, never a number from an unrelated cell.

**No typed inputs.** The sheet is read-only by design: the prediction machinery
(``Group_Prediction_Interval``, ``Predictor_Columns()``, ``Fit_Context()``) is
sheet-scoped and cannot be invoked across sheets, so a typed "shared prediction
inputs" row here could not compute anything for any model — it would be a
control that silently does nothing. The row's inputs are shown as text (read
from the target sheet's own band) and compared, which is what a comparison
actually needs. There is consequently no ``INPUT_COLOR`` cell on this sheet; the
only edit surface is the Name Manager.
"""
from __future__ import annotations

from collections.abc import Sequence

import xlwings as xw

from .catalog_schema import CatalogFunction
from .regression_layout import (
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
from .sheet_styles import CF_DARK_RED_TEXT as _CF_RED_TEXT
from .sheet_styles import CF_LIGHT_RED_FILL as _CF_RED_FILL
from .sheet_styles import SUBHDR_COLOR as _SUBHDR
from .workbook_helpers import (
    add_expression_format,
    anchor_comment_right_of_cell,
    border_box,
    col_letter,
    drop_local_name,
    f,
    get_or_create_sheet,
    note_dimensions,
    quoted_sheet_name,
    rc,
    reset_column_groups,
    reset_generated_sheet,
    safe_activate,
    section_heading,
)

SHEET_NAME = "Model Comparison"

# ── Row grammar ───────────────────────────────────────────────────────────────
# Same three-row frozen band as the Regression sheet: row 1 = zone labels, row 2
# = the status/verdict cells (one per gate, in the column that gate's flag lives
# in), row 3 = column headers everywhere, row 4+ = data. Freeze at A4.
_ROW_ZONE_LABEL = 1
_ROW_STATUS = 2
_ROW_HEADER = 3
_ROW_FIRST = 4

# The build ships one registered row in the production workbook (the single
# Regression sheet) and a few unregistered template rows. A template row's
# anchor name points at this sheet's own $A$1, whose text is not
# "MODEL SPECIFICATION", so its guard fails and every cell in the row reads
# blank until the name is pointed at a real model sheet.
_REGISTERED_ROWS = 1
_TEMPLATE_ROWS = 3
_ROW_LAST = _ROW_FIRST + _REGISTERED_ROWS + _TEMPLATE_ROWS - 1

# Where the flag bands, number formats and conditional formats stop. Generous on
# purpose: these bands are what the row-2 verdicts count over, so the writer can
# ship additional pre-wired template rows (or be rebuilt with a larger
# `template_rows`) without having to resize every rule/range.
# Blank cells are counted as neither TRUE nor FALSE, so a wide band costs
# nothing at calculation time.
_BAND_LAST_ROW = 103

# ── Columns ───────────────────────────────────────────────────────────────────
# Zone 1 — MODEL REGISTRY. The identity facts (n, k, the declared response
# variable) live here rather than in a statistic zone: nothing invalidates them,
# so a gate that shaded them would be lying about what a gate means.
_C_MODEL_SHEET = 1
_C_ANCHOR = 2
_C_MODEL = 3
_C_RESPONSE = 4
_C_RESPONSE_DECLARED = 5
_C_RESPONSE_SPACE = 6
_C_BACK_TRANSFORM = 7
_C_OBSERVATIONS = 8
_C_DESIGN_COLUMNS = 9
_C_SAME_SET = 10
_C_SAME_SPACE = 11
_C_SAME_METHOD = 12
_GAP_1 = 13
# Zone 2 — UNIT-SPACE STATISTICS: the Regression sheet's own
# Comparison_Headline_GoF triplet. In response units, so readable across a
# response-transform change; moved by the back-transform method, which is why
# the method is part of this zone's gate.
_C_UNIT_R_SQUARED = 14
_C_UNIT_ADJUSTED_R_SQUARED = 15
_C_UNIT_RMSE = 16
_GAP_2 = 17
# Zone 3 — FIT-SPACE STATISTICS: measured against the fitted (possibly
# transformed) response, so comparable only within one response space.
_C_F_STATISTIC = 18
_C_SIGNIFICANCE_F = 19
_C_R_SQUARED = 20
_C_ADJUSTED_R_SQUARED = 21
_C_STANDARD_ERROR = 22
_C_AIC = 23
_C_BIC = 24
_C_AICC = 25
_C_PRESS = 26
_C_PRESS_R_SQUARED = 27
_GAP_3 = 28
# Zone 4 — PREDICTION COMPARISON (original units).
_C_PRED_POINT = 29
_C_PRED_CI_LOWER = 30
_C_PRED_CI_UPPER = 31
_C_PRED_PI_LOWER = 32
_C_PRED_PI_UPPER = 33
_C_PRED_INPUTS = 34
_C_PRED_MATCH = 35
_LAST_COL = _C_PRED_MATCH

_ZONES: tuple[tuple[int, int], ...] = (
    (_C_MODEL_SHEET, _C_SAME_METHOD),         # A:L   — Model Registry
    (_C_UNIT_R_SQUARED, _C_UNIT_RMSE),        # N:P   — Unit-Space Statistics
    (_C_F_STATISTIC, _C_PRESS_R_SQUARED),     # R:AA  — Fit-Space Statistics
    (_C_PRED_POINT, _C_PRED_MATCH),           # AC:AI — Prediction Comparison
)

# The ungrouped gap columns (width 2) between the zones: derived from the zones,
# then asserted against the named constants, so a zone edit that closes or
# widens a gap fails at import rather than shipping two zones that Excel fuses
# into one collapse control.
_GAP_COLUMNS: tuple[int, ...] = tuple(_ZONES[i][1] + 1 for i in range(len(_ZONES) - 1))
assert _GAP_COLUMNS == (_GAP_1, _GAP_2, _GAP_3)
assert all(
    _ZONES[i + 1][0] - _ZONES[i][1] == 2 for i in range(len(_ZONES) - 1)
), "each zone must be separated from the next by exactly one gap column"

_COLUMN_GROUPS: tuple[tuple[int, int], ...] = _ZONES

# Content-column widths, keyed on the layout constant — never on a letter. The
# assertions below are the same contract the Regression sheet's table carries:
# every content column sized exactly once, no gap column sized at all, so a
# layout shift fails at import instead of shipping a column of the wrong width.
_COLUMN_WIDTHS: tuple[tuple[int, float], ...] = (
    (_C_MODEL_SHEET, 22.0),
    (_C_ANCHOR, 24.0),
    (_C_MODEL, 46.0),
    (_C_RESPONSE, 18.0),
    (_C_RESPONSE_DECLARED, 18.0),
    (_C_RESPONSE_SPACE, 20.0),
    (_C_BACK_TRANSFORM, 14.0),
    (_C_OBSERVATIONS, 8.0),
    (_C_DESIGN_COLUMNS, 8.0),
    (_C_SAME_SET, 10.0),
    (_C_SAME_SPACE, 11.0),
    (_C_SAME_METHOD, 12.0),
    (_C_UNIT_R_SQUARED, 12.0),
    (_C_UNIT_ADJUSTED_R_SQUARED, 12.0),
    (_C_UNIT_RMSE, 12.0),
    (_C_F_STATISTIC, 11.0),
    (_C_SIGNIFICANCE_F, 13.0),
    (_C_R_SQUARED, 10.0),
    (_C_ADJUSTED_R_SQUARED, 10.0),
    (_C_STANDARD_ERROR, 11.0),
    (_C_AIC, 11.0),
    (_C_BIC, 11.0),
    (_C_AICC, 11.0),
    (_C_PRESS, 12.0),
    (_C_PRESS_R_SQUARED, 12.0),
    (_C_PRED_POINT, 13.0),
    (_C_PRED_CI_LOWER, 11.0),
    (_C_PRED_CI_UPPER, 11.0),
    (_C_PRED_PI_LOWER, 11.0),
    (_C_PRED_PI_UPPER, 11.0),
    (_C_PRED_INPUTS, 34.0),
    (_C_PRED_MATCH, 11.0),
)

_CONTENT_COLUMNS: tuple[int, ...] = tuple(
    col for first, last in _ZONES for col in range(first, last + 1)
)
assert len(_COLUMN_WIDTHS) == len(_CONTENT_COLUMNS)
assert sorted(col for col, _ in _COLUMN_WIDTHS) == sorted(_CONTENT_COLUMNS), (
    "every zone content column must be sized exactly once"
)
assert not set(_GAP_COLUMNS) & {col for col, _ in _COLUMN_WIDTHS}, (
    "a gap column must stay unsized — its width is set with the other gaps"
)

# ── The reader position table ─────────────────────────────────────────────────
# Comparison_Field's documented index (1-24) → the sheet column it lands in. The
# index half lives in lambda_functions.json and is pinned there against the
# regression_layout constants (tests/test_comparison_offsets.py); this table is
# the other half, and the pair is what makes a row's cells a loop rather than
# twenty-four near-identical hand-written writes.
#
# The table is in INDEX order, not layout order: the columns are grouped by
# comparability, so index 24 (the declared response variable, the set test's key)
# displays in the registry zone while index 6 (F) displays in the fit-space zone.
_FIELD_COLUMNS: tuple[tuple[int, int], ...] = (
    (1, _C_RESPONSE),
    (2, _C_RESPONSE_SPACE),
    (3, _C_BACK_TRANSFORM),
    (4, _C_OBSERVATIONS),
    (5, _C_DESIGN_COLUMNS),
    (6, _C_F_STATISTIC),
    (7, _C_SIGNIFICANCE_F),
    (8, _C_UNIT_R_SQUARED),
    (9, _C_UNIT_ADJUSTED_R_SQUARED),
    (10, _C_UNIT_RMSE),
    (11, _C_R_SQUARED),
    (12, _C_ADJUSTED_R_SQUARED),
    (13, _C_STANDARD_ERROR),
    (14, _C_PRESS),
    (15, _C_PRESS_R_SQUARED),
    (16, _C_AIC),
    (17, _C_BIC),
    (18, _C_AICC),
    (19, _C_PRED_POINT),
    (20, _C_PRED_CI_LOWER),
    (21, _C_PRED_CI_UPPER),
    (22, _C_PRED_PI_LOWER),
    (23, _C_PRED_PI_UPPER),
    (24, _C_RESPONSE_DECLARED),
)
_FIELD_COUNT = len(_FIELD_COLUMNS)
assert [index for index, _ in _FIELD_COLUMNS] == list(range(1, _FIELD_COUNT + 1)), (
    "the position table must be dense and in order from 1 — CHOOSE's index is "
    "its position, so a gap or a swap silently reads the wrong statistic"
)
assert len({col for _, col in _FIELD_COLUMNS}) == _FIELD_COUNT, (
    "two fields cannot share a column — one would overwrite the other"
)
assert {col for _, col in _FIELD_COLUMNS} == set(_CONTENT_COLUMNS) - {
    _C_MODEL_SHEET,
    _C_ANCHOR,
    _C_MODEL,
    _C_SAME_SET,
    _C_SAME_SPACE,
    _C_SAME_METHOD,
    _C_PRED_INPUTS,
    _C_PRED_MATCH,
}, (
    "every content column is either a reader field, a derived cell, or a gate — "
    "a column in none of those is a column nothing writes"
)

# The indices the gate formulas read, and the columns they live in. The set test
# is the only one that reads TWO fields: the declared response variable (24) and
# n (4).
_INDEX_RESPONSE = 1
_INDEX_OBSERVATIONS = 4
_INDEX_DESIGN_COLUMNS = 5
_INDEX_RESPONSE_DECLARED = 24
assert _INDEX_RESPONSE_DECLARED == _FIELD_COUNT, (
    "field 24 is the last CHOOSE branch; if the table grew, the set test's key "
    "moved with it"
)

# ── The target-sheet read ─────────────────────────────────────────────────────
# The prediction-inputs band on the target sheet, as an OFFSET from the anchor.
# Derived from regression_layout's own constants — the band is a Regression-sheet
# fact and there is exactly one place it is written down.
_INPUTS_ROWS_OFFSET = _PRED_INPUT_FIRST_ROW - _ROW_RESPONSE_READOUT
_INPUTS_COLS_OFFSET = _C_AK - _C_AF
_INPUTS_HEIGHT = _PRED_INPUT_LAST_ROW - _PRED_INPUT_FIRST_ROW + 1
assert _INPUTS_HEIGHT > 0
assert _INPUTS_ROWS_OFFSET > 0, "the prediction band sits below the response readout"

# ── Headers ───────────────────────────────────────────────────────────────────
_HEADERS: tuple[tuple[int, str], ...] = (
    (_C_MODEL_SHEET, "Model Sheet"),
    (_C_ANCHOR, "Anchor"),
    (_C_MODEL, "Model"),
    (_C_RESPONSE, "Response"),
    (_C_RESPONSE_DECLARED, "Response Variable"),
    (_C_RESPONSE_SPACE, "Response Space"),
    (_C_BACK_TRANSFORM, "Back-Transform"),
    (_C_OBSERVATIONS, "n"),
    (_C_DESIGN_COLUMNS, "k"),
    (_C_SAME_SET, "Same Set?"),
    (_C_SAME_SPACE, "Same Space?"),
    (_C_SAME_METHOD, "Same Method?"),
    (_C_UNIT_R_SQUARED, "R² (unit)"),
    (_C_UNIT_ADJUSTED_R_SQUARED, "Adj R² (unit)"),
    (_C_UNIT_RMSE, "RMSE (unit)"),
    (_C_F_STATISTIC, "F"),
    (_C_SIGNIFICANCE_F, "Significance F"),
    (_C_R_SQUARED, "R²"),
    (_C_ADJUSTED_R_SQUARED, "Adj R²"),
    (_C_STANDARD_ERROR, "Std Error"),
    (_C_AIC, "AIC"),
    (_C_BIC, "BIC"),
    (_C_AICC, "AICc"),
    (_C_PRESS, "PRESS"),
    (_C_PRESS_R_SQUARED, "PRESS R²"),
    (_C_PRED_POINT, "Point Estimate"),
    (_C_PRED_CI_LOWER, "CI Lower"),
    (_C_PRED_CI_UPPER, "CI Upper"),
    (_C_PRED_PI_LOWER, "PI Lower"),
    (_C_PRED_PI_UPPER, "PI Upper"),
    (_C_PRED_INPUTS, "Prediction Inputs"),
    (_C_PRED_MATCH, "Same Inputs?"),
)
assert tuple(col for col, _ in _HEADERS) == _CONTENT_COLUMNS, (
    "the header list must cover every content column, in layout order"
)

# Number formats, keyed on the layout constant, drawn from the `_FMT_*` table in
# `regression_layout.py` — the formats the **Regression sheet** applies to the
# same cells.
#
# That sharing is the point, not tidiness. Every entry below except the two
# counts is a statistic MIRRORED off a target Regression sheet, and this sheet's
# whole claim is that these are that sheet's own numbers. Showing one of them at
# a different precision — 976.1166 beside 976.1 — invites the reader to see a
# difference that is not there, and the pair exists precisely so differences can
# be read. So each format travels with its statistic rather than being chosen
# per sheet.
#
# Two entries are deliberately not mirrored numbers:
#   * `_C_DESIGN_COLUMNS` (k) reads the spec block's Σ total, which the spec
#     block leaves General. It is an integer count, so `_FMT_COUNT` renders it
#     identically, and pinning a format keeps this column's own reading uniform.
#   * `_C_OBSERVATIONS` (n) reads AB9, which IS formatted — as `_FMT_COUNT`.
#
# `tests/test_sheet_comparison_sheet.py` drives BOTH writers and asserts they
# emit the same format for every mirrored statistic, so a drift on either side
# fails the suite instead of shipping two readings of one number.
_NUMBER_FORMATS: tuple[tuple[int, str], ...] = (
    (_C_OBSERVATIONS, _FMT_COUNT),
    (_C_DESIGN_COLUMNS, _FMT_COUNT),
    (_C_UNIT_R_SQUARED, _FMT_STAT),
    (_C_UNIT_ADJUSTED_R_SQUARED, _FMT_STAT),
    (_C_UNIT_RMSE, _FMT_STAT),
    (_C_F_STATISTIC, _FMT_F_STATISTIC),
    (_C_SIGNIFICANCE_F, _FMT_SIGNIFICANCE_F),
    (_C_R_SQUARED, _FMT_STAT),
    (_C_ADJUSTED_R_SQUARED, _FMT_STAT),
    (_C_STANDARD_ERROR, _FMT_STAT),
    (_C_AIC, _FMT_STAT),
    (_C_BIC, _FMT_STAT),
    (_C_AICC, _FMT_STAT),
    (_C_PRESS, _FMT_STAT),
    (_C_PRESS_R_SQUARED, _FMT_STAT),
    (_C_PRED_POINT, _FMT_STAT),
    (_C_PRED_CI_LOWER, _FMT_STAT),
    (_C_PRED_CI_UPPER, _FMT_STAT),
    (_C_PRED_PI_LOWER, _FMT_STAT),
    (_C_PRED_PI_UPPER, _FMT_STAT),
)
assert {col for col, _ in _NUMBER_FORMATS} <= set(_CONTENT_COLUMNS)

# ── The four gates ────────────────────────────────────────────────────────────
# Each statistic zone is readable only while the gates beneath it hold. The
# conditional formats below read these, and so do the row-2 verdicts.
_GATE_SET = "comparison set"
_GATE_SPACE = "response space"
_GATE_METHOD = "back-transform method"
_GATE_INPUTS = "prediction point"

# Sheet-scoped flag-band names → the flag column they cover and the gate they
# summarise. The bands are what the row-2 verdicts count over; they run past the
# shipped rows on purpose (see _BAND_LAST_ROW). Layout order, so the row-2
# verdicts march left to right in the same order as the zones they gate.
_FLAG_BANDS: tuple[tuple[str, int, str], ...] = (
    ("Comparison_Set_Flags", _C_SAME_SET, _GATE_SET),
    ("Comparison_Space_Flags", _C_SAME_SPACE, _GATE_SPACE),
    ("Comparison_Method_Flags", _C_SAME_METHOD, _GATE_METHOD),
    ("Comparison_Prediction_Flags", _C_PRED_MATCH, _GATE_INPUTS),
)


# ── Small address builders ────────────────────────────────────────────────────

def _flag_ref(col: int, row: int) -> str:
    """``$H4`` — an absolute column with a RELATIVE row.

    The relative row is what lets ONE conditional-format rule over a whole band
    adjust per row as Excel shifts the expression down the range, which is why
    this is deliberately not the absolute form below.
    """
    return f"${col_letter(col)}{row}"


def _abs_row_ref(col: int, row: int) -> str:
    return f"${col_letter(col)}${row}"


def _abs_range(first_col: int, last_col: int, first_row: int = _ROW_FIRST,
               last_row: int = _BAND_LAST_ROW) -> str:
    return f"{_abs_row_ref(first_col, first_row)}:{_abs_row_ref(last_col, last_row)}"


def _anchor_name(row: int) -> str:
    """The sheet-scoped name carrying this row's target reference."""
    return f"Comp_Anchor_{row - _ROW_FIRST + 1}"


# ── Sheet furniture ───────────────────────────────────────────────────────────

def _write_zone_labels(sheet: xw.Sheet) -> None:
    for (first_col, _last), label in zip(
        _ZONES,
        (
            "MODEL REGISTRY",
            "UNIT-SPACE STATISTICS",
            "FIT-SPACE STATISTICS",
            "PREDICTION COMPARISON",
        ),
    ):
        section_heading(sheet, _ROW_ZONE_LABEL, first_col, label)
    for first_col, comment in _ZONE_NOTES:
        _set_note(sheet, _ROW_ZONE_LABEL, first_col, comment)


def _write_headers(sheet: xw.Sheet) -> None:
    for col, label in _HEADERS:
        cell = sheet.range(rc(_ROW_HEADER, col))
        cell.value = label
        cell.api.Font.Bold = True
        cell.color = _SUBHDR
        cell.api.WrapText = True
    for first_col, last_col in _ZONES:
        border_box(sheet, _ROW_HEADER, first_col, _BAND_LAST_ROW, last_col)


def _write_status_cells(sheet: xw.Sheet) -> None:
    """Row 2: one verdict per gate, in the column that gate's flag lives in."""
    for name, col, gate in _FLAG_BANDS:
        f(sheet, _ROW_STATUS, col, f'=Comparison_Flag_Status({name},"{gate}")')
        cell = sheet.range(rc(_ROW_STATUS, col))
        cell.api.WrapText = True
        cell.api.Font.Bold = True
        _set_note(sheet, _ROW_STATUS, col, _STATUS_NOTES[gate], label=gate)


# ── Row writing ───────────────────────────────────────────────────────────────

def _write_registry_row(sheet: xw.Sheet, row: int) -> None:
    """One model row: every cell derived from its own ``Comp_Anchor_<n>`` name.

    Four cell families, each written once:

    * the two **derived registry** cells (sheet name, anchor address), which read
      the reference itself rather than its value;
    * the **reader** cells — one per ``Comparison_Field`` index plus the model
      string — each ``IFERROR``-blanked, so a row pointed at a non-model sheet
      is blank rather than wrong;
    * the two **derived text** cells (the target's prediction inputs, and the
      model string's sibling readouts);
    * the four **gate** cells, which compare this row against the first
      registered row and make no ``Comparison_Field`` call of their own.
    """
    name = _anchor_name(row)
    first = _ROW_FIRST

    def self_ref(col: int) -> str:
        return _abs_row_ref(col, row)

    def ref_ref(col: int) -> str:
        return _abs_row_ref(col, first)

    # The one guard. Comparison_Field runs its A1 probe before any offset, so
    # field 1 (the anchor's own cell) is #N/A exactly when the anchor does not
    # point at a Regression-shaped sheet. Only the sheet-name cell needs it
    # explicitly: every other reader is IFERROR-blanked anyway, and CELL would
    # happily return THIS sheet's name for a template anchor pointing here.
    readable = f"NOT(ISNA(Comparison_Field({name},{_INDEX_RESPONSE})))"

    # The target's own sheet name, for the label and the hyperlink. CELL is the
    # only way to ask a REFERENCE what sheet it lives on, and the name it returns
    # is a LABEL first: the whole point of `Comp_Anchor_<n>` is that the row
    # says which sheet it reads, so a repointed name is visible on the sheet
    # rather than only in the Name Manager. The hyperlink is built from that
    # name, NOT from the `[Book]Sheet!cell` path CELL hands back — nothing here
    # derives an address from CELL, and `nm` is empty on an unsaved workbook,
    # which the LET's `nm=""` branch handles rather than emitting a broken link.
    #
    # The quoting is the Excel-side mirror of `workbook_helpers.quoted_sheet_name`:
    # a sheet name may contain an apostrophe (Excel's own "Bob's Data"), and every
    # apostrophe inside a quoted sheet name must be doubled or the `'...'!`
    # reference terminates early and the link points at a sheet that does not
    # exist. Only the LINK LOCATION is escaped — the display argument shows the
    # real name, apostrophes and all.
    f(
        sheet,
        row,
        _C_MODEL_SHEET,
        "=LET("
        f"ok,{readable},"
        f'nm,IFERROR(TEXTAFTER(CELL("filename",{name}),"]"),""),'
        f'IF(ok,IF(nm="","",HYPERLINK("#\'"&SUBSTITUTE(nm,"\'","\'\'")&"\'!'
        f'{_abs_row_ref(_C_MODEL_SHEET, 1)}",nm)),"")'
        ")",
    )
    # The anchor's own address, read off the reference. Always shown — for a
    # mis-pointed anchor this is the cell the row is actually reading, which is
    # exactly what someone debugging the row needs to see. The inner fallback is
    # for a reference on THIS sheet: CELL returns `$A$1` with no `[book]Sheet!`
    # prefix, so there is no `]` to split on and TEXTAFTER would fail — and a
    # template row's anchor points here.
    f(
        sheet,
        row,
        _C_ANCHOR,
        f'=IFERROR(TEXTAFTER(CELL("address",{name}),"]"),'
        f'IFERROR(CELL("address",{name}),""))',
    )
    # The model string is the target's own assembled formula, not a re-derivation.
    f(sheet, row, _C_MODEL, f'=IFERROR(Model_Formula_String({name}),"")')
    for index, col in _FIELD_COLUMNS:
        f(sheet, row, col, f'=IFERROR(Comparison_Field({name},{index}),"")')

    # Gate 1 — same DECLARED response variable AND same observation count. The
    # declared name, not the fitted label: a Log fit and a level fit of one
    # response must compare equal here, or the unit-space zone they exist to
    # compare would be shaded. A row is in the comparison set only if both hold;
    # anything else and the row answers a different question, which invalidates
    # every statistic zone.
    obs = self_ref(_C_OBSERVATIONS)
    f(
        sheet,
        row,
        _C_SAME_SET,
        f'=IF(OR({obs}="",{ref_ref(_C_OBSERVATIONS)}=""),"",'
        f"AND({self_ref(_C_RESPONSE_DECLARED)}={ref_ref(_C_RESPONSE_DECLARED)},"
        f'{obs}={ref_ref(_C_OBSERVATIONS)}))',
    )

    # Gate 2 — same response space. Gates the fit-space zone only: the unit-space
    # columns are already in response units and stay readable across this change.
    # Evaluated for every registered row independently of the other gates: they
    # name different axes, and nesting one inside another would make a verdict
    # depend on another's subject.
    f(
        sheet,
        row,
        _C_SAME_SPACE,
        f'=IF({obs}="","",'
        f"{self_ref(_C_RESPONSE_SPACE)}={ref_ref(_C_RESPONSE_SPACE)})",
    )

    # Gate 3 — same back-transform method. Gates the unit-space zone and the
    # prediction zone: the method moves exactly the numbers expressed in original
    # units, while leaving the underlying fit identical, so two rows differing
    # only here differ by a display choice rather than by a model.
    f(
        sheet,
        row,
        _C_SAME_METHOD,
        f'=IF({obs}="","",'
        f"{self_ref(_C_BACK_TRANSFORM)}={ref_ref(_C_BACK_TRANSFORM)})",
    )

    # The target's own prediction inputs, as text, so two rows evaluated at
    # different points are visibly different rather than quietly incomparable.
    f(
        sheet,
        row,
        _C_PRED_INPUTS,
        f'=IF({obs}="","",TEXTJOIN(", ",TRUE,OFFSET({name},'
        f"{_INPUTS_ROWS_OFFSET},{_INPUTS_COLS_OFFSET},{_INPUTS_HEIGHT},1)))",
    )

    # Gate 4 — the same prediction point: same predictor count, same values. Two
    # models compared at different input vectors are being compared on their
    # inputs, not their fits.
    f(
        sheet,
        row,
        _C_PRED_MATCH,
        f'=IF({obs}="","",'
        f"AND({self_ref(_C_DESIGN_COLUMNS)}={ref_ref(_C_DESIGN_COLUMNS)},"
        f"{self_ref(_C_PRED_INPUTS)}={ref_ref(_C_PRED_INPUTS)}))",
    )


def _write_row_formats(sheet: xw.Sheet) -> None:
    for col, fmt in _NUMBER_FORMATS:
        letter = col_letter(col)
        sheet.range(f"{letter}{_ROW_FIRST}:{letter}{_BAND_LAST_ROW}").number_format = fmt
    # The flags and the two counts are booleans and small integers: centred
    # reads as a status column rather than as data.
    for col in (
        _C_OBSERVATIONS,
        _C_DESIGN_COLUMNS,
        _C_SAME_SET,
        _C_SAME_SPACE,
        _C_SAME_METHOD,
        _C_PRED_MATCH,
    ):
        sheet.range(rc(_ROW_FIRST, col), rc(_BAND_LAST_ROW, col)).api.HorizontalAlignment = -4108  # xlCenter


def _write_gate_formats(sheet: xw.Sheet) -> None:
    """One red rule per statistic zone, plus one on each flag column.

    Every expression begins by testing that the row is registered
    (``$J4<>""``). Without it an unregistered row — whose flags are all ``""`` —
    would satisfy ``<>TRUE`` and light up, and a blank template row would read as
    a failed comparison instead of an empty one.
    """
    set_flag = _flag_ref(_C_SAME_SET, _ROW_FIRST)
    space_flag = _flag_ref(_C_SAME_SPACE, _ROW_FIRST)
    method_flag = _flag_ref(_C_SAME_METHOD, _ROW_FIRST)
    inputs_flag = _flag_ref(_C_PRED_MATCH, _ROW_FIRST)
    registered = f'{set_flag}<>""'

    # Each statistic zone holds while the gates beneath it hold, and `_ZONES[1:]`
    # is the three statistic zones in layout order — these three expressions are
    # written in that same order, so the zone and its gate cannot drift apart.
    zone_rules: tuple[str, ...] = (
        f'=AND({registered},NOT(AND({set_flag},{method_flag})))',
        f'=AND({registered},NOT(AND({set_flag},{space_flag})))',
        f'=AND({registered},NOT(AND({set_flag},{method_flag},{inputs_flag})))',
    )
    for expression, (first_col, last_col) in zip(zone_rules, _ZONES[1:]):
        add_expression_format(
            sheet,
            _abs_range(first_col, last_col),
            expression,
            fill=_CF_RED_FILL,
            font_color=_CF_RED_TEXT,
        )

    # The flag columns flag themselves, so a FALSE is visible without reading
    # across to the zone it invalidates.
    for col in (_C_SAME_SET, _C_SAME_SPACE, _C_SAME_METHOD, _C_PRED_MATCH):
        flag = _flag_ref(col, _ROW_FIRST)
        add_expression_format(
            sheet,
            _abs_range(col, col),
            f'=AND({flag}<>"",{flag}=FALSE)',
            fill=_CF_RED_FILL,
            font_color=_CF_RED_TEXT,
        )


# ── Names ─────────────────────────────────────────────────────────────────────

_WIRING_NAME_COMMENTS = {
    "Comparison_Set_Flags": (
        "Row-2 verdict band: the Same Set? column, counted by "
        "Comparison_Flag_Status. Runs past the shipped rows so a row added later "
        "is still counted."
    ),
    "Comparison_Space_Flags": (
        "Row-2 verdict band: the Same Space? column, counted by "
        "Comparison_Flag_Status."
    ),
    "Comparison_Method_Flags": (
        "Row-2 verdict band: the Same Method? column, counted by "
        "Comparison_Flag_Status."
    ),
    "Comparison_Prediction_Flags": (
        "Row-2 verdict band: the Same Inputs? column, counted by "
        "Comparison_Flag_Status."
    ),
}


def _anchor_comment(target: str | None) -> str:
    """The Name Manager comment for one row's anchor name."""
    if target is None:
        return (
            "This row's model reference - UNREGISTERED. Point this name at "
            "another Analysis sheet's response-label cell (that sheet's own "
            "Comparison_Anchor, e.g. 'Regression'!$AF$3) and the row fills in. "
            "It points at this sheet's own A1 for now, so the row reads blank."
        )
    return (
        f"This row's model reference: {target}'s response-label cell. Editing "
        "this one name repoints the row at a different model sheet - every cell "
        "in the row follows."
    )


def _setup_local_names(
    sheet: xw.Sheet,
    closures: Sequence[CatalogFunction],
    registry: Sequence[str],
    row_last: int,
) -> None:
    """Install this sheet's names: wiring bands, then the row anchors.

    Wiring first — the row-2 status formulas read the flag bands — then each
    row's anchor, then the catalog closures, matching the order Excel has to
    resolve them in.
    """
    sname = quoted_sheet_name(SHEET_NAME)

    for name, col, _gate in _FLAG_BANDS:
        drop_local_name(sheet, name)
        _nm = sheet.api.Names.Add(Name=name, RefersTo=f"={sname}!{_abs_range(col, col)}")
        _nm.Comment = _WIRING_NAME_COMMENTS[name]

    # Row anchors. Every one is a REFERENCE — that is the whole mechanism — and
    # the unregistered ones deliberately point at a cell whose text is not
    # "MODEL SPECIFICATION", so their guard fails loudly into NA() instead of
    # reading a plausible-looking number from an unrelated cell.
    #
    # The reference is the target's response-label CELL, not that sheet's own
    # sheet-scoped Comparison_Anchor name — even though the name is the
    # documented interface and reads as the more robust choice. The reason is
    # mechanical: OFFSET must be handed a reference, and a reference that
    # arrives through a name in a cell is dereferenced before OFFSET sees it
    # (see the module docstring). Pointing at 'Sheet'!Comparison_Anchor would be
    # a SECOND name hop on top of the anchor name itself, and whether Excel
    # carries a reference through two hops is not something this build has
    # verified. The address costs nothing here: _A_RESPONSE_READOUT and that
    # name's own RefersTo are both derived from the same regression_layout
    # constants, so the two cannot drift.
    unregistered = f"={sname}!{_abs_row_ref(_C_MODEL_SHEET, 1)}"
    for row in range(_ROW_FIRST, row_last + 1):
        index = row - _ROW_FIRST
        target = registry[index] if index < len(registry) else None
        refers_to = (
            f"={quoted_sheet_name(target)}!{_A_RESPONSE_READOUT}"
            if target is not None
            else unregistered
        )
        name = _anchor_name(row)
        drop_local_name(sheet, name)
        _nm = sheet.api.Names.Add(Name=name, RefersTo=refers_to)
        _nm.Comment = _anchor_comment(target)

    for closure in closures:
        drop_local_name(sheet, closure.name)
        _nm = sheet.api.Names.Add(
            Name=closure.name, RefersTo="=" + closure.formula_display.lstrip("=")
        )
        _nm.Comment = closure.notes


# ── Notes ─────────────────────────────────────────────────────────────────────

_ZONE_NOTES: tuple[tuple[int, str], ...] = (
    (
        _C_UNIT_R_SQUARED,
        "Goodness of fit in the response's OWN units — the columns to read when "
        "one row is a Log fit and another is not. These are not invariant to a "
        "change of response transform (the fitted function itself changes, so "
        "the numbers move), but a difference between two rows here is a real "
        "difference in how each model predicts the actual response, which is "
        "what makes them readable across that change. They ARE moved by the "
        "back-transform method while the fit stays identical, so this zone is "
        "gated on the method as well as on the comparison set.",
    ),
    (
        _C_F_STATISTIC,
        "Statistics measured against the FITTED, possibly transformed response. "
        "Valid only when the rows share BOTH the comparison set AND the response "
        "space: R² here is computed against a different left-hand side once the "
        "transform differs, F tests a different hypothesis, and the log-Jacobian "
        "alone shifts AIC by tens of thousands. When a row is shaded, read the "
        "unit-space columns beside it instead — that is what they are for.",
    ),
    (
        _C_PRED_POINT,
        "Predictions in original units, valid when the rows share the comparison "
        "set AND the back-transform method AND were evaluated at the same "
        "prediction point. Two models predicted at different input vectors are "
        "being compared on their inputs, not their fits — see the Same Inputs? "
        "column.",
    ),
)

_STATUS_NOTES = {
    _GATE_SET: (
        "How many registered rows are in the same comparison set as the first "
        "one. A row is in the set when its Response Variable and its n both "
        "match the first registered row's — the same DECLARED response fitted "
        "to the same observations. FALSE in a row means that row answers a "
        "different question: none of its statistics belong beside the first "
        "row's, and every statistic zone on the row is shaded. Blank when no row "
        "is registered; blank rows are counted as neither."
    ),
    _GATE_SPACE: (
        "How many registered rows share the first one's response space. FALSE "
        "means the response was transformed differently — a Log fit beside an "
        "untransformed one — so the fit-space statistics (F, significance F, R², "
        "adjusted R², standard error, AIC, BIC, AICc, PRESS, PRESS R²) are not "
        "comparable: they are measured against a different left-hand side, and "
        "AIC carries the log-Jacobian as well. The unit-space columns beside "
        "them still are, which is what they are for. Compare those rows on the "
        "unit-space columns only, or refit one of them in the same space."
    ),
    _GATE_METHOD: (
        "How many registered rows back-transform the same way as the first one "
        "(Duan or Naive). The fit is unaffected by this choice — R², AIC and the "
        "rest are identical either way — but every number expressed in original "
        "units moves: the unit-space triplet, the point estimate and its bounds. "
        "FALSE means two rows are being compared on a display choice rather than "
        "on their models, so the unit-space and prediction zones are shaded. "
        "Set both sheets' Back-Transform cells to the same method, or read only "
        "the fit-space columns across these rows."
    ),
    _GATE_INPUTS: (
        "How many registered rows predict the same point as the first one. A row "
        "predicts the same point when it has the same number of predictors and "
        "the same values in its own Prediction Inputs band — read from the "
        "target sheet and shown here as text. FALSE means the two rows were "
        "evaluated at different input vectors, so their predictions answer "
        "different questions; comparing them compares the inputs, not the "
        "models. Bring both sheets' Prediction Inputs bands to the same point "
        "before reading this zone across rows."
    ),
}

_HEADER_NOTES: tuple[tuple[int, str], ...] = (
    (
        _C_MODEL_SHEET,
        "The sheet this row reads, taken from the anchor reference itself, with "
        "a link straight to it. A blank cell means the anchor name does not "
        "point at a Regression-shaped sheet.",
    ),
    (
        _C_ANCHOR,
        "The address this row reads: the target model sheet's response-label "
        "cell (that sheet's own Comparison_Anchor). The reference itself is the "
        "sheet-scoped name Comp_Anchor_<row> — edit THAT name in Formulas then "
        "Name Manager to point this row at a different model sheet, and every "
        "cell in the row follows. A name pointing at something that is not a "
        "Regression-shaped sheet leaves the row blank rather than reading an "
        "unrelated cell, because the readers probe for the MODEL SPECIFICATION "
        "heading at the target's A1 before reading anything.",
    ),
    (
        _C_MODEL,
        "The target sheet's own model formula string, read from the Model "
        "Formula cell on that sheet. The target computes it from its own "
        "specification; nothing is re-derived here.",
    ),
    (
        _C_RESPONSE,
        "The response as FITTED: a logged fit of MPG reads “Ln(MPG)”, a level "
        "fit reads “MPG”. This is what the model was actually fitted against, "
        "which is why it is the column that shows a transform at a glance. It is "
        "not the comparison key — see Response Variable.",
    ),
    (
        _C_RESPONSE_DECLARED,
        "The response variable as DECLARED in the target's spec block, before "
        "any transform: “MPG” whether that sheet fits MPG or Ln(MPG). This is "
        "the key the Same Set? test reads, because two rows fitting the same "
        "variable in different spaces are the pair the unit-space columns exist "
        "to compare, and comparing the fitted labels would call them different "
        "responses.",
    ),
    (
        _C_RESPONSE_SPACE,
        "The target sheet's Response Space readout: whether its fit is in the "
        "response's own units or in a transformed space. Two rows whose space "
        "differs cannot be compared on the fit-space statistics. This readout is "
        "why that distinction is available without parsing a transform column "
        "across sheets.",
    ),
    (
        _C_BACK_TRANSFORM,
        "The target sheet's back-transform method (Duan or Naive). A choice "
        "about how a fitted value is returned to original units, not about the "
        "fit: two rows that differ only here have identical coefficients and "
        "identical fit-space statistics, but different unit-space ones and "
        "different predictions. It reads “Duan” on sheets with no transform "
        "too, where it changes nothing — the Same Method? flag is then TRUE and "
        "correctly ignored.",
    ),
    (
        _C_OBSERVATIONS,
        "Observations in the target's fit, after its own Include, Filter and "
        "Log-drop masking. A different count means a different sample, and a "
        "different sample means a different problem.",
    ),
    (
        _C_DESIGN_COLUMNS,
        "Design-matrix columns in the target's fit, intercept included — the "
        "target sheet's own Design Columns total. An identity fact about the "
        "model rather than a statistic: no gate invalidates it, and it is shown "
        "here so a nested pair is obvious at a glance.",
    ),
    (
        _C_SAME_SET,
        "TRUE when this row's Response Variable and n both match the first "
        "registered row's — the same declared response fitted to the same "
        "observations. All three statistic zones are gated on it.",
    ),
    (
        _C_SAME_SPACE,
        "TRUE when this row's Response Space matches the first registered row's. "
        "Gates the fit-space zone only: the unit-space columns are already in "
        "response units and stay readable across a transform change.",
    ),
    (
        _C_SAME_METHOD,
        "TRUE when this row's back-transform method matches the first registered "
        "row's. Gates the unit-space zone and the prediction zone, both of which "
        "are expressed in original units and therefore move with the method, "
        "while the fit-space statistics do not.",
    ),
    (
        _C_UNIT_R_SQUARED,
        "Goodness of fit in response units — the column to read when one row is "
        "a Log fit and another is not.",
    ),
    (
        _C_F_STATISTIC,
        "The overall F statistic and, beside it, its p-value: both computed "
        "against the fitted response, so a Log fit's F tests a different "
        "hypothesis from a level fit's. Fit-space, like R².",
    ),
    (
        _C_R_SQUARED,
        "Fit-space goodness of fit: measured against the (possibly transformed) "
        "response the target fitted. Valid only within one comparison set and "
        "one response space — an AIC from a Log fit and one from a non-Log fit "
        "are not on the same scale.",
    ),
    (
        _C_PRED_POINT,
        "The target's prediction in original units, at its own Prediction Inputs "
        "band. Back-transformed there: Duan-smeared for the point estimate, "
        "Naive for the bounds. Comparable across a transform change, but only "
        "across the same method and at the same prediction point — see Same "
        "Inputs?.",
    ),
    (
        _C_PRED_INPUTS,
        "The values typed in the target sheet's own Prediction Inputs band, in "
        "order. Shown so it is obvious at a glance whether two rows were "
        "evaluated at the same point. This sheet has no inputs of its own: the "
        "prediction machinery is sheet-scoped and cannot be driven from here.",
    ),
    (
        _C_PRED_MATCH,
        "TRUE when this row predicts the same point as the first registered row: "
        "the same predictor count and the same input values. Gates the "
        "prediction zone — two models compared at different input vectors are "
        "being compared on their inputs, not their fits.",
    ),
)


def _set_note(
    sheet: xw.Sheet, row: int, col: int, text: str, *, label: str | None = None
) -> None:
    """Replace the cell's note with a sized box anchored to its right.

    The same helper pair the Regression sheet uses, with no size overrides — a
    note box that looks different on two sheets of one workbook reads as a
    different kind of thing.
    """
    cell = sheet.range(rc(row, col))
    cell_api = cell.api
    try:
        cell_api.ClearComments()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    cell_api.AddComment(text)
    try:
        cell_api.Comment.Visible = False
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    width, height = note_dimensions(label if label is not None else text, text, None)
    anchor_comment_right_of_cell(sheet, row, col, width, height)


def _header_label(col: int) -> str:
    for header_col, label in _HEADERS:
        if header_col == col:
            return label
    return col_letter(col)


def _annotate(sheet: xw.Sheet) -> None:
    for col, text in _HEADER_NOTES:
        _set_note(sheet, _ROW_HEADER, col, text, label=_header_label(col))


def _freeze_header_band(sheet: xw.Sheet) -> None:
    """Freeze the three-row header band at A4.

    The Regression sheet's own pattern, not ``safe_freeze_top_row`` (which
    freezes one row): clear any stale state first so re-freezing a rebuilt sheet
    is idempotent, then freeze at the SELECTION. Setting ``SplitRow`` instead
    would persist in the saved workbook as ``state="frozenSplit"`` — draggable
    split bars over a pinned header — rather than a true frozen pane, and the
    saved-XML pane state is how this is checked.
    """
    try:
        window = sheet.api.Application.ActiveWindow
        if window.FreezePanes:
            window.FreezePanes = False
        if window.Split:
            window.Split = False
        sheet.range(rc(1, 1)).select()
        sheet.range(f"A{_ROW_FIRST}").select()
        window.FreezePanes = True
    except Exception:  # pylint: disable=broad-exception-caught
        pass


# ── Entry point ───────────────────────────────────────────────────────────────

def write_comparison_sheet(
    workbook: xw.Book,
    closures: Sequence[CatalogFunction] = (),
    *,
    registry: Sequence[str] = (),
    template_rows: int = _TEMPLATE_ROWS,
) -> None:
    """Create or rebuild the Model Comparison sheet.

    Parameters
    ----------
    workbook : xw.Book
        The target workbook. Every sheet named in ``registry`` must already
        exist and be Regression-shaped — the readers probe for it.
    closures : Sequence[CatalogFunction]
        Sheet-scoped catalog entries owned by this sheet, in document order.
        Empty today (the three Model Comparison readers are workbook-scoped),
        but threaded through so a future sheet-scoped addition installs in the
        same place the Regression sheet installs its constructors.
    registry : Sequence[str]
        Sheet names to register, in row order. The FIRST one is the reference
        row every gate compares against. The production workbook passes one
        name; the test-model artifact passes its curated subset.
    template_rows : int
        Unregistered rows appended below the registry, ready to be pointed at
        other sheets by editing their ``Comp_Anchor_<n>`` names.
    """
    row_last = _ROW_FIRST + len(registry) + max(template_rows, 0) - 1

    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)
    reset_column_groups(sheet)
    safe_activate(sheet)

    # Names before cells: every row cell reads its own anchor name, and the
    # row-2 status cells read the flag bands.
    _setup_local_names(sheet, closures, registry, row_last)

    _write_zone_labels(sheet)
    _write_headers(sheet)
    _write_status_cells(sheet)
    for row in range(_ROW_FIRST, row_last + 1):
        _write_registry_row(sheet, row)
    _write_row_formats(sheet)
    _write_gate_formats(sheet)
    _annotate(sheet)

    for column, width in _COLUMN_WIDTHS:
        letter = col_letter(column)
        sheet.range(f"{letter}:{letter}").column_width = width
    for gap_col in _GAP_COLUMNS:
        letter = col_letter(gap_col)
        sheet.range(f"{letter}:{letter}").column_width = 2

    # One outline group per zone. The gap columns are deliberately left out of
    # every group: Excel fuses contiguous same-level grouped columns into a
    # single collapse control, so without a gap a zone could not be collapsed
    # independently of its neighbour.
    for first_col, last_col in _ZONES:
        sheet.api.Columns(f"{col_letter(first_col)}:{col_letter(last_col)}").Group()

    _freeze_header_band(sheet)

"""Every OFFSET in the two Model Comparison readers points where it claims.

**The failure this closes.** ``Model_Formula_String`` and ``Comparison_Field``
read another Regression sheet by offsetting from an anchor cell, and those
(row, column) offsets are **literal numbers inside a JSON string body** —
``lambda_functions.json`` is data, so no import can reach the numbers and no
``regression_layout`` constant is in scope where they were typed. That is
exactly the class ``AGENTS.md`` calls out for the two Log tokens: a value that
must agree with a Python constant, duplicated into a file the interpreter never
loads.

So a column insertion right of AF — the ordinary repair for a layout change —
moves a statistic, leaves the offset behind, and the reader silently returns
whatever now occupies the old cell. The formula still parses; the number is
wrong. ``docs/AGENTS`` calls that a silent-wrong-answer bug, and it is the one
kind of failure this repo treats as worse than a build error.

**What it asserts.** Every offset in the two bodies, recomputed from the
constants that name those cells:

* the **guard** — ``OFFSET(anchor, -2, -31)`` reaches ``A1``, which every
  Regression-shaped sheet holds as the literal "MODEL SPECIFICATION";
* the **23-entry position table** of ``Comparison_Field``, in order, against
  the ``_A_*`` anchor naming each statistic;
* the **two band offsets** of field 24 — the ``INDEX(...)``/``XMATCH(...)``
  pair that reads the declared response variable out of the target's spec
  block, sized from ``_SPEC_BAND_HEIGHT`` and positioned from the spec block's
  own first row and the anchor's;
* the **label pair** of ``Model_Formula_String`` against the Model Formula
  readout cell.

Every expectation is derived, never spelled: the test computes each offset as
``target − anchor`` from two ``_abs_ref`` strings, so moving a zone or a column
fails here rather than in a workbook nobody re-opened. The anchor itself is not
assumed — it is read from the same constant the sheet-scoped
``Comparison_Anchor`` name is built from, which ``tests/test_sheet_writers.py``
pins against ``AF3``.

Two offsets in a body are *ranged* (five arguments, not two), and the parser
below handles both shapes deliberately: it matches any argument list — then
asserts that every ``OFFSET(`` in the body was consumed, so a body that grew a
third shape fails here instead of having its new offsets silently ignored,
which is the blindness this whole file exists to prevent.

No Excel, no xlwings — JSON plus the imported layout constants, so it runs in
the Linux CI job beside the other committed-artifact checks.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from lambda_catalog.regression_layout import (
    _A_ADJUSTED_R_SQUARED,
    _A_AIC,
    _A_AICC,
    _A_BACK_TRANSFORM_METHOD,
    _A_BIC,
    _A_DESIGN_COLUMNS_TOTAL,
    _A_F_STATISTIC,
    _A_OBSERVATIONS,
    _A_PRED_CI_LOWER,
    _A_PRED_CI_UPPER,
    _A_PRED_PI_LOWER,
    _A_PRED_PI_UPPER,
    _A_PRED_POINT_ESTIMATE,
    _A_PRESS,
    _A_PRESS_R_SQUARED,
    _A_RESPONSE_READOUT,
    _A_RESPONSE_SPACE,
    _A_R_SQUARED,
    _A_SIGNIFICANCE_F,
    _A_STANDARD_ERROR,
    _A_UNIT_ADJUSTED_R_SQUARED,
    _A_UNIT_RMSE,
    _A_UNIT_R_SQUARED,
    _C_A,
    _C_AF,
    _C_MODEL_FORMULA,
    _C_SPEC_ROLE,
    _C_SPEC_VARIABLE,
    _ROW_MODEL_FORMULA,
    _ROW_RESPONSE_READOUT,
    _SPEC_BAND_HEIGHT,
    _abs_ref,
)
from lambda_catalog.spec_layout import _FIRST_DATA_ROW, _ROLE_RESPONSE
from lambda_catalog.workbook_helpers import col_letter

ROOT_DIR = Path(__file__).resolve().parents[1]
_CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

# ``OFFSET(anchor,0,0)`` and ``OFFSET(anchor,1,-31,11,1)`` — every argument list
# these two bodies use, in either shape. The coverage assertion in ``_offsets``
# is what keeps the shape-agnostic match honest: a body that grew an OFFSET this
# pattern cannot parse fails there rather than being silently skipped.
_OFFSET_RE = re.compile(r"OFFSET\(\s*anchor\s*((?:\s*,\s*-?\d+)*)\s*\)")
_INT_RE = re.compile(r"-?\d+")

# The anchor's own address: the response-name readout the sheet-scoped
# Comparison_Anchor name points at, and the value every offset below is
# measured from. Not spelled 'AF3' anywhere here.
_ANCHOR = _A_RESPONSE_READOUT

# The two literals the readers compare against and hand back. The guard probe
# reads A1, which write_sheet_regression writes as the zone heading; the label
# reads the Model Formula readout in the §4b materialization band.
_GUARD_ADDRESS = _abs_ref(1, _C_A)
_GUARD_LITERAL = "MODEL SPECIFICATION"
_FORMULA_LABEL_ADDRESS = _abs_ref(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA)

# The position table of Comparison_Field, in the order the description
# documents and the CHOOSE evaluates: index 1 first. This tuple is the whole
# claim of the test — each entry is the layout constant naming that cell, so
# the expectation is derived twice over (once here, once as an offset).
_FIELD_LAYOUT: tuple[str, ...] = (
    _A_RESPONSE_READOUT,             # 1  Response name (as FITTED — "Ln(MPG)")
    _A_RESPONSE_SPACE,               # 2  Response space
    _A_BACK_TRANSFORM_METHOD,        # 3  Back-transform method
    _A_OBSERVATIONS,                 # 4  Observations n
    _A_DESIGN_COLUMNS_TOTAL,         # 5  Design columns k
    _A_F_STATISTIC,                  # 6  F statistic
    _A_SIGNIFICANCE_F,               # 7  Significance F
    _A_UNIT_R_SQUARED,               # 8  Unit-space R²
    _A_UNIT_ADJUSTED_R_SQUARED,      # 9  Unit-space adjusted R²
    _A_UNIT_RMSE,                    # 10 Unit-space RMSE
    _A_R_SQUARED,                    # 11 R²
    _A_ADJUSTED_R_SQUARED,           # 12 Adjusted R²
    _A_STANDARD_ERROR,               # 13 Standard error
    _A_PRESS,                        # 14 PRESS
    _A_PRESS_R_SQUARED,              # 15 PRESS R²
    _A_AIC,                          # 16 AIC
    _A_BIC,                          # 17 BIC
    _A_AICC,                         # 18 AICc
    _A_PRED_POINT_ESTIMATE,          # 19 Point estimate (original units)
    _A_PRED_CI_LOWER,                # 20 CI lower
    _A_PRED_CI_UPPER,                # 21 CI upper
    _A_PRED_PI_LOWER,                # 22 PI lower
    _A_PRED_PI_UPPER,                # 23 PI upper
)

_FIELD_COUNT = len(_FIELD_LAYOUT)

# Field 24 is the only field that is not one cell on the target sheet: it reads
# the declared response variable out of the target's spec block, which means a
# RANGED offset for each of the two bands plus an XMATCH to find the Response
# row inside them. Both bands start at the spec block's first data row and are
# `_SPEC_BAND_HEIGHT` tall; the row offset is measured from the anchor.
_SPEC_ROW_OFFSET = _FIRST_DATA_ROW - _ROW_RESPONSE_READOUT
_SPEC_VARIABLE_OFFSET: tuple[int, ...] = (
    _SPEC_ROW_OFFSET, _C_SPEC_VARIABLE - _C_AF, _SPEC_BAND_HEIGHT, 1,
)
_SPEC_ROLE_OFFSET: tuple[int, ...] = (
    _SPEC_ROW_OFFSET, _C_SPEC_ROLE - _C_AF, _SPEC_BAND_HEIGHT, 1,
)
_RESPONSE_BAND_OFFSETS: tuple[tuple[int, ...], ...] = (
    _SPEC_VARIABLE_OFFSET,
    _SPEC_ROLE_OFFSET,
)

# The documented index range: the 23 cell fields, plus field 24.
_INDEX_RANGE = f"1-{_FIELD_COUNT + 1}"

# House floors: a detector that stopped matching would otherwise pass on an
# empty set, and a renamed entry would pass on an empty body list.
_MIN_FIELD_COUNT = 24


def _row_col(address: str) -> tuple[int, int]:
    """``$AF$3`` → ``(3, 32)``. The inverse of ``_abs_ref``, for this test only."""
    match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", address)
    assert match, f"not an A1 address: {address!r}"
    letters, row = match.group(1), int(match.group(2))
    col = 0
    for char in letters:
        col = col * 26 + (ord(char) - ord("A") + 1)
    return row, col


def _offset_to(target: str) -> tuple[int, int]:
    """The (row, column) offset a reader must use to reach ``target``."""
    anchor_row, anchor_col = _row_col(_ANCHOR)
    target_row, target_col = _row_col(target)
    return target_row - anchor_row, target_col - anchor_col


def _catalog_entries() -> dict[str, dict]:
    catalog = json.loads(_CATALOG_PATH.read_text(encoding="utf-8"))
    return {entry["name"]: entry for entry in catalog["functions"]}


def _offsets(entry: dict) -> list[tuple[int, ...]]:
    """Every OFFSET in a body, in source order, as its numeric arguments.

    The coverage assertion is the point: ``_OFFSET_RE`` only matches an OFFSET
    whose first argument is ``anchor``, so a future third shape would be
    invisible to it. Counting the literal ``OFFSET(`` occurrences against the
    matches turns that silence into a failure.
    """
    body = entry["formula_display"]
    matches = _OFFSET_RE.findall(body)
    assert body.count("OFFSET(") == len(matches), (
        f"{entry['name']} has {body.count('OFFSET(')} OFFSET calls but only "
        f"{len(matches)} offset from `anchor` — an offset this test cannot see "
        "is an offset it cannot pin"
    )
    return [tuple(int(v) for v in _INT_RE.findall(args)) for args in matches]


def test_the_address_parser_round_trips_through_col_letter() -> None:
    """A local parser is a second opinion — make sure it agrees with the first."""
    for col in (1, 26, 27, 32, 38, 52, 53, 80, 702, 703):
        assert _row_col(f"${col_letter(col)}$7") == (7, col)


def test_the_anchor_is_the_response_readout_the_sheet_name_points_at() -> None:
    """The offsets are relative to this cell, so the cell has to be the right one."""
    assert _ANCHOR == _abs_ref(_ROW_RESPONSE_READOUT, _C_AF)

def test_the_guard_probe_reaches_the_model_specification_heading() -> None:
    """``OFFSET(anchor,-2,-31)`` is A1 — the literal every Regression sheet holds.

    This is the whole reason a mis-pointed anchor is loud instead of silently
    wrong: the probe reads a fixed string that no other sheet shape carries.
    """
    assert _offset_to(_GUARD_ADDRESS) == (-2, -31)
    assert _GUARD_LITERAL == "MODEL SPECIFICATION"


def test_model_formula_string_offsets_match_the_layout() -> None:
    """Guard, then the Model Formula readout — one offset each."""
    offsets = _offsets(_catalog_entries()["Model_Formula_String"])
    assert offsets == [_offset_to(_GUARD_ADDRESS), _offset_to(_FORMULA_LABEL_ADDRESS)], (
        "Model_Formula_String's offsets no longer match the guard cell and the "
        f"Model Formula readout at {_FORMULA_LABEL_ADDRESS} — a layout shift "
        "moved one of them"
    )


def test_comparison_field_position_table_matches_the_layout() -> None:
    """All 23 cell statistics, in documented order, each at its layout constant.

    Reported as the first failing index rather than a bare list inequality:
    the useful sentence is "field 16 (AIC) moved", not two 23-tuples to diff
    by eye.
    """
    offsets = _offsets(_catalog_entries()["Comparison_Field"])
    assert offsets, "Comparison_Field's body has no OFFSET at all"
    assert offsets[0] == _offset_to(_GUARD_ADDRESS), (
        "Comparison_Field's first OFFSET must be the A1 guard probe"
    )
    fields = offsets[1 : 1 + _FIELD_COUNT]
    assert len(fields) == _FIELD_COUNT, (
        f"Comparison_Field has {len(offsets) - 1} offsets after the guard but "
        f"the position table lists {_FIELD_COUNT} cell fields (plus "
        f"{len(_RESPONSE_BAND_OFFSETS)} band offsets for field 24) — the body "
        "and this table disagree about the table's length"
    )
    for index, (address, found) in enumerate(zip(_FIELD_LAYOUT, fields), start=1):
        expected = _offset_to(address)
        assert found == expected, (
            f"Comparison_Field field {index} ({address}) is off: body says "
            f"OFFSET(anchor,{found[0]},{found[1]}), the layout says "
            f"OFFSET(anchor,{expected[0]},{expected[1]}) — a column or row "
            "moved and the JSON body did not follow"
        )


def test_field_24_reads_the_declared_response_variable_from_the_spec_block() -> None:
    """The set test's key, and the only field read as a RANGE.

    Field 1 is the response as FITTED — ``Ln(MPG)`` for a logged fit, ``MPG``
    for a level fit of the same data — so two rows of one comparison set would
    compare unequal on it. Field 24 reads the declared name instead, which
    means the target's spec block: a row offset to its first data row, a
    ``_SPEC_BAND_HEIGHT``-tall band in the Variable column, and an XMATCH over
    the matching Role band to find which row carries the Response role.

    The role literal is pinned to ``spec_layout._ROLE_RESPONSE`` across the
    Python/JSON gap, exactly as the two Log tokens are: the JSON is data, so no
    import can reach it, and a renamed token would leave the XMATCH matching
    nothing — which would return #N/A and blank every row's set flag without
    ever failing a build.
    """
    offsets = _offsets(_catalog_entries()["Comparison_Field"])
    assert offsets[1 + _FIELD_COUNT :] == list(_RESPONSE_BAND_OFFSETS), (
        "field 24's two band offsets no longer match the spec block's first "
        f"data row ({_FIRST_DATA_ROW}), its Variable/Role columns, and "
        f"_SPEC_BAND_HEIGHT ({_SPEC_BAND_HEIGHT})"
    )
    assert _SPEC_BAND_HEIGHT > 1, (
        "a one-row band cannot hold a Response row that is not the first one"
    )
    body = _catalog_entries()["Comparison_Field"]["formula_display"]
    role_band = f"OFFSET(anchor,{','.join(str(v) for v in _SPEC_ROLE_OFFSET)})"
    variable_band = f"OFFSET(anchor,{','.join(str(v) for v in _SPEC_VARIABLE_OFFSET)})"
    expected = f'INDEX({variable_band},XMATCH("{_ROLE_RESPONSE}",{role_band}))'
    assert expected in body, (
        "field 24 must INDEX the Variable band at the XMATCH position of the "
        f"Response role in the Role band; expected to find {expected!r} in the "
        "body"
    )
    assert body.count(f'"{_ROLE_RESPONSE}"') == 1, (
        "the Response role token appears in the body other than once — it is "
        "one lookup key, and a second spelling is a second thing to keep in step"
    )


def test_the_position_table_is_the_documented_length() -> None:
    """The floor, and the one number the description pins in prose."""
    assert _FIELD_COUNT + 1 >= _MIN_FIELD_COUNT
    entry = _catalog_entries()["Comparison_Field"]
    assert _INDEX_RANGE in entry["notes"], (
        "the notes must state the index range the table actually has: "
        f"{_INDEX_RANGE}"
    )
    for field in ("description", "yields"):
        assert _INDEX_RANGE in entry[field], (
            f"{field} must state the index range the table actually has: "
            f"{_INDEX_RANGE}"
        )
    args = {a["name"]: a["description"] for a in entry["arguments"]}
    assert _INDEX_RANGE in args["i"]


@pytest.mark.parametrize(
    "name", ["Model_Formula_String", "Comparison_Field"]
)
def test_both_readers_are_workbook_scoped_and_sheet_agnostic(name: str) -> None:
    """Workbook scope is only safe while the body reads nothing unqualified.

    ``Model_Formula`` is sheet-scoped for the opposite reason: its body reads
    ``Spec_Role``/``Header_Names``/``Source_Data`` and friends, which resolve
    against the calling sheet. A body that grew one such name would silently
    read the wrong sheet in the test-model workbook, where 50 sheets are
    Regression-shaped — the trap that already bit ``Fit_Sample_Include``.
    """
    entry = _catalog_entries()[name]
    # `.get`, not `[...]`: the JSON writes `scope` only when it is NOT workbook
    # (the 25 Regression-scoped entries), so reading the key directly would
    # KeyError on every workbook-scoped entry — including these two. The
    # default is the loader's own (catalog_schema.CatalogFunction.scope).
    assert entry.get("scope", "workbook") == "workbook"
    body = entry["formula_display"]
    for unqualified in (
        "Spec_Role",
        "Spec_Transform",
        "Spec_Include",
        "Header_Names",
        "Source_Data",
        "Allow_Intercept",
        "Constructed_Column_Names",
        "Constructed_Column_Transforms",
        "Predictor_Columns",
        "Fit_Context",
    ):
        assert unqualified not in body, (
            f"{name} is workbook-scoped but reads the sheet-scoped name "
            f"{unqualified} — an unqualified name resolves against the CALLING "
            "sheet, so every Regression-shaped sheet would read a different one"
        )

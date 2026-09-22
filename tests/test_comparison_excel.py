"""Workbook-backed oracle for the v3.4 Model Comparison sheet.

**What only Excel can answer.** `tests/test_comparison_offsets.py` recomputes
every `OFFSET` in the two readers from the layout constants, and
`tests/test_sheet_writers.py` pins the sheet's structure headless. Neither
evaluates a single formula, so neither can answer the one question the whole
design rests on: **does `OFFSET`, handed a cross-sheet reference through a
sheet-scoped name, actually read the OTHER sheet?** That is the mechanism the
anchor story is made of, and it was settled by a probe rather than by reasoning
— `OFFSET` preserves its reference argument's sheet context, and a cell merely
*containing* a cross-sheet reference is a value on this sheet, not a reference
to that one. If that were wrong, every row would read a plausible number off the
wrong sheet, which is the worst failure this repository recognises.

**Why it reads the artifact rather than building one.** The comparison sheet's
gates are meaningful only against real Regression-shaped sheets with real fits,
and building 50 of them inside a test would duplicate `build_test_models.py`
badly and slowly. So this test reads the artifact that script produces —
gitignored, and built on the machines that can build it — and skips with the
build command when it is absent. The gate cases it asserts are exactly the
curated registry `scripts/build_test_models.py` documents, so the two files name
the same six sheets and the same three gates on purpose.

Gated behind ``RUN_EXCEL_INTEGRATION=1``, like the other COM tests: the
GitHub-hosted runner has no Office, so this skips in CI and runs only where
desktop Excel exists.
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest
import xlwings as xw

from lambda_catalog.regression_layout import (
    _A_AIC,
    _A_DESIGN_COLUMNS_TOTAL,
    _A_OBSERVATIONS,
    _A_PRED_POINT_ESTIMATE,
    _A_RESPONSE_READOUT,
    _A_R_SQUARED,
    _A_UNIT_R_SQUARED,
    _C_MODEL_FORMULA,
    _ROW_MODEL_FORMULA,
    _abs_ref,
)
from lambda_catalog.write_sheet_comparison import (
    SHEET_NAME,
    _C_AIC,
    _C_DESIGN_COLUMNS,
    _C_MODEL,
    _C_MODEL_SHEET,
    _C_OBSERVATIONS,
    _C_PRED_MATCH,
    _C_PRED_POINT,
    _C_RESPONSE,
    _C_RESPONSE_DECLARED,
    _C_R_SQUARED,
    _C_SAME_METHOD,
    _C_SAME_SET,
    _C_SAME_SPACE,
    _C_UNIT_R_SQUARED,
    _FIELD_COLUMNS,
    _ROW_FIRST,
    _ROW_STATUS,
)
from tests.script_loader import load_script_module

build_test_models = load_script_module("build_test_models")

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="set RUN_EXCEL_INTEGRATION=1 to run Excel COM integration tests",
)

# The reference row: the first registered row every gate compares against.
_REFERENCE_ROW = _ROW_FIRST

# Which gate each curated row is in the registry to break, relative to row 1 —
# and the gates it must NOT break, which is the half that would catch a gate
# wired to the wrong column. Read off `_COMPARISON_REGISTRY`'s own comment.
_GATE_EXPECTATIONS: dict[str, dict[int, bool]] = {
    # Same declared response (mpg), same 392 rows, fitted logged: the fit-space
    # zone is unreadable and the unit-space zone beside it still is.
    "M05 Log-Log NA Masking": {_C_SAME_SET: True, _C_SAME_SPACE: False},
    # Same set, same space, one more design column: its prediction answers a
    # different point.
    "M14 Mixed Cat And Continuous": {
        _C_SAME_SET: True,
        _C_SAME_SPACE: True,
        _C_PRED_MATCH: False,
    },
    # A different sample (245 vs 392 rows): nothing on the row is comparable.
    "M15 Filter Degenerate Cat": {_C_SAME_SET: False},
    # The Duan/Naive pair on the Life Expectancy Log response — Set-F against
    # the M01 reference (different response, different sample), shipped as a
    # second comparison group.
    "L02 Log Response Duan": {_C_SAME_SET: False},
    "L03 Log Response Naive": {_C_SAME_SET: False},
}

# One statistic per zone, checked against the TARGET sheet's own cell rather
# than against another reader on this sheet: (field index, this sheet's column,
# the target address constant). The index is cross-checked against
# `_FIELD_COLUMNS` below, so the two tables cannot drift.
_SPOT_CHECKS: tuple[tuple[int, int, str], ...] = (
    (1, _C_RESPONSE, _A_RESPONSE_READOUT),
    (4, _C_OBSERVATIONS, _A_OBSERVATIONS),
    (5, _C_DESIGN_COLUMNS, _A_DESIGN_COLUMNS_TOTAL),
    (8, _C_UNIT_R_SQUARED, _A_UNIT_R_SQUARED),
    (11, _C_R_SQUARED, _A_R_SQUARED),
    (16, _C_AIC, _A_AIC),
    (19, _C_PRED_POINT, _A_PRED_POINT_ESTIMATE),
)

# The anchor name a row reads. Not imported from the writer: this is the
# user-facing half of the add-a-model story, and the test edits these names.
# The writer numbers them 1-based from the first data row (`_anchor_name` in
# write_sheet_comparison.py — row `_ROW_FIRST` reads `Comp_Anchor_1`), so the
# `+ 1` is the naming contract, not an offset. Kept as one local helper because
# three open-coded copies of this expression is how it drifted a row out.
def _anchor_name(row: int) -> str:
    return f"Comp_Anchor_{row - _ROW_FIRST + 1}"


# A row well past the anchor band, used ONLY to mint a name. The writer creates
# one anchor per band row, so every row of this registry already owns a live
# name (`Comp_Anchor_1` … for rows 4, 5, …); repointing one of THOSE would
# silently rewrite a registered row's target instead of testing a fresh name,
# and the corruption would look like a pass. Row 98 is past every row the sheet
# materializes, so `_anchor_name(98)` is free. The test writes nothing at row
# 98 — the number exists to name an anchor, not to address a cell.
_ROGUE_ROW = 98
_ROGUE_ANCHOR = _anchor_name(_ROGUE_ROW)


def _artifact_or_skip() -> Path:
    path = Path(build_test_models.DEFAULT_WORKBOOK_PATH)
    if not path.exists():
        pytest.skip(
            f"{path.name} has not been built. Run `python scripts/"
            "build_test_models.py` first — this test reads the materialized "
            "artifact rather than writing 50 Regression sheets itself."
        )
    return path


def _start_excel_or_skip() -> xw.App:
    try:
        app = xw.App(visible=False, add_book=False)
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False
        return app
    except Exception as exc:  # pragma: no cover - depends on the host's Excel
        pytest.skip(f"Excel COM is unavailable: {exc}")


def _flag(value) -> bool | None:
    """A gate cell as a tri-state: ``True``, ``False``, or ``None`` when blank.

    Blank is a state, not a missing value: an unregistered row's gates are
    deliberately empty so it is counted as neither in nor out of the verdict.
    """
    if value in (True, 1, 1.0, "TRUE", "True"):
        return True
    if value in (False, 0, 0.0, "FALSE", "False"):
        return False
    return None


def _same(a, b, *, places: int = 9) -> bool:
    """Numeric equality at Excel's own working precision, else exact equality.

    A statistic read off a live sheet and the same statistic computed at another
    formula's end can differ in the last bit or two; the sheet's own QC scale
    scores in decimal places for exactly this reason, so this compares at a
    tighter version of the same idea rather than demanding bit equality.
    """
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 10 ** (-places)
    return a == b


def _cell_state(sheet: xw.Sheet, row: int, column: int, into_row: int) -> str:
    """One cell's state, asked of Excel: ``"NA"``, ``"BLANK"`` or ``"PRESENT"``.

    **Never read an Excel error back through xlwings.** On this version an error
    cell and a blank one both arrive as ``None``, so ``cell.value`` cannot tell
    the loud case (the guard's ``#N/A``) from the deliberate one (an
    unregistered template row, blank by design) — which is the entire
    distinction the anchor tests exist to make. It cannot tell either from a
    real result, either. Excel states the state as text instead, which is the
    same probe ``test_categorical_model_construction_excel.py`` uses for ``#N/A``
    propagation, and for the same reason.

    ``"PRESENT"`` rather than "a number": the column probed most here is the
    Model Sheet column, which holds sheet NAMES, so the third state has to cover
    text — and a mis-pointed anchor reading a plausible value of EITHER kind is
    the failure being guarded against.

    ``into_row`` is a scratch row in the SAME column as the cell probed, so the
    probe can name its subject with ``_abs_ref`` rather than a spelled address.
    """
    source = _abs_ref(row, column)
    sheet.range((into_row, column)).api.Formula2 = (
        f'=IF(ISNA({source}),"NA",IF({source}="","BLANK","PRESENT"))'
    )
    return sheet.range((into_row, column)).value


def test_the_registry_rows_read_their_target_sheets(tmp_path: Path) -> None:
    """Every row reproduces its target sheet's own numbers, cell for cell.

    The core assertion this whole design exists to make possible: a sheet-scoped
    name whose `RefersTo` is a cross-sheet reference, handed to `OFFSET` by a
    workbook-scoped LAMBDA, reads THAT sheet. A mis-pointed or dereferenced
    anchor would return a plausible number from the wrong place, so the check is
    against the target's own cells — read directly here, not through the same
    reader — rather than against another cell on the comparison sheet.
    """
    artifact = _artifact_or_skip()
    registry = tuple(build_test_models._COMPARISON_REGISTRY)

    # The two tables have to agree before either is used: this test's spot
    # checks and the sheet's own field-to-column map.
    by_index = dict(_FIELD_COLUMNS)
    for index, column, _address in _SPOT_CHECKS:
        assert by_index[index] == column, (
            f"field {index} is column {by_index.get(index)} on the sheet but "
            f"this test reads column {column}"
        )

    app = _start_excel_or_skip()
    try:
        book = app.books.open(str(artifact))
        try:
            sheet = book.sheets[SHEET_NAME]
            book.app.api.CalculateFullRebuild()

            for offset, target in enumerate(registry):
                row = _ROW_FIRST + offset
                name = f"row {row} ({target})"

                assert sheet.range((row, _C_MODEL_SHEET)).value == target, (
                    f"{name}: the Model Sheet cell does not name the sheet this "
                    "row is registered to — the anchor name and the registry "
                    "have come apart"
                )

                # The model string is the target's OWN assembled formula, read
                # by the reader rather than re-derived here.
                target_book = book.sheets[target]
                assert sheet.range((row, _C_MODEL)).value == target_book.range(
                    _abs_ref(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA)
                ).value, (
                    f"{name}: Model_Formula_String does not reproduce the "
                    "target sheet's own Model Formula readout"
                )

                # The declared response variable — field 24, read as a RANGE
                # out of the target's spec block. The key the set gate uses.
                assert sheet.range((row, _C_RESPONSE_DECLARED)).value not in (
                    "",
                    None,
                ), (
                    f"{name}: the declared-response cell is empty — the field-24 "
                    "band/XMATCH read did not resolve, which blanks the set flag "
                    "without failing any build"
                )

                for _index, column, address in _SPOT_CHECKS:
                    expected = target_book.range(address).value
                    actual = sheet.range((row, column)).value
                    assert _same(actual, expected), (
                        f"{name}: column {column} reads {actual!r} but {target}!"
                        f"{address} holds {expected!r} — the cross-sheet offset "
                        "resolved to the wrong cell"
                    )
        finally:
            # Never saved: the test repoints names below, and the artifact is
            # meant to stay as built.
            book.close()
    finally:
        app.quit()


def test_the_shipped_registry_breaks_the_gates_it_claims(tmp_path: Path) -> None:
    """Row 1 is the reference; each curated row breaks the gate it is here for.

    Both halves matter. A gate that fires on everything is as useless as one
    that fires on nothing, so each expectation names the gates that must hold
    TRUE beside the one that must be FALSE — which is what would catch a gate
    wired to the wrong cell.
    """
    artifact = _artifact_or_skip()
    registry = tuple(build_test_models._COMPARISON_REGISTRY)

    app = _start_excel_or_skip()
    try:
        book = app.books.open(str(artifact))
        try:
            sheet = book.sheets[SHEET_NAME]
            book.app.api.CalculateFullRebuild()

            # The reference row is in every set by construction.
            for column in (_C_SAME_SET, _C_SAME_SPACE, _C_SAME_METHOD):
                assert _flag(sheet.range((_REFERENCE_ROW, column)).value) is True, (
                    f"the reference row's gate in column {column} is not TRUE"
                )

            # Every expectation must name a registered row. The loop below
            # `continue`s on an unknown target, so a sheet renamed in the
            # registry would silently drop that row's coverage — the failure
            # this assertion exists to make loud.
            assert set(_GATE_EXPECTATIONS) <= set(registry), (
                "gate expectations name rows absent from the registry: "
                f"{sorted(set(_GATE_EXPECTATIONS) - set(registry))}"
            )

            for offset, target in enumerate(registry):
                expected = _GATE_EXPECTATIONS.get(target)
                if expected is None:
                    continue
                row = _ROW_FIRST + offset
                for column, value in expected.items():
                    actual = _flag(sheet.range((row, column)).value)
                    assert actual is value, (
                        f"row {row} ({target}) column {column}: expected "
                        f"{value}, got {actual!r}"
                    )

            # Every gate column has a row-2 verdict, and a verdict that fired
            # names the axis it is about.
            for column, subject in (
                (_C_SAME_SET, "comparison set"),
                (_C_SAME_SPACE, "response space"),
                (_C_SAME_METHOD, "back-transform method"),
            ):
                text = sheet.range((_ROW_STATUS, column)).value
                assert text, f"the {subject} verdict on row 2 is empty"
                assert subject in text, (
                    f"the {subject} verdict names a different subject: {text!r}"
                )
        finally:
            book.close()
    finally:
        app.quit()


def test_a_mis_pointed_anchor_is_loud_and_a_template_row_is_blank() -> None:
    """Three anchor states, in one workbook: right, wrong, and unset.

    * a **template row** (no registered target) is blank, not wrong — its anchor
      points at this sheet's own header cell, whose text is not the guard
      literal, so the readers fail into `NA()`;
    * a **mis-pointed anchor** — a Regression-COLUMN reference on a sheet that is
      not Regression-shaped at all — returns `#N/A` from the reader, which is the
      guard doing its job: a wrong target can never read a plausible number.
      Read through Excel (see `_cell_state`), because xlwings hands back the
      same `None` for this error as for the blank template row above;
    * the same name, repointed at a **real** target, resolves.

    The last step is also the add-a-model story under test: repointing one
    sheet-scoped name in Name Manager is the whole of registering a model.
    """
    artifact = _artifact_or_skip()
    registry = tuple(build_test_models._COMPARISON_REGISTRY)
    template_row = _ROW_FIRST + len(registry)

    app = _start_excel_or_skip()
    try:
        book = app.books.open(str(artifact))
        try:
            sheet = book.sheets[SHEET_NAME]
            book.app.api.CalculateFullRebuild()

            # An unregistered template row: nothing on it, and no error.
            assert sheet.range((template_row, _C_MODEL_SHEET)).value in ("", None)
            assert sheet.range((template_row, _C_MODEL)).value in ("", None)
            assert _flag(sheet.range((template_row, _C_SAME_SET)).value) is None

            # A rogue anchor: correct column, sheet that holds no spec block.
            rogue_sheet = "Mileage Data"
            rogue_ref = f"={xlwings_quote(rogue_sheet)}!{_A_RESPONSE_READOUT}"
            _add_name(sheet, _ROGUE_ANCHOR, rogue_ref)
            probe_row = template_row + 1
            probe = sheet.range((probe_row, _C_MODEL_SHEET))
            probe.api.Formula2 = f"=Comparison_Field({_ROGUE_ANCHOR},1)"
            book.app.api.CalculateFullRebuild()
            # Read through Excel rather than through xlwings: `.value` returns
            # None for this error, which is also what it returns for the blank
            # template row above — so reading it directly cannot tell the loud
            # case from the deliberate one, and a number would slip past both.
            state = _cell_state(sheet, probe_row, _C_MODEL_SHEET, probe_row + 1)
            assert state == "NA", (
                "a mis-pointed anchor must read #N/A — that is the guard doing "
                "its job — never a plausible value, and never blank: a BLANK "
                "here would mean the wrong target read as nothing instead of as "
                f"an error. Got {state!r}."
            )

            # The same name, pointed at a real Regression sheet, resolves.
            # Asserted against the TARGET sheet's own cell, read directly —
            # field 1 is the response readout the name itself points at, so
            # comparing it to the sheet's NAME would be false by construction.
            real = registry[0]
            _add_name(
                sheet, _ROGUE_ANCHOR, f"={xlwings_quote(real)}!{_A_RESPONSE_READOUT}"
            )
            book.app.api.CalculateFullRebuild()
            target_readout = book.sheets[real].range(_A_RESPONSE_READOUT).value
            assert _same(probe.value, target_readout), (
                f"the repointed anchor should read {real}'s own "
                f"{_A_RESPONSE_READOUT} ({target_readout!r}), got {probe.value!r}"
            )
        finally:
            book.close()
    finally:
        app.quit()


def test_two_rows_one_method_apart_shade_the_method_axis_only() -> None:
    """The Method gate, isolated — the one case this artifact cannot ship.

    Every curated case that differs in back-transform method also differs in its
    response, so it fires the Set gate first and the method axis is never read
    clean. The pair that isolates it is `L02` / `L03` — the Duan and Naive fits
    of the same Log response on the same 2,768 rows — so this test repoints the
    REFERENCE row at `L02` and an unregistered template row at `L03`. Between
    those two rows the set and space gates must hold TRUE while the method gate
    goes FALSE, which is the whole claim: the fit is identical and only the
    numbers expressed in original units move.

    Repointing those two names is also the shipped add-a-model story, exercised
    end to end.
    """
    artifact = _artifact_or_skip()
    registry = tuple(build_test_models._COMPARISON_REGISTRY)
    template_row = _ROW_FIRST + len(registry)
    reference_anchor = _anchor_name(_REFERENCE_ROW)
    template_anchor = _anchor_name(template_row)

    app = _start_excel_or_skip()
    try:
        book = app.books.open(str(artifact))
        try:
            sheet = book.sheets[SHEET_NAME]
            assert "L02 Log Response Duan" in registry

            _add_name(
                sheet,
                reference_anchor,
                f"={xlwings_quote('L02 Log Response Duan')}!{_A_RESPONSE_READOUT}",
            )
            _add_name(
                sheet,
                template_anchor,
                f"={xlwings_quote('L03 Log Response Naive')}!{_A_RESPONSE_READOUT}",
            )
            book.app.api.CalculateFullRebuild()

            # The two rows are genuinely the same comparison: same declared
            # response, same rows, same response space.
            assert _flag(sheet.range((_REFERENCE_ROW, _C_SAME_SET)).value) is True
            assert _flag(sheet.range((template_row, _C_SAME_SET)).value) is True
            assert _flag(sheet.range((template_row, _C_SAME_SPACE)).value) is True

            # ... and differ in exactly one thing.
            assert (
                _flag(sheet.range((template_row, _C_SAME_METHOD)).value) is False
            ), "L02 and L03 differ in back-transform method — the gate did not fire"

            verdict = sheet.range((_ROW_STATUS, _C_SAME_METHOD)).value
            assert verdict and "differ in back-transform method" in verdict, (
                f"the method verdict did not report a difference: {verdict!r}"
            )
        finally:
            book.close()
    finally:
        app.quit()


def xlwings_quote(sheet_name: str) -> str:
    """A sheet name as Excel wants it inside a formula: quoted, doubled quotes."""
    return "'" + sheet_name.replace("'", "''") + "'"


def _add_name(sheet: xw.Sheet, name: str, refers_to: str) -> None:
    """Repoint (or create) a sheet-scoped name — the user's Name Manager edit."""
    try:
        sheet.api.Names(name).Delete()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    sheet.api.Names.Add(Name=name, RefersTo=refers_to)

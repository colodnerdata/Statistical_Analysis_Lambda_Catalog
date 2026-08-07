"""Tests for the cartesian Beta grid-search writer.

This module exists separately from ``test_sheet_writers.py`` (which is
dedicated to the standard writer) so the cartesian writer can grow its own
unit coverage without swelling the standard-writer suite.
"""
# pylint: disable=missing-function-docstring,protected-access
from typing import Any, cast

from tests.recording_sheet import RecordingSheet

from lambda_catalog.workbook_helpers import col_letter
from lambda_catalog.write_sheet_univariate_cartesian import (
    _BODY_COL_START,
    _C_GS_BETA,
    _GS_C_BEST,
    _GS_R_BODY,
    _GS_R_HDR,
    _GS_R_P1,
    _GS_R_P2,
    _TABLE_NAME_S1,
    _TABLE_NAME_S2,
    _nll_body_formula,
    _write_grid_stage_cartesian,
    _write_two_stage_grid_search_cartesian,
)


def _as_sheet(sheet: RecordingSheet) -> Any:
    return cast(Any, sheet)


def _named_range(sheet: RecordingSheet, name: str) -> str:
    """Return the RefersTo of a sheet-scoped defined name, or '' if absent."""
    for item in sheet.api.Names.items:
        if item.Name.split("!", 1)[-1] == name:
            return cast(str, item.RefersTo)
    return ""


def _col(n: int) -> int:
    return _C_GS_BETA + n


def _row(rr: int) -> int:
    return 1 + rr  # r0 == 1 for Stage 1


def test_nll_body_formula_is_row_major_three_columns() -> None:
    formula = _nll_body_formula(1, _C_GS_BETA, 20)

    assert formula.startswith("=LET(")
    assert "MAKEARRAY(400,3," in formula
    # Alpha is the slow axis — INT((r-1)/n)+1
    assert "INT((r-1)/20)+1" in formula
    # Beta is the fast axis — MOD(r-1,n)+1
    assert "MOD(r-1,20)+1" in formula
    # Three-column projection via IFS
    assert "IFS(c=1,alpha,c=2,beta,TRUE,nll)" in formula
    # Per-cell NLL matches the standard writer's expression; here the rescaled
    # z comes from the enclosing LET rather than from the grid spill.
    assert "NLL_Beta(z,alpha,beta)+COUNT(d)*LN(scale_)" in formula


def test_grid_stage_writes_headers_seqs_body_and_named_ranges() -> None:
    sheet = RecordingSheet(name="Univariate")

    _write_grid_stage_cartesian(
        _as_sheet(sheet),
        row_start=1,
        col_start=_C_GS_BETA,
        title="Beta Stage 1",
        body_name="UV_BETA_S1",
        p1_label="Alpha (α)",
        p2_label="Beta (β)",
        p1_min=0.2,
        p1_max=50.0,
        p2_min=0.2,
        p2_max=50.0,
        beta_grid_size=10,
        stage_num=1,
    )

    # Body anchor is at the first body row + first body col
    body_row = _row(_GS_R_BODY)         # 6
    body_col = _col(_BODY_COL_START)    # 91 (CM)
    body_formula = cast(str, sheet.cell(body_row, body_col).api.Formula2)
    assert "MAKEARRAY(100,3," in body_formula

    # Three named ranges — one per spilled column — fixed A1 range so the
    # aggregation functions (``MIN``, ``TOCOL``, ``Grid_Argument_Minimum``)
    # see the cells. (The ``#`` spill suffix makes those functions silently
    # treat the range as empty.)
    col_a = col_letter(body_col)
    col_b = col_letter(body_col + 1)
    col_c = col_letter(body_col + 2)
    body_last_row = body_row + 99   # n*n - 1 = 99 for n=10
    assert _named_range(sheet, "UV_BETA_S1_ALPHA") == f"='Univariate'!${col_a}${body_row}:${col_a}${body_last_row}"
    assert _named_range(sheet, "UV_BETA_S1_BETA") == f"='Univariate'!${col_b}${body_row}:${col_b}${body_last_row}"
    assert _named_range(sheet, "UV_BETA_S1_NLL") == f"='Univariate'!${col_c}${body_row}:${col_c}${body_last_row}"

    # Best alpha / best beta reference the named ranges and use INT/MOD math.
    # The body location is computed inline as
    # ``XMATCH(MIN(NLL), TOCOL(NLL))`` rather than via the
    # ``Grid_Argument_Minimum`` LAMBDA; the LAMBDA returns an ``HSTACK``
    # whose intermediate ``INDEX(...,1,2)`` element can be cached as a stale
    # scalar during the dynamic-array spill's materialization, while the
    # inline ``XMATCH`` chain keeps the dependency on ``MIN`` direct.
    p1_row = _row(_GS_R_P1)             # 3
    p2_row = _row(_GS_R_P2)             # 4
    best_alpha = cast(str, sheet.cell(p1_row, _col(_GS_C_BEST)).api.Formula2)
    best_beta = cast(str, sheet.cell(p2_row, _col(_GS_C_BEST)).api.Formula2)
    assert best_alpha.startswith("=INDEX(UV_BETA_S1_ALPHA,INT((")
    assert "INT((XMATCH(MIN(UV_BETA_S1_NLL),TOCOL(UV_BETA_S1_NLL))-1)/10)+1)" in best_alpha
    assert best_beta.startswith("=INDEX(UV_BETA_S1_BETA,MOD(")
    assert "MOD(XMATCH(MIN(UV_BETA_S1_NLL),TOCOL(UV_BETA_S1_NLL))-1,10)+1)" in best_beta

    # Three body headers at the body column row (above the data spill)
    hdr_row = _row(_GS_R_HDR)           # 5
    assert sheet.cell(hdr_row, _col(_BODY_COL_START) + 0).value == "Alpha"
    assert sheet.cell(hdr_row, _col(_BODY_COL_START) + 1).value == "Beta"
    assert sheet.cell(hdr_row, _col(_BODY_COL_START) + 2).value == "NLL"

    # No ListObject is registered — Excel rejects dynamic-array spills inside
    # a structured-table body, so the body is a flat 3-column spill region
    # rather than a real Excel table.
    assert sheet.api.ListObjects.items == []


def test_two_stage_driver_writes_named_ranges_for_both_stages() -> None:
    sheet = RecordingSheet(name="Univariate")

    _write_two_stage_grid_search_cartesian(
        _as_sheet(sheet),
        col_start=_C_GS_BETA,
        body_prefix="UV_BETA",
        p1_label="Alpha (α)",
        p2_label="Beta (β)",
        p1_min=0.2,
        p1_max=50.0,
        p2_min=0.2,
        p2_max=50.0,
        beta_grid_size=10,
    )

    # Each stage registers the three body named ranges
    assert _named_range(sheet, "UV_BETA_S1_ALPHA")
    assert _named_range(sheet, "UV_BETA_S1_BETA")
    assert _named_range(sheet, "UV_BETA_S1_NLL")
    assert _named_range(sheet, "UV_BETA_S2_ALPHA")
    assert _named_range(sheet, "UV_BETA_S2_BETA")
    assert _named_range(sheet, "UV_BETA_S2_NLL")

    # No ListObjects — see test_grid_stage_writes_headers_seqs_body_and_named_ranges
    assert sheet.api.ListObjects.items == []


def test_grid_stage_uses_n_squared_as_boundary_guard_upper_bound() -> None:
    sheet = RecordingSheet(name="Univariate")

    _write_grid_stage_cartesian(
        _as_sheet(sheet),
        row_start=1,
        col_start=_C_GS_BETA,
        title="Stage",
        body_name="UV_BETA_S1",
        p1_label="Alpha (α)",
        p2_label="Beta (β)",
        p1_min=0.2,
        p1_max=50.0,
        p2_min=0.2,
        p2_max=50.0,
        beta_grid_size=10,
        stage_num=1,
    )

    p1_row = _row(_GS_R_P1)
    best_alpha = sheet.cell(p1_row, _col(_GS_C_BEST)).api
    # The boundary-guard CF rule's Formula1 references the argmin location
    # (=INDEX(Grid_Argument_Minimum(...),1,2)) and asserts it is in
    # [1, n*n]. n=10 → n*n=100.
    cf_rules = best_alpha.FormatConditions.items
    assert any("=100" in rule.Formula1 for rule in cf_rules), (
        "Expected the boundary-guard CF to use n*n (=100) as the upper bound, "
        f"but found rules: {[rule.Formula1 for rule in cf_rules]}"
    )


def test_data_feed_and_local_names_match_standard_writer() -> None:
    """The cartesian writer mirrors the standard writer's two name registrations.

    The catalog LAMBDAs (``NLL_Beta``, ``Grid_Argument_Minimum``) need
    ``UV_Data`` and ``UV_Include`` to resolve when the workbook opens; the
    cartesian writer registers both via the same ``_setup_local_names`` /
    ``_write_data_feed`` pair the standard writer uses.
    """
    sheet = RecordingSheet(name="Univariate")

    from lambda_catalog.write_sheet_univariate_cartesian import (
        _setup_local_names,
        _write_data_feed,
    )

    _write_data_feed(_as_sheet(sheet))
    _setup_local_names(_as_sheet(sheet))

    assert _named_range(sheet, "UV_Data") == "='Univariate'!$A$4#"
    assert _named_range(sheet, "UV_Include") == "='Univariate'!$B$4#"


def test_table_name_constants_are_kept_for_discoverability() -> None:
    """``Stage_{1,2}_Beta_Param_Search`` are reserved names for any future
    registration that may not be a real ``ListObject`` (e.g. a workbook-
    scoped marker pointing at the NLL column). They are still imported here
    so renaming them elsewhere surfaces a NameError at import time.
    """
    assert _TABLE_NAME_S1 == "Stage_1_Beta_Param_Search"
    assert _TABLE_NAME_S2 == "Stage_2_Beta_Param_Search"
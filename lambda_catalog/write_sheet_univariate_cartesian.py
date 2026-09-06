"""Write the univariate sheet using a Cartesian product approach for the Beta distribution grid search.

This module provides an alternative to the standard Univariate sheet writer that
replaces the Beta distribution's two-dimensional grid search with a Cartesian
product approach: two vertical ``SEQUENCE`` spills (alpha, beta) sit
side by side, and the ``N×N`` body is a single ``MAKEARRAY`` spill to their
right that evaluates ``NLL_Beta`` at each ``(alpha_i, beta_j)``. Stage 2 mirrors
Stage 1 with bounds derived from the Stage-1 optimum (best ± step), exactly as
``_write_two_stage_grid_search`` does for the standard writer.

Only the Beta distribution zone is written — the rest of the Univariate sheet
(histograms, other distributions, charts) is omitted, since this module exists
to compare two implementations of the same Beta fit.
"""

from __future__ import annotations

import xlwings as xw

from lambda_catalog.sheet_styles import (
    HEADER_COLOR as _HEADER,
    INPUT_COLOR as _INPUT,
    SUBHDR_COLOR as _SUBHDR,
)
from lambda_catalog.workbook_helpers import (
    col_letter,
    f,
    rc,
    val,
)

_FMT_INT = "0"
_FMT_1DP = "0.0"
_FMT_SCI_1DP = "0.0E+00"

_UNIVARIATE_SHEET_NAME = "Univariate"

# Anchor at the same column the standard writer uses so the fitting-table
# Best-cell references (computed from ``_C_GS_BETA``) line up regardless of
# which writer built the sheet.
_C_GS_BETA = 88    # CJ — first column of the Beta fit zone (matches write_sheet_univariate)
_ROW_FIT_ZONE = 1  # first row of the Beta fit zone

# Stage row anchors — mirror the standard writer so the fitting-table row
# math (offset 2 and 3 from r0) keeps landing on the parameter table.
_GS_R_CONTROL_HDR = 1
_GS_R_P1 = 2
_GS_R_P2 = 3
_GS_R_HDR = 4
_GS_R_BODY = 5

# Column anchors within a stage.
_GS_C_MINNLL = 0   # Min NLL label/value
_GS_C_N_GRID = 1   # Rows/Columns label/value
_GS_C_PARAM = 3    # Parameter label
_GS_C_INPUT = 4    # Input (informational only, kept so the parameter table
                   # keeps the same shape as the standard writer's for
                   # visual comparison)
_GS_C_MIN = 5
_GS_C_MAX = 6
_GS_C_STEP = 7
_GS_C_BEST = 8

# Column positions of the two vertical SEQUENCE spills and the body.
_AXIS_ALPHA_COL = 0   # alpha SEQUENCE (vertical) — to the LEFT
_AXIS_BETA_COL = 1    # beta SEQUENCE (vertical) — to the LEFT
_BODY_AXIS_GAP = 2    # CL — gap column between the axis SEQUENCEs and the body
_BODY_COL_START = 3   # CM — alpha column of body ListObject (column offset)
_BODY_COL_END = _BODY_COL_START + 2   # CO — NLL column of body ListObject

_TABLE_NAME_S1 = "Stage_1_Beta_Param_Search"
_TABLE_NAME_S2 = "Stage_2_Beta_Param_Search"
_BLOCK_GAP_R = 1      # gap row between Stage 1 and Stage 2

# Minimal data-feed columns on the Univariate sheet — the standard writer
# exposes the source-data column and an include mask at A4# / B4# and registers
# sheet-scoped names ``UV_Data`` / ``UV_Include`` so the catalog LAMBDA closures
# (NLL_Beta, Grid_Argument_Minimum, …) can read them. The cartesian writer only
# needs those two names, so we mirror just enough of the standard layout to
# register them.
_DATA_COL = 1        # column A — spill of the source values
_INCLUDE_COL = 2     # column B — spill of the ISNUMBER include mask
_DATA_ROW_START = 4  # first data row (matches the standard writer's anchor)
_SOURCE_TABLE = "LifeExpectancyData"
_SOURCE_COLUMN = "Life expectancy"


def _subheader_row(sheet: xw.Sheet, row: int, c1: int, c2: int) -> None:
    """Apply the shared sub-header styling to a contiguous cell range."""
    sheet.range(rc(row, c1), rc(row, c2)).color = _SUBHDR
    sheet.range(rc(row, c1), rc(row, c2)).api.Font.Bold = True


def _setup_local_names(sheet: xw.Sheet) -> None:
    """Register ``UV_Data`` and ``UV_Include`` so the catalog LAMBDAs can read them.

    Mirrors the standard writer's ``_setup_local_names`` for the only two names
    the cartesian body formula needs (``UV_Data`` and ``UV_Include``). The data
    and include-mask spills live at ``A4#`` and ``B4#`` — written by
    ``_write_data_feed`` below — and the names point at those spill anchors.
    """
    sname = sheet.name
    f_col = col_letter(_DATA_COL)
    i_col = col_letter(_INCLUDE_COL)
    nm = sheet.api.Names.Add(
        Name="UV_Data",
        RefersTo=f"='{sname}'!${f_col}${_DATA_ROW_START}#",
    )
    nm.Comment = (
        "The raw data column spill, unfiltered — the completeness mask is "
        "UV_Include."
    )
    nm = sheet.api.Names.Add(
        Name="UV_Include",
        RefersTo=f"='{sname}'!${i_col}${_DATA_ROW_START}#",
    )
    nm.Comment = "Completeness filter mask spill: TRUE where the raw value is numeric."


def _write_data_feed(sheet: xw.Sheet) -> None:
    """Write the Life Expectancy data spill at A4 and the include mask at B4.

    These formulas are the same ones the standard writer uses at A4 and B4 to
    pull the Life Expectancy values from the ``LifeExpectancyData`` table and
    produce a local filter mask. The cartesian writer reuses them rather than
    re-inventing them, because ``UV_Data`` / ``UV_Include`` already bind to
    ``A4#`` / ``B4#`` in ``_setup_local_names``.
    """
    val(sheet, _DATA_ROW_START - 1, _DATA_COL, _SOURCE_COLUMN)
    f(
        sheet,
        _DATA_ROW_START,
        _DATA_COL,
        f'=IF({_SOURCE_TABLE}[{_SOURCE_COLUMN}]="","",{_SOURCE_TABLE}[{_SOURCE_COLUMN}])',
    )
    val(sheet, _DATA_ROW_START - 1, _INCLUDE_COL, "Include")
    f(
        sheet,
        _DATA_ROW_START,
        _INCLUDE_COL,
        f"=ISNUMBER({_SOURCE_TABLE}[{_SOURCE_COLUMN}])",
    )


def _stage_a1(row_start: int, col_start: int, dr: int, dc: int) -> str:
    """Absolute A1 ref offset from (row_start, col_start) by (dr, dc)."""
    return f"${col_letter(col_start + dc)}${row_start + dr}"


def _nll_body_formula(
    row_start: int,
    col_start: int,
    n: int,
) -> str:
    """Build the row-major body formula for one cartesian stage.

    Returns a single ``MAKEARRAY(n*n, 3, LAMBDA(r, c, …))`` spill anchored at
    ``(row_start + _GS_R_BODY, col_start + _BODY_COL_START)``. Each row emits
    ``(alpha_i, beta_j, NLL_ij)`` in row-major order — alpha is the slow axis
    (changes every ``n`` rows) and beta is the fast axis (cycles every ``n``
    rows). The outer ``LET`` rescales ``UV_Data`` to ``(0, 1)`` once per spill,
    matching the per-cell NLL expression the standard writer's body uses but
    lifted out of the per-cell body so rescaling is computed once, not once per
    grid point.
    """
    body_row_top = row_start + _GS_R_BODY
    body_row_bot = body_row_top + n - 1
    alpha_se = (
        f"${col_letter(col_start + _AXIS_ALPHA_COL)}${body_row_top}"
        f":${col_letter(col_start + _AXIS_ALPHA_COL)}${body_row_bot}"
    )
    beta_se = (
        f"${col_letter(col_start + _AXIS_BETA_COL)}${body_row_top}"
        f":${col_letter(col_start + _AXIS_BETA_COL)}${body_row_bot}"
    )

    n_rows = n * n
    return (
        "=LET("
        "d,FILTER(UV_Data,UV_Include),"
        "range_,MAX(d)-MIN(d),"
        "pad,range_*0.001,"
        "scale_,range_+2*pad,"
        "z,(d-MIN(d)+pad)/scale_,"
        f"MAKEARRAY({n_rows},3,LAMBDA(r,c,"
        "LET("
        f"ai,INT((r-1)/{n})+1,"
        f"bi,MOD(r-1,{n})+1,"
        f"alpha,INDEX({alpha_se},ai),"
        f"beta,INDEX({beta_se},bi),"
        "nll,NLL_Beta(z,alpha,beta)+COUNT(d)*LN(scale_),"
        "IFS(c=1,alpha,c=2,beta,TRUE,nll)"
        "))))"
    )


def _write_grid_stage_cartesian(
    sheet: xw.Sheet,
    row_start: int,
    col_start: int,
    title: str,
    body_name: str,
    p1_label: str,
    p2_label: str,
    p1_min,
    p1_max,
    p2_min,
    p2_max,
    beta_grid_size: int = 20,
    stage_num: int = 1,
) -> dict:
    """Write one grid-search stage for the Beta distribution using Cartesian product.

    Layout (offsets from ``row_start``/``col_start``):

    ============= =============================================================
    row 0         Stage title merged across all stage columns.
    row 1         ``Min NLL:``, ``Rows/Columns``, parameter-table headers
                  (``Parameter | Input | Min | Max | Step | Best``).
    row 2         Alpha parameter row.
    row 3         Beta parameter row.
    row 4         Header row — ``Alpha`` over the alpha axis, ``Beta`` over the
                  beta axis, then the ListObject headers (``Alpha | Beta |
                  NLL``) at ``_BODY_COL_START.._BODY_COL_END``.
    rows 5..5+N-1 Alpha SEQUENCE (col 0), beta SEQUENCE (col 1), body
                  ListObject rows 1..N (cols ``_BODY_COL_START.._BODY_COL_END``).
    rows 5..4+N*N Body ListObject rows 2..N² (the row-major cartesian product
                  of alpha × beta; the SEQUENCE spills stop at row 5+N-1 but
                  the body keeps going).
    ============= =============================================================

    Parameters mirror ``_write_grid_stage``: ``p1_min``/``p1_max`` and
    ``p2_min``/``p2_max`` are either numeric defaults (Stage 1) or Excel
    formula strings referencing Stage 1 (Stage 2). The returned dict holds the
    same keys the standard writer returns so the higher-level two-stage
    driver can pass best/step values between stages.
    """
    n = beta_grid_size
    n_rows = n * n
    r0, c0 = row_start, col_start
    body_col_start = c0 + _BODY_COL_START
    body_col_end = c0 + _BODY_COL_END
    stage_last_col = max(body_col_end, c0 + _GS_C_BEST)

    # ── Stage title ──────────────────────────────────────────────────────────
    val(sheet, r0, c0, title)
    sheet.range(rc(r0, c0), rc(r0, stage_last_col)).merge()
    title_range = sheet.range(rc(r0, c0), rc(r0, stage_last_col))
    title_range.color = _HEADER
    title_range.api.Font.Bold = True

    control_hdr_row = r0 + _GS_R_CONTROL_HDR
    p1_row = r0 + _GS_R_P1
    p2_row = r0 + _GS_R_P2

    # ── Fixed-area tables: Min NLL + Rows/Columns ────────────────────────────
    val(sheet, control_hdr_row, c0 + _GS_C_MINNLL, "Min NLL:")
    val(sheet, control_hdr_row, c0 + _GS_C_N_GRID, "Rows/Columns")
    _subheader_row(sheet, control_hdr_row, c0 + _GS_C_MINNLL, c0 + _GS_C_N_GRID)

    val(sheet, p1_row, c0 + _GS_C_N_GRID, n)
    sheet.range(rc(p1_row, c0 + _GS_C_N_GRID)).number_format = _FMT_INT
    n_grid_ref = _stage_a1(r0, c0, _GS_R_P1, _GS_C_N_GRID)

    # ── Parameter table ─────────────────────────────────────────────────────
    for dc, label in [
        (_GS_C_PARAM, "Parameter"),
        (_GS_C_INPUT, "Input"),
        (_GS_C_MIN, "Min"),
        (_GS_C_MAX, "Max"),
        (_GS_C_STEP, "Step Size"),
        (_GS_C_BEST, "Best"),
    ]:
        val(sheet, control_hdr_row, c0 + dc, label)
    _subheader_row(sheet, control_hdr_row, c0 + _GS_C_PARAM, c0 + _GS_C_BEST)

    val(sheet, p1_row, c0 + _GS_C_PARAM, p1_label)
    val(sheet, p2_row, c0 + _GS_C_PARAM, p2_label)
    # The visible ``Input`` cell is informational only.
    val(sheet, p1_row, c0 + _GS_C_INPUT, 1.0)
    val(sheet, p2_row, c0 + _GS_C_INPUT, 1.0)

    for row, p_min, p_max in [
        (p1_row, p1_min, p1_max),
        (p2_row, p2_min, p2_max),
    ]:
        if isinstance(p_min, str):
            f(sheet, row, c0 + _GS_C_MIN, p_min)
        else:
            val(sheet, row, c0 + _GS_C_MIN, p_min)
        if isinstance(p_max, str):
            f(sheet, row, c0 + _GS_C_MAX, p_max)
        else:
            val(sheet, row, c0 + _GS_C_MAX, p_max)

        if stage_num == 1:
            sheet.range(rc(row, c0 + _GS_C_MIN)).color = _INPUT
            sheet.range(rc(row, c0 + _GS_C_MAX)).color = _INPUT

    min_p1_ref = _stage_a1(r0, c0, _GS_R_P1, _GS_C_MIN)
    max_p1_ref = _stage_a1(r0, c0, _GS_R_P1, _GS_C_MAX)
    min_p2_ref = _stage_a1(r0, c0, _GS_R_P2, _GS_C_MIN)
    max_p2_ref = _stage_a1(r0, c0, _GS_R_P2, _GS_C_MAX)
    step_p1_ref = _stage_a1(r0, c0, _GS_R_P1, _GS_C_STEP)
    step_p2_ref = _stage_a1(r0, c0, _GS_R_P2, _GS_C_STEP)

    f(sheet, p1_row, c0 + _GS_C_STEP,
      f"=({max_p1_ref}-{min_p1_ref})/({n_grid_ref}-1)")
    f(sheet, p2_row, c0 + _GS_C_STEP,
      f"=({max_p2_ref}-{min_p2_ref})/({n_grid_ref}-1)")

    # ── Header row (above the SEQUENCEs and body) ───────────────────────────
    hdr_row = r0 + _GS_R_HDR
    val(sheet, hdr_row, c0 + _AXIS_ALPHA_COL, "Alpha")
    val(sheet, hdr_row, c0 + _AXIS_BETA_COL, "Beta")
    sheet.range(rc(hdr_row, c0 + _AXIS_ALPHA_COL), rc(hdr_row, c0 + _AXIS_BETA_COL)).api.Font.Bold = True
    sheet.range(rc(hdr_row, c0 + _AXIS_ALPHA_COL), rc(hdr_row, c0 + _AXIS_BETA_COL)).color = _SUBHDR
    # Body ListObject headers — three cells (Alpha, Beta, NLL) at columns
    # _BODY_COL_START.._BODY_COL_END. The ListObject registration later in this
    # function names these headers as the table's column-header row, so the
    # data spills one row below them.
    val(sheet, hdr_row, body_col_start + 0, "Alpha")
    val(sheet, hdr_row, body_col_start + 1, "Beta")
    val(sheet, hdr_row, body_col_start + 2, "NLL")
    _subheader_row(sheet, hdr_row, body_col_start, body_col_end)

    # ── Alpha SEQUENCE (vertical, column-axis of the body) ───────────────────
    body_row_start = r0 + _GS_R_BODY
    f(
        sheet,
        body_row_start,
        c0 + _AXIS_ALPHA_COL,
        f"=SEQUENCE({n_grid_ref},1,{min_p1_ref},{step_p1_ref})",
    )
    sheet.range(
        rc(body_row_start, c0 + _AXIS_ALPHA_COL),
        rc(body_row_start + n - 1, c0 + _AXIS_ALPHA_COL),
    ).number_format = _FMT_1DP

    # ── Beta SEQUENCE (vertical, column-axis of the body) ────────────────────
    f(
        sheet,
        body_row_start,
        c0 + _AXIS_BETA_COL,
        f"=SEQUENCE({n_grid_ref},1,{min_p2_ref},{step_p2_ref})",
    )
    sheet.range(
        rc(body_row_start, c0 + _AXIS_BETA_COL),
        rc(body_row_start + n - 1, c0 + _AXIS_BETA_COL),
    ).number_format = _FMT_1DP

    # ── Body (single MAKEARRAY spill, N²×3, anchored at col _BODY_COL_START)
    # The LAMBDA inside MAKEARRAY returns (alpha_i, beta_j, NLL_ij) tuples in
    # row-major order — alpha is the slow axis (changes every n rows) and
    # beta is the fast axis (cycles every n rows). For n=20 the spill is 400
    # rows × 3 columns, anchored at ``$CM$5`` and reaching ``$CO$404``.
    f(
        sheet,
        body_row_start,
        body_col_start,
        _nll_body_formula(r0, c0, n),
    )
    alpha_body_range = sheet.range(
        rc(body_row_start, body_col_start + 0),
        rc(body_row_start + n_rows - 1, body_col_start + 0),
    )
    beta_body_range = sheet.range(
        rc(body_row_start, body_col_start + 1),
        rc(body_row_start + n_rows - 1, body_col_start + 1),
    )
    nll_body_range = sheet.range(
        rc(body_row_start, body_col_start + 2),
        rc(body_row_start + n_rows - 1, body_col_start + 2),
    )
    alpha_body_range.number_format = _FMT_1DP
    beta_body_range.number_format = _FMT_1DP
    nll_body_range.number_format = _FMT_SCI_1DP

    # ── Named ranges for the body (one per spilled column). The body is a
    # fixed ``n*n`` × 1 spill so the named ranges use an explicit A1 range,
    # not the ``#`` spill suffix — Excel's aggregation functions (``MIN``,
    # ``COUNT``, ``TOCOL``, and the catalog ``Grid_Argument_Minimum`` which
    # uses ``TOCOL`` internally) silently treat ``#`` references as empty,
    # and the ``Grid_Argument_Minimum(UV_BETA_S1_NLL)`` call would return
    # ``NA()`` for every cell.
    sname = sheet.name
    col_a = col_letter(body_col_start + 0)
    col_b = col_letter(body_col_start + 1)
    col_c = col_letter(body_col_start + 2)
    body_last_row = body_row_start + n_rows - 1
    body_alpha_a1 = f"${col_a}${body_row_start}:${col_a}${body_last_row}"
    body_beta_a1 = f"${col_b}${body_row_start}:${col_b}${body_last_row}"
    body_nll_a1 = f"${col_c}${body_row_start}:${col_c}${body_last_row}"
    name_alpha = f"{body_name}_ALPHA"
    name_beta = f"{body_name}_BETA"
    name_nll = f"{body_name}_NLL"
    nm = sheet.api.Names.Add(Name=name_alpha, RefersTo=f"='{sname}'!{body_alpha_a1}")
    nm.Comment = f"{body_name}: the Alpha candidate values (fixed A1 range over the spill)"
    nm = sheet.api.Names.Add(Name=name_beta, RefersTo=f"='{sname}'!{body_beta_a1}")
    nm.Comment = f"{body_name}: the Beta candidate values (fixed A1 range over the spill)"
    nm = sheet.api.Names.Add(Name=name_nll, RefersTo=f"='{sname}'!{body_nll_a1}")
    nm.Comment = f"{body_name}: the NLL body — Min NLL and Best α/β recovery read it"

    # ── Min NLL, Best alpha, Best beta (read from the NLL column name) ─────
    # The body is row-major, so the 1-based body-row index of the min NLL is
    # ``XMATCH(MIN(NLL), TOCOL(NLL))``. We inline that instead of going through
    # ``Grid_Argument_Minimum`` because that LAMBDA returns an ``HSTACK`` —
    # a 1×3 array whose ``INDEX(...,1,2)`` element can be cached as a stale
    # scalar at build-time when the body's dynamic spill is still materializing.
    # The inline ``XMATCH`` chain avoids the HSTACK intermediate and keeps the
    # dependency on ``MIN`` direct, which Excel's dependency graph updates
    # atomically with the spill.
    f(
        sheet,
        p1_row,
        c0 + _GS_C_MINNLL,
        f'=IFERROR(MIN({name_nll}),"—")',
    )
    sheet.range(rc(p1_row, c0 + _GS_C_MINNLL)).number_format = _FMT_SCI_1DP

    flat_loc = f"XMATCH(MIN({name_nll}),TOCOL({name_nll}))"
    f(
        sheet,
        p1_row,
        c0 + _GS_C_BEST,
        f"=INDEX({name_alpha},INT(({flat_loc}-1)/{n})+1)",
    )
    f(
        sheet,
        p2_row,
        c0 + _GS_C_BEST,
        f"=INDEX({name_beta},MOD({flat_loc}-1,{n})+1)",
    )
    sheet.range(rc(p1_row, c0 + _GS_C_BEST), rc(p2_row, c0 + _GS_C_BEST)).number_format = _FMT_1DP

    # ── Fixed-area table: align number formats ──────────────────────────────
    fixed_values = sheet.range(
        rc(p1_row, c0 + _GS_C_MIN),
        rc(p2_row, c0 + _GS_C_BEST),
    )
    fixed_values.number_format = _FMT_1DP

    # ── Boundary guard CF (red fill when the optimum lies on a grid edge) ───
    # With the row-major body, the upper bound is the row count (n²), not n.
    for row in (p1_row, p2_row):
        cell_api = sheet.range(rc(row, c0 + _GS_C_BEST)).api
        cf = cell_api.FormatConditions.Add(
            Type=2,  # xlExpression
            Formula1=f"=OR({flat_loc}=1,{flat_loc}={n_rows})",
        )
        cf.Interior.Color = 0x0000FF   # red (BGR)
        cf.Font.Color = 0xFFFFFF       # white

    # ── Body heatmap on the NLL column only (green=low → red=high/overflow)
    try:
        nll_body_range.api.FormatConditions.Delete()
        cs = nll_body_range.api.FormatConditions.AddColorScale(3)
        cs.ColorScaleCriteria(1).Type = 1
        cs.ColorScaleCriteria(1).FormatColor.Color = 0x63BE7B
        cs.ColorScaleCriteria(2).Type = 5
        cs.ColorScaleCriteria(2).Value = 50
        cs.ColorScaleCriteria(2).FormatColor.Color = 0xFFEB84
        cs.ColorScaleCriteria(3).Type = 2
        cs.ColorScaleCriteria(3).FormatColor.Color = 0xF8696B
    except Exception:
        pass

    # ── Note on ListObject registration ───────────────────────────────────────
    # The original plan called for registering a real Excel ``ListObject``
    # (``Stage_{1,2}_Beta_Param_Search``) so the body shows up in Power
    # Query, filter dropdowns, etc. Excel rejects that combination: the
    # body is a single ``MAKEARRAY`` dynamic-array spill, and a
    # ``ListObject`` body must contain static cells — registering the
    # table around the spill collapses it to a one-cell scalar and
    # orphans the rest of the formula. The named ranges registered above
    # (``UV_BETA_S{1,2}_ALPHA`` / ``_BETA`` / ``_NLL``) still give external
    # consumers a way to discover and read each column by name; only the
    # Excel table / filter UI is missing.

    return {
        "best_p1": _stage_a1(r0, c0, _GS_R_P1, _GS_C_BEST),
        "best_p2": _stage_a1(r0, c0, _GS_R_P2, _GS_C_BEST),
        "min_p1": min_p1_ref,
        "max_p1": max_p1_ref,
        "min_p2": min_p2_ref,
        "max_p2": max_p2_ref,
        "step_p1": step_p1_ref,
        "step_p2": step_p2_ref,
        "n_grid": n_grid_ref,
        "corner": _stage_a1(r0, c0, _GS_R_BODY, _BODY_COL_START),
    }


def _write_two_stage_grid_search_cartesian(
    sheet: xw.Sheet,
    col_start: int,
    body_prefix: str,
    p1_label: str,
    p2_label: str,
    p1_min: float,
    p1_max: float,
    p2_min: float,
    p2_max: float,
    beta_grid_size: int = 20,
) -> None:
    """Write Stage 1 and Stage 2 grid-search blocks for Beta using Cartesian product.

    Stage 2 uses the same data-driven bounds the standard writer uses:
    ``[best - step, best + step]`` for each parameter, floored at 0.001.
    """
    r0 = _ROW_FIT_ZONE
    # The cartesian body has ``beta_grid_size ** 2`` rows (one per (alpha, beta)
    # pair) under the header row at ``_GS_R_BODY = 4``. Stage 2 starts one gap
    # row past the end of Stage 1's body so the two ListObjects don't collide.
    stage2_row_start = r0 + _GS_R_BODY + (beta_grid_size * beta_grid_size) + _BLOCK_GAP_R

    s1 = _write_grid_stage_cartesian(
        sheet,
        row_start=r0,
        col_start=col_start,
        title=f"Beta Grid-Search MLE  —  Stage 1  ({p1_label} × {p2_label})",
        body_name=f"{body_prefix}_S1",
        p1_label=p1_label,
        p2_label=p2_label,
        p1_min=p1_min,
        p1_max=p1_max,
        p2_min=p2_min,
        p2_max=p2_max,
        beta_grid_size=beta_grid_size,
        stage_num=1,
    )

    _write_grid_stage_cartesian(
        sheet,
        row_start=stage2_row_start,
        col_start=col_start,
        title=f"Beta Grid-Search MLE  —  Stage 2  (refined)",
        body_name=f"{body_prefix}_S2",
        p1_label=p1_label,
        p2_label=p2_label,
        p1_min=f"=MAX(0.001,{s1['best_p1']}-{s1['step_p1']})",
        p1_max=f"={s1['best_p1']}+{s1['step_p1']}",
        p2_min=f"=MAX(0.001,{s1['best_p2']}-{s1['step_p2']})",
        p2_max=f"={s1['best_p2']}+{s1['step_p2']}",
        beta_grid_size=beta_grid_size,
        stage_num=2,
    )


def write_univariate_sheet_cartesian(
    workbook: xw.Book,
    sheet_notes: dict[str, str] | None = None,
    beta_grid_size: int = 20,
) -> xw.Sheet:
    """Create or replace the Univariate sheet and write Beta-only content using Cartesian product.

    Parameters
    ----------
    workbook : xw.Book
        Open xlwings workbook to write into.
    sheet_notes : dict[str, str] | None
        Plain-language tooltips for column headers and zone labels.
    beta_grid_size : int, optional
        Size of the Beta distribution grid search in each dimension (default: 20).
        The total number of NLL evaluations per stage is ``beta_grid_size ** 2``.

    Returns
    -------
    xw.Sheet
        The written sheet.
    """
    del sheet_notes  # no annotations in this minimal Beta-only writer

    if _UNIVARIATE_SHEET_NAME in {s.name for s in workbook.sheets}:
        workbook.sheets[_UNIVARIATE_SHEET_NAME].delete()

    sheet = workbook.sheets.add(
        name=_UNIVARIATE_SHEET_NAME,
        after=workbook.sheets[-1],
    )

    val(sheet, _ROW_FIT_ZONE, _C_GS_BETA, "Univariate Analysis (Beta-only / Cartesian)")
    sheet.range(rc(_ROW_FIT_ZONE, _C_GS_BETA)).api.Font.Bold = True

    _write_data_feed(sheet)
    _setup_local_names(sheet)

    _write_two_stage_grid_search_cartesian(
        sheet,
        col_start=_C_GS_BETA,
        body_prefix="UV_BETA",
        p1_label="Alpha (α)",
        p2_label="Beta (β)",
        p1_min=0.2,
        p1_max=50.0,
        p2_min=0.2,
        p2_max=50.0,
        beta_grid_size=beta_grid_size,
    )

    return sheet

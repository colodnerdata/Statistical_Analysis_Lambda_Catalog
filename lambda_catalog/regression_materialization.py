"""The ARCHITECTURE §4b materialization band for the Regression template.

Extracted from ``write_sheet_regression.py`` (reunify-workbook Part 5.2, the
twin of Part 5.1's ``regression_charts.py``) so the writer is its zone writers +
orchestrator. This one function writes the whole band that sits past the charts:
the fixed-size **Model Context** block, the materialized ``Sample_Include`` row
mask, and the terminal **Constructed Design Matrix** — each separated by an
ungrouped gutter column.

Every column and row it addresses is a constant from ``regression_layout.py``
(Part 1.1) — ``_LAST_CHART_COLUMN`` and the ``_C_*`` zone columns, the
``_MATERIALIZATION_*`` rows, ``_MODEL_CONTEXT_ELEMENTS`` — so the band tracks the
chart anchor rather than any letter written here. The function depends on those
constants, the standard cell-write helpers, the catalog reader for the
sheet-scoped closures, and the sheet object; nothing else. Importing them here
creates no cycle back into ``write_sheet_regression``.

``write_sheet_regression`` re-exports ``_write_materialization_zone`` so its own
call site in ``write_regression_output_sheet`` and the three test import sites
keep working unchanged.
"""
from __future__ import annotations

import xlwings as xw

from .catalog_schema import CatalogFunction, load_catalog_document
from .regression_layout import (
    REGRESSION_SHEET_NAME,
    _CHART_RIGHT_OFFSET_PT,
    _C_BB,
    _C_DESIGN_MATRIX,
    _C_DESIGN_MATRIX_NAMES,
    _C_GUTTER_AFTER_CHARTS,
    _C_GUTTER_AFTER_CONTEXT,
    _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    _C_MODEL_CONTEXT,
    _C_MODEL_CONTEXT_LABEL,
    _C_MODEL_FORMULA,
    _C_MODEL_FORMULA_LABEL,
    _C_SAMPLE_INCLUDE_MATERIALIZED,
    _DEFINITIONS_PATH,
    _DESIGN_MATRIX_COLUMN_WIDTH,
    _DESIGN_MATRIX_INTERCEPT_HEADER,
    _DESIGN_MATRIX_SIZED_COLUMNS,
    _LAST_CHART_COLUMN,
    _MATERIALIZATION_FIRST_ROW,
    _MATERIALIZATION_HEADER_ROW,
    _MATERIALIZATION_SPILL_ROW,
    _MODEL_CONTEXT_ELEMENTS,
    _MODEL_CONTEXT_LABEL_WIDTH,
    _MODEL_CONTEXT_LAST_ROW,
    _MODEL_CONTEXT_ROWS,
    _MODEL_CONTEXT_VALUE_WIDTH,
    _ROW_MODEL_CONTEXT_CHECK,
    _ROW_MODEL_FORMULA,
    _SAMPLE_INCLUDE_HEADER,
    _SAMPLE_INCLUDE_MATERIALIZED_WIDTH,
)
from .workbook_helpers import (
    a1,
    bold_row,
    border_box,
    col_letter,
    drop_local_name,
    f,
    quoted_sheet_name,
    rc,
    section_heading,
    val,
)


def _add_spill_reader(
    sheet: xw.Sheet, sheet_name: str, name: str, column: int
) -> None:
    """Register a sheet-scoped thunk over one materialized spill.

    The v3.2 half that makes the band pay for itself. Excel does not memoize a
    name whose RefersTo is a formula, so ``Design_Columns()`` written into
    forty-four cells constructs the design matrix forty-four times. These names
    let a call site read the ONE materialized copy instead.

    ``=LAMBDA('Sheet'!$BW$3#)`` — the anchor plus the spill operator, so the
    reference tracks the array's live extent in both dimensions without a count
    cell. That is the part no other name in this workbook does; ``Fit_Context``
    wraps a FIXED range and sidesteps the question.

    **``#`` inside a defined-name RefersTo works — confirmed in Excel on #223**,
    which is what the two-call-site spike existed to establish. Both shapes
    resolve: the 1-D mask (``ROWS(Source_Data)`` tall) and the 2-D design matrix
    (``Observations`` x ``$O$1``). Treat this as settled rather than re-testing
    it; what still needs Excel is whether a MIGRATED CALL SITE reads the right
    rows, which is a different question and the reason the rest of the rewiring
    goes zone by zone against the cell-by-cell verifier.

    The escape hatch if it ever regresses is the OFFSET-sized-by-a-count-cell
    form every ``RegChart*`` range already uses — same call syntax at every call
    site, so only this function would change. The dimensions it needs are live
    cells and were confirmed on the same run: ``$AB$8`` for height, ``$O$1`` for
    the design matrix's width.

    Wrapped in ``LAMBDA`` rather than left as a bare range name so call sites
    read ``Fit_Design_Columns()``, matching ``Fit_Context()`` and the
    constructor closures beside it. A bare range name would work too but would
    make the sheet's formulas inconsistent about which names are callable.

    Parameters
    ----------
    sheet : xw.Sheet
        Sheet to register the name on.
    sheet_name : str
        Bare sheet name; this function quotes it. Every generated test-model
        sheet has a space in its name ("M01 Baseline Categoricals"), and an
        unquoted RefersTo is an invalid formula that makes Excel reject the
        whole ``Names.Add``. Quoting a name that does not need it is always
        legal, so the quotes live here rather than at each call site — the
        same reasoning ``_setup_local_names`` records for its own ``sname``.
    name : str
        The defined name to create, dropping any stale copy first.
    column : int
        1-based column of the spill's anchor cell; the row is always
        ``_MATERIALIZATION_SPILL_ROW``.
    """
    anchor = (
        f"{quoted_sheet_name(sheet_name)}"
        f"!${col_letter(column)}${_MATERIALIZATION_SPILL_ROW}"
    )
    drop_local_name(sheet, name)
    sheet.api.Names.Add(Name=name, RefersTo=f"=LAMBDA({anchor}#)")


def _write_materialization_zone(
    sheet: xw.Sheet,
    closures: tuple[CatalogFunction, ...] | None = None,
) -> None:
    """Write the ARCHITECTURE §4b materialization band at the sheet's far right.

    Two names share the model context, by design (v3.0 stage two):

    ``Model_Context(...)`` is the WORKBOOK-scoped constructor — the one
    definition of the default context (Has_Intercept=TRUE, DF_Absorbed=0,
    both transforms "None"). It takes optional overrides, is what every
    engine's omitted-[Context] default routes through, and is what the MLR
    test sheets call with a per-row flag.

    ``Fit_Context()`` is the SHEET-scoped reader — a zero-arg thunk over the
    FIXED range holding the materialized context block. It is what the ~30
    Regression sheet call sites pass, so they read the actual spec-derived
    context (the C2 Allow_Intercept toggle, the Absorbed_Degrees_Of_Freedom()
    closure, the spec-block transform summaries) rather than the constructor
    default. Splitting the names keeps ``Model_Context`` unshadowed: a single
    sheet-scoped thunk named ``Model_Context`` would make ``Model_Context()``
    in a sheet cell resolve to the materialized values while the same token in
    a carrier's omitted-default resolved to the workbook constructor — the
    invisible shadowing the v3.0 release exists to remove.

    The context is materialized ONCE (Excel does not memoize a name whose
    RefersTo is a formula, so a constructor inside ~30 engine calls runs ~30
    times); ``Fit_Context`` reads the fixed range those cells occupy, so the
    ~30 call sites that pass ``Fit_Context()`` all read the one materialized
    block.

    The block is written as ONE CELL PER ELEMENT, each with its own label in
    the column to its left, and boxed — it is a fixed-size table, not a data
    range. A single spill would buy nothing (the height is a build-time
    constant) and would cost correctness: a spill is one dependency node that
    Excel vacates and re-spills whenever the spec block changes, and while it
    is vacated the range behind ``Fit_Context`` is transiently blank, so
    every engine reading it sees a torn context. Independent cells
    recalculate independently and are never vacated.

    Elements 3-4 (Response_Transform, Predictor_Transform) are populated from
    the spec block now but have no engine reader until the v3.3 unit-space
    dispatcher; the row order is the contract that is expensive to change
    later, so all four rows land together. An error in an unconsumed row is
    contained to its own cell now that each element stands alone, and the
    engines read only elements 1-2 through the accessors, so a bad name
    surfaces as a visible cell error (caught by the headless verifier and by
    the block's own health-check row) without shifting a single fitted number.

    The two data-dependent zones — the ``Sample_Include`` row mask and the
    terminal Constructed Design Matrix — are SURFACED here: each carries a
    column-header row at ``_MATERIALIZATION_HEADER_ROW`` and spills from
    ``_MATERIALIZATION_SPILL_ROW``, full height and row-aligned with the source
    table, so the mask reads straight across into the design-matrix row beside
    it. Both replaced a ``"reserved"`` placeholder that held the position while
    stage 3 established the layout and the pre-flight width guard. Neither is
    grouped: they hold spills, and a collapsed outline group over a spill range
    leaves it stale on recalculation, so the model refits on old values. Only
    the fixed-height Model Context block — individual cells, no spill — carries
    a group, and it is the only zone that ships collapsed.

    Surfacing the values is NOT the same as rewiring the readers, and the two
    halves ship separately. ``Fit_Sample_Include`` and ``Fit_Design_Columns``
    are the readers: sheet-scoped thunks over the two spills, registered below
    beside ``Fit_Context``, which is the same idea over the fixed-height
    context block. They use the dynamic-array spill operator (``#``) inside a
    ``LAMBDA`` defined-name RefersTo — a combination used nowhere else in this
    workbook, which is why the v3.2 remainder spiked two call sites before
    migrating the rest. It resolves correctly in Excel (confirmed on #223), so
    the mechanism is settled and the remaining risk is per-call-site, not
    structural.

    **New names, not promotions, and that is structural.** ``Sample_Include``
    and ``Design_Columns`` keep their meanings because the spill cells here
    ARE ``=Sample_Include()`` and ``=Design_Columns()`` — promoting either name
    to read its own spill would make the cell producing it self-referential.
    ``Sample_Include`` additionally keeps an optional ``apply_log_domain``
    argument that a materialized column cannot express: ``Log_Domain_Status``
    calls ``Sample_Include(FALSE)`` for the mask BEFORE the positivity layer,
    and only the default is materialized here.

    A wrong range in one of these names does not error — it returns numbers
    from the wrong rows — so the migration goes zone by zone against the
    cell-by-cell spec verifier rather than in one sweep.

    The design matrix's header row is split across two cells:
    ``Design_Columns()`` is one column wider than
    ``Constructed_Column_Names()`` when the intercept is on, because the
    constructor prepends the ones column, so the anchor cell names that column
    and the names spill starts beside it. With ``Allow_Intercept`` FALSE there
    is no ones column and the names sit one column right of the values they
    label — a labelling offset in an otherwise unread display zone, called out
    on the heading cell's note.

    Plain cell writes + defined-name registration only; no chart/COM API, so
    this is exercised in unit tests via ``RecordingSheet`` without Excel. The
    chart-footprint clearance check is the one Excel-only step and is guarded.
    """
    sname = sheet.name
    if closures is None:
        closures = load_catalog_document(_DEFINITIONS_PATH).functions_for_sheet(
            REGRESSION_SHEET_NAME
        )

    # ── Model Context (fixed-size labelled block) ────────────────────────────
    # One row per element, label left of value, heading on row 1, the whole
    # thing boxed — the same treatment every other fixed-size block on this
    # sheet gets (Regression Statistics, Diagnostics, Prediction Interval).
    # Element 1 is the C2 Allow_Intercept toggle (named range). Element 2 is
    # the Absorbed_Degrees_Of_Freedom() Regression closure. Elements 3-4 are
    # the spec-block transform summaries (see _RESPONSE_TRANSFORM_FORMULA /
    # _PREDICTOR_TRANSFORM_FORMULA), both reading worksheet-scoped spec names
    # on this sheet. _MODEL_CONTEXT_ELEMENTS is the single source of the row
    # order, the labels, and the height.
    ctx_col = col_letter(_C_MODEL_CONTEXT)
    section_heading(sheet, 1, _C_MODEL_CONTEXT_LABEL, "MODEL CONTEXT")
    # The value column's heading cell carries the fill but no text, so the
    # heading reads as one banner across the two-column block.
    section_heading(sheet, 1, _C_MODEL_CONTEXT, "")

    for offset, element in enumerate(_MODEL_CONTEXT_ELEMENTS):
        row = _MATERIALIZATION_FIRST_ROW + offset
        val(sheet, row, _C_MODEL_CONTEXT_LABEL, element.label)
        f(sheet, row, _C_MODEL_CONTEXT, f"={element.formula}")

    # The materialized block occupies _MATERIALIZATION_FIRST_ROW ..
    # _MODEL_CONTEXT_LAST_ROW; the sheet-scoped reader Fit_Context reads that
    # fixed range (no spill operator — the height is a structural constant, so
    # a fixed range is exact and avoids the dynamic-array-in-a-name question
    # entirely). Drop any stale "Model_Context" or "Fit_Context" name before
    # re-adding, so a rebuild never leaves a shadow.
    ctx_ref = (
        f"{quoted_sheet_name(sname)}!${ctx_col}${_MATERIALIZATION_FIRST_ROW}"
        f":${ctx_col}${_MODEL_CONTEXT_LAST_ROW}"
    )
    drop_local_name(sheet, "Model_Context")
    drop_local_name(sheet, "Fit_Context")
    sheet.api.Names.Add(
        Name="Fit_Context",
        RefersTo=f"=LAMBDA({ctx_ref})",
    )
    # Health check, one row under the block and inside its box. The height
    # half is the build-time invariant (_MODEL_CONTEXT_ROWS); the error half
    # is what decomposition made worth checking — with four independent cells
    # a broken spec name errors in ONE of them and leaves the other three
    # looking fine, so the block reports whether every element resolved.
    val(sheet, _ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT_LABEL, "Context OK")
    f(
        sheet,
        _ROW_MODEL_CONTEXT_CHECK,
        _C_MODEL_CONTEXT,
        f"=AND(ROWS(Fit_Context())={_MODEL_CONTEXT_ROWS},"
        "SUMPRODUCT(--ISERROR(Fit_Context()))=0)",
    )
    border_box(
        sheet, 1, _C_MODEL_CONTEXT_LABEL, _ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT
    )

    # ── Sample_Include (materialized row mask) ───────────────────────────────
    # The mask spills full-height and row-aligned with the source table (the
    # row-mask contract), so it reads straight across into the design-matrix
    # rows beside it. This SURFACES the value; it does not rewire the closure —
    # Sample_Include() is still evaluated per call site, and promoting it to a
    # thunk over this spill stays deferred (see the docstring).
    section_heading(sheet, 1, _C_SAMPLE_INCLUDE_MATERIALIZED, "Sample Include")
    val(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        _SAMPLE_INCLUDE_HEADER,
    )
    bold_row(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
    )
    f(
        sheet,
        _MATERIALIZATION_SPILL_ROW,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        "=Sample_Include()",
    )
    # Fit_Sample_Include — the reader over the spill written above. 1-D (n x 1,
    # one row per SOURCE row, not per included row), so the fallback if `#`
    # misbehaves is a single-count OFFSET. Registered whether or not any call
    # site reads it yet: the name is what the migration repoints cells AT.
    _add_spill_reader(
        sheet,
        sname,
        "Fit_Sample_Include",
        _C_SAMPLE_INCLUDE_MATERIALIZED,
    )
    # Document what the column is for on the heading cell — it is a read-only
    # view of the mask every engine applies, not a second place to edit it.
    try:
        sheet.range(rc(1, _C_SAMPLE_INCLUDE_MATERIALIZED)).api.AddComment(
            "TRUE for every source row the model is fitted on. Read-only view "
            "of the row mask the engines apply — change it through the spec "
            "block (Include, and any Role=Filter column), not here. Full "
            "height and row-aligned with the source table, so it reads "
            "straight across into the design matrix to its right."
        )
    except Exception:  # pylint: disable=broad-except
        pass

    # ── Constructed Design Matrix (terminal zone) ────────────────────────────
    # The zone that terminates the band. Its width is unbounded and one
    # dropdown away — Country as a Categorical Predictor is 156 columns, and
    # interactions multiply — which is why nothing may ever be placed to its
    # right. The zone ships expanded, not collapsed: an unbounded-width zone
    # left open is a scrolling hazard, but hiding the columns a full-height
    # spill occupies leaves Design_Columns() stale on recalculation, so every
    # engine reading the matrix would fit on old values. The scrolling is the
    # accepted cost; the zone stays expanded.
    #
    # Establishing the zone and MATERIALIZING into it were deliberately
    # separate steps: the position and the width guard that reads the spec's
    # pre-flight column count are what a later release could not add without
    # moving columns again, so they landed first and the spill is a formula
    # change into columns that already exist.
    #
    # The header row is split across two cells because the matrix is one
    # column wider than its names when the intercept is on — the anchor cell
    # names the constructor's ones column and the names spill starts beside
    # it. See _DESIGN_MATRIX_INTERCEPT_HEADER.
    section_heading(sheet, 1, _C_DESIGN_MATRIX, "Constructed Design Matrix")
    f(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_DESIGN_MATRIX,
        _DESIGN_MATRIX_INTERCEPT_HEADER,
    )
    f(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_DESIGN_MATRIX_NAMES,
        "=Constructed_Column_Names()",
    )
    bold_row(
        sheet, _MATERIALIZATION_HEADER_ROW, _C_DESIGN_MATRIX, _C_DESIGN_MATRIX_NAMES
    )
    f(sheet, _MATERIALIZATION_SPILL_ROW, _C_DESIGN_MATRIX, "=Design_Columns()")
    # Fit_Design_Columns — the reader over the design-matrix spill. 2-D
    # (n x k, both dimensions dynamic), which is the harder of the two shapes:
    # the `#` form is dimension-agnostic, but an OFFSET fallback would need a
    # height AND a width ($AB$8 and $O$1 respectively). Spiking both shapes is
    # what tells us which spelling the migration can rely on.
    _add_spill_reader(sheet, sname, "Fit_Design_Columns", _C_DESIGN_MATRIX)

    # ── Model Formula readout ────────────────────────────────────────────────
    # Row 1 of this zone, right of its heading — the one row the design matrix
    # itself can never reach (its names and values spill from the rows below
    # and grow rightward), so a caption here is never displaced by an ordinary
    # modelling choice, and with WrapText OFF it overflows across as much of an
    # empty row 1 as the string needs. See _ROW_MODEL_FORMULA.
    #
    # The string is assembled by the sheet-scoped Model_Formula() closure
    # (lambda_functions.json, scope "Regression") — the cell holds a call, not
    # a 300-character expression, so the assembly rules live in the catalog
    # with every other spec-derived constructor and the LAMBDA_functions sheet
    # documents them.
    #
    # WrapText is set FALSE explicitly rather than left at the default: this
    # cell holds the longest string on the sheet, and inheriting a wrap (from a
    # future band-wide format, or from a copy of this block) would make one
    # caption dictate a row height — exactly what moving it off row 2 of the
    # Regression Outputs zone was meant to stop.
    section_heading(sheet, _ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL, "Model Formula")
    f(sheet, _ROW_MODEL_FORMULA, _C_MODEL_FORMULA, "=Model_Formula()")
    sheet.range(
        rc(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL),
        rc(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA),
    ).api.WrapText = False
    try:
        sheet.range(rc(1, _C_DESIGN_MATRIX)).api.AddComment(
            "The design matrix the engines actually fit: Design_Columns(), "
            "full height and row-aligned with the source table and with the "
            "Sample Include mask to its left (the engines apply that mask "
            "themselves). Terminal §4b zone — nothing may ever be placed to "
            "its right, because its width is unbounded and one dropdown away. "
            "Left expanded and ungrouped on purpose: collapsing the columns a "
            "spill occupies leaves the matrix stale on recalculation.\n\n"
            "Header row: the anchor cell names the intercept column the "
            "constructor prepends, and Constructed_Column_Names() spills from "
            "the column beside it. With Allow Intercept FALSE there is no "
            "ones column, so the names sit one column right of the values "
            "they label."
        )
    except Exception:  # pylint: disable=broad-except
        pass

    # Both data-dependent zones spill from the same row so the band reads
    # across — the mask value beside its design-matrix row, both aligned to
    # the source table rows. Asserted rather than merely intended.
    assert _MATERIALIZATION_SPILL_ROW == _MATERIALIZATION_HEADER_ROW + 1
    assert _MATERIALIZATION_FIRST_ROW == 2

    # ── Column widths + outline groups ───────────────────────────────────────
    # Only the Model Context zone is grouped, and only it ships collapsed. The
    # width-2 gutters stay ungrouped, and the first gutter (after the charts)
    # is structural — it keeps the floating chart anchors out of the
    # collapsible outline group.
    for gutter in (
        _C_GUTTER_AFTER_CHARTS,
        _C_GUTTER_AFTER_CONTEXT,
        _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    ):
        sheet.range(f"{col_letter(gutter)}:{col_letter(gutter)}").column_width = 2
    for content, width in (
        (_C_MODEL_CONTEXT_LABEL, _MODEL_CONTEXT_LABEL_WIDTH),
        (_C_MODEL_CONTEXT, _MODEL_CONTEXT_VALUE_WIDTH),
        (_C_SAMPLE_INCLUDE_MATERIALIZED, _SAMPLE_INCLUDE_MATERIALIZED_WIDTH),
    ):
        sheet.range(f"{col_letter(content)}:{col_letter(content)}").column_width = width
    # The Model Context zone is grouped as the label/value PAIR, not per
    # column, so it collapses as a unit — grouping the value column alone
    # would strand its labels beside a collapsed column. It is the band's only
    # group: it is a fixed-height block of individual cells, so hiding it
    # hides no spill.
    #
    # Sample_Include and the Constructed Design Matrix are NOT grouped and NOT
    # collapsed. Both are full-height dynamic-array spills, and a collapsed
    # group over a spill range is the state in which Excel stops recalculating
    # the model — the hidden columns keep the stale arrays, and the ~30 engine
    # call sites reading across them fit on stale values. Do not re-add a
    # Group()/ShowDetail call for either one.
    context_band = f"{col_letter(_C_MODEL_CONTEXT_LABEL)}:{col_letter(_C_MODEL_CONTEXT)}"
    sheet.api.Columns(context_band).Group()
    # ShowDetail is an ActiveWindow-free property on the range, but it still
    # needs a real outline underneath, so guard it the way every other cosmetic
    # COM call on this sheet is guarded — a workbook that opens with the block
    # expanded is a nuisance, not a broken build.
    try:
        sheet.api.Columns(context_band).ShowDetail = False
    except Exception:  # pylint: disable=broad-except
        pass

    # Width only for the terminal zone — sized across the bounded band the
    # width guard already bounds, with no outline over it.
    matrix_band = (
        f"{col_letter(_C_DESIGN_MATRIX)}:"
        f"{col_letter(_C_DESIGN_MATRIX + _DESIGN_MATRIX_SIZED_COLUMNS - 1)}"
    )
    sheet.range(matrix_band).column_width = _DESIGN_MATRIX_COLUMN_WIDTH

    # ── Chart-footprint clearance assertion (Excel only) ────────────────────
    # _LAST_CHART_COLUMN is a conservative bound; this verifies the column
    # past the footprint actually clears the computed chart right edge, so a
    # chart resize that would overlap the context block fails the build.
    #
    # The geometry LOOKUP is best-effort — COM geometry (sheet.range(...).left)
    # is unavailable headless, so that raise is swallowed and the check is
    # skipped there (the conservative constant keeps the layout safe by
    # construction). But the clearance ASSERT itself must NOT be swallowed:
    # wrapping it in the same broad except would make the guard a no-op in
    # Excel, the one place it can actually run. So acquire the geometry under
    # the guard, then assert outside it.
    try:
        chart_right = (
            sheet.range(a1(1, _C_BB)).left + _CHART_RIGHT_OFFSET_PT
        )
        clear_left = sheet.range(a1(1, _LAST_CHART_COLUMN + 1)).left
    except Exception:  # pylint: disable=broad-except — headless / no COM geometry
        chart_right = None
        clear_left = None
    if chart_right is not None and clear_left is not None:
        assert clear_left >= chart_right, (
            f"chart footprint ({chart_right:.0f}pt) overlaps the materialization "
            f"zone (column {col_letter(_LAST_CHART_COLUMN + 1)} left edge "
            f"{clear_left:.0f}pt); raise _LAST_CHART_COLUMN"
        )

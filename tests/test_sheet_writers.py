"""Tests for workbook-writing logic that do not require a live Excel process."""
# pylint: disable=import-outside-toplevel,missing-function-docstring,protected-access
from pathlib import Path
from typing import cast

import xlwings as xw

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
    HEADER_COLOR,
    INPUT_COLOR,
)
from lambda_catalog.workbook_helpers import (
    _NOTE_MAX_WIDTH,
    _NOTE_MIN_WIDTH,
    add_expression_format,
    col_letter,
    excel_color,
    note_dimensions as _note_dimensions,
    rc,
)
from lambda_catalog.write_sheet_regression import (
    _C_AA,
    _C_AB,
    _C_AC,
    _C_AD,
    _C_AE,
    _C_AF,
    _C_AG,
    _C_AH,
    _C_AJ,
    _C_AK,
    _C_AL,
    _C_AN,
    _C_AO,
    _C_AQ,
    _C_AR,
    _C_AX,
    _C_AY,
    _C_AZ,
    _C_BA,
    _C_BB,
    _A_BACK_TRANSFORM_METHOD,
    _C_DESIGN_MATRIX,
    _C_DESIGN_MATRIX_NAMES,
    _C_SPEC_DESIGN_COLUMNS,
    _C_SPEC_INTERACTION_OPERATION,
    _C_SPEC_INTERACTION_TERM,
    _C_SPEC_SEQUENCE_PERIOD,
    _C_P,
    _C_Q,
    _C_S,
    _C_X,
    _C_Y,
    _COLUMN_WIDTHS,
    _GAP_COLUMNS,
    _ZONES,
    _DESIGN_MATRIX_GROUPED_COLUMNS,
    _DESIGN_MATRIX_GROUPED_WIDTH,
    _DESIGN_MATRIX_INTERCEPT_HEADER,
    _DESIGN_MATRIX_MAX_COLUMNS,
    _DESIGN_MATRIX_SOFT_CELLS,
    _DESIGN_MATRIX_SOFT_COLUMNS,
    _MATERIALIZATION_FIRST_ROW,
    _MATERIALIZATION_HEADER_ROW,
    _MATERIALIZATION_SPILL_ROW,
    _SAMPLE_INCLUDE_HEADER,
    _SAMPLE_INCLUDE_MATERIALIZED_WIDTH,
    _C_GUTTER_AFTER_CHARTS,
    _C_GUTTER_AFTER_CONTEXT,
    _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    _C_MODEL_CONTEXT,
    _C_MODEL_CONTEXT_LABEL,
    _C_MODEL_FORMULA,
    _C_MODEL_FORMULA_LABEL,
    _ROW_MODEL_FORMULA,
    _C_SAMPLE_INCLUDE_MATERIALIZED,
    _C_CHART_LABEL_NAME,
    _C_CHART_TITLE,
    _C_CHART_XLABEL,
    _C_CHART_YLABEL,
    _CHART_Y_TICK_FORMATS,
    _CHART_Y_TICK_FORMAT_DEFAULT,
    _MODEL_CONTEXT_ELEMENTS,
    _MODEL_CONTEXT_LABEL_WIDTH,
    _MODEL_CONTEXT_LAST_ROW,
    _MODEL_CONTEXT_ROWS,
    _MODEL_CONTEXT_VALUE_WIDTH,
    _ROW_MODEL_CONTEXT_CHECK,
    _PRED_INPUT_FIRST_ROW,
    _PRED_INPUT_LAST_ROW,
    _NOTE_SIZE_OVERRIDES,
    _ROW_CHART_LABELS,
    _diagnostic_chart_specs,
    _setup_local_names as _setup_regression_names,
    _write_chart_label_cells,
    _write_coefficients,
    _write_diagnostics,
    _write_design_matrix_width_guard,
    _write_prediction_interval,
    _write_prediction_inputs,
    _write_regression_outputs_header,
    _write_residuals,
    _BACK_TRANSFORM_METHODS,
    _write_unit_space_block,
    _write_materialization_zone,
)
from lambda_catalog.write_sheet_univariate import (
    _C_FIT_FIRST,
    _HIST_COLUMNS,
    _STAT_ROWS,
    _dist_rows,
    _setup_local_names,
    _write_data_zone,
    _write_descriptive_stats,
    _write_fitting_table,
    _write_grid_stage,
    _write_histogram_table,
    _write_qq_data,
    _write_weibull_grid_search,
)
from tests.recording_sheet import RecordingName, RecordingNames, RecordingSheet

ROOT_DIR = Path(__file__).resolve().parents[1]


def _as_xw_sheet(sheet: RecordingSheet) -> xw.Sheet:
    return cast(xw.Sheet, sheet)


class BrokenIndexRecordingNames(RecordingNames):
    def __call__(self, index: int | str) -> RecordingName:
        if isinstance(index, int):
            raise OSError("broken integer enumeration")
        return super().__call__(index)


def _formula(sheet: RecordingSheet, row: int, col: int) -> str:
    return cast(str, sheet.cell(row, col).api.Formula2)


def test_note_dimensions_clamps_width_to_configured_bounds() -> None:
    tiny_width, _ = _note_dimensions("tiny", "x")
    assert tiny_width >= _NOTE_MIN_WIDTH

    huge_width, _ = _note_dimensions("huge", "x" * 5000)
    assert huge_width == _NOTE_MAX_WIDTH


def test_note_dimensions_height_grows_with_multi_paragraph_text() -> None:
    paragraph = "word " * 40
    _, single_height = _note_dimensions("single", paragraph)
    _, two_height = _note_dimensions("two", paragraph + "\n" + paragraph)

    assert two_height > single_height


def test_note_dimensions_override_replaces_only_the_given_axis() -> None:
    label = "Test Note Label"
    text = "some note text that would otherwise be auto-sized"
    default_width, default_height = _note_dimensions("Unlabeled", text)

    _NOTE_SIZE_OVERRIDES[label] = (default_width + 50.0, None)
    try:
        width, height = _note_dimensions(label, text, _NOTE_SIZE_OVERRIDES)
    finally:
        del _NOTE_SIZE_OVERRIDES[label]

    assert width == default_width + 50.0
    assert height == default_height


def test_regression_names_register_spec_wiring_and_constructors() -> None:
    sheet = RecordingSheet(name="Regression")

    _setup_regression_names(_as_xw_sheet(sheet))

    names = [item.Name.split("!", 1)[-1] for item in sheet.api.Names.items]
    # Spec wiring precedes the closures, which precede the Regression-only names.
    assert names.index("Spec_Include") < names.index("Sample_Include")
    assert names.index("Sample_Include") < names.index("Predictor_Columns")
    assert names.index("Predictor_Columns") < names.index("Zero_Predictors_Selected")
    # The v1 hard-wired names are gone.
    for legacy in ("All_Xs", "Coefficient_Name_Col", "Ind_Var_Include", "y",
                   "Regression_Sample_Include", "data_identifiers"):
        assert legacy not in names, legacy

    predictor_formula = sheet.api.Names.by_short_name("Predictor_Columns").RefersTo
    assert predictor_formula.startswith("=LAMBDA(LET(")
    assert "Dummy_Levels(" in predictor_formula

    # The width probe counts PREDICTOR columns. Design_Columns() would report 1
    # in exactly this state — the intercept stage still runs when the predictor
    # stage is empty — and the zero-predictor branch would never fire.
    zero_formula = sheet.api.Names.by_short_name("Zero_Predictors_Selected").RefersTo
    assert zero_formula == "=LAMBDA(IFERROR(COLUMNS(Predictor_Columns()),0)=0)"

    assert sheet.api.Names.by_short_name("Allow_Intercept").RefersTo == (
        "='Regression'!$C$2"
    )
    # Single-quoted, like every other name on the sheet. These four
    # Regression-only names used to be registered UNQUOTED, which happened to
    # work only because "Regression" is a single word — a sheet name with a
    # space made the RefersTo an invalid formula and Excel rejected the
    # Names.Add outright. See
    # tests/test_test_model_sheets.py::test_every_refers_to_quotes_a_sheet_name_containing_spaces.
    assert sheet.api.Names.by_short_name("alpha").RefersTo == "='Regression'!$AB$12"


def test_regression_chart_names_size_to_the_observation_cell() -> None:
    sheet = RecordingSheet(name="Regression")

    _setup_regression_names(_as_xw_sheet(sheet))

    fit_y = sheet.api.Names.by_short_name("RegChartFitY").RefersTo
    assert fit_y == (
        "=OFFSET('Regression'!$AP$2,1,0,"
        "MAX(IFERROR('Regression'!$AB$8,1),1),1)"
    )
    press = sheet.api.Names.by_short_name("RegChartPRESSResid").RefersTo
    assert "$AX$2" in press
    cooks_flag = sheet.api.Names.by_short_name("RegChartCookDistFlag").RefersTo
    assert "$AY$2" in cooks_flag
    obs_label = sheet.api.Names.by_short_name("RegChartObsLabel").RefersTo
    assert "$AN$2" in obs_label
    assert {
        name: sheet.api.Names.by_short_name(name).Comment
        for name in (
            "RegChartQQX",
            "RegChartQQY",
            "RegChartFitY",
            "RegChartResid",
            "RegChartActY",
            "RegChartScaleLoc",
            "RegChartCookDist",
            "RegChartLeverage",
            "RegChartStudResid",
            "RegChartPRESSResid",
            "RegChartCookDistFlag",
            "RegChartObsLabel",
        )
    } == {
        "RegChartQQX": "Normal Q-Q chart: X values (theoretical quantiles, Normal Scores Ranked)",
        "RegChartQQY": "Normal Q-Q chart: Y values (Studentized Residuals Ranked)",
        "RegChartFitY": (
            "Predicted Y: X values for the Residuals vs. Fitted, Actual vs. Predicted, "
            "and Scale-Location charts"
        ),
        "RegChartResid": "Residuals vs. Fitted chart: Y values (Residuals)",
        "RegChartActY": "Actual vs. Predicted chart: Y values (Actual Y)",
        "RegChartScaleLoc": "Scale-Location chart: Y values (sqrt of abs Studentized Residual)",
        "RegChartCookDist": "Cook's Distance chart: bar values",
        "RegChartLeverage": "Studentized Residuals vs. Leverage chart: X values (Hat Diagonal)",
        "RegChartStudResid": "Studentized Residuals vs. Leverage chart: Y values",
        "RegChartPRESSResid": "PRESS Residuals chart: bar values",
        "RegChartCookDistFlag": (
            "Cook's Distance chart: flagged-point overlay for data labels (D > 4/n or D > 0.9)"
        ),
        "RegChartObsLabel": (
            "Cook's Distance chart: observation identifier for flagged-point data labels"
        ),
    }


def test_chart_label_cells_reference_live_statistics_and_stay_ordered() -> None:
    """Chart Title / X-Axis / Y-Axis formula cells are written one row per
    chart, in chart_specs order, and the dynamic titles reference the correct
    sheet cells (response name, QQ correlation, Cook's D threshold, etc.)."""
    sheet = RecordingSheet(name="Regression")

    _write_chart_label_cells(_as_xw_sheet(sheet))

    specs = _diagnostic_chart_specs()
    assert len(specs) == 7
    for i, (key, _chart_type, _x_addr, _y_addr, title_formula, x_label_formula,
            y_label_formula, _grid_row, _grid_col) in enumerate(specs):
        row = _ROW_CHART_LABELS + i
        assert sheet.ranges[((row, _C_CHART_LABEL_NAME),)].state.value == key
        assert sheet.ranges[((row, _C_CHART_TITLE),)].state.formula2 == title_formula
        assert sheet.ranges[((row, _C_CHART_XLABEL),)].state.formula2 == x_label_formula
        assert sheet.ranges[((row, _C_CHART_YLABEL),)].state.formula2 == y_label_formula

    cooks_row = _ROW_CHART_LABELS + [s[0] for s in specs].index("Cook's Distance")
    cooks_title = sheet.ranges[((cooks_row, _C_CHART_TITLE),)].state.formula2
    assert "MIN(4/$AB$8,0.9)" in cooks_title

    qq_row = _ROW_CHART_LABELS + [s[0] for s in specs].index("Normal Q-Q")
    qq_title = sheet.ranges[((qq_row, _C_CHART_TITLE),)].state.formula2
    assert "$AE$10" in qq_title

    fitted_row = _ROW_CHART_LABELS + [s[0] for s in specs].index("Residuals vs. Fitted")
    fitted_title = sheet.ranges[((fitted_row, _C_CHART_TITLE),)].state.formula2
    assert "$AF$2" in fitted_title


def test_scale_location_y_axis_always_shows_one_decimal() -> None:
    # Scale-Location plots sqrt(|studentized residual|), which almost always
    # sits inside 0-2. Under the default "0" tick format the axis renders
    # 0/1/2 and throws away the spread the chart exists to show, so it is
    # pinned to one decimal. The chart writer is COM-only and cannot run
    # headless, so the format TABLE is what gets asserted — which is the
    # reason the formats are a declarative dict rather than a chain of
    # `if key ==` overrides inside the writer.
    assert _CHART_Y_TICK_FORMATS["Scale-Location"] == "0.0"

    # Every key in the table must name a real chart, or a renamed chart would
    # silently fall back to the default instead of failing.
    keys = {spec[0] for spec in _diagnostic_chart_specs()}
    assert set(_CHART_Y_TICK_FORMATS) <= keys

    # Everything else keeps the default; Cook's Distance is the one other
    # exception (1e-4 to 1e-1 and heavily right-skewed, so scientific).
    assert _CHART_Y_TICK_FORMAT_DEFAULT == "0"
    assert set(_CHART_Y_TICK_FORMATS) == {"Scale-Location", "Cook's Distance"}


def test_intercept_only_n_does_not_depend_on_filter() -> None:
    sheet = RecordingSheet(name="Regression")

    _setup_regression_names(_as_xw_sheet(sheet))

    intercept_only_n_formula = sheet.api.Names.by_short_name("Intercept_Only_N").RefersTo
    assert "FILTER" not in intercept_only_n_formula
    # SUMPRODUCT over the computed mask: COUNTIF needs a range reference and
    # Sample_Include() is an array; SUMPRODUCT never errors on an empty mask.
    assert "SUMPRODUCT(N(Sample_Include()))" in intercept_only_n_formula


def test_prediction_interval_binds_constructed_inputs_in_the_cell_formula() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_prediction_interval(_as_xw_sheet(sheet))

    formula = sheet.cell(3, _C_AK).api.Formula2
    assert formula is not None
    assert formula.startswith("=IF(Zero_Predictors_Selected(),")
    assert "IFERROR" not in formula
    # Inputs correspond 1:1 to constructed columns — TAKE exactly k rows,
    # no intercept slot (group-mean recovery never uses one). v2.2 Log
    # wiring: whichever of those raw values belong to a Log-transformed
    # column are Ln_Positive-transformed before the call, per the
    # per-constructed-column flag from Constructed_Column_Transforms()
    # (TRANSPOSE'd to match the AK band's column-vector shape).
    assert (
        f"LET(raw,TAKE($AK${_PRED_INPUT_FIRST_ROW}:"
        f"$AK${_PRED_INPUT_LAST_ROW},COLUMNS(Predictor_Columns()))," in formula
    )
    assert "trn,TRANSPOSE(Constructed_Column_Transforms())," in formula
    assert 'pred_input,IF(trn="Log",Ln_Positive(raw),raw),' in formula
    assert (
        "Group_Prediction_Interval(Predictor_Columns(),Response_Column(),pred_input,"
        "Prediction_Group_Column(),$AK$12,"
        "Sample_Include(),alpha,Fit_Context())"
    ) in formula
    assert "Intercept_Only_Point()" in formula
    # The zero-predictor closed form also splits mean-CI from new-obs-PI now.
    assert "se_mean,Intercept_Only_S()/SQRT(Intercept_Only_N())" in formula
    assert "se_new,Intercept_Only_S()*SQRT(1+1/Intercept_Only_N())" in formula


def test_prediction_interval_writes_fe_group_selector_and_readouts() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_prediction_interval(_as_xw_sheet(sheet))

    assert sheet.cell(12, _C_AJ).value == "FE Group"
    fe_group = sheet.cell(12, _C_AK).api.Formula2
    assert fe_group == (
        "=INDEX(SORT(UNIQUE(FILTER(Prediction_Group_Column(),Sample_Include()))),1,1)"
    )
    conditions = sheet.range("$AK$12").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == [
        "=ISNA(MATCH($AK$12,Prediction_Group_Column(),0))"
    ]
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    assert sheet.cell(13, _C_AJ).value == "Group Mean (y)"
    assert sheet.cell(13, _C_AK).api.Formula2 == (
        "=Group_Mean_At(Response_Column(),Prediction_Group_Column(),"
        "$AK$12,Sample_Include())"
    )
    assert sheet.cell(14, _C_AJ).value == "Group Count"
    assert sheet.cell(14, _C_AK).api.Formula2 == (
        "=Group_Count_At(Prediction_Group_Column(),$AK$12,Sample_Include())"
    )


def test_materialization_zone_materializes_model_context() -> None:
    # The §4b band materializes the context ONCE and exposes it through a
    # sheet-scoped reader name (Fit_Context), so the ~30 engine call sites that
    # pass Fit_Context() all read the one materialized block instead of
    # rebuilding the context ~30 times. Model_Context stays the workbook-scoped
    # constructor (the omitted-[Context] default); the split keeps it unshadowed.
    # Plain cell writes + one Names.Add only (no COM chart API), so this runs
    # headless via RecordingSheet. ``closures=()`` skips the catalog disk load the
    # default path would do (the writes don't use it).
    sheet = RecordingSheet(name="Regression")

    _write_materialization_zone(_as_xw_sheet(sheet), closures=())

    ctx_col = col_letter(_C_MODEL_CONTEXT)

    # Section headings at row 1, styled as section headings. The Model Context
    # heading names the block from its LABEL column and the value cell beside
    # it carries the fill only, so the heading reads as one banner across the
    # two-column block.
    assert sheet.cell(1, _C_MODEL_CONTEXT_LABEL).value == "MODEL CONTEXT"
    assert sheet.cell(1, _C_MODEL_CONTEXT_LABEL).color == HEADER_COLOR
    assert sheet.cell(1, _C_MODEL_CONTEXT).value == ""
    assert sheet.cell(1, _C_MODEL_CONTEXT).color == HEADER_COLOR
    assert sheet.cell(1, _C_SAMPLE_INCLUDE_MATERIALIZED).value == "Sample Include"

    # Each context element is its OWN cell, labelled in the column to its left.
    # Not a VSTACK spill: a spill is a single dependency node that Excel vacates
    # and re-spills on every spec-block edit, and while it is vacated the fixed
    # range behind Fit_Context is blank — every engine reading it would see a
    # torn context. Independent cells never get vacated.
    assert [element.contract_name for element in _MODEL_CONTEXT_ELEMENTS] == [
        "Has_Intercept",
        "DF_Absorbed",
        "Response_Transform",
        "Predictor_Transform",
    ]
    formulas = {}
    for offset, element in enumerate(_MODEL_CONTEXT_ELEMENTS):
        row = _MATERIALIZATION_FIRST_ROW + offset
        assert sheet.cell(row, _C_MODEL_CONTEXT_LABEL).value == element.label
        formula = sheet.cell(row, _C_MODEL_CONTEXT).api.Formula2
        assert formula == f"={element.formula}"
        assert "VSTACK" not in formula
        formulas[element.contract_name] = formula

    # Elements 1-2 (Allow_Intercept toggle, the Absorbed_Degrees_Of_Freedom()
    # closure) feed today's engines; elements 3-4 are the spec-block transform
    # summaries that have no reader until v3.3 but land now so the row order is
    # fixed. They must NOT be the old '="None"' placeholders.
    assert formulas["Has_Intercept"] == "=Allow_Intercept"
    assert formulas["DF_Absorbed"] == "=Absorbed_Degrees_Of_Freedom()"
    # Response_Transform: INDEX/XMATCH over the spec rows for the response
    # role, mirroring _RESPONSE_LOG_FORMULA, with an IFERROR->"None".
    response = formulas["Response_Transform"]
    assert response != '="None"'
    assert 'XMATCH("Response (y)",TAKE(Spec_Role,COLUMNS(Source_Data)))' in response
    assert "INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data))" in response
    # Predictor_Transform: a LET summarising the INCLUDED CONTINUOUS predictors
    # as None/Log/Mixed, masked to Continuous so a Categorical dummy can't
    # manufacture a false "Mixed".
    predictor = formulas["Predictor_Transform"]
    assert predictor != '="None"'
    assert "LET(n_c,COLUMNS(Source_Data)," in predictor
    assert '(rl="Predictor (x)")*(inc=TRUE)*(typ="Continuous")' in predictor
    assert 'IF(nL=0,"None",IF(nN=0,"Log","Mixed")' in predictor

    # The sheet-scoped reader Fit_Context reads the FIXED materialized range —
    # no spill operator inside the LAMBDA RefersTo, so the dynamic-array-in-a-
    # name question is sidestepped and the height is a structural constant. The
    # legacy "Model_Context" sheet name is dropped (a pre-split build would have
    # left one) so a rebuild never leaves a shadow alongside Fit_Context.
    assert _MODEL_CONTEXT_LAST_ROW == _MATERIALIZATION_FIRST_ROW + _MODEL_CONTEXT_ROWS - 1
    assert sheet.api.Names.by_short_name("Fit_Context").RefersTo == (
        f"=LAMBDA('Regression'!${ctx_col}${_MATERIALIZATION_FIRST_ROW}"
        f":${ctx_col}${_MODEL_CONTEXT_LAST_ROW})"
    )
    assert "Model_Context" not in {
        item.Name.split("!", 1)[-1] for item in sheet.api.Names.items
    }

    # Health check one row under the block: the height is the build-time
    # invariant, and — the part decomposition made worth checking — every
    # element resolved. With independent cells a broken spec name errors in one
    # of them and leaves the other three looking fine.
    assert _ROW_MODEL_CONTEXT_CHECK == _MODEL_CONTEXT_LAST_ROW + 1
    assert sheet.cell(_ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT_LABEL).value == "Context OK"
    assert sheet.cell(_ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT).api.Formula2 == (
        f"=AND(ROWS(Fit_Context())={_MODEL_CONTEXT_ROWS},"
        "SUMPRODUCT(--ISERROR(Fit_Context()))=0)"
    )

    # The block is of FIXED size, so it is boxed like every other fixed-size
    # block on this sheet — heading row through the health-check row, both
    # columns.
    box = sheet.range(
        (1, _C_MODEL_CONTEXT_LABEL), (_ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT)
    ).api
    assert set(box._borders) == {7, 8, 9, 10}

    # Sample_Include is SURFACED at its final §4b position: a header row and a
    # full-height spill of the live closure. Surfacing the value is not the
    # same as rewiring the reader — promoting the closure to a thunk over this
    # spill stays deferred (Excel-verified, not blind), so nothing on the sheet
    # reads this column and Sample_Include() is still evaluated per call site.
    assert (
        sheet.cell(_MATERIALIZATION_HEADER_ROW, _C_SAMPLE_INCLUDE_MATERIALIZED).value
        == _SAMPLE_INCLUDE_HEADER
    )
    assert _formula(
        sheet, _MATERIALIZATION_SPILL_ROW, _C_SAMPLE_INCLUDE_MATERIALIZED
    ) == "=Sample_Include()"

    # The zone sits past the chart footprint with a structural gutter after
    # the charts; Model Context is a label/value pair and Sample_Include is one
    # column, separated by width-2 ungrouped gutters.
    assert _C_GUTTER_AFTER_CHARTS < _C_MODEL_CONTEXT_LABEL < _C_MODEL_CONTEXT
    assert _C_MODEL_CONTEXT < _C_GUTTER_AFTER_CONTEXT
    assert _C_MODEL_CONTEXT_LABEL - _C_GUTTER_AFTER_CHARTS == 1
    assert _C_GUTTER_AFTER_CONTEXT < _C_SAMPLE_INCLUDE_MATERIALIZED < _C_GUTTER_AFTER_SAMPLE_INCLUDE
    assert _C_GUTTER_AFTER_CONTEXT - _C_MODEL_CONTEXT == 1
    assert _C_SAMPLE_INCLUDE_MATERIALIZED - _C_GUTTER_AFTER_CONTEXT == 1

    # The terminal zone: the Constructed Design Matrix, spilled at its final
    # position. §4b's ordering rule is that the band runs in increasing width
    # and TERMINATES here — nothing may ever be placed to its right, because
    # its width is unbounded and one dropdown away.
    assert sheet.cell(1, _C_DESIGN_MATRIX).value == "Constructed Design Matrix"
    assert sheet.cell(1, _C_DESIGN_MATRIX).color == HEADER_COLOR
    assert _formula(
        sheet, _MATERIALIZATION_SPILL_ROW, _C_DESIGN_MATRIX
    ) == "=Design_Columns()"
    assert _C_DESIGN_MATRIX - _C_GUTTER_AFTER_SAMPLE_INCLUDE == 1

    # The header row is SPLIT across two cells, because Design_Columns() is one
    # column wider than Constructed_Column_Names() whenever the intercept is
    # on: the constructor prepends the ones column, and the names closure
    # describes the constructed predictor columns only. The anchor cell names
    # that ones column; the names spill starts in the column beside it.
    assert _formula(sheet, _MATERIALIZATION_HEADER_ROW, _C_DESIGN_MATRIX) == (
        _DESIGN_MATRIX_INTERCEPT_HEADER
    )
    assert "Allow_Intercept" in _DESIGN_MATRIX_INTERCEPT_HEADER
    assert _C_DESIGN_MATRIX_NAMES - _C_DESIGN_MATRIX == 1
    assert _formula(
        sheet, _MATERIALIZATION_HEADER_ROW, _C_DESIGN_MATRIX_NAMES
    ) == "=Constructed_Column_Names()"

    # Both data-dependent zones head on one row and spill from the next, so
    # the band reads across — the mask value beside its design-matrix row.
    assert _MATERIALIZATION_FIRST_ROW == 2
    assert _MATERIALIZATION_HEADER_ROW == _MATERIALIZATION_FIRST_ROW
    assert _MATERIALIZATION_SPILL_ROW == _MATERIALIZATION_HEADER_ROW + 1
    # The fixed-height Model Context block needs no header row, so it still
    # starts its elements on the band's first row — its last element row and
    # health check sit below the spills' header row, in their own columns.
    assert _MODEL_CONTEXT_LAST_ROW >= _MATERIALIZATION_SPILL_ROW


def test_materialization_zones_ship_collapsed() -> None:
    # All three §4b content zones ship collapsed so the materialization band
    # stays out of the way until explicitly expanded. The terminal zone is the
    # worst offender because its width is one dropdown away from hundreds of
    # columns, but the two bounded zones are part of the same secondary band.
    #
    # One outline group per ZONE, not per column: Model Context is grouped as
    # the label/value pair so it collapses as a unit; grouping the value column
    # alone would leave its labels stranded beside a collapsed column.
    sheet = RecordingSheet(name="Regression")

    _write_materialization_zone(_as_xw_sheet(sheet), closures=())

    matrix_band = (
        f"{col_letter(_C_DESIGN_MATRIX)}:"
        f"{col_letter(_C_DESIGN_MATRIX + _DESIGN_MATRIX_GROUPED_COLUMNS - 1)}"
    )
    assert sheet.column_groups == [
        f"{col_letter(_C_MODEL_CONTEXT_LABEL)}:{col_letter(_C_MODEL_CONTEXT)}",
        f"{col_letter(_C_SAMPLE_INCLUDE_MATERIALIZED)}:"
        f"{col_letter(_C_SAMPLE_INCLUDE_MATERIALIZED)}",
        matrix_band,
    ]
    assert sheet.column_show_detail == {
        f"{col_letter(_C_MODEL_CONTEXT_LABEL)}:{col_letter(_C_MODEL_CONTEXT)}": False,
        f"{col_letter(_C_SAMPLE_INCLUDE_MATERIALIZED)}:"
        f"{col_letter(_C_SAMPLE_INCLUDE_MATERIALIZED)}": False,
        matrix_band: False,
    }
    # The gutters stay OUT of every group, or the zones would fuse into one
    # outline and lose independent collapse.
    for gutter in (
        _C_GUTTER_AFTER_CHARTS,
        _C_GUTTER_AFTER_CONTEXT,
        _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    ):
        letter = col_letter(gutter)
        assert not any(letter in band.split(":") for band in sheet.column_groups)


def test_materialization_zone_widths_use_named_constants() -> None:
    """§4b emitted widths stay aligned with the writer's width constants."""
    sheet = RecordingSheet(name="Regression")

    _write_materialization_zone(_as_xw_sheet(sheet), closures=())

    assert (
        sheet.range(f"{col_letter(_C_MODEL_CONTEXT_LABEL)}:{col_letter(_C_MODEL_CONTEXT_LABEL)}").column_width
        == _MODEL_CONTEXT_LABEL_WIDTH
    )
    assert (
        sheet.range(f"{col_letter(_C_MODEL_CONTEXT)}:{col_letter(_C_MODEL_CONTEXT)}").column_width
        == _MODEL_CONTEXT_VALUE_WIDTH
    )
    assert (
        sheet.range(
            f"{col_letter(_C_SAMPLE_INCLUDE_MATERIALIZED)}:{col_letter(_C_SAMPLE_INCLUDE_MATERIALIZED)}"
        ).column_width
        == _SAMPLE_INCLUDE_MATERIALIZED_WIDTH
    )

    matrix_band = (
        f"{col_letter(_C_DESIGN_MATRIX)}:"
        f"{col_letter(_C_DESIGN_MATRIX + _DESIGN_MATRIX_GROUPED_COLUMNS - 1)}"
    )
    assert sheet.range(matrix_band).column_width == _DESIGN_MATRIX_GROUPED_WIDTH


def test_width_guard_reads_the_spec_not_the_constructed_matrix() -> None:
    # The whole point of the pre-flight guard: a matrix too wide to fit
    # cannot be built in order to be measured. Both thresholds therefore
    # read the spec block's own Design Columns audit total, never
    # COLUMNS(Design_Columns()).
    sheet = RecordingSheet(name="Regression")

    _write_design_matrix_width_guard(_as_xw_sheet(sheet))

    total = _formula(sheet, 1, _C_SPEC_DESIGN_COLUMNS)
    assert total == (
        "=SUM(TAKE(Spec_Design_Columns,COLUMNS(Source_Data)))+N(Allow_Intercept)"
    )
    assert sheet.cell(1, _C_SPEC_INTERACTION_OPERATION).value == "Σ Design Columns"

    status = _formula(sheet, 2, _C_SPEC_INTERACTION_TERM)
    assert "Design_Columns()" not in status
    assert "$O$1" in status                      # reads the audit total
    assert "n,ROWS(Source_Data)" in status       # cell count needs n as well
    assert f"IF(k>{_DESIGN_MATRIX_MAX_COLUMNS}," in status
    assert f"OR(k>{_DESIGN_MATRIX_SOFT_COLUMNS},n*k>{_DESIGN_MATRIX_SOFT_CELLS})" in status
    assert status.endswith('"")))')

    # The hard limit is DERIVED from the layout constants — it is exactly the
    # column budget left of Excel's right edge once the materialization band
    # is placed, so moving a zone moves the limit with it.
    assert _DESIGN_MATRIX_MAX_COLUMNS == 16384 - _C_DESIGN_MATRIX + 1

    # Red outranks yellow via StopIfTrue, and both the status line and the
    # total carry the flag — the total is the number a user actually reads.
    for cell in ("$M$2", "$O$1"):
        rules = sheet.range(cell).api.FormatConditions.items
        assert [c.Formula1 for c in rules] == [
            '=ISNUMBER(SEARCH("ERROR",$M$2))',
            '=ISNUMBER(SEARCH("WARNING",$M$2))',
        ], cell
        assert rules[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
        assert rules[0].StopIfTrue is True
        assert rules[1].Interior.Color == excel_color(CF_YELLOW_FILL)


def test_prediction_prefills_index_the_single_training_mean_spill() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_prediction_inputs(_as_xw_sheet(sheet))

    # The Training Mean spill is the ONE Predictor_Columns() evaluation for the whole
    # prefill band; it owns column AI downward so it can never collide with
    # another spill when the source data or spec changes.
    assert sheet.cell(17, _C_AL).value == "Training Mean"
    means = _formula(sheet, _PRED_INPUT_FIRST_ROW, _C_AL)
    # v2.2 Log wiring: a Log-transformed column's mean is EXP'd back to
    # input space (the geometric mean) so the AH prefill — which just
    # INDEXes this spill — doesn't get double-logged when the row-3
    # prediction formula applies Ln_Positive to it.
    assert means == (
        "=IFERROR(TRANSPOSE(LET(m,BYCOL(FILTER(Predictor_Columns(),Sample_Include()),"
        "LAMBDA(c,AVERAGE(c))),t,Constructed_Column_Transforms(),"
        'IF(t="Log",EXP(m),m))),"")'
    )

    # Perf tripwire: Predictor_Columns() is a full design-matrix construction on every
    # call, so no prefill cell may invoke it — 50 cells × 2 calls made the
    # workbook's first full calculation take ~20 minutes.
    for row in (_PRED_INPUT_FIRST_ROW, _PRED_INPUT_LAST_ROW):
        prefill = _formula(sheet, row, _C_AK)
        assert "Predictor_Columns()" not in prefill
        assert f"INDEX($AL${_PRED_INPUT_FIRST_ROW}#" in prefill
        assert f"IFERROR(ROWS($AL${_PRED_INPUT_FIRST_ROW}#),0)" in prefill


def test_write_coefficients_adds_intercept_only_closed_form_branch() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_coefficients(_as_xw_sheet(sheet))

    label_formula = _formula(sheet, 21, _C_AA)
    assert label_formula.startswith("=IF(Zero_Predictors_Selected(),")
    assert 'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),"Intercept",NA())' in label_formula
    # Level-qualified names come from the constructor twin (a row vector).
    assert 'VSTACK("Intercept",TRANSPOSE(Constructed_Column_Names()))' in label_formula

    coefficient_formula = _formula(sheet, 21, _C_AB)
    assert "IF(AND(Allow_Intercept,Intercept_Only_N()>=1),Intercept_Only_Point(),NA())" in coefficient_formula
    assert "Coefficients(Design_Columns(),Design_Response(),Sample_Include())" in coefficient_formula

    se_formula = _formula(sheet, 21, _C_AC)
    assert "IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_SE(),NA())" in se_formula

    beta_formula = _formula(sheet, 21, _C_AH)
    assert 'IF(Allow_Intercept,"",NA())' in beta_formula


def test_regression_outputs_header_writes_predicted_variable_readout() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_regression_outputs_header(_as_xw_sheet(sheet))

    assert sheet.cell(2, _C_AE).value == "Predicted Variable"
    assert sheet.cell(2, _C_AE).api.Font.Bold is True
    assert sheet.cell(2, _C_AE).color == HEADER_COLOR
    # Derived response name — the header of the Role=Response spec row.
    readout = sheet.cell(2, _C_AF).api.Formula2
    assert readout is not None
    assert readout.startswith("=LET(n_c,COLUMNS(Source_Data),")
    assert 'XMATCH("Response (y)"' in readout
    # v2.2 Log wiring: relabelled Ln(name) when the Response row's
    # Transform is Log, so the readout never lies about what's fitted.
    assert 'IF(INDEX(TAKE(Spec_Transform,n_c),p)="Log","Ln("&h&")",h)' in readout
    assert sheet.cell(2, _C_AF).api.Font.Bold is True
    assert sheet.cell(2, _C_AF).color == HEADER_COLOR


def test_write_residuals_writes_row_labels_and_diagnostics() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_residuals(_as_xw_sheet(sheet))

    # Static header; Row_Labels() supplies its own per-row content and
    # no-Identifier fallback, so only an all-FALSE mask is absorbed here.
    assert sheet.cell(2, _C_AN).value == "Observation"
    assert sheet.cell(3, _C_AN).api.Formula2 == (
        "=IFERROR(FILTER(Row_Labels(),Sample_Include()),NA())"
    )
    # The diagnostics columns shift one slot right of the identifiers column.
    # Response-scale headers (Y, Predicted Y, Residuals, PRESS Residual) are
    # nested IF conditionals on both the FE-count gate (a suffix naming the
    # active Fixed Effects variable) and the Response row's Transform (a
    # "(Log)" suffix, v2.2) — plain label with neither, both suffixes when
    # both apply. Y gets its own "Deviation from <FE> Avg." wording (the
    # Within transform only subtracts the group mean, it never divides by a
    # standard deviation, so "Within" alone would be accurate but vaguer,
    # and a "St Devs" framing would be outright wrong).
    header_formula = sheet.cell(2, _C_AO).api.Formula2
    assert header_formula is not None
    assert '"Y"&" (Deviation from "' in header_formula
    assert '" Avg., Log)"' in header_formula
    assert '" Avg.)"' in header_formula
    assert '"Y (Log)"' in header_formula
    assert header_formula.endswith(',"Y")))')
    # Lock in the gates themselves — the same FE-role count detector the
    # DW/BFN trigger cells use, and the Response row's own Transform cell —
    # so an always-true/false or unrelated condition would not satisfy this.
    assert "Spec_Role" in header_formula
    assert '"Fixed Effects"' in header_formula
    assert "Spec_Transform" in header_formula
    assert '"Response (y)"' in header_formula
    # The FE variable's own name, pulled the same way the spec feedback
    # block's "FE Variable" cell does.
    assert "Header_Names" in header_formula
    assert header_formula.startswith("=IF(AND(SUMPRODUCT(N(TAKE(Spec_Role,")

    # A dimensionless diagnostic (Hat Diagonal) gets ONLY the FE suffix —
    # it isn't in response units either way, so a Log-transformed response
    # must not relabel it. Its suffix names the FE variable ("Within
    # Country"), not the bare "(Within)" token.
    hat_header = sheet.cell(2, _C_AR).api.Formula2
    assert hat_header == (
        '=IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))>0,'
        '"Hat Diagonal"&" (Within "&IFERROR(INDEX(TOROW(Header_Names),'
        'XMATCH("Fixed Effects",TAKE(Spec_Role,COLUMNS(Source_Data)))),"FE")&")",'
        '"Hat Diagonal")'
    )
    assert "Spec_Transform" not in hat_header
    assert sheet.cell(3, _C_AO).api.Formula2 == (
        "=Dependent_Variable(Design_Response(),Sample_Include())"
    )
    assert sheet.cell(3, _C_AQ).api.Formula2 == (
        "=Residuals(Design_Columns(),Design_Response(),Sample_Include())"
    )
    assert sheet.cell(3, _C_AR).api.Formula2 == (
        "=Hat_Diagonal(Design_Columns(),Sample_Include())"
    )
    assert sheet.cell(3, _C_AX).api.Formula2 == (
        "=LOOCV_Residual(Design_Columns(),Design_Response(),Sample_Include())"
    )
    # Cook's Distance (Flagged): NA()'d below both influence cutoffs, so the
    # Cook's Distance chart's overlay series only labels flagged points.
    assert sheet.cell(3, _C_AY).api.Formula2 == (
        "=IF(AT3#>MIN(4/$AB$8,0.9),AT3#,NA())"
    )


def test_write_unit_space_block_writes_section_input_and_three_gof_cells() -> None:
    """v3.3 unit-space block at AG3:AH9 — section heading, Method input,
    Smearing Factor, R²/Adj R²/RMSE in original units, Response Space readout.
    """
    sheet = RecordingSheet(name="Regression")
    _write_unit_space_block(_as_xw_sheet(sheet))

    # Section heading on row 3 spans AG3 (a single value cell here — the
    # section_heading helper writes it as a label, not a merge).
    assert sheet.cell(3, _C_AG).value == "UNIT-SPACE FIT"
    # Row 4: Back-Transform Method input. A LITERAL "Duan", never a formula —
    # the cell is an input, and a formula reading its own $AH$4 address is a
    # circular reference that Excel resolves to 0 with iterative calculation
    # off, feeding an unrecognised method to every consumer below. The
    # Duan/Naive constraint lives in the list validation, not in the cell.
    assert sheet.cell(4, _C_AG).value == "Back-Transform"
    assert sheet.cell(4, _C_AH).value == "Duan"
    method_formula = sheet.cell(4, _C_AH).api.Formula2
    assert method_formula is None or not str(method_formula).startswith("="), (
        "AH4 must hold a literal default, not a self-referential formula"
    )
    # The dropdown's item list. Excel's xlValidateList takes the items as a
    # bare comma-separated string; the quotes VBA examples show around it are
    # that language's string delimiters. Passing them through COM made them
    # literal and the dropdown offered `"Duan` and `Naive"` — both accepted by
    # the validation and rejected by every consumer, since neither matches a
    # recognised method. Assert the parsed items, not the raw string, so the
    # check is about what the user is offered.
    rules = sheet.cell(4, _C_AH).api.Validation.rules
    assert len(rules) == 1, "AH4 must carry exactly one validation rule"
    items = str(rules[0]["Formula1"]).split(",")
    assert items == list(_BACK_TRANSFORM_METHODS), (
        f"dropdown offers {items!r}, not the two supported methods"
    )
    assert '"' not in str(rules[0]["Formula1"]), (
        "quote characters in Formula1 become part of the list items"
    )
    assert sheet.cell(4, _C_AH).api.Validation.IgnoreBlank is False
    # Rows 5–8: the four GoF statistics; each formula lifts the smearing
    # factor's own X/Y/Include/Context wiring rather than re-stating it.
    for row, label in [
        (5, "Smearing Factor"),
        (6, "R Square (Unit)"),
        (7, "Adj R Square (Unit)"),
        (8, "RMSE (Unit)"),
    ]:
        assert sheet.cell(row, _C_AG).value == label, row
    smearing_formula = sheet.cell(5, _C_AH).api.Formula2
    assert smearing_formula == (
        "=Smearing_Factor(Design_Columns(),Design_Response(),"
        "Sample_Include(),Fit_Context())"
    )
    # The three GoF cells pass the Method input back through the catalog
    # entry-point so the Duan/Naive toggle actually changes the readout.
    for row in (6, 7, 8):
        formula = sheet.cell(row, _C_AH).api.Formula2
        assert formula is not None
        assert "Fit_Context()" in formula
        assert _A_BACK_TRANSFORM_METHOD in formula  # the $AH$4 anchor
    # Row 9: Response Space readout — "Original units (back-transformed)"
    # under Log, "Same as fit space" otherwise.
    assert sheet.cell(9, _C_AG).value == "Response Space"
    response_space_formula = sheet.cell(9, _C_AH).api.Formula2
    assert "Context_Response_Transform(Fit_Context())" in response_space_formula
    assert '"Original units (back-transformed)"' in response_space_formula
    assert '"Same as fit space"' in response_space_formula


def test_write_residuals_appends_unit_space_columns_az_ba() -> None:
    """v3.3: Predicted Y (Original Units) and Residual (Original Units) on
    columns AZ and BA. Plain labels (no (Log) / (Within) suffix), back-
    transformation routed through Unit_Space_Predictions / Unit_Space_Residuals
    with the AH4 Method input.
    """
    sheet = RecordingSheet(name="Regression")
    _write_residuals(_as_xw_sheet(sheet))

    # Static headers — no Log/Within conditional; the unit-space columns
    # are by construction in original units, so the conditional suffix
    # logic that decorates the AO–AY columns does not apply here.
    assert sheet.cell(2, _C_AZ).value == "Predicted Y (Original Units)"
    assert sheet.cell(2, _C_BA).value == "Residual (Original Units)"
    # Neither header leaks the With_FE/Log conditional fragments.
    for col in (_C_AZ, _C_BA):
        header = sheet.cell(2, col).value
        assert isinstance(header, str)
        assert "(Log)" not in header
        assert "(Within" not in header
    # Row 3 formulas call the catalog Unit_Space_* functions and feed the
    # Method toggle from the unit-space block.
    az_formula = sheet.cell(3, _C_AZ).api.Formula2
    ba_formula = sheet.cell(3, _C_BA).api.Formula2
    assert az_formula == (
        "=Unit_Space_Predictions(Design_Columns(),Design_Response(),"
        "Response_Column(),Sample_Include(),Fit_Context(),"
        f"{_A_BACK_TRANSFORM_METHOD})"
    )
    assert ba_formula == (
        "=Unit_Space_Residuals(Design_Columns(),Design_Response(),"
        "Response_Column(),Sample_Include(),Fit_Context(),"
        f"{_A_BACK_TRANSFORM_METHOD})"
    )


def test_column_widths_cover_every_zone_column_and_no_gap_column() -> None:
    """Widths are keyed on the layout constants, one per content column.

    The table used to be keyed on literal letters, and the layout break that
    shifted every zone right of the spec block three columns over left the
    keys behind: the Predictor Summary's name column was sized for a stats
    value, the Regression Outputs' diagnostics labels got the width meant for
    its values, and the whole Prediction Outputs zone fell off the table and
    rendered at Excel's default width. Nothing failed — the build sized a
    different column, which is exactly the silent-wrong-answer mode the `_C_*`
    constants exist to prevent. Pin the coverage so the next shift fails here.
    """
    widths = dict(_COLUMN_WIDTHS)
    assert len(widths) == len(_COLUMN_WIDTHS), "no column may be sized twice"

    # Every content column of every zone past the spec block is sized exactly
    # once. (The spec block A–O owns its own widths in
    # write_sheet_model_construction; column I is the one override here.)
    content = {
        col for first, last in _ZONES[1:] for col in range(first, last + 1)
    } | {_C_P, _C_Q}  # zone 1's spec-feedback columns; A–O is the spec block
    assert content <= set(widths)
    # ...and nothing else is sized except the column-I overlay and the BB
    # post-zone gutter.
    assert set(widths) - content == {_C_SPEC_SEQUENCE_PERIOD, _C_BB}
    # Gap columns are sized by the _GAP_COLUMNS loop (width 2) and must not
    # appear here, or a stale entry would silently widen a separator.
    assert not set(widths) & set(_GAP_COLUMNS)

    # The columns whose width the stale letter keys landed on the wrong side
    # of. Each is the widest thing in its zone: a label or name column beside
    # narrow numeric ones.
    assert widths[_C_S] == 24           # constructed column names
    assert widths[_C_X] == widths[_C_Y] == 9   # GVIF / Tolerance — plain stats
    assert widths[_C_AD] == 24          # "BFN Panel Durbin-Watson" (23)
    assert widths[_C_AJ] == 24          # "PREDICTION INTERVAL" labels + names spill
    assert widths[_C_AK] == 16          # interval + prediction input values
    assert widths[_C_AL] == 14          # Original Units / Training Mean
    assert widths[_C_SPEC_SEQUENCE_PERIOD] == 38  # the I2 Verdict switch


def test_regression_outputs_header_no_longer_holds_the_model_formula() -> None:
    """The Model Formula readout left AA2/AB2.

    Row 2 of this zone wraps and then AutoFits, so the longest string on the
    sheet sitting in a 12-wide column set the height of the whole header row.
    The caption moved to the §4b materialization band; the header keeps only
    the zone heading and the Predicted Variable readout.
    """
    sheet = RecordingSheet(name="Regression")
    _write_regression_outputs_header(_as_xw_sheet(sheet))

    assert sheet.cell(1, _C_AA).value == "REGRESSION OUTPUTS"
    assert sheet.cell(2, _C_AA).value is None
    assert sheet.cell(2, _C_AB).api.Formula2 is None
    # The Predicted Variable readout stays put.
    assert sheet.cell(2, _C_AE).value == "Predicted Variable"


def test_materialization_zone_writes_the_model_formula_readout() -> None:
    """The Model Formula caption, on row 1 of the design-matrix zone.

    Right of that zone's own heading — header two columns over, readout one
    column past the header — holding a call to the sheet-scoped
    `Model_Formula()` closure rather than a 300-character expression, with
    wrap explicitly OFF so the string overflows across an empty row 1 instead
    of dictating a row height the way it did at AB2.
    """
    sheet = RecordingSheet(name="Regression")
    _write_materialization_zone(_as_xw_sheet(sheet), closures=())

    assert _ROW_MODEL_FORMULA == 1
    assert _C_MODEL_FORMULA_LABEL == _C_DESIGN_MATRIX + 2
    assert _C_MODEL_FORMULA == _C_MODEL_FORMULA_LABEL + 1
    assert sheet.cell(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL).value == "Model Formula"
    assert sheet.cell(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL).api.Font.Bold is True
    assert sheet.cell(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL).color == HEADER_COLOR
    assert _formula(sheet, _ROW_MODEL_FORMULA, _C_MODEL_FORMULA) == "=Model_Formula()"
    assert sheet.range(
        (_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL),
        (_ROW_MODEL_FORMULA, _C_MODEL_FORMULA),
    ).api.WrapText is False


def test_model_formula_readout_is_clear_of_the_design_matrix_body() -> None:
    """Row 1 is the one row in that zone the matrix can never reach.

    The design matrix's names spill on `_MATERIALIZATION_HEADER_ROW` and its
    values on `_MATERIALIZATION_SPILL_ROW`, both growing RIGHTWARD from the
    zone anchor — so a caption on row 1 is not displaced by an ordinary
    modelling choice, and the §4b rule that nothing may sit to the RIGHT of
    this zone is not in play: the caption is above the body, inside the
    zone's own columns.

    The header/readout gap is load-bearing too. `"Model Formula"` is wider
    than one 12-wide design-matrix column, so a readout immediately beside
    the header would clip it; the clearance is what lets both render. The
    assertion below is the invariant (readout strictly right of header), not
    the current width of the gap.
    """
    assert _ROW_MODEL_FORMULA < _MATERIALIZATION_HEADER_ROW
    assert _MATERIALIZATION_HEADER_ROW < _MATERIALIZATION_SPILL_ROW
    # Right of the zone heading, and clear of the split header cells the
    # matrix itself uses on the row below.
    assert _C_DESIGN_MATRIX < _C_MODEL_FORMULA_LABEL < _C_MODEL_FORMULA
    assert _C_MODEL_FORMULA_LABEL > _C_DESIGN_MATRIX_NAMES
    assert _C_MODEL_FORMULA - _C_MODEL_FORMULA_LABEL >= 1


def test_model_formula_closure_assembles_the_spec_derived_caption() -> None:
    """The catalog body that replaced the inline AB2 expression.

    Sheet-scoped (it reads this sheet's Spec_* names and the
    Constructed_Column_Names() closure beside them), and built from the same
    four pieces the cell formula used to concatenate inline.
    """
    document = load_catalog_document(ROOT_DIR / "lambda_functions.json")
    closure = next(
        fn for fn in document.functions_for_sheet("Regression")
        if fn.name == "Model_Formula"
    )
    body = closure.formula_display.replace(" ", "").replace("\n", "")

    assert closure.scope == "Regression"
    assert closure.arguments == ()
    # Response side — the XMATCH-on-Role lookup, wrapped as "Ln(name)" on Log.
    assert 'XMATCH("Response(y)",TAKE(Spec_Role,n_c))' in body
    assert '"Ln("&h&")"' in body
    assert 'INDEX(TOROW(Header_Names),p)' in body
    assert '"~"' in body  # the " ~ " separator, whitespace-stripped
    # RHS — always "1 + " with intercept, "0 + " without.
    assert "IF(Allow_Intercept," in body
    # Predictor list — TEXTJOIN over the constructed column names.
    assert "TEXTJOIN(" in body
    assert "Constructed_Column_Names()" in body
    # FE suffix — gated on the Fixed Effects count, names the active variable.
    assert 'SUMPRODUCT(N(TAKE(Spec_Role,n_c)="FixedEffects"))' in body
    assert '"FE"' in body  # the IFERROR fallback when no FE row is declared
    # A sheet-scoped body must never qualify a sheet: a workbook with several
    # Regression-shaped sheets would then read one sheet's spec everywhere.
    assert "'Regression'!" not in closure.formula_display


def test_setup_local_names_registers_comparison_anchor_headline_and_formula() -> None:
    """v3.3: Comparison_Anchor, Comparison_Headline_GoF, Comparison_Model_Formula
    are sheet-scoped names that the v3.4 Model Comparison sheet reads from.
    Comparison_Anchor → AF2 (the response-name readout), Comparison_Headline_GoF
    → AH6:AH8 (the three unit-space GoF statistics), Comparison_Model_Formula
    → the Model Formula readout in the §4b band (it moved off AB2; v3.4 reads
    the NAME, which is why the move costs its consumer nothing).
    """
    sheet = RecordingSheet(name="Regression")
    _setup_regression_names(_as_xw_sheet(sheet), closures=())

    assert sheet.api.Names.by_short_name("Comparison_Anchor").RefersTo == (
        "='Regression'!$AF$2"
    )
    assert sheet.api.Names.by_short_name("Comparison_Headline_GoF").RefersTo == (
        "='Regression'!$AH$6:$AH$8"
    )
    assert sheet.api.Names.by_short_name("Comparison_Model_Formula").RefersTo == (
        f"='Regression'!${col_letter(_C_MODEL_FORMULA)}${_ROW_MODEL_FORMULA}"
    )


def test_diagnostics_durbin_watson_is_gated_on_a_sequence_flag() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_diagnostics(_as_xw_sheet(sheet))

    assert sheet.cell(11, _C_AD).value == "Durbin-Watson"
    dw_formula = cast(str, sheet.cell(11, _C_AE).api.Formula2)
    # All off-spec states show an explicit text token, never NA() or "".
    assert '"n/a — requires Sequence"' in dw_formula      # zero flags
    assert '"n/a — multiple Sequence flags"' in dw_formula  # two-plus flags
    assert '"n/a — FE active"' in dw_formula              # panel: BFN takes over
    assert "NA()" not in dw_formula
    # Gate keys on the Sequence-flag and FE-variable counts over the live spec
    # rows, evaluated once via LET; requires EXACTLY one flag and NO Fixed
    # Effects variable before computing the single-series statistic.
    assert "SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))" in dw_formula
    assert (
        'SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))'
    ) in dw_formula
    assert "IF(seq_flags=0," in dw_formula
    assert "IF(seq_flags>1," in dw_formula
    assert "IF(fe_vars>0," in dw_formula
    # With exactly one flag, DW is computed along the declared axis, not row order.
    assert (
        "Durbin_Watson_By(Design_Columns(),Design_Response(),Sequence_Column(),"
        "Sample_Include())"
    ) in dw_formula
    # Scalar numeric cell (the token is text and ignores the format).
    assert sheet.range(rc(11, _C_AE), rc(11, _C_AE)).number_format == "0.000"


def test_diagnostics_bfn_panel_dw_is_self_guarded_on_sequence_and_fe() -> None:
    # The second cell of the serial-correlation trigger matrix. Self-guarding:
    # every state it can show is visible in its own formula — it is NOT
    # dispatched from the DW cell or a shared selector.
    sheet = RecordingSheet(name="Regression")

    _write_diagnostics(_as_xw_sheet(sheet))

    assert sheet.cell(12, _C_AD).value == "BFN Panel Durbin-Watson"
    bfn_formula = cast(str, sheet.cell(12, _C_AE).api.Formula2)
    # The four trigger-matrix states, each an explicit token (never NA()/""):
    assert '"n/a — requires Sequence"' in bfn_formula       # no Sequence axis
    assert '"n/a — multiple Sequence flags"' in bfn_formula  # spec error
    assert '"n/a — no fixed effects"' in bfn_formula        # Sequence, no FE
    assert '"n/a — multiple FE variables"' in bfn_formula   # two-way: out of scope
    assert "NA()" not in bfn_formula
    # Guards key on the same LET-bound counts as the DW cell above.
    assert "SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))" in bfn_formula
    assert (
        'SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))'
    ) in bfn_formula
    assert "IF(seq_flags=0," in bfn_formula
    assert "IF(seq_flags>1," in bfn_formula
    assert "IF(fe_vars=0," in bfn_formula
    assert "IF(fe_vars>1," in bfn_formula
    # Active state: the panel statistic on the resolved grouping key and
    # Sequence axis, with the spec's visible Period In Use (column J) —
    # scalar out, no spill. The group argument routes through
    # Serial_Correlation_Group() (the grouping-key resolver, the single
    # retargeting point), never the FE column accessor directly.
    assert (
        "BFN_Panel_Durbin_Watson(Design_Columns(),Design_Response(),"
        "Serial_Correlation_Group(),Sequence_Column(),Base_Period_Delta(),"
        "Sample_Include())"
    ) in bfn_formula
    assert "Fixed_Effects_Column()" not in bfn_formula
    assert sheet.range(rc(12, _C_AE), rc(12, _C_AE)).number_format == "0.000"


def test_univariate_filter_reads_blanks_from_the_source_table() -> None:
    sheet = RecordingSheet()

    _write_data_zone(_as_xw_sheet(sheet))

    assert sheet.cell(4, 1).api.Formula2 == (
        '=IF(LifeExpectancyData[Life expectancy]="","",'
        "LifeExpectancyData[Life expectancy])"
    )
    assert sheet.cell(4, 2).api.Formula2 == "=ISNUMBER(LifeExpectancyData[Life expectancy])"


def test_expression_format_records_formula_and_font_options() -> None:
    sheet = RecordingSheet()

    condition = add_expression_format(
        _as_xw_sheet(sheet),
        "AB3:AB10",
        "=AB3>2*$P$6",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
        bold_font=True,
        stop_if_true=True,
    )

    assert condition.Type == 2
    assert condition.Formula1 == "=AB3>2*$P$6"
    assert condition.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert condition.Font.Color == excel_color(CF_DARK_RED_TEXT)
    assert condition.Font.Bold is True
    assert condition.StopIfTrue is True


def test_histogram_writer_records_method_cell_formulas() -> None:
    sheet = RecordingSheet()

    _write_histogram_table(_as_xw_sheet(sheet), 21, "Sturges")

    assert sheet.cell(2, 22).value == "Method"
    assert sheet.cell(2, 22).color == HEADER_COLOR
    assert sheet.cell(2, 23).value == "Sturges"
    assert sheet.cell(2, 23).color == HEADER_COLOR
    assert sheet.cell(3, 22).value == "Bins:"
    assert sheet.cell(3, 23).api.Formula2 == "=Number_Of_Histogram_Bins(UV_Data,W2,UV_Include)"
    assert sheet.cell(5, 22).api.Formula2 == "=Upper_Bin_Edges(UV_Data,W2,UV_Include)"
    assert sheet.cell(5, 23).api.Formula2 == "=Bin_Counts(UV_Data,W2,UV_Include)"
    assert sheet.cell(5, 21).api.Formula2 == "=Bin_Lower_Edges(UV_Data,W2,UV_Include)"
    cdf_cols = [21 + offset for offset, _, _, distribution in _HIST_COLUMNS if distribution]
    cdf_formulas = [sheet.cell(5, col).api.Formula2 for col in cdf_cols]
    assert all(formula and formula.startswith("=LET(edges,V5#,lower,U5#") for formula in cdf_formulas)
    assert all("HSTACK" not in formula for formula in cdf_formulas if formula)
    assert "CDF_Normal(edges,$I$5,$K$5,lower)" in _formula(sheet, 5, 24)
    assert "CDF_BetaPERT(edges,$I$12,$K$12,$M$12,lower)" in _formula(sheet, 5, 31)


def test_local_name_setup_removes_legacy_globals_and_uses_method_cells() -> None:
    from lambda_catalog.write_sheet_univariate import _HIST_BLOCKS

    histogram_names = [
        f"{prefix}_{suffix}"
        for prefix, _ in _HIST_BLOCKS
        for _, _, suffix, _ in _HIST_COLUMNS
    ]
    sheet = RecordingSheet(
        global_names=[
            "UV_Data",
            "UV_Include",
            "UV_n",
            *histogram_names,
            "UnrelatedName",
        ]
    )

    _setup_local_names(_as_xw_sheet(sheet))

    assert [item.Name for item in sheet.book.api.Names.items] == ["UnrelatedName"]
    names = sheet.api.Names
    assert names.by_short_name("UV_Data").RefersTo == "='Univariate'!$A$4#"
    for hist_name in histogram_names:
        refers_to = names.by_short_name(hist_name).RefersTo
        assert refers_to.startswith("=OFFSET('Univariate'!$")
        assert ",1,0,MAX(IFERROR(Number_Of_Histogram_Bins(" in refers_to
        assert refers_to.endswith(",1),1),1)")
    assert "$W$2" in names.by_short_name("UV_Sturges_Upper_Edges").RefersTo
    assert "$AI$2" in names.by_short_name("UV_Scott_Upper_Edges").RefersTo
    assert "$AU$2" in names.by_short_name("UV_FD_Upper_Edges").RefersTo
    assert names.by_short_name("UV_Sturges_Upper_Edges").Comment == (
        "Sturges Method histogram chart: category (X) axis bin edges"
    )
    assert names.by_short_name("UV_Sturges_Counts").Comment == (
        "Sturges Method histogram chart: bar values (bin counts)"
    )
    assert names.by_short_name("UV_Scott_Upper_Edges").Comment == (
        "Scott Method histogram chart: category (X) axis bin edges"
    )
    assert names.by_short_name("UV_Scott_Counts").Comment == (
        "Scott Method histogram chart: bar values (bin counts)"
    )
    assert names.by_short_name("UV_FD_Upper_Edges").Comment == (
        "Freedman-Diaconis Method histogram chart: category (X) axis bin edges"
    )
    assert names.by_short_name("UV_FD_Counts").Comment == (
        "Freedman-Diaconis Method histogram chart: bar values (bin counts)"
    )
    assert names.by_short_name("UV_Sturges_Normal_CDF").RefersTo == (
        "=OFFSET('Univariate'!$X$4,1,0,"
        "MAX(IFERROR(Number_Of_Histogram_Bins(UV_Data,'Univariate'!$W$2,UV_Include),1),1),1)"
    )
    assert getattr(names.by_short_name("UV_Sturges_Normal_CDF"), "Comment", None) is None


def test_local_name_setup_drops_legacy_globals_without_enumerating_all_workbook_names() -> None:
    sheet = RecordingSheet()
    sheet.book.api.Names = BrokenIndexRecordingNames(
        names=["UV_Data", "UV_Include", "UnrelatedName"]
    )

    _setup_local_names(_as_xw_sheet(sheet))

    assert [item.Name for item in sheet.book.api.Names.items] == ["UnrelatedName"]
    assert sheet.api.Names.by_short_name("UV_Data").RefersTo == "='Univariate'!$A$4#"


def test_local_name_setup_creates_expected_count_overlay_names() -> None:
    sheet = RecordingSheet()

    _setup_local_names(_as_xw_sheet(sheet))

    names = sheet.api.Names
    expected = names.by_short_name("UV_Sturges_Normal_Expected")
    assert expected.RefersTo == (
        "=OFFSET('Univariate'!$X$4,1,0,"
        "MAX(IFERROR(Number_Of_Histogram_Bins(UV_Data,'Univariate'!$W$2,UV_Include),1),1),1)"
        "*'Univariate'!$E$14"
    )
    assert expected.Comment == (
        "Sturges Method histogram chart: Normal overlay line "
        "(expected counts = bin probability × n)"
    )
    for prefix in ("UV_Sturges", "UV_Scott", "UV_FD"):
        for _, _, _, distribution in _HIST_COLUMNS:
            if not distribution:
                continue
            refers_to = names.by_short_name(f"{prefix}_{distribution}_Expected").RefersTo
            assert refers_to.startswith("=OFFSET('Univariate'!$")
            assert refers_to.endswith("*'Univariate'!$E$14")


def test_local_name_setup_creates_qq_chart_ranges() -> None:
    sheet = RecordingSheet()

    _setup_local_names(_as_xw_sheet(sheet))

    names = sheet.api.Names
    # The Q-Q zone leads the right-hand band at BE:BN, ahead of the fit zones.
    sample = names.by_short_name("UV_QQ_Sample")
    assert sample.RefersTo == (
        "=OFFSET('Univariate'!$BF$4,1,0,MAX(IFERROR('Univariate'!$E$14,1),1),1)"
    )
    assert sample.Comment == (
        "Q-Q plots: sorted sample values (Y axis, shared by all eight charts)"
    )
    normal = names.by_short_name("UV_QQ_Normal")
    assert normal.RefersTo == (
        "=OFFSET('Univariate'!$BG$4,1,0,MAX(IFERROR('Univariate'!$E$14,1),1),1)"
    )
    assert normal.Comment == "Normal Q-Q plot: theoretical quantiles (X axis)"
    assert names.by_short_name("UV_QQ_BetaPERT").RefersTo == (
        "=OFFSET('Univariate'!$BN$4,1,0,MAX(IFERROR('Univariate'!$E$14,1),1),1)"
    )
    # The plotting-position column is an intermediate and never charted.
    assert all(item.Name.split("!", 1)[-1] != "UV_QQ_P" for item in names.items)


def test_qq_data_zone_formulas_reference_fit_table_parameters() -> None:
    sheet = RecordingSheet()

    _write_qq_data(_as_xw_sheet(sheet))

    assert sheet.cell(1, 57).value == "Q-Q Plot Data"
    assert ((1, 57), (1, 66)) in sheet.merges
    assert sheet.cell(4, 57).value == "P"
    assert sheet.cell(4, 58).value == "Sample"
    assert sheet.cell(4, 66).value == "BetaPERT"

    assert _formula(sheet, 5, 57) == "=LET(n_,UV_n,IF(n_<=0,NA(),(SEQUENCE(n_)-0.5)/n_))"
    assert _formula(sheet, 5, 58) == "=SORT(FILTER(UV_Data,UV_Include))"
    assert _formula(sheet, 5, 59) == "=LET(p_,$BE$5#,NORM.INV(p_,$I$5,$K$5))"
    assert _formula(sheet, 5, 60) == "=LET(p_,$BE$5#,LOGNORM.INV(p_,$I$6,$K$6))"
    assert _formula(sheet, 5, 61) == "=LET(p_,$BE$5#,-LN(1-p_)/$I$7)"
    assert _formula(sheet, 5, 62) == "=LET(p_,$BE$5#,$K$8*(-LN(1-p_))^(1/$I$8))"
    assert _formula(sheet, 5, 63) == "=LET(p_,$BE$5#,GAMMA.INV(p_,$I$9,1/$K$9))"
    triangular = _formula(sheet, 5, 64)
    assert triangular.startswith("=LET(p_,$BE$5#,mn,$I$10,md,$K$10,mx,$M$10,")
    assert "mn+SQRT(p_*(mx-mn)*(md-mn))" in triangular
    assert "mx-SQRT((1-p_)*(mx-mn)*(mx-md))" in triangular
    beta = _formula(sheet, 5, 65)
    assert "mn,$E$9,range_,$E$11" in beta
    assert "BETA.INV(p_,$I$11,$K$11)*scale_+mn-pad" in beta
    betapert = _formula(sheet, 5, 66)
    assert "mn,$I$12,md,$K$12,mx,$M$12" in betapert
    # λ=4 PERT mapping — the μ-based form is 0/0 at a symmetric mode
    assert "alpha_param,1+4*(md-mn)/(mx-mn+1E-30)" in betapert
    assert "beta_param,1+4*(mx-md)/(mx-mn+1E-30)" in betapert
    assert "BETA.INV(p_,alpha_param,beta_param)*(mx-mn)+mn" in betapert

    assert sheet.range((5, 57), (2004, 57)).number_format == "0.0000"
    assert sheet.range((5, 58), (2004, 66)).number_format == "0.0"


def test_missing_count_formula_uses_unfiltered_active_range() -> None:
    formulas = dict(_STAT_ROWS)
    assert formulas["Missing"] == "=Missing_Count(UV_Data)"


def test_univariate_number_formats_are_one_decimal_or_integer_unless_nll() -> None:
    sheet = RecordingSheet()

    _write_data_zone(_as_xw_sheet(sheet))
    assert sheet.range((4, 1), (2003, 1)).number_format == "0.0"
    assert sheet.range((4, 2), (2003, 2)).number_format == "0"

    _write_descriptive_stats(_as_xw_sheet(sheet))
    assert sheet.cell(4, 5).number_format == "0.0"
    assert sheet.cell(14, 5).number_format == "0"
    assert sheet.cell(15, 5).number_format == "0"

    _write_histogram_table(_as_xw_sheet(sheet), 21, "Sturges")
    assert sheet.cell(3, 23).number_format == "0"
    assert sheet.range((5, 21), (2003, 21)).number_format == "0.0"
    assert sheet.range((5, 22), (2003, 22)).number_format == "0.0"
    assert sheet.range((5, 23), (2003, 23)).number_format == "0"

    _write_fitting_table(_as_xw_sheet(sheet))
    assert sheet.cell(5, 9).number_format == "0.0"
    assert sheet.cell(5, 14).number_format == "0.0E+00"
    assert sheet.cell(5, 15).number_format == "0"
    assert sheet.cell(5, 16).number_format == "0.0"

    _write_weibull_grid_search(_as_xw_sheet(sheet))
    # Weibull zone BP:BX, stage 1: BQ3 grid points, BT3:BX4 controls,
    # BP12:BP31 shape axis, BQ12:BQ31 profile NLL, BP3 Min NLL.
    assert sheet.cell(3, 69).number_format == "0"
    assert sheet.range((3, 72), (4, 76)).number_format == "0.0"
    assert sheet.range((12, 68), (31, 68)).number_format == "0.0"
    assert sheet.range((12, 69), (31, 69)).number_format == "0.0E+00"
    assert sheet.cell(3, 68).number_format == "0.0E+00"

    # Beta zone CJ:DD, stage 1 keeps the grid formats: CJ5 corner,
    # CK5:DD5 alpha row, CJ6:CJ25 beta column, CK6:DD25 body.
    assert sheet.cell(3, 89).number_format == "0"
    assert sheet.cell(5, 88).number_format == "0.0E+00"
    assert sheet.range((5, 89), (5, 108)).number_format == "0.0"
    assert sheet.range((6, 88), (25, 88)).number_format == "0.0"
    assert sheet.range((6, 89), (25, 108)).number_format == "0.0E+00"


def test_histogram_chart_title_cells_reference_method_headers():
    """Chart title formula cells at G14/G34/G54 reference the correct method header columns."""
    from lambda_catalog.write_sheet_univariate import (
        _write_histogram_chart_title_cells,
        _ROW_CHART1_TITLE, _ROW_CHART2_TITLE, _ROW_CHART3_TITLE,
    )
    sheet = RecordingSheet()
    _write_histogram_chart_title_cells(_as_xw_sheet(sheet))
    f1 = sheet.ranges[((_ROW_CHART1_TITLE, _C_FIT_FIRST),)].state.formula2
    f2 = sheet.ranges[((_ROW_CHART2_TITLE, _C_FIT_FIRST),)].state.formula2
    f3 = sheet.ranges[((_ROW_CHART3_TITLE, _C_FIT_FIRST),)].state.formula2
    assert f1 == '=W2&" Method Histogram"'
    assert f2 == '=AI2&" Method Histogram"'
    assert f3 == '=AU2&" Method Histogram"'


def test_weibull_grid_search_uses_final_layout_and_named_bodies() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    # Each fit owns a column zone and stacks its two stages inside it. Weibull
    # BP:BX and Gamma BZ:CH are nine wide; Beta CJ:DD is the only 21-wide zone.
    # Profile stage 2 is 5 rows down; Beta's is a full grid block + gap (26).
    assert ((1, 68), (1, 76)) in sheet.merges
    assert ((6, 68), (6, 76)) in sheet.merges
    assert ((1, 78), (1, 86)) in sheet.merges
    assert ((6, 78), (6, 86)) in sheet.merges
    assert ((1, 88), (1, 108)) in sheet.merges
    assert ((27, 88), (27, 108)) in sheet.merges

    assert sheet.cell(2, 68).value == "Min NLL:"
    assert sheet.cell(2, 69).value == "Grid Points"
    assert sheet.cell(3, 69).value == 20
    assert sheet.cell(2, 70).value is None
    assert [sheet.cell(2, col).value for col in range(71, 77)] == [
        "Parameter", "Start", "Min", "Max", "Step Size", "Best",
    ]
    assert sheet.cell(3, 71).value == "Shape (k)"
    assert sheet.cell(4, 71).value == "Scale (λ)"
    assert sheet.cell(3, 81).value == "Shape (α)"
    assert sheet.cell(4, 81).value == "Rate (β)"

    # Beta keeps the two-input Data Table control row ("Input", not "Start").
    assert sheet.cell(2, 89).value == "Rows/Columns"
    assert [sheet.cell(2, col).value for col in range(91, 97)] == [
        "Parameter", "Input", "Min", "Max", "Step Size", "Best",
    ]
    assert sheet.cell(3, 91).value == "Alpha (α)"
    assert sheet.cell(4, 91).value == "Beta (β)"

    # Both profile bodies share rows and sit side by side under the two halves
    # of the control block: stage 1 at offsets 0/1, stage 2 at offsets 3/4.
    assert sheet.cell(11, 68).value == "Shape (k)"
    assert sheet.cell(11, 69).value == "Profile NLL"
    assert sheet.cell(11, 71).value == "Shape (k)"
    assert sheet.cell(11, 72).value == "Profile NLL"
    assert sheet.cell(11, 78).value == "Shape (α)"
    assert sheet.cell(11, 79).value == "Profile NLL"

    names = sheet.api.Names
    assert names.by_short_name("UV_WB_S1").RefersTo == "='Univariate'!$BQ$12:$BQ$31"
    assert names.by_short_name("UV_WB_S1_Axis").RefersTo == "='Univariate'!$BP$12:$BP$31"
    assert names.by_short_name("UV_WB_S2").RefersTo == "='Univariate'!$BT$12:$BT$31"
    assert names.by_short_name("UV_WB_S2_Axis").RefersTo == "='Univariate'!$BS$12:$BS$31"
    assert names.by_short_name("UV_GAMMA_S1").RefersTo == "='Univariate'!$CA$12:$CA$31"
    assert names.by_short_name("UV_GAMMA_S1_Axis").RefersTo == "='Univariate'!$BZ$12:$BZ$31"
    assert names.by_short_name("UV_GAMMA_S2").RefersTo == "='Univariate'!$CD$12:$CD$31"
    assert names.by_short_name("UV_GAMMA_S2_Axis").RefersTo == "='Univariate'!$CC$12:$CC$31"
    assert names.by_short_name("UV_BETA_S1").RefersTo == "='Univariate'!$CK$6:$DD$25"
    assert names.by_short_name("UV_BETA_S2").RefersTo == "='Univariate'!$CK$32:$DD$51"


def test_band_zones_are_derived_in_order_from_one_table() -> None:
    """Zone starts, gaps, and the sheet's last column all come from _BAND_ZONES.

    Pinned because every address in the band is relative to these: a width
    change to any zone must not silently relocate the ones after it, and the
    fit zones must stay last so a grid resize displaces nothing.
    """
    from lambda_catalog.write_sheet_univariate import (
        _BAND_COL, _BAND_GAP_COLS, _BAND_LAST_COL, _BAND_ZONES,
    )

    assert [name for name, _ in _BAND_ZONES] == ["qq", "weibull", "gamma", "beta"]
    assert _BAND_COL == {"qq": 57, "weibull": 68, "gamma": 78, "beta": 88}
    assert _BAND_GAP_COLS == (67, 77, 87)
    assert _BAND_LAST_COL == 108


def test_profile_chart_ranges_are_offset_sized_by_the_grid_point_cell() -> None:
    """Each profile-NLL chart reads OFFSET ranges over BOTH of its curves.

    Stage 2 re-samples 20 points across ±1 stage-1 step, so it is the region the
    search actually resolved; the chart plots it as a second series. Each range
    is sized by its own stage's Grid Points cell — stage 1 by row 3, stage 2 by
    row 8, five rows down where that control block sits.
    """
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    names = sheet.api.Names
    assert names.by_short_name("UV_Profile_WB_S1_Axis").RefersTo == (
        "=OFFSET('Univariate'!$BP$11,1,0,MAX(IFERROR('Univariate'!$BQ$3,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_WB_S1_NLL").RefersTo == (
        "=OFFSET('Univariate'!$BQ$11,1,0,MAX(IFERROR('Univariate'!$BQ$3,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_WB_S2_Axis").RefersTo == (
        "=OFFSET('Univariate'!$BS$11,1,0,MAX(IFERROR('Univariate'!$BQ$8,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_WB_S2_NLL").RefersTo == (
        "=OFFSET('Univariate'!$BT$11,1,0,MAX(IFERROR('Univariate'!$BQ$8,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_GAMMA_S1_Axis").RefersTo == (
        "=OFFSET('Univariate'!$BZ$11,1,0,MAX(IFERROR('Univariate'!$CA$3,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_GAMMA_S1_NLL").RefersTo == (
        "=OFFSET('Univariate'!$CA$11,1,0,MAX(IFERROR('Univariate'!$CA$3,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_GAMMA_S2_Axis").RefersTo == (
        "=OFFSET('Univariate'!$CC$11,1,0,MAX(IFERROR('Univariate'!$CA$8,1),1),1)"
    )
    assert names.by_short_name("UV_Profile_GAMMA_S2_NLL").RefersTo == (
        "=OFFSET('Univariate'!$CD$11,1,0,MAX(IFERROR('Univariate'!$CA$8,1),1),1)"
    )


def test_profile_charts_anchor_one_clear_row_below_their_zone() -> None:
    """The charts sit under their own fit zone, not in the G:T chart band.

    Bodies end at row 31, row 32 stays blank, so the anchor is row 33 — Weibull
    at BP33 and Gamma at BZ33, each at its zone's first column. Pinned because
    the anchor is derived; an off-by-one would silently overlap the last body
    row or leave a two-row gap.
    """
    from lambda_catalog.write_sheet_univariate import (
        _BAND_COL, _N_PROFILE, _PS_R_BODY, _ROW_FIT_ZONE, _ROW_PROFILE_CHART,
    )

    last_body_row = _ROW_FIT_ZONE + _PS_R_BODY + _N_PROFILE - 1
    assert last_body_row == 31
    assert _ROW_PROFILE_CHART == last_body_row + 2 == 33
    assert (_BAND_COL["weibull"], _BAND_COL["gamma"]) == (68, 78)  # BP, BZ


def test_profile_stage_formulas_reference_visible_controls() -> None:
    """Weibull/Gamma stage 1: start, bracket, axis, and profiled-out partner."""
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    # Closed-form starting value, and the 3× bracket built from it. Weibull
    # zone BP:BX — Start BT, Min BU, Max BV, Step BW, Best BX.
    assert sheet.cell(3, 72).api.Formula2 == (
        "=IFERROR(LET(x,SORT(FILTER(UV_Data,UV_Include)),n,ROWS(x),"
        "p,(SEQUENCE(n)-0.5)/n,SLOPE(LN(-LN(1-p)),LN(x))),2)"
    )
    # Gamma zone BZ:CH — Start CD.
    assert sheet.cell(3, 82).api.Formula2 == (
        "=IFERROR(LET(x,FILTER(UV_Data,UV_Include),"
        "s,LN(AVERAGE(x))-AVERAGE(LN(x)),"
        "IF(s<=0,1,(3-s+SQRT((s-3)^2+24*s))/(12*s))),1)"
    )
    assert sheet.cell(3, 73).api.Formula2 == "=MAX(0.001,$BT$3/3)"
    assert sheet.cell(3, 74).api.Formula2 == "=$BT$3*3"
    assert sheet.cell(3, 75).api.Formula2 == "=($BV$3-$BU$3)/($BQ$3-1)"

    # The profiled-out parameter is solved, not searched: no bounds, no step.
    assert sheet.cell(4, 73).api.Formula2 is None
    assert sheet.cell(4, 74).api.Formula2 is None
    assert sheet.cell(4, 75).api.Formula2 is None

    # One column of trial shapes, one column of profile NLL beside it.
    assert sheet.cell(12, 68).api.Formula2 == "=SEQUENCE($BQ$3,1,$BU$3,$BW$3)"
    # The sample is bound once per cell and passed to both the partner and the
    # NLL call — the naive form re-filters the full input range twice per cell.
    # NLL's optional [filter] is omitted because x is already the included
    # numeric sample, so its ISNUMBER default filters nothing.
    assert sheet.cell(12, 69).api.Formula2 == (
        "=IFERROR(LET(x,FILTER(UV_Data,UV_Include),p,$BP$12,"
        "NLL_Weibull(x,p,((AVERAGE(x^p))^(1/p)))),1E+15)"
    )
    assert sheet.cell(31, 69).api.Formula2 == (
        "=IFERROR(LET(x,FILTER(UV_Data,UV_Include),p,$BP$31,"
        "NLL_Weibull(x,p,((AVERAGE(x^p))^(1/p)))),1E+15)"
    )
    assert sheet.cell(12, 79).api.Formula2 == (
        "=IFERROR(LET(x,FILTER(UV_Data,UV_Include),p,$BZ$12,"
        "NLL_Gamma(x,p,(p/AVERAGE(x)))),1E+15)"
    )

    assert sheet.cell(3, 68).api.Formula2 == (
        '=IFERROR(TAKE(Grid_Argument_Minimum(UV_WB_S1),,1),"—")'
    )

    # A single-column grid puts the minimum in the row slot, so the searched
    # value is INDEX(axis, row_location) — Grid_Search_Optimum's column half
    # would read the "Profile NLL" header instead of a parameter value.
    assert sheet.cell(3, 76).api.Formula2 == (
        '=IFERROR(INDEX(UV_WB_S1_Axis,'
        'INDEX(Grid_Argument_Minimum(UV_WB_S1),1,2)),"—")'
    )
    assert sheet.cell(4, 76).api.Formula2 == (
        '=IFERROR(((AVERAGE(FILTER(UV_Data,UV_Include)^$BX$3))^(1/$BX$3)),"—")'
    )
    assert sheet.cell(9, 86).api.Formula2 == (
        '=IFERROR(($CH$8/AVERAGE(FILTER(UV_Data,UV_Include))),"—")'
    )

    # Stage 2 sits 5 rows below stage 1 in the same zone, refining to ±1 step.
    assert sheet.cell(8, 72).api.Formula2 == "=$BX$3"
    assert sheet.cell(8, 73).api.Formula2 == "=MAX(0.001,$BX$3-$BW$3)"
    assert sheet.cell(8, 74).api.Formula2 == "=$BX$3+$BW$3"


def test_beta_grid_formulas_reference_visible_controls() -> None:
    """Beta stays a two-input Data Table, corner formula and both axes.

    Beta's zone is CJ:DD and its stage 2 stacks 26 rows down — a full grid
    block plus a gap row — where the profile fits' stage 2 is 5 rows down.
    """
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    assert sheet.cell(3, 95).api.Formula2 == "=($CP$3-$CO$3)/($CK$3-1)"
    assert sheet.cell(4, 95).api.Formula2 == "=($CP$4-$CO$4)/($CK$3-1)"
    assert sheet.cell(5, 89).api.Formula2 == "=SEQUENCE(1,$CK$3,$CO$3,$CQ$3)"
    assert sheet.cell(6, 88).api.Formula2 == "=SEQUENCE($CK$3,1,$CO$4,$CQ$4)"
    assert sheet.cell(3, 88).api.Formula2 == (
        '=IFERROR(TAKE(Grid_Argument_Minimum(UV_BETA_S1),,1),"—")'
    )
    assert sheet.cell(3, 96).api.Formula2 == "=Grid_Search_Optimum(UV_BETA_S1)"
    assert sheet.cell(4, 96).api.Formula2 is None
    assert "NLL_Beta(z,$CN$3,$CN$4)" in _formula(sheet, 5, 88)

    assert sheet.cell(29, 93).api.Formula2 == "=MAX(0.001,$CR$3-$CQ$3)"
    assert sheet.cell(29, 94).api.Formula2 == "=$CR$3+$CQ$3"
    assert sheet.cell(30, 93).api.Formula2 == "=MAX(0.001,$CR$4-$CQ$4)"
    assert sheet.cell(30, 94).api.Formula2 == "=$CR$4+$CQ$4"


def test_weibull_grid_uses_visible_inputs_borders_and_boundary_rules() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    # Beta's two stages are the artifact's only Data Tables — the Weibull and
    # Gamma stages are 1-D profile columns and wire no Data Table object.
    assert sheet.tables == [
        {
            "range": ((5, 88), (25, 108)),
            "row_input": ((3, 92),),
            "column_input": ((4, 92),),
        },
        {
            "range": ((31, 88), (51, 108)),
            "row_input": ((29, 92),),
            "column_input": ((30, 92),),
        },
    ]

    for address in (
        ((2, 68), (3, 68)),     # Min NLL
        ((2, 69), (3, 69)),     # Grid Points
        ((2, 71), (4, 76)),     # parameter table
        ((11, 68), (31, 69)),   # stage 1 axis + body
        ((11, 71), (31, 72)),   # stage 2 axis + body
    ):
        assert set(sheet.range(*address).api._borders) == {7, 8, 9, 10}
    assert sheet.range((2, 70), (4, 70)).api._borders == {}

    # Only the searched parameter has a boundary to hit; the profiled-out
    # partner is solved in closed form, so it carries no rule.
    shape_rule = sheet.cell(3, 76).api.FormatConditions.items[0].Formula1
    assert shape_rule == (
        "=OR(INDEX(Grid_Argument_Minimum(UV_WB_S1),1,2)=1,"
        "INDEX(Grid_Argument_Minimum(UV_WB_S1),1,2)=$BQ$3)"
    )
    assert sheet.cell(4, 76).api.FormatConditions.items == []
    assert len(sheet.range((12, 69), (31, 69)).api.FormatConditions.color_scales) == 1

    # Beta keeps both boundary rules — its two parameters are both searched.
    beta_alpha_rule = sheet.cell(3, 96).api.FormatConditions.items[0].Formula1
    beta_beta_rule = sheet.cell(4, 96).api.FormatConditions.items[0].Formula1
    assert beta_alpha_rule == (
        "=OR(INDEX(Grid_Argument_Minimum(UV_BETA_S1),1,3)=1,"
        "INDEX(Grid_Argument_Minimum(UV_BETA_S1),1,3)=$CK$3)"
    )
    assert beta_beta_rule == (
        "=OR(INDEX(Grid_Argument_Minimum(UV_BETA_S1),1,2)=1,"
        "INDEX(Grid_Argument_Minimum(UV_BETA_S1),1,2)=$CK$3)"
    )


def test_weibull_bounds_and_summary_reference_final_best_cells() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))
    rows = {name: row for row, name, *_ in _dist_rows(5)}
    weibull_row = next(item for item in _dist_rows(5) if item[1] == "Weibull")
    gamma_row = next(item for item in _dist_rows(5) if item[1] == "Gamma")
    beta_row = next(item for item in _dist_rows(5) if item[1] == "Beta")

    # Each fit now has its own Best column and its own stage 2 row, so these
    # three pairs no longer share a column the way they did under the
    # stage-major layout. Weibull BX, Gamma CH, Beta CR.
    assert rows["Weibull"] == weibull_row[0]
    assert weibull_row[3] == "=$BX$8"
    assert weibull_row[5] == "=$BX$9"
    assert rows["Gamma"] == gamma_row[0]
    assert gamma_row[3] == "=$CH$8"
    assert gamma_row[5] == "=$CH$9"
    assert rows["Beta"] == beta_row[0]
    assert beta_row[3] == "=$CR$29"
    assert beta_row[5] == "=$CR$30"
    assert "NLL_Beta" in beta_row[8]
    assert "COUNT(d)*LN(scale_)" in beta_row[8]

    # Stage 1 bounds are input-coloured overrides; stage 2's are derived.
    assert sheet.cell(3, 73).color == INPUT_COLOR
    assert sheet.cell(3, 74).color == INPUT_COLOR
    assert sheet.cell(8, 73).color is None
    assert sheet.cell(8, 74).color is None


def test_grid_stage_returns_visible_step_and_count_references() -> None:
    sheet = RecordingSheet()

    refs = _write_grid_stage(
        _as_xw_sheet(sheet),
        row_start=1,
        col_start=29,
        title="Stage",
        body_name="UV_TEST",
        p1_label="P1",
        p2_label="P2",
        nll_formula=lambda p1, p2: f"=NLL_Test({p1},{p2})",
        p1_min=0.5,
        p1_max=10.0,
        p2_min=0.1,
        p2_max=20.0,
    )

    assert refs["step_p1"] == "$AJ$3"
    assert refs["step_p2"] == "$AJ$4"
    assert refs["n_grid"] == "$AD$3"


def test_betapert_cdf_expr_uses_valid_let_variable_names() -> None:
    """BetaPERT CDF must not use `a1`/`a2` as LET names — they are A1-style cell refs."""
    betapert_row = next(item for item in _dist_rows(5) if item[1] == "BetaPERT")
    cdf_expr = betapert_row[10]

    assert "alpha_param" in cdf_expr
    assert "beta_param" in cdf_expr
    # a1 / a2 must not appear as LET variable definitions
    assert "a1," not in cdf_expr
    assert ",a1," not in cdf_expr
    assert "a2," not in cdf_expr
    assert ",a2," not in cdf_expr
    # parentheses must balance
    assert cdf_expr.count("(") == cdf_expr.count(")")

"""Tests for workbook-writing logic that do not require a live Excel process."""
# pylint: disable=import-outside-toplevel,missing-function-docstring,protected-access
from typing import cast

import xlwings as xw

from lambda_catalog.sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_LIGHT_RED_FILL,
    HEADER_COLOR,
    INPUT_COLOR,
)
from lambda_catalog.workbook_helpers import add_expression_format, excel_color
from lambda_catalog.write_sheet_mlr_observation_test import _section_formula
from lambda_catalog.write_sheet_regression import (
    _C_S,
    _C_T,
    _C_U,
    _C_W,
    _C_X,
    _C_Z,
    _C_AC,
    _C_AD,
    _C_AE,
    _C_AF,
    _C_AH,
    _C_AI,
    _C_AJ,
    _C_AP,
    _PRED_INPUT_FIRST_ROW,
    _PRED_INPUT_LAST_ROW,
    _setup_local_names as _setup_regression_names,
    _write_coefficients,
    _write_prediction_interval,
    _write_prediction_inputs,
    _write_regression_outputs_header,
    _write_residuals,
)
from lambda_catalog.write_sheet_mlr_scalar_test import _actual_formula
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
    _write_weibull_grid_search,
)
from tests.recording_sheet import RecordingSheet


def _as_xw_sheet(sheet: RecordingSheet) -> xw.Sheet:
    return cast(xw.Sheet, sheet)


def _formula(sheet: RecordingSheet, row: int, col: int) -> str:
    return cast(str, sheet.cell(row, col).api.Formula2)


def test_scalar_formula_maps_include_to_the_sheet_filter() -> None:
    formula = _actual_formula("Observations", ("Y", "Include"))
    assert formula == "=Observations(y, Regression_Sample_Include)"


def test_observation_y_only_formulas_reuse_first_spill() -> None:
    anchors: dict[str, int] = {}

    first = _section_formula(1, True, "Rank_Fraction", 3, 3, anchors)
    second = _section_formula(5, False, "Rank_Fraction", 158, 3, anchors)
    prediction = _section_formula(5, False, "Predictions", 158, 6, anchors)

    assert first == "=Rank_Fraction(y,Regression_Sample_Include)"
    assert second == "=$C$3#"
    assert prediction == (
        "=LET(x_s,OFFSET(y,0,1,ROWS(y),5),"
        "Predictions(x_s,y,FALSE,Regression_Sample_Include))"
    )


def test_regression_names_register_spec_wiring_and_constructors() -> None:
    sheet = RecordingSheet(name="Regression")

    _setup_regression_names(_as_xw_sheet(sheet))

    names = [item.Name.split("!", 1)[-1] for item in sheet.api.Names.items]
    # Spec wiring precedes the closures, which precede the Regression-only names.
    assert names.index("Spec_Include") < names.index("Sample_Include")
    assert names.index("Sample_Include") < names.index("X_s")
    assert names.index("X_s") < names.index("Zero_Predictors_Selected")
    # The v1 hard-wired names are gone.
    for legacy in ("All_Xs", "Coefficient_Name_Col", "Ind_Var_Include", "y",
                   "Regression_Sample_Include", "data_identifiers"):
        assert legacy not in names, legacy

    x_s_formula = sheet.api.Names.by_short_name("X_s").RefersTo
    assert x_s_formula.startswith("=LAMBDA(LET(")
    assert "Dummy_Levels(" in x_s_formula

    zero_formula = sheet.api.Names.by_short_name("Zero_Predictors_Selected").RefersTo
    assert zero_formula == "=LAMBDA(IFERROR(COLUMNS(X_s()),0)=0)"

    assert sheet.api.Names.by_short_name("Allow_Intercept").RefersTo == (
        "='Regression'!$C$2"
    )
    assert sheet.api.Names.by_short_name("alpha").RefersTo == "=Regression!$T$12"


def test_regression_chart_names_size_to_the_observation_cell() -> None:
    sheet = RecordingSheet(name="Regression")

    _setup_regression_names(_as_xw_sheet(sheet))

    fit_y = sheet.api.Names.by_short_name("RegChartFitY").RefersTo
    assert fit_y == (
        "=OFFSET('Regression'!$AG$2,1,0,"
        "MAX(IFERROR('Regression'!$T$8,1),1),1)"
    )
    press = sheet.api.Names.by_short_name("RegChartPRESSResid").RefersTo
    assert "$AP$2" in press


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

    formula = sheet.cell(3, 29).api.Formula2
    assert formula is not None
    assert formula.startswith("=IF(Zero_Predictors_Selected(),")
    assert "IFERROR" not in formula
    # Inputs correspond 1:1 to constructed columns — TAKE exactly k rows,
    # no Include-filter needed.
    assert "LET(pred_input,VSTACK($AC$12,TAKE($AC$13:$AC$62,COLUMNS(X_s())))" in formula
    assert "Prediction_Interval(X_s(),Response_Column(),pred_input" in formula
    assert "Intercept_Only_Point()" in formula


def test_prediction_prefills_index_the_single_training_mean_spill() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_prediction_inputs(_as_xw_sheet(sheet))

    # The Training Mean spill is the ONE X_s() evaluation for the whole
    # prefill band; it owns column AD downward so it can never collide with
    # another spill when the source data or spec changes.
    assert sheet.cell(11, _C_AD).value == "Training Mean"
    means = _formula(sheet, _PRED_INPUT_FIRST_ROW, _C_AD)
    assert means == (
        "=IFERROR(TRANSPOSE(BYCOL(FILTER(X_s(),Sample_Include()),"
        'LAMBDA(c,AVERAGE(c)))),"")'
    )

    # Perf tripwire: X_s() is a full design-matrix construction on every
    # call, so no prefill cell may invoke it — 50 cells × 2 calls made the
    # workbook's first full calculation take ~20 minutes.
    for row in (_PRED_INPUT_FIRST_ROW, _PRED_INPUT_LAST_ROW):
        prefill = _formula(sheet, row, _C_AC)
        assert "X_s()" not in prefill
        assert f"INDEX($AD${_PRED_INPUT_FIRST_ROW}#" in prefill
        assert f"IFERROR(ROWS($AD${_PRED_INPUT_FIRST_ROW}#),0)" in prefill


def test_write_coefficients_adds_intercept_only_closed_form_branch() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_coefficients(_as_xw_sheet(sheet))

    label_formula = _formula(sheet, 21, _C_S)
    assert label_formula.startswith("=IF(Zero_Predictors_Selected(),")
    assert 'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),"Intercept",NA())' in label_formula
    # Level-qualified names come from the constructor twin (a row vector).
    assert 'VSTACK("Intercept",TRANSPOSE(Constructed_Column_Names()))' in label_formula

    coefficient_formula = _formula(sheet, 21, _C_T)
    assert "IF(AND(Allow_Intercept,Intercept_Only_N()>=1),Intercept_Only_Point(),NA())" in coefficient_formula
    assert "Coefficients(X_s(),Response_Column(),Allow_Intercept,Sample_Include())" in coefficient_formula

    se_formula = _formula(sheet, 21, _C_U)
    assert "IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_SE(),NA())" in se_formula

    beta_formula = _formula(sheet, 21, _C_Z)
    assert 'IF(Allow_Intercept,"",NA())' in beta_formula


def test_regression_outputs_header_writes_predicted_variable_readout() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_regression_outputs_header(_as_xw_sheet(sheet))

    assert sheet.cell(2, _C_W).value == "Predicted Variable"
    assert sheet.cell(2, _C_W).api.Font.Bold is True
    assert sheet.cell(2, _C_W).color == HEADER_COLOR
    # Derived response name — the header of the Role=Response spec row.
    readout = sheet.cell(2, _C_X).api.Formula2
    assert readout is not None
    assert readout.startswith("=IFERROR(INDEX(TOROW(Header_Names),")
    assert 'XMATCH("Response (y)"' in readout
    assert sheet.cell(2, _C_X).api.Font.Bold is True
    assert sheet.cell(2, _C_X).color == HEADER_COLOR


def test_write_residuals_writes_row_labels_and_diagnostics() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_residuals(_as_xw_sheet(sheet))

    # Static header; Row_Labels() supplies its own per-row content and
    # no-Identifier fallback, so only an all-FALSE mask is absorbed here.
    assert sheet.cell(2, _C_AE).value == "Observation"
    assert sheet.cell(3, _C_AE).api.Formula2 == (
        "=IFERROR(FILTER(Row_Labels(),Sample_Include()),NA())"
    )
    # The diagnostics columns shift one slot right of the identifiers column.
    assert sheet.cell(2, _C_AF).value == "Y"
    assert sheet.cell(3, _C_AF).api.Formula2 == (
        "=Dependent_Variable(Response_Column(),Sample_Include())"
    )
    assert sheet.cell(3, _C_AH).api.Formula2 == (
        "=Residuals(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"
    )
    assert sheet.cell(3, _C_AJ).api.Formula2 == (
        "=Hat_Diagonal(X_s(),Allow_Intercept,Sample_Include())"
    )
    assert sheet.cell(3, _C_AI).api.Formula2 == (
        "=LOOCV_Residual(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"
    )
    assert sheet.cell(3, _C_AP).api.Formula2 == (
        "=LOOCV_Residual(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"
    )


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
    assert names.by_short_name("UV_Sturges_Normal_CDF").RefersTo == (
        "=OFFSET('Univariate'!$X$4,1,0,"
        "MAX(IFERROR(Number_Of_Histogram_Bins(UV_Data,'Univariate'!$W$2,UV_Include),1),1),1)"
    )


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
    assert sheet.cell(3, 58).number_format == "0"
    assert sheet.range((3, 61), (4, 65)).number_format == "0.0"
    assert sheet.cell(5, 57).number_format == "0.0E+00"
    assert sheet.range((5, 58), (5, 77)).number_format == "0.0"
    assert sheet.range((6, 57), (25, 57)).number_format == "0.0"
    assert sheet.range((6, 58), (25, 77)).number_format == "0.0E+00"
    assert sheet.cell(3, 57).number_format == "0.0E+00"


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

    assert ((1, 57), (1, 77)) in sheet.merges
    assert ((1, 79), (1, 99)) in sheet.merges
    assert sheet.cell(2, 57).value == "Min NLL:"
    assert sheet.cell(2, 58).value == "Rows/Columns"
    assert sheet.cell(3, 58).value == 20
    assert sheet.cell(2, 59).value is None
    assert [sheet.cell(2, col).value for col in range(60, 66)] == [
        "Parameter", "Input", "Min", "Max", "Step Size", "Best",
    ]
    assert sheet.cell(3, 60).value == "Shape (k)"
    assert sheet.cell(29, 60).value == "Shape (α)"
    assert sheet.cell(30, 60).value == "Rate (β)"
    assert sheet.cell(55, 60).value == "Alpha (α)"
    assert sheet.cell(56, 60).value == "Beta (β)"
    assert sheet.cell(4, 60).value == "Scale (λ)"

    names = sheet.api.Names
    assert names.by_short_name("UV_WB_S1").RefersTo == "='Univariate'!$BF$6:$BY$25"
    assert names.by_short_name("UV_WB_S2").RefersTo == "='Univariate'!$CB$6:$CU$25"
    assert names.by_short_name("UV_GAMMA_S1").RefersTo == "='Univariate'!$BF$32:$BY$51"
    assert names.by_short_name("UV_GAMMA_S2").RefersTo == "='Univariate'!$CB$32:$CU$51"
    assert names.by_short_name("UV_BETA_S1").RefersTo == "='Univariate'!$BF$58:$BY$77"
    assert names.by_short_name("UV_BETA_S2").RefersTo == "='Univariate'!$CB$58:$CU$77"


def test_weibull_grid_formulas_reference_visible_controls() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    assert sheet.cell(3, 64).api.Formula2 == "=($BK$3-$BJ$3)/($BF$3-1)"
    assert sheet.cell(4, 64).api.Formula2 == "=($BK$4-$BJ$4)/($BF$3-1)"
    assert sheet.cell(5, 58).api.Formula2 == "=SEQUENCE(1,$BF$3,$BJ$3,$BL$3)"
    assert sheet.cell(6, 57).api.Formula2 == "=SEQUENCE($BF$3,1,$BJ$4,$BL$4)"
    assert sheet.cell(3, 57).api.Formula2 == (
        '=IFERROR(TAKE(Grid_Argument_Minimum(UV_WB_S1),,1),"—")'
    )
    assert sheet.cell(3, 65).api.Formula2 == "=Grid_Search_Optimum(UV_WB_S1)"
    assert sheet.cell(4, 65).api.Formula2 is None
    assert sheet.cell(5, 57).api.Formula2 == (
        "=NLL_Weibull(UV_Data,$BI$3,$BI$4,UV_Include)"
    )
    assert sheet.cell(31, 57).api.Formula2 == (
        "=NLL_Gamma(UV_Data,$BI$29,$BI$30,UV_Include)"
    )
    assert "NLL_Beta(z,$BI$55,$BI$56)" in _formula(sheet, 57, 57)

    assert sheet.cell(3, 84).api.Formula2 == "=MAX(0.001,$BM$3-$BL$3)"
    assert sheet.cell(3, 85).api.Formula2 == "=$BM$3+$BL$3"
    assert sheet.cell(4, 84).api.Formula2 == "=MAX(0.001,$BM$4-$BL$4)"
    assert sheet.cell(4, 85).api.Formula2 == "=$BM$4+$BL$4"


def test_weibull_grid_uses_visible_inputs_borders_and_boundary_rules() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))

    assert sheet.tables[:2] == [
        {
            "range": ((5, 57), (25, 77)),
            "row_input": ((3, 61),),
            "column_input": ((4, 61),),
        },
        {
            "range": ((5, 79), (25, 99)),
            "row_input": ((3, 83),),
            "column_input": ((4, 83),),
        },
    ]
    assert sheet.tables[2:] == [
        {
            "range": ((31, 57), (51, 77)),
            "row_input": ((29, 61),),
            "column_input": ((30, 61),),
        },
        {
            "range": ((31, 79), (51, 99)),
            "row_input": ((29, 83),),
            "column_input": ((30, 83),),
        },
        {
            "range": ((57, 57), (77, 77)),
            "row_input": ((55, 61),),
            "column_input": ((56, 61),),
        },
        {
            "range": ((57, 79), (77, 99)),
            "row_input": ((55, 83),),
            "column_input": ((56, 83),),
        },
    ]

    for address in (
        ((2, 57), (3, 57)),
        ((2, 58), (3, 58)),
        ((2, 60), (4, 65)),
        ((5, 57), (25, 77)),
    ):
        assert set(sheet.range(*address).api._borders) == {7, 8, 9, 10}
    assert sheet.range((2, 59), (4, 59)).api._borders == {}

    shape_rule = sheet.cell(3, 65).api.FormatConditions.items[0].Formula1
    scale_rule = sheet.cell(4, 65).api.FormatConditions.items[0].Formula1
    assert shape_rule == "=OR(INDEX(Grid_Argument_Minimum(UV_WB_S1),1,3)=1,INDEX(Grid_Argument_Minimum(UV_WB_S1),1,3)=$BF$3)"
    assert scale_rule == "=OR(INDEX(Grid_Argument_Minimum(UV_WB_S1),1,2)=1,INDEX(Grid_Argument_Minimum(UV_WB_S1),1,2)=$BF$3)"
    assert len(sheet.range((6, 58), (25, 77)).api.FormatConditions.color_scales) == 1


def test_weibull_bounds_and_summary_reference_final_best_cells() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(_as_xw_sheet(sheet))
    rows = {name: row for row, name, *_ in _dist_rows(5)}
    weibull_row = next(item for item in _dist_rows(5) if item[1] == "Weibull")
    gamma_row = next(item for item in _dist_rows(5) if item[1] == "Gamma")
    beta_row = next(item for item in _dist_rows(5) if item[1] == "Beta")

    assert rows["Weibull"] == weibull_row[0]
    assert weibull_row[3] == "=$CI$3"
    assert weibull_row[5] == "=$CI$4"
    assert rows["Gamma"] == gamma_row[0]
    assert gamma_row[3] == "=$CI$29"
    assert gamma_row[5] == "=$CI$30"
    assert rows["Beta"] == beta_row[0]
    assert beta_row[3] == "=$CI$55"
    assert beta_row[5] == "=$CI$56"
    assert "NLL_Beta" in beta_row[8]
    assert "COUNT(d)*LN(scale_)" in beta_row[8]
    assert sheet.cell(3, 62).color == INPUT_COLOR
    assert sheet.cell(3, 63).color == INPUT_COLOR
    assert sheet.cell(3, 84).color is None
    assert sheet.cell(3, 86).color is None


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

"""Tests for workbook-writing logic that do not require a live Excel process."""
from lambda_catalog.sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_LIGHT_RED_FILL,
    HEADER_COLOR,
    INPUT_COLOR,
)
from lambda_catalog.workbook_helpers import add_expression_format, excel_color
from lambda_catalog.write_sheet_regression import (
    _setup_local_names as _setup_regression_names,
    _write_prediction_interval,
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


def test_scalar_formula_maps_include_to_the_sheet_filter() -> None:
    formula = _actual_formula("Observations", ("Y", "Include"))
    assert formula == "=Observations(y, Regression_Sample_Include)"


def test_regression_predictor_name_preserves_a_multicolumn_range() -> None:
    sheet = RecordingSheet(name="Regression")

    _setup_regression_names(sheet)

    local_name_order = [item.Name.split("!", 1)[-1] for item in sheet.api.Names.items]
    assert local_name_order.index("Ind_Var_Include") < local_name_order.index("x_s")
    x_s_formula = sheet.api.Names.by_short_name("x_s").RefersTo
    assert x_s_formula.startswith("=LAMBDA(")
    assert "TRANSPOSE(FILTER(TRANSPOSE(All_Xs)" in x_s_formula
    assert "TAKE(All_Xs,,1)" in x_s_formula


def test_prediction_interval_binds_selected_inputs_in_the_cell_formula() -> None:
    sheet = RecordingSheet(name="Regression")

    _write_prediction_interval(sheet)

    formula = sheet.cell(3, 22).api.Formula2
    assert formula is not None
    assert formula.startswith("=LET(pred_input,VSTACK($V$12,")
    assert "FILTER($V$13:$V$30" in formula
    assert "Prediction_Interval(x_s(),y,pred_input" in formula


def test_univariate_filter_reads_blanks_from_the_source_table() -> None:
    sheet = RecordingSheet()

    _write_data_zone(sheet)

    assert sheet.cell(4, 1).api.Formula2 == (
        '=IF(LifeExpectancyData[Life expectancy]="","",'
        "LifeExpectancyData[Life expectancy])"
    )
    assert sheet.cell(4, 2).api.Formula2 == (
        "=MAP(LifeExpectancyData[Life expectancy],Data_Completeness)"
    )


def test_expression_format_records_formula_and_font_options() -> None:
    sheet = RecordingSheet()

    condition = add_expression_format(
        sheet,
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

    _write_histogram_table(sheet, 21, "Sturges")

    assert sheet.cell(2, 21).value == "Method"
    assert sheet.cell(2, 21).color == HEADER_COLOR
    assert sheet.cell(2, 22).value == "Sturges"
    assert sheet.cell(2, 22).color == HEADER_COLOR
    assert sheet.cell(3, 21).value == "Bins:"
    assert sheet.cell(3, 22).api.Formula2 == "=num_histogram_bins(UV_Data,V2,UV_Include)"
    assert sheet.cell(5, 21).api.Formula2 == "=Bin_Edges(UV_Data,V2,UV_Include)"
    assert sheet.cell(5, 22).api.Formula2 == "=Bin_Counts(UV_Data,U5#,UV_Include)"
    cdf_cols = [21 + offset for offset, _, _, distribution in _HIST_COLUMNS if distribution]
    cdf_formulas = [sheet.cell(5, col).api.Formula2 for col in cdf_cols]
    assert all(formula and formula.startswith("=LET(edges,U5#") for formula in cdf_formulas)
    assert all("HSTACK" not in formula for formula in cdf_formulas if formula)
    assert "CDF_Normal(edges,$I$5,$K$5,lower)" in sheet.cell(5, 23).api.Formula2
    assert "CDF_BetaPERT(edges,$I$12,$K$12,$M$12,lower)" in sheet.cell(5, 30).api.Formula2


def test_local_name_setup_removes_legacy_globals_and_uses_method_cells() -> None:
    histogram_names = [
        f"{prefix}_{suffix}"
        for prefix in ("UV_Sturges", "UV_Scott", "UV_FD")
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

    _setup_local_names(sheet)

    assert [item.Name for item in sheet.book.api.Names.items] == ["UnrelatedName"]
    names = sheet.api.Names
    assert names.by_short_name("UV_Data").RefersTo == "='Univariate'!$A$4#"
    for hist_name in histogram_names:
        refers_to = names.by_short_name(hist_name).RefersTo
        assert refers_to.startswith("=OFFSET('Univariate'!$")
        assert ",1,0,MAX(IFERROR(num_histogram_bins(" in refers_to
        assert refers_to.endswith(",1),1),1)")
    assert "$V$2" in names.by_short_name("UV_Sturges_Edges").RefersTo
    assert "$AG$2" in names.by_short_name("UV_Scott_Edges").RefersTo
    assert "$AR$2" in names.by_short_name("UV_FD_Edges").RefersTo
    assert names.by_short_name("UV_Sturges_Normal_CDF").RefersTo == (
        "=OFFSET('Univariate'!$W$4,1,0,"
        "MAX(IFERROR(num_histogram_bins(UV_Data,'Univariate'!$V$2,UV_Include),1),1),1)"
    )


def test_missing_count_formula_uses_unfiltered_active_range() -> None:
    formulas = dict(_STAT_ROWS)
    assert formulas["Missing"] == "=Missing_Count(UV_Data)"


def test_univariate_number_formats_are_one_decimal_or_integer_unless_nll() -> None:
    sheet = RecordingSheet()

    _write_data_zone(sheet)
    assert sheet.range((4, 1), (2003, 1)).number_format == "0.0"
    assert sheet.range((4, 2), (2003, 2)).number_format == "0"

    _write_descriptive_stats(sheet)
    assert sheet.cell(4, 5).number_format == "0.0"
    assert sheet.cell(14, 5).number_format == "0"
    assert sheet.cell(15, 5).number_format == "0"

    _write_histogram_table(sheet, 21, "Sturges")
    assert sheet.cell(3, 22).number_format == "0"
    assert sheet.range((5, 21), (2003, 21)).number_format == "0.0"
    assert sheet.range((5, 22), (2003, 22)).number_format == "0"

    _write_fitting_table(sheet)
    assert sheet.cell(5, 9).number_format == "0.0"
    assert sheet.cell(5, 14).number_format == "0.0E+00"
    assert sheet.cell(5, 15).number_format == "0"
    assert sheet.cell(5, 16).number_format == "0.0"

    _write_weibull_grid_search(sheet)
    assert sheet.cell(3, 55).number_format == "0"
    assert sheet.range((3, 58), (4, 62)).number_format == "0.0"
    assert sheet.cell(5, 54).number_format == "0.0E+00"
    assert sheet.range((5, 55), (5, 74)).number_format == "0.0"
    assert sheet.range((6, 54), (25, 54)).number_format == "0.0"
    assert sheet.range((6, 55), (25, 74)).number_format == "0.0E+00"
    assert sheet.cell(3, 54).number_format == "0.0E+00"


def test_histogram_chart_title_cells_reference_method_headers():
    """Chart title formula cells at G14/G34/G54 reference the correct method header columns."""
    from lambda_catalog.write_sheet_univariate import (
        _write_histogram_chart_title_cells,
        _ROW_CHART1_TITLE, _ROW_CHART2_TITLE, _ROW_CHART3_TITLE,
    )
    sheet = RecordingSheet()
    _write_histogram_chart_title_cells(sheet)
    f1 = sheet.ranges[((_ROW_CHART1_TITLE, _C_FIT_FIRST),)].state.formula2
    f2 = sheet.ranges[((_ROW_CHART2_TITLE, _C_FIT_FIRST),)].state.formula2
    f3 = sheet.ranges[((_ROW_CHART3_TITLE, _C_FIT_FIRST),)].state.formula2
    assert f1 == '=V2&" Method Histogram"'
    assert f2 == '=AG2&" Method Histogram"'
    assert f3 == '=AR2&" Method Histogram"'


def test_weibull_grid_search_uses_final_layout_and_named_bodies() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert ((1, 54), (1, 74)) in sheet.merges
    assert ((1, 76), (1, 96)) in sheet.merges
    assert sheet.cell(2, 54).value == "Min NLL:"
    assert sheet.cell(2, 55).value == "Rows/Columns"
    assert sheet.cell(3, 55).value == 20
    assert sheet.cell(2, 56).value is None
    assert [sheet.cell(2, col).value for col in range(57, 63)] == [
        "Parameter", "Input", "Min", "Max", "Step Size", "Best",
    ]
    assert sheet.cell(3, 57).value == "Shape (k)"
    assert sheet.cell(29, 57).value == "Shape (α)"
    assert sheet.cell(30, 57).value == "Rate (β)"
    assert sheet.cell(55, 57).value == "Alpha (α)"
    assert sheet.cell(56, 57).value == "Beta (β)"
    assert sheet.cell(4, 57).value == "Scale (λ)"

    names = sheet.api.Names
    assert names.by_short_name("UV_WB_S1").RefersTo == "='Univariate'!$BC$6:$BV$25"
    assert names.by_short_name("UV_WB_S2").RefersTo == "='Univariate'!$BY$6:$CR$25"
    assert names.by_short_name("UV_GAMMA_S1").RefersTo == "='Univariate'!$BC$32:$BV$51"
    assert names.by_short_name("UV_GAMMA_S2").RefersTo == "='Univariate'!$BY$32:$CR$51"
    assert names.by_short_name("UV_BETA_S1").RefersTo == "='Univariate'!$BC$58:$BV$77"
    assert names.by_short_name("UV_BETA_S2").RefersTo == "='Univariate'!$BY$58:$CR$77"


def test_weibull_grid_formulas_reference_visible_controls() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert sheet.cell(3, 61).api.Formula2 == "=($BH$3-$BG$3)/($BC$3-1)"
    assert sheet.cell(4, 61).api.Formula2 == "=($BH$4-$BG$4)/($BC$3-1)"
    assert sheet.cell(5, 55).api.Formula2 == "=SEQUENCE(1,$BC$3,$BG$3,$BI$3)"
    assert sheet.cell(6, 54).api.Formula2 == "=SEQUENCE($BC$3,1,$BG$4,$BI$4)"
    assert sheet.cell(3, 54).api.Formula2 == (
        '=IFERROR(TAKE(Grid_Argmin(UV_WB_S1),,1),"—")'
    )
    assert sheet.cell(3, 62).api.Formula2 == "=Grid_Search_Optimum(UV_WB_S1)"
    assert sheet.cell(4, 62).api.Formula2 is None
    assert sheet.cell(5, 54).api.Formula2 == (
        "=NLL_Weibull(UV_Data,$BF$3,$BF$4,UV_Include)"
    )
    assert sheet.cell(31, 54).api.Formula2 == (
        "=NLL_Gamma(UV_Data,$BF$29,$BF$30,UV_Include)"
    )
    assert "NLL_Beta(z,$BF$55,$BF$56)" in sheet.cell(57, 54).api.Formula2

    assert sheet.cell(3, 81).api.Formula2 == "=MAX(0.001,$BJ$3-$BI$3)"
    assert sheet.cell(3, 82).api.Formula2 == "=$BJ$3+$BI$3"
    assert sheet.cell(4, 81).api.Formula2 == "=MAX(0.001,$BJ$4-$BI$4)"
    assert sheet.cell(4, 82).api.Formula2 == "=$BJ$4+$BI$4"


def test_weibull_grid_uses_visible_inputs_borders_and_boundary_rules() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert sheet.tables[:2] == [
        {
            "range": ((5, 54), (25, 74)),
            "row_input": ((3, 58),),
            "column_input": ((4, 58),),
        },
        {
            "range": ((5, 76), (25, 96)),
            "row_input": ((3, 80),),
            "column_input": ((4, 80),),
        },
    ]
    assert sheet.tables[2:] == [
        {
            "range": ((31, 54), (51, 74)),
            "row_input": ((29, 58),),
            "column_input": ((30, 58),),
        },
        {
            "range": ((31, 76), (51, 96)),
            "row_input": ((29, 80),),
            "column_input": ((30, 80),),
        },
        {
            "range": ((57, 54), (77, 74)),
            "row_input": ((55, 58),),
            "column_input": ((56, 58),),
        },
        {
            "range": ((57, 76), (77, 96)),
            "row_input": ((55, 80),),
            "column_input": ((56, 80),),
        },
    ]

    for address in (
        ((2, 54), (3, 54)),
        ((2, 55), (3, 55)),
        ((2, 57), (4, 62)),
        ((5, 54), (25, 74)),
    ):
        assert set(sheet.range(*address).api._borders) == {7, 8, 9, 10}
    assert sheet.range((2, 56), (4, 56)).api._borders == {}

    shape_rule = sheet.cell(3, 62).api.FormatConditions.items[0].Formula1
    scale_rule = sheet.cell(4, 62).api.FormatConditions.items[0].Formula1
    assert shape_rule == "=OR(INDEX(Grid_Argmin(UV_WB_S1),1,3)=1,INDEX(Grid_Argmin(UV_WB_S1),1,3)=$BC$3)"
    assert scale_rule == "=OR(INDEX(Grid_Argmin(UV_WB_S1),1,2)=1,INDEX(Grid_Argmin(UV_WB_S1),1,2)=$BC$3)"
    assert len(sheet.range((6, 55), (25, 74)).api.FormatConditions.color_scales) == 1


def test_weibull_bounds_and_summary_reference_final_best_cells() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)
    rows = {name: row for row, name, *_ in _dist_rows(5)}
    weibull_row = next(item for item in _dist_rows(5) if item[1] == "Weibull")
    gamma_row = next(item for item in _dist_rows(5) if item[1] == "Gamma")
    beta_row = next(item for item in _dist_rows(5) if item[1] == "Beta")

    assert rows["Weibull"] == weibull_row[0]
    assert weibull_row[3] == "=$CF$3"
    assert weibull_row[5] == "=$CF$4"
    assert rows["Gamma"] == gamma_row[0]
    assert gamma_row[3] == "=$CF$29"
    assert gamma_row[5] == "=$CF$30"
    assert rows["Beta"] == beta_row[0]
    assert beta_row[3] == "=$CF$55"
    assert beta_row[5] == "=$CF$56"
    assert "NLL_Beta" in beta_row[8]
    assert "COUNT(d)*LN(scale_)" in beta_row[8]
    assert sheet.cell(3, 59).color == INPUT_COLOR
    assert sheet.cell(3, 60).color == INPUT_COLOR
    assert sheet.cell(3, 81).color is None
    assert sheet.cell(3, 83).color is None


def test_grid_stage_returns_visible_step_and_count_references() -> None:
    sheet = RecordingSheet()

    refs = _write_grid_stage(
        sheet,
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

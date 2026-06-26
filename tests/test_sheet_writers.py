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
    _STAT_ROWS,
    _dist_rows,
    _setup_local_names,
    _write_data_zone,
    _write_descriptive_stats,
    _write_fitting_table,
    _write_grid_stage,
    _write_histogram_table,
    _write_weibull_grid_search,
    UNIVARIATE_SHEET_NAME,
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

    _write_histogram_table(sheet, 7, 8, "Sturges")

    assert sheet.cell(2, 7).value == "Method"
    assert sheet.cell(2, 7).color == HEADER_COLOR
    assert sheet.cell(2, 8).value == "Sturges"
    assert sheet.cell(2, 8).color == HEADER_COLOR
    assert sheet.cell(3, 7).value == "Bins:"
    assert sheet.cell(3, 8).api.Formula2 == "=num_histogram_bins(UV_Data,H2,UV_Include)"
    assert sheet.cell(5, 7).api.Formula2 == "=Bin_Edges(UV_Data,H2,UV_Include)"
    assert sheet.cell(5, 8).api.Formula2 == "=Bin_Counts(UV_Data,G5#,UV_Include)"


def test_local_name_setup_removes_legacy_globals_and_uses_method_cells() -> None:
    sheet = RecordingSheet(
        global_names=[
            "UV_Data",
            "UV_Include",
            "UV_n",
            "UV_Sturges_Edges",
            "UV_Sturges_Counts",
            "UV_Scott_Edges",
            "UV_Scott_Counts",
            "UV_FD_Edges",
            "UV_FD_Counts",
            "UnrelatedName",
        ]
    )

    _setup_local_names(sheet)

    assert [item.Name for item in sheet.book.api.Names.items] == ["UnrelatedName"]
    names = sheet.api.Names
    assert names.by_short_name("UV_Data").RefersTo == "='Univariate'!$A$4#"
    for hist_name in (
        "UV_Sturges_Edges",
        "UV_Sturges_Counts",
        "UV_Scott_Edges",
        "UV_Scott_Counts",
        "UV_FD_Edges",
        "UV_FD_Counts",
    ):
        assert "num_histogram_bins(" in names.by_short_name(hist_name).RefersTo
    assert "$H$2" in names.by_short_name("UV_Sturges_Edges").RefersTo
    assert "$K$2" in names.by_short_name("UV_Scott_Edges").RefersTo
    assert "$N$2" in names.by_short_name("UV_FD_Edges").RefersTo


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

    _write_histogram_table(sheet, 7, 8, "Sturges")
    assert sheet.cell(3, 8).number_format == "0"
    assert sheet.range((5, 7), (2003, 7)).number_format == "0.0"
    assert sheet.range((5, 8), (2003, 8)).number_format == "0"

    _write_fitting_table(sheet)
    assert sheet.cell(5, 19).number_format == "0.0"
    assert sheet.cell(5, 24).number_format == "0.0E+00"
    assert sheet.cell(5, 25).number_format == "0"
    assert sheet.cell(5, 26).number_format == "0.0"

    _write_weibull_grid_search(sheet)
    assert sheet.cell(3, 30).number_format == "0"
    assert sheet.range((3, 33), (4, 37)).number_format == "0.0"
    assert sheet.cell(5, 29).number_format == "0.0E+00"
    assert sheet.range((5, 30), (5, 49)).number_format == "0.0"
    assert sheet.range((6, 29), (25, 29)).number_format == "0.0"
    assert sheet.range((6, 30), (25, 49)).number_format == "0.0E+00"
    assert sheet.cell(3, 29).number_format == "0.0E+00"


def test_histogram_chart_title_cells_reference_method_headers():
    """Chart title formula cells at Q14/Q34/Q54 reference the correct method header columns."""
    from lambda_catalog.write_sheet_univariate import (
        _write_histogram_chart_title_cells,
        _ROW_CHART1_TITLE, _ROW_CHART2_TITLE, _ROW_CHART3_TITLE,
        _C_Q,
    )
    sheet = RecordingSheet()
    _write_histogram_chart_title_cells(sheet)
    # f() writes via sheet.range(rc(row, col)) → key = ((row, col),)
    f1 = sheet.ranges[((_ROW_CHART1_TITLE, _C_Q),)].state.formula2
    f2 = sheet.ranges[((_ROW_CHART2_TITLE, _C_Q),)].state.formula2
    f3 = sheet.ranges[((_ROW_CHART3_TITLE, _C_Q),)].state.formula2
    assert f1 is not None and "H" in f1
    assert f2 is not None and "K" in f2
    assert f3 is not None and "N" in f3


def test_weibull_grid_search_uses_final_layout_and_named_bodies() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert ((1, 29), (1, 49)) in sheet.merges
    assert ((1, 51), (1, 71)) in sheet.merges
    assert sheet.cell(2, 29).value == "Min NLL:"
    assert sheet.cell(2, 30).value == "Rows/Columns"
    assert sheet.cell(3, 30).value == 20
    assert sheet.cell(2, 31).value is None
    assert [sheet.cell(2, col).value for col in range(32, 38)] == [
        "Parameter", "Input", "Min", "Max", "Step Size", "Best",
    ]
    assert sheet.cell(3, 32).value == "Shape (k)"
    assert sheet.cell(29, 32).value == "Shape (α)"
    assert sheet.cell(30, 32).value == "Rate (β)"
    assert sheet.cell(55, 32).value == "Alpha (α)"
    assert sheet.cell(56, 32).value == "Beta (β)"
    assert sheet.cell(4, 32).value == "Scale (λ)"

    names = sheet.api.Names
    assert names.by_short_name("UV_WB_S1").RefersTo == "='Univariate'!$AD$6:$AW$25"
    assert names.by_short_name("UV_WB_S2").RefersTo == "='Univariate'!$AZ$6:$BS$25"
    assert names.by_short_name("UV_GAMMA_S1").RefersTo == "='Univariate'!$AD$32:$AW$51"
    assert names.by_short_name("UV_GAMMA_S2").RefersTo == "='Univariate'!$AZ$32:$BS$51"
    assert names.by_short_name("UV_BETA_S1").RefersTo == "='Univariate'!$AD$58:$AW$77"
    assert names.by_short_name("UV_BETA_S2").RefersTo == "='Univariate'!$AZ$58:$BS$77"


def test_weibull_grid_formulas_reference_visible_controls() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert sheet.cell(3, 36).api.Formula2 == "=($AI$3-$AH$3)/($AD$3-1)"
    assert sheet.cell(4, 36).api.Formula2 == "=($AI$4-$AH$4)/($AD$3-1)"
    assert sheet.cell(5, 30).api.Formula2 == "=SEQUENCE(1,$AD$3,$AH$3,$AJ$3)"
    assert sheet.cell(6, 29).api.Formula2 == "=SEQUENCE($AD$3,1,$AH$4,$AJ$4)"
    assert sheet.cell(3, 29).api.Formula2 == (
        '=IFERROR(TAKE(Grid_Argmin(UV_WB_S1),,1),"—")'
    )
    assert sheet.cell(3, 37).api.Formula2 == "=Grid_Search_Optimum(UV_WB_S1)"
    assert sheet.cell(4, 37).api.Formula2 is None
    assert sheet.cell(5, 29).api.Formula2 == (
        "=NLL_Weibull(UV_Data,$AG$3,$AG$4,UV_Include)"
    )
    assert sheet.cell(31, 29).api.Formula2 == (
        "=NLL_Gamma(UV_Data,$AG$29,$AG$30,UV_Include)"
    )
    assert "NLL_Beta(z,$AG$55,$AG$56)" in sheet.cell(57, 29).api.Formula2

    assert sheet.cell(3, 56).api.Formula2 == "=MAX(0.001,$AK$3-$AJ$3)"
    assert sheet.cell(3, 57).api.Formula2 == "=$AK$3+$AJ$3"
    assert sheet.cell(4, 56).api.Formula2 == "=MAX(0.001,$AK$4-$AJ$4)"
    assert sheet.cell(4, 57).api.Formula2 == "=$AK$4+$AJ$4"


def test_weibull_grid_uses_visible_inputs_borders_and_boundary_rules() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert sheet.tables[:2] == [
        {
            "range": ((5, 29), (25, 49)),
            "row_input": ((3, 33),),
            "column_input": ((4, 33),),
        },
        {
            "range": ((5, 51), (25, 71)),
            "row_input": ((3, 55),),
            "column_input": ((4, 55),),
        },
    ]
    assert sheet.tables[2:] == [
        {
            "range": ((31, 29), (51, 49)),
            "row_input": ((29, 33),),
            "column_input": ((30, 33),),
        },
        {
            "range": ((31, 51), (51, 71)),
            "row_input": ((29, 55),),
            "column_input": ((30, 55),),
        },
        {
            "range": ((57, 29), (77, 49)),
            "row_input": ((55, 33),),
            "column_input": ((56, 33),),
        },
        {
            "range": ((57, 51), (77, 71)),
            "row_input": ((55, 55),),
            "column_input": ((56, 55),),
        },
    ]

    for address in (
        ((2, 29), (3, 29)),
        ((2, 30), (3, 30)),
        ((2, 32), (4, 37)),
        ((5, 29), (25, 49)),
    ):
        assert set(sheet.range(*address).api._borders) == {7, 8, 9, 10}
    assert sheet.range((2, 31), (4, 31)).api._borders == {}

    shape_rule = sheet.cell(3, 37).api.FormatConditions.items[0].Formula1
    scale_rule = sheet.cell(4, 37).api.FormatConditions.items[0].Formula1
    assert shape_rule == "=OR(INDEX(Grid_Argmin(UV_WB_S1),1,3)=1,INDEX(Grid_Argmin(UV_WB_S1),1,3)=$AD$3)"
    assert scale_rule == "=OR(INDEX(Grid_Argmin(UV_WB_S1),1,2)=1,INDEX(Grid_Argmin(UV_WB_S1),1,2)=$AD$3)"
    assert len(sheet.range((6, 30), (25, 49)).api.FormatConditions.color_scales) == 1


def test_weibull_bounds_and_summary_reference_final_best_cells() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)
    rows = {name: row for row, name, *_ in _dist_rows(5)}
    weibull_row = next(item for item in _dist_rows(5) if item[1] == "Weibull")
    gamma_row = next(item for item in _dist_rows(5) if item[1] == "Gamma")
    beta_row = next(item for item in _dist_rows(5) if item[1] == "Beta")

    assert rows["Weibull"] == weibull_row[0]
    assert weibull_row[3] == "=$BG$3"
    assert weibull_row[5] == "=$BG$4"
    assert rows["Gamma"] == gamma_row[0]
    assert gamma_row[3] == "=$BG$29"
    assert gamma_row[5] == "=$BG$30"
    assert rows["Beta"] == beta_row[0]
    assert beta_row[3] == "=$BG$55"
    assert beta_row[5] == "=$BG$56"
    assert "NLL_Beta" in beta_row[8]
    assert "COUNT(d)*LN(scale_)" in beta_row[8]
    assert sheet.cell(3, 34).color == INPUT_COLOR
    assert sheet.cell(3, 35).color == INPUT_COLOR
    assert sheet.cell(3, 56).color is None
    assert sheet.cell(3, 58).color is None


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

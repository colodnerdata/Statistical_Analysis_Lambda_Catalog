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


def test_betapert_gof_cdf_uses_valid_let_names() -> None:
    betapert = next(row for row in _dist_rows(5) if row[1] == "BetaPERT")
    cdf_expr = betapert[-1]

    assert "a1," not in cdf_expr
    assert "a2," not in cdf_expr
    assert "alpha_p," in cdf_expr
    assert "beta_p," in cdf_expr


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
    assert sheet.cell(3, 32).number_format == "0"
    assert sheet.range((3, 35), (4, 39)).number_format == "0.0"
    assert sheet.cell(5, 31).number_format == "0.0E+00"
    assert sheet.range((5, 32), (5, 51)).number_format == "0.0"
    assert sheet.range((6, 31), (25, 31)).number_format == "0.0"
    assert sheet.range((6, 32), (25, 51)).number_format == "0.0E+00"
    assert sheet.cell(3, 31).number_format == "0.0E+00"


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
    assert f1 == '=H2&" Method Histogram"'
    assert f2 == '=K2&" Method Histogram"'
    assert f3 == '=N2&" Method Histogram"'


def test_weibull_grid_search_uses_final_layout_and_named_bodies() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert ((1, 31), (1, 51)) in sheet.merges
    assert ((1, 53), (1, 73)) in sheet.merges
    assert sheet.cell(2, 31).value == "Min NLL:"
    assert sheet.cell(2, 32).value == "Rows/Columns"
    assert sheet.cell(3, 32).value == 20
    assert sheet.cell(2, 33).value is None
    assert [sheet.cell(2, col).value for col in range(34, 40)] == [
        "Parameter", "Input", "Min", "Max", "Step Size", "Best",
    ]
    assert sheet.cell(3, 34).value == "Shape (k)"
    assert sheet.cell(29, 34).value == "Shape (α)"
    assert sheet.cell(30, 34).value == "Rate (β)"
    assert sheet.cell(55, 34).value == "Alpha (α)"
    assert sheet.cell(56, 34).value == "Beta (β)"
    assert sheet.cell(4, 34).value == "Scale (λ)"

    names = sheet.api.Names
    assert names.by_short_name("UV_WB_S1").RefersTo == "='Univariate'!$AF$6:$AY$25"
    assert names.by_short_name("UV_WB_S2").RefersTo == "='Univariate'!$BB$6:$BU$25"
    assert names.by_short_name("UV_GAMMA_S1").RefersTo == "='Univariate'!$AF$32:$AY$51"
    assert names.by_short_name("UV_GAMMA_S2").RefersTo == "='Univariate'!$BB$32:$BU$51"
    assert names.by_short_name("UV_BETA_S1").RefersTo == "='Univariate'!$AF$58:$AY$77"
    assert names.by_short_name("UV_BETA_S2").RefersTo == "='Univariate'!$BB$58:$BU$77"


def test_weibull_grid_formulas_reference_visible_controls() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert sheet.cell(3, 38).api.Formula2 == "=($AK$3-$AJ$3)/($AF$3-1)"
    assert sheet.cell(4, 38).api.Formula2 == "=($AK$4-$AJ$4)/($AF$3-1)"
    assert sheet.cell(5, 32).api.Formula2 == "=SEQUENCE(1,$AF$3,$AJ$3,$AL$3)"
    assert sheet.cell(6, 31).api.Formula2 == "=SEQUENCE($AF$3,1,$AJ$4,$AL$4)"
    assert sheet.cell(3, 31).api.Formula2 == (
        '=IFERROR(TAKE(Grid_Argmin(UV_WB_S1),,1),"—")'
    )
    assert sheet.cell(3, 39).api.Formula2 == "=Grid_Search_Optimum(UV_WB_S1)"
    assert sheet.cell(4, 39).api.Formula2 is None
    assert sheet.cell(5, 31).api.Formula2 == (
        "=NLL_Weibull(UV_Data,$AI$3,$AI$4,UV_Include)"
    )
    assert sheet.cell(31, 31).api.Formula2 == (
        "=NLL_Gamma(UV_Data,$AI$29,$AI$30,UV_Include)"
    )
    assert "NLL_Beta(z,$AI$55,$AI$56)" in sheet.cell(57, 31).api.Formula2

    assert sheet.cell(3, 58).api.Formula2 == "=MAX(0.001,$AM$3-$AL$3)"
    assert sheet.cell(3, 59).api.Formula2 == "=$AM$3+$AL$3"
    assert sheet.cell(4, 58).api.Formula2 == "=MAX(0.001,$AM$4-$AL$4)"
    assert sheet.cell(4, 59).api.Formula2 == "=$AM$4+$AL$4"


def test_weibull_grid_uses_visible_inputs_borders_and_boundary_rules() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)

    assert sheet.tables[:2] == [
        {
            "range": ((5, 31), (25, 51)),
            "row_input": ((3, 35),),
            "column_input": ((4, 35),),
        },
        {
            "range": ((5, 53), (25, 73)),
            "row_input": ((3, 57),),
            "column_input": ((4, 57),),
        },
    ]
    assert sheet.tables[2:] == [
        {
            "range": ((31, 31), (51, 51)),
            "row_input": ((29, 35),),
            "column_input": ((30, 35),),
        },
        {
            "range": ((31, 53), (51, 73)),
            "row_input": ((29, 57),),
            "column_input": ((30, 57),),
        },
        {
            "range": ((57, 31), (77, 51)),
            "row_input": ((55, 35),),
            "column_input": ((56, 35),),
        },
        {
            "range": ((57, 53), (77, 73)),
            "row_input": ((55, 57),),
            "column_input": ((56, 57),),
        },
    ]

    for address in (
        ((2, 31), (3, 31)),
        ((2, 32), (3, 32)),
        ((2, 34), (4, 39)),
        ((5, 31), (25, 51)),
    ):
        assert set(sheet.range(*address).api._borders) == {7, 8, 9, 10}
    assert sheet.range((2, 33), (4, 33)).api._borders == {}

    shape_rule = sheet.cell(3, 39).api.FormatConditions.items[0].Formula1
    scale_rule = sheet.cell(4, 39).api.FormatConditions.items[0].Formula1
    assert shape_rule == "=OR(INDEX(Grid_Argmin(UV_WB_S1),1,3)=1,INDEX(Grid_Argmin(UV_WB_S1),1,3)=$AF$3)"
    assert scale_rule == "=OR(INDEX(Grid_Argmin(UV_WB_S1),1,2)=1,INDEX(Grid_Argmin(UV_WB_S1),1,2)=$AF$3)"
    assert len(sheet.range((6, 32), (25, 51)).api.FormatConditions.color_scales) == 1


def test_weibull_bounds_and_summary_reference_final_best_cells() -> None:
    sheet = RecordingSheet()

    _write_weibull_grid_search(sheet)
    rows = {name: row for row, name, *_ in _dist_rows(5)}
    weibull_row = next(item for item in _dist_rows(5) if item[1] == "Weibull")
    gamma_row = next(item for item in _dist_rows(5) if item[1] == "Gamma")
    beta_row = next(item for item in _dist_rows(5) if item[1] == "Beta")

    assert rows["Weibull"] == weibull_row[0]
    assert weibull_row[3] == "=$BI$3"
    assert weibull_row[5] == "=$BI$4"
    assert rows["Gamma"] == gamma_row[0]
    assert gamma_row[3] == "=$BI$29"
    assert gamma_row[5] == "=$BI$30"
    assert rows["Beta"] == beta_row[0]
    assert beta_row[3] == "=$BI$55"
    assert beta_row[5] == "=$BI$56"
    assert "NLL_Beta" in beta_row[8]
    assert "COUNT(d)*LN(scale_)" in beta_row[8]
    assert sheet.cell(3, 36).color == INPUT_COLOR
    assert sheet.cell(3, 37).color == INPUT_COLOR
    assert sheet.cell(3, 58).color is None
    assert sheet.cell(3, 60).color is None


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

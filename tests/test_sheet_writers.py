"""Tests for workbook-writing logic that do not require a live Excel process."""
from lambda_catalog.sheet_styles import CF_DARK_RED_TEXT, CF_LIGHT_RED_FILL, HEADER_COLOR
from lambda_catalog.workbook_helpers import add_expression_format, excel_color
from lambda_catalog.write_sheet_regression import (
    _setup_local_names as _setup_regression_names,
    _write_prediction_interval,
)
from lambda_catalog.write_sheet_mlr_scalar_test import _actual_formula
from lambda_catalog.write_sheet_univariate import (
    _STAT_ROWS,
    _setup_local_names,
    _write_data_zone,
    _write_descriptive_stats,
    _write_fitting_table,
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
    assert sheet.range((5, 32), (6, 35)).number_format == "0.0"
    assert sheet.cell(7, 29).number_format == "0.0E+00"
    assert sheet.range((7, 30), (7, 49)).number_format == "0.0"
    assert sheet.range((8, 29), (27, 29)).number_format == "0.0"
    assert sheet.range((8, 30), (27, 49)).number_format == "0.0E+00"
    assert sheet.cell(5, 29).number_format == "0.0E+00"

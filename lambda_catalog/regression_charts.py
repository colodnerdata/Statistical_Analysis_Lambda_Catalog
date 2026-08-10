"""The 7 diagnostic charts for the Regression template.

Extracted from ``write_sheet_regression.py`` (reunify-workbook Part 5.1) so the
writer is its zone writers + orchestrator. These three functions build the chart
*content* only; the chart *constants* (``_XL_*``, ``_CHART_*``,
``_CHART_Y_TICK_FORMATS``, the label/anchor columns) and the layout anchors the
title formulas reference live in ``regression_layout.py``. The functions depend
on the layout constants, the standard cell-write helpers, the header fill
color, and the sheet object -- nothing else -- so importing them here creates no
cycle back into ``write_sheet_regression``.

``write_sheet_regression`` re-exports all three names so existing importers
(notably the unit tests, and the call sites in ``write_regression_output_sheet``)
keep working unchanged.
"""
from __future__ import annotations

from typing import Any

import xlwings as xw

from .regression_layout import (
    REGRESSION_SHEET_NAME,
    _A_ADJUSTED_R_SQUARED,
    _A_MEAN_LEVERAGE,
    _A_PRESS,
    _A_QQ_CORRELATION,
    _A_RESPONSE_READOUT,
    _CHART_GAP,
    _CHART_HEIGHT,
    _CHART_WIDTH,
    _CHART_Y_TICK_FORMAT_DEFAULT,
    _CHART_Y_TICK_FORMATS,
    _C_BB,
    _C_CHART_LABEL_NAME,
    _C_CHART_TITLE,
    _C_CHART_XLABEL,
    _C_CHART_YLABEL,
    _COOKS_CUTOFF,
    _ROW_CHART_LABELS,
    _XL_CATEGORY,
    _XL_COLUMN_CLUSTERED,
    _XL_LINE,
    _XL_VALUE,
    _XL_XY_SCATTER,
    _XL_XY_SCATTER_LINES_NO_MARKERS,
)
from .sheet_styles import HEADER_COLOR as _HEADER
from .workbook_helpers import a1, col_letter, excel_color, f, val


def _diagnostic_chart_specs(
    sheet_name: str = REGRESSION_SHEET_NAME,
) -> list[tuple[str, str, str | None, str, str, str, str, int, int]]:
    """Static spec for the 7 regression diagnostic charts.

    Each tuple is (key, chart_type, x_addr, y_addr, title_formula,
    x_label_formula, y_label_formula, grid_row, grid_col).

    `key` is a stable internal identifier — used for series naming, the
    gridline-mode lookup, and per-chart branching in `_write_diagnostic_charts`
    — and is never itself displayed. `title_formula` / `x_label_formula` /
    `y_label_formula` are Excel formulas (written verbatim into the chart's
    label cells by `_write_chart_label_cells`) that MAY reference live sheet
    statistics, so the displayed title can vary with the fitted model while
    `key` stays fixed.

    ``sheet_name`` qualifies every named-range and cell reference the specs
    emit. Chart SERIES formulas live above the sheet layer, so they must
    carry the sheet prefix even for worksheet-scoped names — which means a
    spec built for the wrong sheet name points a chart at another sheet's
    data and still parses. It defaults to the production sheet; generated
    test-model sheets pass their own.
    """

    def _name_ref(local_name: str) -> str:
        return f"='{sheet_name}'!{local_name}"

    return [
        (
            "Residuals vs. Fitted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartResid"),
            f'="Residuals vs. Fitted — "&{_A_RESPONSE_READOUT}',
            '="Fitted Values"', '="Residuals"',
            1, 1,
        ),
        (
            "Normal Q-Q", "scatter",
            _name_ref("RegChartQQX"),
            _name_ref("RegChartQQY"),
            f'="Normal Q-Q  (r = "&TEXT({_A_QQ_CORRELATION},"0.000")&")"'
            f'&IF({_A_QQ_CORRELATION}<0.95,"  — check normality","")',
            '="Theoretical Quantiles"', '="Studentized Residuals"',
            1, 2,
        ),
        (
            "Actual vs. Predicted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartActY"),
            f'="Actual vs. Predicted — "&{_A_RESPONSE_READOUT}'
            f'&"  (Adj. R² = "&TEXT({_A_ADJUSTED_R_SQUARED},"0.000")&")"',
            f'="Predicted "&{_A_RESPONSE_READOUT}',
            f'="Actual "&{_A_RESPONSE_READOUT}',
            2, 1,
        ),
        (
            "Scale-Location", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartScaleLoc"),
            '="Scale-Location"',
            '="Fitted Values"', '="√|Studentized Residual|"',
            2, 2,
        ),
        (
            "Cook's Distance", "bar",
            None,
            _name_ref("RegChartCookDist"),
            # The IFERROR is around TEXT, not inside _COOKS_CUTOFF. The cutoff
            # is deliberately NA() in the zero-predictor state so every
            # COMPARISON against it fails closed, but TEXT(NA(),…) is #N/A and
            # would propagate through the concatenation, leaving the whole
            # chart title rendering as "#N/A". The em dash is the same
            # not-available token the Univariate fit tables use.
            '="Cook\'s Distance  (flag: D > "'
            f'&IFERROR(TEXT({_COOKS_CUTOFF},"0.000"),"—")&")"',
            '="Observation"', '="Cook\'s Distance"',
            3, 1,
        ),
        (
            "Studentized Residuals vs. Leverage", "scatter",
            _name_ref("RegChartLeverage"),
            _name_ref("RegChartStudResid"),
            '="Studentized Residuals vs. Leverage  (mean leverage = "'
            f'&TEXT({_A_MEAN_LEVERAGE},"0.000")&")"',
            '="Leverage (Hat Diagonal)"', '="Studentized Residuals"',
            3, 2,
        ),
        (
            "PRESS Residuals", "bar",
            None,
            _name_ref("RegChartPRESSResid"),
            f'="PRESS Residuals  (PRESS = "&TEXT({_A_PRESS},"#,##0")&")"',
            '="Observation"', '="PRESS Residual"',
            4, 1,
        ),
    ]


def _write_chart_label_cells(sheet: xw.Sheet) -> None:
    """Write the Chart Title / X-Axis Title / Y-Axis Title formula cells for
    the 7 diagnostic charts, one row per chart starting at _ROW_CHART_LABELS.

    `_write_diagnostic_charts` binds each chart's title and axis titles to
    these cells via `.Formula` rather than embedding label strings directly
    into the chart-construction call, so tuning a chart's label is a formula
    edit here, not a change to the COM chart-building code. Plain cell
    writes only (no chart/COM API), so this is exercised directly in unit
    tests via `RecordingSheet` without Excel.
    """
    for i, spec in enumerate(_diagnostic_chart_specs(sheet.name)):
        key, _chart_type, _x_addr, _y_addr, title_formula, x_label_formula, y_label_formula, _grid_row, _grid_col = spec
        row = _ROW_CHART_LABELS + i
        val(sheet, row, _C_CHART_LABEL_NAME, key)
        f(sheet, row, _C_CHART_TITLE, title_formula)
        f(sheet, row, _C_CHART_XLABEL, x_label_formula)
        f(sheet, row, _C_CHART_YLABEL, y_label_formula)


def _write_diagnostic_charts(sheet: xw.Sheet) -> None:  # pylint: disable=too-many-locals,too-many-statements
    """Create 7 pre-built diagnostic charts to the right of the Residual Output section."""
    start_left = sheet.range(a1(1, _C_BB)).left
    start_top = sheet.range("A3").top

    col_step = _CHART_WIDTH + _CHART_GAP
    row_step = _CHART_HEIGHT + _CHART_GAP

    def _pos(grid_row: int, grid_col: int) -> tuple[float, float]:
        return (
            start_left + (grid_col - 1) * col_step,
            start_top + (grid_row - 1) * row_step,
        )

    # The LIVE sheet's name, not REGRESSION_SHEET_NAME: this writer runs on
    # generated test-model sheets too, and a hardcoded constant here would
    # point every one of their charts at the production Regression sheet.
    sname = sheet.name

    def _name_ref(local_name: str) -> str:
        return f"='{sname}'!{local_name}"

    def _label_ref(col: int, row: int) -> str:
        return f"='{sname}'!${col_letter(col)}${row}"

    chart_specs = _diagnostic_chart_specs(sname)

    # Per-chart gridline strategy:
    # - Use Y major gridlines for residual magnitude judgment on residual and bar charts.
    # - Use both axes on comparative scatter plots where position relative to both axes matters.
    gridline_modes = {
        "Residuals vs. Fitted": "none",
        "Normal Q-Q": "both",
        "Actual vs. Predicted": "both",
        "Scale-Location": "y",
        "Cook's Distance": "y",
        "Studentized Residuals vs. Leverage": "both",
        "PRESS Residuals": "y",
    }

    def _add_identity_line(chart: Any, name_ref: str) -> None:
        """Add a dotted y=x reference series using one column for both axes.

        Pointing XValues and Values at the same named range guarantees every
        plotted point sits exactly on the identity line — a real data series
        stays correct if the chart is resized, moved, or the axis scaling
        changes, unlike a shape drawn at fixed plot-area pixel coordinates.
        """
        series = chart.SeriesCollection().NewSeries()
        series.XValues = name_ref
        series.Values = name_ref
        series.Name = "Identity"
        series.ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS
        series.Format.Line.ForeColor.RGB = excel_color((120, 120, 120))
        series.Format.Line.DashStyle = 3  # msoLineRoundDot
        series.Format.Line.Weight = 1.25

    for i, spec in enumerate(chart_specs):
        (key, chart_type, x_addr, y_addr, _title_formula,
         _x_label_formula, _y_label_formula, grid_row, grid_col) = spec
        label_row = _ROW_CHART_LABELS + i
        left, top = _pos(grid_row, grid_col)
        co = sheet.api.ChartObjects().Add(left, top, _CHART_WIDTH, _CHART_HEIGHT)
        chart = co.Chart

        chart.ChartType = _XL_XY_SCATTER if chart_type == "scatter" else _XL_COLUMN_CLUSTERED

        sc = chart.SeriesCollection()
        for j in range(sc.Count, 0, -1):
            sc.Item(j).Delete()

        series = chart.SeriesCollection().NewSeries()
        if x_addr is not None:
            series.XValues = x_addr
        series.Values = y_addr
        series.Name = key
        # Bar charts (Cook's Distance, PRESS Residuals) have no markers to resize.
        if chart_type == "scatter":
            series.MarkerSize = 4

        # All charts: Header-style title (bold, 14 pt, light-blue fill).
        # Title and both axis titles are bound to the formula cells written by
        # _write_chart_label_cells (row `label_row`, cols _C_CHART_TITLE /
        # _C_CHART_XLABEL / _C_CHART_YLABEL) rather than set from a literal
        # string, so their content can reference live sheet statistics.
        chart.HasLegend = False
        chart.HasTitle = True
        chart.ChartTitle.Formula = _label_ref(_C_CHART_TITLE, label_row)
        chart.ChartTitle.Font.Bold = True
        chart.ChartTitle.Font.Size = 14
        chart.ChartTitle.Format.Fill.Visible = True
        chart.ChartTitle.Format.Fill.Solid()
        chart.ChartTitle.Format.Fill.ForeColor.RGB = excel_color(_HEADER)

        x_axis = chart.Axes(_XL_CATEGORY)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Formula = _label_ref(_C_CHART_XLABEL, label_row)
        x_axis.TickLabels.NumberFormat = "0"

        y_axis = chart.Axes(_XL_VALUE)
        y_axis.HasTitle = True
        y_axis.AxisTitle.Formula = _label_ref(_C_CHART_YLABEL, label_row)
        # One unconditional assignment from the per-chart table, so a chart's
        # y-axis format is whatever _CHART_Y_TICK_FORMATS says it is.
        y_tick_format = _CHART_Y_TICK_FORMATS.get(key, _CHART_Y_TICK_FORMAT_DEFAULT)
        y_axis.TickLabels.NumberFormat = y_tick_format

        gridline_mode = gridline_modes.get(key, "none")
        x_axis.HasMajorGridlines = gridline_mode == "both"
        y_axis.HasMajorGridlines = gridline_mode in {"y", "both"}

        if key == "Cook's Distance":
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone

            # Overlay series for selective data labels: NA()'d rows in
            # RegChartCookDistFlag plot/label nothing, so only points past
            # the F(0.5, p, n−p) cutoff get a marker+label. ChartType=xlLine
            # (rather than the chart's own xlColumnClustered) keeps this
            # series off the bar cluster — sharing the category axis
            # without narrowing/shifting the real bars — which makes this
            # a Column+Line combo chart.
            flag_series = chart.SeriesCollection().NewSeries()
            flag_series.XValues = _name_ref("RegChartObsLabel")
            flag_series.Values = _name_ref("RegChartCookDistFlag")
            flag_series.ChartType = _XL_LINE
            flag_series.Name = "Flagged (D > F(0.5, p, n-p))"
            flag_series.Format.Line.Visible = False  # msoFalse — no connecting line
            flag_series.MarkerStyle = -4142          # xlMarkerStyleNone — label only
            flag_series.HasDataLabels = True
            dls = flag_series.DataLabels()
            dls.ShowCategoryName = True  # observation identifier, e.g. "United States"
            dls.ShowValue = True         # ...plus the Cook's D value
            dls.NumberFormat = y_tick_format  # same format as the y-axis ticks
            dls.Position = 0              # xlLabelPositionAbove
        if key == "Studentized Residuals vs. Leverage":
            x_axis.TickLabels.NumberFormat = "0.00"
        if key == "PRESS Residuals":
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone

        # Both identity-line charts leave axis limits at Excel's defaults. The
        # reference line is a real data series with XValues and Values pointed
        # at the same range, so every point sits on y=x whatever the axes do —
        # forcing the two scales equal only made it render at a visual 45°, and
        # it did so from values read back via Evaluate() during the
        # sheet-writing phase, which runs under XL_CALCULATION_MANUAL and so
        # sees stale or unfit numbers.
        if key == "Normal Q-Q":
            if x_addr is None:
                raise AssertionError("Normal Q-Q chart requires an x-axis range")
            identity_ref: str = x_addr
            _add_identity_line(chart, identity_ref)
        if key == "Actual vs. Predicted":
            if x_addr is None:
                raise AssertionError("Actual vs. Predicted chart requires an x-axis range")
            identity_ref = x_addr
            _add_identity_line(chart, identity_ref)
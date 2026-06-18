"""
write_sheet_univariate.py
Writes the Univariate Analysis sheet into any target workbook.

Design decision — fitting approach
───────────────────────────────────
MLE throughout.  For Normal, Lognormal, and Exponential the MLE estimators are
closed-form (sample mean/SD on the raw or log-transformed data).  For
Triangular and BetaPERT the likelihood is non-differentiable at the mode, so
direct min/mode/max estimation is used — the result is a valid parameter set
that the NLL formula evaluates against.  Grid-search MLE for the two-parameter
shape distributions (Weibull, Gamma, Beta) is deferred to v2.1.

Sheet layout (four horizontal zones)
──────────────────────────────────────
  Col A          — Data input: row-3 header, data in A4:A2003 (2000 rows)
  Col B          — thin gap (width 2)
  Col C–D        — Descriptive Statistics: 12 stat label/value rows
  Col E          — thin gap (width 2)
  Col F–G        — Sturges histogram table (edge | count)
  Col H          — thin gap (width 2)
  Col I–J        — Scott histogram table
  Col K          — thin gap (width 2)
  Col L–M        — Freedman-Diaconis histogram table
  Col N          — thin gap (width 2)
  Col O–Z        — Distribution Fitting summary table
                   P: distribution name | Q: θ₁ label | R: θ₁ value |
                   S: θ₂ label | T: θ₂ value | U: θ₃ label | V: θ₃ value |
                   W: NLL | X: k | Y: AIC | Z: BIC

Sheet-scoped named ranges
─────────────────────────
  UV_Data        — data input column  $A$4:$A$2003
  UV_n           — COUNT(UV_Data)  — scalar numeric count
  UV_Sturges_Edges, UV_Sturges_Counts — OFFSET-based chart series ranges
  UV_Scott_Edges,   UV_Scott_Counts
  UV_FD_Edges,      UV_FD_Counts
"""
from __future__ import annotations

from typing import Any

import xlwings as xw

# ── Colors ────────────────────────────────────────────────────────────────────
_HEADER = (202, 237, 251)    # section headings (same as Regression sheet)
_INPUT  = (251, 226, 213)    # user-editable cells (orange)
_SUBHDR = (220, 230, 241)    # sub-section / column header row

# ── Column indices (1-based) ─────────────────────────────────────────────────

# Zone 1: Data Input
_C_A = 1    # data values

# Zone 2: Descriptive Statistics
_C_C = 3    # stat labels
_C_D = 4    # stat values

# Zone 3: Histogram Tables
_C_F = 6    # Sturges edges
_C_G = 7    # Sturges counts
_C_I = 9    # Scott edges
_C_J = 10   # Scott counts
_C_L = 12   # FD edges
_C_M = 13   # FD counts

# Zone 4: Distribution Fitting
_C_P = 16   # distribution name
_C_Q = 17   # θ₁ label
_C_R = 18   # θ₁ value
_C_S = 19   # θ₂ label
_C_T = 20   # θ₂ value
_C_U = 21   # θ₃ label
_C_V = 22   # θ₃ value
_C_W = 23   # NLL
_C_X = 24   # k (param count)
_C_Y = 25   # AIC
_C_Z = 26   # BIC

# ── Row anchors ───────────────────────────────────────────────────────────────
_ROW_TITLE        = 1   # "Univariate Analysis"
_ROW_SECTION_HDR  = 3   # section headings across all zones
_ROW_COL_HDRS     = 4   # column sub-headers (histograms, fitting table)
_ROW_DATA_START   = 4   # data values begin (col A)
_ROW_STATS_START  = 4   # descriptive stat rows begin (col C–D)
_ROW_HIST_START   = 5   # histogram spill formulas begin
_ROW_DIST_START   = 5   # distribution fitting rows begin

# ── Data capacity ─────────────────────────────────────────────────────────────
_DATA_ROWS = 2000
_DATA_END  = _ROW_DATA_START + _DATA_ROWS - 1   # last data row = 2003

# ── Chart constants ───────────────────────────────────────────────────────────
_XL_COLUMN_CLUSTERED = 51   # xlColumnClustered
_XL_CATEGORY         = 1    # horizontal axis type
_XL_VALUE            = 2    # vertical axis type
_CHART_WIDTH         = 280.0
_CHART_HEIGHT        = 220.0
_CHART_GAP_H         = 10.0
_CHART_GAP_V         = 10.0

UNIVARIATE_SHEET_NAME = "Univariate"


# ── Cell helpers ──────────────────────────────────────────────────────────────

def _rc(row: int, col: int) -> tuple[int, int]:
    return (row, col)


def _col_letter(col: int) -> str:
    result = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        result = chr(ord("A") + rem) + result
    return result


def _a1(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"


def _val(sheet: xw.Sheet, row: int, col: int, value: object) -> None:
    sheet.range(_rc(row, col)).value = value


def _f(sheet: xw.Sheet, row: int, col: int, formula: str) -> None:
    sheet.range(_rc(row, col)).api.Formula2 = formula


def _bold(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(_rc(row, col)).api.Font.Bold = True


def _bold_row(sheet: xw.Sheet, row: int, c1: int, c2: int) -> None:
    sheet.range(_rc(row, c1), _rc(row, c2)).api.Font.Bold = True


def _section_heading(sheet: xw.Sheet, row: int, col: int, label: str) -> None:
    _val(sheet, row, col, label)
    sheet.range(_rc(row, col)).api.Font.Bold = True
    sheet.range(_rc(row, col)).color = _HEADER


def _subheader_row(sheet: xw.Sheet, row: int, c1: int, c2: int) -> None:
    sheet.range(_rc(row, c1), _rc(row, c2)).color = _SUBHDR
    sheet.range(_rc(row, c1), _rc(row, c2)).api.Font.Bold = True


def _format_input(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(_rc(row, col)).color = _INPUT


def _border_box(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    rng = sheet.range(_rc(r1, c1), _rc(r2, c2)).api
    for edge in [7, 8, 9, 10]:
        rng.Borders(edge).LineStyle = 1
        rng.Borders(edge).Weight = 2


# ── Column widths ─────────────────────────────────────────────────────────────

def _set_column_widths(sheet: xw.Sheet) -> None:
    widths = {
        _C_A: 14,    # data
        2:    2,     # gap
        _C_C: 16,    # stat labels
        _C_D: 12,    # stat values
        5:    2,     # gap
        _C_F: 12,    # Sturges edge
        _C_G: 10,    # Sturges count
        8:    2,     # gap
        _C_I: 12,    # Scott edge
        _C_J: 10,    # Scott count
        11:   2,     # gap
        _C_L: 12,    # FD edge
        _C_M: 10,    # FD count
        14:   2,     # gap
        15:   2,     # gap
        _C_P: 14,    # distribution name
        _C_Q: 10,    # θ₁ label
        _C_R: 12,    # θ₁ value
        _C_S: 10,    # θ₂ label
        _C_T: 12,    # θ₂ value
        _C_U: 10,    # θ₃ label
        _C_V: 12,    # θ₃ value
        _C_W: 12,    # NLL
        _C_X: 6,     # k
        _C_Y: 12,    # AIC
        _C_Z: 12,    # BIC
    }
    for col, w in widths.items():
        sheet.range(_rc(1, col), _rc(1, col)).column_width = w


# ── Sheet-scoped named range management ──────────────────────────────────────

def _drop_local_name(sheet: xw.Sheet, name: str) -> None:
    for idx in range(sheet.api.Names.Count, 0, -1):
        local = sheet.api.Names(idx).Name.split("!", 1)[-1]
        if local.lower() == name.lower():
            sheet.api.Names(idx).Delete()


def _setup_local_names(sheet: xw.Sheet) -> None:
    """Register sheet-scoped named ranges used by formulas and charts."""
    sname = sheet.name
    data_col = _col_letter(_C_A)

    # UV_Data: full input column (includes blanks; LAMBDAs filter via ISNUMBER)
    _drop_local_name(sheet, "UV_Data")
    sheet.api.Names.Add(
        Name="UV_Data",
        RefersTo=f"='{sname}'!${data_col}${_ROW_DATA_START}:${data_col}${_DATA_END}",
    )

    # UV_n: count of numeric entries in the data column
    _drop_local_name(sheet, "UV_n")
    sheet.api.Names.Add(
        Name="UV_n",
        RefersTo=f"=COUNT('{sname}'!${data_col}${_ROW_DATA_START}:${data_col}${_DATA_END})",
    )

    # Chart series ranges: OFFSET-based, sized by calling the bin-count
    # LAMBDAs directly inside the named-range formula.  This avoids hidden
    # helper cells that would conflict with the descriptive stats column.
    for name, col_ltr, start_row, size_formula in [
        ("UV_Sturges_Edges",  _col_letter(_C_F), _ROW_HIST_START, "Sturges_Bins(UV_Data)"),
        ("UV_Sturges_Counts", _col_letter(_C_G), _ROW_HIST_START, "Sturges_Bins(UV_Data)"),
        ("UV_Scott_Edges",    _col_letter(_C_I), _ROW_HIST_START, "Scott_Bins(UV_Data)"),
        ("UV_Scott_Counts",   _col_letter(_C_J), _ROW_HIST_START, "Scott_Bins(UV_Data)"),
        ("UV_FD_Edges",       _col_letter(_C_L), _ROW_HIST_START, "FD_Bins(UV_Data)"),
        ("UV_FD_Counts",      _col_letter(_C_M), _ROW_HIST_START, "FD_Bins(UV_Data)"),
    ]:
        _drop_local_name(sheet, name)
        sheet.api.Names.Add(
            Name=name,
            RefersTo=(
                f"=OFFSET('{sname}'!${col_ltr}${start_row},"
                f"0,0,{size_formula},1)"
            ),
        )


# ── Zone 1: data input column ─────────────────────────────────────────────────

def _write_data_zone(sheet: xw.Sheet) -> None:
    _section_heading(sheet, _ROW_SECTION_HDR, _C_A, "Data")
    # Shade entire input range so the user sees where to paste
    sheet.range(
        _rc(_ROW_DATA_START, _C_A), _rc(_DATA_END, _C_A)
    ).color = _INPUT
    _border_box(sheet, _ROW_SECTION_HDR, _C_A, _DATA_END, _C_A)


# ── Zone 2: descriptive statistics ───────────────────────────────────────────

# Stat label, formula (using Excel functions; UV_Data is the named range)
_STAT_ROWS: list[tuple[str, str]] = [
    ("Mean",      "=AVERAGE(UV_Data)"),
    ("Median",    "=MEDIAN(UV_Data)"),
    ("Mode",      "=IFERROR(MODE.SNGL(UV_Data),\"N/A\")"),
    ("Std Dev",   "=STDEV.S(UV_Data)"),
    ("Variance",  "=VAR.S(UV_Data)"),
    ("Min",       "=MIN(UV_Data)"),
    ("Max",       "=MAX(UV_Data)"),
    ("Range",     "=MAX(UV_Data)-MIN(UV_Data)"),
    ("Skewness",  "=INDEX(Skewness(UV_Data),1)"),
    ("Kurtosis",  "=INDEX(Kurtosis(UV_Data),1)"),
    ("Count",     "=COUNT(UV_Data)"),
    ("Missing",   "=Missing_Count(UV_Data)"),
]

def _write_descriptive_stats(sheet: xw.Sheet) -> None:
    _section_heading(sheet, _ROW_SECTION_HDR, _C_C, "Descriptive Statistics")

    for i, (label, formula) in enumerate(_STAT_ROWS):
        row = _ROW_STATS_START + i
        _val(sheet, row, _C_C, label)
        _f(sheet, row, _C_D, formula)
        sheet.range(_rc(row, _C_D)).number_format = (
            "0" if label in ("Count", "Missing") else "0.0000"
        )

    # Border around the whole stats table
    last_row = _ROW_STATS_START + len(_STAT_ROWS) - 1
    _border_box(sheet, _ROW_SECTION_HDR, _C_C, last_row, _C_D)
    _subheader_row(sheet, _ROW_SECTION_HDR, _C_C, _C_D)


# ── Zone 3: histogram tables ──────────────────────────────────────────────────

def _write_histogram_table(
    sheet: xw.Sheet,
    col_edge: int,
    col_count: int,
    method: str,
    short_label: str,
) -> None:
    """Write one histogram bin table (edges + counts) for the given method."""
    # Section heading spans edge and count columns
    _section_heading(sheet, _ROW_SECTION_HDR, col_edge, short_label)
    sheet.range(_rc(_ROW_SECTION_HDR, col_edge), _rc(_ROW_SECTION_HDR, col_count)).merge()

    # Column sub-headers
    _val(sheet, _ROW_COL_HDRS, col_edge,  "Upper Edge")
    _val(sheet, _ROW_COL_HDRS, col_count, "Count")
    _subheader_row(sheet, _ROW_COL_HDRS, col_edge, col_count)

    # Spill formulas
    _f(sheet, _ROW_HIST_START, col_edge,
       f"=Bin_Edges(UV_Data,\"{method}\")")
    edge_spill_ref = f"{_col_letter(col_edge)}{_ROW_HIST_START}#"
    _f(sheet, _ROW_HIST_START, col_count,
       f"=Bin_Counts(UV_Data,{edge_spill_ref})")

    # Number format for edges
    sheet.range(_rc(_ROW_HIST_START, col_edge)).number_format = "0.00"
    sheet.range(_rc(_ROW_HIST_START, col_count)).number_format = "0"


def _write_histograms(sheet: xw.Sheet) -> None:
    # Three side-by-side bin tables; each method heading sits in the shared
    # section-header row alongside "Data" and "Descriptive Statistics".
    _write_histogram_table(sheet, _C_F, _C_G, "Sturges", "Sturges")
    _write_histogram_table(sheet, _C_I, _C_J, "Scott",   "Scott")
    _write_histogram_table(sheet, _C_L, _C_M, "FD",      "Freedman-Diaconis")


# ── Zone 4: distribution fitting summary table ────────────────────────────────
# Closed-form distributions only.  Each row: name, param labels/values, NLL, k, AIC, BIC.
# The formulas reference UV_Data directly; the LAMBDA functions handle numeric filtering.

# Columns headers and number formats
_FIT_COL_HDRS = [
    (_C_P, "Distribution"),
    (_C_Q, "θ₁"),
    (_C_R, "Value"),
    (_C_S, "θ₂"),
    (_C_T, "Value"),
    (_C_U, "θ₃"),
    (_C_V, "Value"),
    (_C_W, "NLL"),
    (_C_X, "k"),
    (_C_Y, "AIC"),
    (_C_Z, "BIC"),
]

_FIT_NUMBER_FORMATS: dict[int, str] = {
    _C_R: "0.0000",
    _C_T: "0.0000",
    _C_V: "0.0000",
    _C_W: "0.00",
    _C_X: "0",
    _C_Y: "0.00",
    _C_Z: "0.00",
}

# Distribution definitions.
# Each entry: (name, θ₁_label, θ₁_formula, θ₂_label, θ₂_formula,
#              θ₃_label, θ₃_formula, nll_formula, k)
# θ₂/θ₃ may be "" (blank).  nll_formula uses R, T, V column letter
# references at the distribution's own row.

def _dist_rows(base_row: int) -> list[tuple]:
    """Return distribution row specs.  base_row = row of first distribution."""

    def _r(row: int) -> str:
        return f"${_col_letter(_C_R)}${row}"

    def _t(row: int) -> str:
        return f"${_col_letter(_C_T)}${row}"

    def _v(row: int) -> str:
        return f"${_col_letter(_C_V)}${row}"

    def _n(row: int) -> str:
        return f"${_col_letter(_C_X)}${row}"

    rows: list[tuple] = []
    dist_specs = [
        # (name, θ₁_lbl, θ₁_formula, θ₂_lbl, θ₂_formula, θ₃_lbl, θ₃_formula, nll_builder, k)
        (
            "Normal",
            "Mean",    "=AVERAGE(UV_Data)",
            "Std Dev", "=STDEV.S(UV_Data)",
            "",        "",
            lambda r: f"=NLL_Normal(UV_Data,{_r(r)},{_t(r)})",
            2,
        ),
        (
            "Lognormal",
            "μ_ln",  "=AVERAGE(LN(FILTER(UV_Data,ISNUMBER(UV_Data))))",
            "σ_ln",  "=STDEV.S(LN(FILTER(UV_Data,ISNUMBER(UV_Data))))",
            "",      "",
            lambda r: f"=NLL_Lognormal(UV_Data,{_r(r)},{_t(r)})",
            2,
        ),
        (
            "Exponential",
            "Rate",  "=1/AVERAGE(UV_Data)",
            "",      "",
            "",      "",
            lambda r: f"=NLL_Exponential(UV_Data,{_r(r)})",
            1,
        ),
        (
            "Triangular",
            "Min",  "=MIN(UV_Data)",
            "Mode", "=IFERROR(MODE.SNGL(UV_Data),MEDIAN(UV_Data))",
            "Max",  "=MAX(UV_Data)",
            lambda r: f"=NLL_Triangular(UV_Data,{_r(r)},{_t(r)},{_v(r)})",
            3,
        ),
        (
            "BetaPERT",
            "Min",  "=MIN(UV_Data)",
            "Mode", "=IFERROR(MODE.SNGL(UV_Data),MEDIAN(UV_Data))",
            "Max",  "=MAX(UV_Data)",
            lambda r: f"=NLL_BetaPERT(UV_Data,{_r(r)},{_t(r)},{_v(r)})",
            3,
        ),
    ]

    for i, (name, l1, f1, l2, f2, l3, f3, nll_fn, k) in enumerate(dist_specs):
        row = base_row + i
        rows.append((row, name, l1, f1, l2, f2, l3, f3, nll_fn(row), k))

    return rows


def _write_fitting_table(sheet: xw.Sheet) -> None:
    _section_heading(sheet, _ROW_SECTION_HDR, _C_P, "Distribution Fitting — Closed-Form MLE")

    # Column headers
    for col, label in _FIT_COL_HDRS:
        _val(sheet, _ROW_COL_HDRS, col, label)
    _subheader_row(sheet, _ROW_COL_HDRS, _C_P, _C_Z)

    # Distribution rows
    dist_data = _dist_rows(_ROW_DIST_START)
    for row, name, l1, f1, l2, f2, l3, f3, nll_f, k in dist_data:
        _val(sheet, row, _C_P, name)
        _val(sheet, row, _C_Q, l1)
        if f1:
            _f(sheet, row, _C_R, f1)
        _val(sheet, row, _C_S, l2)
        if f2:
            _f(sheet, row, _C_T, f2)
        _val(sheet, row, _C_U, l3)
        if f3:
            _f(sheet, row, _C_V, f3)
        _f(sheet, row, _C_W, nll_f)
        _val(sheet, row, _C_X, k)
        _f(sheet, row, _C_Y, f"=GoF_AIC(${_col_letter(_C_W)}${row},${_col_letter(_C_X)}${row})")
        _f(sheet, row, _C_Z,
           f"=GoF_BIC(${_col_letter(_C_W)}${row},${_col_letter(_C_X)}${row},UV_n)")

        # Number formats
        for col, fmt in _FIT_NUMBER_FORMATS.items():
            sheet.range(_rc(row, col)).number_format = fmt

    # Border around the table
    last_row = _ROW_DIST_START + len(dist_data) - 1
    _border_box(sheet, _ROW_SECTION_HDR, _C_P, last_row, _C_Z)

    # Highlight the best-fit row (lowest AIC) with conditional formatting
    aic_range = sheet.range(
        _rc(_ROW_DIST_START, _C_Y),
        _rc(last_row, _C_Y),
    )
    row_range = sheet.range(
        _rc(_ROW_DIST_START, _C_P),
        _rc(last_row, _C_Z),
    )
    # Use a formula-based rule: highlight the entire row where AIC = MIN(AIC column)
    aic_col_letter = _col_letter(_C_Y)
    aic_min_range  = f"${aic_col_letter}${_ROW_DIST_START}:${aic_col_letter}${last_row}"
    formula = f"=${aic_col_letter}{_ROW_DIST_START}=MIN({aic_min_range})"
    cf = row_range.api.FormatConditions.Add(Type=2, Formula1=formula)  # xlExpression=2
    cf.Interior.Color = 0xC6EFCE  # light green fill
    cf.Font.Color     = 0x276228  # dark green text


# ── Histogram charts ──────────────────────────────────────────────────────────

def _add_histogram_chart(
    sheet: xw.Sheet,
    chart_left: float,
    chart_top: float,
    title: str,
    edges_name: str,
    counts_name: str,
) -> None:
    """Insert one column-chart histogram below the bin tables."""
    sname = sheet.name
    chart_obj = sheet.charts.add(
        left=chart_left,
        top=chart_top,
        width=_CHART_WIDTH,
        height=_CHART_HEIGHT,
    )
    chart = chart_obj.chart
    chart.chart_type = _XL_COLUMN_CLUSTERED
    chart.has_title = True
    chart.chart_title.text = title

    # Clear default series
    while chart.series_collection.count > 0:
        chart.series_collection(1).delete()

    series = chart.series_collection.new_series()
    series.name      = title
    # Reference sheet-scoped names for X (categories) and Y (values)
    series.formula   = (
        f"=SERIES(\"{title}\","
        f"'{sname}'!{edges_name},"
        f"'{sname}'!{counts_name},"
        f"1)"
    )

    # Gap width 0 → bars touch (histogram style)
    series.api.GapWidth = 0

    chart.axes(_XL_CATEGORY).has_title = True
    chart.axes(_XL_CATEGORY).axis_title.text = "Upper Edge"
    chart.axes(_XL_VALUE).has_title = True
    chart.axes(_XL_VALUE).axis_title.text = "Count"

    # Remove legend (single series)
    chart.has_legend = False


def _write_histogram_charts(sheet: xw.Sheet) -> None:
    """Add three histogram charts below the bin tables."""
    # Anchor charts below the histogram tables; top is at row ~35 equivalent
    # Use pixel-based positioning: each row ≈ 15 pt, chart starts after some rows
    top_offset = (_ROW_HIST_START + 35) * 15.0  # approximate top in points

    for i, (title, edges_name, counts_name, col) in enumerate([
        ("Sturges",           "UV_Sturges_Edges", "UV_Sturges_Counts", _C_F),
        ("Scott",             "UV_Scott_Edges",   "UV_Scott_Counts",   _C_I),
        ("Freedman-Diaconis", "UV_FD_Edges",      "UV_FD_Counts",      _C_L),
    ]):
        left = sum(
            sheet.range(_rc(1, c)).column_width * 7.5  # approx pts per char
            for c in range(1, col)
        )
        _add_histogram_chart(
            sheet,
            chart_left=left,
            chart_top=top_offset,
            title=title,
            edges_name=edges_name,
            counts_name=counts_name,
        )


# ── Row height and freeze ─────────────────────────────────────────────────────

def _finalize_sheet(sheet: xw.Sheet) -> None:
    sheet.range(_rc(1, 1)).api.EntireRow.RowHeight = 20
    sheet.range(_rc(_ROW_SECTION_HDR, 1)).api.EntireRow.RowHeight = 18
    sheet.range(_rc(_ROW_COL_HDRS, 1)).api.EntireRow.RowHeight = 18
    # Freeze panes below the column header row and to the right of col A
    sheet.range(_rc(_ROW_HIST_START, _C_C)).api.Select()
    sheet.book.app.api.ActiveWindow.FreezePanes = True


# ── Main entry point ──────────────────────────────────────────────────────────

def write_univariate_sheet(workbook: xw.Book) -> xw.Sheet:
    """Create or replace the Univariate sheet and write all content.

    Parameters
    ----------
    workbook : xw.Book
        Open xlwings workbook to write into.

    Returns
    -------
    xw.Sheet
        The written sheet.
    """
    # Delete existing sheet if present
    if UNIVARIATE_SHEET_NAME in {s.name for s in workbook.sheets}:
        workbook.sheets[UNIVARIATE_SHEET_NAME].delete()

    # Add sheet after the last existing sheet
    sheet = workbook.sheets.add(
        name=UNIVARIATE_SHEET_NAME,
        after=workbook.sheets[-1],
    )

    # Title row
    _val(sheet, _ROW_TITLE, _C_A, "Univariate Analysis")
    sheet.range(_rc(_ROW_TITLE, _C_A)).api.Font.Bold = True
    sheet.range(_rc(_ROW_TITLE, _C_A)).api.Font.Size = 14

    # Column widths
    _set_column_widths(sheet)

    # Sheet-scoped named ranges (must exist before formulas reference them)
    _setup_local_names(sheet)

    # Write zones
    _write_data_zone(sheet)
    _write_descriptive_stats(sheet)
    _write_histograms(sheet)
    _write_fitting_table(sheet)

    # Charts (skipped silently if chart API raises; e.g. headless builds)
    try:
        _write_histogram_charts(sheet)
    except Exception:
        pass

    _finalize_sheet(sheet)

    return sheet

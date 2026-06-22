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

Sheet layout
────────────
  Row 1          — Title "Univariate Analysis"
  Row 2          — Method/zone headings: Sturges Method | Scott Method | Freedman-Diaconis Method | Distribution Fitting/Comparison
  Row 3          — Section headings: Data | Filter | Descriptive Statistics | Bins; pane freeze anchored here
  Row 4          — Column sub-headers (table headers for histogram and fitting tables)
  Row 4+         — Data: raw column spill (col A), filter mask (col B), stats (D–E), histogram bins (G–N), fitting (Q–AA)

  Col A          — Data: LifeExpectancyData[Life expectancy] spill in A4 (unfiltered; spill range = $A$4#)
  Col B          — Filter: MAP(A4#,Data_Completeness) per-row mask — local UV_Include source ($B$4#)
  Col C          — thin gap (width 2); freeze pane left boundary
  Col D–E        — Descriptive Statistics (12 stat rows)
  Col F          — thin gap (width 2)
  Col G–H        — Sturges histogram table (edge | count)
  Col I          — thin gap (width 2)
  Col J–K        — Scott histogram table
  Col L          — thin gap (width 2)
  Col M–N        — Freedman-Diaconis histogram table
  Col O          — thin gap (width 2)
  Col P          — thin gap (width 2)
  Col Q–AA       — Distribution Fitting summary table

Sheet-scoped named ranges
─────────────────────────
  UV_Data        — spill range of the raw column formula ($A$4#, unfiltered)
  UV_Include     — local filter mask spill ($B$4#) — Data_Completeness applied per row
  UV_n           — IFERROR(COUNT(FILTER(UV_Data,UV_Include)), 0)
  UV_Sturges_Edges, UV_Sturges_Counts — OFFSET-based chart series ranges
  UV_Scott_Edges,   UV_Scott_Counts
  UV_FD_Edges,      UV_FD_Counts
"""
from __future__ import annotations


import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR
from .workbook_helpers import (
    a1, bold, bold_row, border_box, col_letter, drop_local_name,
    f, format_input, rc, section_heading, val,
)

# ── Column indices (1-based) ─────────────────────────────────────────────────

# Zone 1: Data Input + Filter
_C_A = 1    # data values
_C_B = 2    # local filter mask (MAP(A4#,Data_Completeness))

# Zone 2: Descriptive Statistics
_C_D = 4    # stat labels
_C_E = 5    # stat values

# Zone 3: Histogram Tables
_C_G = 7    # Sturges edges
_C_H = 8    # Sturges counts
_C_J = 10   # Scott edges
_C_K = 11   # Scott counts
_C_M = 13   # FD edges
_C_N = 14   # FD counts

# Zone 4: Distribution Fitting
_C_Q = 17   # distribution name
_C_R = 18   # θ₁ label
_C_S = 19   # θ₁ value
_C_T = 20   # θ₂ label
_C_U = 21   # θ₂ value
_C_V = 22   # θ₃ label
_C_W = 23   # θ₃ value
_C_X = 24   # NLL
_C_Y = 25   # k (param count)
_C_Z = 26   # AIC
_C_AA = 27   # BIC

# Zone 5: Weibull Grid-Search MLE (cols AC onward)
_N_GRID   = 20             # grid points per axis per stage (20×20 = 400/stage)
_C_GS     = 29             # col AC — first col of grid-search region (AB=28 is gap)
_GS_W     = _N_GRID + 1   # cols per stage = 21  (1 param2 col + N param1 cols)
_GS_GAP_C = 1              # gap col between Stage 1 and Stage 2
_C_GS_S2  = _C_GS + _GS_W + _GS_GAP_C   # col AY = 51

# Weibull block row anchor (aligns with other section headers; = _ROW_SECTION_HDR = 3)
_ROW_GS_WB   = 3

# Within-stage row offsets from block row_start
_GS_R_P1_BND = 1   # param1 (shape) bounds row
_GS_R_P2_BND = 2   # param2 (scale) bounds row
_GS_R_MINNLL = 3   # min-NLL display row
_GS_R_BEST   = 4   # best-param display row
_GS_R_AUX    = 5   # auxiliary row (argmin cells + Data Table input placeholders)
_GS_R_HDR    = 6   # grid header row (corner cell + param1 SEQUENCE)
_GS_R_BODY   = 7   # first grid body row

# Within-stage col offsets relative to stage col_start
# col_start+0  : param2 column (labels, param2 SEQUENCE, corner)
# col_start+1  : first param1 col (SEQUENCE + first body col)
# col_start+N  : last param1 col / body col
# Best-param display in _GS_R_BEST row:
_GS_C_BEST_P1 = 1   # best param1 value offset
_GS_C_LBL_P2  = 2   # "Best p2:" label offset
_GS_C_BEST_P2 = 3   # best param2 value offset
# Auxiliary cells in _GS_R_AUX row:
_GS_C_ROW_OFF = 0   # row-offset of grid minimum (1-indexed)
_GS_C_COL_OFF = 1   # col-offset of grid minimum (1-indexed)
_GS_C_DT_ROW  = 2   # Data Table row-input placeholder (param1 substitute)
_GS_C_DT_COL  = 3   # Data Table col-input placeholder (param2 substitute)

# ── Row anchors ───────────────────────────────────────────────────────────────
_ROW_TITLE        = 1   # "Univariate Analysis" + zone-level "Histograms" label
_ROW_METHOD_HDR   = 2   # method/zone headings: "Sturges Method", "Scott Method", etc.
_ROW_SECTION_HDR  = 3   # section headings: "Data", "Descriptive Statistics"
_ROW_COL_HDRS     = 4   # column sub-headers (histograms, fitting table)
_ROW_DATA_START   = 4   # data values begin (col A)
_ROW_STATS_START  = 4   # descriptive stat rows begin (col D–E)
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


def _subheader_row(sheet: xw.Sheet, row: int, c1: int, c2: int) -> None:
    sheet.range(rc(row, c1), rc(row, c2)).color = _SUBHDR
    sheet.range(rc(row, c1), rc(row, c2)).api.Font.Bold = True


# ── Column widths ─────────────────────────────────────────────────────────────

def _set_column_widths(sheet: xw.Sheet) -> None:
    widths = {
        _C_A: 14,    # data
        _C_B: 10,    # filter mask
        3:    2,     # gap (C) between filter and stats
        _C_D: 16,    # stat labels
        _C_E: 12,    # stat values
        6:    2,     # gap
        _C_G: 12,    # Sturges edge
        _C_H: 10,    # Sturges count
        9:    2,     # gap
        _C_J: 12,    # Scott edge
        _C_K: 10,    # Scott count
        12:   2,     # gap
        _C_M: 12,    # FD edge
        _C_N: 10,    # FD count
        15:   2,     # gap
        16:   2,     # gap
        _C_Q: 14,    # distribution name
        _C_R: 10,    # θ₁ label
        _C_S: 12,    # θ₁ value
        _C_T: 10,    # θ₂ label
        _C_U: 12,    # θ₂ value
        _C_V: 10,    # θ₃ label
        _C_W: 12,    # θ₃ value
        _C_X: 12,    # NLL
        _C_Y: 6,     # k
        _C_Z: 12,    # AIC
        _C_AA: 12,   # BIC
        28:   2,     # gap (AB) before grid-search section
    }
    for col, w in widths.items():
        sheet.range(rc(1, col), rc(1, col)).column_width = w

    # Grid-search columns: label/param2 cols are wider; body cols are narrow
    # Stage 1: AB(28)–AV(48); gap AW(49); Stage 2: AX(50)–BR(70)
    for stage_start in (_C_GS, _C_GS_S2):
        # param2 / label column
        sheet.range(rc(1, stage_start), rc(1, stage_start)).column_width = 12
        # param1 / body columns (N_GRID cols)
        for c in range(stage_start + 1, stage_start + _N_GRID + 1):
            sheet.range(rc(1, c), rc(1, c)).column_width = 6
    # gap col AX (col 50) between Stage 1 (AC–AW) and Stage 2 (AY onward)
    sheet.range(rc(1, _C_GS + _GS_W), rc(1, _C_GS + _GS_W)).column_width = 2


# ── Sheet-scoped named range management ──────────────────────────────────────

def _drop_wb_name(sheet: xw.Sheet, name: str) -> None:
    """Remove a workbook-scoped (non-local) name if present."""
    wb_names = sheet.book.api.Names
    for idx in range(wb_names.Count, 0, -1):
        n = wb_names(idx)
        if "!" not in n.Name and n.Name.lower() == name.lower():
            n.Delete()


def _setup_local_names(sheet: xw.Sheet) -> None:
    """Register sheet-scoped named ranges used by formulas and charts."""
    sname = sheet.name

    # UV_Data: spill range of the raw column formula in A4; unfiltered — filter is UV_Include
    drop_local_name(sheet, "UV_Data")
    sheet.api.Names.Add(
        Name="UV_Data",
        RefersTo=f"='{sname}'!${col_letter(_C_A)}${_ROW_DATA_START}#",
    )

    # UV_Include: local filter mask — spill of the Data_Completeness formula in B4
    drop_local_name(sheet, "UV_Include")
    sheet.api.Names.Add(
        Name="UV_Include",
        RefersTo=f"='{sname}'!${col_letter(_C_B)}${_ROW_DATA_START}#",
    )

    # UV_n: count of numeric included observations; used by GoF_BIC and chart range sizing
    drop_local_name(sheet, "UV_n")
    sheet.api.Names.Add(
        Name="UV_n",
        RefersTo=f"=IFERROR(COUNT(FILTER('{sname}'!${col_letter(_C_A)}${_ROW_DATA_START}#,UV_Include)),0)",
    )

    # Chart series ranges: OFFSET-based, sized by the method value stored in row 2.
    for name, col_ltr, start_row, size_formula in [
        ("UV_Sturges_Edges",  col_letter(_C_G), _ROW_HIST_START,
         f'n_histogram_bins(UV_Data,${col_letter(_C_H)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_Sturges_Counts", col_letter(_C_H), _ROW_HIST_START,
         f'n_histogram_bins(UV_Data,${col_letter(_C_H)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_Scott_Edges",    col_letter(_C_J), _ROW_HIST_START,
         f'n_histogram_bins(UV_Data,${col_letter(_C_K)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_Scott_Counts",   col_letter(_C_K), _ROW_HIST_START,
         f'n_histogram_bins(UV_Data,${col_letter(_C_K)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_FD_Edges",       col_letter(_C_M), _ROW_HIST_START,
         f'n_histogram_bins(UV_Data,${col_letter(_C_N)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_FD_Counts",      col_letter(_C_N), _ROW_HIST_START,
         f'n_histogram_bins(UV_Data,${col_letter(_C_N)}${_ROW_METHOD_HDR},UV_Include)'),
    ]:
        drop_local_name(sheet, name)
        sheet.api.Names.Add(
            Name=name,
            RefersTo=(
                f"=OFFSET('{sname}'!${col_ltr}${start_row},"
                f"0,0,{size_formula},1)"
            ),
        )


# ── Zone 1: data input column ─────────────────────────────────────────────────

def _write_data_zone(sheet: xw.Sheet) -> None:
    section_heading(sheet, _ROW_SECTION_HDR, _C_A, "Data")
    val(sheet, _ROW_COL_HDRS, _C_A, "Life expectancy")
    f(sheet, _ROW_DATA_START, _C_A, "=LifeExpectancyData[Life expectancy]")

    # Filter column — local Data_Completeness mask; UV_Include is defined as $B$4#
    section_heading(sheet, _ROW_SECTION_HDR, _C_B, "Filter")
    val(sheet, _ROW_COL_HDRS, _C_B, "Include")
    f(sheet, _ROW_DATA_START, _C_B,
       f"=MAP({col_letter(_C_A)}{_ROW_DATA_START}#,Data_Completeness)")


# ── Zone 2: descriptive statistics ───────────────────────────────────────────

# Stat label, formula — native Excel functions wrap UV_Data in FILTER(UV_Data,UV_Include);
# custom LAMBDAs receive UV_Include as their optional filter/Include argument.
_STAT_ROWS: list[tuple[str, str]] = [
    ("Mean",      "=AVERAGE(FILTER(UV_Data,UV_Include))"),
    ("Median",    "=MEDIAN(FILTER(UV_Data,UV_Include))"),
    ("Mode",      "=IFERROR(MODE.SNGL(FILTER(UV_Data,UV_Include)),\"N/A\")"),
    ("Std Dev",   "=STDEV.S(FILTER(UV_Data,UV_Include))"),
    ("Variance",  "=VAR.S(FILTER(UV_Data,UV_Include))"),
    ("Min",       "=MIN(FILTER(UV_Data,UV_Include))"),
    ("Max",       "=MAX(FILTER(UV_Data,UV_Include))"),
    ("Range",     "=LET(d,FILTER(UV_Data,UV_Include),MAX(d)-MIN(d))"),
    ("Skewness",  "=INDEX(Skewness(UV_Data,UV_Include),1)"),
    ("Kurtosis",  "=INDEX(Kurtosis(UV_Data,UV_Include),1)"),
    ("Count",     "=COUNT(FILTER(UV_Data,UV_Include))"),
    ("Missing",   "=Missing_Count(UV_Data,UV_Include)"),
]

def _write_descriptive_stats(sheet: xw.Sheet) -> None:
    section_heading(sheet, _ROW_SECTION_HDR, _C_D, "Descriptive Statistics")

    # Column sub-headers
    val(sheet, _ROW_COL_HDRS, _C_D, "Statistic")
    val(sheet, _ROW_COL_HDRS, _C_E, "Value")
    _subheader_row(sheet, _ROW_COL_HDRS, _C_D, _C_E)

    for i, (label, formula) in enumerate(_STAT_ROWS):
        row = _ROW_STATS_START + i
        val(sheet, row, _C_D, label)
        f(sheet, row, _C_E, formula)
        sheet.range(rc(row, _C_E)).number_format = (
            "0" if label in ("Count", "Missing") else "0.0000"
        )

    last_row = _ROW_STATS_START + len(_STAT_ROWS) - 1
    border_box(sheet, _ROW_SECTION_HDR, _C_D, last_row, _C_E)


# ── Zone 3: histogram tables ──────────────────────────────────────────────────

def _write_histogram_table(
    sheet: xw.Sheet,
    col_edge: int,
    col_count: int,
    method: str,
) -> None:
    """Write one histogram bin table (edges + counts) for the given method."""
    # Method heading row: label + value, both styled as section headings
    section_heading(sheet, _ROW_METHOD_HDR, col_edge, "Method:")
    val(sheet, _ROW_METHOD_HDR, col_count, method)
    sheet.range(rc(_ROW_METHOD_HDR, col_count)).color = _HEADER
    sheet.range(rc(_ROW_METHOD_HDR, col_count)).api.Font.Bold = True

    val(sheet, _ROW_COL_HDRS, col_edge,  "Upper Edge")
    val(sheet, _ROW_COL_HDRS, col_count, "Count")
    _subheader_row(sheet, _ROW_COL_HDRS, col_edge, col_count)

    # Row 3: bin-count display — "Bins" label + n_histogram_bins formula
    val(sheet, _ROW_SECTION_HDR, col_edge, "Bins:")
    method_cell = a1(_ROW_METHOD_HDR, col_count)
    f(sheet, _ROW_SECTION_HDR, col_count, f"=n_histogram_bins(UV_Data,{method_cell},UV_Include)")
    sheet.range(rc(_ROW_SECTION_HDR, col_count)).number_format = "0"

    # Spill formulas — method must be explicit (can't skip it to reach the 3rd filter arg)
    f(sheet, _ROW_HIST_START, col_edge, f"=Bin_Edges(UV_Data,{method_cell},UV_Include)")
    edge_spill_ref = f"{col_letter(col_edge)}{_ROW_HIST_START}#"
    f(sheet, _ROW_HIST_START, col_count,
       f"=Bin_Counts(UV_Data,{edge_spill_ref},UV_Include)")

    sheet.range(rc(_ROW_HIST_START, col_edge)).number_format = "0.00"
    sheet.range(rc(_ROW_HIST_START, col_count)).number_format = "0"


def _write_histograms(sheet: xw.Sheet) -> None:
    # Zone super-heading in title row, merged across all three histogram tables
    section_heading(sheet, _ROW_TITLE, _C_G, "Histograms")
    sheet.range(rc(_ROW_TITLE, _C_G), rc(_ROW_TITLE, _C_N)).merge()

    # Each table writes its own method heading at _ROW_METHOD_HDR (row 2)
    _write_histogram_table(sheet, _C_G, _C_H, "Sturges")
    _write_histogram_table(sheet, _C_J, _C_K, "Scott")
    _write_histogram_table(sheet, _C_M, _C_N, "FD")


# ── Zone 4: distribution fitting summary table ────────────────────────────────

_FIT_COL_HDRS = [
    (_C_Q, "Distribution"),
    (_C_R, "θ₁"),
    (_C_S, "Value"),
    (_C_T, "θ₂"),
    (_C_U, "Value"),
    (_C_V, "θ₃"),
    (_C_W, "Value"),
    (_C_X, "NLL"),
    (_C_Y, "k"),
    (_C_Z, "AIC"),
    (_C_AA, "BIC"),
]

_FIT_NUMBER_FORMATS: dict[int, str] = {
    _C_S: "0.0000",
    _C_U: "0.0000",
    _C_W: "0.0000",
    _C_X: "0.00",
    _C_Y: "0",
    _C_Z: "0.00",
    _C_AA: "0.00",
}


def _dist_rows(base_row: int) -> list[tuple]:
    """Return distribution row specs.  base_row = row of first distribution."""

    def _r(row: int) -> str:
        return f"${col_letter(_C_S)}${row}"

    def _t(row: int) -> str:
        return f"${col_letter(_C_U)}${row}"

    def _v(row: int) -> str:
        return f"${col_letter(_C_W)}${row}"

    def _n(row: int) -> str:
        return f"${col_letter(_C_Y)}${row}"

    rows: list[tuple] = []
    dist_specs = [
        (
            "Normal",
            "Mean",    "=AVERAGE(FILTER(UV_Data,UV_Include))",
            "Std Dev", "=STDEV.S(FILTER(UV_Data,UV_Include))",
            "",        "",
            lambda r: f"=NLL_Normal(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
        ),
        (
            "Lognormal",
            "μ_ln",  "=AVERAGE(LN(FILTER(UV_Data,(UV_Include)*(ISNUMBER(UV_Data)))))",
            "σ_ln",  "=STDEV.S(LN(FILTER(UV_Data,(UV_Include)*(ISNUMBER(UV_Data)))))",
            "",      "",
            lambda r: f"=NLL_Lognormal(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
        ),
        (
            "Exponential",
            "Rate",  "=1/AVERAGE(FILTER(UV_Data,UV_Include))",
            "",      "",
            "",      "",
            lambda r: f"=NLL_Exponential(UV_Data,{_r(r)},UV_Include)",
            1,
        ),
        (
            # NLL_Triangular and NLL_BetaPERT both have PDF = 0 at the boundary
            # points (x = min or x = max), which means any dataset where data
            # includes its own minimum or maximum (i.e. always) would produce
            # LN(0) → error → 1E+15.  Expanding the support by 0.1 % of the
            # data range gives boundary data points positive (tiny) density.
            "Triangular",
            "Min",  "=LET(d,FILTER(UV_Data,UV_Include),MIN(d)-(MAX(d)-MIN(d))*0.001)",
            "Mode", "=IFERROR(MODE.SNGL(FILTER(UV_Data,UV_Include)),MEDIAN(FILTER(UV_Data,UV_Include)))",
            "Max",  "=LET(d,FILTER(UV_Data,UV_Include),MAX(d)+(MAX(d)-MIN(d))*0.001)",
            lambda r: f"=NLL_Triangular(UV_Data,{_r(r)},{_t(r)},{_v(r)},UV_Include)",
            3,
        ),
        (
            "BetaPERT",
            "Min",  "=LET(d,FILTER(UV_Data,UV_Include),MIN(d)-(MAX(d)-MIN(d))*0.001)",
            "Mode", "=IFERROR(MODE.SNGL(FILTER(UV_Data,UV_Include)),MEDIAN(FILTER(UV_Data,UV_Include)))",
            "Max",  "=LET(d,FILTER(UV_Data,UV_Include),MAX(d)+(MAX(d)-MIN(d))*0.001)",
            lambda r: f"=NLL_BetaPERT(UV_Data,{_r(r)},{_t(r)},{_v(r)},UV_Include)",
            3,
        ),
        # Weibull: params come from Stage 2 grid-search MLE (see Zone 5)
        (
            "Weibull",
            "shape",
            f"=${col_letter(_C_GS_S2 + _GS_C_BEST_P1)}${_ROW_GS_WB + _GS_R_BEST}",
            "scale",
            f"=${col_letter(_C_GS_S2 + _GS_C_BEST_P2)}${_ROW_GS_WB + _GS_R_BEST}",
            "",    "",
            lambda r: f"=NLL_Weibull(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
        ),
    ]

    for i, (name, l1, f1, l2, f2, l3, f3, nll_fn, k) in enumerate(dist_specs):
        row = base_row + i
        rows.append((row, name, l1, f1, l2, f2, l3, f3, nll_fn(row), k))

    return rows


def _write_fitting_table(sheet: xw.Sheet) -> None:
    # Zone heading at row 2, merged across all fitting columns
    section_heading(sheet, _ROW_METHOD_HDR, _C_Q, "Distribution Fitting/Comparison")
    sheet.range(rc(_ROW_METHOD_HDR, _C_Q), rc(_ROW_METHOD_HDR, _C_AA)).merge()

    for col, label in _FIT_COL_HDRS:
        val(sheet, _ROW_COL_HDRS, col, label)
    _subheader_row(sheet, _ROW_COL_HDRS, _C_Q, _C_AA)

    dist_data = _dist_rows(_ROW_DIST_START)
    for row, name, l1, f1, l2, f2, l3, f3, nll_f, k in dist_data:
        val(sheet, row, _C_Q, name)
        val(sheet, row, _C_R, l1)
        if f1:
            f(sheet, row, _C_S, f1)
        val(sheet, row, _C_T, l2)
        if f2:
            f(sheet, row, _C_U, f2)
        val(sheet, row, _C_V, l3)
        if f3:
            f(sheet, row, _C_W, f3)
        f(sheet, row, _C_X, nll_f)
        val(sheet, row, _C_Y, k)
        f(sheet, row, _C_Z, f"=GoF_AIC(${col_letter(_C_X)}${row},${col_letter(_C_Y)}${row})")
        f(sheet, row, _C_AA,
           f"=GoF_BIC(${col_letter(_C_X)}${row},${col_letter(_C_Y)}${row},UV_n)")

        for col, fmt in _FIT_NUMBER_FORMATS.items():
            sheet.range(rc(row, col)).number_format = fmt

    # Border around the table (col headers through last data row)
    last_row = _ROW_DIST_START + len(dist_data) - 1
    border_box(sheet, _ROW_COL_HDRS, _C_Q, last_row, _C_AA)

    # Highlight best-fit row (lowest AIC) with conditional formatting
    aic_col_letter = col_letter(_C_Z)
    aic_min_range  = f"${aic_col_letter}${_ROW_DIST_START}:${aic_col_letter}${last_row}"
    formula = f"=${'$'}{aic_col_letter}{_ROW_DIST_START}=MIN({aic_min_range})"
    row_range = sheet.range(rc(_ROW_DIST_START, _C_Q), rc(last_row, _C_AA))
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

    while chart.series_collection.count > 0:
        chart.series_collection(1).delete()

    series = chart.series_collection.new_series()
    series.name      = title
    series.formula   = (
        f"=SERIES(\"{title}\","
        f"'{sname}'!{edges_name},"
        f"'{sname}'!{counts_name},"
        f"1)"
    )

    series.api.GapWidth = 0

    chart.axes(_XL_CATEGORY).has_title = True
    chart.axes(_XL_CATEGORY).axis_title.text = "Upper Edge"
    chart.axes(_XL_VALUE).has_title = True
    chart.axes(_XL_VALUE).axis_title.text = "Count"

    chart.has_legend = False


def _write_histogram_charts(sheet: xw.Sheet) -> None:
    """Add three histogram charts below the bin tables."""
    top_offset = (_ROW_HIST_START + 35) * 15.0

    for i, (title, edges_name, counts_name, col) in enumerate([
        ("Sturges",           "UV_Sturges_Edges", "UV_Sturges_Counts", _C_G),
        ("Scott",             "UV_Scott_Edges",   "UV_Scott_Counts",   _C_J),
        ("Freedman-Diaconis", "UV_FD_Edges",      "UV_FD_Counts",      _C_M),
    ]):
        left = sum(
            sheet.range(rc(1, c)).column_width * 7.5
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


# ── Zone 5: Weibull grid-search MLE ──────────────────────────────────────────
#
# Layout within one stage block (row_start, col_start):
#   row+0  : section heading (merged across all stage cols)
#   row+1  : shape (k) bounds  — min at col+1, max at col+2
#   row+2  : scale (λ) bounds  — min at col+1, max at col+2
#   row+3  : Min NLL display   — value at col+1
#   row+4  : Best shape        — col+1; Best scale — col+2 (label) col+3 (value)
#   row+5  : row_offset (col+0), col_offset (col+1), dt_row_input (col+2), dt_col_input (col+3)
#   row+6  : corner cell (col+0), param1 SEQUENCE → spills right N cols
#   row+7…+6+N : param2 SEQUENCE (col+0), Data Table body (cols+1…+N)
#
# Stage 1: col_start = _C_GS  = 28 (AB)
# Stage 2: col_start = _C_GS_S2 = 50 (AX)  — bounds computed from Stage 1 best ± 1 step
#
# Named ranges registered here:
#   UV_WB_S1 = Stage 1 body  (AC10:AV29 for default layout)
#   UV_WB_S2 = Stage 2 body  (AY10:BR29)


def _gs_a1(row_start: int, col_start: int, dr: int, dc: int) -> str:
    """Absolute A1 ref offset from (row_start, col_start) by (dr, dc)."""
    return f"${col_letter(col_start + dc)}${row_start + dr}"


def _write_grid_stage(
    sheet: xw.Sheet,
    row_start: int,
    col_start: int,
    title: str,
    body_name: str,
    p1_min,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    p1_max,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    p2_min,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    p2_max,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    editable_bounds: bool = True,
) -> dict:
    """Write one 20×20 NLL grid-search stage for Weibull(shape, scale).

    Returns a dict of absolute A1 refs for key cells, used by Stage 2 formulas
    and the fitting summary:
        best_p1, best_p2, min_p1, max_p1, min_p2, max_p2,
        row_offset, col_offset, corner, p1_seq, p2_seq
    """
    sname = sheet.name
    n = _N_GRID

    # ── Control cells ────────────────────────────────────────────────────────
    r0, c0 = row_start, col_start
    last_col = c0 + n  # inclusive last col (param1 col N)

    # Title row — merged across full stage width
    val(sheet, r0, c0, title)
    sheet.range(rc(r0, c0), rc(r0, last_col)).merge()
    sheet.range(rc(r0, c0)).api.Font.Bold = True
    sheet.range(rc(r0, c0)).color = _HEADER

    # Param bounds rows
    val(sheet, r0 + _GS_R_P1_BND, c0, "shape (k) range:")
    val(sheet, r0 + _GS_R_P2_BND, c0, "scale (λ) range:")
    for dr, p_min, p_max in [
        (_GS_R_P1_BND, p1_min, p1_max),
        (_GS_R_P2_BND, p2_min, p2_max),
    ]:
        row = r0 + dr
        if isinstance(p_min, str):
            f(sheet, row, c0 + 1, p_min)
        else:
            val(sheet, row, c0 + 1, p_min)
        if isinstance(p_max, str):
            f(sheet, row, c0 + 2, p_max)
        else:
            val(sheet, row, c0 + 2, p_max)
        if editable_bounds:
            sheet.range(rc(row, c0 + 1)).color = _INPUT
            sheet.range(rc(row, c0 + 2)).color = _INPUT
        sheet.range(rc(row, c0 + 1)).number_format = "0.0000"
        sheet.range(rc(row, c0 + 2)).number_format = "0.0000"

    # Min NLL row — COUNT guard needed because MIN(empty range) = 0, not an error
    val(sheet, r0 + _GS_R_MINNLL, c0, "Min NLL:")
    f(sheet, r0 + _GS_R_MINNLL, c0 + 1,
       f'=IF(COUNT({body_name})=0,"—",MIN({body_name}))')
    sheet.range(rc(r0 + _GS_R_MINNLL, c0 + 1)).number_format = "0.00"

    # Auxiliary row: row_offset, col_offset, dt_row_input, dt_col_input
    aux_row   = r0 + _GS_R_AUX
    r_off_ref = _gs_a1(r0, c0, _GS_R_AUX, _GS_C_ROW_OFF)   # row_offset cell
    c_off_ref = _gs_a1(r0, c0, _GS_R_AUX, _GS_C_COL_OFF)   # col_offset cell
    corner_ref = _gs_a1(r0, c0, _GS_R_HDR, 0)               # corner cell
    p1_seq_ref = _gs_a1(r0, c0, _GS_R_HDR, 1)               # param1 SEQUENCE anchor
    p2_seq_ref = _gs_a1(r0, c0, _GS_R_BODY, 0)              # param2 SEQUENCE anchor

    # row_offset: 1-indexed row position of minimum within body
    f(sheet, aux_row, c0 + _GS_C_ROW_OFF,
       f"=IFERROR(MIN(IF({body_name}=MIN({body_name}),"
       f"ROW({body_name})-ROW({corner_ref}))),1)")
    # col_offset: 1-indexed column position of minimum within body
    f(sheet, aux_row, c0 + _GS_C_COL_OFF,
       f"=IFERROR(MIN(IF({body_name}=MIN({body_name}),"
       f"COLUMN({body_name})-COLUMN({corner_ref}))),1)")
    # Data Table input placeholder cells (written as values; Excel substitutes them)
    sheet.range(rc(aux_row, c0 + _GS_C_DT_ROW)).value = 1.0   # param1 (shape)
    sheet.range(rc(aux_row, c0 + _GS_C_DT_COL)).value = 1.0   # param2 (scale)
    dt_row_ref = _gs_a1(r0, c0, _GS_R_AUX, _GS_C_DT_ROW)
    dt_col_ref = _gs_a1(r0, c0, _GS_R_AUX, _GS_C_DT_COL)

    # Best-param display row (references the SEQUENCE spill anchors)
    best_row = r0 + _GS_R_BEST
    val(sheet, best_row, c0, "Best shape:")
    f(sheet, best_row, c0 + _GS_C_BEST_P1,
       f"=IFERROR(INDEX({p1_seq_ref}#,1,{c_off_ref}),\"—\")")
    val(sheet, best_row, c0 + _GS_C_LBL_P2, "Best scale:")
    f(sheet, best_row, c0 + _GS_C_BEST_P2,
       f"=IFERROR(INDEX({p2_seq_ref}#,{r_off_ref}),\"—\")")
    for dc in (_GS_C_BEST_P1, _GS_C_BEST_P2):
        sheet.range(rc(best_row, c0 + dc)).number_format = "0.0000"

    # Boundary guard: red fill on best-param cells when minimum is on grid edge
    for dc, off_ref in [(_GS_C_BEST_P1, c_off_ref), (_GS_C_BEST_P2, r_off_ref)]:
        cell_api = sheet.range(rc(best_row, c0 + dc)).api
        cf = cell_api.FormatConditions.Add(
            Type=2,  # xlExpression
            Formula1=f"=OR({off_ref}=1,{off_ref}={n})",
        )
        cf.Interior.Color = 0x0000FF   # red (BGR)
        cf.Font.Color     = 0xFFFFFF   # white

    # ── Grid: header row (corner + param1 SEQUENCE) ──────────────────────────
    hdr_row = r0 + _GS_R_HDR
    min_p1_ref = _gs_a1(r0, c0, _GS_R_P1_BND, 1)
    max_p1_ref = _gs_a1(r0, c0, _GS_R_P1_BND, 2)
    min_p2_ref = _gs_a1(r0, c0, _GS_R_P2_BND, 1)
    max_p2_ref = _gs_a1(r0, c0, _GS_R_P2_BND, 2)

    # Corner cell: NLL formula referencing the two Data Table input placeholders
    f(sheet, hdr_row, c0,
       f"=NLL_Weibull(UV_Data,{dt_row_ref},{dt_col_ref},UV_Include)")
    sheet.range(rc(hdr_row, c0)).number_format = "0.00"

    # Param1 (shape) SEQUENCE — spills right across N cols
    f(sheet, hdr_row, c0 + 1,
       f"=SEQUENCE(1,{n},{min_p1_ref},({max_p1_ref}-{min_p1_ref})/({n}-1))")
    sheet.range(rc(hdr_row, c0 + 1)).number_format = "0.00"

    # ── Grid: body rows (param2 SEQUENCE + Data Table body) ──────────────────
    body_row_start = r0 + _GS_R_BODY
    body_row_end   = body_row_start + n - 1
    body_col_start = c0 + 1
    body_col_end   = c0 + n

    # Param2 (scale) SEQUENCE — spills down N rows in col c0
    f(sheet, body_row_start, c0,
       f"=SEQUENCE({n},1,{min_p2_ref},({max_p2_ref}-{min_p2_ref})/({n}-1))")
    sheet.range(rc(body_row_start, c0)).number_format = "0.0000"

    # Register named range for the body BEFORE setting up the Data Table
    body_range = sheet.range(
        rc(body_row_start, body_col_start),
        rc(body_row_end,   body_col_end),
    )
    drop_local_name(sheet, body_name)
    sheet.api.Names.Add(
        Name=body_name,
        RefersTo=(
            f"='{sname}'!${col_letter(body_col_start)}${body_row_start}"
            f":${col_letter(body_col_end)}${body_row_end}"
        ),
    )

    # Two-input Data Table wiring is disabled — wire manually via the Excel UI.
    # row_input_cell = sheet.range(rc(aux_row, c0 + _GS_C_DT_ROW))
    # col_input_cell = sheet.range(rc(aux_row, c0 + _GS_C_DT_COL))
    # full_table_range = sheet.range(
    #     rc(hdr_row, c0),
    #     rc(body_row_end, body_col_end),
    # )
    # full_table_range.api.Table(
    #     RowInput=row_input_cell.api,
    #     ColumnInput=col_input_cell.api,
    # )

    # Number format for the body
    body_range.number_format = "0.00"

    # Heatmap: 3-color scale (green=low NLL → yellow=mid → red=high / overflow)
    try:
        body_range.api.FormatConditions.Delete()
        cs = body_range.api.FormatConditions.AddColorScale(3)
        cs.ColorScaleCriteria(1).Type = 1             # xlConditionValueLowestValue
        cs.ColorScaleCriteria(1).FormatColor.Color = 0x63BE7B   # green
        cs.ColorScaleCriteria(2).Type = 5             # xlConditionValuePercentile
        cs.ColorScaleCriteria(2).Value = 50
        cs.ColorScaleCriteria(2).FormatColor.Color = 0xFFEB84   # yellow
        cs.ColorScaleCriteria(3).Type = 2             # xlConditionValueHighestValue
        cs.ColorScaleCriteria(3).FormatColor.Color = 0xF8696B   # red
    except Exception:
        pass

    # Border around the whole stage block
    border_box(sheet, r0, c0, body_row_end, last_col)

    return {
        "best_p1":    _gs_a1(r0, c0, _GS_R_BEST, _GS_C_BEST_P1),
        "best_p2":    _gs_a1(r0, c0, _GS_R_BEST, _GS_C_BEST_P2),
        "min_p1":     _gs_a1(r0, c0, _GS_R_P1_BND, 1),
        "max_p1":     _gs_a1(r0, c0, _GS_R_P1_BND, 2),
        "min_p2":     _gs_a1(r0, c0, _GS_R_P2_BND, 1),
        "max_p2":     _gs_a1(r0, c0, _GS_R_P2_BND, 2),
        "row_offset": _gs_a1(r0, c0, _GS_R_AUX, _GS_C_ROW_OFF),
        "col_offset": _gs_a1(r0, c0, _GS_R_AUX, _GS_C_COL_OFF),
        "corner":     _gs_a1(r0, c0, _GS_R_HDR, 0),
        "p1_seq":     _gs_a1(r0, c0, _GS_R_HDR, 1),
        "p2_seq":     _gs_a1(r0, c0, _GS_R_BODY, 0),
    }


def _write_weibull_grid_search(sheet: xw.Sheet) -> None:
    """Write Stage 1 and Stage 2 Weibull grid-search MLE blocks side by side."""
    n = _N_GRID
    rs = _ROW_GS_WB   # row_start = 3

    # Stage 1 — user-editable bounds, initial wide search
    s1 = _write_grid_stage(
        sheet,
        row_start   = rs,
        col_start   = _C_GS,
        title       = "Weibull Grid-Search MLE  —  Stage 1  (shape × scale)",
        body_name   = "UV_WB_S1",
        p1_min      = 0.5,     # shape min (user can adjust)
        p1_max      = 10.0,    # shape max
        p2_min      = 0.1,     # scale min
        p2_max      = "=IFERROR(2*AVERAGE(FILTER(UV_Data,UV_Include)),10)",  # data-driven scale max
        editable_bounds = True,
    )

    # Stage 2 — bounds auto-computed as Stage 1 best ± one Stage-1 grid step
    # Step = (max - min) / (N - 1) per axis; zoom centres on Stage 1 best param.
    step_p1 = f"({s1['max_p1']}-{s1['min_p1']})/({n}-1)"
    step_p2 = f"({s1['max_p2']}-{s1['min_p2']})/({n}-1)"

    _write_grid_stage(
        sheet,
        row_start   = rs,
        col_start   = _C_GS_S2,
        title       = "Weibull Grid-Search MLE  —  Stage 2  (refined)",
        body_name   = "UV_WB_S2",
        p1_min      = f"=MAX(0.001,{s1['best_p1']}-{step_p1})",
        p1_max      = f"={s1['best_p1']}+{step_p1}",
        p2_min      = f"=MAX(0.001,{s1['best_p2']}-{step_p2})",
        p2_max      = f"={s1['best_p2']}+{step_p2}",
        editable_bounds = False,
    )

    # Section heading label over the gap column between stages
    val(sheet, rs, _C_GS + _GS_W, "")


# ── Row height and freeze ─────────────────────────────────────────────────────

def _finalize_sheet(sheet: xw.Sheet) -> None:
    sheet.range(rc(1, 1)).api.EntireRow.RowHeight = 20
    sheet.range(rc(_ROW_METHOD_HDR, 1)).api.EntireRow.RowHeight = 18
    sheet.range(rc(_ROW_SECTION_HDR, 1)).api.EntireRow.RowHeight = 18
    sheet.range(rc(_ROW_COL_HDRS, 1)).api.EntireRow.RowHeight = 18
    # Freeze at C4: rows 1-3 and cols A-B (data + filter) always visible
    sheet.range(rc(_ROW_DATA_START, 3)).api.Select()
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
    if UNIVARIATE_SHEET_NAME in {s.name for s in workbook.sheets}:
        workbook.sheets[UNIVARIATE_SHEET_NAME].delete()

    sheet = workbook.sheets.add(
        name=UNIVARIATE_SHEET_NAME,
        after=workbook.sheets[-1],
    )

    val(sheet, _ROW_TITLE, _C_A, "Univariate Analysis")
    sheet.range(rc(_ROW_TITLE, _C_A)).api.Font.Bold = True
    sheet.range(rc(_ROW_TITLE, _C_A)).api.Font.Size = 14

    _set_column_widths(sheet)

    # Data zone must be written first so the FILTER spill in A4 is in place
    # before _setup_local_names registers UV_Data ($A$4#) and UV_n.
    _write_data_zone(sheet)
    _setup_local_names(sheet)

    _write_descriptive_stats(sheet)
    _write_histograms(sheet)
    _write_fitting_table(sheet)

    # Grid-search MLE for two-parameter distributions (skipped silently on headless builds)
    try:
        _write_weibull_grid_search(sheet)
    except Exception:
        pass

    # Charts (skipped silently if chart API raises; e.g. headless builds)
    try:
        _write_histogram_charts(sheet)
    except Exception:
        pass

    _finalize_sheet(sheet)

    return sheet

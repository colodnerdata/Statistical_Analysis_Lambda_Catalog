"""
write_sheet_univariate.py
Writes the Univariate Analysis sheet into any target workbook.

Design decision — fitting approach
───────────────────────────────────
MLE throughout.  For Normal, Lognormal, and Exponential the MLE estimators are
closed-form (sample mean/SD on the raw or log-transformed data).  For
Triangular and BetaPERT the likelihood is non-differentiable at the mode, so
direct min/mode/max estimation is used — the result is a valid parameter set
that the NLL formula evaluates against.  Weibull, Gamma, and Beta use two-stage,
two-input Data Table grid searches.

Sheet layout
────────────
  Row 1          — Title "Univariate Analysis"
  Row 2          — Method labels/values: Sturges | Scott | FD | Distribution Fitting/Comparison
  Row 3          — Section headings: Data | Filter | Descriptive Statistics | Bins; pane freeze anchored here
  Row 4          — Column sub-headers (table headers for histogram and fitting tables)
  Row 4+         — Data: raw column spill (col A), filter mask (col B), stats (D–E), histogram bins (G–N), fitting (Q–AA)

  Col A          — Data: LifeExpectancyData[Life expectancy] spill in A4 (unfiltered; spill range = $A$4#)
  Col B          — Filter: Data_Completeness over the source table column — local UV_Include source ($B$4#)
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
  Col AB         — thin gap before the grid-search section
  Col AC–AW      — Stage 1 controls and 20×20 Data Table
  Col AX         — thin gap between grid-search stages
  Col AY–BS      — Stage 2 controls and 20×20 Data Table

  Grid-search rows 1–5 — stage titles, compact controls, and Data Table headings
  Grid-search rows 6–25 — 20×20 Data Table bodies, repeated per distribution

Sheet-scoped named ranges
─────────────────────────
  UV_Data        — spill range of the raw column formula ($A$4#, unfiltered)
  UV_Include     — local filter mask spill ($B$4#) — Data_Completeness applied per row
  UV_n           — IFERROR(COUNT(FILTER(UV_Data,UV_Include)), 0)
  UV_Sturges_Edges, UV_Sturges_Counts — OFFSET-based chart series ranges
  UV_Scott_Edges,   UV_Scott_Counts
  UV_FD_Edges,      UV_FD_Counts
  UV_WB_S1       — Stage 1 Weibull Data Table body (AD6:AW25)
  UV_WB_S2       — Stage 2 Weibull Data Table body (AZ6:BS25)
  UV_GAMMA_S1/S2 — Stage 1/2 Gamma Data Table bodies
  UV_BETA_S1/S2  — Stage 1/2 Beta Data Table bodies
"""
from __future__ import annotations


import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR
from .workbook_helpers import (
    a1, border_box, col_letter, drop_local_name,
    f, rc, section_heading, val,
)

# ── Column indices (1-based) ─────────────────────────────────────────────────

# Zone 1: Data Input + Filter
_C_A = 1    # data values
_C_B = 2    # local filter mask over the source table column

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

# Zone 5: two-parameter Grid-Search MLE (cols AC onward)
_N_GRID   = 20             # grid points per axis per stage (20×20 = 400/stage)
_C_GS     = 29             # col AC — first col of grid-search region (AB=28 is gap)
_GS_W     = _N_GRID + 1   # cols per stage = 21  (1 param2 col + N param1 cols)
_GS_GAP_C = 1              # gap col between Stage 1 and Stage 2
_C_GS_S2  = _C_GS + _GS_W + _GS_GAP_C   # col AY = 51

# Within-stage row offsets from block row_start
_GS_R_CONTROL_HDR = 1   # Min NLL label + parameter-table headers
_GS_R_P1          = 2   # column parameter row
_GS_R_P2          = 3   # row parameter row
_GS_R_HDR         = 4   # Data Table corner + column-parameter SEQUENCE
_GS_R_BODY        = 5   # first Data Table body row

# Grid-search block row anchors (stage title rows)
_GS_BLOCK_H      = _GS_R_BODY + _N_GRID
_GS_BLOCK_GAP_R  = 1
_ROW_GS_WB       = 1
_ROW_GS_GAMMA    = _ROW_GS_WB + _GS_BLOCK_H + _GS_BLOCK_GAP_R
_ROW_GS_BETA     = _ROW_GS_GAMMA + _GS_BLOCK_H + _GS_BLOCK_GAP_R

# Fixed-area column offsets relative to stage col_start
_GS_C_MINNLL = 0   # two-cell vertical Min NLL table
_GS_C_N_GRID = 1   # two-cell vertical Rows/Columns table
_GS_C_SPACER = 2   # intentionally blank between fixed-area tables
_GS_C_PARAM  = 3   # Parameter
_GS_C_INPUT  = 4   # visible Data Table substitution cells
_GS_C_MIN    = 5   # parameter range minimum
_GS_C_MAX    = 6   # parameter range maximum
_GS_C_STEP   = 7   # endpoint-inclusive parameter step size
_GS_C_BEST   = 8   # Grid_Search_Optimum spill anchor / result column

# ── Row anchors ───────────────────────────────────────────────────────────────
_ROW_TITLE        = 1   # "Univariate Analysis" + zone-level "Histograms" label
_ROW_METHOD_HDR   = 2   # method labels and values for Sturges, Scott, and FD
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

UNIVARIATE_SHEET_NAME = "Univariate"

_FMT_INT = "0"
_FMT_1DP = "0.0"
_FMT_SCI_1DP = "0.0E+00"


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

    # Grid-search stages: compact fixed-area controls above narrow Data Tables.
    # Stage 1: AC–AW; gap AX; Stage 2: AY–BS.
    for stage_start in (_C_GS, _C_GS_S2):
        for c in range(stage_start, stage_start + _N_GRID + 1):
            sheet.range(rc(1, c), rc(1, c)).column_width = 6
        sheet.range(rc(1, stage_start)).column_width = 12                  # Min NLL / row axis
        sheet.range(rc(1, stage_start + _GS_C_N_GRID)).column_width = 14   # Rows/Columns
        sheet.range(rc(1, stage_start + _GS_C_PARAM)).column_width = 13   # Parameter
        for dc in (_GS_C_INPUT, _GS_C_MIN, _GS_C_MAX, _GS_C_STEP, _GS_C_BEST):
            sheet.range(rc(1, stage_start + dc)).column_width = 10

    # Gap col AX between Stage 1 and Stage 2.
    sheet.range(rc(1, _C_GS + _GS_W), rc(1, _C_GS + _GS_W)).column_width = 2


def _autofit_column_widths(sheet: xw.Sheet) -> None:
    """Autofit all populated layout columns, then restore intentional gaps."""
    last_col = _C_GS_S2 + _GS_W - 1
    sheet.range(rc(_ROW_TITLE, _C_A), rc(_DATA_END, last_col)).columns.autofit()

    for col in (3, 6, 9, 12, 15, 16, 28, _C_GS + _GS_W):
        sheet.range(rc(1, col), rc(1, col)).column_width = 2


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

    # Remove obsolete workbook-scoped definitions before recreating local names.
    for name in ("UV_Data", "UV_Include", "UV_n"):
        _drop_wb_name(sheet, name)

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
         f'num_histogram_bins(UV_Data,${col_letter(_C_H)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_Sturges_Counts", col_letter(_C_H), _ROW_HIST_START,
         f'num_histogram_bins(UV_Data,${col_letter(_C_H)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_Scott_Edges",    col_letter(_C_J), _ROW_HIST_START,
         f'num_histogram_bins(UV_Data,${col_letter(_C_K)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_Scott_Counts",   col_letter(_C_K), _ROW_HIST_START,
         f'num_histogram_bins(UV_Data,${col_letter(_C_K)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_FD_Edges",       col_letter(_C_M), _ROW_HIST_START,
         f'num_histogram_bins(UV_Data,${col_letter(_C_N)}${_ROW_METHOD_HDR},UV_Include)'),
        ("UV_FD_Counts",      col_letter(_C_N), _ROW_HIST_START,
         f'num_histogram_bins(UV_Data,${col_letter(_C_N)}${_ROW_METHOD_HDR},UV_Include)'),
    ]:
        _drop_wb_name(sheet, name)
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
    f(
        sheet,
        _ROW_DATA_START,
        _C_A,
        '=IF(LifeExpectancyData[Life expectancy]="","",LifeExpectancyData[Life expectancy])',
    )
    sheet.range(rc(_ROW_DATA_START, _C_A), rc(_DATA_END, _C_A)).number_format = _FMT_1DP

    # Filter column — local Data_Completeness mask; UV_Include is defined as $B$4#
    section_heading(sheet, _ROW_SECTION_HDR, _C_B, "Filter")
    val(sheet, _ROW_COL_HDRS, _C_B, "Include")
    f(sheet, _ROW_DATA_START, _C_B,
       "=MAP(LifeExpectancyData[Life expectancy],Data_Completeness)")
    sheet.range(rc(_ROW_DATA_START, _C_B), rc(_DATA_END, _C_B)).number_format = _FMT_INT


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
    ("Missing",   "=Missing_Count(UV_Data)"),
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
            _FMT_INT if label in ("Count", "Missing") else _FMT_1DP
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
    # Row 2 keeps the label and method in separate cells for formula references.
    section_heading(sheet, _ROW_METHOD_HDR, col_edge, "Method")
    val(sheet, _ROW_METHOD_HDR, col_count, method)
    sheet.range(rc(_ROW_METHOD_HDR, col_count)).color = _HEADER
    sheet.range(rc(_ROW_METHOD_HDR, col_count)).api.Font.Bold = True

    val(sheet, _ROW_COL_HDRS, col_edge,  "Upper Edge")
    val(sheet, _ROW_COL_HDRS, col_count, "Count")
    _subheader_row(sheet, _ROW_COL_HDRS, col_edge, col_count)

    # Row 3: bin-count display — "Bins" label + num_histogram_bins formula
    val(sheet, _ROW_SECTION_HDR, col_edge, "Bins:")
    method_cell = a1(_ROW_METHOD_HDR, col_count)
    f(sheet, _ROW_SECTION_HDR, col_count, f"=num_histogram_bins(UV_Data,{method_cell},UV_Include)")
    sheet.range(rc(_ROW_SECTION_HDR, col_count)).number_format = _FMT_INT

    # Spill formulas — method must be explicit (can't skip it to reach the 3rd filter arg)
    f(sheet, _ROW_HIST_START, col_edge, f"=Bin_Edges(UV_Data,{method_cell},UV_Include)")
    edge_spill_ref = f"{col_letter(col_edge)}{_ROW_HIST_START}#"
    f(sheet, _ROW_HIST_START, col_count,
       f"=Bin_Counts(UV_Data,{edge_spill_ref},UV_Include)")

    sheet.range(rc(_ROW_HIST_START, col_edge), rc(_DATA_END, col_edge)).number_format = _FMT_1DP
    sheet.range(rc(_ROW_HIST_START, col_count), rc(_DATA_END, col_count)).number_format = _FMT_INT


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
    _C_S: _FMT_1DP,
    _C_U: _FMT_1DP,
    _C_W: _FMT_1DP,
    _C_X: _FMT_SCI_1DP,
    _C_Y: _FMT_INT,
    _C_Z: _FMT_1DP,
    _C_AA: _FMT_1DP,
}


def _final_grid_best_refs(row_start: int) -> tuple[str, str]:
    """Return final-stage best column- and row-parameter cells for a grid block."""
    best_col = col_letter(_C_GS_S2 + _GS_C_BEST)
    return (
        f"=${best_col}${row_start + _GS_R_P1}",
        f"=${best_col}${row_start + _GS_R_P2}",
    )


def _nll_beta_rescaled_formula(alpha_ref: str, beta_ref: str) -> str:
    """NLL_Beta on min/max-rescaled data, corrected back to original scale."""
    return (
        "=IFERROR(LET("
        "d,FILTER(UV_Data,UV_Include),"
        "range_,MAX(d)-MIN(d),"
        "pad,range_*0.001,"
        "scale_,range_+2*pad,"
        "z,(d-MIN(d)+pad)/scale_,"
        f"NLL_Beta(z,{alpha_ref},{beta_ref})+COUNT(d)*LN(scale_)"
        "),1E+15)"
    )


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

    weibull_shape_ref, weibull_scale_ref = _final_grid_best_refs(_ROW_GS_WB)
    gamma_shape_ref, gamma_rate_ref = _final_grid_best_refs(_ROW_GS_GAMMA)
    beta_alpha_ref, beta_beta_ref = _final_grid_best_refs(_ROW_GS_BETA)

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
            # Params come from Stage 2 grid-search MLE (see Zone 5)
            "Weibull",
            "shape",
            weibull_shape_ref,
            "scale",
            weibull_scale_ref,
            "",    "",
            lambda r: f"=NLL_Weibull(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
        ),
        (
            # Params come from Stage 2 grid-search MLE (see Zone 5)
            "Gamma",
            "shape",
            gamma_shape_ref,
            "rate",
            gamma_rate_ref,
            "",    "",
            lambda r: f"=NLL_Gamma(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
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
            # Params come from Stage 2 grid-search MLE on min/max-rescaled data.
            "Beta",
            "alpha",
            beta_alpha_ref,
            "beta",
            beta_beta_ref,
            "",    "",
            lambda r: _nll_beta_rescaled_formula(_r(r), _t(r)),
            2,
        ),
        (
            "BetaPERT",
            "Min",  "=LET(d,FILTER(UV_Data,UV_Include),MIN(d)-(MAX(d)-MIN(d))*0.001)",
            "Mode", "=IFERROR(MODE.SNGL(FILTER(UV_Data,UV_Include)),MEDIAN(FILTER(UV_Data,UV_Include)))",
            "Max",  "=LET(d,FILTER(UV_Data,UV_Include),MAX(d)+(MAX(d)-MIN(d))*0.001)",
            lambda r: f"=NLL_BetaPERT(UV_Data,{_r(r)},{_t(r)},{_v(r)},UV_Include)",
            3,
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
    formula = f"=${aic_col_letter}{_ROW_DIST_START}=MIN({aic_min_range})"
    row_range = sheet.range(rc(_ROW_DIST_START, _C_Q), rc(last_row, _C_AA))
    cf = row_range.api.FormatConditions.Add(Type=2, Formula1=formula)  # xlExpression=2
    cf.Interior.Color = 0xC6EFCE  # light green fill
    cf.Font.Color     = 0x276228  # dark green text


# ── Histogram charts ──────────────────────────────────────────────────────────

def _add_histogram_chart(
    sheet: xw.Sheet,
    chart_left: float,
    chart_top: float,
    chart_width: float,
    chart_height: float,
    title_cell: str,
    edges_name: str,
    counts_name: str,
) -> None:
    """Insert one gapless column chart from the pre-binned histogram table."""
    sname = sheet.name
    chart_obj = sheet.charts.add(
        left=chart_left,
        top=chart_top,
        width=chart_width,
        height=chart_height,
    )
    chart = chart_obj.chart
    chart.api.ChartType = _XL_COLUMN_CLUSTERED
    chart.has_title = True
    title_ref = f"'{sname}'!${title_cell}$2"
    try:
        chart.api.ChartTitle.Formula = f"={title_ref}"
    except Exception:
        chart.chart_title.text = f"={title_ref}"

    while chart.series_collection.count > 0:
        chart.series_collection(1).delete()

    series = chart.series_collection.new_series()
    series.formula   = (
        f"=SERIES({title_ref},"
        f"'{sname}'!{edges_name},"
        f"'{sname}'!{counts_name},"
        f"1)"
    )

    try:
        series.api.GapWidth = 0
    except Exception:
        pass

    chart.axes(_XL_CATEGORY).has_title = True
    chart.axes(_XL_CATEGORY).axis_title.text = "Upper Edge"
    chart.axes(_XL_VALUE).has_title = True
    chart.axes(_XL_VALUE).axis_title.text = "Count"

    chart.has_legend = False


def _write_histogram_charts(sheet: xw.Sheet) -> None:
    """Add three histogram charts in the Q:AA chart band."""
    for title_cell, edges_name, counts_name, row_start, row_end in [
        ("H", "UV_Sturges_Edges", "UV_Sturges_Counts", 14, 32),
        ("K", "UV_Scott_Edges",   "UV_Scott_Counts",   34, 42),
        ("N", "UV_FD_Edges",      "UV_FD_Counts",      44, 52),
    ]:
        chart_range = sheet.range(rc(row_start, _C_Q), rc(row_end, _C_AA))
        _add_histogram_chart(
            sheet,
            chart_left=chart_range.left,
            chart_top=chart_range.top,
            chart_width=chart_range.width,
            chart_height=chart_range.height,
            title_cell=title_cell,
            edges_name=edges_name,
            counts_name=counts_name,
        )


# ── Zone 5: two-parameter grid-search MLE ────────────────────────────────────
#
# Layout within one stage block (row_start, col_start):
#   row+0  : section heading merged across the full stage width
#   row+1  : Min NLL and Rows/Columns labels; blank col+2; parameter headers at col+3…col+8
#   row+2  : control values; column parameter: Parameter | Input | Min | Max | Step Size | Best
#   row+3  : row parameter:                      Parameter | Input | Min | Max | Step Size | Best
#   row+4  : Data Table corner at col+0; column-parameter SEQUENCE spills across col+1…col+N
#   row+5…+4+N : row-parameter SEQUENCE at col+0; Data Table body at col+1…col+N
#
# Fixed-area tables:
#   Min NLL table       — col+0, rows+1…+2; value uses TAKE(Grid_Argmin(...),,1)
#   Rows/Columns table  — col+1, rows+1…+2; generated value documents physical grid size
#   Parameter table     — cols+3…+8, rows+1…+3
#   Blank spacer column — col+2, rows+1…+3
#   Best column         — Grid_Search_Optimum(...) spills from column parameter to row parameter
#
# The visible Input cells are the actual RowInput and ColumnInput cells supplied
# to Excel's two-input Data Table object.  No hidden auxiliary row is required.
#
# Stage 1: col_start = _C_GS = 29 (AC); body = AD6:AW25
# Stage 2: col_start = _C_GS_S2 = 51 (AY); body = AZ6:BS25
#
# Named ranges registered here:
#   *_S1 = Stage 1 Data Table body only
#   *_S2 = Stage 2 Data Table body only


def _gs_a1(row_start: int, col_start: int, dr: int, dc: int) -> str:
    """Absolute A1 ref offset from (row_start, col_start) by (dr, dc)."""
    return f"${col_letter(col_start + dc)}${row_start + dr}"


def _write_grid_stage(
    sheet: xw.Sheet,
    row_start: int,
    col_start: int,
    title: str,
    body_name: str,
    p1_label: str,
    p2_label: str,
    nll_formula,
    p1_min,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    p1_max,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    p2_min,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    p2_max,   # numeric default (Stage 1) or Excel formula string (Stage 2)
    editable_bounds: bool = True,
) -> dict:
    """Write one 20×20 NLL grid-search stage for a two-parameter fit.

    Returns absolute A1 references used by the refined Stage 2 search and by
    the distribution-fitting summary: best_p1, best_p2, min_p1, max_p1,
    min_p2, max_p2, step_p1, step_p2, n_grid, corner, p1_seq, and p2_seq.
    """
    sname = sheet.name
    n = _N_GRID
    r0, c0 = row_start, col_start
    last_col = c0 + n

    # ── Stage title ──────────────────────────────────────────────────────────
    val(sheet, r0, c0, title)
    sheet.range(rc(r0, c0), rc(r0, last_col)).merge()
    sheet.range(rc(r0, c0)).api.Font.Bold = True
    sheet.range(rc(r0, c0)).color = _HEADER

    control_hdr_row = r0 + _GS_R_CONTROL_HDR
    p1_row = r0 + _GS_R_P1
    p2_row = r0 + _GS_R_P2

    # ── Fixed-area tables ────────────────────────────────────────────────────
    # Left tables: Min NLL and the generated physical grid dimension.
    val(sheet, control_hdr_row, c0 + _GS_C_MINNLL, "Min NLL:")
    val(sheet, control_hdr_row, c0 + _GS_C_N_GRID, "Rows/Columns")
    _subheader_row(
        sheet,
        control_hdr_row,
        c0 + _GS_C_MINNLL,
        c0 + _GS_C_N_GRID,
    )
    val(sheet, p1_row, c0 + _GS_C_N_GRID, n)
    sheet.range(rc(p1_row, c0 + _GS_C_N_GRID)).number_format = _FMT_INT
    n_grid_ref = _gs_a1(r0, c0, _GS_R_P1, _GS_C_N_GRID)

    # Right table: Parameter | Input | Min | Max | Step Size | Best.
    for dc, label in [
        (_GS_C_PARAM, "Parameter"),
        (_GS_C_INPUT, "Input"),
        (_GS_C_MIN, "Min"),
        (_GS_C_MAX, "Max"),
        (_GS_C_STEP, "Step Size"),
        (_GS_C_BEST, "Best"),
    ]:
        val(sheet, control_hdr_row, c0 + dc, label)
    _subheader_row(
        sheet,
        control_hdr_row,
        c0 + _GS_C_PARAM,
        c0 + _GS_C_BEST,
    )

    val(sheet, p1_row, c0 + _GS_C_PARAM, p1_label)
    val(sheet, p2_row, c0 + _GS_C_PARAM, p2_label)

    # The visible Input cells are the Data Table's substitution cells.
    val(sheet, p1_row, c0 + _GS_C_INPUT, 1.0)
    val(sheet, p2_row, c0 + _GS_C_INPUT, 1.0)
    dt_row_ref = _gs_a1(r0, c0, _GS_R_P1, _GS_C_INPUT)
    dt_col_ref = _gs_a1(r0, c0, _GS_R_P2, _GS_C_INPUT)

    # Parameter bounds.
    for row, p_min, p_max in [
        (p1_row, p1_min, p1_max),
        (p2_row, p2_min, p2_max),
    ]:
        if isinstance(p_min, str):
            f(sheet, row, c0 + _GS_C_MIN, p_min)
        else:
            val(sheet, row, c0 + _GS_C_MIN, p_min)
        if isinstance(p_max, str):
            f(sheet, row, c0 + _GS_C_MAX, p_max)
        else:
            val(sheet, row, c0 + _GS_C_MAX, p_max)

        if editable_bounds:
            sheet.range(rc(row, c0 + _GS_C_MIN)).color = _INPUT
            sheet.range(rc(row, c0 + _GS_C_MAX)).color = _INPUT

    min_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _GS_C_MIN)
    max_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _GS_C_MAX)
    min_p2_ref = _gs_a1(r0, c0, _GS_R_P2, _GS_C_MIN)
    max_p2_ref = _gs_a1(r0, c0, _GS_R_P2, _GS_C_MAX)
    step_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _GS_C_STEP)
    step_p2_ref = _gs_a1(r0, c0, _GS_R_P2, _GS_C_STEP)
    f(sheet, p1_row, c0 + _GS_C_STEP,
      f"=({max_p1_ref}-{min_p1_ref})/({n_grid_ref}-1)")
    f(sheet, p2_row, c0 + _GS_C_STEP,
      f"=({max_p2_ref}-{min_p2_ref})/({n_grid_ref}-1)")

    fixed_values = sheet.range(
        rc(p1_row, c0 + _GS_C_INPUT),
        rc(p2_row, c0 + _GS_C_BEST),
    )
    fixed_values.number_format = _FMT_1DP

    # ── Data Table axes and body ─────────────────────────────────────────────
    hdr_row = r0 + _GS_R_HDR
    body_row_start = r0 + _GS_R_BODY
    body_row_end = body_row_start + n - 1
    body_col_start = c0 + 1
    body_col_end = c0 + n

    corner_ref = _gs_a1(r0, c0, _GS_R_HDR, 0)
    p1_seq_ref = _gs_a1(r0, c0, _GS_R_HDR, 1)
    p2_seq_ref = _gs_a1(r0, c0, _GS_R_BODY, 0)

    # Corner formula references the visible Input cells.
    f(
        sheet,
        hdr_row,
        c0,
        nll_formula(dt_row_ref, dt_col_ref),
    )
    sheet.range(rc(hdr_row, c0)).number_format = _FMT_SCI_1DP

    # Column parameter (shape) values immediately above the body.
    f(
        sheet,
        hdr_row,
        c0 + 1,
        f"=SEQUENCE(1,{n_grid_ref},{min_p1_ref},{step_p1_ref})",
    )
    sheet.range(rc(hdr_row, body_col_start), rc(hdr_row, body_col_end)).number_format = _FMT_1DP

    # Row parameter (scale) values immediately to the left of the body.
    f(
        sheet,
        body_row_start,
        c0,
        f"=SEQUENCE({n_grid_ref},1,{min_p2_ref},{step_p2_ref})",
    )
    sheet.range(rc(body_row_start, c0), rc(body_row_end, c0)).number_format = _FMT_1DP

    body_range = sheet.range(
        rc(body_row_start, body_col_start),
        rc(body_row_end, body_col_end),
    )
    drop_local_name(sheet, body_name)
    sheet.api.Names.Add(
        Name=body_name,
        RefersTo=(
            f"='{sname}'!${col_letter(body_col_start)}${body_row_start}"
            f":${col_letter(body_col_end)}${body_row_end}"
        ),
    )

    full_table_range = sheet.range(
        rc(hdr_row, c0),
        rc(body_row_end, body_col_end),
    )
    try:
        full_table_range.api.Table(
            RowInput=sheet.range(rc(p1_row, c0 + _GS_C_INPUT)).api,
            ColumnInput=sheet.range(rc(p2_row, c0 + _GS_C_INPUT)).api,
        )
    except Exception as e:
        raise RuntimeError(f"Failed to wire two-input Data Table for stage {title!r}") from e

    body_range.number_format = _FMT_SCI_1DP

    # ── LAMBDA-driven outputs ────────────────────────────────────────────────
    # Min NLL is the first column returned by Grid_Argmin.
    f(
        sheet,
        p1_row,
        c0 + _GS_C_MINNLL,
        f'=IFERROR(TAKE(Grid_Argmin({body_name}),,1),"—")',
    )
    sheet.range(rc(p1_row, c0 + _GS_C_MINNLL)).number_format = _FMT_SCI_1DP

    # Best column-parameter and row-parameter values spill vertically.
    f(
        sheet,
        p1_row,
        c0 + _GS_C_BEST,
        f"=Grid_Search_Optimum({body_name})",
    )

    # Boundary guard: red fill when the optimum lies on a grid edge.
    boundary_specs = [
        (p1_row, 3),  # column location controls the column parameter
        (p2_row, 2),  # row location controls the row parameter
    ]
    for row, argmin_col in boundary_specs:
        location = f"INDEX(Grid_Argmin({body_name}),1,{argmin_col})"
        cell_api = sheet.range(rc(row, c0 + _GS_C_BEST)).api
        cf = cell_api.FormatConditions.Add(
            Type=2,  # xlExpression
            Formula1=f"=OR({location}=1,{location}={n_grid_ref})",
        )
        cf.Interior.Color = 0x0000FF   # red (BGR)
        cf.Font.Color = 0xFFFFFF       # white

    # Heatmap: green=low NLL → yellow=mid → red=high / overflow.
    try:
        body_range.api.FormatConditions.Delete()
        cs = body_range.api.FormatConditions.AddColorScale(3)
        cs.ColorScaleCriteria(1).Type = 1
        cs.ColorScaleCriteria(1).FormatColor.Color = 0x63BE7B
        cs.ColorScaleCriteria(2).Type = 5
        cs.ColorScaleCriteria(2).Value = 50
        cs.ColorScaleCriteria(2).FormatColor.Color = 0xFFEB84
        cs.ColorScaleCriteria(3).Type = 2
        cs.ColorScaleCriteria(3).FormatColor.Color = 0xF8696B
    except Exception:
        pass

    # Separate borders for the two fixed-area tables and the Data Table.
    border_box(
        sheet,
        control_hdr_row,
        c0 + _GS_C_MINNLL,
        p1_row,
        c0 + _GS_C_MINNLL,
    )
    border_box(
        sheet,
        control_hdr_row,
        c0 + _GS_C_N_GRID,
        p1_row,
        c0 + _GS_C_N_GRID,
    )
    border_box(
        sheet,
        control_hdr_row,
        c0 + _GS_C_PARAM,
        p2_row,
        c0 + _GS_C_BEST,
    )
    border_box(sheet, hdr_row, c0, body_row_end, body_col_end)

    return {
        "best_p1": _gs_a1(r0, c0, _GS_R_P1, _GS_C_BEST),
        "best_p2": _gs_a1(r0, c0, _GS_R_P2, _GS_C_BEST),
        "min_p1": _gs_a1(r0, c0, _GS_R_P1, _GS_C_MIN),
        "max_p1": _gs_a1(r0, c0, _GS_R_P1, _GS_C_MAX),
        "min_p2": _gs_a1(r0, c0, _GS_R_P2, _GS_C_MIN),
        "max_p2": _gs_a1(r0, c0, _GS_R_P2, _GS_C_MAX),
        "step_p1": step_p1_ref,
        "step_p2": step_p2_ref,
        "n_grid": n_grid_ref,
        "corner": corner_ref,
        "p1_seq": p1_seq_ref,
        "p2_seq": p2_seq_ref,
    }


def _write_two_stage_grid_search(
    sheet: xw.Sheet,
    row_start: int,
    dist_name: str,
    body_prefix: str,
    p1_label: str,
    p2_label: str,
    nll_formula,
    p1_min,
    p1_max,
    p2_min,
    p2_max,
) -> None:
    """Write Stage 1 and Stage 2 grid-search blocks for one distribution."""
    s1 = _write_grid_stage(
        sheet,
        row_start   = row_start,
        col_start   = _C_GS,
        title       = f"{dist_name} Grid-Search MLE  —  Stage 1  ({p1_label} × {p2_label})",
        body_name   = f"{body_prefix}_S1",
        p1_label    = p1_label,
        p2_label    = p2_label,
        nll_formula = nll_formula,
        p1_min      = p1_min,
        p1_max      = p1_max,
        p2_min      = p2_min,
        p2_max      = p2_max,
        editable_bounds = True,
    )

    _write_grid_stage(
        sheet,
        row_start   = row_start,
        col_start   = _C_GS_S2,
        title       = f"{dist_name} Grid-Search MLE  —  Stage 2  (refined)",
        body_name   = f"{body_prefix}_S2",
        p1_label    = p1_label,
        p2_label    = p2_label,
        nll_formula = nll_formula,
        p1_min      = f"=MAX(0.001,{s1['best_p1']}-{s1['step_p1']})",
        p1_max      = f"={s1['best_p1']}+{s1['step_p1']}",
        p2_min      = f"=MAX(0.001,{s1['best_p2']}-{s1['step_p2']})",
        p2_max      = f"={s1['best_p2']}+{s1['step_p2']}",
        editable_bounds = False,
    )

    # Section heading label over the gap column between stages.
    val(sheet, row_start, _C_GS + _GS_W, "")


def _write_two_parameter_grid_search(sheet: xw.Sheet) -> None:
    """Write two-stage grid-search MLE blocks for all two-parameter fits."""
    _write_two_stage_grid_search(
        sheet,
        row_start   = _ROW_GS_WB,
        dist_name   = "Weibull",
        body_prefix = "UV_WB",
        p1_label    = "Shape (k)",
        p2_label    = "Scale (λ)",
        nll_formula = lambda p1, p2: f"=NLL_Weibull(UV_Data,{p1},{p2},UV_Include)",
        p1_min      = 0.5,
        p1_max      = 10.0,
        p2_min      = 0.1,
        p2_max      = "=IFERROR(2*AVERAGE(FILTER(UV_Data,UV_Include)),10)",
    )

    _write_two_stage_grid_search(
        sheet,
        row_start   = _ROW_GS_GAMMA,
        dist_name   = "Gamma",
        body_prefix = "UV_GAMMA",
        p1_label    = "Shape (α)",
        p2_label    = "Rate (β)",
        nll_formula = lambda p1, p2: f"=NLL_Gamma(UV_Data,{p1},{p2},UV_Include)",
        p1_min      = 0.5,
        p1_max      = 100.0,
        p2_min      = 0.001,
        p2_max      = 2.0,
    )

    _write_two_stage_grid_search(
        sheet,
        row_start   = _ROW_GS_BETA,
        dist_name   = "Beta",
        body_prefix = "UV_BETA",
        p1_label    = "Alpha (α)",
        p2_label    = "Beta (β)",
        nll_formula = _nll_beta_rescaled_formula,
        p1_min      = 0.2,
        p1_max      = 50.0,
        p2_min      = 0.2,
        p2_max      = 50.0,
    )


def _write_weibull_grid_search(sheet: xw.Sheet) -> None:
    """Compatibility wrapper for the two-parameter grid-search section."""
    _write_two_parameter_grid_search(sheet)


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

    _write_weibull_grid_search(sheet)
    _autofit_column_widths(sheet)

    # Charts (skipped silently if chart API raises; e.g. headless builds)
    try:
        _write_histogram_charts(sheet)
    except Exception:
        pass

    _finalize_sheet(sheet)

    return sheet

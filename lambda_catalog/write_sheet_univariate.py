"""
write_sheet_univariate.py
Writes the Univariate Analysis sheet into any target workbook.

Design decision — fitting approach
───────────────────────────────────
MLE throughout.  For Normal, Lognormal, and Exponential the MLE estimators are
closed-form (sample mean/SD on the raw or log-transformed data).  For
Triangular and BetaPERT the likelihood is non-differentiable at the mode, so
direct min/mode/max estimation is used — the result is a valid parameter set
that the NLL formula evaluates against.  Weibull, Gamma, and Beta use two-stage
searches.

**Weibull and Gamma search one dimension** (the grid shrink).  Their scale /
rate parameter is *profiled out* in closed form at every trial shape, so each
stage evaluates a 20-point profile-NLL curve rather than a 20×20 surface:

  Weibull   λ̂(k) = ((1/n)·Σxᵢᵏ)^(1/k)
  Gamma     β̂(α) = α / x̄

Profiling is still genuine MLE — the profile maximizer is the joint maximizer —
so the two returned parameters are the same estimates the 2-D grid produced,
at 20 evaluations per stage instead of 400.  Each stage 1 brackets a
closed-form starting value (Weibull: probability-plot regression of
ln(−ln(1−F̂)) on ln x, F̂ the Hazen positions Zone 5 already uses; Gamma:
Minka's approximation from s = ln(AM/GM)); stage 2 refines to ±1 stage-1 step.  The two-dimensional
NLL heatmap is replaced by a **profile-NLL line chart** per distribution.

**Beta stays two-dimensional** — both conditional MLEs involve digamma — and
keeps its two two-input Data Tables.  They are the only Data Tables in the
artifact.

Sheet layout
────────────
  Row 1          — Title "Univariate Analysis"
  Row 2          — Distribution Fitting/Comparison heading (G:S) | Method labels for histograms
  Row 3          — Section headings: Data | Filter | Descriptive Statistics | Bins; pane freeze
  Row 4          — Column sub-headers

  Col A          — Data: LifeExpectancyData[Life expectancy] spill in A4
  Col B          — Filter: numeric-data mask ($B$4#)
  Col C          — thin gap (width 2); freeze pane left boundary
  Col D–E        — Descriptive Statistics (12 stat rows)
  Col F          — thin gap (width 2)
  Col G–S        — Distribution Fitting summary table (Name,θ₁/₂/₃ labels+values,NLL,k,AIC,BIC,A-D,K-S)
  Col T          — thin gap (width 2)
  Col U–AE       — Sturges histogram (11 cols: Lower Edges,Upper Edges,Count,Normal…BetaPERT CDF probabilities)
  Col AF         — thin gap (width 2)
  Col AG–AQ      — Scott histogram (11 cols)
  Col AR         — thin gap (width 2)
  Col AS–BC      — Freedman-Diaconis histogram (11 cols)
  Col BD         — thin gap (width 2)

  The band from BE right is one zone per topic, each followed by a gap column,
  all derived from the ordered _BAND_ZONES table — never hard-code these:

  Col BE–BN      — Q-Q plot data (P, sorted Sample, 8 theoretical-quantile columns)
  Col BP–BX      — Weibull fit zone: both profile stages (9 cols, rows 1–31)
  Col BZ–CH      — Gamma fit zone: both profile stages (9 cols, rows 1–31)
  Col CJ–DD      — Beta fit zone: both 20×20 Data Table stages (21 cols, rows 1–51)

  One zone per distribution, with that fit's two stages stacked inside it, so a
  fit reads as a single self-contained block.  The fit zones come last because
  they are the only ones whose width is a tunable (_N_GRID / _N_PROFILE), so a
  resize displaces nothing.

  Charts anchored under the fitting table:
    G14, G34, G54          — histogram combo charts (count bars + 8 fitted overlay lines)
    rows 74-153            — eight per-distribution Q-Q plots in a 4×2 grid
                             (left chart G:O, right chart P:T per row)
    rows 154-173           — Weibull and Gamma profile-NLL line charts
                             (left chart G:O, right chart P:T)

Sheet-scoped named ranges
─────────────────────────
  UV_Data        — spill range of the raw column formula ($A$4#, unfiltered)
  UV_Include     — local filter mask spill ($B$4#)
  UV_n           — IFERROR(COUNT(FILTER(UV_Data,UV_Include)), 0)
  UV_Sturges_*, UV_Scott_*, UV_FD_* — OFFSET-based histogram column ranges
  UV_*_<Dist>_Expected — histogram CDF-delta column × Count stat cell (expected
                   counts); feeds the fitted-distribution overlay line series
  UV_QQ_Sample, UV_QQ_<Dist> — OFFSET-based Q-Q chart column ranges
  UV_WB_S1/S2    — Stage 1/2 Weibull profile-NLL bodies (20×1), side by side
  UV_WB_S1/S2_Axis — Stage 1/2 Weibull shape axes (20×1), left of each body
  UV_GAMMA_S1/S2  — Stage 1/2 Gamma profile-NLL bodies (20×1), side by side
  UV_GAMMA_S1/S2_Axis — Stage 1/2 Gamma shape axes (20×1)
  UV_BETA_S1/S2   — Stage 1/2 Beta Data Table bodies (20×20)
  UV_Profile_WB_*, UV_Profile_GAMMA_* — OFFSET-based stage-1 axis/NLL ranges
                   feeding the two profile-NLL line charts
"""
from __future__ import annotations

from typing import NamedTuple

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT, SUBHDR_COLOR as _SUBHDR
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS, a1, anchor_comment_right_of_cell, border_box,
    col_letter, drop_local_name, excel_color, f, note_dimensions, rc,
    section_heading, set_column_widths, val,
)

# ── Column indices (1-based) ─────────────────────────────────────────────────

# Zone 1: Data Input + Filter
_C_A = 1    # data values
_C_B = 2    # local filter mask over the source table column

# Zone 2: Descriptive Statistics
_C_D = 4    # stat labels
_C_E = 5    # stat values

# Zone 3: Distribution Fitting (G:S = cols 7-19)
_C_DIST_NAME = 7    # G — distribution name
_C_T1_LBL    = 8    # H — θ₁ label
_C_T1_VAL    = 9    # I — θ₁ value
_C_T2_LBL    = 10   # J — θ₂ label
_C_T2_VAL    = 11   # K — θ₂ value
_C_T3_LBL    = 12   # L — θ₃ label
_C_T3_VAL    = 13   # M — θ₃ value
_C_NLL       = 14   # N — NLL
_C_K_PARAM   = 15   # O — k (param count)
_C_AIC       = 16   # P — AIC
_C_BIC       = 17   # Q — BIC
_C_AD        = 18   # R — A-D
_C_KS        = 19   # S — K-S

_C_FIT_FIRST = _C_DIST_NAME   # first fitting column
_C_FIT_LAST  = _C_KS          # last fitting column
_C_QQ_PLOT_AREA_END = _C_K_PARAM   # Sets boundary for Q-Q plots to keep them more square.

# Zone 4: Histogram Tables (11 cols each, with gap cols between)
_HIST_W = 11   # columns per histogram block

_C_STUR  = 21   # U  — first col of Sturges block (U:AE = 21-31)
_C_SCOTT = 33   # AG — first col of Scott block (AG:AQ = 33-43)
_C_FD    = 45   # AS — first col of FD block (AS:BC = 45-55)

# Within each 11-col histogram block (offsets from block start)
_HB_LOWER = 0   # Lower Edges (pre-computed once per block; referenced by all CDF columns)
_HB_EDGE  = 1   # Upper Edges
_HB_COUNT = 2   # Count
_HB_NORM  = 3   # Normal CDF prob
_HB_LOGN  = 4   # Lognormal CDF prob
_HB_EXP   = 5   # Exponential CDF prob
_HB_WB    = 6   # Weibull CDF prob
_HB_GAM   = 7   # Gamma CDF prob
_HB_TRI   = 8   # Triangular CDF prob
_HB_BETA  = 9   # Beta CDF prob
_HB_PERT  = 10  # BetaPERT CDF prob

class _BlockColumnSpec(NamedTuple):
    """One column's facts within a repeating block: offset, header text,
    named-range suffix, and distribution label (empty for non-CDF columns).

    Callers unpack these by position (``for offset, header, suffix,
    distribution in _HIST_COLUMNS``), so field order must stay
    offset/header/suffix/distribution.
    """

    offset: int
    header: str
    suffix: str
    distribution: str = ""


_HIST_COLUMNS = [
    _BlockColumnSpec(_HB_LOWER, "Lower Edges", "Lower_Edges", ""),
    _BlockColumnSpec(_HB_EDGE,  "Upper Edges", "Upper_Edges", ""),
    _BlockColumnSpec(_HB_COUNT, "Count", "Counts", ""),
    _BlockColumnSpec(_HB_NORM,  "Normal", "Normal_CDF", "Normal"),
    _BlockColumnSpec(_HB_LOGN,  "Lognormal", "Lognormal_CDF", "Lognormal"),
    _BlockColumnSpec(_HB_EXP,   "Exponential", "Exponential_CDF", "Exponential"),
    _BlockColumnSpec(_HB_WB,    "Weibull", "Weibull_CDF", "Weibull"),
    _BlockColumnSpec(_HB_GAM,   "Gamma", "Gamma_CDF", "Gamma"),
    _BlockColumnSpec(_HB_TRI,   "Triangular", "Triangular_CDF", "Triangular"),
    _BlockColumnSpec(_HB_BETA,  "Beta", "Beta_CDF", "Beta"),
    _BlockColumnSpec(_HB_PERT,  "BetaPERT", "BetaPERT_CDF", "BetaPERT"),
]

_HIST_BLOCKS = [
    ("UV_Sturges", _C_STUR),
    ("UV_Scott", _C_SCOTT),
    ("UV_FD", _C_FD),
]

# Zone 6: two-parameter MLE search — Beta's 2-D grid dimensions
_N_GRID   = 20             # grid points per axis per stage (20×20 = 400/stage)
_GS_W     = _N_GRID + 1   # cols per Beta stage = 21  (1 row-axis col + N body cols)

# Within-stage row offsets from block row_start
_GS_R_CONTROL_HDR = 1   # Min NLL label + parameter-table headers
_GS_R_P1          = 2   # column parameter row
_GS_R_P2          = 3   # row parameter row
_GS_R_HDR         = 4   # Data Table corner + column-parameter SEQUENCE
_GS_R_BODY        = 5   # first Data Table body row

# Beta's two stages stack vertically inside its own column zone.
_GS_BLOCK_H      = _GS_R_BODY + _N_GRID
_GS_BLOCK_GAP_R  = 1
_GS_R_STAGE2     = _GS_BLOCK_H + _GS_BLOCK_GAP_R   # 26 rows below Stage 1

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

# ── One-dimensional profile-NLL search (Weibull, Gamma) ──────────────────────
#
# A profile zone reuses the 2-D stage's column offsets so the two writers read
# alike; only offset 4 differs in meaning (the Data Table substitution `Input`
# becomes the closed-form `Start`).  The profiled-out parameter's Min/Max/Step
# cells are left empty because it is solved, not searched.
#
# Unlike the 2-D stage, a profile stage's control block and its body do not
# share a column origin: the two stages' control blocks stack vertically in one
# 9-column band, and their two bodies sit side by side underneath, each under
# one half of the control block and sharing its offset-2 spacer.
_N_PROFILE   = 20   # profile points per stage (20 per stage, not 20×20)
_PS_C_MINNLL = _GS_C_MINNLL   # Min NLL label/value
_PS_C_POINTS = _GS_C_N_GRID   # "Grid Points" label/value
_PS_C_PARAM  = _GS_C_PARAM    # Parameter
_PS_C_START  = _GS_C_INPUT    # closed-form starting value for the searched parameter
_PS_C_MIN    = _GS_C_MIN      # searched-parameter range minimum
_PS_C_MAX    = _GS_C_MAX      # searched-parameter range maximum
_PS_C_STEP   = _GS_C_STEP     # endpoint-inclusive step size
_PS_C_BEST   = _GS_C_BEST     # best searched / profiled parameter values
_PS_W        = _PS_C_BEST + 1   # columns a profile zone occupies (9)

# Profile-stage row offsets from the zone's row anchor.  A control block is
# four rows (title, header, searched param, profiled param); the stages stack
# with one blank row between, then one blank row before the shared body.
_PS_R_CONTROL_H    = _GS_R_P2 + 1              # 4
_PS_R_STAGE_STRIDE = _PS_R_CONTROL_H + 1       # 5 — Stage 2 title row offset
_PS_R_BODY_HDR     = 2 * _PS_R_STAGE_STRIDE    # 10 — axis / Profile NLL headers
_PS_R_BODY         = _PS_R_BODY_HDR + 1        # 11 — first body row

# Body column pairs (axis, NLL) per stage, relative to the zone's first column.
# Stage 1 sits under the Min NLL / Grid Points half of the control block and
# Stage 2 under the Parameter / Start half, sharing the offset-2 spacer.
_PS_BODY_COLS = (
    (_PS_C_MINNLL, _PS_C_POINTS),
    (_PS_C_PARAM, _PS_C_START),
)

# Stage 1 brackets the closed-form start by this multiplicative factor either
# side.  Both starting estimates land within a few percent of the MLE, so 3×
# is a wide bracket; the boundary-hit guard on the Best cell flags the rare
# sample where it is not, and the Min/Max cells stay editable overrides.
_PROFILE_BRACKET = 3

# ── Right-hand zone band ─────────────────────────────────────────────────────
#
# One ordered table is the single source of truth for every column in the band
# right of the histograms: zone starts, gap columns, and the sheet's last
# column all derive from it, and reordering the band means reordering this list
# and nothing else.  Never hard-code a column letter from it.
#
# The three fit zones come last on purpose.  They are the only zones whose
# width is a tunable — _N_GRID and _N_PROFILE set it, and Beta's pending shrink
# to a ~12×12 grid will change its width by nine columns — so keeping them at
# the end means a resize displaces nothing.  Same reasoning as the Regression
# sheet's rule that nothing may sit right of the design-matrix zone.  The Q-Q
# data is a fixed 10 columns and is read alongside the fitting table, so it
# leads.  Each zone is followed by one gap column; the histogram band's own
# trailing gap (BD) serves as the leading one.
#
# The band is Zone 5 (Q-Q data) then Zone 6 (the three fit zones) — the two
# swapped physical order when Q-Q moved ahead of the fits, and the numbering
# was swapped with them so it still reads left to right.
_QQ_W = 10   # Zone 5 width: P, Sample, 8 theoretical-quantile columns

_BAND_FIRST_COL = 57   # BE — first column after the FD histogram block's gap
_BAND_GAP_W = 1
_BAND_ZONES = (
    ("qq", _QQ_W),        # BE–BN
    ("weibull", _PS_W),   # BP–BX
    ("gamma", _PS_W),     # BZ–CH
    ("beta", _GS_W),      # CJ–DD
)


def _derive_band_columns() -> tuple[dict[str, int], tuple[int, ...], int]:
    """Return (zone first columns, gap columns, last populated column)."""
    starts: dict[str, int] = {}
    gaps: list[int] = []
    col = _BAND_FIRST_COL
    for index, (name, width) in enumerate(_BAND_ZONES):
        starts[name] = col
        col += width
        if index < len(_BAND_ZONES) - 1:
            gaps.append(col)
            col += _BAND_GAP_W
    return starts, tuple(gaps), col - 1


_BAND_COL, _BAND_GAP_COLS, _BAND_LAST_COL = _derive_band_columns()

_C_QQ        = _BAND_COL["qq"]        # 57  (BE)
_C_GS_WB     = _BAND_COL["weibull"]   # 68  (BP)
_C_GS_GAMMA  = _BAND_COL["gamma"]     # 78  (BZ)
_C_GS_BETA   = _BAND_COL["beta"]      # 88  (CJ)

# Within the Q-Q block: (offset, header, name suffix, distribution)
_QQ_P      = 0   # Hazen plotting positions (i - 0.5)/n
_QQ_SAMPLE = 1   # sorted included sample values (Y axis of every Q-Q chart)
_QQ_COLUMNS = [
    _BlockColumnSpec(_QQ_P,      "P", "P", ""),
    _BlockColumnSpec(_QQ_SAMPLE, "Sample", "Sample", ""),
    _BlockColumnSpec(2,  "Normal", "Normal", "Normal"),
    _BlockColumnSpec(3,  "Lognormal", "Lognormal", "Lognormal"),
    _BlockColumnSpec(4,  "Exponential", "Exponential", "Exponential"),
    _BlockColumnSpec(5,  "Weibull", "Weibull", "Weibull"),
    _BlockColumnSpec(6,  "Gamma", "Gamma", "Gamma"),
    _BlockColumnSpec(7,  "Triangular", "Triangular", "Triangular"),
    _BlockColumnSpec(8,  "Beta", "Beta", "Beta"),
    _BlockColumnSpec(9,  "BetaPERT", "BetaPERT", "BetaPERT"),
]

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
_XL_COLUMN_CLUSTERED = 51      # xlColumnClustered
_XL_LINE             = 4       # xlLine — overlay series in the histogram combo charts
_XL_XY_SCATTER       = -4169   # xlXYScatter — Q-Q plots
_XL_XY_SCATTER_LINES = 74      # xlXYScatterLines — profile-NLL curves
_XL_XY_SCATTER_LINES_NO_MARKERS = 75   # identity reference series on Q-Q plots
_XL_CATEGORY         = 1    # horizontal axis type
_XL_VALUE            = 2    # vertical axis type
_XL_LEGEND_BOTTOM    = -4107   # xlLegendPositionBottom
_XL_MARKER_NONE      = -4142   # xlMarkerStyleNone

# Chart title cells — row where chart title formula and chart block begin
_ROW_CHART1_TITLE = 14   # G14 — Sturges histogram chart title
_ROW_CHART2_TITLE = 34   # G34 — Scott histogram chart title
_ROW_CHART3_TITLE = 54   # G54 — FD histogram chart title

# Q-Q plot charts sit in a 4×2 grid below the FD histogram chart: four rows
# of two charts, the left chart spanning cols G:O and the right cols P:T
# (rows 74-93, 94-113, 114-133, 134-153).
_ROW_QQ_CHART_START = _ROW_CHART3_TITLE + 20   # 74
_QQ_CHART_ROWS      = 20
_QQ_CHART_BANDS = (
    (7, 15),    # G:O — left chart of each pair
    (16, 20),   # P:T — right chart of each pair
)

# The two profile-NLL line charts (Weibull, Gamma) sit in one more row of the
# same two bands, directly under the last Q-Q pair.
_N_QQ_CHART_ROWS_USED = 4
_ROW_PROFILE_CHART_START = _ROW_QQ_CHART_START + _N_QQ_CHART_ROWS_USED * _QQ_CHART_ROWS
_PROFILE_CHART_ROWS = _QQ_CHART_ROWS

UNIVARIATE_SHEET_NAME = "Univariate"

_FMT_INT = "0"
_FMT_1DP = "0.0"
_FMT_4DP = "0.0000"
_FMT_SCI_1DP = "0.0E+00"


def _subheader_row(sheet: xw.Sheet, row: int, c1: int, c2: int) -> None:
    sheet.range(rc(row, c1), rc(row, c2)).color = _SUBHDR
    sheet.range(rc(row, c1), rc(row, c2)).api.Font.Bold = True


# ── Column widths ─────────────────────────────────────────────────────────────

# Per-offset widths within one 11-col histogram block (Lower/Upper/Count get
# their own widths; the 8 CDF probability columns share one width).
_HIST_BLOCK_WIDTHS: dict[int, float] = {
    _HB_LOWER: 12,
    _HB_EDGE: 12,
    _HB_COUNT: 10,
    **{offset: 12 for offset in range(_HB_NORM, _HIST_W)},
}


def _set_column_widths(sheet: xw.Sheet) -> None:
    widths = {
        _C_A: 14,         # data
        _C_B: 10,         # filter mask
        3:    2,          # gap (C) between filter and stats
        _C_D: 16,         # stat labels
        _C_E: 12,         # stat values
        6:    2,          # gap (F)
        _C_DIST_NAME: 14, # distribution name (G)
        _C_T1_LBL: 10,   # θ₁ label (H)
        _C_T1_VAL: 12,   # θ₁ value (I)
        _C_T2_LBL: 10,   # θ₂ label (J)
        _C_T2_VAL: 12,   # θ₂ value (K)
        _C_T3_LBL: 10,   # θ₃ label (L)
        _C_T3_VAL: 12,   # θ₃ value (M)
        _C_NLL: 12,      # NLL (N)
        _C_K_PARAM: 6,   # k (O)
        _C_AIC: 12,      # AIC (P)
        _C_BIC: 12,      # BIC (Q)
        _C_AD: 12,       # A-D (R)
        _C_KS: 12,       # K-S (S)
        20:   2,          # gap (T)
    }
    set_column_widths(sheet, widths.items())

    # Histogram blocks: 11 cols each with gap cols between/after
    for block_start in (_C_STUR, _C_SCOTT, _C_FD):
        set_column_widths(
            sheet, ((block_start + offset, w) for offset, w in _HIST_BLOCK_WIDTHS.items())
        )
        gap_col = block_start + _HIST_W
        set_column_widths(sheet, ((gap_col, 2),))

    # Q-Q plot data block (the histogram band's trailing gap precedes it).
    set_column_widths(sheet, ((_C_QQ + offset, 12) for offset in range(_QQ_W)))

    # Fit zones: compact fixed-area controls above the narrow bodies. Every
    # zone shares the same control-column offsets, so one loop covers the two
    # 9-column profile zones and Beta's 21-column grid zone; only Beta has
    # columns past _PS_W, and those are its Data Table body.
    for zone_start, zone_width in (
        (_C_GS_WB, _PS_W),
        (_C_GS_GAMMA, _PS_W),
        (_C_GS_BETA, _GS_W),
    ):
        set_column_widths(sheet, ((c, 6) for c in range(zone_start, zone_start + zone_width)))
        set_column_widths(
            sheet,
            (
                (zone_start, 12),                    # Min NLL / row axis / profile axis
                (zone_start + _GS_C_N_GRID, 14),      # Rows/Columns / Grid Points / profile NLL
                (zone_start + _GS_C_PARAM, 13),       # Parameter / profile Stage 2 axis
            ),
        )
        set_column_widths(
            sheet, ((zone_start + dc, 10) for dc in (_GS_C_INPUT, _GS_C_MIN, _GS_C_MAX, _GS_C_STEP, _GS_C_BEST))
        )

    set_column_widths(sheet, ((gap_col, 2) for gap_col in _BAND_GAP_COLS))


def _autofit_column_widths(sheet: xw.Sheet) -> None:
    """Autofit all populated layout columns, then restore intentional gaps."""
    sheet.range(rc(_ROW_TITLE, _C_A), rc(_DATA_END, _BAND_LAST_COL)).columns.autofit()

    gap_cols = [3, 6, 20]  # C, F, T
    for block_start in (_C_STUR, _C_SCOTT, _C_FD):
        gap_cols.append(block_start + _HIST_W)  # AF, AR, BD
    gap_cols.extend(_BAND_GAP_COLS)             # BO, BY, CI
    for col in gap_cols:
        sheet.range(rc(1, col), rc(1, col)).column_width = 2


# ── Sheet-scoped named range management ──────────────────────────────────────

def _drop_wb_name(sheet: xw.Sheet, name: str) -> None:
    """Remove a workbook-scoped (non-local) name if present."""
    wb_names = sheet.book.api.Names
    try:
        n = wb_names(name)
    except OPEN_WORKBOOK_ERRORS:
        return
    if "!" not in n.Name and n.Name.lower() == name.lower():
        n.Delete()


def _setup_local_names(sheet: xw.Sheet) -> None:
    """Register sheet-scoped named ranges used by formulas and charts."""
    sname = sheet.name

    # Remove obsolete workbook-scoped definitions before recreating local names.
    local_names = ["UV_Data", "UV_Include", "UV_n"]
    for prefix, _ in _HIST_BLOCKS:
        local_names.extend(f"{prefix}_{suffix}" for _, _, suffix, _ in _HIST_COLUMNS)
        local_names.extend(
            f"{prefix}_{distribution}_Expected"
            for _, _, _, distribution in _HIST_COLUMNS
            if distribution
        )
    local_names.extend(
        f"UV_QQ_{suffix}" for _, _, suffix, _ in _QQ_COLUMNS if suffix != "P"
    )

    for name in local_names:
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

    # Histogram column ranges: OFFSET-based, anchored at the header row like
    # Regression chart ranges, and sized by the method value stored in row 2.
    # Upper_Edges and Counts feed the histogram chart SERIES formulas, so they
    # get a Name Manager comment saying which chart they belong to.
    method_labels = {
        "UV_Sturges": "Sturges",
        "UV_Scott": "Scott",
        "UV_FD": "Freedman-Diaconis",
    }
    for prefix, block_start in _HIST_BLOCKS:
        method_ref = f"'{sname}'!${col_letter(block_start + _HB_COUNT)}${_ROW_METHOD_HDR}"
        size_formula = f"Number_Of_Histogram_Bins(UV_Data,{method_ref},UV_Include)"
        chart_comments = {
            "Upper_Edges": (
                f"{method_labels[prefix]} Method histogram chart: "
                "category (X) axis bin edges"
            ),
            "Counts": (
                f"{method_labels[prefix]} Method histogram chart: "
                "bar values (bin counts)"
            ),
        }
        for offset, _, suffix, distribution in _HIST_COLUMNS:
            name = f"{prefix}_{suffix}"
            col_ltr = col_letter(block_start + offset)
            _drop_wb_name(sheet, name)
            drop_local_name(sheet, name)
            nm = sheet.api.Names.Add(
                Name=name,
                RefersTo=(
                    f"=OFFSET('{sname}'!${col_ltr}${_ROW_COL_HDRS},"
                    f"1,0,MAX(IFERROR({size_formula},1),1),1)"
                ),
            )
            if suffix in chart_comments:
                nm.Comment = chart_comments[suffix]

            if not distribution:
                continue

            # Expected-count named formula: the CDF-delta column (per-bin
            # probability mass) scaled by the Count stat cell.  This puts the
            # fitted-distribution overlay line series on the same count axis
            # as the histogram bars.
            expected_name = f"{prefix}_{distribution}_Expected"
            drop_local_name(sheet, expected_name)
            count_ref = f"'{sname}'!${col_letter(_C_E)}${_ROW_STAT_COUNT}"
            nm = sheet.api.Names.Add(
                Name=expected_name,
                RefersTo=(
                    f"=OFFSET('{sname}'!${col_ltr}${_ROW_COL_HDRS},"
                    f"1,0,MAX(IFERROR({size_formula},1),1),1)*{count_ref}"
                ),
            )
            nm.Comment = (
                f"{method_labels[prefix]} Method histogram chart: {distribution} "
                "overlay line (expected counts = bin probability × n)"
            )

    # Q-Q chart column ranges: OFFSET-based, anchored at the header row and
    # sized by the Count stat cell (one point per included observation).
    qq_size = f"MAX(IFERROR('{sname}'!${col_letter(_C_E)}${_ROW_STAT_COUNT},1),1)"
    for offset, _, suffix, distribution in _QQ_COLUMNS:
        if suffix == "P":
            continue   # plotting positions are an intermediate, never charted
        name = f"UV_QQ_{suffix}"
        col_ltr = col_letter(_C_QQ + offset)
        drop_local_name(sheet, name)
        nm = sheet.api.Names.Add(
            Name=name,
            RefersTo=(
                f"=OFFSET('{sname}'!${col_ltr}${_ROW_COL_HDRS},1,0,{qq_size},1)"
            ),
        )
        nm.Comment = (
            f"{distribution} Q-Q plot: theoretical quantiles (X axis)"
            if distribution
            else "Q-Q plots: sorted sample values (Y axis, shared by all eight charts)"
        )


def _dist_fit_rows_by_name() -> dict[str, int]:
    return {name: row for row, name, *_ in _dist_rows(_ROW_DIST_START)}


def _cdf_column_formula(
    edge_spill_ref: str,
    lower_spill_ref: str,
    distribution: str,
) -> str:
    """Build a single-column spill formula for one histogram CDF probability.

    Both edge_spill_ref (upper edges) and lower_spill_ref (lower edges) are
    already computed on the sheet — each is a spill reference to the relevant
    histogram column.  The CDF formula simply references them, so no FILTER or
    Bin_Lower_Edges computation is repeated across the 8 CDF columns in a block.

    Beta additionally references the already-computed Min ($E$9) and Range ($E$11)
    stat cells for the data-range rescaling it requires.
    """
    t1 = col_letter(_C_T1_VAL)  # I — θ₁ value
    t2 = col_letter(_C_T2_VAL)  # K — θ₂ value
    t3 = col_letter(_C_T3_VAL)  # M — θ₃ value

    fit_rows = _dist_fit_rows_by_name()

    expressions = {
        "Normal": (
            f"CDF_Normal(edges,${t1}${fit_rows['Normal']},"
            f"${t2}${fit_rows['Normal']},lower)"
        ),
        "Lognormal": (
            f"CDF_Lognormal(edges,${t1}${fit_rows['Lognormal']},"
            f"${t2}${fit_rows['Lognormal']},lower)"
        ),
        "Exponential": (
            f"CDF_Exponential(edges,${t1}${fit_rows['Exponential']},lower)"
        ),
        "Weibull": (
            f"CDF_Weibull(edges,${t1}${fit_rows['Weibull']},"
            f"${t2}${fit_rows['Weibull']},lower)"
        ),
        "Gamma": (
            f"CDF_Gamma(edges,${t1}${fit_rows['Gamma']},"
            f"${t2}${fit_rows['Gamma']},lower)"
        ),
        "Triangular": (
            f"CDF_Triangular(edges,${t1}${fit_rows['Triangular']},"
            f"${t2}${fit_rows['Triangular']},${t3}${fit_rows['Triangular']},lower)"
        ),
        "Beta": (
            f"CDF_Beta(edges,${t1}${fit_rows['Beta']},"
            f"${t2}${fit_rows['Beta']},dmin,drange,lower)"
        ),
        "BetaPERT": (
            f"CDF_BetaPERT(edges,${t1}${fit_rows['BetaPERT']},"
            f"${t2}${fit_rows['BetaPERT']},${t3}${fit_rows['BetaPERT']},lower)"
        ),
    }

    # Beta needs dmin/drange for data-range rescaling.  Rather than computing
    # FILTER(UV_Data,UV_Include) inline again, reference the Descriptive Statistics
    # cells (Min and Range) that were already computed in the stats zone.
    if distribution == "Beta":
        stat_col = col_letter(_C_E)
        min_ref   = f"${stat_col}${_ROW_STAT_MIN}"    # e.g. $E$9
        range_ref = f"${stat_col}${_ROW_STAT_RANGE}"  # e.g. $E$11
        beta_vars = f"dmin,{min_ref},drange,{range_ref},"
    else:
        beta_vars = ""

    return (
        "=LET("
        f"edges,{edge_spill_ref},"
        f"lower,{lower_spill_ref},"
        f"{beta_vars}"
        f"{expressions[distribution]}"
        ")"
    )


# ── Zone 1: data input column ─────────────────────────────────────────────────

def _write_data_zone(sheet: xw.Sheet) -> None:
    section_heading(sheet, _ROW_SECTION_HDR, _C_A, "Data")
    # Override section-heading color: A3 and the A4:B(_DATA_END) input block use INPUT_COLOR
    # to signal that this is where users paste their own dataset.
    sheet.range(rc(_ROW_SECTION_HDR, _C_A)).color = _INPUT
    sheet.range(rc(_ROW_DATA_START, _C_A), rc(_DATA_END, _C_B)).color = _INPUT

    val(sheet, _ROW_COL_HDRS, _C_A, "Life expectancy")
    f(
        sheet,
        _ROW_DATA_START,
        _C_A,
        '=IF(LifeExpectancyData[Life expectancy]="","",LifeExpectancyData[Life expectancy])',
    )
    sheet.range(rc(_ROW_DATA_START, _C_A), rc(_DATA_END, _C_A)).number_format = _FMT_1DP

    # Filter column — local numeric mask; UV_Include is defined as $B$4#
    section_heading(sheet, _ROW_SECTION_HDR, _C_B, "Filter")
    val(sheet, _ROW_COL_HDRS, _C_B, "Include")
    f(sheet, _ROW_DATA_START, _C_B, "=ISNUMBER(LifeExpectancyData[Life expectancy])")
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

# Row positions of stat cells referenced by CDF/quantile formulas and chart
# named ranges (avoid repeating FILTER calls)
_ROW_STAT_MIN   = _ROW_STATS_START + next(i for i, (n, _) in enumerate(_STAT_ROWS) if n == "Min")
_ROW_STAT_RANGE = _ROW_STATS_START + next(i for i, (n, _) in enumerate(_STAT_ROWS) if n == "Range")
_ROW_STAT_COUNT = _ROW_STATS_START + next(i for i, (n, _) in enumerate(_STAT_ROWS) if n == "Count")

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


# ── Zone 4: histogram tables (11-col blocks with CDF probabilities) ──────────

def _write_histogram_table(
    sheet: xw.Sheet,
    col_start: int,
    method: str,
) -> None:
    """Write one 11-column histogram block (lower edges, upper edges, counts, 8 CDF probabilities)."""
    col_lower = col_start + _HB_LOWER
    col_edge = col_start + _HB_EDGE
    col_count = col_start + _HB_COUNT
    col_cdf_first = col_start + _HB_NORM
    col_last = col_start + _HIST_W - 1

    # Row 2: method label and value
    section_heading(sheet, _ROW_METHOD_HDR, col_edge, "Method")
    val(sheet, _ROW_METHOD_HDR, col_count, method)
    sheet.range(rc(_ROW_METHOD_HDR, col_count)).color = _HEADER
    sheet.range(rc(_ROW_METHOD_HDR, col_count)).api.Font.Bold = True

    # Row 4: column headers
    for offset, header, _, _ in _HIST_COLUMNS:
        val(sheet, _ROW_COL_HDRS, col_start + offset, header)
    _subheader_row(sheet, _ROW_COL_HDRS, col_lower, col_last)

    # Row 3: bin-count display
    val(sheet, _ROW_SECTION_HDR, col_edge, "Bins:")
    method_cell = a1(_ROW_METHOD_HDR, col_count)
    f(sheet, _ROW_SECTION_HDR, col_count,
      f"=Number_Of_Histogram_Bins(UV_Data,{method_cell},UV_Include)")
    sheet.range(rc(_ROW_SECTION_HDR, col_count)).number_format = _FMT_INT

    # Upper-edge, count, and lower-edge spill formulas.
    # Each function takes (data, method, filter) and derives its own edges from
    # Bin_Edges internally. The resulting spill ranges are still referenced by
    # the CDF columns below to avoid recomputing edges eight times per block.
    f(sheet, _ROW_HIST_START, col_edge, f"=Upper_Bin_Edges(UV_Data,{method_cell},UV_Include)")
    edge_spill_ref = f"{col_letter(col_edge)}{_ROW_HIST_START}#"
    f(sheet, _ROW_HIST_START, col_count,
      f"=Bin_Counts(UV_Data,{method_cell},UV_Include)")
    f(sheet, _ROW_HIST_START, col_lower,
      f"=Bin_Lower_Edges(UV_Data,{method_cell},UV_Include)")
    lower_spill_ref = f"{col_letter(col_lower)}{_ROW_HIST_START}#"

    # CDF probability columns — one spill formula per column.
    for offset, _, _, distribution in _HIST_COLUMNS:
        if not distribution:
            continue
        f(
            sheet,
            _ROW_HIST_START,
            col_start + offset,
            _cdf_column_formula(edge_spill_ref, lower_spill_ref, distribution),
        )

    # Number formats
    sheet.range(rc(_ROW_HIST_START, col_lower), rc(_DATA_END, col_lower)).number_format = _FMT_1DP
    sheet.range(rc(_ROW_HIST_START, col_edge), rc(_DATA_END, col_edge)).number_format = _FMT_1DP
    sheet.range(rc(_ROW_HIST_START, col_count), rc(_DATA_END, col_count)).number_format = _FMT_INT
    sheet.range(
        rc(_ROW_HIST_START, col_cdf_first), rc(_DATA_END, col_last)
    ).number_format = _FMT_4DP


def _write_histograms(sheet: xw.Sheet) -> None:
    # Zone super-heading in title row, merged across all three histogram blocks
    section_heading(sheet, _ROW_TITLE, _C_STUR, "Histograms")
    sheet.range(
        rc(_ROW_TITLE, _C_STUR), rc(_ROW_TITLE, _C_FD + _HIST_W - 1)
    ).merge()

    _write_histogram_table(sheet, _C_STUR, "Sturges")
    _write_histogram_table(sheet, _C_SCOTT, "Scott")
    _write_histogram_table(sheet, _C_FD, "FD")


# ── Zone 5: Q-Q plot data ─────────────────────────────────────────────────────

def _qq_column_formula(p_spill_ref: str, distribution: str) -> str:
    """Build a single-column spill formula of theoretical quantiles.

    Evaluates the fitted distribution's inverse CDF at the plotting positions
    in p_spill_ref, referencing the same fit-table parameter cells the CDF
    columns use.  Native inverse functions (NORM.INV, LOGNORM.INV, GAMMA.INV,
    BETA.INV) where Excel has them; closed-form inverses for Exponential,
    Weibull, and Triangular.

    Beta inverts the same 0.1 %-padded data-range rescaling CDF_Beta applies,
    reading the Min ($E$9) and Range ($E$11) stat cells; BetaPERT applies the
    same PERT reparameterization as its A-D/K-S formulas before BETA.INV.
    """
    t1 = col_letter(_C_T1_VAL)  # I — θ₁ value
    t2 = col_letter(_C_T2_VAL)  # K — θ₂ value
    t3 = col_letter(_C_T3_VAL)  # M — θ₃ value

    fit_rows = _dist_fit_rows_by_name()
    r = fit_rows[distribution]

    stat_col = col_letter(_C_E)
    min_ref   = f"${stat_col}${_ROW_STAT_MIN}"    # e.g. $E$9
    range_ref = f"${stat_col}${_ROW_STAT_RANGE}"  # e.g. $E$11

    # (extra LET variables, quantile expression) per distribution
    expressions = {
        "Normal": ("", f"NORM.INV(p_,${t1}${r},${t2}${r})"),
        "Lognormal": ("", f"LOGNORM.INV(p_,${t1}${r},${t2}${r})"),
        "Exponential": ("", f"-LN(1-p_)/${t1}${r}"),
        "Weibull": ("", f"${t2}${r}*(-LN(1-p_))^(1/${t1}${r})"),
        "Gamma": ("", f"GAMMA.INV(p_,${t1}${r},1/${t2}${r})"),
        "Triangular": (
            f"mn,${t1}${r},md,${t2}${r},mx,${t3}${r},",
            "IF(p_<(md-mn)/(mx-mn),"
            "mn+SQRT(p_*(mx-mn)*(md-mn)),"
            "mx-SQRT((1-p_)*(mx-mn)*(mx-md)))",
        ),
        "Beta": (
            f"mn,{min_ref},range_,{range_ref},pad,range_*0.001,scale_,range_+2*pad,",
            f"BETA.INV(p_,${t1}${r},${t2}${r})*scale_+mn-pad",
        ),
        # λ=4 PERT mapping — algebraically identical to the μ-based
        # reparameterization but with no removable 0/0 singularity at a
        # symmetric mode (md = (mn+mx)/2), where the μ form degenerates to
        # α = β = 0 and BETA.INV returns #NUM!.
        "BetaPERT": (
            f"mn,${t1}${r},md,${t2}${r},mx,${t3}${r},"
            "alpha_param,1+4*(md-mn)/(mx-mn+1E-30),"
            "beta_param,1+4*(mx-md)/(mx-mn+1E-30),",
            "BETA.INV(p_,alpha_param,beta_param)*(mx-mn)+mn",
        ),
    }

    extra_vars, expr = expressions[distribution]
    return f"=LET(p_,{p_spill_ref},{extra_vars}{expr})"


def _write_qq_data(sheet: xw.Sheet) -> None:
    """Write plotting positions, sorted sample, and 8 theoretical-quantile columns."""
    col_p = _C_QQ + _QQ_P
    col_sample = _C_QQ + _QQ_SAMPLE
    col_last = _C_QQ + _QQ_W - 1

    section_heading(sheet, _ROW_TITLE, _C_QQ, "Q-Q Plot Data")
    sheet.range(rc(_ROW_TITLE, _C_QQ), rc(_ROW_TITLE, col_last)).merge()

    for offset, header, _, _ in _QQ_COLUMNS:
        val(sheet, _ROW_COL_HDRS, _C_QQ + offset, header)
    _subheader_row(sheet, _ROW_COL_HDRS, _C_QQ, col_last)

    # Hazen plotting positions (i - 0.5)/n — the convention QQ_Correlation and
    # Normal_Scores already use in the regression Q-Q machinery.  Guarded for
    # the empty-data case: SEQUENCE(0) would spill #CALC! and cascade into
    # every quantile column; NA() instead makes the charts skip the points.
    f(sheet, _ROW_HIST_START, col_p,
      "=LET(n_,UV_n,IF(n_<=0,NA(),(SEQUENCE(n_)-0.5)/n_))")
    p_spill_ref = f"${col_letter(col_p)}${_ROW_HIST_START}#"

    f(sheet, _ROW_HIST_START, col_sample, "=SORT(FILTER(UV_Data,UV_Include))")

    for offset, _, _, distribution in _QQ_COLUMNS:
        if not distribution:
            continue
        f(
            sheet,
            _ROW_HIST_START,
            _C_QQ + offset,
            _qq_column_formula(p_spill_ref, distribution),
        )

    # Spills run one row past _DATA_END when every data row is included
    # (2000 values from row 5), so format through _DATA_END + 1.
    last_spill_row = _DATA_END + 1
    sheet.range(rc(_ROW_HIST_START, col_p), rc(last_spill_row, col_p)).number_format = _FMT_4DP
    sheet.range(
        rc(_ROW_HIST_START, col_sample), rc(last_spill_row, col_last)
    ).number_format = _FMT_1DP


# ── Zone 3: distribution fitting summary table ────────────────────────────────

_FIT_COL_HDRS = [
    (_C_DIST_NAME, "Distribution"),
    (_C_T1_LBL, "θ₁"),
    (_C_T1_VAL, "Value"),
    (_C_T2_LBL, "θ₂"),
    (_C_T2_VAL, "Value"),
    (_C_T3_LBL, "θ₃"),
    (_C_T3_VAL, "Value"),
    (_C_NLL, "NLL"),
    (_C_K_PARAM, "k"),
    (_C_AIC, "AIC"),
    (_C_BIC, "BIC"),
    (_C_AD, "A-D"),
    (_C_KS, "K-S"),
]

_FIT_NUMBER_FORMATS: dict[int, str] = {
    _C_T1_VAL: _FMT_1DP,
    _C_T2_VAL: _FMT_1DP,
    _C_T3_VAL: _FMT_1DP,
    _C_NLL: _FMT_SCI_1DP,
    _C_K_PARAM: _FMT_INT,
    _C_AIC: _FMT_1DP,
    _C_BIC: _FMT_1DP,
    _C_AD: _FMT_4DP,
    _C_KS: _FMT_4DP,
}


# Where each fit's Stage 2 control block sits: (zone first column, Stage 2 row).
# This is the one place the fitting table and the search writers agree on, so
# the table can never reference a block the writers did not put there.  Both
# halves are per-distribution now: with one column zone per fit, the Best column
# differs across fits, and Beta's Stage 2 is 26 rows down where the profile
# fits' is 5.
_STAGE2_ANCHORS: dict[str, tuple[int, int]] = {
    "Weibull": (_C_GS_WB, _PS_R_STAGE_STRIDE),
    "Gamma": (_C_GS_GAMMA, _PS_R_STAGE_STRIDE),
    "Beta": (_C_GS_BETA, _GS_R_STAGE2),
}

# Every fit zone's row anchor. The zones sit side by side, so they all start
# on the same row and differ only in column.
_ROW_FIT_ZONE = 1


def _final_grid_best_refs(distribution: str) -> tuple[str, str]:
    """Return the Stage 2 Best cells for a fit, as fitting-table formulas.

    Reads ``_STAGE2_ANCHORS`` so the reference follows the zone the search
    writer actually produced.  The two rows within a stage are shared by both
    stage writers (``_GS_R_P1`` / ``_GS_R_P2``); the column and the stage row
    are per-distribution.
    """
    zone_col, stage2_row = _STAGE2_ANCHORS[distribution]
    best_col = col_letter(zone_col + _GS_C_BEST)
    title_row = _ROW_FIT_ZONE + stage2_row
    return (
        f"=${best_col}${title_row + _GS_R_P1}",
        f"=${best_col}${title_row + _GS_R_P2}",
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
        return f"${col_letter(_C_T1_VAL)}${row}"

    def _t(row: int) -> str:
        return f"${col_letter(_C_T2_VAL)}${row}"

    def _v(row: int) -> str:
        return f"${col_letter(_C_T3_VAL)}${row}"

    def _n(row: int) -> str:
        return f"${col_letter(_C_K_PARAM)}${row}"

    weibull_shape_ref, weibull_scale_ref = _final_grid_best_refs("Weibull")
    gamma_shape_ref, gamma_rate_ref = _final_grid_best_refs("Gamma")
    beta_alpha_ref, beta_beta_ref = _final_grid_best_refs("Beta")

    rows: list[tuple] = []
    dist_specs = [
        (
            "Normal",
            "Mean",    "=AVERAGE(FILTER(UV_Data,UV_Include))",
            "Std Dev", "=STDEV.S(FILTER(UV_Data,UV_Include))",
            "",        "",
            lambda r: f"=NLL_Normal(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
            lambda r: f"NORM.DIST(UV_Data,{_r(r)},{_t(r)},TRUE)",
        ),
        (
            "Lognormal",
            "μ_ln",  "=AVERAGE(LN(FILTER(UV_Data,(UV_Include)*(ISNUMBER(UV_Data)))))",
            "σ_ln",  "=STDEV.S(LN(FILTER(UV_Data,(UV_Include)*(ISNUMBER(UV_Data)))))",
            "",      "",
            lambda r: f"=NLL_Lognormal(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
            lambda r: f"LOGNORM.DIST(UV_Data,{_r(r)},{_t(r)},TRUE)",
        ),
        (
            "Exponential",
            "Rate",  "=1/AVERAGE(FILTER(UV_Data,UV_Include))",
            "",      "",
            "",      "",
            lambda r: f"=NLL_Exponential(UV_Data,{_r(r)},UV_Include)",
            1,
            lambda r: f"EXPON.DIST(UV_Data,{_r(r)},TRUE)",
        ),
        (
            # Params come from Stage 2 grid-search MLE (see Zone 6)
            "Weibull",
            "shape",
            weibull_shape_ref,
            "scale",
            weibull_scale_ref,
            "",    "",
            lambda r: f"=NLL_Weibull(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
            lambda r: f"WEIBULL.DIST(UV_Data,{_r(r)},{_t(r)},TRUE)",
        ),
        (
            # Params come from Stage 2 grid-search MLE (see Zone 6)
            "Gamma",
            "shape",
            gamma_shape_ref,
            "rate",
            gamma_rate_ref,
            "",    "",
            lambda r: f"=NLL_Gamma(UV_Data,{_r(r)},{_t(r)},UV_Include)",
            2,
            lambda r: f"GAMMA.DIST(UV_Data,{_r(r)},1/{_t(r)},TRUE)",
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
            lambda r: (
                f"IFERROR(IF(UV_Data<{_r(r)},0,"
                f"IF(UV_Data<{_t(r)},"
                f"(UV_Data-{_r(r)})^2/(({_v(r)}-{_r(r)})*({_t(r)}-{_r(r)})+1E-30),"
                f"IF(UV_Data<={_v(r)},"
                f"1-(({_v(r)}-UV_Data)^2/(({_v(r)}-{_r(r)})*({_v(r)}-{_t(r)})+1E-30)),1))),0.5)"
            ),
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
            lambda r: (
                f"LET(d,UV_Data,mn,MIN(FILTER(d,UV_Include)),"
                f"range_,MAX(FILTER(d,UV_Include))-mn,"
                f"pad,MAX(range_*0.001,1E-30),scale_,range_+2*pad,"
                f"z,(d-mn+pad)/scale_,"
                f"IFERROR(BETA.DIST(z,{_r(r)},{_t(r)},TRUE),0.5))"
            ),
        ),
        (
            "BetaPERT",
            "Min",  "=LET(d,FILTER(UV_Data,UV_Include),MIN(d)-(MAX(d)-MIN(d))*0.001)",
            "Mode", "=IFERROR(MODE.SNGL(FILTER(UV_Data,UV_Include)),MEDIAN(FILTER(UV_Data,UV_Include)))",
            "Max",  "=LET(d,FILTER(UV_Data,UV_Include),MAX(d)+(MAX(d)-MIN(d))*0.001)",
            lambda r: f"=NLL_BetaPERT(UV_Data,{_r(r)},{_t(r)},{_v(r)},UV_Include)",
            3,
            # λ=4 PERT mapping (same as NLL_BetaPERT) — the μ-based form is
            # 0/0 at a symmetric mode and degenerates to α = β = 0.
            lambda r: (
                f"LET(mn,{_r(r)},md,{_t(r)},mx,{_v(r)},"
                f"alpha_param,1+4*(md-mn)/(mx-mn+1E-30),"
                f"beta_param,1+4*(mx-md)/(mx-mn+1E-30),"
                f"BETA.DIST((UV_Data-mn)/(mx-mn+1E-30),alpha_param,beta_param,TRUE))"
            ),
        ),
    ]

    for i, (name, l1, f1, l2, f2, l3, f3, nll_fn, k, cdf_fn) in enumerate(dist_specs):
        row = base_row + i
        rows.append((row, name, l1, f1, l2, f2, l3, f3, nll_fn(row), k, cdf_fn(row)))

    return rows


def _write_fitting_table(sheet: xw.Sheet) -> None:
    # Zone heading at row 2, merged across all fitting columns
    section_heading(sheet, _ROW_METHOD_HDR, _C_FIT_FIRST, "Distribution Fitting/Comparison")
    sheet.range(rc(_ROW_METHOD_HDR, _C_FIT_FIRST), rc(_ROW_METHOD_HDR, _C_FIT_LAST)).merge()

    for col, label in _FIT_COL_HDRS:
        val(sheet, _ROW_COL_HDRS, col, label)
    _subheader_row(sheet, _ROW_COL_HDRS, _C_FIT_FIRST, _C_FIT_LAST)

    dist_data = _dist_rows(_ROW_DIST_START)
    for row, name, l1, f1, l2, f2, l3, f3, nll_f, k, cdf_expr in dist_data:
        val(sheet, row, _C_DIST_NAME, name)
        val(sheet, row, _C_T1_LBL, l1)
        if f1:
            f(sheet, row, _C_T1_VAL, f1)
        val(sheet, row, _C_T2_LBL, l2)
        if f2:
            f(sheet, row, _C_T2_VAL, f2)
        val(sheet, row, _C_T3_LBL, l3)
        if f3:
            f(sheet, row, _C_T3_VAL, f3)
        f(sheet, row, _C_NLL, nll_f)
        val(sheet, row, _C_K_PARAM, k)
        f(sheet, row, _C_AIC,
          f"=GoF_AIC(${col_letter(_C_NLL)}${row},${col_letter(_C_K_PARAM)}${row})")
        f(sheet, row, _C_BIC,
          f"=GoF_BIC(${col_letter(_C_NLL)}${row},${col_letter(_C_K_PARAM)}${row},UV_n)")
        f(sheet, row, _C_AD,
          f"=GoF_Anderson_Darling(UV_Data,{cdf_expr},UV_Include)")
        f(sheet, row, _C_KS,
          f"=GoF_Kolmogorov_Smirnov(UV_Data,{cdf_expr},UV_Include)")

        for col, fmt in _FIT_NUMBER_FORMATS.items():
            sheet.range(rc(row, col)).number_format = fmt

    # Border around the table (col headers through last data row)
    last_row = _ROW_DIST_START + len(dist_data) - 1
    border_box(sheet, _ROW_COL_HDRS, _C_FIT_FIRST, last_row, _C_FIT_LAST)

    # Highlight best-fit row (lowest AIC) with conditional formatting
    aic_col_letter = col_letter(_C_AIC)
    aic_min_range = f"${aic_col_letter}${_ROW_DIST_START}:${aic_col_letter}${last_row}"
    formula = f"=${aic_col_letter}{_ROW_DIST_START}=MIN({aic_min_range})"
    row_range = sheet.range(rc(_ROW_DIST_START, _C_FIT_FIRST), rc(last_row, _C_FIT_LAST))
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
    prefix: str,
) -> None:
    """Insert one histogram combo chart: gapless count bars plus one smoothed,
    markerless overlay line per fitted distribution (expected counts)."""
    sname = sheet.name
    co = sheet.api.ChartObjects().Add(chart_left, chart_top, chart_width, chart_height)
    chart = co.Chart

    while chart.SeriesCollection().Count > 0:
        chart.SeriesCollection(1).Delete()

    chart.ChartType = _XL_COLUMN_CLUSTERED
    chart.ChartGroups(1).GapWidth = 0

    edges_ref = f"='{sname}'!{prefix}_Upper_Edges"
    series = chart.SeriesCollection().NewSeries()
    series.XValues = edges_ref
    series.Values  = f"='{sname}'!{prefix}_Counts"
    series.Name    = "Count"

    # Fitted-distribution overlays sourced from the *_Expected named formulas
    # (per-bin probability mass × n), so they share the bars' count axis.
    for _, _, _, distribution in _HIST_COLUMNS:
        if not distribution:
            continue
        line = chart.SeriesCollection().NewSeries()
        line.ChartType = _XL_LINE
        line.XValues = edges_ref
        line.Values = f"='{sname}'!{prefix}_{distribution}_Expected"
        line.Name = distribution
        line.Smooth = True
        line.MarkerStyle = _XL_MARKER_NONE

    chart.HasLegend = True
    chart.Legend.Position = _XL_LEGEND_BOTTOM
    chart.HasTitle = True
    title_row = {
        "UV_Sturges": _ROW_CHART1_TITLE,
        "UV_Scott": _ROW_CHART2_TITLE,
        "UV_FD": _ROW_CHART3_TITLE,
    }[prefix]
    chart.ChartTitle.Formula = f"='{sname}'!${col_letter(_C_FIT_FIRST)}${title_row}"
    x_axis = chart.Axes(_XL_CATEGORY)
    x_axis.HasTitle = True
    x_axis.AxisTitle.Text = "Upper Edges"
    y_axis = chart.Axes(_XL_VALUE)
    y_axis.HasTitle = True
    y_axis.AxisTitle.Text = "Count"


def _write_histogram_chart_title_cells(sheet: xw.Sheet) -> None:
    """Write chart title formula cells at G14, G34, G54 for the three histograms."""
    for row, block_start in [
        (_ROW_CHART1_TITLE, _C_STUR),
        (_ROW_CHART2_TITLE, _C_SCOTT),
        (_ROW_CHART3_TITLE, _C_FD),
    ]:
        count_col = block_start + _HB_COUNT
        f(sheet, row, _C_FIT_FIRST,
          f'={col_letter(count_col)}{_ROW_METHOD_HDR}&" Method Histogram"')


def _write_histogram_charts(sheet: xw.Sheet) -> None:
    """Insert three histogram combo charts for the Sturges, Scott, and FD tables."""
    for prefix, row_start, row_end in [
        ("UV_Sturges", _ROW_CHART1_TITLE, _ROW_CHART1_TITLE + 19),
        ("UV_Scott",   _ROW_CHART2_TITLE, _ROW_CHART2_TITLE + 19),
        ("UV_FD",      _ROW_CHART3_TITLE, _ROW_CHART3_TITLE + 19),
    ]:
        chart_range = sheet.range(rc(row_start, _C_FIT_FIRST), rc(row_end, _C_FIT_LAST))
        _add_histogram_chart(
            sheet,
            chart_left=chart_range.left,
            chart_top=chart_range.top,
            chart_width=chart_range.width,
            chart_height=chart_range.height,
            prefix=prefix,
        )


def _set_equal_qq_axis_scale(sheet: xw.Sheet, x_axis, y_axis, x_name: str) -> None:
    """Set equal min/max scales on both axes so the identity series reads as 45°.

    Same pattern as the Regression sheet's Normal Q-Q chart. Guarded per chart:
    the two-parameter quantile columns depend on Data Table grid-search output,
    which may not evaluate during a manual-calculation build — in that case the
    axes are left auto-scaling rather than failing the remaining charts.
    """
    sname = sheet.name
    try:
        common_min = float(
            sheet.api.Evaluate(f"=MIN('{sname}'!{x_name},'{sname}'!UV_QQ_Sample)")
        )
        common_max = float(
            sheet.api.Evaluate(f"=MAX('{sname}'!{x_name},'{sname}'!UV_QQ_Sample)")
        )
    except Exception:
        return

    if common_max <= common_min:
        return

    x_axis.MinimumScale = common_min
    x_axis.MaximumScale = common_max
    y_axis.MinimumScale = common_min
    y_axis.MaximumScale = common_max


def _write_qq_charts(sheet: xw.Sheet) -> None:
    """Insert eight Q-Q scatter charts in a 4×2 grid below the histogram charts.

    Charts fill row-major in fit-table order: Normal | Lognormal on the first
    row, Exponential | Weibull on the second, and so on — left charts span
    G:O and right charts P:T (_QQ_CHART_BANDS).
    """
    sname = sheet.name
    distributions = [d for _, _, _, d in _QQ_COLUMNS if d]
    for i, distribution in enumerate(distributions):
        grid_row, grid_col = divmod(i, len(_QQ_CHART_BANDS))
        col_first, col_last = _QQ_CHART_BANDS[grid_col]
        row_start = _ROW_QQ_CHART_START + grid_row * _QQ_CHART_ROWS
        chart_range = sheet.range(
            rc(row_start, col_first),
            rc(row_start + _QQ_CHART_ROWS - 1, col_last),
        )
        co = sheet.api.ChartObjects().Add(
            chart_range.left, chart_range.top, chart_range.width, chart_range.height
        )
        chart = co.Chart

        while chart.SeriesCollection().Count > 0:
            chart.SeriesCollection(1).Delete()

        chart.ChartType = _XL_XY_SCATTER

        theoretical_ref = f"='{sname}'!UV_QQ_{distribution}"
        series = chart.SeriesCollection().NewSeries()
        series.XValues = theoretical_ref
        series.Values = f"='{sname}'!UV_QQ_Sample"
        series.Name = distribution
        series.MarkerSize = 4

        # y = x reference as a real data series (never a drawn shape): pointing
        # both axes at the theoretical-quantile range guarantees every point
        # sits on the identity line across the chart's X extent.
        identity = chart.SeriesCollection().NewSeries()
        identity.XValues = theoretical_ref
        identity.Values = theoretical_ref
        identity.Name = "Identity"
        identity.ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS
        identity.Format.Line.ForeColor.RGB = excel_color((120, 120, 120))
        identity.Format.Line.DashStyle = 3  # msoLineRoundDot
        identity.Format.Line.Weight = 1.25

        chart.HasLegend = False
        chart.HasTitle = True
        chart.ChartTitle.Text = f"{distribution} Q-Q Plot"
        x_axis = chart.Axes(_XL_CATEGORY)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = "Theoretical Quantiles"
        y_axis = chart.Axes(_XL_VALUE)
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = "Sample Quantiles"

        _set_equal_qq_axis_scale(sheet, x_axis, y_axis, f"UV_QQ_{distribution}")




# ── Zone 6: two-parameter grid-search MLE ────────────────────────────────────
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
#   Min NLL table       — col+0, rows+1…+2; value uses TAKE(Grid_Argument_Minimum(...),,1)
#   Rows/Columns table  — col+1, rows+1…+2; generated value documents physical grid size
#   Parameter table     — cols+3…+8, rows+1…+3
#   Blank spacer column — col+2, rows+1…+3
#   Best column         — Grid_Search_Optimum(...) spills from column parameter to row parameter
#
# The visible Input cells are the actual RowInput and ColumnInput cells supplied
# to Excel's two-input Data Table object.  No hidden auxiliary row is required.
#
# Beta is the only fit using this writer.  Both its stages share the Beta zone
# (_C_GS_BETA) and stack vertically, Stage 2 _GS_R_STAGE2 rows below Stage 1:
#   Stage 1: row 1,  body = CK6:DD25
#   Stage 2: row 27, body = CK32:DD51
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

    The body is a two-input Excel Data Table wired via ``Range.Table`` — the
    corner NLL formula is substituted across the row and column axes.  Beta is
    the only distribution that still needs a two-dimensional search (both of
    its conditional MLEs involve digamma, so neither parameter profiles out in
    closed form); Weibull and Gamma go through ``_write_profile_stage``.
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
    _drop_wb_name(sheet, body_name)
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
    # Min NLL is the first column returned by Grid_Argument_Minimum.
    f(
        sheet,
        p1_row,
        c0 + _GS_C_MINNLL,
        f'=IFERROR(TAKE(Grid_Argument_Minimum({body_name}),,1),"—")',
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
        location = f"INDEX(Grid_Argument_Minimum({body_name}),1,{argmin_col})"
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
    col_start: int,
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
    """Write Stage 1 and Stage 2 grid-search blocks for one distribution.

    Both stages share this fit's own column zone and stack vertically, Stage 2
    ``_GS_R_STAGE2`` rows below Stage 1 — a full grid block plus one gap row.
    """
    r0 = _ROW_FIT_ZONE
    s1 = _write_grid_stage(
        sheet,
        row_start   = r0,
        col_start   = col_start,
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
        row_start   = r0 + _GS_R_STAGE2,
        col_start   = col_start,
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


# ── One-dimensional profile-NLL search ───────────────────────────────────────
#
# A fit's whole zone (col_start = c0, rows from _ROW_FIT_ZONE = r0).  Both
# stages' control blocks stack in the same 9 columns; both bodies share rows
# and sit side by side beneath, each under one half of the control block.
#
#   r0+0  : Stage 1 heading merged across c0…c0+_PS_W-1
#   r0+1  : Min NLL and Grid Points labels; blank c0+2; parameter headers
#   r0+2  : control values; searched parameter: Parameter | Start | Min | Max |
#           Step Size | Best
#   r0+3  : profiled-out parameter: Parameter | (no bounds) | Best
#   r0+5…+8 : Stage 2, the same four rows _PS_R_STAGE_STRIDE lower
#   r0+10 : body headers — Stage 1 axis/NLL at c0+0/c0+1, Stage 2 at c0+3/c0+4
#   r0+11…+10+N : both stages' axis SEQUENCE and profile NLL, side by side
#
# Splitting the control anchor from the body anchor is what allows this: the
# stage writer takes `row_start`/`col_start` for its controls and a separate
# `body_row_start`/`body_cols` for its curve.
#
# The searched parameter's partner is eliminated in closed form, so each body
# cell is a single NLL call at (trial value, partner(trial value)) — the same
# catalog NLL LAMBDA the fitting table reports, which is what guarantees the
# curve's minimum and the table's NLL agree exactly.
#
# Named ranges registered here:
#   *_S1 / *_S2        = stage profile-NLL body (N×1)
#   *_S1_Axis / *_S2_Axis = stage searched-parameter axis (N×1)


# The included sample, as an expression.  Body cells bind it to a LET name
# once and pass that name down; the standalone `Best` cells, which evaluate the
# partner a single time, inline it.
_FILTERED_DATA = "FILTER(UV_Data,UV_Include)"


def _weibull_profile_scale(shape_ref: str, data: str = _FILTERED_DATA) -> str:
    """Closed-form Weibull scale at a given shape: λ̂(k) = ((1/n)·Σxᵢᵏ)^(1/k)."""
    return f"((AVERAGE({data}^{shape_ref}))^(1/{shape_ref}))"


def _gamma_profile_rate(shape_ref: str, data: str = _FILTERED_DATA) -> str:
    """Closed-form Gamma rate at a given shape: β̂(α) = α / x̄."""
    return f"({shape_ref}/AVERAGE({data}))"


# Closed-form starting values for the searched parameter.  Both are plug-in
# estimators accurate to a few percent, which is what makes a 20-point bracket
# enough; each falls back to a neutral constant if the sample cannot support it.
# The plotting positions are built the way Zone 5's `P` column builds them —
# SORT once, then Hazen (i-0.5)/n — rather than through `Rank_Fraction`.  Two
# reasons: `Rank_Fraction` returns i/n, which is exactly 1 at the sample
# maximum and makes LN(-LN(1-p)) an error on every sample; and it is O(n²)
# (BYROW over SUMPRODUCT), where the sorted form is O(n log n).
_WEIBULL_SHAPE_START = (
    "=IFERROR(LET("
    "x,SORT(FILTER(UV_Data,UV_Include)),"
    "n,ROWS(x),"
    "p,(SEQUENCE(n)-0.5)/n,"
    "SLOPE(LN(-LN(1-p)),LN(x))"
    "),2)"
)

_GAMMA_SHAPE_START = (
    "=IFERROR(LET("
    "x,FILTER(UV_Data,UV_Include),"
    "s,LN(AVERAGE(x))-AVERAGE(LN(x)),"
    "IF(s<=0,1,(3-s+SQRT((s-3)^2+24*s))/(12*s))"
    "),1)"
)


def _write_profile_stage(
    sheet: xw.Sheet,
    row_start: int,
    col_start: int,
    title: str,
    body_name: str,
    p1_label: str,
    p2_label: str,
    nll_formula,
    partner_formula,
    p1_start,   # Excel formula string for the searched parameter's start value
    p1_min,     # Excel formula string
    p1_max,     # Excel formula string
    body_row_start: int,
    body_cols: tuple[int, int],   # (axis, profile-NLL) columns, absolute
    editable_bounds: bool = True,
) -> dict:
    """Write one 20-point profile-NLL search stage for a two-parameter fit.

    ``partner_formula(ref, data)`` returns the Excel expression (no leading
    ``=``) for the profiled-out parameter at the searched value ``ref``, over
    the sample expression ``data``.  ``nll_formula(data, p1, p2)`` returns the
    catalog NLL call over an **already-filtered** sample, so its ``[filter]``
    argument is omitted.

    Unlike ``_write_grid_stage``, the control block and the body are positioned
    independently: ``row_start`` / ``col_start`` place the control block, while
    ``body_row_start`` and ``body_cols`` place the axis and profile-NLL columns.
    Both stages of a fit share one column zone and one set of body rows, so
    their control blocks stack vertically while their bodies sit side by side.

    Returns the absolute A1 references the refined Stage 2 and the fitting
    summary read: best_p1, best_p2, min_p1, max_p1, step_p1, n_points, axis,
    and body.
    """
    sname = sheet.name
    n = _N_PROFILE
    r0, c0 = row_start, col_start
    last_col = c0 + _PS_W - 1

    # ── Stage title ──────────────────────────────────────────────────────────
    val(sheet, r0, c0, title)
    sheet.range(rc(r0, c0), rc(r0, last_col)).merge()
    sheet.range(rc(r0, c0)).api.Font.Bold = True
    sheet.range(rc(r0, c0)).color = _HEADER

    control_hdr_row = r0 + _GS_R_CONTROL_HDR
    p1_row = r0 + _GS_R_P1
    p2_row = r0 + _GS_R_P2

    # ── Fixed-area tables ────────────────────────────────────────────────────
    val(sheet, control_hdr_row, c0 + _PS_C_MINNLL, "Min NLL:")
    val(sheet, control_hdr_row, c0 + _PS_C_POINTS, "Grid Points")
    _subheader_row(sheet, control_hdr_row, c0 + _PS_C_MINNLL, c0 + _PS_C_POINTS)
    val(sheet, p1_row, c0 + _PS_C_POINTS, n)
    sheet.range(rc(p1_row, c0 + _PS_C_POINTS)).number_format = _FMT_INT
    n_points_ref = _gs_a1(r0, c0, _GS_R_P1, _PS_C_POINTS)

    for dc, label in [
        (_PS_C_PARAM, "Parameter"),
        (_PS_C_START, "Start"),
        (_PS_C_MIN, "Min"),
        (_PS_C_MAX, "Max"),
        (_PS_C_STEP, "Step Size"),
        (_PS_C_BEST, "Best"),
    ]:
        val(sheet, control_hdr_row, c0 + dc, label)
    _subheader_row(sheet, control_hdr_row, c0 + _PS_C_PARAM, c0 + _PS_C_BEST)

    val(sheet, p1_row, c0 + _PS_C_PARAM, p1_label)
    val(sheet, p2_row, c0 + _PS_C_PARAM, p2_label)

    # Searched-parameter start and bracket.  The Min/Max cells hold formulas
    # but stay input-coloured on Stage 1: they are the one place a user can
    # widen the search when the boundary guard fires, by typing over them.
    f(sheet, p1_row, c0 + _PS_C_START, p1_start)
    f(sheet, p1_row, c0 + _PS_C_MIN, p1_min)
    f(sheet, p1_row, c0 + _PS_C_MAX, p1_max)
    if editable_bounds:
        sheet.range(rc(p1_row, c0 + _PS_C_MIN)).color = _INPUT
        sheet.range(rc(p1_row, c0 + _PS_C_MAX)).color = _INPUT

    min_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _PS_C_MIN)
    max_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _PS_C_MAX)
    step_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _PS_C_STEP)
    f(sheet, p1_row, c0 + _PS_C_STEP,
      f"=({max_p1_ref}-{min_p1_ref})/({n_points_ref}-1)")

    sheet.range(
        rc(p1_row, c0 + _PS_C_START),
        rc(p2_row, c0 + _PS_C_BEST),
    ).number_format = _FMT_1DP

    # ── Axis and profile-NLL body ────────────────────────────────────────────
    # Positioned from the caller's body anchors, not from the control block:
    # the two stages of a fit share these rows and sit in adjacent column pairs.
    hdr_row = body_row_start - 1
    body_row_end = body_row_start + n - 1
    axis_col, body_col = body_cols

    val(sheet, hdr_row, axis_col, p1_label)
    val(sheet, hdr_row, body_col, "Profile NLL")
    _subheader_row(sheet, hdr_row, axis_col, body_col)

    f(sheet, body_row_start, axis_col,
      f"=SEQUENCE({n_points_ref},1,{min_p1_ref},{step_p1_ref})")
    axis_range = sheet.range(rc(body_row_start, axis_col), rc(body_row_end, axis_col))
    axis_range.number_format = _FMT_1DP

    body_range = sheet.range(rc(body_row_start, body_col), rc(body_row_end, body_col))

    axis_name = f"{body_name}_Axis"
    for name, first_col in ((body_name, body_col), (axis_name, axis_col)):
        _drop_wb_name(sheet, name)
        drop_local_name(sheet, name)
        sheet.api.Names.Add(
            Name=name,
            RefersTo=(
                f"='{sname}'!${col_letter(first_col)}${body_row_start}"
                f":${col_letter(first_col)}${body_row_end}"
            ),
        )

    # One NLL call per trial value, with the partner parameter substituted in
    # closed form.  The included sample is bound once per cell as `x` and
    # passed to both the partner and the NLL call: the partner would otherwise
    # re-filter the full 2,000-row input range, and the catalog NLL LAMBDA
    # filters its `data` argument again internally, so the naive form pays two
    # full-range FILTERs per cell — comparable to the likelihood evaluation
    # itself.  `x` is already the numeric included sample (UV_Include *is*
    # `ISNUMBER(...)`), so omitting NLL's optional `[filter]` leaves its
    # `ISNUMBER(x)` default filtering nothing.  The outer IFERROR is the same
    # undefined-NLL sentinel the 2-D bodies inherit from the NLL LAMBDAs; it is
    # explicit here because LET propagates an error in `x` past them.
    # Written as a single N×1 ``Formula2`` array so the stage costs one COM
    # round-trip rather than N.
    axis_cell_refs = [
        f"${col_letter(axis_col)}${body_row_start + i}" for i in range(n)
    ]
    body_range.api.Formula2 = [
        [
            f"=IFERROR(LET(x,{_FILTERED_DATA},p,{ref},"
            f"{nll_formula('x', 'p', partner_formula('p', 'x'))}),1E+15)"
        ]
        for ref in axis_cell_refs
    ]
    body_range.number_format = _FMT_SCI_1DP

    # ── LAMBDA-driven outputs ────────────────────────────────────────────────
    f(
        sheet,
        p1_row,
        c0 + _PS_C_MINNLL,
        f'=IFERROR(TAKE(Grid_Argument_Minimum({body_name}),,1),"—")',
    )
    sheet.range(rc(p1_row, c0 + _PS_C_MINNLL)).number_format = _FMT_SCI_1DP

    # A single-column grid puts the minimum's location in the row slot, so the
    # searched value is INDEX(axis, row_location).  Grid_Search_Optimum is not
    # usable here: its column-parameter half reads the cell above the body,
    # which on a 1-D stage is the "Profile NLL" header, not a parameter value.
    row_location = f"INDEX(Grid_Argument_Minimum({body_name}),1,2)"
    best_p1_ref = _gs_a1(r0, c0, _GS_R_P1, _PS_C_BEST)
    f(
        sheet,
        p1_row,
        c0 + _PS_C_BEST,
        f'=IFERROR(INDEX({axis_name},{row_location}),"—")',
    )
    f(
        sheet,
        p2_row,
        c0 + _PS_C_BEST,
        f'=IFERROR({partner_formula(best_p1_ref)},"—")',
    )

    # Boundary guard: red fill when the optimum lies on an end of the bracket.
    # Only the searched parameter can hit a boundary — the partner is solved,
    # not searched, so it has no grid edge to land on.
    cf = sheet.range(rc(p1_row, c0 + _PS_C_BEST)).api.FormatConditions.Add(
        Type=2,  # xlExpression
        Formula1=f"=OR({row_location}=1,{row_location}={n_points_ref})",
    )
    cf.Interior.Color = 0x0000FF   # red (BGR)
    cf.Font.Color = 0xFFFFFF       # white

    # Same green→yellow→red scale the 2-D heatmap uses, read down the curve.
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

    border_box(sheet, control_hdr_row, c0 + _PS_C_MINNLL, p1_row, c0 + _PS_C_MINNLL)
    border_box(sheet, control_hdr_row, c0 + _PS_C_POINTS, p1_row, c0 + _PS_C_POINTS)
    border_box(sheet, control_hdr_row, c0 + _PS_C_PARAM, p2_row, c0 + _PS_C_BEST)
    border_box(sheet, hdr_row, axis_col, body_row_end, body_col)

    return {
        "best_p1": best_p1_ref,
        "best_p2": _gs_a1(r0, c0, _GS_R_P2, _PS_C_BEST),
        "min_p1": min_p1_ref,
        "max_p1": max_p1_ref,
        "step_p1": step_p1_ref,
        "n_points": n_points_ref,
        "axis": axis_name,
        "body": body_name,
        "axis_hdr": f"${col_letter(axis_col)}${hdr_row}",
        "body_hdr": f"${col_letter(body_col)}${hdr_row}",
    }


def _write_two_stage_profile_search(
    sheet: xw.Sheet,
    col_start: int,
    dist_name: str,
    body_prefix: str,
    p1_label: str,
    p2_label: str,
    nll_formula,
    partner_formula,
    p1_start: str,
) -> dict:
    """Write Stage 1 and Stage 2 profile-search blocks for one distribution.

    Both stages live in this fit's own 9-column zone: their control blocks
    stack vertically (Stage 2 ``_PS_R_STAGE_STRIDE`` rows below Stage 1) and
    their bodies sit side by side beneath, sharing the same rows so the two
    curves can be read against each other.

    Returns the Stage 1 references the profile-NLL chart plots.
    """
    r0 = _ROW_FIT_ZONE
    c0 = col_start
    body_row_start = r0 + _PS_R_BODY
    s1_body_cols, s2_body_cols = (
        (c0 + axis, c0 + body) for axis, body in _PS_BODY_COLS
    )
    start_ref = _gs_a1(r0, c0, _GS_R_P1, _PS_C_START)

    s1 = _write_profile_stage(
        sheet,
        row_start   = r0,
        col_start   = c0,
        title       = f"{dist_name} Profile-NLL MLE  —  Stage 1  ({p1_label}; {p2_label} profiled out)",
        body_name   = f"{body_prefix}_S1",
        p1_label    = p1_label,
        p2_label    = p2_label,
        nll_formula = nll_formula,
        partner_formula = partner_formula,
        p1_start    = p1_start,
        p1_min      = f"=MAX(0.001,{start_ref}/{_PROFILE_BRACKET})",
        p1_max      = f"={start_ref}*{_PROFILE_BRACKET}",
        body_row_start = body_row_start,
        body_cols   = s1_body_cols,
        editable_bounds = True,
    )

    _write_profile_stage(
        sheet,
        row_start   = r0 + _PS_R_STAGE_STRIDE,
        col_start   = c0,
        title       = f"{dist_name} Profile-NLL MLE  —  Stage 2  (refined)",
        body_name   = f"{body_prefix}_S2",
        p1_label    = p1_label,
        p2_label    = p2_label,
        nll_formula = nll_formula,
        partner_formula = partner_formula,
        p1_start    = f"={s1['best_p1']}",
        p1_min      = f"=MAX(0.001,{s1['best_p1']}-{s1['step_p1']})",
        p1_max      = f"={s1['best_p1']}+{s1['step_p1']}",
        body_row_start = body_row_start,
        body_cols   = s2_body_cols,
        editable_bounds = False,
    )

    # OFFSET-based chart ranges over the Stage 1 curve — the wide search, where
    # the basin, the interior minimum, and any boundary hit are visible.  Sized
    # by the Grid Points cell so a change there resizes the plotted series,
    # following the Regression sheet's RegChart* pattern.
    sname = sheet.name
    size = f"MAX(IFERROR('{sname}'!{s1['n_points']},1),1)"
    for suffix, header_ref in (
        ("Axis", s1["axis_hdr"]),
        ("NLL", s1["body_hdr"]),
    ):
        name = f"UV_Profile_{body_prefix.removeprefix('UV_')}_{suffix}"
        _drop_wb_name(sheet, name)
        drop_local_name(sheet, name)
        sheet.api.Names.Add(
            Name=name,
            RefersTo=f"=OFFSET('{sname}'!{header_ref},1,0,{size},1)",
        )

    return s1


# ``nll_formula`` takes the sample expression as its first argument and omits
# the catalog LAMBDA's optional ``[filter]``: the profile bodies pass an
# already-filtered sample, so re-applying a mask would only cost a second pass.
_PROFILE_SEARCHES = [
    {
        "col_start": _C_GS_WB,
        "dist_name": "Weibull",
        "body_prefix": "UV_WB",
        "p1_label": "Shape (k)",
        "p2_label": "Scale (λ)",
        "nll_formula": lambda data, p1, p2: f"NLL_Weibull({data},{p1},{p2})",
        "partner_formula": _weibull_profile_scale,
        "p1_start": _WEIBULL_SHAPE_START,
    },
    {
        "col_start": _C_GS_GAMMA,
        "dist_name": "Gamma",
        "body_prefix": "UV_GAMMA",
        "p1_label": "Shape (α)",
        "p2_label": "Rate (β)",
        "nll_formula": lambda data, p1, p2: f"NLL_Gamma({data},{p1},{p2})",
        "partner_formula": _gamma_profile_rate,
        "p1_start": _GAMMA_SHAPE_START,
    },
]


def _write_two_parameter_grid_search(sheet: xw.Sheet) -> None:
    """Write the two-stage MLE search blocks for all two-parameter fits.

    Weibull and Gamma search one dimension: their scale / rate parameter is
    profiled out in closed form, so each stage is a 20-point profile-NLL curve
    (``_write_profile_stage``) rather than a 20×20 surface.  Beta stays
    two-dimensional — both of its conditional MLEs involve digamma — and keeps
    the artifact's only two Data Tables (``_write_grid_stage``).
    """
    for spec in _PROFILE_SEARCHES:
        _write_two_stage_profile_search(sheet, **spec)

    _write_two_stage_grid_search(
        sheet,
        col_start   = _C_GS_BETA,
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


def _write_profile_charts(sheet: xw.Sheet) -> None:
    """Insert the Weibull and Gamma profile-NLL line charts.

    One chart per one-dimensional fit, plotting Stage 1's profile NLL against
    the searched parameter. This is what replaces the 2-D NLL heatmap those two
    distributions used to carry: a 1-D search would reduce the heatmap to a
    single colour strip, whereas the curve shows the basin, the interior
    minimum, and a boundary hit directly.
    """
    sname = sheet.name
    for i, spec in enumerate(_PROFILE_SEARCHES):
        col_first, col_last = _QQ_CHART_BANDS[i % len(_QQ_CHART_BANDS)]
        row_start = _ROW_PROFILE_CHART_START + (i // len(_QQ_CHART_BANDS)) * _PROFILE_CHART_ROWS
        chart_range = sheet.range(
            rc(row_start, col_first),
            rc(row_start + _PROFILE_CHART_ROWS - 1, col_last),
        )
        co = sheet.api.ChartObjects().Add(
            chart_range.left, chart_range.top, chart_range.width, chart_range.height
        )
        chart = co.Chart

        while chart.SeriesCollection().Count > 0:
            chart.SeriesCollection(1).Delete()

        chart.ChartType = _XL_XY_SCATTER_LINES

        stem = spec["body_prefix"].removeprefix("UV_")
        series = chart.SeriesCollection().NewSeries()
        series.XValues = f"='{sname}'!UV_Profile_{stem}_Axis"
        series.Values = f"='{sname}'!UV_Profile_{stem}_NLL"
        series.Name = f"{spec['dist_name']} profile NLL"
        series.MarkerSize = 4

        chart.HasLegend = False
        chart.HasTitle = True
        chart.ChartTitle.Text = f"{spec['dist_name']} Profile NLL"
        x_axis = chart.Axes(_XL_CATEGORY)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = spec["p1_label"]
        y_axis = chart.Axes(_XL_VALUE)
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = "Profile NLL"


# ── Row height and freeze ─────────────────────────────────────────────────────

def _set_note(
    sheet: xw.Sheet, row: int, col: int, text: str, *, label: str | None = None
) -> None:
    """Replace the cell's note/comment text with a plain-language explanation.

    Sized and anchored to the right of the cell via the shared
    `note_dimensions` + `anchor_comment_right_of_cell` helpers (same
    heuristic as the regression sheet's notes). The Univariate notes run
    130–260 characters, so without explicit sizing the default Excel
    comment box clips them.
    """
    cell = sheet.range(rc(row, col))
    cell_api = cell.api
    try:
        cell_api.ClearComments()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    cell_api.AddComment(text)
    try:
        cell_api.Comment.Visible = False
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    width, height = note_dimensions(label if label is not None else text, text)
    anchor_comment_right_of_cell(sheet, row, col, width, height)


def _annotate_univariate_terms(sheet: xw.Sheet, sheet_notes: dict[str, str]) -> None:
    """Attach plain-language notes to key labels on the Univariate sheet.

    Sites are restricted to column headers and zone headings (per the
    project convention: notes go on labels, never on data cells). Method
    labels, fitting-table headers, Q-Q zone headers, and the grid-search
    control labels each get one short note pulled from sheet_notes.
    """
    note_cells = [
        # Three histogram method labels (row 2, one per binning block)
        (_ROW_METHOD_HDR, _C_STUR, "Sturges"),
        (_ROW_METHOD_HDR, _C_SCOTT, "Scott"),
        (_ROW_METHOD_HDR, _C_FD, "FD"),
        # Distribution fitting table — acronym headers in row 4
        (_ROW_COL_HDRS, _C_NLL, "NLL"),
        (_ROW_COL_HDRS, _C_K_PARAM, "k"),
        (_ROW_COL_HDRS, _C_AIC, "AIC"),
        (_ROW_COL_HDRS, _C_BIC, "BIC"),
        (_ROW_COL_HDRS, _C_AD, "A-D"),
        (_ROW_COL_HDRS, _C_KS, "K-S"),
        # Q-Q plot data zone (row 1 zone heading, row 4 column sub-headers)
        (_ROW_TITLE, _C_QQ, "Q-Q Plot Data"),
        (_ROW_COL_HDRS, _C_QQ + _QQ_P, "P"),
        (_ROW_COL_HDRS, _C_QQ + _QQ_SAMPLE, "Sample"),
    ]

    # Profile-search control and body headers, Stage 1 of each 1-D fit.
    s1_axis_col, s1_body_col = _PS_BODY_COLS[0]
    for spec in _PROFILE_SEARCHES:
        zone_col = spec["col_start"]
        note_cells.append(
            (_ROW_FIT_ZONE + _GS_R_CONTROL_HDR, zone_col + _PS_C_START, "Start")
        )
        note_cells.append(
            (_ROW_FIT_ZONE + _PS_R_BODY_HDR, zone_col + s1_body_col, "Profile NLL")
        )

    for row, col, key in note_cells:
        note_text = sheet_notes.get(key)
        if note_text is not None:
            _set_note(sheet, row, col, note_text, label=key)


def _finalize_sheet(sheet: xw.Sheet) -> None:
    sheet.range(rc(1, 1)).api.EntireRow.RowHeight = 20
    sheet.range(rc(_ROW_METHOD_HDR, 1)).api.EntireRow.RowHeight = 18
    sheet.range(rc(_ROW_SECTION_HDR, 1)).api.EntireRow.RowHeight = 18
    sheet.range(rc(_ROW_COL_HDRS, 1)).api.EntireRow.RowHeight = 18
    # Freeze at C4: rows 1-3 and cols A-B (data + filter) always visible
    sheet.range(rc(_ROW_DATA_START, 3)).api.Select()
    sheet.book.app.api.ActiveWindow.FreezePanes = True


# ── Main entry point ──────────────────────────────────────────────────────────

def write_univariate_sheet(
    workbook: xw.Book,
    sheet_notes: dict[str, str] | None = None,
) -> xw.Sheet:
    """Create or replace the Univariate sheet and write all content.

    Parameters
    ----------
    workbook : xw.Book
        Open xlwings workbook to write into.
    sheet_notes : dict[str, str] | None
        Plain-language tooltips for column headers and zone labels,
        keyed by label text. Pulled from the ``univariate_sheet_notes``
        key in lambda_functions.json by the build script. Optional —
        no notes are attached when ``None`` or empty (the caller's
        ``{}`` fallback keeps the standalone build free of dict wiring).

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
    _write_qq_data(sheet)

    _write_weibull_grid_search(sheet)
    _autofit_column_widths(sheet)

    _write_histogram_chart_title_cells(sheet)
    try:
        _write_histogram_charts(sheet)
    except Exception:
        pass
    try:
        _write_qq_charts(sheet)
    except Exception:
        pass
    try:
        _write_profile_charts(sheet)
    except Exception:
        pass

    # Notes attach after the labels are on the sheet, but before freeze
    # panes so the comment anchors are not displaced by a later resize.
    _annotate_univariate_terms(sheet, sheet_notes or {})

    _finalize_sheet(sheet)

    return sheet

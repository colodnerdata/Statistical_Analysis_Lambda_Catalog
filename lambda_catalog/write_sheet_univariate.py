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
  Col BP–BS      — Weibull fit zone: both profile stages side by side (4 cols)
  Col BU–BX      — Gamma fit zone: both profile stages side by side (4 cols)
  Col BY–CD      — Beta fit zone: both Full_Factorial stages side by side (6 cols)

  One zone per distribution, with that fit's two stages SIDE BY SIDE inside it
  (S1 left, S2 right), so a dynamic Beta body height (N² rows) never makes a
  Stage 2 row anchor depend on Stage 1's spill.  The fit zones come last because
  they are the only ones whose width is a tunable (_N_GRID / _N_PROFILE), so a
  resize displaces nothing.  Gamma sits flush against Beta (no gutter) so the two
  read as one block.

  Charts anchored under the fitting table:
    G14, G34, G54          — histogram combo charts (count bars + 8 fitted overlay lines)
    rows 74-153            — eight per-distribution Q-Q plots in a 4×2 grid
                             (left chart G:O, right chart P:T per row)

  Charts anchored ABOVE their fit zone's body, in the band between the control
  block and the body header (rows 13–30), one zone wide:
    BP13:BS30, BU13:BX30   — Weibull and Gamma profile-NLL charts, each plotting
                             both stages (Stage 2 with + markers). BY13:CD30 is
                             reserved blank for a potential future Beta chart.

Sheet-scoped named ranges
─────────────────────────
  UV_Data        — spill range of the raw column formula ($A$4#, unfiltered)
  UV_Include     — local filter mask spill ($B$4#)
  UV_n           — IFERROR(COUNT(FILTER(UV_Data,UV_Include)), 0)
  UV_Sturges_*, UV_Scott_*, UV_FD_* — OFFSET-based histogram column ranges
  UV_*_<Dist>_Expected — histogram CDF-delta column × Count stat cell (expected
                   counts); feeds the fitted-distribution overlay line series
  UV_QQ_Sample, UV_QQ_<Dist> — OFFSET-based Q-Q chart column ranges
  UV_WB_S1/S2    — Stage 1/2 Weibull profile-NLL bodies (20×1)
  UV_WB_S1/S2_Axis — Stage 1/2 Weibull shape axes (20×1)
  UV_GAMMA_S1/S2  — Stage 1/2 Gamma profile-NLL bodies (20×1)
  UV_GAMMA_S1/S2_Axis — Stage 1/2 Gamma shape axes (20×1)
  UV_BETA_S{1,2}_{Alpha,Beta,NLL} — Stage 1/2 Beta Full_Factorial body columns
                   (N²×1 each), OFFSET ranges sized by the in-sheet Grid Points
                   cell so they track the live grid. Replaces the old Data-Table
                   UV_BETA_S1/S2 ranges.
  UV_Profile_<WB|GAMMA>_<S1|S2>_<Axis|NLL> — OFFSET-based chart ranges, one
                   pair per stage, each sized by its own stage's Grid Points
                   cell; the two series of that fit's profile-NLL chart
"""
from __future__ import annotations

from typing import NamedTuple

import xlwings as xw

from .build_common import MIN_GRID_POINTS
from .sheet_styles import CF_DARK_RED_TEXT as _CF_RED_TEXT
from .sheet_styles import CF_LIGHT_RED_FILL as _CF_RED_FILL
from .sheet_styles import HEADER_COLOR as _HEADER
from .sheet_styles import INPUT_COLOR as _INPUT
from .sheet_styles import SUBHDR_COLOR as _SUBHDR
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    a1,
    add_expression_format,
    anchor_comment_right_of_cell,
    border_box,
    col_letter,
    drop_local_name,
    excel_color,
    f,
    note_dimensions,
    rc,
    section_heading,
    set_column_widths,
    val,
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

# Zone 6: two-parameter MLE search — grid dimensions
#
# Two parameter searches share the right-hand band.  Every stage of every fit
# is TWO dynamic-array spills — a Full_Factorial grid and a BYROW NLL column
# that reads it via the `#` operator — sized by an in-sheet "Grid Points" cell,
# so N is editable live and the NLL column follows the grid to whatever height
# the user asks for:
#   • Weibull and Gamma search ONE dimension (the shape). Their scale / rate
#     parameter is profiled out in closed form, so each stage is an
#     Full_Factorial(N, Min, Max) axis → N rows — the d=1 reduction of the same
#     grid Beta uses — beside a BYROW profile-NLL column (Step documents the
#     spacing and brackets Stage 2; the axis reads Min/Max/N directly).
#     _N_PROFILE is the default written into that cell.
#   • Beta searches TWO dimensions (both conditional MLEs involve digamma).
#     Each stage is a Full_Factorial(N × N) grid → N² rows in two columns
#     (Alpha | Beta), beside a BYROW NLL column.  _N_GRID is the default.
# Both stages sit SIDE BY SIDE inside one column zone, so a dynamic body height
# never makes a Stage 2 row anchor depend on Stage 1's spill.
_N_PROFILE = 20   # default profile points per stage (Weibull / Gamma; editable live)
_N_GRID    = 10    # default Beta grid points per axis (N² rows/stage; editable live)

# Floor on every Grid Points cell, re-exported from build_common so the flag,
# the in-sheet Validation, and the invalid-N conditional format cannot drift
# apart.  See MIN_GRID_POINTS there for why the floor is 2 and not 1.
_MIN_GRID_POINTS = MIN_GRID_POINTS

# ── Fit-zone row skeleton (offsets from _ROW_FIT_ZONE = 1) ───────────────────
# Every fit zone uses the same skeleton: a control field-list, a chart band,
# then the body.  The profile-NLL chart sits ABOVE the body (rows 13–30),
# between the control block and the body — the reverse of the old layout,
# which put it below.  Beta reserves the same band blank for a future chart.
_R_TITLE         = 0    # row 1 — fit title (merged across the zone)
_R_STAGE_SUBHDR  = 2    # row 3 — "Stage 1 (Wide Scope)" / "Stage 2 (refined)"
_R_CONTROL_FIRST = 3    # row 4 — first control field
_R_CHART_TOP     = 12   # row 13 — chart top (Weibull/Gamma); reserved blank (Beta)
_R_CHART_BOTTOM  = 29   # row 30 — chart bottom
_R_BODY_HDR      = 31   # row 32 — body column headers
_R_BODY          = 32   # row 33 — first body row (N rows for profile; N² for Beta)

# Every fit zone's row anchor.  The zones sit side by side, so they all start
# on the same row and differ only in column.
_ROW_FIT_ZONE = 1

# ── Weibull / Gamma profile search (4-col zone, two stages side by side) ─────
# col0 = labels + S1 axis body; col1 = S1 value + S1 NLL body;
# col2 = S2 value + S2 axis body; col3 = S2 NLL body (no control cell).
# Control is a vertical field-list (label col0, S1 value col1, S2 value col2).
_PR_C_LABEL  = 0
_PR_C_S1     = 1   # S1 control value / S1 NLL body
_PR_C_S2     = 2   # S2 control value / S2 axis body
_PR_C_S2_NLL = 3   # S2 NLL body (no control cell)
_PR_W        = 4   # profile zone width

# Profile-zone control field rows (offsets 3–10 → rows 4–11).  Recovery and the
# profiled-out partner use the same Grid_Argument_Minimum mechanism as before.
_PR_R_GRID_POINTS = 3   # row 4 — Grid Points (default _N_PROFILE; editable live)
_PR_R_MIN         = 4
_PR_R_MAX         = 5
_PR_R_START       = 6   # closed-form starting value (searched parameter)
_PR_R_MINNLL      = 7
_PR_R_STEP        = 8
_PR_R_BEST_P1     = 9   # Optimal Shape (k) / (α) — searched parameter's MLE
_PR_R_BEST_P2     = 10  # Optimal Scale (λ) / Rate (β) — profiled out, closed form

# ── Beta Full_Factorial search (6-col zone, two stages side by side) ────────
# 3 cols/stage (Alpha | Beta | NLL).  col0 = labels + S1 Alpha body;
# col1 = S1 α value + S1 Beta body; col2 = S1 β value + S1 NLL body;
# col3 = S2 α value + S2 Alpha body; col4 = S2 β value + S2 Beta body;
# col5 = S2 NLL body (no control cell).  α values live in the α cols (1/3),
# β values in the β cols (2/4), so the two-parameter control block is compact.
_BETA_C_LABEL  = 0
_BETA_C_S1_A   = 1   # S1 α value / S1 Beta body
_BETA_C_S1_B   = 2   # S1 β value / S1 NLL body
_BETA_C_S2_A   = 3   # S2 α value / S2 Alpha body
_BETA_C_S2_B   = 4   # S2 β value / S2 Beta body
_BETA_C_S2_NLL = 5   # S2 NLL body (no control cell)
_BETA_W        = 6   # Beta zone width

# Beta control field rows (offsets 3–11 → rows 4–12): two searched parameters.
# α Min/Max/Step in the α cols (S1 col1, S2 col3); β Min/Max/Step in the β cols
# (S1 col2, S2 col4); Min NLL in the NLL cols (S1 col2, S2 col5); Best α in the
# α cols and Best β in the β cols (row 12).  Rows 4–12 clear the chart band.
_BETA_R_GRID_POINTS = 3   # row 4
_BETA_R_A_MIN       = 4
_BETA_R_A_MAX       = 5
_BETA_R_A_STEP      = 6
_BETA_R_B_MIN       = 7
_BETA_R_B_MAX       = 8
_BETA_R_B_STEP      = 9
_BETA_R_MINNLL      = 10  # row 11 — Min NLL (S1 col2, S2 col5)
_BETA_R_BEST        = 11  # row 12 — Best α (α cols) and Best β (β cols)

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
# width is a tunable (_N_GRID / _N_PROFILE), so keeping them at the end means a
# resize displaces nothing.  Same reasoning as the Regression sheet's rule that
# nothing may sit right of the design-matrix zone.  The Q-Q data is a fixed 10
# columns and is read alongside the fitting table, so it leads.
#
# Each zone is followed by one gap column, with ONE exception: Beta (BY:CD)
# follows Gamma (BU:BX) with NO gutter between them, per the owner's explicit
# column letters (the two stages sit side by side across the pair, so a gutter
# would split a single visual block).  The per-zone ``gutter_after`` flag below
# encodes that one relaxation; the histogram band's own trailing gap (BD) serves
# as the leading one before the Q-Q zone.
#
# The band is Zone 5 (Q-Q data) then Zone 6 (the three fit zones) — the two
# swapped physical order when Q-Q moved ahead of the fits, and the numbering
# was swapped with them so it still reads left to right.
_QQ_W = 10   # Zone 5 width: P, Sample, 8 theoretical-quantile columns

_BAND_FIRST_COL = 57   # BE — first column after the FD histogram block's gap
_BAND_GAP_W = 1
# (zone name, width, gutter_after) — width is the content span; gutter_after
# is False only for gamma, which sits flush against Beta.
_BAND_ZONES = (
    ("qq", _QQ_W, True),        # BE–BN  → gap BO
    ("weibull", _PR_W, True),    # BP–BS  → gap BT
    ("gamma", _PR_W, False),     # BU–BX  → no gap (Beta flush to the right)
    ("beta", _BETA_W, False),    # BY–CD  → last zone
)


def _derive_band_columns() -> tuple[dict[str, int], tuple[int, ...], int]:
    """Return (zone first columns, gap columns, last populated column)."""
    starts: dict[str, int] = {}
    gaps: list[int] = []
    col = _BAND_FIRST_COL
    for index, (name, width, gutter_after) in enumerate(_BAND_ZONES):
        starts[name] = col
        col += width
        if index < len(_BAND_ZONES) - 1 and gutter_after:
            gaps.append(col)
            col += _BAND_GAP_W
    return starts, tuple(gaps), col - 1


_BAND_COL, _BAND_GAP_COLS, _BAND_LAST_COL = _derive_band_columns()

_C_QQ        = _BAND_COL["qq"]        # 57  (BE)
_C_GS_WB     = _BAND_COL["weibull"]   # 68  (BP)
_C_GS_GAMMA  = _BAND_COL["gamma"]     # 73  (BU)
_C_GS_BETA   = _BAND_COL["beta"]      # 77  (BY)

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
_XL_MARKER_PLUS      = 9       # xlMarkerStylePlus — the refined Stage 2 series

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

# Each profile-NLL chart sits ABOVE its fit zone's body, in the band between the
# control block (rows 4–11/12) and the body header (row 32): the curve it draws
# is the body column directly below it, so the two read together. The chart is
# one zone wide. Beta reserves the same band blank (BY13:CD30) for a potential
# future chart.
_ROW_PROFILE_CHART  = _ROW_FIT_ZONE + _R_CHART_TOP    # 13
_PROFILE_CHART_ROWS = _R_CHART_BOTTOM - _R_CHART_TOP + 1   # 18 (rows 13–30)

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

    # Fit zones: compact controls above the bodies. Weibull/Gamma are 4-col
    # profile zones (two stages side by side); Beta is a 6-col Full_Factorial
    # zone.  The label column (col0) carries the field names, so it is wider;
    # the value/body columns stay narrow.  Autofit refines these afterwards.
    for zone_start, zone_width, label_w in (
        (_C_GS_WB, _PR_W, 22),
        (_C_GS_GAMMA, _PR_W, 22),
        (_C_GS_BETA, _BETA_W, 16),
    ):
        set_column_widths(sheet, ((c, 10) for c in range(zone_start, zone_start + zone_width)))
        set_column_widths(sheet, ((zone_start, label_w),))

    set_column_widths(sheet, ((gap_col, 2) for gap_col in _BAND_GAP_COLS))


def _autofit_column_widths(sheet: xw.Sheet) -> None:
    """Autofit all populated layout columns, then restore intentional gaps."""
    sheet.range(rc(_ROW_TITLE, _C_A), rc(_DATA_END, _BAND_LAST_COL)).columns.autofit()

    gap_cols = [3, 6, 20]  # C, F, T
    for block_start in (_C_STUR, _C_SCOTT, _C_FD):
        gap_cols.append(block_start + _HIST_W)  # AF, AR, BD
    gap_cols.extend(_BAND_GAP_COLS)             # BO, BT (no gap between Gamma and Beta)
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


# Where each fit's Stage 2 Best cells sit: (zone first column, Best-P1 column
# offset, Best-P2 column offset, Best-P1 row offset, Best-P2 row offset).  This
# is the one place the fitting table and the search writers agree on, so the
# table can never reference a block the writers did not put there.
#
# Stages sit SIDE BY SIDE, so Stage 2's Best cells are at FIXED control rows in
# the Stage 2 value column(s), independent of Beta's dynamic N² body height.  Weibull/Gamma put both Best cells in the one Stage 2
# value column (Optimal Shape row 10, Optimal Scale row 11); Beta splits Best α
# and Best β across the two Stage 2 value columns, both on row 12.
_STAGE2_ANCHORS: dict[str, tuple[int, int, int, int, int]] = {
    "Weibull": (_C_GS_WB, _PR_C_S2, _PR_C_S2, _PR_R_BEST_P1, _PR_R_BEST_P2),
    "Gamma":   (_C_GS_GAMMA, _PR_C_S2, _PR_C_S2, _PR_R_BEST_P1, _PR_R_BEST_P2),
    "Beta":    (_C_GS_BETA, _BETA_C_S2_A, _BETA_C_S2_B, _BETA_R_BEST, _BETA_R_BEST),
}


def _final_grid_best_refs(distribution: str, *, beta_grid_size: int = _N_GRID) -> tuple[str, str]:
    """Return the Stage 2 Best cells for a fit, as fitting-table formulas.

    Reads ``_STAGE2_ANCHORS`` so the reference follows the zone the search
    writer actually produced.  Stages sit side by side, so the Stage 2 Best
    cells anchor at fixed control rows regardless of grid size; ``beta_grid_size``
    is accepted for call-site compatibility but does not move the anchor.
    """
    _ = beta_grid_size  # side-by-side stages anchor at fixed rows; N is read in-sheet
    zone_col, p1_col, p2_col, p1_row, p2_row = _STAGE2_ANCHORS[distribution]
    return (
        f"=${col_letter(zone_col + p1_col)}${_ROW_FIT_ZONE + p1_row}",
        f"=${col_letter(zone_col + p2_col)}${_ROW_FIT_ZONE + p2_row}",
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


def _dist_rows(base_row: int, *, beta_grid_size: int = _N_GRID) -> list[tuple]:
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
    beta_alpha_ref, beta_beta_ref = _final_grid_best_refs("Beta", beta_grid_size=beta_grid_size)

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


def _write_fitting_table(sheet: xw.Sheet, *, beta_grid_size: int = _N_GRID) -> None:
    # Zone heading at row 2, merged across all fitting columns
    section_heading(sheet, _ROW_METHOD_HDR, _C_FIT_FIRST, "Distribution Fitting/Comparison")
    sheet.range(rc(_ROW_METHOD_HDR, _C_FIT_FIRST), rc(_ROW_METHOD_HDR, _C_FIT_LAST)).merge()

    for col, label in _FIT_COL_HDRS:
        val(sheet, _ROW_COL_HDRS, col, label)
    _subheader_row(sheet, _ROW_COL_HDRS, _C_FIT_FIRST, _C_FIT_LAST)

    dist_data = _dist_rows(_ROW_DIST_START, beta_grid_size=beta_grid_size)
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

# ── Zone 6: two-parameter MLE search ─────────────────────────────────────────
#
# Every fit is one column zone whose two stages sit SIDE BY SIDE (S1 left,
# S2 right), so a dynamic Beta body height (N² rows) never makes a Stage 2 row
# anchor depend on Stage 1's spill.  Each zone shares one row skeleton:
#
#   r0+0   : fit title merged across the zone (_R_TITLE)
#   r0+2   : stage sub-headers — "Stage 1 (Wide Scope)" / "Stage 2 (refined)"
#            (_R_STAGE_SUBHDR), one over each stage's value column(s)
#   r0+3…  : control field-list — label in col0, Stage 1 value(s) and Stage 2
#            value(s) in the value columns (_R_CONTROL_FIRST …)
#   r0+12..29 : profile-NLL chart (Weibull/Gamma), one zone wide, ABOVE the
#            body (_R_CHART_TOP .. _R_CHART_BOTTOM).  Beta reserves this band
#            blank for a potential future chart.
#   r0+31  : body column headers (_R_BODY_HDR)
#   r0+32… : body — N rows for Weibull/Gamma; N² rows for Beta (_R_BODY)
#
# Two writers, one body shape.  Each stage is a Full_Factorial spill beside a
# BYROW NLL spill that reads it through the `#` operator, so the NLL column is
# always exactly as tall as the grid it scores and both follow the stage's live
# Grid Points cell.  OFFSET named ranges track that same cell, and recovery is
# Grid_Argument_Minimum over the materialized NLL column.
#   _write_profile_fit  — Weibull/Gamma, 4-col zone.  One dimension searched;
#            scale/rate profiled out in closed form.  Body = an
#            Full_Factorial(N, Min, Max) axis (N rows) beside its BYROW column.
#   _write_beta_fit     — Beta, 6-col zone (3 cols/stage: Alpha | Beta | NLL).
#            Two dimensions searched.  Body = a Full_Factorial grid (N²×2,
#            Alpha | Beta) across the first two cols beside its BYROW column.
#            The grid stands alone as the pure Cartesian product; the NLL reads
#            it by reference.


def _gs_a1(row_start: int, col_start: int, dr: int, dc: int) -> str:
    """Absolute A1 ref offset from (row_start, col_start) by (dr, dc)."""
    return f"${col_letter(col_start + dc)}${row_start + dr}"


# The included sample, as an expression.  Body cells bind it to a LET name
# once and pass that name down; the standalone ``Best`` cells, which evaluate the
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
# The Weibull start uses sorted Hazen plotting positions (i-0.5)/n — not
# Rank_Fraction (i/n, which is exactly 1 at the max and errors LN(-LN(1-p))).
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


# Window over which a profile body's number formats, colour scale, and border
# box are painted.  Both fit writers share one rule: the body is a live spill
# whose height follows the stage's Grid Points cell, so the static formatting is
# applied over the DEFAULT-size window and rows that a live N-increase adds
# beyond it stay unshaded.  Cosmetic, not correctness — the named ranges and the
# recovery formulas track the real height.  (Beta's counterpart is
# _BETA_BODY_CF_ROWS_CAP, squared because its body is N² rows.)
_PROFILE_BODY_CF_ROWS_CAP = _N_PROFILE


# Excel COM validation constants, declared locally the way write_spec_block.py
# declares its own (the xlwings constants module is not imported here).
_XL_VALIDATE_WHOLE_NUMBER = 1
_XL_VALID_ALERT_STOP = 1
_XL_GREATER_EQUAL = 7


def _guard_grid_points_cell(sheet: xw.Sheet, row: int, col: int) -> None:
    """Fill, validate, and flag a Grid Points cell.

    The cell drives a live spill, so it is an input: it takes ``INPUT_COLOR``,
    a whole-number Validation with a Stop alert at ``_MIN_GRID_POINTS``, and a
    conditional format that turns red on anything the Validation cannot see.
    Both guards are wanted — a Stop alert catches typing, and the CF catches a
    paste, which bypasses Validation entirely.  ``--beta-grid-size`` applies the
    same floor at build time via ``build_common.positive_grid_size``.
    """
    cell = sheet.range(rc(row, col))
    cell.color = _INPUT

    # Unguarded, like write_spec_block's validations: a try/except here would
    # swallow a malformed rule and ship a cell with no validation at all, which
    # no test could see.
    cell.api.Validation.Delete()
    cell.api.Validation.Add(
        Type=_XL_VALIDATE_WHOLE_NUMBER,
        AlertStyle=_XL_VALID_ALERT_STOP,
        Operator=_XL_GREATER_EQUAL,
        Formula1=str(_MIN_GRID_POINTS),
    )
    cell.api.Validation.IgnoreBlank = False

    ref = f"${col_letter(col)}${row}"
    add_expression_format(
        sheet,
        rc(row, col),
        f"=OR(NOT(ISNUMBER({ref})),{ref}<{_MIN_GRID_POINTS},INT({ref})<>{ref})",
        fill=_CF_RED_FILL,
        font_color=_CF_RED_TEXT,
    )


def _write_profile_fit(
    sheet: xw.Sheet,
    *,
    col_start: int,
    dist_name: str,
    body_prefix: str,
    p1_label: str,
    p2_label: str,
    nll_formula,
    partner_formula,
    p1_start: str,
) -> dict:
    """Write both stages of a 1-D profile-NLL fit, side by side in a 4-col zone.

    col0 carries the field labels and the S1 axis body; col1 the S1 values and
    S1 NLL body; col2 the S2 values and S2 axis body; col3 the S2 NLL body.
    Control is a vertical field-list (rows 4–11); the profile-NLL chart occupies
    rows 13–30; the body starts at row 33 and runs N rows, where N is the
    stage's live Grid Points cell (default ``_N_PROFILE``).  Each stage is two
    spills: an ``Full_Factorial(N, Min, Max)`` axis — the d=1 reduction of the
    same grid Beta uses — and a ``BYROW`` profile-NLL column that reads that
    axis through the ``#`` operator, so the two heights always agree.  Step
    (=(Max-Min)/(N-1)) documents the spacing and brackets Stage 2's Min/Max
    around Stage 1's optimum by ±1 step.  Both stages use the same
    ``Grid_Argument_Minimum`` recovery and the profiled-out partner closed form.

    ``partner_formula(ref, data)`` returns the Excel expression (no leading ``=``)
    for the profiled-out parameter at the searched value ``ref``; ``nll_formula(
    data, p1, p2)`` returns the catalog NLL call over an already-filtered sample
    (its ``[filter]`` is omitted).
    """
    sname = sheet.name
    n = _N_PROFILE
    r0, c0 = _ROW_FIT_ZONE, col_start
    last_col = c0 + _PR_W - 1
    body_hdr_row = r0 + _R_BODY_HDR
    body_row = r0 + _R_BODY
    # The default-size formatting window (see _PROFILE_BODY_CF_ROWS_CAP).
    body_row_end = body_row + max(n, _PROFILE_BODY_CF_ROWS_CAP) - 1

    # ── Title + stage sub-headers ─────────────────────────────────────────────
    val(sheet, r0, c0, f"{dist_name} Profile-NLL MLE  —  {p1_label}; {p2_label} profiled out")
    sheet.range(rc(r0, c0), rc(r0, last_col)).merge()
    title_range = sheet.range(rc(r0, c0), rc(r0, last_col))
    title_range.color = _HEADER
    title_range.api.Font.Bold = True

    sub_row = r0 + _R_STAGE_SUBHDR
    val(sheet, sub_row, c0 + _PR_C_S1, "Stage 1 (Wide Scope)")
    val(sheet, sub_row, c0 + _PR_C_S2, "Stage 2 (refined)")
    _subheader_row(sheet, sub_row, c0 + _PR_C_S1, c0 + _PR_C_S2)

    # ── Control field-list (rows 4–11): label col0, S1 value col1, S2 value col2 ─
    fields = [
        (_PR_R_GRID_POINTS, "Grid Points"),
        (_PR_R_MIN,         "Min"),
        (_PR_R_MAX,         "Max"),
        (_PR_R_START,       "Start"),
        (_PR_R_MINNLL,      "Min NLL"),
        (_PR_R_STEP,        "Step Size"),
        (_PR_R_BEST_P1,     f"Optimal {p1_label}"),
        (_PR_R_BEST_P2,     f"Optimal {p2_label} (profiled)"),
    ]
    for r_off, label in fields:
        val(sheet, r0 + r_off, c0, label)

    s1_n_ref    = _gs_a1(r0, c0, _PR_R_GRID_POINTS, _PR_C_S1)
    s1_min_ref  = _gs_a1(r0, c0, _PR_R_MIN,        _PR_C_S1)
    s1_max_ref  = _gs_a1(r0, c0, _PR_R_MAX,        _PR_C_S1)
    s1_step_ref = _gs_a1(r0, c0, _PR_R_STEP,       _PR_C_S1)
    s1_start_ref= _gs_a1(r0, c0, _PR_R_START,      _PR_C_S1)
    s1_best1_ref= _gs_a1(r0, c0, _PR_R_BEST_P1,     _PR_C_S1)
    s1_best2_ref= _gs_a1(r0, c0, _PR_R_BEST_P2,     _PR_C_S1)
    s2_n_ref    = _gs_a1(r0, c0, _PR_R_GRID_POINTS, _PR_C_S2)
    s2_min_ref  = _gs_a1(r0, c0, _PR_R_MIN,        _PR_C_S2)
    s2_max_ref  = _gs_a1(r0, c0, _PR_R_MAX,        _PR_C_S2)
    s2_step_ref = _gs_a1(r0, c0, _PR_R_STEP,       _PR_C_S2)
    s2_best1_ref= _gs_a1(r0, c0, _PR_R_BEST_P1,     _PR_C_S2)
    s2_best2_ref= _gs_a1(r0, c0, _PR_R_BEST_P2,     _PR_C_S2)

    body_s1 = f"{body_prefix}_S1"
    body_s2 = f"{body_prefix}_S2"
    axis_s1 = f"{body_prefix}_S1_Axis"
    axis_s2 = f"{body_prefix}_S2_Axis"

    # Stage 1 control values.  Min/Max stay input-coloured: they are the one
    # place a user widens the search when the boundary guard fires.
    val(sheet, r0 + _PR_R_GRID_POINTS, c0 + _PR_C_S1, n)
    sheet.range(rc(r0 + _PR_R_GRID_POINTS, c0 + _PR_C_S1)).number_format = _FMT_INT
    _guard_grid_points_cell(sheet, r0 + _PR_R_GRID_POINTS, c0 + _PR_C_S1)
    f(sheet, r0 + _PR_R_START, c0 + _PR_C_S1, p1_start)
    f(sheet, r0 + _PR_R_MIN, c0 + _PR_C_S1, f"=MAX(0.001,{s1_start_ref}/{_PROFILE_BRACKET})")
    f(sheet, r0 + _PR_R_MAX, c0 + _PR_C_S1, f"={s1_start_ref}*{_PROFILE_BRACKET}")
    sheet.range(rc(r0 + _PR_R_MIN, c0 + _PR_C_S1)).color = _INPUT
    sheet.range(rc(r0 + _PR_R_MAX, c0 + _PR_C_S1)).color = _INPUT
    f(sheet, r0 + _PR_R_STEP, c0 + _PR_C_S1,
      f"=({s1_max_ref}-{s1_min_ref})/({s1_n_ref}-1)")
    f(sheet, r0 + _PR_R_MINNLL, c0 + _PR_C_S1,
      f'=IFERROR(TAKE(Grid_Argument_Minimum({body_s1}),,1),"—")')
    f(sheet, r0 + _PR_R_BEST_P1, c0 + _PR_C_S1,
      f'=IFERROR(INDEX({axis_s1},INDEX(Grid_Argument_Minimum({body_s1}),1,2)),"—")')
    f(sheet, r0 + _PR_R_BEST_P2, c0 + _PR_C_S1,
      f'=IFERROR({partner_formula(s1_best1_ref)},"—")')

    # Stage 2 control values — bracket Stage 1's optimum by ±1 step.
    f(sheet, r0 + _PR_R_GRID_POINTS, c0 + _PR_C_S2, f"={s1_n_ref}")
    sheet.range(rc(r0 + _PR_R_GRID_POINTS, c0 + _PR_C_S2)).number_format = _FMT_INT
    f(sheet, r0 + _PR_R_START, c0 + _PR_C_S2, f"={s1_best1_ref}")
    f(sheet, r0 + _PR_R_MIN, c0 + _PR_C_S2, f"=MAX(0.001,{s1_best1_ref}-{s1_step_ref})")
    f(sheet, r0 + _PR_R_MAX, c0 + _PR_C_S2, f"={s1_best1_ref}+{s1_step_ref}")
    f(sheet, r0 + _PR_R_STEP, c0 + _PR_C_S2,
      f"=({s2_max_ref}-{s2_min_ref})/({s2_n_ref}-1)")
    f(sheet, r0 + _PR_R_MINNLL, c0 + _PR_C_S2,
      f'=IFERROR(TAKE(Grid_Argument_Minimum({body_s2}),,1),"—")')
    f(sheet, r0 + _PR_R_BEST_P1, c0 + _PR_C_S2,
      f'=IFERROR(INDEX({axis_s2},INDEX(Grid_Argument_Minimum({body_s2}),1,2)),"—")')
    f(sheet, r0 + _PR_R_BEST_P2, c0 + _PR_C_S2,
      f'=IFERROR({partner_formula(s2_best1_ref)},"—")')

    sheet.range(
        rc(r0 + _PR_R_MIN, c0 + _PR_C_S1),
        rc(r0 + _PR_R_BEST_P2, c0 + _PR_C_S2),
    ).number_format = _FMT_1DP
    sheet.range(rc(r0 + _PR_R_MINNLL, c0 + _PR_C_S1)).number_format = _FMT_SCI_1DP
    sheet.range(rc(r0 + _PR_R_MINNLL, c0 + _PR_C_S2)).number_format = _FMT_SCI_1DP

    # ── Body (rows 33–52): axes + profile NLL, two stages side by side ─────────
    val(sheet, body_hdr_row, c0 + _PR_C_LABEL, p1_label)
    val(sheet, body_hdr_row, c0 + _PR_C_S1, "Profile NLL")
    val(sheet, body_hdr_row, c0 + _PR_C_S2, p1_label)
    val(sheet, body_hdr_row, c0 + _PR_C_S2_NLL, "Profile NLL")
    _subheader_row(sheet, body_hdr_row, c0, last_col)

    # Each stage's axis is Full_Factorial(N, Min, Max) — N evenly-spaced points
    # from Min to Max.  This is the d=1 reduction of the same grid Beta uses:
    # The axis reads Min, Max and N directly; Step (below) documents the spacing
    # (=(Max-Min)/(N-1)) and brackets Stage 2.  Full_Factorial's MAX(1,N-1)
    # divisor makes the N=1 case a single point at Min, matching SEQUENCE.
    f(sheet, body_row, c0 + _PR_C_LABEL,
      f"=Full_Factorial({s1_n_ref},{s1_min_ref},{s1_max_ref})")
    f(sheet, body_row, c0 + _PR_C_S2,
      f"=Full_Factorial({s2_n_ref},{s2_min_ref},{s2_max_ref})")
    sheet.range(rc(body_row, c0 + _PR_C_LABEL), rc(body_row_end, c0 + _PR_C_LABEL)).number_format = _FMT_1DP
    sheet.range(rc(body_row, c0 + _PR_C_S2), rc(body_row_end, c0 + _PR_C_S2)).number_format = _FMT_1DP

    # One NLL call per trial value, partner substituted in closed form, as a
    # single BYROW spill over the stage's axis spill.  Reading the axis through
    # `<anchor>#` is what keeps the two columns the same height: raise Grid
    # Points and the axis grows, and the NLL column grows with it.
    #
    # Three details are load-bearing.  INDEX(r,1,1) scalarizes the row BYROW
    # hands the callback (a 1x1 array passed straight into the NLL call or the
    # partner closed form broadcasts instead of evaluating).  IFERROR sits
    # INSIDE the LAMBDA, so a single non-evaluable trial value costs its own row
    # rather than collapsing the column.  And the filtered sample is bound once
    # per stage as `x` and passed to both the partner and the NLL call — the
    # per-row form paid a full-range FILTER for every point.
    def _stage_nll(axis_anchor: str) -> str:
        return (
            f"=LET(x,{_FILTERED_DATA},"
            f"BYROW({axis_anchor}#,LAMBDA(r,LET(p,INDEX(r,1,1),"
            f"IFERROR({nll_formula('x', 'p', partner_formula('p', 'x'))},1E+15)))))"
        )

    s1_axis_anchor = f"${col_letter(c0 + _PR_C_LABEL)}${body_row}"
    s2_axis_anchor = f"${col_letter(c0 + _PR_C_S2)}${body_row}"
    f(sheet, body_row, c0 + _PR_C_S1, _stage_nll(s1_axis_anchor))
    f(sheet, body_row, c0 + _PR_C_S2_NLL, _stage_nll(s2_axis_anchor))

    s1_body_range = sheet.range(rc(body_row, c0 + _PR_C_S1), rc(body_row_end, c0 + _PR_C_S1))
    s2_body_range = sheet.range(rc(body_row, c0 + _PR_C_S2_NLL), rc(body_row_end, c0 + _PR_C_S2_NLL))
    s1_body_range.number_format = _FMT_SCI_1DP
    s2_body_range.number_format = _FMT_SCI_1DP

    # ── OFFSET named ranges (height N, tracking the live Grid Points cell) ────
    # The same shape _write_beta_fit uses for UV_BETA_S*, without the square: a
    # profile axis is N points, not N².  Sizing these off the cell rather than
    # pinning them to the default window is what lets Min NLL, the Optimal
    # Shape recovery, and the boundary guard all see a resized grid.
    for name, col_off, n_ref in (
        (body_s1, _PR_C_S1,     s1_n_ref), (axis_s1, _PR_C_LABEL, s1_n_ref),
        (body_s2, _PR_C_S2_NLL, s2_n_ref), (axis_s2, _PR_C_S2,    s2_n_ref),
    ):
        cl = col_letter(c0 + col_off)
        _drop_wb_name(sheet, name)
        drop_local_name(sheet, name)
        sheet.api.Names.Add(
            Name=name,
            RefersTo=(
                f"=OFFSET('{sname}'!${cl}${body_row},0,0,"
                f"MAX(IFERROR('{sname}'!{n_ref},1),1),1)"
            ),
        )

    # ── OFFSET chart ranges (sized by each stage's Grid Points cell) ───────────
    stem = body_prefix.removeprefix("UV_")
    for stage, n_cell, axis_col_off, nll_col_off in (
        ("S1", s1_n_ref, _PR_C_LABEL, _PR_C_S1),
        ("S2", s2_n_ref, _PR_C_S2, _PR_C_S2_NLL),
    ):
        size = f"MAX(IFERROR('{sname}'!{n_cell},1),1)"
        for suffix, col_off in (("Axis", axis_col_off), ("NLL", nll_col_off)):
            name = f"UV_Profile_{stem}_{stage}_{suffix}"
            cl = col_letter(c0 + col_off)
            _drop_wb_name(sheet, name)
            drop_local_name(sheet, name)
            sheet.api.Names.Add(
                Name=name,
                RefersTo=f"=OFFSET('{sname}'!${cl}${body_hdr_row},1,0,{size},1)",
            )

    # ── Boundary guard (Best P1 cells) + NLL colour scale ─────────────────────
    for best_row, best_col, body_name, n_ref in (
        (r0 + _PR_R_BEST_P1, c0 + _PR_C_S1, body_s1, s1_n_ref),
        (r0 + _PR_R_BEST_P1, c0 + _PR_C_S2, body_s2, s2_n_ref),
    ):
        loc = f"INDEX(Grid_Argument_Minimum({body_name}),1,2)"
        cf = sheet.range(rc(best_row, best_col)).api.FormatConditions.Add(
            Type=2,  # xlExpression
            Formula1=f"=OR({loc}=1,{loc}={n_ref})",
        )
        cf.Interior.Color = 0x0000FF   # red (BGR)
        cf.Font.Color = 0xFFFFFF       # white

    for body_range in (s1_body_range, s2_body_range):
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

    border_box(sheet, r0 + _PR_R_GRID_POINTS, c0, r0 + _PR_R_BEST_P2, c0 + _PR_C_S2)
    border_box(sheet, body_hdr_row, c0 + _PR_C_LABEL, body_row_end, c0 + _PR_C_S1)
    border_box(sheet, body_hdr_row, c0 + _PR_C_S2, body_row_end, c0 + _PR_C_S2_NLL)

    return {
        "s1": {"best_p1": s1_best1_ref, "best_p2": s1_best2_ref,
                "step_p1": s1_step_ref, "n_points": s1_n_ref},
        "s2": {"best_p1": s2_best1_ref, "best_p2": s2_best2_ref,
                "step_p1": s2_step_ref, "n_points": s2_n_ref},
    }


# Window over which the Beta NLL colour scale is applied — the N² counterpart
# of _PROFILE_BODY_CF_ROWS_CAP, and the same rule: painted over the default-size
# window so the shipped artifact is fully shaded, with rows a live N-increase
# adds beyond it staying unshaded (cosmetic, not correctness).
_BETA_BODY_CF_ROWS_CAP = 20 * 20


def _write_beta_fit(sheet: xw.Sheet, beta_grid_size: int = _N_GRID) -> dict:
    """Write both stages of Beta's 2-D Full_Factorial search, side by side.

    6-col zone (3 cols/stage: Alpha | Beta | NLL): each stage's body is TWO
    spills at row 33 — a ``Full_Factorial`` grid (N²×2, Alpha | Beta) across the
    stage's first two cols, and a ``BYROW`` NLL spill in the third col that
    reads the grid via the ``#`` operator — so the two stages grow down
    independently.  N is an in-sheet cell (editable live); OFFSET named ranges
    track N².  Recovery is ``Grid_Argument_Minimum`` over the materialized NLL
    column, the same mechanism the profile fits use.
    """
    sname = sheet.name
    n = beta_grid_size
    r0, c0 = _ROW_FIT_ZONE, _C_GS_BETA
    last_col = c0 + _BETA_W - 1
    body_hdr_row = r0 + _R_BODY_HDR
    body_row = r0 + _R_BODY

    def a1(r_off: int, c_off: int) -> str:
        return _gs_a1(r0, c0, r_off, c_off)

    s1_n    = a1(_BETA_R_GRID_POINTS, _BETA_C_S1_A)
    s1_amin = a1(_BETA_R_A_MIN,  _BETA_C_S1_A)
    s1_amax = a1(_BETA_R_A_MAX,  _BETA_C_S1_A)
    s1_astep= a1(_BETA_R_A_STEP, _BETA_C_S1_A)
    s1_bmin = a1(_BETA_R_B_MIN,  _BETA_C_S1_B)
    s1_bmax = a1(_BETA_R_B_MAX,  _BETA_C_S1_B)
    s1_bstep= a1(_BETA_R_B_STEP, _BETA_C_S1_B)
    s1_best_a = a1(_BETA_R_BEST, _BETA_C_S1_A)
    s1_best_b = a1(_BETA_R_BEST, _BETA_C_S1_B)
    s2_n    = a1(_BETA_R_GRID_POINTS, _BETA_C_S2_A)
    s2_amin = a1(_BETA_R_A_MIN,  _BETA_C_S2_A)
    s2_amax = a1(_BETA_R_A_MAX,  _BETA_C_S2_A)
    s2_astep= a1(_BETA_R_A_STEP, _BETA_C_S2_A)
    s2_bmin = a1(_BETA_R_B_MIN,  _BETA_C_S2_B)
    s2_bmax = a1(_BETA_R_B_MAX,  _BETA_C_S2_B)
    s2_bstep= a1(_BETA_R_B_STEP, _BETA_C_S2_B)
    s2_best_a = a1(_BETA_R_BEST, _BETA_C_S2_A)
    s2_best_b = a1(_BETA_R_BEST, _BETA_C_S2_B)

    # ── Title + stage sub-headers ─────────────────────────────────────────────
    val(sheet, r0, c0, "Beta Full_Factorial MLE  —  Alpha (α) × Beta (β); two stages side by side")
    sheet.range(rc(r0, c0), rc(r0, last_col)).merge()
    title_range = sheet.range(rc(r0, c0), rc(r0, last_col))
    title_range.color = _HEADER
    title_range.api.Font.Bold = True

    sub_row = r0 + _R_STAGE_SUBHDR
    val(sheet, sub_row, c0 + _BETA_C_S1_A, "Stage 1 (Wide Scope)")
    val(sheet, sub_row, c0 + _BETA_C_S2_A, "Stage 2 (refined)")
    _subheader_row(sheet, sub_row, c0 + _BETA_C_S1_A, c0 + _BETA_C_S2_A)

    # ── Control field-list (rows 4–12) ────────────────────────────────────────
    # α values in the α cols (S1 col1, S2 col3); β values in the β cols (S1 col2,
    # S2 col4); Min NLL in the NLL cols (S1 col2, S2 col5); Best α/β on row 12.
    fields = [
        (_BETA_R_GRID_POINTS, "Grid Points"),
        (_BETA_R_A_MIN, "α Min"),
        (_BETA_R_A_MAX, "α Max"),
        (_BETA_R_A_STEP, "α Step"),
        (_BETA_R_B_MIN, "β Min"),
        (_BETA_R_B_MAX, "β Max"),
        (_BETA_R_B_STEP, "β Step"),
        (_BETA_R_MINNLL, "Min NLL"),
        (_BETA_R_BEST, "Best (α, β)"),
    ]
    for r_off, label in fields:
        val(sheet, r0 + r_off, c0, label)

    # Stage 1 control values (numeric defaults; Min/Max editable overrides).
    val(sheet, r0 + _BETA_R_GRID_POINTS, c0 + _BETA_C_S1_A, n)
    sheet.range(rc(r0 + _BETA_R_GRID_POINTS, c0 + _BETA_C_S1_A)).number_format = _FMT_INT
    _guard_grid_points_cell(sheet, r0 + _BETA_R_GRID_POINTS, c0 + _BETA_C_S1_A)
    val(sheet, r0 + _BETA_R_A_MIN, c0 + _BETA_C_S1_A, 0.2)
    val(sheet, r0 + _BETA_R_A_MAX, c0 + _BETA_C_S1_A, 50.0)
    val(sheet, r0 + _BETA_R_B_MIN, c0 + _BETA_C_S1_B, 0.2)
    val(sheet, r0 + _BETA_R_B_MAX, c0 + _BETA_C_S1_B, 50.0)
    f(sheet, r0 + _BETA_R_A_STEP, c0 + _BETA_C_S1_A, f"=({s1_amax}-{s1_amin})/({s1_n}-1)")
    f(sheet, r0 + _BETA_R_B_STEP, c0 + _BETA_C_S1_B, f"=({s1_bmax}-{s1_bmin})/({s1_n}-1)")
    for r_off, c_off in (
        (_BETA_R_A_MIN, _BETA_C_S1_A), (_BETA_R_A_MAX, _BETA_C_S1_A),
        (_BETA_R_B_MIN, _BETA_C_S1_B), (_BETA_R_B_MAX, _BETA_C_S1_B),
    ):
        sheet.range(rc(r0 + r_off, c0 + c_off)).color = _INPUT

    # Stage 2 control values — bracket Stage 1's optimum by ±1 step.
    f(sheet, r0 + _BETA_R_GRID_POINTS, c0 + _BETA_C_S2_A, f"={s1_n}")
    sheet.range(rc(r0 + _BETA_R_GRID_POINTS, c0 + _BETA_C_S2_A)).number_format = _FMT_INT
    f(sheet, r0 + _BETA_R_A_MIN, c0 + _BETA_C_S2_A, f"=MAX(0.001,{s1_best_a}-{s1_astep})")
    f(sheet, r0 + _BETA_R_A_MAX, c0 + _BETA_C_S2_A, f"={s1_best_a}+{s1_astep}")
    f(sheet, r0 + _BETA_R_B_MIN, c0 + _BETA_C_S2_B, f"=MAX(0.001,{s1_best_b}-{s1_bstep})")
    f(sheet, r0 + _BETA_R_B_MAX, c0 + _BETA_C_S2_B, f"={s1_best_b}+{s1_bstep}")
    f(sheet, r0 + _BETA_R_A_STEP, c0 + _BETA_C_S2_A, f"=({s2_amax}-{s2_amin})/({s2_n}-1)")
    f(sheet, r0 + _BETA_R_B_STEP, c0 + _BETA_C_S2_B, f"=({s2_bmax}-{s2_bmin})/({s2_n}-1)")

    sheet.range(
        rc(r0 + _BETA_R_A_MIN, c0 + _BETA_C_S1_A),
        rc(r0 + _BETA_R_B_STEP, c0 + _BETA_C_S2_B),
    ).number_format = _FMT_1DP

    # ── Body spills (row 33): Full_Factorial grid + a separate BYROW NLL col ─────
    # The grid is a Full_Factorial(N, mins, maxs) spill → N²×2 (col1 = α slow
    # axis, col2 = β fast axis) across the stage's first two cols (Alpha | Beta).
    # The NLL is a SEPARATE spill in the stage's third col: a BYROW over the grid
    # spill (via the `#` operator) that evaluates NLL_Beta at each (α,β).  Two
    # spills rather than one lets the grid stand alone as the pure Cartesian
    # product and the NLL read it by reference — the same shape the profile fits
    # use for their axis and NLL columns.
    # Stage 1 grid at col0 (BY33), NLL at col2 (CA33); Stage 2 grid at col3 (CB33),
    # NLL at col5 (CD33) — same row, independent height.
    val(sheet, body_hdr_row, c0 + _BETA_C_LABEL, "Alpha (α)")
    val(sheet, body_hdr_row, c0 + _BETA_C_S1_A, "Beta (β)")
    val(sheet, body_hdr_row, c0 + _BETA_C_S1_B, "NLL")
    val(sheet, body_hdr_row, c0 + _BETA_C_S2_A, "Alpha (α)")
    val(sheet, body_hdr_row, c0 + _BETA_C_S2_B, "Beta (β)")
    val(sheet, body_hdr_row, c0 + _BETA_C_S2_NLL, "NLL")
    _subheader_row(sheet, body_hdr_row, c0, last_col)

    def _stage_grid(n_ref: str, amin: str, amax: str, bmin: str, bmax: str) -> str:
        # Full_Factorial grid: N²×2 (Alpha | Beta) spill at the stage's first col.
        return f"=Full_Factorial({n_ref},VSTACK({amin},{bmin}),VSTACK({amax},{bmax}))"

    def _stage_nll(grid_anchor: str) -> str:
        # BYROW NLL over the grid spill — a separate spill in the NLL col.  The
        # scaled sample z is bound once here and reused across every row; the
        # grid_anchor is the absolute ref of the stage's grid spill (`<anchor>#`
        # expands to the full N²×2 grid, so each row ab is an (α,β) pair).
        return (
            "=LET("
            "d,FILTER(UV_Data,UV_Include),"
            "range_,MAX(d)-MIN(d),"
            "pad,range_*0.001,"
            "scale_,range_+2*pad,"
            "z,(d-MIN(d)+pad)/scale_,"
            f"BYROW({grid_anchor}#,LAMBDA(ab,IFERROR("
            "NLL_Beta(z,INDEX(ab,1,1),INDEX(ab,1,2))+COUNT(d)*LN(scale_),1E+15)))"
            ")"
        )

    s1_grid_anchor = a1(_R_BODY, _BETA_C_LABEL)   # $BY$33
    s2_grid_anchor = a1(_R_BODY, _BETA_C_S2_A)     # $CB$33
    f(sheet, body_row, c0 + _BETA_C_LABEL,
      _stage_grid(s1_n, s1_amin, s1_amax, s1_bmin, s1_bmax))
    f(sheet, body_row, c0 + _BETA_C_S1_B,
      _stage_nll(s1_grid_anchor))
    f(sheet, body_row, c0 + _BETA_C_S2_A,
      _stage_grid(s2_n, s2_amin, s2_amax, s2_bmin, s2_bmax))
    f(sheet, body_row, c0 + _BETA_C_S2_NLL,
      _stage_nll(s2_grid_anchor))

    # Number formats + colour scale over the default-size body window.
    cf_row_end = body_row + max(n * n, _BETA_BODY_CF_ROWS_CAP) - 1
    for c_off, fmt in (
        (_BETA_C_LABEL, _FMT_1DP), (_BETA_C_S1_A, _FMT_1DP), (_BETA_C_S1_B, _FMT_SCI_1DP),
        (_BETA_C_S2_A, _FMT_1DP), (_BETA_C_S2_B, _FMT_1DP), (_BETA_C_S2_NLL, _FMT_SCI_1DP),
    ):
        sheet.range(rc(body_row, c0 + c_off), rc(cf_row_end, c0 + c_off)).number_format = fmt

    # ── OFFSET named ranges (height N², tracking the live Grid Points cell) ───
    name_s1_alpha, name_s1_beta, name_s1_nll = "UV_BETA_S1_Alpha", "UV_BETA_S1_Beta", "UV_BETA_S1_NLL"
    name_s2_alpha, name_s2_beta, name_s2_nll = "UV_BETA_S2_Alpha", "UV_BETA_S2_Beta", "UV_BETA_S2_NLL"
    for name, col_off, n_ref in (
        (name_s1_alpha, _BETA_C_LABEL,  s1_n),
        (name_s1_beta,  _BETA_C_S1_A,   s1_n),
        (name_s1_nll,   _BETA_C_S1_B,   s1_n),
        (name_s2_alpha, _BETA_C_S2_A,   s2_n),
        (name_s2_beta,  _BETA_C_S2_B,   s2_n),
        (name_s2_nll,   _BETA_C_S2_NLL, s2_n),
    ):
        cl = col_letter(c0 + col_off)
        _drop_wb_name(sheet, name)
        drop_local_name(sheet, name)
        sheet.api.Names.Add(
            Name=name,
            RefersTo=(
                f"=OFFSET('{sname}'!${cl}${body_row},0,0,"
                f"MAX(IFERROR('{sname}'!{n_ref},1),1)^2,1)"
            ),
        )

    # ── Recovery: Min NLL, Best α, Best β (via Grid_Argument_Minimum) ─────────
    for nll_name, alpha_name, beta_name, minnll_coff, best_a_coff, best_b_coff, n_ref in (
        (name_s1_nll, name_s1_alpha, name_s1_beta, _BETA_C_S1_B, _BETA_C_S1_A, _BETA_C_S1_B, s1_n),
        (name_s2_nll, name_s2_alpha, name_s2_beta, _BETA_C_S2_NLL, _BETA_C_S2_A, _BETA_C_S2_B, s2_n),
    ):
        loc = f"INDEX(Grid_Argument_Minimum({nll_name}),1,2)"
        f(sheet, r0 + _BETA_R_MINNLL, c0 + minnll_coff,
          f'=IFERROR(TAKE(Grid_Argument_Minimum({nll_name}),,1),"—")')
        f(sheet, r0 + _BETA_R_BEST, c0 + best_a_coff,
          f'=IFERROR(INDEX({alpha_name},{loc}),"—")')
        f(sheet, r0 + _BETA_R_BEST, c0 + best_b_coff,
          f'=IFERROR(INDEX({beta_name},{loc}),"—")')
        sheet.range(rc(r0 + _BETA_R_MINNLL, c0 + minnll_coff)).number_format = _FMT_SCI_1DP
        sheet.range(rc(r0 + _BETA_R_BEST, c0 + best_a_coff)).number_format = _FMT_1DP
        sheet.range(rc(r0 + _BETA_R_BEST, c0 + best_b_coff)).number_format = _FMT_1DP
        # Boundary guard on Best α and Best β: red when the optimum is on an
        # edge of the N²-long grid (location 1 or N²).
        for best_coff in (best_a_coff, best_b_coff):
            cell_api = sheet.range(rc(r0 + _BETA_R_BEST, c0 + best_coff)).api
            cf = cell_api.FormatConditions.Add(
                Type=2,  # xlExpression
                Formula1=f"=OR({loc}=1,{loc}={n_ref}^2)",
            )
            cf.Interior.Color = 0x0000FF   # red (BGR)
            cf.Font.Color = 0xFFFFFF       # white

    # NLL colour scale on both NLL columns over the body window.
    for nll_coff in (_BETA_C_S1_B, _BETA_C_S2_NLL):
        try:
            rng = sheet.range(rc(body_row, c0 + nll_coff), rc(cf_row_end, c0 + nll_coff))
            rng.api.FormatConditions.Delete()
            cs = rng.api.FormatConditions.AddColorScale(3)
            cs.ColorScaleCriteria(1).Type = 1
            cs.ColorScaleCriteria(1).FormatColor.Color = 0x63BE7B
            cs.ColorScaleCriteria(2).Type = 5
            cs.ColorScaleCriteria(2).Value = 50
            cs.ColorScaleCriteria(2).FormatColor.Color = 0xFFEB84
            cs.ColorScaleCriteria(3).Type = 2
            cs.ColorScaleCriteria(3).FormatColor.Color = 0xF8696B
        except Exception:
            pass

    border_box(sheet, r0 + _BETA_R_GRID_POINTS, c0, r0 + _BETA_R_BEST, c0 + _BETA_C_S2_B)

    return {
        "s1": {"best_a": s1_best_a, "best_b": s1_best_b, "n_points": s1_n},
        "s2": {"best_a": s2_best_a, "best_b": s2_best_b, "n_points": s2_n},
    }


# ``nll_formula`` takes the sample expression as its first argument and omits the
# catalog LAMBDA's optional ``[filter]``: the profile bodies pass an already-
# filtered sample, so re-applying a mask would only cost a second pass.
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


def _write_two_parameter_grid_search(sheet: xw.Sheet, beta_grid_size: int = _N_GRID) -> None:
    """Write the two-stage MLE search blocks for all two-parameter fits.

    Weibull and Gamma search one dimension: their scale / rate parameter is
    profiled out in closed form, so each stage is an N-point profile-NLL curve
    (``_write_profile_fit``).  Beta stays two-dimensional — both of its
    conditional MLEs involve digamma — and uses an N² ``Full_Factorial`` grid
    per stage (``_write_beta_fit``).  Every stage is a grid spill beside a
    ``BYROW`` NLL column that reads it, sized live by a Grid Points cell.
    """
    for spec in _PROFILE_SEARCHES:
        _write_profile_fit(sheet, **spec)

    _write_beta_fit(sheet, beta_grid_size=beta_grid_size)


def _write_weibull_grid_search(sheet: xw.Sheet, beta_grid_size: int = _N_GRID) -> None:
    """Compatibility wrapper for the two-parameter grid-search section."""
    _write_two_parameter_grid_search(sheet, beta_grid_size=beta_grid_size)


def _write_profile_charts(sheet: xw.Sheet) -> None:
    """Insert the Weibull and Gamma profile-NLL line charts.

    One chart per one-dimensional fit, anchored ABOVE its fit zone's body — in
    the band between the control block (rows 4–11) and the body header (row 32),
    one zone wide — so the curve and the body column it is drawn from read
    together.  (Beta reserves the same band, BY13:CD30, blank for a potential
    future chart.)

    This is what replaces the 2-D NLL heatmap those two distributions used to
    carry: a 1-D search would reduce the heatmap to a single colour strip,
    whereas the curve shows the basin, the interior minimum, and a boundary hit
    directly.

    Each chart carries **both stages**. Stage 1 is the wide bracket that locates
    the basin; Stage 2 re-samples 20 points across ±1 Stage 1 step around the
    winner and is what actually fixes the reported parameter. Plotting both, with
    ``+`` markers on Stage 2, shows where the search spent its resolution — the
    refined region is otherwise invisible, being a couple of pixels of the
    Stage 1 curve.
    """
    sname = sheet.name
    for spec in _PROFILE_SEARCHES:
        zone_col = spec["col_start"]
        chart_range = sheet.range(
            rc(_ROW_PROFILE_CHART, zone_col),
            rc(_ROW_PROFILE_CHART + _PROFILE_CHART_ROWS - 1, zone_col + _PR_W - 1),
        )
        co = sheet.api.ChartObjects().Add(
            chart_range.left, chart_range.top, chart_range.width, chart_range.height
        )
        chart = co.Chart

        while chart.SeriesCollection().Count > 0:
            chart.SeriesCollection(1).Delete()

        chart.ChartType = _XL_XY_SCATTER_LINES

        stem = spec["body_prefix"].removeprefix("UV_")
        stage1 = chart.SeriesCollection().NewSeries()
        stage1.XValues = f"='{sname}'!UV_Profile_{stem}_S1_Axis"
        stage1.Values = f"='{sname}'!UV_Profile_{stem}_S1_NLL"
        stage1.Name = f"{spec['dist_name']} profile NLL"
        stage1.MarkerSize = 4

        # Stage 2 over the same axes, marked with + so the refined region is
        # distinguishable where it overlaps the Stage 1 curve.
        stage2 = chart.SeriesCollection().NewSeries()
        stage2.XValues = f"='{sname}'!UV_Profile_{stem}_S2_Axis"
        stage2.Values = f"='{sname}'!UV_Profile_{stem}_S2_NLL"
        stage2.Name = f"{spec['dist_name']} profile NLL — refined"
        stage2.MarkerStyle = _XL_MARKER_PLUS
        stage2.MarkerSize = 5

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
    for spec in _PROFILE_SEARCHES:
        zone_col = spec["col_start"]
        note_cells.append(
            (_ROW_FIT_ZONE + _PR_R_START, zone_col + _PR_C_LABEL, "Start")
        )
        note_cells.append(
            (_ROW_FIT_ZONE + _R_BODY_HDR, zone_col + _PR_C_S1, "Profile NLL")
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
    beta_grid_size: int = _N_GRID,
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
    _write_fitting_table(sheet, beta_grid_size=beta_grid_size)
    _write_qq_data(sheet)

    _write_weibull_grid_search(sheet, beta_grid_size=beta_grid_size)
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

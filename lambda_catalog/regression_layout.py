"""
regression_layout.py

Layout constants for the Regression sheet -- column indices, row anchors,
zone/gap/width tables, chart constants, materialization zone constants, and
the model-context element table.

These constants are the shared contract between the sheet writer
(write_sheet_regression.py), the spec-block QC analyzers, the test-model
sheet I/O, and the build scripts. Every A1 address quoted in a formula is
derived from these constants via _abs_ref / _band, never spelled out.

Extracted from write_sheet_regression.py to make the layout contract an
explicit, discoverable module rather than something buried at the top of a
3,200-line writer file.
"""
from __future__ import annotations

from pathlib import Path
from typing import NamedTuple

from .workbook_helpers import MAX_EXCEL_ROW, col_letter
from .spec_layout import (
    _C_DESIGN_COLUMNS as _C_SPEC_DESIGN_COLUMNS,
)
from .spec_layout import (
    _C_SEQUENCE_PERIOD as _C_SPEC_SEQUENCE_PERIOD,
)
from .spec_layout import (
    _DEFAULT_TRANSFORM,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    _TRANSFORM_LOG,
    _TRANSFORM_LOG_DROP,
)

REGRESSION_SHEET_NAME = "Regression"

# The catalog file backing this sheet's constructor closures (scope
# "Regression"). Used when a caller does not pass them in explicitly.
_DEFINITIONS_PATH = Path(__file__).resolve().parent.parent / "lambda_functions.json"

# ── 1-based column indices ─────────────────────────────────────────────────────
# Every constant matches its actual column letter (e.g. _C_K is column K).
#
# ── Column-layout paradigm ─────────────────────────────────────────────────────
# The sheet is a strip of content zones. Between any two adjacent zones sits
# ONE dedicated GAP column (width 2) that is deliberately left OUT of every
# outline group. That ungrouped column is what makes the two zones collapse
# independently: Excel merges a contiguous run of same-level grouped columns
# into a single outline, so two zones with no ungrouped column between them
# share one collapse control. The gap columns are R, Z, AI, AM — one before
# each zone that follows the spec block.
#
#   A–O   Model Specification   | P, Q spec feedback + I Verdict overlay
#   | R gap | S–Y Predictor Summary
#   | Z gap | AA–AH Regression Outputs | AI gap | AJ–AL Prediction Outputs
#   | AM gap | AN–AY Residual Output
#
# The layout-break MAJOR appended three columns to the spec block — M
# (Interaction Term), N (Interaction Operation), and O (the Design Columns
# audit) — so every zone after it moved three columns right. That shift is
# the break: A–L keep their letters and meanings, everything past them does
# not.
#
# The gap columns and the per-zone content spans below are the single source
# of truth for both column widths and outline grouping (see _ZONES / _GAP_COLUMNS).

# Zone 1: Model Specification — columns A–O are owned by the shared spec-block
# writers in spec_layout (imported above); only the section
# heading cell is written here.
_C_A = 1    # spec: Variable labels / A1 zone heading / A2 Intercept label

# Spec feedback (P, Q, plus the I Verdict overlay): the delta spectrum
# (Sequence_Delta_Spectrum() spill at P2:Q?) sits in P and Q. The combined
# verdict switch overlays column I (the Sequence_Period spec column on
# the spec data rows; the row-1/row-2 cells are above the spec table and
# free). I1 holds the "Verdict" header (bold), I2 the priority-ordered
# switch formula (off-grid outranks regularity, both outrank no-natural
# and calendar; red CF outranks yellow via StopIfTrue). The E1 cell
# carries the multi-flag Sequence error status (moved here from H2 when the
# spec data area became a structured table — that table is gone, but the
# status stayed put). The design-matrix width guard writes M2 (status) and
# N1/O1 (label/total).
_C_P = 16   # spec feedback: Δ header / spectrum spill
_C_Q = 17   # spec feedback: Count header / spectrum spill

# Gap before Predictor Summary.
_C_R = 18   # thin gap (ungrouped — splits the spec and predictor outlines)

# Zone 2: Predictor Summary (constructed columns)
_C_S = 19   # constructed column names (level-qualified)
_C_T = 20   # Pearson R
_C_U = 21   # Spearman R
_C_V = 22   # Skewness
_C_W = 23   # Kurtosis
_C_X = 24   # GVIF
_C_Y = 25   # Tolerance

# Gap before Regression Outputs.
_C_Z = 26   # thin gap (ungrouped)

# Zone 3: Regression Outputs
_C_AA = 27  # labels (stats / ANOVA / coefficients)
_C_AB = 28  # stat values / ANOVA df / coefficient values
_C_AC = 29  # ANOVA SS / coefficient SE
_C_AD = 30  # diagnostics labels / ANOVA MS / coefficient t-stat
_C_AE = 31  # Predicted Variable label (AE2) / diagnostics values / ANOVA F / coefficient p-value
_C_AF = 32  # predicted variable readout (AF2) / ANOVA Sig F / coefficient CI lower
_C_AG = 33  # coefficient CI upper
_C_AH = 34  # Beta Weights

# Gap before Prediction Outputs.
_C_AI = 35  # thin gap (ungrouped)

# Zone 4: Prediction Outputs
_C_AJ = 36  # prediction interval labels / prediction input labels
_C_AK = 37  # prediction interval values / prediction input values
_C_AL = 38  # Training Mean — the per-constructed-column means spill (AL19).
            # The spill owns column AL downward, so it can never collide with
            # another spill when the source data or spec changes.

# Gap before Residual Output.
_C_AM = 39  # thin gap (ungrouped)

# Zone 5: Residual Output
# _C_AN holds Row_Labels() (the identifiers spill); _C_AO onward hold the
# residual-diagnostic columns. Row-2 headers are written by _write_residuals
# (see the note_cells table); row-3 formulas are the source of truth for what
# each column actually contains.
_C_AN = 40  # row identifiers (Row_Labels() spill)
_C_AO = 41  # Y (actual dependent variable)
_C_AP = 42  # Predicted Y
_C_AQ = 43  # Residuals
_C_AR = 44  # Hat Diagonal
_C_AS = 45  # Studentized Residuals
_C_AT = 46  # Cook's Distance
_C_AU = 47  # Normal Scores Ranked
_C_AV = 48  # Studentized Residuals Ranked
_C_AW = 49  # Scale-Location
_C_AX = 50  # PRESS Residual
_C_AY = 51  # Cook's Distance (Flagged) — chart data-label helper column
_C_AZ = 52  # Predicted Y (Original Units) — Residual Output zone (v3.3 unit-space)
_C_BA = 53  # Residual (Original Units) — Residual Output zone (v3.3 unit-space)
_C_BB = 54  # chart anchor — formerly _C_AZ; everything past it shifts right by 2

# The constructed-column count is spec-dependent (19 on the default WHO spec),
# so bands that v1 sized with the fixed k=18 now cover a generous fixed range.
_PRED_INPUT_FIRST_ROW = 19
_PRED_INPUT_LAST_ROW = 62
_FORMAT_BAND_LAST_ROW = 62

# ── Cell anchors referenced from formulas ─────────────────────────────────────
# Conditional-format expressions, chart titles, and OFFSET-based named ranges
# all need A1 addresses as literal text inside a formula string. Spelling those
# letters out by hand is what makes a column insertion a silent-wrong-answer
# bug rather than a build failure: the formula still parses, it just reads a
# different cell. Every such address is therefore BUILT from the _C_* constants
# and the row anchors below, so the layout constants stay the single source of
# truth for formulas exactly as they already are for widths and outlines.
_ROW_DATA_FIRST = 3            # first spilled data row in every full-height zone
_ROW_ADJUSTED_R_SQUARED = 6
_ROW_STANDARD_ERROR = 7
_ROW_OBSERVATIONS = 8
_ROW_PRESS = 4
_ROW_PRESS_R_SQUARED = 5
_ROW_MEAN_LEVERAGE = 6
_ROW_QQ_CORRELATION = 10
_ROW_ALPHA = 12                # confidence-level input
_ROW_SIGNIFICANCE_F = 15
_ROW_ANOVA_RESIDUAL_DF = 16    # ANOVA table: the Residual row's df cell
_ROW_COEFF_FIRST = 21
_ROW_RESPONSE_READOUT = 2      # Predicted Variable readout
_ROW_FE_GROUP = 12             # prediction-zone FE group selector


def _abs_ref(row: int, col: int) -> str:
    """``$D$4``-style absolute address for embedding in a formula string."""
    return f"${col_letter(col)}${row}"


def _band(col: int, first_row: int = _ROW_DATA_FIRST) -> str:
    """``D3:D1048576``-style full-height band for a conditional-format range."""
    letter = col_letter(col)
    return f"{letter}{first_row}:{letter}{MAX_EXCEL_ROW}"


# Statistics / diagnostics cells other formulas key on.
_A_ALPHA = _abs_ref(_ROW_ALPHA, _C_AB)
_A_OBSERVATIONS = _abs_ref(_ROW_OBSERVATIONS, _C_AB)
_A_STANDARD_ERROR = _abs_ref(_ROW_STANDARD_ERROR, _C_AB)
_A_ADJUSTED_R_SQUARED = _abs_ref(_ROW_ADJUSTED_R_SQUARED, _C_AB)
_A_PRESS = _abs_ref(_ROW_PRESS, _C_AE)
_A_PRESS_R_SQUARED = _abs_ref(_ROW_PRESS_R_SQUARED, _C_AE)
_A_MEAN_LEVERAGE = _abs_ref(_ROW_MEAN_LEVERAGE, _C_AE)
_A_QQ_CORRELATION = _abs_ref(_ROW_QQ_CORRELATION, _C_AE)
_A_SIGNIFICANCE_F = _abs_ref(_ROW_SIGNIFICANCE_F, _C_AF)
_A_RESPONSE_READOUT = _abs_ref(_ROW_RESPONSE_READOUT, _C_AF)
_A_FE_GROUP = _abs_ref(_ROW_FE_GROUP, _C_AK)
# The ANOVA Residual df cell (n − p − absorbed) and the spec block's Σ Design
# Columns total (p, intercept column included). Together they are the reference
# F distribution for Cook's Distance — see _COOKS_CUTOFF below.
_A_RESIDUAL_DF = _abs_ref(_ROW_ANOVA_RESIDUAL_DF, _C_AB)
_A_DESIGN_COLUMNS_TOTAL = _abs_ref(1, _C_SPEC_DESIGN_COLUMNS)

# ── The Cook's Distance influence cutoff ──────────────────────────────────────
# D_i > F(0.5, p, n−p): the median of the reference F distribution, which is the
# standard rule and the one this sheet screens on. Written once and reused by the
# AT/AY conditional formats, the AY flag column, and the chart title, so the three
# can never disagree about what "flagged" means.
#
# The numerator df is p — the design matrix's COLUMN COUNT, intercept included —
# not the ANOVA Regression df at $AB$15, which is p−1. Cooks_Distance divides by
# COLUMNS(Design_Matrix(X, Include)), so p is what makes the statistic and its
# reference distribution the same quantity. $O$1 already totals exactly that
# (Σ of the spec's Design Columns audit plus N(Allow_Intercept)).
#
# The IFERROR is load-bearing, not defensive habit: under Zero_Predictors_Selected()
# the design collapses, F.INV sees a zero df and returns #NUM!. NA() makes every
# comparison against the cutoff fail closed — nothing flagged — where a raw #NUM!
# would propagate into the flag column and light the whole band.
#
# F.INV needs no _xlfn. prefix here. That prefix belongs to catalog formulas
# written into lambda_functions.json (see lambda_formula_parser); a formula
# handed to Formula2 goes through Excel's own parser, exactly like the T.INV.2T
# calls in the coefficient block below.
_COOKS_CUTOFF = f"IFERROR(F.INV(0.5,{_A_DESIGN_COLUMNS_TOTAL},{_A_RESIDUAL_DF}),NA())"
# Unit-space block anchors (v3.3). The Method cell lives at row 4 of the
# block — sibling to the section heading at row 3 — so the rest of the block
# (rows 5–9) and the prediction column (AL) can reference a single source.
_A_BACK_TRANSFORM_METHOD = _abs_ref(4, _C_AH)
# The two back-transform methods, and the default written into AH4 as a
# LITERAL. The cell is an input: it must never hold a formula that reads its
# own address (a circular reference), and the validation list below is what
# constrains a typed value.
_BACK_TRANSFORM_METHODS = ("Duan", "Naive")
_BACK_TRANSFORM_DEFAULT = _BACK_TRANSFORM_METHODS[0]

# Content zones as (first_col, last_col) spans — the single source of truth for
# the outline groups. Each pair becomes one collapsible column group; the gap
# columns between them (below) stay ungrouped so the zones collapse
# independently. Zone 1 includes the P/Q spec feedback columns and the
# column-I Verdict overlay so the spec block and its feedback share one
# outline (one click to collapse the spec block together — a spec edit
# doesn't need the feedback visible).
_ZONES: tuple[tuple[int, int], ...] = (
    (_C_A, _C_Q),                 # A:Q   — Model Specification + feedback
    (_C_S, _C_Y),                 # S:Y   — Predictor Summary
    (_C_AA, _C_AH),               # AA:AH — Regression Outputs
    (_C_AJ, _C_AL),               # AJ:AL — Prediction Outputs
    (_C_AN, _C_BA),               # AN:BA — Residual Output (was AN:AY; v3.3 added AZ/BA)
)

# The ungrouped gap columns (width 2) that separate the zones above. Derived as
# the single column between each pair of adjacent zones; asserted to be exactly
# one column wide so a future zone edit that closes or widens a gap fails loudly.
_GAP_COLUMNS: tuple[int, ...] = tuple(
    _ZONES[i][1] + 1 for i in range(len(_ZONES) - 1)
)
assert all(
    _ZONES[i + 1][0] - _ZONES[i][1] == 2 for i in range(len(_ZONES) - 1)
), "each zone must be separated from the next by exactly one gap column"

# Column outline groups, one per zone — the thin gap columns stay ungrouped
# so each zone collapses independently. Cleared and re-applied on every
# rebuild (Cells.Clear does NOT remove outline levels, so re-grouping
# without ClearOutline would deepen the outline each build).
_COLUMN_GROUPS: tuple[tuple[int, int], ...] = _ZONES


# ── Content-column widths ─────────────────────────────────────────────────────
# One entry per content column, KEYED ON THE LAYOUT CONSTANT, never on a
# hard-coded letter. This table used to be a dict of literal letters, and the
# layout-break MAJOR that shifted every zone right of the spec block three
# columns over did not move its keys: the Predictor Summary's name column got
# the width meant for a stats column, the Regression Outputs' diagnostics
# labels got the width meant for its values, and the whole Prediction Outputs
# zone (AJ–AL) fell off the table and rendered at Excel's default 8.43. A
# letter key is exactly the silent-wrong-answer failure the `_C_*` constants
# exist to prevent — the build still succeeds, it just sizes a different
# column — so the widths now derive from the same constants the zones do, and
# the coverage assertion below fails the next shift loudly.
#
# Widths are sized to the widest label/value each column actually carries;
# row-2 headers wrap (see write_regression_output_sheet), so a header longer
# than its column costs height rather than truncation.
_COLUMN_WIDTHS: tuple[tuple[int, float], ...] = (
    # ── Spec block (A–O) ────────────────────────────────────────────────────
    # The spec block owns its own widths (_set_spec_block_column_widths in
    # ``spec_layout``). Column I is the ONE
    # Regression-only override: this sheet overlays the combined Sequence
    # Verdict switch on I2, a long message and the widest cell on the sheet.
    (_C_SPEC_SEQUENCE_PERIOD, 38),
    # ── Zone 1 feedback (P, Q) — the Δ spectrum spill ───────────────────────
    (_C_P, 10),        # Δ header / spectrum column 1
    (_C_Q, 8),         # Count header / spectrum column 2
    # ── Zone 2: Predictor Summary (S–Y) ─────────────────────────────────────
    (_C_S, 24),        # constructed column names (e.g. "Status[Developed]")
    (_C_T, 9),         # Pearson R
    (_C_U, 9),         # Spearman R
    (_C_V, 9),         # Skewness
    (_C_W, 9),         # Kurtosis
    (_C_X, 9),         # GVIF
    (_C_Y, 9),         # Tolerance
    # ── Zone 3: Regression Outputs (AA–AH) ──────────────────────────────────
    # Column roles vary by sub-table (statistics, ANOVA, coefficients,
    # diagnostics, unit-space fit); each width is sized for the widest
    # label/value used anywhere in that column.
    (_C_AA, 22),       # labels — "Adjusted R Square" (17), "Status[Developing]" (18)
    (_C_AB, 12),       # stat values, Alpha, ANOVA df, Coefficients
    (_C_AC, 12),       # ANOVA SS, Std Error
    (_C_AD, 24),       # diagnostics labels ("BFN Panel Durbin-Watson" = 23) + MS + t Stat
    (_C_AE, 16),       # diagnostics values + "Predicted Variable" heading + F + P-value
    (_C_AF, 14),       # response-name readout + Significance F + Lower 95%
    (_C_AG, 16),       # Upper 95% + unit-space labels ("Adj R Square (Unit)" = 19)
    (_C_AH, 16),       # Beta Weight + unit-space values + the Duan/Naive toggle
    # ── Zone 4: Prediction Outputs (AJ–AL) ──────────────────────────────────
    (_C_AJ, 24),       # "PREDICTION INTERVAL"/"PREDICTION INPUTS" labels + names spill
    (_C_AK, 16),       # Fit Space interval values + prediction input values
    (_C_AL, 14),       # Original Units + Training Mean values
    # ── Zone 5: Residual Output (AN–BA) ─────────────────────────────────────
    # AN holds Row_Labels() — identifier strings like "United States" (13).
    (_C_AN, 16),       # row identifiers
    (_C_AO, 9),        # Y
    (_C_AP, 9),        # Predicted Y
    (_C_AQ, 12),       # Residuals
    (_C_AR, 9),        # Hat Diagonal
    (_C_AS, 14),       # Studentized Residuals
    (_C_AT, 17),       # Cook's Distance
    (_C_AU, 14),       # Normal Scores Ranked
    (_C_AV, 12),       # Studentized Residuals Ranked
    (_C_AW, 10),       # Scale-Location
    (_C_AX, 12),       # PRESS Residual
    (_C_AY, 12),       # Cook's Distance (Flagged) — chart data-label helper column
    (_C_AZ, 14),       # Predicted Y (Original Units) — v3.3
    (_C_BA, 14),       # Residual (Original Units) — v3.3
    # ── Post-zone gutter ────────────────────────────────────────────────────
    # BB is NOT a content column and NOT a zone gap — it is the gutter that
    # bounds the row-2 header wrap (last content column = BA) and anchors the
    # diagnostic charts. Sized here so it reads as a deliberate margin rather
    # than a default-width column.
    (_C_BB, 15),
)

# Every content column in every zone gets exactly one width, and no width is
# assigned to a gap column. This is the guard the letter-keyed dict did not
# have: after the next zone shift, a stale entry lands outside its zone (or a
# zone column loses its entry) and the build fails here instead of shipping a
# sheet whose columns are sized for the previous layout.
_WIDTH_COLUMNS: tuple[int, ...] = tuple(col for col, _ in _COLUMN_WIDTHS)
assert len(set(_WIDTH_COLUMNS)) == len(_WIDTH_COLUMNS), (
    "a column may be assigned a width only once"
)
assert not (set(_WIDTH_COLUMNS) & set(_GAP_COLUMNS)), (
    "gap columns are sized by the _GAP_COLUMNS loop, not by _COLUMN_WIDTHS"
)
assert set(_WIDTH_COLUMNS) >= {
    col
    for first, last in _ZONES[1:]
    for col in range(first, last + 1)
} | {_C_P, _C_Q}, "every content column outside the spec block needs an explicit width"


# ── Chart constants ───────────────────────────────────────────────────────────
_XL_XY_SCATTER = -4169       # Excel xlXYScatter
_XL_XY_SCATTER_LINES_NO_MARKERS = 75  # Excel xlXYScatterLinesNoMarkers
_XL_COLUMN_CLUSTERED = 51    # Excel xlColumnClustered
_XL_LINE = 4                 # Excel xlLine
_XL_CATEGORY = 1             # horizontal axis
_XL_VALUE = 2                # vertical axis
# Office msoChartFieldRange — the "Value From Cells" data-label field, inserted
# into a DataLabels TextRange via InsertChartField. See the Cook's Distance
# branch of _write_diagnostic_charts.
_MSO_CHART_FIELD_RANGE = 7
_CHART_WIDTH = 310.0         # points
_CHART_HEIGHT = 310.0        # points
_CHART_GAP = 10.0            # gap between charts in points

# Y-axis tick-label number format, per chart. Keyed by the chart_specs key,
# with _CHART_Y_TICK_FORMAT_DEFAULT for anything not listed — declarative
# rather than a chain of `if key ==` overrides, so the format a chart gets is
# readable in one place and pinnable by a unit test (the chart writer itself
# is COM-only and cannot run headless).
#
# The default "0" suits the charts whose y values span the response's own
# scale. It is wrong for any chart whose y values live in a narrow band near
# zero, where integer ticks collapse the axis to two or three labels:
#
#   Scale-Location   sqrt(|studentized residual|), almost always inside 0-2 —
#                    "0" renders 0/1/2 and throws away the spread the chart
#                    exists to show. One decimal is the right resolution: the
#                    diagnostic reads the TREND in that spread, not exact
#                    values, and the CF thresholds on the source column sit at
#                    1.414 / 1.732.
#   Cook's Distance  typically 1e-4 to 1e-1 and heavily right-skewed, so it
#                    takes scientific notation instead.
_CHART_Y_TICK_FORMAT_DEFAULT = "0"
_CHART_Y_TICK_FORMATS: dict[str, str] = {
    "Scale-Location": "0.0",
    "Cook's Distance": "0.0E+00",
}

# Chart label formula cells — one row per diagnostic chart, well below the
# 2-col x 4-row chart grid's pixel footprint (row_step=320pt starting at row
# 3, ~85 rows at default row height) so nothing ever renders on top of them.
# Columns sit past _C_BB (the chart anchor). v3.3 shifted BB to BB+14 (= 68)
# to keep the chart anchor letter stable after the AZ/BA unit-space columns
# replaced the pre-v3.3 AZ gutter.
_C_CHART_LABEL_NAME = _C_BB + 1   # BC — human-readable chart name (doc only)
_C_CHART_TITLE = _C_BB + 2        # BD — Chart Title formula
_C_CHART_XLABEL = _C_BB + 3       # BE — X-Axis Title formula
_C_CHART_YLABEL = _C_BB + 4       # BF — Y-Axis Title formula
_ROW_CHART_LABELS = 95     # first of 7 rows, one per chart in chart_specs order

# ── §4b materialization zone ──────────────────────────────────────────────────
# Per ARCHITECTURE §4b, the Regression sheet carries a band of materialized
# artifacts at its far right — values computed once into a spill range, read by
# the formulas that would otherwise recompute them. Excel does not memoize a name
# whose RefersTo is a formula, so a constructor called inside ~30 engine
# functions runs ~30 times; a materialized cell runs once.
#
# Zone order (increasing width, terminating in the unbounded zone):
#   charts | gutter | Model Context | gutter | Sample_Include | gutter | matrix →
#           (label + value, 4 rows)       (n x 1)            (n x k, unbounded)
#
# Only the Model Context zone is grouped, and it is the only one that ships
# COLLAPSED. It is a bounded, fixed-height block of individual cells — nothing
# spills into it — so hiding it is free, and the far-right band is secondary
# reading surface that should open out of the way.
#
# The two DATA-DEPENDENT zones (the Sample_Include row mask and the terminal
# Constructed Design Matrix) are deliberately left UNGROUPED and visible. Both
# hold full-height dynamic-array spills, and a collapsed outline group over a
# spill range is the configuration in which Excel fails to recompute the model:
# the hidden columns leave the spilled arrays stale, and every engine reading
# across them fits on the stale values. A scrolling nuisance is the accepted
# cost — correct recalculation is not negotiable, and the §4b ordering rule
# (nothing right of the design matrix) is what keeps an expanded terminal zone
# from displacing anything.
#
# Gutters remain width-2 ungrouped separators — the first (after the charts) is
# structural, keeping the floating chart anchors out of the collapsible group.
#
# The chart footprint needs an explicit bound. _C_BB is the chart ANCHOR, not
# its extent: the seven diagnostic charts are floating objects tiled in a
# _CHART_GRID_COLS x _CHART_GRID_ROWS grid, whose right edge sits
# _CHART_RIGHT_OFFSET_PT points past BB's left edge. _LAST_CHART_COLUMN is a
# conservative column index past which that footprint is clear, so the
# full-height materialization spills are never drawn under a chart. A guarded
# build-time assertion verifies the column past the footprint actually clears
# the computed chart right edge — without it a chart resize silently overlaps
# the context block, and the zone start column cannot be trusted.
_CHART_GRID_COLS = 2
_CHART_GRID_ROWS = 4
_CHART_RIGHT_OFFSET_PT = (
    (_CHART_GRID_COLS - 1) * (_CHART_WIDTH + _CHART_GAP) + _CHART_WIDTH
)
# Conservative clear-past-the-footprint bound; asserted against the measured
# chart right edge in _write_materialization_zone. Tracks the chart anchor, so
# a zone shift moves it automatically instead of silently under-reserving.
# v3.3: BA absorbed two Residual-Output content columns (Predicted Y /
# Residual Original Units), so BB=54 + 14 = 68 keeps the same BP
# column-letter end value the chart footprint was sized against.
_LAST_CHART_COLUMN = _C_BB + 14   # BP

# Bounded materialization columns + their ungrouped gutters, then the terminal
# Constructed Design Matrix, which runs unbounded to the sheet's right edge.
# Nothing may ever sit right of it (ARCHITECTURE §4b ordering rule): its width
# is one dropdown away from 156 columns and interactions multiply, so any zone
# placed after it would be displaced by an ordinary modeling choice.
_C_GUTTER_AFTER_CHARTS = _LAST_CHART_COLUMN + 1                      # structural
_C_MODEL_CONTEXT_LABEL = _C_GUTTER_AFTER_CHARTS + 1                  # element labels
_C_MODEL_CONTEXT = _C_MODEL_CONTEXT_LABEL + 1                        # element values
_C_GUTTER_AFTER_CONTEXT = _C_MODEL_CONTEXT + 1
_C_SAMPLE_INCLUDE_MATERIALIZED = _C_GUTTER_AFTER_CONTEXT + 1         # n x 1 spill
_C_GUTTER_AFTER_SAMPLE_INCLUDE = _C_SAMPLE_INCLUDE_MATERIALIZED + 1
_C_DESIGN_MATRIX = _C_GUTTER_AFTER_SAMPLE_INCLUDE + 1                # n x k, terminal
# Header cell for the constructed column names, one column right of the
# matrix anchor — the anchor's own header cell names the intercept column.
_C_DESIGN_MATRIX_NAMES = _C_DESIGN_MATRIX + 1

# The band's first occupied row, under the row-1 zone headings. What that row
# HOLDS differs by zone, because the zones are different kinds of thing:
#
#   Model Context      fixed-height label/value table, no column header, so
#                      this is its first ELEMENT row (rows continue down
#                      through _MODEL_CONTEXT_LAST_ROW)
#   Sample_Include     column-header row, spilling from the row beneath
#   Design Matrix      column-header row, spilling from the row beneath
#
# so this is deliberately NOT "the first data row" of every zone.
_MATERIALIZATION_FIRST_ROW = 2

# The read-across contract lives on the SPILL row, not the header row: the two
# data-dependent zones spill full-height and row-aligned with the source table,
# so the mask value sits beside its own design-matrix row with the gutters as
# visual separators. Both zones must therefore start their spill on the same
# row — asserted in the build.
_MATERIALIZATION_HEADER_ROW = _MATERIALIZATION_FIRST_ROW
_MATERIALIZATION_SPILL_ROW = _MATERIALIZATION_HEADER_ROW + 1

# Header text over the materialized row mask.
_SAMPLE_INCLUDE_HEADER = "In Sample"

# The design matrix's header row is split across two cells because
# Design_Columns() is one column WIDER than Constructed_Column_Names() when
# the intercept is on: the constructor prepends the ones column (see the
# HSTACK(ones, demeaned) branch of Design_Columns), and the names closure
# describes the constructed predictor columns only — the same asymmetry the
# coefficients table resolves with VSTACK("Intercept",...). So the first
# design-matrix column gets its own header cell and the names spill starts one
# column right of it.
_DESIGN_MATRIX_INTERCEPT_HEADER = '=IF(Allow_Intercept,"Intercept","")'


# ── The model context block ───────────────────────────────────────────────────
# The context is a FIXED-SIZE table, one row per element, written as INDEPENDENT
# CELLS — deliberately not a single VSTACK spill.
#
# A spill buys nothing here (the height is a build-time constant, not data-
# dependent) and costs correctness: one formula producing four cells means the
# whole block is a single dependency node that Excel must vacate and re-spill
# whenever the spec block changes. During that re-spill the four cells backing
# Fit_Context's fixed range are transiently blank/#SPILL!, and the ~30 engine
# call sites reading that range through Fit_Context() see the torn value — the
# race this decomposition removes. Four independent cells recalculate
# independently and are never vacated, so the range always holds four values.
#
# It also makes the block readable: each element gets its own labelled row
# instead of four anonymous spilled cells, so the sheet shows what it computed.
#
# Both Log tokens report here as plain "Log". The context feeds the unit-space
# dispatcher, which keys on the (response_transform, predictor_transform) PAIR;
# the two tokens produce the identical Ln(x) column and differ only in the row
# mask, so telling them apart here would double the dispatcher's axis to
# distinguish two cases with the same back-transformation. See
# spec_layout._TRANSFORM_LOG_DROP.
_RESPONSE_TRANSFORM_FORMULA = (
    "IFERROR(IF(INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data)),"
    f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data))))'
    f'="{_TRANSFORM_LOG_DROP}","{_TRANSFORM_LOG}",'
    "INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data)),"
    f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data))))),"None")'
)
# None/Log/Mixed over the INCLUDED CONTINUOUS predictors — masked to Continuous
# so a Categorical dummy's transform never yields a false "Mixed".
_PREDICTOR_TRANSFORM_FORMULA = (
    "LET(n_c,COLUMNS(Source_Data),"
    "rl,TAKE(Spec_Role,n_c),"
    "inc,TAKE(Spec_Include,n_c),"
    "typ,TAKE(Spec_Type,n_c),"
    "trn,TAKE(Spec_Transform,n_c),"
    f'mask,(rl="{_ROLE_PREDICTOR}")*(inc=TRUE)*(typ="Continuous"),'
    # N() of an OR() would collapse the per-column array to one scalar, so the
    # two tokens are summed as indicators and clamped instead.
    f'nL,SUMPRODUCT(mask*(((trn="{_TRANSFORM_LOG}")'
    f'+(trn="{_TRANSFORM_LOG_DROP}"))>0)),'
    f'nN,SUMPRODUCT(mask*N(trn="{_DEFAULT_TRANSFORM}")),'
    'IF(nL=0,"None",IF(nN=0,"Log","Mixed")))'
)


class ModelContextElement(NamedTuple):
    """One row of the materialized model context.

    ``contract_name`` is the versioned public-contract element name that the
    workbook-scoped ``Context_*`` accessors index by position — append-only,
    never insert. ``label`` is what the sheet's label column displays, and
    ``formula`` is the expression materialized into the value cell. Keeping
    all three in one record is what stops the displayed label from drifting
    away from the value beside it.
    """

    contract_name: str
    label: str
    formula: str


# Elements 1-2 (the C2 Allow_Intercept toggle and the
# Absorbed_Degrees_Of_Freedom() closure) feed today's engines; elements 3-4
# (the spec-block transform summaries) have no engine reader until the v3.3
# unit-space dispatcher but land now so the row order is fixed.
# Header note for the Prediction Inputs band. Interaction columns (spec M/N,
# wired at v3.1) appear here as ordinary rows named "left:right", and their
# value is an independent input like every other row's.
_BACK_TRANSFORM_NOTE = (
    "Back-Transform Method — how a Log-transformed response is returned to "
    "original units. Applies to the point estimate only.\n\n"
    "Duan = Duan (1983) smearing — estimates the conditional MEAN.\n"
    "Naive = textbook EXP(ŷ) — the conditional MEDIAN, biased low "
    "for the mean.\n\n"
    "CI/PI bounds are back-transformed with EXP alone under BOTH settings, "
    "because a bound is a quantile and EXP preserves quantiles. So under "
    "Duan the point estimate does not sit at the centre of its interval. "
    "That gap is correct, not a defect: the mean and the median of a "
    "skewed distribution are different numbers."
)

_PREDICTION_INPUT_NOTE = (
    "Prediction Inputs — one row per constructed design-matrix column, "
    "pre-filled with that column's Training Mean. Type a raw, real-world "
    "value; a Log-transformed row is logged for you.\n\n"
    "Interaction rows (named \"left \u00d7 right\", or with \u2212 / \u00f7 "
    "for a Difference or Ratio) are independent inputs, NOT "
    "recomputed from the two operand rows. Leave the whole band at its "
    "defaults and the prediction sits on the model's own centre. If you "
    "override an operand, override its interaction rows to match — "
    "otherwise the prediction mixes a new operand value with the old "
    "interaction value."
)

_MODEL_CONTEXT_ELEMENTS: tuple[ModelContextElement, ...] = (
    ModelContextElement("Has_Intercept", "Allow Intercept", "Allow_Intercept"),
    ModelContextElement(
        "DF_Absorbed", "Absorbed DF", "Absorbed_Degrees_Of_Freedom()"
    ),
    ModelContextElement(
        "Response_Transform", "Response Transform", _RESPONSE_TRANSFORM_FORMULA
    ),
    ModelContextElement(
        "Predictor_Transform", "Predictor Transform", _PREDICTOR_TRANSFORM_FORMULA
    ),
)

# Model Context is a bounded, fixed-height cache: the element table IS the
# height, so Fit_Context's fixed range and the on-sheet health check below both
# derive from it and a future element can never leave one of them behind.
_MODEL_CONTEXT_ROWS = len(_MODEL_CONTEXT_ELEMENTS)
_MODEL_CONTEXT_LAST_ROW = _MATERIALIZATION_FIRST_ROW + _MODEL_CONTEXT_ROWS - 1
# Health-check row, immediately under the block and inside its border box.
_ROW_MODEL_CONTEXT_CHECK = _MODEL_CONTEXT_LAST_ROW + 1

_MODEL_CONTEXT_LABEL_WIDTH = 20.0
_MODEL_CONTEXT_VALUE_WIDTH = 14.0
_SAMPLE_INCLUDE_MATERIALIZED_WIDTH = 14.0
_DESIGN_MATRIX_COLUMN_WIDTH = 12.0
_CONSTRUCTED_DESIGN_MATRIX_LABEL_WIDTH = 24.0
_MODEL_FORMULA_LABEL_WIDTH = 14.0

# ── The Model Formula readout ─────────────────────────────────────────────────
# The assembled "<response> ~ 1 + <predictors> [| <FE>]" string — a LABEL for
# the model, and the v3.4 Model Comparison sheet's per-row caption
# (Comparison_Model_Formula points here).
#
# Sits on ROW 1 of the terminal Constructed Design Matrix zone, right of
# that zone's own heading: header two columns right of it, the readout one
# column right of the header. Row 1 is the one row in this zone that no
# amount of design matrix can reach — the names spill on
# _MATERIALIZATION_HEADER_ROW and the values on _MATERIALIZATION_SPILL_ROW, and
# both grow RIGHTWARD from there, never up — so this placement does not breach
# the §4b ordering rule (nothing is placed to the right of the zone; the
# caption is placed ABOVE its body, inside the zone's own columns).
#
# Which is the point of putting it here: with WrapText OFF and nothing else on
# row 1 to its right, the string overflows across as many empty columns as it
# needs. The gap between header and readout is what keeps the header itself
# readable — "Model Formula" is in a fixed 14 point width design-matrix
# column, so the readout starting immediately beside it would clip the header
# instead.
#
# Both columns derive from _C_DESIGN_MATRIX, so the caption tracks the zone.
_ROW_MODEL_FORMULA = 1
_C_MODEL_FORMULA_LABEL = _C_DESIGN_MATRIX + 1
_C_MODEL_FORMULA = _C_MODEL_FORMULA_LABEL + 1

# ── The design-matrix width guard ─────────────────────────────────────────────
# Two thresholds, both computed PRE-FLIGHT from the spec block's Design
# Columns audit total rather than from COLUMNS(Design_Columns()). That is the
# whole point of the audit column: constructing a 16,000-column array in order
# to discover it does not fit is the failure being prevented.
_MAX_EXCEL_COLUMN = 16384

# HARD ERROR — the practical limit where Excel cannot invert the Gram
# matrix. Empirically, k = 205 on n = 2909 rows produces all-nan results
# (verified by the L07 guard-state case). Gram_Inverse is O(k^3) in MMULT,
# so 200 columns is near the wall on real datasets. A model reaching this
# many constructed columns needs dimensionality reduction, not a wider sheet.
_DESIGN_MATRIX_MAX_COLUMNS = 200

# SOFT WARNING — k constructed columns, or n x k materialized cells,
# whichever trips first. This fires before the hard error, warning the user
# that recalculation is getting slow (every materialized cell recalculates
# on any input change) and that they are approaching the Gram inversion
# limit. On the WHO data with Country as a Categorical Predictor the matrix
# is roughly 2,938 x 156 = 458,000 live cells.
_DESIGN_MATRIX_SOFT_COLUMNS = 100
_DESIGN_MATRIX_SOFT_CELLS = 200_000

# How much of the terminal zone gets an explicit column width. The zone runs to
# the sheet's right edge, but sizing out to column 16,384 would bloat the sheet
# for a width no usable model reaches, so the sized band stops at the soft
# column threshold — past it the width guard has already fired. (This band used
# to be the terminal zone's outline group as well; the grouping is gone, because
# a collapsed group over the design-matrix spill is what stops Excel
# recalculating the model. The width is all that remains.)
_DESIGN_MATRIX_SIZED_COLUMNS = _DESIGN_MATRIX_SOFT_COLUMNS

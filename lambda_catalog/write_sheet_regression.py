"""
write_sheet_regression.py
Writes the spec-driven Regression sheet into any target workbook.

Layout (five horizontal zones, each preceded — after the spec block — by a
single ungrouped GAP column so the zones collapse independently; see the
"Column-layout paradigm" note above the _C_* constants):
  Col A–O        — Model Specification: the declarative spec block shared with
                   write_sheet_model_construction (Variable / Role / Include /
                   Type / Reference Level / Order / Transform / Sequence /
                   Sequence Period / Period In Use / Levels / Reference In Use /
                   Interaction Term / Interaction Operation / Design Columns).
                   Row 2 = Intercept control (label A2, Allow_Intercept toggle
                   C2) and the Sequence status line (E1); I1 carries the
                   "Verdict" header and I2 the combined verdict switch
                   (the row-1/row-2 cells of column I are above the spec
                   table (SpecTable), so the verdict overlays the
                   Sequence_Period column's input cells without
                   disturbing the spec rows); M2 carries the design-matrix
                   width-guard status and N1/O1 its Σ label and total;
                   P holds the Δ header and Q holds the Count header for
                   the Sequence_Delta_Spectrum() spill at P2:Q?
  Col R          — thin gap (width 2, ungrouped)
  Col S–Y        — Predictor Summary: level-qualified constructed names (S) +
                   Pearson R, Spearman R, Skewness, Kurtosis, GVIF, Tolerance —
                   computed on the CONSTRUCTED design matrix (dummies included);
                   GVIF/Tolerance share one value across a categorical
                   predictor's dummy columns (Generalized VIF)
  Col Z          — thin gap (width 2, ungrouped)
  Col AA–AH      — Regression Outputs: Predicted Variable readout (AE2:AF2),
                   Statistics (AA–AB rows 3–8), Diagnostics (AD–AE rows 3–12;
                   the serial-correlation pair DW/BFN at rows 11–12),
                   Alpha input (AB12), ANOVA Table (rows 13–17, AA–AF),
                   Coefficients (rows 19+, AA–AH), Beta Weights (AH)
  Col AI         — thin gap (width 2, ungrouped)
  Col AJ–AL      — Prediction Outputs: Prediction Interval (AJ1:AK14, boxed —
                   point/CI/PI rows 3-11, FE Group selector + ybar_i/T_i
                   readouts rows 12-14), Prediction Inputs (AJ16+, one row
                   per constructed column, no Intercept row), Training Mean
                   spill (AL19 — the single Predictor_Columns() evaluation
                   the orange AK prefills INDEX into; owns column AL downward
                   so it can never collide with another spill)
  Col AM         — thin gap (width 2, ungrouped)
  Col AN–AZ      — Residual Output: heading + Row_Labels() identifiers in AN;
                   11 diagnostics columns (AO–AY, the last being the Cook's
                   Distance flagged-point overlay feeding chart data labels),
                   plus AZ gutter; spills downward from row 3
  Col BA–BD      — chart title / axis-label formula cells (below the chart grid)
  Col BQ →       — the ARCHITECTURE §4b materialization band: the Model
                   Context block (a boxed, fixed-size label/value pair, one
                   labelled cell per context element), the materialized
                   Sample_Include row mask, and the terminal Constructed
                   Design Matrix — the last two headed on row 2 and spilling
                   from row 3 — each in its own outline group separated by
                   ungrouped gutters. Nothing may ever be placed right of the
                   design matrix.

Every A1 address quoted above is DERIVED in code from the _C_* column
constants (see _abs_ref / _band and the _A_* anchors), never spelled out in a
formula string — a column insertion that shifted one would otherwise leave a
formula that still parses and reads the wrong cell.

The spec block replaces the v1 A–B Model Selection zone (predictor toggles +
Allow_Intercept in B2). Everything the v1 sheet hard-wired is now derived:
    y                          → Response_Column()
    All_Xs / x_s()             → Source_Data / Design_Columns() (the
                                 spec-driven design-matrix constructor; it
                                 owns the intercept column from v3.0)
    Regression_Sample_Include  → Sample_Include()
    Coefficient_Name_Col       → Constructed_Column_Names()
    data_identifiers           → Row_Labels()
The constructor closures come from lambda_functions.json (scope
"Regression") and are registered sheet-scoped here, exactly as the Model
Construction sheet registered them; the spec-block writers are imported from
write_sheet_model_construction so the two sheets can never drift.
"""
# pylint: disable=too-many-lines
from __future__ import annotations

from pathlib import Path
from typing import Any, NamedTuple

import xlwings as xw

from .catalog_schema import CatalogFunction, load_catalog_document
from .sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_DARK_YELLOW_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
    HEADER_COLOR as _HEADER,
    INPUT_COLOR as _INPUT,
)
from .workbook_helpers import (
    MAX_EXCEL_ROW, a1, add_expression_format, anchor_comment_right_of_cell,
    bold, bold_row, border_box, col_letter, drop_local_name, excel_color, f,
    format_input, note_dimensions, rc, safe_activate, section_heading, val,
)
from .write_sheet_model_construction import (
    SPEC_DATASET_PROFILES,
    SpecDatasetProfile,
    _INCLUDE_NOTE,
    _LABEL_NOTE,
    _LEVELS_NOTE,
    _PERIOD_IN_USE_NOTE,
    _REF_IN_USE_NOTE,
    _REFERENCE_NOTE,
    _ROLE_NOTE,
    _SEQUENCE_PERIOD_NOTE,
    _DESIGN_COLUMNS_NOTE,
    _FIXED_EFFECTS_COUNT_FORMULA,
    _FIXED_EFFECTS_NAME_FORMULA,
    _INTERACTION_OPERATION_NOTE,
    _INTERACTION_TERM_NOTE,
    _RESERVED_NOTE,
    _RESPONSE_LOG_FORMULA,
    _RESPONSE_NAME_FORMULA,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    _SEQUENCE_FLAG_COUNT_FORMULA,
    _SEQUENCE_NOTE,
    _TRANSFORM_NOTE,
    _TYPE_NOTE,
    _C_INCLUDE as _C_SPEC_INCLUDE,
    _C_LABEL as _C_SPEC_LABEL,
    _C_LEVELS as _C_SPEC_LEVELS,
    _C_PERIOD_IN_USE as _C_SPEC_PERIOD_IN_USE,
    _C_REFERENCE as _C_SPEC_REFERENCE,
    _C_ROLE as _C_SPEC_ROLE,
    _C_SEQUENCE_PERIOD as _C_SPEC_SEQUENCE_PERIOD,
    _C_DESIGN_COLUMNS as _C_SPEC_DESIGN_COLUMNS,
    _C_INTERACTION_OPERATION as _C_SPEC_INTERACTION_OPERATION,
    _C_INTERACTION_TERM as _C_SPEC_INTERACTION_TERM,
    _C_ORDER as _C_SPEC_ORDER,
    _C_REF_IN_USE as _C_SPEC_REF_IN_USE,
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
    _C_SPEC_LAST,
    _C_TRANSFORM as _C_SPEC_TRANSFORM,
    _C_TYPE as _C_SPEC_TYPE,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _HEADER_ROW as _SPEC_HEADER_ROW,
    _set_sheet_scoped_names as _set_spec_scoped_names,
    _set_spec_block_column_widths,
    _set_spec_block_optional_outline_group,
    _write_intercept_control,
    _write_spec_block,
    _write_spec_feedback,
)

# ── Conditional-formatting helpers ────────────────────────────────────────────

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
# writers in write_sheet_model_construction (imported above); only the section
# heading cell is written here.
_C_A = 1    # spec: Variable labels / A1 zone heading / A2 Intercept label

# Spec feedback (P, Q, plus the I Verdict overlay): the delta spectrum
# (Sequence_Delta_Spectrum() spill at P2:Q?) sits in P and Q. The combined
# verdict switch overlays column I (the Sequence_Period spec column on
# the spec data rows; the row-1/row-2 cells are above the spec table and
# free). I1 holds the "Verdict" header (bold), I2 the priority-ordered
# switch formula (off-grid outranks regularity, both outrank no-natural
# and calendar; red CF outranks yellow via StopIfTrue). The E1 cell
# carries the multi-flag Sequence error status (moved here from H2 when
# the spec data area became a structured table, SpecTable). The design-
# matrix width guard writes M2 (status) and N1/O1 (label/total).
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


# ── Chart constants ───────────────────────────────────────────────────────────
_XL_XY_SCATTER = -4169       # Excel xlXYScatter
_XL_XY_SCATTER_LINES_NO_MARKERS = 75  # Excel xlXYScatterLinesNoMarkers
_XL_COLUMN_CLUSTERED = 51    # Excel xlColumnClustered
_XL_LINE = 4                 # Excel xlLine
_XL_CATEGORY = 1             # horizontal axis
_XL_VALUE = 2                # vertical axis
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
# The Model Context zone is two columns (labels then values) and Sample_Include
# is one; both ship EXPANDED. The terminal Constructed Design Matrix ships
# COLLAPSED, because an unbounded-width zone that cannot be collapsed is a
# scrolling hazard. Gutters are width-2 ungrouped separators so each zone
# collapses independently — the first gutter (after the charts) is structural,
# keeping the floating chart anchors out of every collapsible outline group.
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
_RESPONSE_TRANSFORM_FORMULA = (
    'IFERROR(INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data)),'
    f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data)))),"None")'
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
    'nL,SUMPRODUCT(mask*N(trn="Log")),'
    'nN,SUMPRODUCT(mask*N(trn="None")),'
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

# ── The design-matrix width guard ─────────────────────────────────────────────
# Two thresholds, both computed PRE-FLIGHT from the spec block's Design
# Columns audit total rather than from COLUMNS(Design_Columns()). That is the
# whole point of the audit column: constructing a 16,000-column array in order
# to discover it does not fit is the failure being prevented.
_MAX_EXCEL_COLUMN = 16384

# HARD ERROR — the widest matrix that still fits on the sheet, derived from
# the layout constants rather than hard-coded. Equivalently
# 16,384 - (_LAST_CHART_COLUMN + 6), the six columns being the three gutters
# plus the Model Context label/value pair and the Sample_Include column.
_DESIGN_MATRIX_MAX_COLUMNS = _MAX_EXCEL_COLUMN - _C_DESIGN_MATRIX + 1
assert _DESIGN_MATRIX_MAX_COLUMNS == _MAX_EXCEL_COLUMN - (_LAST_CHART_COLUMN + 6)

# SOFT WARNING — k constructed columns, or n x k materialized cells,
# whichever trips first. Gram_Inverse is O(k^3) in MMULT, so the practical
# wall is in the hundreds; a model that reaches 16k columns has been unusable
# for a long time already. The cell count is the second threshold because
# materialization cost scales with n too: on the WHO data with Country as a
# Categorical Predictor the matrix is roughly 2,938 x 156 = 458,000 live cells
# that recalculate on any input change.
_DESIGN_MATRIX_SOFT_COLUMNS = 200
_DESIGN_MATRIX_SOFT_CELLS = 500_000

# The collapse control over the terminal zone covers a bounded band: an
# outline group has to name its columns, and grouping out to column 16,384
# would bloat the sheet for a width no usable model reaches. The band is
# sized to the soft column threshold — past it the guard has already fired.
_DESIGN_MATRIX_GROUPED_COLUMNS = _DESIGN_MATRIX_SOFT_COLUMNS

# ── Visual formatting helpers ─────────────────────────────────────────────────


def _input_range(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    sheet.range(rc(r1, c1), rc(r2, c2)).color = _INPUT


# ── Note (cell comment) sizing ─────────────────────────────────────────────────
# Excel's default new-comment box (≈128x74pt) is too small for most of the
# plain-language notes on this sheet, so every note is explicitly sized and
# positioned instead of left at the Excel default. Width/height are guessed
# from the text length (see _note_dimensions); either axis can be overridden
# Per-note size overrides for the Regression sheet's notes. Key is the note's
# label — the sheet_notes key for statistical-term notes (e.g. "Durbin-Watson"),
# or the human-readable label passed at the spec-block call sites (e.g. "Reserved",
# "Transform"). The shared heuristic in `workbook_helpers.note_dimensions` handles
# the rest; this dict is just the hand-tuning knob for outliers.
_NOTE_SIZE_OVERRIDES: dict[str, tuple[float | None, float | None]] = {
    # "Durbin-Watson": (320.0, 170.0),  # example manual override (width, height)
}


def _set_note(
    sheet: xw.Sheet, row: int, col: int, text: str, *, label: str | None = None
) -> None:
    """Replace the cell's note/comment text with a plain-language explanation.

    The note box is sized to fit `text` (via `note_dimensions`) and anchored
    directly to the right of the cell, rather than left at Excel's small
    default box and default offset position.
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
    width, height = note_dimensions(
        label if label is not None else text, text, _NOTE_SIZE_OVERRIDES
    )
    anchor_comment_right_of_cell(sheet, row, col, width, height)


def _annotate_statistical_terms(sheet: xw.Sheet, sheet_notes: dict[str, str]) -> None:
    """Attach plain-language notes to key statistical labels on the sheet."""
    note_cells = [
        (2, _C_T, "Pearson R"),
        (2, _C_U, "Spearman R"),
        (2, _C_V, "Skewness"),
        (2, _C_W, "Kurtosis"),
        (2, _C_X, "GVIF"),
        (2, _C_Y, "Tolerance"),
        (4, _C_AA, "Multiple R"),
        (5, _C_AA, "R Square"),
        (6, _C_AA, "Adjusted R Square"),
        (7, _C_AA, "Standard Error"),
        (8, _C_AA, "Observations"),
        (2, _C_AA, "Model Formula"),
        (4, _C_AD, "PRESS"),
        (5, _C_AD, "PRESS R²"),
        (6, _C_AD, "Mean Leverage"),
        (7, _C_AD, "AIC"),
        (8, _C_AD, "BIC"),
        (9, _C_AD, "AICc"),
        (10, _C_AD, "QQ Correlation"),
        (11, _C_AD, "Durbin-Watson"),
        (12, _C_AD, "BFN Panel Durbin-Watson"),
        (12, _C_AA, "Alpha"),
        (14, _C_AB, "df"),
        (14, _C_AC, "SS"),
        (14, _C_AD, "MS"),
        (14, _C_AE, "F"),
        (14, _C_AF, "Significance F"),
        (15, _C_AA, "Regression"),
        (16, _C_AA, "Residual"),
        (17, _C_AA, "Total"),
        (20, _C_AB, "Coefficients"),
        (20, _C_AC, "Std Error"),
        (20, _C_AD, "t Stat"),
        (20, _C_AE, "P-value"),
        (20, _C_AF, "Lower 95%"),
        (20, _C_AG, "Upper 95%"),
        (20, _C_AH, "Beta Weight"),
        (17, _C_AL, "Training Mean"),
        (4, _C_AG, "Back-Transform"),
        (5, _C_AG, "Smearing Factor"),
        (6, _C_AG, "R Square (Unit)"),
        (7, _C_AG, "Adj R Square (Unit)"),
        (8, _C_AG, "RMSE (Unit)"),
        (9, _C_AG, "Response Space"),
        (3, _C_AJ, "Point Estimate"),
        (4, _C_AJ, "SE (Mean)"),
        (5, _C_AJ, "SE (New Obs)"),
        (6, _C_AJ, "t Critical"),
        (7, _C_AJ, "CI Lower"),
        (8, _C_AJ, "CI Upper"),
        (9, _C_AJ, "PI Lower"),
        (10, _C_AJ, "PI Upper"),
        (11, _C_AJ, "Confidence Level"),
        (12, _C_AJ, "FE Group"),
        (13, _C_AJ, "Group Mean (y)"),
        (14, _C_AJ, "Group Count"),
        (3, _C_AL, "Point Estimate (Original Units)"),
        (2, _C_AO, "Y"),
        (2, _C_AP, "Predicted Y"),
        (2, _C_AQ, "Residuals"),
        (2, _C_AR, "Hat Diagonal"),
        (2, _C_AS, "Studentized Residuals"),
        (2, _C_AT, "Cook's Distance"),
        (2, _C_AU, "Normal Scores Ranked"),
        (2, _C_AV, "Studentized Residuals Ranked"),
        (2, _C_AW, "Scale-Location"),
        (2, _C_AX, "PRESS Residual"),
        (2, _C_AY, "Cook's Distance (Flagged)"),
        (2, _C_AZ, "Predicted Y (Original Units)"),
        (2, _C_BA, "Residual (Original Units)"),
    ]

    for row, col, key in note_cells:
        note_text = sheet_notes.get(key)
        if note_text is not None:
            _set_note(sheet, row, col, note_text, label=key)


def _write_significance_conditional_formatting(sheet: xw.Sheet) -> None:
    """Flag nonsignificant coefficient and overall-model P-values."""

    coefficient_p_values = _band(_C_AE, _ROW_COEFF_FIRST)
    p_anchor = f"{col_letter(_C_AE)}{_ROW_COEFF_FIRST}"
    significance_f = _A_SIGNIFICANCE_F

    sheet.range(coefficient_p_values).api.FormatConditions.Delete()
    sheet.range(significance_f).api.FormatConditions.Delete()

    # Individual coefficient P-values above alpha. The p-value reference is
    # RELATIVE (no $) so the one rule re-anchors per row down the band; the
    # alpha reference is absolute so every row tests the same input cell.
    add_expression_format(
        sheet,
        coefficient_p_values,
        f"=AND(ISNUMBER({p_anchor}),{p_anchor}>{_A_ALPHA})",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Overall regression P-value above alpha.
    add_expression_format(
        sheet,
        significance_f,
        f"=AND(ISNUMBER({significance_f}),{significance_f}>{_A_ALPHA})",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_residual_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply diagnostic cutoffs to the residual-output columns."""

    addresses = {
        "hat":                _band(_C_AR),
        "studentized":        _band(_C_AS),
        "cooks":              _band(_C_AT),
        "studentized_ranked": _band(_C_AV),
        "scale_location":     _band(_C_AW),
        "press_residual":     _band(_C_AX),
        "cooks_flag":         _band(_C_AY),
    }
    # Relative per-row anchors for the banded rules (see the note on the
    # coefficient P-value rule above: no $, so the rule walks down the band).
    hat = f"{col_letter(_C_AR)}{_ROW_DATA_FIRST}"
    scale_location = f"{col_letter(_C_AW)}{_ROW_DATA_FIRST}"
    press_residual = f"{col_letter(_C_AX)}{_ROW_DATA_FIRST}"

    # Remove existing rules so repeated builds do not duplicate them.
    for address in addresses.values():
        sheet.range(address).api.FormatConditions.Delete()

    # ── Hat diagonal ─────────────────────────────────────────────────────────
    # The mean-leverage cell holds p/n.
    # > 2p/n: light-red fill and dark-red text.
    add_expression_format(
        sheet,
        addresses["hat"],
        f"=AND(ISNUMBER({hat}),{hat}>2*{_A_MEAN_LEVERAGE})",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # > 3p/n: additionally bold.
    add_expression_format(
        sheet,
        addresses["hat"],
        f"=AND(ISNUMBER({hat}),{hat}>3*{_A_MEAN_LEVERAGE})",
        bold_font=True,
    )

    # ── Studentized residuals ────────────────────────────────────────────────
    for column, address in [
        (col_letter(_C_AS), addresses["studentized"]),
        (col_letter(_C_AV), addresses["studentized_ranked"]),
    ]:
        # 2 < |r| < 3: light-yellow fill and dark-yellow text.
        add_expression_format(
            sheet,
            address,
            (
                f"=AND("
                f"ISNUMBER({column}3),"
                f"ABS({column}3)>2,"
                f"ABS({column}3)<3"
                f")"
            ),
            fill=CF_YELLOW_FILL,
            font_color=CF_DARK_YELLOW_TEXT,
        )

        # |r| >= 3: light-red fill and dark-red text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({column}3),ABS({column}3)>=3)",
            fill=CF_LIGHT_RED_FILL,
            font_color=CF_DARK_RED_TEXT,
        )

    # ── Cook's distance (and its "Flagged" duplicate) ───────────────────────
    # The Observations cell holds n. The Flagged column mirrors Cook's D
    # (NA() below both cutoffs), so the same two rules applied to it just
    # recolor whatever value the source column already produced there — kept
    # visually consistent with the column it duplicates.
    for column, address in [
        (col_letter(_C_AT), addresses["cooks"]),
        (col_letter(_C_AY), addresses["cooks_flag"]),
    ]:
        cell = f"{column}{_ROW_DATA_FIRST}"
        # 4/n < D <= 0.9: light-yellow fill and dark-yellow text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({cell}),{cell}>4/{_A_OBSERVATIONS},{cell}<=0.9)",
            fill=CF_YELLOW_FILL,
            font_color=CF_DARK_YELLOW_TEXT,
        )

        # D > 0.9: light-red fill and dark-red text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({cell}),{cell}>0.9)",
            fill=CF_LIGHT_RED_FILL,
            font_color=CF_DARK_RED_TEXT,
        )

    # ── Scale-Location: SQRT(|Studentized|) ─────────────────────────────────
    # SQRT(2) ≈ 1.414 corresponds to |Studentized| = 2.
    # SQRT(3) ≈ 1.732 corresponds to |Studentized| = 3.
    add_expression_format(
        sheet,
        addresses["scale_location"],
        f"=AND(ISNUMBER({scale_location}),"
        f"{scale_location}>1.414,{scale_location}<=1.732)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["scale_location"],
        f"=AND(ISNUMBER({scale_location}),{scale_location}>1.732)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS Residual: e_i / (1 - h_i) ─────────────────────────────────────
    # The Standard Error cell holds the regression's SE.
    # |PRESS| > 2*SE: mild concern; > 3*SE: strong concern.
    add_expression_format(
        sheet,
        addresses["press_residual"],
        f"=AND(ISNUMBER({press_residual}),"
        f"ABS({press_residual})>2*{_A_STANDARD_ERROR},"
        f"ABS({press_residual})<=3*{_A_STANDARD_ERROR})",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["press_residual"],
        f"=AND(ISNUMBER({press_residual}),"
        f"ABS({press_residual})>3*{_A_STANDARD_ERROR})",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_model_diagnostic_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply rule-of-thumb formatting to GVIF, PRESS R², and QQ Correlation."""

    gvif_address = _band(_C_X)
    gvif = f"{col_letter(_C_X)}{_ROW_DATA_FIRST}"
    press_r2_address = _A_PRESS_R_SQUARED
    qq_corr_address = _A_QQ_CORRELATION

    # Prevent duplicate rules when rebuilding the sheet.
    sheet.range(gvif_address).api.FormatConditions.Delete()
    sheet.range(press_r2_address).api.FormatConditions.Delete()
    sheet.range(qq_corr_address).api.FormatConditions.Delete()

    # ── GVIF ────────────────────────────────────────────────────────────────
    # 5 < GVIF <= 10: possible multicollinearity; review. Thresholds are exact
    # for a continuous (single-column) predictor and reflect the combined
    # block for a multi-level categorical one (raw, unstandardized GVIF).
    add_expression_format(
        sheet,
        gvif_address,
        f"=AND(ISNUMBER({gvif}),{gvif}>5,{gvif}<=10)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # GVIF > 10: strong multicollinearity warning.
    add_expression_format(
        sheet,
        gvif_address,
        f"=AND(ISNUMBER({gvif}),{gvif}>10)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS R² ─────────────────────────────────────────────────────────────
    # Negative PRESS R² means cross-validated predictions perform worse than
    # predicting the outcome mean.
    add_expression_format(
        sheet,
        press_r2_address,
        f"=AND(ISNUMBER({press_r2_address}),{press_r2_address}<0)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── QQ Correlation ────────────────────────────────────────────────────────
    # Pearson r of sorted scaled residuals vs. normal quantiles; near 1.0 = normal errors.
    # < 0.98: mild departure (yellow); < 0.95: stronger departure (red).
    add_expression_format(
        sheet,
        qq_corr_address,
        f"=AND(ISNUMBER({qq_corr_address}),"
        f"{qq_corr_address}<0.98,{qq_corr_address}>=0.95)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        qq_corr_address,
        f"=AND(ISNUMBER({qq_corr_address}),{qq_corr_address}<0.95)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


# ── Local name management ─────────────────────────────────────────────────────

# v1 sheet-scoped names replaced by the spec-driven constructors — dropped on
# every rebuild so a workbook carried forward never resolves against them.
_LEGACY_LOCAL_NAMES = (
    "All_Xs",
    "Coefficient_Name_Col",
    "Ind_Var_Filter",
    "Ind_Var_Include",
    "fil",
    "y",
    "Regression_Sample_Include",
    "data_identifiers",
    "pred_input",
)


def _setup_local_names(
    sheet: xw.Sheet,
    closures: tuple[CatalogFunction, ...] | None = None,
    source_table_ref: str = "=MileageData[#All]",
) -> None:
    """Register sheet-scoped names used by every formula on this sheet.

    The spec wiring (Source_Data / Header_Names / Spec_* / Allow_Intercept)
    and the constructor closures (Sample_Include / Response_Column /
    Row_Labels / Predictor_Columns / Constructed_Column_Names) are registered by
    ``_set_sheet_scoped_names`` from write_sheet_model_construction; this
    function adds the Regression-only names on top.
    """
    sname = sheet.name

    if closures is None:
        closures = load_catalog_document(_DEFINITIONS_PATH).functions_for_sheet(
            REGRESSION_SHEET_NAME
        )

    for legacy in _LEGACY_LOCAL_NAMES:
        drop_local_name(sheet, legacy)

    _set_spec_scoped_names(sheet, closures, source_table_ref=source_table_ref)

    # Zero_Predictors_Selected(): TRUE when the spec contributes no
    # predictor columns — no included Predictor rows, or every included
    # Categorical degenerate. Predictor_Columns() errors in that state (DROP
    # of the sentinel-only accumulator), so the width probe wraps IFERROR
    # rather than counting Include toggles the way v1 did.
    #
    # It must probe Predictor_Columns(), NOT Design_Columns(): with the
    # intercept relocated into the constructor, Design_Columns() returns the
    # lone ones column in exactly this state, so counting its columns would
    # report 1 and the zero-predictor branch would never fire.
    drop_local_name(sheet, "Zero_Predictors_Selected")
    sheet.api.Names.Add(
        Name="Zero_Predictors_Selected",
        RefersTo="=LAMBDA(IFERROR(COLUMNS(Predictor_Columns()),0)=0)",
    )

    # alpha: the confidence-level input cell in the Statistics block.
    drop_local_name(sheet, "alpha")
    sheet.api.Names.Add(
        Name="alpha",
        RefersTo=f"={sname}!{_A_ALPHA}",
    )

    # ── Intercept-only closed-form helpers ──────────────────────────────────
    # Used by _write_coefficients and _write_prediction_interval when
    # Zero_Predictors_Selected() is TRUE and Allow_Intercept is TRUE: an
    # intercept-only OLS model (Y = b0 + error) is still statistically
    # well-defined even though Predictor_Columns() has nothing to construct.
    # Bypasses Predictor_Columns()/Coefficients()/Prediction_Interval()
    # entirely since Excel cannot represent a valid zero-column array.
    #
    # Design_Columns() does return a well-formed ones column in this state
    # (the intercept stage runs even when the predictor stage is empty), so
    # the engines could in principle fit it directly. The closed-form bypass
    # is kept because it is what the shipped behaviour was verified against;
    # retiring it is a follow-up, not part of the relocation.
    # Intercept_Only_N uses SUMPRODUCT over the computed mask (COUNTIF needs
    # a range reference, and Sample_Include() is an array) so it never
    # errors, even when the mask has zero TRUE rows — callers guard on its
    # value before invoking the FILTER/STDEV.S-based helpers below.
    drop_local_name(sheet, "Intercept_Only_N")
    sheet.api.Names.Add(
        Name="Intercept_Only_N",
        RefersTo="=LAMBDA(SUMPRODUCT(N(Sample_Include())))",
    )

    drop_local_name(sheet, "Intercept_Only_Point")
    sheet.api.Names.Add(
        Name="Intercept_Only_Point",
        RefersTo="=LAMBDA(AVERAGE(FILTER(Response_Column(),Sample_Include())))",
    )

    drop_local_name(sheet, "Intercept_Only_S")
    sheet.api.Names.Add(
        Name="Intercept_Only_S",
        RefersTo="=LAMBDA(STDEV.S(FILTER(Response_Column(),Sample_Include())))",
    )

    drop_local_name(sheet, "Intercept_Only_SE")
    sheet.api.Names.Add(
        Name="Intercept_Only_SE",
        RefersTo="=LAMBDA(Intercept_Only_S()/SQRT(Intercept_Only_N()))",
    )

    drop_local_name(sheet, "Intercept_Only_DF")
    sheet.api.Names.Add(
        Name="Intercept_Only_DF",
        RefersTo="=LAMBDA(Intercept_Only_N()-1)",
    )

    # ── Chart data ranges (OFFSET-based, sized to n = the Observations cell) ───
    # These worksheet-scoped names feed chart SERIES formulas as
    # ='Regression'!<Name>, avoiding full-column references that degrade
    # performance and avoiding the unsupported # spill operator in chart formulas.
    # The third element becomes the Name Manager comment so a user browsing
    # names can tell which chart each range feeds.
    for _name, _col_ltr, _comment in [
        ("RegChartQQX", col_letter(_C_AU),
         "Normal Q-Q chart: X values (theoretical quantiles, Normal Scores Ranked)"),
        ("RegChartQQY", col_letter(_C_AV),
         "Normal Q-Q chart: Y values (Studentized Residuals Ranked)"),
        ("RegChartFitY", col_letter(_C_AP),
         "Predicted Y: X values for the Residuals vs. Fitted, Actual vs. Predicted, and Scale-Location charts"),
        ("RegChartResid", col_letter(_C_AQ),
         "Residuals vs. Fitted chart: Y values (Residuals)"),
        ("RegChartActY", col_letter(_C_AO),
         "Actual vs. Predicted chart: Y values (Actual Y)"),
        ("RegChartScaleLoc", col_letter(_C_AW),
         "Scale-Location chart: Y values (sqrt of abs Studentized Residual)"),
        ("RegChartCookDist", col_letter(_C_AT),
         "Cook's Distance chart: bar values"),
        ("RegChartLeverage", col_letter(_C_AR),
         "Studentized Residuals vs. Leverage chart: X values (Hat Diagonal)"),
        ("RegChartStudResid", col_letter(_C_AS),
         "Studentized Residuals vs. Leverage chart: Y values"),
        ("RegChartPRESSResid", col_letter(_C_AX),
         "PRESS Residuals chart: bar values"),
        ("RegChartCookDistFlag", col_letter(_C_AY),
         "Cook's Distance chart: flagged-point overlay for data labels (D > 4/n or D > 0.9)"),
        ("RegChartObsLabel", col_letter(_C_AN),
         "Cook's Distance chart: observation identifier for flagged-point data labels"),
    ]:
        drop_local_name(sheet, _name)
        _nm = sheet.api.Names.Add(
            Name=_name,
            RefersTo=(
                f"=OFFSET('{sname}'!${_col_ltr}$2,1,0,"
                f"MAX(IFERROR('{sname}'!{_A_OBSERVATIONS},1),1),1)"
            ),
        )
        _nm.Comment = _comment

    # ── v3.3 Comparison_* names (committed public interface for v3.4 Model
    # Comparison). The unit-space block (AG3:AH9) and the Model Formula cell
    # (AA2:AB2) are the surfaces v3.4 reads from — Comparison_Anchor is the
    # response-name readout that pairs two models; Comparison_Headline_GoF
    # is the three unit-space goodness-of-fit numbers (R², adjusted R²,
    # RMSE in original units) so the comparison sheet can rank alternatives
    # without rewriting the same formulas; Comparison_Model_Formula is the
    # assembled formula string used as the comparison sheet's per-row label.
    for _name, _refers_to, _comment in [
        (
            "Comparison_Anchor",
            f"={sname}!{_A_RESPONSE_READOUT}",
            "Response-name readout (the displayed response label) — pairs two models in v3.4 Model Comparison",
        ),
        (
            "Comparison_Headline_GoF",
            f"={sname}!$AH$6:$AH$8",
            "Unit-space goodness-of-fit triplet (R², adjusted R², RMSE) — feeds the v3.4 Model Comparison headline row",
        ),
        (
            "Comparison_Model_Formula",
            f"={sname}!$AB$2",
            "Assembled model formula string (response ~ predictors [| FE]) — feeds the v3.4 Model Comparison per-row label",
        ),
    ]:
        drop_local_name(sheet, _name)
        _nm = sheet.api.Names.Add(Name=_name, RefersTo=_refers_to)
        _nm.Comment = _comment


# ── Section writers ───────────────────────────────────────────────────────────

def _write_model_specification(sheet: xw.Sheet) -> None:
    """Zone A–O: the shared spec block + row-2 Intercept control.

    The block itself (headers, defaults, dropdowns, CF, the Levels and
    Reference In Use displays) is written by the same functions that build
    the standalone Model Construction sheet, so the two layouts can never
    drift. Only the zone heading and the reserved-column notes are local.

    The spec block's TABLE CREATION (SpecTable) happens at the top of
    ``write_regression_output_sheet`` — names registered after that point
    can bind to the table's columns via SpecTable[Column] structured
    references. Here we just layer the spec feedback (E1 status, M/N
    spectrum, I Verdict overlay), the Intercept control, and the column
    notes on top.
    """
    section_heading(sheet, 1, _C_A, "MODEL SPECIFICATION")
    _write_spec_feedback(sheet)
    _write_intercept_control(sheet)
    _write_design_matrix_width_guard(sheet)
    # Spec-block notes anchor on the header row (row 3) so the tooltip
    # appears when the user hovers the column heading the notes describe,
    # not the first variable row. All twelve spec-block headers carry a
    # plain-language note; the four (Order, Transform, Sequence, Sequence
    # Period) that double as the shipped spec-feature headers use the
    # longer notes defined in write_sheet_model_construction.py.
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_LABEL, _LABEL_NOTE, label="Variable")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_ROLE, _ROLE_NOTE, label="Role")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_INCLUDE, _INCLUDE_NOTE, label="Include")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_TYPE, _TYPE_NOTE, label="Type")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_REFERENCE, _REFERENCE_NOTE, label="Reference Level")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_ORDER, _RESERVED_NOTE, label="Order")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_TRANSFORM, _TRANSFORM_NOTE, label="Transform")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_SEQUENCE, _SEQUENCE_NOTE, label="Sequence")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_SEQUENCE_PERIOD, _SEQUENCE_PERIOD_NOTE, label="Sequence Period")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_PERIOD_IN_USE, _PERIOD_IN_USE_NOTE, label="Period In Use")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_LEVELS, _LEVELS_NOTE, label="Levels")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_REF_IN_USE, _REF_IN_USE_NOTE, label="Reference In Use")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_INTERACTION_TERM, _INTERACTION_TERM_NOTE, label="Interaction Term")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_INTERACTION_OPERATION, _INTERACTION_OPERATION_NOTE, label="Interaction Operation")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_DESIGN_COLUMNS, _DESIGN_COLUMNS_NOTE, label="Design Columns")
    # Prediction Inputs (AJ17). An interaction column (v3.1) gets its own
    # row there like any other constructed column, and its Prediction Value
    # is NOT derived from the two operand rows. Leaving the whole band at
    # its Training Mean defaults is self-consistent (it sits on the design
    # matrix's own centroid); overriding an operand without also overriding
    # the interaction row is not. The band cannot silently recompute it — a
    # prediction row is a user input, and rewriting one input because
    # another changed is the "silently switch" behaviour the spec block
    # exists to avoid — so it says so instead.
    _set_note(sheet, 17, _C_AJ, _PREDICTION_INPUT_NOTE, label="Predictor")


def _write_design_matrix_width_guard(sheet: xw.Sheet) -> None:
    """The ARCHITECTURE §4b pre-flight width guard, in the spec-block area.

    Three cells, all in the free row-1/row-2 band above the spec table
    (SpecTable's own content starts at the row-3 header):

        N1 = "Σ Design Columns"   (bold label)
        O1 = the total            — Σ(spec column O) plus the intercept
                                    column, i.e. exactly COLUMNS(Design_Columns())
        M2 = the guard status     — blank while the model is within limits,
                                    a WARNING line at the soft threshold, an
                                    ERROR line at the hard one

    O1 sits directly above the per-row Design Columns audit it totals, so the
    number and its breakdown read as one column. M2's message overflows
    rightward across N2/O2, which the spec block leaves empty.

    Both thresholds are read from the SPEC, never from
    ``COLUMNS(Design_Columns())`` — a matrix too wide to fit cannot be built
    in order to be measured, which is the failure this guard exists to
    prevent. The hard limit is derived from the layout constants (see
    ``_DESIGN_MATRIX_MAX_COLUMNS``), so moving a zone moves the guard with it.

    The total is a display over displays and feeds no constructor, so it does
    not breach "display derives, never feeds" — see ARCHITECTURE §4.
    """
    total_cell = _abs_ref(1, _C_SPEC_DESIGN_COLUMNS)
    status_cell = _abs_ref(2, _C_SPEC_INTERACTION_TERM)

    val(sheet, 1, _C_SPEC_INTERACTION_OPERATION, "Σ Design Columns")
    bold(sheet, 1, _C_SPEC_INTERACTION_OPERATION)

    # The intercept is not a spec row, so it is added here rather than
    # counted in column O. N() coerces the Allow_Intercept boolean to 1/0;
    # SUM ignores the "" that column O returns on every non-Predictor row.
    f(
        sheet,
        1,
        _C_SPEC_DESIGN_COLUMNS,
        "=SUM(TAKE(Spec_Design_Columns,COLUMNS(Source_Data)))+N(Allow_Intercept)",
    )
    bold(sheet, 1, _C_SPEC_DESIGN_COLUMNS)

    f(
        sheet,
        2,
        _C_SPEC_INTERACTION_TERM,
        (
            f"=LET(k,{total_cell},n,ROWS(Source_Data),"
            f"IF(k>{_DESIGN_MATRIX_MAX_COLUMNS},"
            '"ERROR: the design matrix needs "&k&" columns; only '
            f'{_DESIGN_MATRIX_MAX_COLUMNS} fit right of the materialization '
            "band. Reduce a Categorical predictor's levels or exclude it.\","
            f"IF(OR(k>{_DESIGN_MATRIX_SOFT_COLUMNS},"
            f"n*k>{_DESIGN_MATRIX_SOFT_CELLS}),"
            '"WARNING: "&k&" constructed columns x "&n&" rows = "&(n*k)&'
            '" materialized cells. Gram_Inverse is O(k^3) in MMULT, and every '
            'materialized cell recalculates on any input change.",'
            '"")))'
        ),
    )

    # Red outranks yellow via StopIfTrue, the same priority idiom the
    # Sequence verdict cell uses. Each rule keys on the message's own
    # leading token, so the formula above stays the single source of the
    # thresholds.
    for cell in (status_cell, total_cell):
        add_expression_format(
            sheet,
            cell,
            f'=ISNUMBER(SEARCH("ERROR",{status_cell}))',
            fill=CF_LIGHT_RED_FILL,
            font_color=CF_DARK_RED_TEXT,
            stop_if_true=True,
        )
        add_expression_format(
            sheet,
            cell,
            f'=ISNUMBER(SEARCH("WARNING",{status_cell}))',
            fill=CF_YELLOW_FILL,
            font_color=CF_DARK_YELLOW_TEXT,
        )


def _write_predictor_summary(sheet: xw.Sheet) -> None:
    """Zone P–V: EDA stats for the CONSTRUCTED design-matrix columns."""
    section_heading(sheet, 1, _C_S, "PREDICTOR SUMMARY")

    for col, header in zip(
        [_C_S, _C_T, _C_U, _C_V, _C_W, _C_X, _C_Y],
        ["", "Pearson R", "Spearman R", "Skewness", "Kurtosis", "GVIF", "Tolerance"],
    ):
        val(sheet, 2, col, header)
    bold_row(sheet, 2, _C_S, _C_Y)

    # Spill anchors at row 3 — each spills once per constructed column.
    # Names come from the constructor twin, so dummies are level-qualified
    # and the stats run on the actual design matrix. GVIF/Tolerance are
    # generalized (Fox & Monette): every dummy column from the same
    # categorical predictor shares one value instead of a separate,
    # coding-dependent number per level.
    f(sheet, 3, _C_S, "=TRANSPOSE(Constructed_Column_Names())")
    f(sheet, 3, _C_T, "=Pearson_R(Predictor_Columns(),Response_Column(),Sample_Include())")
    f(sheet, 3, _C_U, "=Spearman_R(Predictor_Columns(),Response_Column(),Sample_Include())")
    f(sheet, 3, _C_V, "=Skewness(Predictor_Columns(),Sample_Include())")
    f(sheet, 3, _C_W, "=Kurtosis(Predictor_Columns(),Sample_Include())")
    f(sheet, 3, _C_X, "=GVIF(Predictor_Columns(),Constructed_Column_Names(),Sample_Include())")
    f(sheet, 3, _C_Y, "=Generalized_Tolerance(Predictor_Columns(),Constructed_Column_Names(),Sample_Include())")

    sheet.range(
        (rc(3, _C_T)), (rc(_FORMAT_BAND_LAST_ROW, _C_Y))
    ).number_format = "0.00"


def _write_regression_outputs_header(sheet: xw.Sheet) -> None:
    section_heading(sheet, 1, _C_AA, "REGRESSION OUTPUTS")
    section_heading(sheet, 2, _C_AE, "Predicted Variable")
    section_heading(sheet, 2, _C_AF, "")
    # Derived response name — the header of the Role=Response spec row.
    f(sheet, 2, _C_AF, f"={_RESPONSE_NAME_FORMULA}")

    # Model Formula cell (v3.3) — AA2 label, AB2 the assembled string. Built
    # entirely from existing pieces so the formula label is always exactly
    # what the model itself is:
    #   response side  = _RESPONSE_NAME_FORMULA (already emits "Ln(name)" when Log)
    #   predictor side = "1 + " + TEXTJOIN(" + ", Constructed_Column_Names())
    #                    (which already emits "Ln(name)" per logged predictor,
    #                     level-qualified dummy names, and "left × right"
    #                     interaction names — the mixed Log/None case renders
    #                     correctly with no extra work)
    #   FE suffix      = " | <FE name>" when a Fixed Effects row is declared
    val(sheet, 2, _C_AA, "Model Formula")
    bold(sheet, 2, _C_AA)
    f(
        sheet,
        2,
        _C_AB,
        (
            f"={_RESPONSE_NAME_FORMULA}"
            '&" ~ "'
            '&IF(Allow_Intercept,"1 + ","0 + ")'
            '&IFERROR(TEXTJOIN(" + ",TRUE,Constructed_Column_Names()),"")'
            f'&IF({_FIXED_EFFECTS_COUNT_FORMULA}>0," | "&{_FIXED_EFFECTS_NAME_FORMULA},"")'
        ),
    )


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    """Cols X–Y, rows 3–8."""
    section_heading(sheet, 3, _C_AA, "REGRESSION STATISTICS")
    # Fit-time X/y (Design_Columns()/Design_Response()): the response is
    # Response_Column() unchanged with no Fixed Effects row and one-way
    # within-demeaned when one is declared
    # — every statistic below reports the "within" flavor under FE, the same
    # convention panel-regression software (e.g. R's plm) uses. Adjusted R²
    # and Standard Error also carry the absorbed df (element 2 of Model_Context,
    # 0 with no FE row) so their df-dependent penalty/divisor is correct.
    for row, label, formula in [
        (4, "Multiple R",        "=Multiple_R(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (5, "R Square",          "=R_Squared(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (6, "Adjusted R Square", "=Adjusted_R_Squared(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (7, "Standard Error",    "=SE_Regression(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (8, "Observations",      "=Observations(Design_Response(),Sample_Include())"),
    ]:
        val(sheet, row, _C_AA, label)
        f(sheet, row, _C_AB, formula)
    sheet.range(rc(4, _C_AB), rc(7, _C_AB)).number_format = "0.0000"
    sheet.range(rc(8, _C_AB), rc(8, _C_AB)).number_format = "0"
    border_box(sheet, 3, _C_AA, 8, _C_AB)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    """Cols AA–AB, rows 3–12."""
    section_heading(sheet, 3, _C_AD, "DIAGNOSTICS")
    for row, label, formula in [
        (4,  "PRESS",          "=PRESS(Design_Columns(),Design_Response(),Sample_Include())"),
        (
            5,
            "PRESS R²",
            "=1-PRESS(Design_Columns(),Design_Response(),Sample_Include())"
            "/SS_Total(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())",
        ),
        (
            6,
            "Mean Leverage",
            "=COLUMNS(Design_Columns())"
            "/Observations(Design_Response(),Sample_Include())",
        ),
        (7,  "AIC",            "=AIC(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (8,  "BIC",            "=BIC(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (9,  "AICc",           "=AICc(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
        (10, "QQ Correlation", "=QQ_Correlation(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"),
    ]:
        val(sheet, row, _C_AD, label)
        f(sheet, row, _C_AE, formula)
    sheet.range(rc(4, _C_AE), rc(10, _C_AE)).number_format = "0.0000"

    # Serial-correlation trigger matrix — two fixed cells (AD11/AE11 plain DW,
    # AD12/AE12 BFN panel DW), each SELF-GUARDING: every state a cell can show is
    # visible in that cell's own formula, never dispatched from a shared
    # selector cell — the audit-friendly form. Off-spec states show an explicit
    # not-applicable token (never NA(), which is reserved for genuine errors,
    # and friendlier than "" to a future Model-Comparison XLOOKUP):
    #
    #   no Sequence axis          → both cells "n/a — requires Sequence"
    #   Sequence + no FE variable → DW computes, BFN "n/a — no fixed effects"
    #   Sequence + FE variable    → BFN computes, DW "n/a — FE active"
    #
    # Durbin-Watson is gated on EXACTLY ONE declared Sequence axis: differencing
    # residuals in physical row order silently assumes row order = time order,
    # which reports spurious autocorrelation under any meaningful non-time sort.
    # Zero flags → "requires Sequence"; two-plus flags is the spec error the E1
    # status line already reports, and computing on Sequence_Column()'s first
    # match would be ambiguous → "multiple Sequence flags". Under fixed effects
    # the single-series statistic is invalid twice over (within-demeaned
    # residuals; no single panel ordering) → "FE active", and the BFN cell
    # below takes over. With exactly one flag and no FE, Durbin_Watson_By sorts
    # residuals by Sequence_Column() before differencing — all array work
    # internal, scalar out, no spill.
    val(sheet, 11, _C_AD, "Durbin-Watson")
    f(
        sheet,
        11,
        _C_AE,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars>0,"n/a — FE active",'
        "Durbin_Watson_By(Design_Columns(),Design_Response(),Sequence_Column(),"
        "Sample_Include())))))",
    )
    sheet.range(rc(11, _C_AE), rc(11, _C_AE)).number_format = "0.000"

    # BFN panel Durbin-Watson (Bhargava–Franzini–Narendranathan 1982): the
    # within-group form for panels under fixed effects. Both this cell and the
    # plain DW cell above read the fit-time pair
    # (Design_Columns()/Design_Response()), and both must: BFN's own contract
    # says "the residuals are within-demeaned" — Residuals(X, Y) only produces
    # within residuals when X/Y already ARE the within-transformed pair. The DW
    # cell above only ever fires in the no-FE state, where Design_Response()
    # reduces to Response_Column() and Design_Columns() to the intercept plus
    # Predictor_Columns(), so reading the fit-time pair there is the same
    # computation stated in the honest way. Differencing is
    # restricted to within-group (group, seq−Δ) pairs via Difference_By inside
    # the LAMBDA, so group seams contribute nothing by construction. Active
    # only when a Sequence axis AND a Fixed Effects variable are declared;
    # multiple FE variables (two-way absorption, out of scope) show a token
    # rather than computing on an arbitrary first match. The group argument is
    # Serial_Correlation_Group() — the grouping-key resolver — NOT the FE
    # column accessor directly: the resolver is the single retargeting point
    # for which dimension partitions residuals (today the Fixed Effects
    # column; its dormant Cluster branch activates at v3.7 without touching
    # this cell). Interpretation reads like DW (near 2 ⇒ no first-order
    # autocorrelation), but its critical values depend on N and T — surfacing
    # those bounds is a recorded open item, and the standard DW bounds must
    # not be presented next to it (the cell note carries that caveat).
    val(sheet, 12, _C_AD, "BFN Panel Durbin-Watson")
    f(
        sheet,
        12,
        _C_AE,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars=0,"n/a — no fixed effects",'
        'IF(fe_vars>1,"n/a — multiple FE variables",'
        "BFN_Panel_Durbin_Watson(Design_Columns(),Design_Response(),"
        "Serial_Correlation_Group(),Sequence_Column(),Base_Period_Delta(),"
        "Sample_Include()))))))",
    )
    sheet.range(rc(12, _C_AE), rc(12, _C_AE)).number_format = "0.000"
    border_box(sheet, 3, _C_AD, 12, _C_AE)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at AB12 — controls prediction interval confidence level."""
    val(sheet, 12, _C_AA, "Alpha")
    bold(sheet, 12, _C_AA)
    val(sheet, 12, _C_AB, 0.05)
    format_input(sheet, 12, _C_AB)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 13–17, cols X–AC."""
    section_heading(sheet, 13, _C_AA, "ANOVA TABLE")

    for col, header in zip(
        [_C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        val(sheet, 14, col, header)
    bold_row(sheet, 14, _C_AA, _C_AF)

    # SST = SSR + SSE must hold under FE too, so every row reads the SAME
    # fit-time pair (Design_Columns()/Design_Response()) — mixing a raw Total SS against
    # within Regression/Residual SS would break the ANOVA identity.
    val(sheet, 15, _C_AA, "Regression")
    f(sheet, 15, _C_AB, "=Regression_Degrees_Of_Freedom(Design_Columns(),Fit_Context())")
    f(sheet, 15, _C_AC, "=SS_Regression(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 15, _C_AD, "=MS_Regression(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 15, _C_AE, "=F_Statistic(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 15, _C_AF, "=F_Statistic_P_Value(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")

    val(sheet, 16, _C_AA, "Residual")
    f(sheet, 16, _C_AB, "=Residual_Degrees_Of_Freedom(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 16, _C_AC, "=SS_Residual(Design_Columns(),Design_Response(),Sample_Include())")
    f(sheet, 16, _C_AD, "=MS_Residual(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")

    val(sheet, 17, _C_AA, "Total")
    f(sheet, 17, _C_AB, "=Total_Degrees_Of_Freedom(Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 17, _C_AC, "=SS_Total(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")

    sheet.range(rc(15, _C_AB), rc(17, _C_AB)).number_format = "0"
    sheet.range(rc(15, _C_AC), rc(17, _C_AC)).number_format = "0.0"
    sheet.range(rc(15, _C_AD), rc(16, _C_AD)).number_format = "0.0"
    sheet.range(rc(15, _C_AE), rc(15, _C_AE)).number_format = "0.0"
    sheet.range(rc(15, _C_AF), rc(15, _C_AF)).number_format = "0.0E+00"
    border_box(sheet, 13, _C_AA, 17, _C_AF)


def _write_coefficients(sheet: xw.Sheet) -> None:
    """Cols X–AE, rows 19+. Spills downward — nothing placed below row 39 in these cols."""
    section_heading(sheet, 19, _C_AA, "COEFFICIENTS")

    for col, header in zip(
        [_C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF, _C_AG, _C_AH],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%", "Beta Weight"],
    ):
        val(sheet, 20, col, header)
    bold_row(sheet, 20, _C_AA, _C_AH)

    # Spill row labels aligned to the constructed columns (level-qualified via
    # the constructor twin). Zero_Predictors_Selected() branch computes a real
    # intercept-only model (label "Intercept") instead of fabricating a result
    # for an unselected variable; NA() when nothing is fit.
    f(
        sheet,
        21,
        _C_AA,
        '=IF(Zero_Predictors_Selected(),'
        'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),"Intercept",NA()),'
        'IF(Allow_Intercept,'
        'VSTACK("Intercept",TRANSPOSE(Constructed_Column_Names())),'
        'VSTACK("",TRANSPOSE(Constructed_Column_Names()))))',
    )

    # Spill anchors at row 21 — pad with blank top row when intercept is disabled;
    # zero-predictor branch uses the closed-form intercept-only statistic, or
    # NA() when there is nothing to fit. The mean (Y) only needs one observation;
    # SE/t/p/CI (Z-AA-AD) need at least two to estimate variance, so they're guarded
    # separately rather than sharing the N>=1 check used for the mean.
    f(sheet, 21, _C_AB,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),Intercept_Only_Point(),NA()),'
       'IF(Allow_Intercept,Coefficients(Design_Columns(),Design_Response(),Sample_Include()),'
       'VSTACK("",Coefficients(Design_Columns(),Design_Response(),Sample_Include()))))')
    f(sheet, 21, _C_AC,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,SE_Coefficients(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()),'
       'VSTACK("",SE_Coefficients(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()))))')
    f(sheet, 21, _C_AD,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_Point()/Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,T_Statistics(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()),'
       'VSTACK("",T_Statistics(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()))))')
    f(sheet, 21, _C_AE,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'T.DIST.2T(ABS(Intercept_Only_Point()/Intercept_Only_SE()),Intercept_Only_DF()),NA()),'
       'IF(Allow_Intercept,P_Values(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()),'
       'VSTACK("",P_Values(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()))))')
    # Confidence_Interval_Lower/Upper's [Context] sits after [Alpha], and
    # Excel LAMBDA calls cannot skip a middle optional argument — 0.05 is
    # passed explicitly here (matching the function's own internal default
    # bit-for-bit) so [Context] can be reached without changing the
    # pre-existing (Alpha-input-independent) 95% CI behavior.
    f(sheet, 21, _C_AF,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()-T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Lower(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()),'
       'VSTACK("",Confidence_Interval_Lower(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()))))')
    f(sheet, 21, _C_AG,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()+T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Upper(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()),'
       'VSTACK("",Confidence_Interval_Upper(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()))))')
    # Beta Weights: k×1 (no intercept row); always prepend blank to align with other columns.
    # No predictor exists to standardize in the zero-predictor branch, so render
    # blank (not an error) when Allow_Intercept is TRUE; NA() when nothing is fit.
    f(sheet, 21, _C_AH,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,"",NA()),'
       'VSTACK("",Beta_Weights(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())))')

    for col in [_C_AB, _C_AC, _C_AD, _C_AF, _C_AG, _C_AH]:
        sheet.range(
            rc(21, col), rc(_FORMAT_BAND_LAST_ROW, col)
        ).number_format = "0.0000"
    sheet.range(
        rc(21, _C_AE), rc(_FORMAT_BAND_LAST_ROW, _C_AE)
    ).number_format = "0.0E+00"


def _write_unit_space_block(sheet: xw.Sheet) -> None:
    """v3.3 unit-space / back-transformation block at AG3:AH9.

    Sits between the Coefficients spill (rows 19+) and the Prediction Outputs
    zone (AJ1+). Pair the catalog's `Unit_Space_*` LAMBDA functions with the
    Back-Transform Method input on row 4 (default "Duan") so the prediction
    column (AL) and the residual-zone original-units columns (AZ, BA) can
    stitch onto a single source. Reads are gated by the response Transform
    string read off Fit_Context(); the Response Space readout on row 9 makes
    the active state visible at a glance.

    Row layout:

    | Row | AG                | AH                                         |
    |-----|-------------------|--------------------------------------------|
    | 3   | "UNIT-SPACE FIT"  | (section heading merged across AG3:AH3)    |
    | 4   | "Back-Transform"  | input: "Duan" / "Naive" (default "Duan")   |
    | 5   | "Smearing Factor" | =Smearing_Factor(Design_Columns(), ...)    |
    | 6   | "R Square (Unit)" | =Unit_Space_R_Squared(...)                 |
    | 7   | "Adj R Square (Unit)" | =Unit_Space_Adjusted_R_Squared(...)    |
    | 8   | "RMSE (Unit)"     | =Unit_Space_RMSE(...)                      |
    | 9   | "Response Space"  | readout (Fit_Context → "Log"/"None")       |
    """
    section_heading(sheet, 3, _C_AG, "UNIT-SPACE FIT")
    val(sheet, 4, _C_AG, "Back-Transform")
    format_input(sheet, 4, _C_AH)
    # A LITERAL default, never a formula. This cell is an INPUT, and the
    # earlier "=IF(OR($AH$4=..." form read its own address: a circular
    # reference, which Excel resolves to 0 with iterative calculation off,
    # so every consumer below (rows 6-8, AL3, AZ, BA) received an
    # unrecognised method and returned #N/A. Same shape as the Alpha input
    # at AB12 (val(sheet, 12, _C_AB, 0.05)) — a typed value is constrained
    # by the list validation, not by the cell re-deriving itself.
    val(sheet, 4, _C_AH, _BACK_TRANSFORM_DEFAULT)
    # Restrict the input to the two supported methods via list validation.
    try:
        sheet.range(rc(4, _C_AH)).api.Validation.Delete()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    try:
        sheet.range(rc(4, _C_AH)).api.Validation.Add(
            Type=3,  # xlValidateList
            AlertStyle=1,
            Formula1=f'"{",".join(_BACK_TRANSFORM_METHODS)}"',
        )
        # IgnoreBlank would let a cleared cell through, and a blank method is
        # not one of the six recognised states — it would silently #N/A the
        # whole block rather than being rejected at entry.
        sheet.range(rc(4, _C_AH)).api.Validation.IgnoreBlank = False
    except Exception:  # pylint: disable=broad-exception-caught
        pass

    for row, label, formula in [
        (
            5,
            "Smearing Factor",
            (
                "=Smearing_Factor(Design_Columns(),Design_Response(),"
                "Sample_Include(),Fit_Context())"
            ),
        ),
        (
            6,
            "R Square (Unit)",
            (
                "=Unit_Space_R_Squared(Design_Columns(),Design_Response(),"
                "Response_Column(),Sample_Include(),Fit_Context(),"
                f"{_A_BACK_TRANSFORM_METHOD})"
            ),
        ),
        (
            7,
            "Adj R Square (Unit)",
            (
                "=Unit_Space_Adjusted_R_Squared(Design_Columns(),Design_Response(),"
                "Response_Column(),Sample_Include(),Fit_Context(),"
                f"{_A_BACK_TRANSFORM_METHOD})"
            ),
        ),
        (
            8,
            "RMSE (Unit)",
            (
                "=Unit_Space_RMSE(Design_Columns(),Design_Response(),"
                "Response_Column(),Sample_Include(),Fit_Context(),"
                f"{_A_BACK_TRANSFORM_METHOD})"
            ),
        ),
    ]:
        val(sheet, row, _C_AG, label)
        f(sheet, row, _C_AH, formula)
        sheet.range(rc(row, _C_AH), rc(row, _C_AH)).number_format = "0.0000"

    val(sheet, 9, _C_AG, "Response Space")
    f(
        sheet,
        9,
        _C_AH,
        '=IF(Context_Response_Transform(Fit_Context())="Log",'
        '"Original units (back-transformed)","Same as fit space")',
    )
    border_box(sheet, 3, _C_AG, 9, _C_AH)


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    """Zone AJ1:AK14: boxed prediction interval output, plus the FE group
    selector and its group-mean/count readouts.

    v2.1 Fixed Effects group-mean recovery (DECISIONS.md "FE point
    prediction" / "prediction interval"): rebuilt from the v2.0 single-CI
    6-row block into 9 rows surfacing BOTH a mean-response CI and a
    new-observation PI (same center, same t-critical, differing by one
    variance term), via Group_Prediction_Interval — the group-mean-recovery
    sibling of Prediction_Interval. With no Fixed Effects row declared,
    Prediction_Group_Column() degenerates to a constant "(all)" group and
    this collapses EXACTLY to the old Prediction_Interval() numbers (verified
    in tests/test_group_prediction_interval.py to floating-point precision) —
    the "build it once" property the whole v2.1 FE design relies on.
    """
    section_heading(sheet, 1, _C_AJ, "PREDICTION OUTPUTS")
    val(sheet, 2, _C_AJ, "PREDICTION INTERVAL")
    bold(sheet, 2, _C_AJ)
    # Sub-headers row 2: AK is "Fit Space" (the log/identity values above),
    # AL is "Original Units" (the back-transformed sibling). Both are
    # unmerged so they read as the column header for the point estimate that
    # follows directly underneath.
    val(sheet, 2, _C_AK, "Fit Space")
    val(sheet, 2, _C_AL, "Original Units")
    bold_row(sheet, 2, _C_AJ, _C_AL)
    for row, label in [
        (3, "Point Estimate"),
        (4, "SE (Mean)"),
        (5, "SE (New Obs)"),
        (6, "t Critical"),
        (7, "CI Lower"),
        (8, "CI Upper"),
        (9, "PI Lower"),
        (10, "PI Upper"),
        (11, "Confidence Level"),
    ]:
        val(sheet, row, _C_AJ, label)
    # Zero_Predictors_Selected() branch: the closed-form intercept-only
    # model, now split into its own mean-CI (S/sqrt(n)) and new-obs-PI
    # (S*sqrt(1+1/n)) the same way the live branch does; NA() when nothing
    # is fit. The live branch takes exactly COLUMNS(Predictor_Columns()) raw predictor
    # values from the Prediction Inputs band — no intercept slot, since
    # group-mean recovery never uses one (the selected group's own mean
    # plays that role) — then auto-logs whichever of those values belong to
    # a Log-transformed Continuous predictor before the call:
    # Constructed_Column_Transforms() gives the per-constructed-column
    # Log/None flag (not a raw spec-row flag, since a Categorical Predictor
    # contributes a variable number of dummy columns that must never be
    # logged), TRANSPOSE'd to match the AK band's column-vector shape, and
    # Ln_Positive is applied elementwise only where it reads "Log". The
    # user always types a raw value in AK (e.g. actual miles), never ln(x).
    f(
        sheet,
        3,
        _C_AK,
        "=IF(Zero_Predictors_Selected(),"
        "IF(AND(Allow_Intercept,Intercept_Only_N()>=2),"
        "LET(point,Intercept_Only_Point(),"
        "se_mean,Intercept_Only_S()/SQRT(Intercept_Only_N()),"
        "se_new,Intercept_Only_S()*SQRT(1+1/Intercept_Only_N()),"
        "t_crit,T.INV.2T(alpha,Intercept_Only_DF()),"
        "VSTACK(point,se_mean,se_new,t_crit,"
        "point-t_crit*se_mean,point+t_crit*se_mean,"
        "point-t_crit*se_new,point+t_crit*se_new,1-alpha)),"
        "NA()),"
        f"LET(raw,TAKE({_abs_ref(_PRED_INPUT_FIRST_ROW, _C_AK)}:"
        f"{_abs_ref(_PRED_INPUT_LAST_ROW, _C_AK)},COLUMNS(Predictor_Columns())),"
        "trn,TRANSPOSE(Constructed_Column_Transforms()),"
        'pred_input,IF(trn="Log",Ln_Positive(raw),raw),'
        "Group_Prediction_Interval(Predictor_Columns(),Response_Column(),pred_input,"
        f"Prediction_Group_Column(),{_A_FE_GROUP},"
        "Sample_Include(),alpha,Fit_Context())))",
    )
    sheet.range(rc(3, _C_AK), rc(11, _C_AK)).number_format = "0.0000"

    # Original Units column (AL): the back-transformed sibling of AK under a
    # Log response. Point estimate (row 3) honors the Method toggle (Duan =
    # conditional mean via smearing factor; Naive = textbook EXP(ŷ)). CI/PI
    # bounds (rows 7–10) ALWAYS use Naive — bounds are quantiles, so EXP(L)
    # to EXP(U) preserves the right coverage; multiplying by smearing would
    # destroy it. Rows 4–6 (SE and t-critical) have no unit-space counterpart
    # and stay blank.
    f(
        sheet,
        3,
        _C_AL,
        (
            "=Back_Transform_Response("
            f"{_abs_ref(3, _C_AK)},"
            "Fit_Context(),"
            f"{_A_BACK_TRANSFORM_METHOD},"
            "Smearing_Factor(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())"
            ")"
        ),
    )
    for row in (7, 8, 9, 10):
        f(
            sheet,
            row,
            _C_AL,
            (
                "=Back_Transform_Response("
                f"{_abs_ref(row, _C_AK)},"
                "Fit_Context(),"
                '"Naive",'
                "1)"
            ),
        )
    sheet.range(rc(3, _C_AL), rc(3, _C_AL)).number_format = "0.0000"
    sheet.range(rc(7, _C_AL), rc(10, _C_AL)).number_format = "0.0000"

    # FE Group selector (row 12): computed-with-override, the same
    # reference-level pattern as the Categorical Reference Level (E) —
    # pre-filled with the alphabetically-first observed group (which is
    # always the "(all)" sentinel itself when no Fixed Effects row is
    # declared, since Prediction_Group_Column() is constant in that state),
    # and directly editable to any other observed group. Red CF flags a
    # typed value that is not among the observed groups.
    val(sheet, 12, _C_AJ, "FE Group")
    f(
        sheet,
        12,
        _C_AK,
        "=INDEX(SORT(UNIQUE(FILTER(Prediction_Group_Column(),Sample_Include()))),1,1)",
    )
    format_input(sheet, 12, _C_AK)
    fe_group_cell = f"${col_letter(_C_AK)}$12"
    add_expression_format(
        sheet,
        fe_group_cell,
        f"=ISNA(MATCH({fe_group_cell},Prediction_Group_Column(),0))",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Group Mean (y) / Group Count (rows 13-14): the ȳᵢ / Tᵢ readouts
    # DECISIONS.md calls for, computed on the selected group directly via
    # Group_Mean_At / Group_Count_At — the same primitives
    # Group_Prediction_Interval uses internally, so these never disagree
    # with what the interval above actually used.
    val(sheet, 13, _C_AJ, "Group Mean (y)")
    f(
        sheet,
        13,
        _C_AK,
        f"=Group_Mean_At(Response_Column(),Prediction_Group_Column(),"
        f"{_A_FE_GROUP},Sample_Include())",
    )
    val(sheet, 14, _C_AJ, "Group Count")
    f(
        sheet,
        14,
        _C_AK,
        f"=Group_Count_At(Prediction_Group_Column(),{_A_FE_GROUP},Sample_Include())",
    )
    sheet.range(rc(13, _C_AK), rc(13, _C_AK)).number_format = "0.0000"
    sheet.range(rc(14, _C_AK), rc(14, _C_AK)).number_format = "0"

    # Caveat row 15: explains the asymmetric CI/PI placement under Duan
    # smearing. Always present (the text is honest under either method) so
    # the user sees the gap between the point estimate and the bounds when
    # the Method toggle is Duan — the visible "off-centre" placement is the
    # plan's whole point.
    sheet.range(rc(15, _C_AJ), rc(15, _C_AL)).merge()
    f(
        sheet,
        15,
        _C_AJ,
        (
            '="Duan = Duan (1983) smearing — estimates the conditional mean. '
            "Naive = textbook EXP(ŷ) — the conditional median, biased for the "
            "mean. CI/PI bounds are back-transformed with EXP alone, so under "
            'Duan the point estimate does not sit at the interval centre."'
        ),
    )
    sheet.range(rc(15, _C_AJ), rc(15, _C_AL)).api.WrapText = True

    border_box(sheet, 1, _C_AJ, 15, _C_AL)


def _write_prediction_inputs(sheet: xw.Sheet) -> None:
    """Zone AJ16+: one prediction-input row per constructed column.

    No Intercept row here (unlike the pre-v2.1 layout): Group_Prediction_Interval's
    pred_input is exactly COLUMNS(Predictor_Columns()) raw predictor values with no
    intercept slot — group-mean recovery never uses one (the selected
    group's own mean plays that role) — so the row that used to hold it
    would be actively misleading now. Row 18 is a blank spacer between the
    headers and the first predictor row.
    """
    section_heading(sheet, 16, _C_AJ, "PREDICTION INPUTS")
    val(sheet, 17, _C_AJ, "Predictor")
    val(sheet, 17, _C_AK, "Prediction Value")
    bold_row(sheet, 17, _C_AJ, _C_AL)
    # The band's header note lives in _write_model_specification alongside
    # every other note: AddComment is a COM-only call, and keeping it out of
    # here is what lets this writer stay exercisable through RecordingSheet.

    # AJ19: spill formula — level-qualified names, one per constructed column
    f(sheet, _PRED_INPUT_FIRST_ROW, _C_AJ, "=TRANSPOSE(Constructed_Column_Names())")

    # AL19: the Training Mean column — per-column means of the filtered design
    # matrix, computed with a SINGLE Predictor_Columns() evaluation. Predictor_Columns() is
    # a full
    # design-matrix construction on every call (Excel does not cache LAMBDA
    # results), so this spill is the one place the means are computed; the
    # orange prefill cells INDEX into it. The earlier design — every prefill
    # cell calling Predictor_Columns() twice (width guard + column mean) — made the
    # workbook's first full calculation pathological (~20 minutes at the save
    # step). The spill owns column AL downward (the AM gap and residual zone
    # start beyond it), so a wider dataset or spec can never make it collide
    # with another spill.
    # Degrades to "" on an empty model, which the prefill guard reads as a
    # one-row spill holding a blank.
    #
    # Log columns need the GEOMETRIC mean here, not the arithmetic mean of
    # the already-logged Predictor_Columns() column: the AH prefill cells below just
    # INDEX into this spill, and the row-3 prediction formula applies
    # Ln_Positive to whatever it finds in AK — so if this spill held the
    # log-space arithmetic mean, the default prediction would silently
    # double-log (ln(ln(x)), not merely un-back-transformed). EXP(mean(ln
    # x)) is exact and self-cancelling: Ln_Positive(EXP(mean(ln x))) =
    # mean(ln x), so the default prediction still lands precisely on
    # Predictor_Columns()'s own centroid, unchanged from the pre-Log-transform behavior.
    # Constructed_Column_Transforms() gives the per-column Log/None flag in
    # the same 1xk shape as the BYCOL means row, so the two combine
    # elementwise before the single TRANSPOSE down into the AL column.
    val(sheet, 17, _C_AL, "Training Mean")
    means_anchor = _abs_ref(_PRED_INPUT_FIRST_ROW, _C_AL)
    f(
        sheet,
        _PRED_INPUT_FIRST_ROW,
        _C_AL,
        (
            "=IFERROR(TRANSPOSE(LET(m,BYCOL(FILTER(Predictor_Columns(),Sample_Include()),"
            "LAMBDA(c,AVERAGE(c))),t,Constructed_Column_Transforms(),"
            'IF(t="Log",EXP(m),m))),"")'
        ),
    )

    # AK19:AK62 — the Training Mean of each constructed column, individually
    # overridable. Each row guards on its position against the means-spill
    # height so rows beyond the live constructed width render blank (the
    # width is spec-dependent — 19 on the default WHO spec). Cheap spill
    # references only: no prefill cell may call Predictor_Columns() itself.
    offset = _PRED_INPUT_FIRST_ROW - 1
    for row in range(_PRED_INPUT_FIRST_ROW, _PRED_INPUT_LAST_ROW + 1):
        f(
            sheet,
            row,
            _C_AK,
            (
                f"=IF(ROW()-{offset}<=IFERROR(ROWS({means_anchor}#),0),"
                f"INDEX({means_anchor}#,ROW()-{offset}),"
                '"")'
            ),
        )

    # Orange for all user-editable prediction value cells (the Training Mean
    # column is computed display, not input)
    _input_range(sheet, _PRED_INPUT_FIRST_ROW, _C_AK, _PRED_INPUT_LAST_ROW, _C_AK)
    sheet.range(
        rc(_PRED_INPUT_FIRST_ROW, _C_AK), rc(_PRED_INPUT_LAST_ROW, _C_AK)
    ).number_format = "0.0000"
    # Column-wide: the means spill height is spec-dependent and the column
    # holds nothing else.
    sheet.range(f"{col_letter(_C_AL)}:{col_letter(_C_AL)}").number_format = "0.0000"


def _write_residuals(sheet: xw.Sheet) -> None:
    """Residual diagnostic table — row identifiers + 11 diagnostics columns starting at AO."""
    section_heading(sheet, 1, _C_AN, "RESIDUAL OUTPUT")

    # AN2: static header — Row_Labels() supplies its own per-row content
    # (joined Identifier columns, or positional Obs. n labels).
    val(sheet, 2, _C_AN, "Observation")

    # Every one of these columns is fit off Design_Columns()/Design_Response(), so once a
    # Fixed Effects row is declared they hold within-demeaned quantities, not
    # the raw response — the header must say so or the table reads as if
    # "Y" - "Predicted Y" silently stopped equaling "Residuals" (see the
    # comment above the AL formula below). Each header is an IF conditional
    # on the same FE-count gate the DW/BFN trigger cells use, so the label
    # flips the moment a spec declares (or removes) a Fixed Effects variable.
    #
    # Response-scale columns (Y, Predicted Y, Residuals, PRESS Residual) get
    # a second, independent "(Log)" suffix when the Response row's Transform
    # is Log — those four are literally in response units (or a difference
    # of two response-unit values), so a log-transformed response leaves
    # them in log space, unlike the Within case they don't silently change
    # meaning, but they DO stop being in the response's original units and
    # must say so. The remaining six columns (Hat Diagonal, Studentized
    # Residuals, Cook's Distance, Normal Scores Ranked, Studentized
    # Residuals Ranked, Scale-Location) are dimensionless/standardized
    # diagnostics — not in response units either way — and do not get the
    # suffix.
    #
    # The Within suffix names the active Fixed Effects variable (e.g.
    # "(Within Country)") rather than the bare "(Within)" token, via
    # _FIXED_EFFECTS_NAME_FORMULA — the same Role=Fixed Effects lookup that
    # fills the spec feedback block's "FE Variable" cell. Y (_C_AO) gets its
    # own wording instead of "Within": Design_Columns()/Design_Response() only SUBTRACT the
    # group mean, they never divide by a standard deviation, so "Y (Within
    # Country)" would read as a demeaning but "St Devs from Avg." would be
    # outright wrong — "Deviation from <FE> Avg." says exactly what the
    # column holds (still response-scale units, just group-centered).
    response_scale_headers = {_C_AO, _C_AP, _C_AQ, _C_AX}
    fe_within_suffix = f'" (Within "&{_FIXED_EFFECTS_NAME_FORMULA}&")"'
    fe_within_log_suffix = f'" (Within "&{_FIXED_EFFECTS_NAME_FORMULA}&", Log)"'
    fe_deviation_suffix = f'" (Deviation from "&{_FIXED_EFFECTS_NAME_FORMULA}&" Avg.)"'
    fe_deviation_log_suffix = (
        f'" (Deviation from "&{_FIXED_EFFECTS_NAME_FORMULA}&" Avg., Log)"'
    )
    for col, header in zip(
        [_C_AO, _C_AP, _C_AQ, _C_AR, _C_AS, _C_AT, _C_AU, _C_AV, _C_AW, _C_AX, _C_AY],
        [
            "Y", "Predicted Y", "Residuals",
            "Hat Diagonal", "Studentized Residuals", "Cook's Distance",
            "Normal Scores Ranked", "Studentized Residuals Ranked",
            "Scale-Location", "PRESS Residual", "Cook's Distance (Flagged)",
        ],
    ):
        if col == _C_AO:
            within_suffix, within_log_suffix = fe_deviation_suffix, fe_deviation_log_suffix
        else:
            within_suffix, within_log_suffix = fe_within_suffix, fe_within_log_suffix
        if col in response_scale_headers:
            formula = (
                f'=IF(AND({_FIXED_EFFECTS_COUNT_FORMULA}>0,{_RESPONSE_LOG_FORMULA}),'
                f'"{header}"&{within_log_suffix},'
                f'IF({_FIXED_EFFECTS_COUNT_FORMULA}>0,"{header}"&{within_suffix},'
                f'IF({_RESPONSE_LOG_FORMULA},"{header} (Log)","{header}")))'
            )
        else:
            formula = (
                f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}>0,"{header}"&{within_suffix},"{header}")'
            )
        f(sheet, 2, col, formula)
    # v3.3 unit-space columns: AZ/BA are the original-units siblings of the
    # fit-space Predicted Y / Residuals columns. They carry no (Log) or
    # (Within …) suffix — they are in original units by construction, so the
    # conditional suffix logic above does not apply to them.
    val(sheet, 2, _C_AZ, "Predicted Y (Original Units)")
    val(sheet, 2, _C_BA, "Residual (Original Units)")
    bold_row(sheet, 2, _C_AN, _C_BA)

    # AN3: row labels — the spec-derived Row_Labels() filtered to the sample.
    # Row_Labels() has its own no-Identifier fallback ("Obs. n"), so the only
    # error left to absorb is an all-FALSE mask.
    f(
        sheet, 3, _C_AN,
        "=IFERROR(FILTER(Row_Labels(),Sample_Include()),NA())",
    )
    # Spill anchors — each spills n rows downward. Fit-time Design_Columns()/Design_Response()
    # throughout, INCLUDING the "Y" column (AL): under FE the whole table
    # must read as one internally consistent block — Residuals (AN) is an
    # independently-computed column, not a literal AL-AM subtraction, but a
    # raw "Y" next to a within-fitted "Predicted Y" would make the table look
    # broken (Residuals would not visually match Y - Predicted Y). The actual
    # observed response is still available via Response_Column() elsewhere
    # (e.g. Intercept_Only_*); this table shows the model's own fit space.
    f(sheet, 3, _C_AO, "=Dependent_Variable(Design_Response(),Sample_Include())")
    f(sheet, 3, _C_AP, "=Predictions(Design_Columns(),Design_Response(),Sample_Include())")
    f(sheet, 3, _C_AQ, "=Residuals(Design_Columns(),Design_Response(),Sample_Include())")
    f(sheet, 3, _C_AR, "=Hat_Diagonal(Design_Columns(),Sample_Include())")
    f(sheet, 3, _C_AS, "=Studentized_Residuals(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 3, _C_AT, "=Cooks_Distance(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 3, _C_AU, "=SORT(Normal_Scores(Design_Response(),Sample_Include()))")
    f(sheet, 3, _C_AV, "=Studentized_Residuals_Ranked(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    # Scale-Location: SQRT(|Studentized_Residuals|) — horizontal spread should be flat.
    f(
        sheet, 3, _C_AW,
        "=SQRT(ABS(Studentized_Residuals(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())))",
    )
    # PRESS Residual equals the leave-one-out residual e_i / (1 - h_i).
    f(sheet, 3, _C_AX, "=LOOCV_Residual(Design_Columns(),Design_Response(),Sample_Include())")
    # Cook's Distance (Flagged): NA()'d except where D exceeds the standard
    # influence cutoffs (4/n or 0.9 — algebraically MIN(4/n, 0.9) since these
    # are two lower bounds on the same "flag it" union, for every n). This
    # feeds the Cook's Distance chart's data-label overlay series (see
    # RegChartCookDistFlag in _setup_local_names / _write_diagnostic_charts):
    # NA() points are skipped for both plotting and labeling, so only the
    # flagged points get a label.
    cooks_col = col_letter(_C_AT)
    f(
        sheet, 3, _C_AY,
        f"=IF({cooks_col}3#>MIN(4/{_A_OBSERVATIONS},0.9),{cooks_col}3#,NA())",
    )
    # AZ: Predicted Y in original units — the unit-space siblings of AP/AR/AQ
    # /AX. Method defaults to Duan smearing when the response is Log, which
    # lifts EXP(ŷ) from a median predictor to a mean predictor; under
    # `None` it reduces to the same Predicted Y as AP. BA: Residual in
    # original units = observed response in original units minus the unit-
    # space predicted. Both carry the Method toggle from `AH4` so flipping
    # Duan → Naive propagates here.
    f(
        sheet, 3, _C_AZ,
        (
            "=Unit_Space_Predictions(Design_Columns(),Design_Response(),"
            "Response_Column(),Sample_Include(),Fit_Context(),"
            f"{_A_BACK_TRANSFORM_METHOD})"
        ),
    )
    f(
        sheet, 3, _C_BA,
        (
            "=Unit_Space_Residuals(Design_Columns(),Design_Response(),"
            "Response_Column(),Sample_Include(),Fit_Context(),"
            f"{_A_BACK_TRANSFORM_METHOD})"
        ),
    )
    # Format every numeric residual-output column — the actual Y (AO) through
    # Residual (Original Units) (BA). Only the AN identifier column (text:
    # country/Obs. labels) is left unformatted.
    sheet.range(f"{col_letter(_C_AO)}:{col_letter(_C_BA)}").number_format = "0.0000"


def _diagnostic_chart_specs() -> list[tuple[str, str, str | None, str, str, str, str, int, int]]:
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
    """
    sname = REGRESSION_SHEET_NAME

    def _name_ref(local_name: str) -> str:
        return f"='{sname}'!{local_name}"

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
            '="Cook\'s Distance  (flag: D > "'
            f'&TEXT(MIN(4/{_A_OBSERVATIONS},0.9),"0.000")&")"',
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
            f'="PRESS Residuals  (PRESS = "&TEXT({_A_PRESS},"#,##0.0000")&")"',
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
    for i, spec in enumerate(_diagnostic_chart_specs()):
        key, _chart_type, _x_addr, _y_addr, title_formula, x_label_formula, y_label_formula, _grid_row, _grid_col = spec
        row = _ROW_CHART_LABELS + i
        val(sheet, row, _C_CHART_LABEL_NAME, key)
        f(sheet, row, _C_CHART_TITLE, title_formula)
        f(sheet, row, _C_CHART_XLABEL, x_label_formula)
        f(sheet, row, _C_CHART_YLABEL, y_label_formula)


def _write_materialization_zone(
    sheet: xw.Sheet,
    closures: tuple[CatalogFunction, ...] | None = None,
) -> None:
    """Write the ARCHITECTURE §4b materialization band at the sheet's far right.

    Two names share the model context, by design (v3.0 stage two):

    ``Model_Context(...)`` is the WORKBOOK-scoped constructor — the one
    definition of the default context (Has_Intercept=TRUE, DF_Absorbed=0,
    both transforms "None"). It takes optional overrides, is what every
    engine's omitted-[Context] default routes through, and is what the MLR
    test sheets call with a per-row flag.

    ``Fit_Context()`` is the SHEET-scoped reader — a zero-arg thunk over the
    FIXED range holding the materialized context block. It is what the ~30
    Regression sheet call sites pass, so they read the actual spec-derived
    context (the C2 Allow_Intercept toggle, the Absorbed_Degrees_Of_Freedom()
    closure, the spec-block transform summaries) rather than the constructor
    default. Splitting the names keeps ``Model_Context`` unshadowed: a single
    sheet-scoped thunk named ``Model_Context`` would make ``Model_Context()``
    in a sheet cell resolve to the materialized values while the same token in
    a carrier's omitted-default resolved to the workbook constructor — the
    invisible shadowing the v3.0 release exists to remove.

    The context is materialized ONCE (Excel does not memoize a name whose
    RefersTo is a formula, so a constructor inside ~30 engine calls runs ~30
    times); ``Fit_Context`` reads the fixed range those cells occupy, so the
    ~30 call sites that pass ``Fit_Context()`` all read the one materialized
    block.

    The block is written as ONE CELL PER ELEMENT, each with its own label in
    the column to its left, and boxed — it is a fixed-size table, not a data
    range. It used to be a single ``VSTACK`` spill, which bought nothing (the
    height is a build-time constant) and cost correctness: a spill is one
    dependency node that Excel vacates and re-spills whenever the spec block
    changes, and while it is vacated the range behind ``Fit_Context`` is
    transiently blank, so every engine reading it sees a torn context. Four
    independent cells recalculate independently and are never vacated.

    Elements 3-4 (Response_Transform, Predictor_Transform) are populated from
    the spec block now but have no engine reader until the v3.3 unit-space
    dispatcher; the row order is the contract that is expensive to change
    later, so all four rows land together. An error in an unconsumed row is
    contained to its own cell now that each element stands alone, and the
    engines read only elements 1-2 through the accessors, so a bad name
    surfaces as a visible cell error (caught by the headless verifier and by
    the block's own health-check row) without shifting a single fitted number.

    The two data-dependent zones — the ``Sample_Include`` row mask and the
    terminal Constructed Design Matrix — are SURFACED here: each carries a
    column-header row at ``_MATERIALIZATION_HEADER_ROW`` and spills from
    ``_MATERIALIZATION_SPILL_ROW``, full height and row-aligned with the source
    table, so the mask reads straight across into the design-matrix row beside
    it. Both replaced a ``"reserved"`` placeholder that held the position while
    stage 3 established the layout, the outline behaviour, and the pre-flight
    width guard.

    Surfacing the values is NOT the same as rewiring the readers, and only the
    first half lands here. ``Sample_Include()`` and ``Design_Columns()`` remain
    live closures evaluated per call site; promoting either to a thunk over its
    spill is still deferred, because that needs the dynamic-array spill
    operator (``#``) inside a ``LAMBDA`` defined-name RefersTo — a combination
    used nowhere else in this workbook and verifiable only with Excel present.
    A wrong guess would break the row-mask contract that keeps every spilled
    array row-aligned, so the optimization lands separately, Excel-verified
    rather than blind. Nothing on the sheet reads these two spills today.

    The design matrix's header row is split across two cells:
    ``Design_Columns()`` is one column wider than
    ``Constructed_Column_Names()`` when the intercept is on, because the
    constructor prepends the ones column, so the anchor cell names that column
    and the names spill starts beside it. With ``Allow_Intercept`` FALSE there
    is no ones column and the names sit one column right of the values they
    label — a labelling offset in an otherwise unread display zone, called out
    on the heading cell's note.

    Plain cell writes + defined-name registration only; no chart/COM API, so
    this is exercised in unit tests via ``RecordingSheet`` without Excel. The
    chart-footprint clearance check is the one Excel-only step and is guarded.
    """
    sname = sheet.name
    if closures is None:
        closures = load_catalog_document(_DEFINITIONS_PATH).functions_for_sheet(
            REGRESSION_SHEET_NAME
        )

    # ── Model Context (fixed-size labelled block) ────────────────────────────
    # One row per element, label left of value, heading on row 1, the whole
    # thing boxed — the same treatment every other fixed-size block on this
    # sheet gets (Regression Statistics, Diagnostics, Prediction Interval).
    # Element 1 is the C2 Allow_Intercept toggle (named range). Element 2 is
    # the Absorbed_Degrees_Of_Freedom() Regression closure. Elements 3-4 are
    # the spec-block transform summaries (see _RESPONSE_TRANSFORM_FORMULA /
    # _PREDICTOR_TRANSFORM_FORMULA), both reading worksheet-scoped spec names
    # on this sheet. _MODEL_CONTEXT_ELEMENTS is the single source of the row
    # order, the labels, and the height.
    ctx_col = col_letter(_C_MODEL_CONTEXT)
    section_heading(sheet, 1, _C_MODEL_CONTEXT_LABEL, "MODEL CONTEXT")
    # The value column's heading cell carries the fill but no text, so the
    # heading reads as one banner across the two-column block.
    section_heading(sheet, 1, _C_MODEL_CONTEXT, "")

    for offset, element in enumerate(_MODEL_CONTEXT_ELEMENTS):
        row = _MATERIALIZATION_FIRST_ROW + offset
        val(sheet, row, _C_MODEL_CONTEXT_LABEL, element.label)
        f(sheet, row, _C_MODEL_CONTEXT, f"={element.formula}")

    # The materialized block occupies _MATERIALIZATION_FIRST_ROW ..
    # _MODEL_CONTEXT_LAST_ROW; the sheet-scoped reader Fit_Context reads that
    # fixed range (no spill operator — the height is a structural constant, so
    # a fixed range is exact and avoids the dynamic-array-in-a-name question
    # entirely). Drop both the legacy "Model_Context" sheet name (left by a
    # pre-split build) and any stale "Fit_Context" before re-adding, so a
    # rebuild never leaves a shadow.
    ctx_ref = (
        f"'{sname}'!${ctx_col}${_MATERIALIZATION_FIRST_ROW}"
        f":${ctx_col}${_MODEL_CONTEXT_LAST_ROW}"
    )
    drop_local_name(sheet, "Model_Context")
    drop_local_name(sheet, "Fit_Context")
    sheet.api.Names.Add(
        Name="Fit_Context",
        RefersTo=f"=LAMBDA({ctx_ref})",
    )
    # Health check, one row under the block and inside its box. The height
    # half is the build-time invariant (_MODEL_CONTEXT_ROWS); the error half
    # is what decomposition made worth checking — with four independent cells
    # a broken spec name errors in ONE of them and leaves the other three
    # looking fine, so the block reports whether every element resolved.
    val(sheet, _ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT_LABEL, "Context OK")
    f(
        sheet,
        _ROW_MODEL_CONTEXT_CHECK,
        _C_MODEL_CONTEXT,
        f"=AND(ROWS(Fit_Context())={_MODEL_CONTEXT_ROWS},"
        "SUMPRODUCT(--ISERROR(Fit_Context()))=0)",
    )
    border_box(
        sheet, 1, _C_MODEL_CONTEXT_LABEL, _ROW_MODEL_CONTEXT_CHECK, _C_MODEL_CONTEXT
    )

    # ── Sample_Include (materialized row mask) ───────────────────────────────
    # The mask spills full-height and row-aligned with the source table (the
    # row-mask contract), so it reads straight across into the design-matrix
    # rows beside it. This SURFACES the value; it does not rewire the closure —
    # Sample_Include() is still evaluated per call site, and promoting it to a
    # thunk over this spill stays deferred (see the docstring).
    section_heading(sheet, 1, _C_SAMPLE_INCLUDE_MATERIALIZED, "Sample Include")
    val(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        _SAMPLE_INCLUDE_HEADER,
    )
    bold_row(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
    )
    f(
        sheet,
        _MATERIALIZATION_SPILL_ROW,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        "=Sample_Include()",
    )
    # Document what the column is for on the heading cell — it is a read-only
    # view of the mask every engine applies, not a second place to edit it.
    try:
        sheet.range(rc(1, _C_SAMPLE_INCLUDE_MATERIALIZED)).api.AddComment(
            "TRUE for every source row the model is fitted on. Read-only view "
            "of the row mask the engines apply — change it through the spec "
            "block (Include, and any Role=Filter column), not here. Full "
            "height and row-aligned with the source table, so it reads "
            "straight across into the design matrix to its right."
        )
    except Exception:  # pylint: disable=broad-except
        pass

    # ── Constructed Design Matrix (terminal zone) ────────────────────────────
    # The zone that terminates the band. Its width is unbounded and one
    # dropdown away — Country as a Categorical Predictor is 156 columns, and
    # interactions multiply — which is why nothing may ever be placed to its
    # right, and why it ships COLLAPSED while the two bounded zones ship
    # expanded: an unbounded-width zone that cannot be collapsed is a
    # scrolling hazard.
    #
    # Establishing the zone and MATERIALIZING into it were deliberately
    # separate steps: the position, the outline behaviour, and the width guard
    # that reads the spec's pre-flight column count are what a later release
    # could not add without moving columns again, so they landed first and the
    # spill is a formula change into columns that already exist.
    #
    # The header row is split across two cells because the matrix is one
    # column wider than its names when the intercept is on — the anchor cell
    # names the constructor's ones column and the names spill starts beside
    # it. See _DESIGN_MATRIX_INTERCEPT_HEADER.
    section_heading(sheet, 1, _C_DESIGN_MATRIX, "Constructed Design Matrix")
    f(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_DESIGN_MATRIX,
        _DESIGN_MATRIX_INTERCEPT_HEADER,
    )
    f(
        sheet,
        _MATERIALIZATION_HEADER_ROW,
        _C_DESIGN_MATRIX_NAMES,
        "=Constructed_Column_Names()",
    )
    bold_row(
        sheet, _MATERIALIZATION_HEADER_ROW, _C_DESIGN_MATRIX, _C_DESIGN_MATRIX_NAMES
    )
    f(sheet, _MATERIALIZATION_SPILL_ROW, _C_DESIGN_MATRIX, "=Design_Columns()")
    try:
        sheet.range(rc(1, _C_DESIGN_MATRIX)).api.AddComment(
            "The design matrix the engines actually fit: Design_Columns(), "
            "full height and row-aligned with the source table and with the "
            "Sample Include mask to its left (the engines apply that mask "
            "themselves). Terminal §4b zone — nothing may ever be placed to "
            "its right, because its width is unbounded and one dropdown away. "
            "Ships collapsed for that reason.\n\n"
            "Header row: the anchor cell names the intercept column the "
            "constructor prepends, and Constructed_Column_Names() spills from "
            "the column beside it. With Allow Intercept FALSE there is no "
            "ones column, so the names sit one column right of the values "
            "they label."
        )
    except Exception:  # pylint: disable=broad-except
        pass

    # Both data-dependent zones spill from the same row so the band reads
    # across — the mask value beside its design-matrix row, both aligned to
    # the source table rows. Asserted rather than merely intended.
    assert _MATERIALIZATION_SPILL_ROW == _MATERIALIZATION_HEADER_ROW + 1
    assert _MATERIALIZATION_FIRST_ROW == 2

    # ── Column widths + outline groups ───────────────────────────────────────
    # The two bounded zones ship EXPANDED (per §4b); the terminal design-matrix
    # zone ships COLLAPSED. The width-2 gutters stay ungrouped so the zones
    # collapse independently, and the first gutter (after the charts) is
    # structural — it keeps the floating chart anchors out of every collapsible
    # outline group.
    for gutter in (
        _C_GUTTER_AFTER_CHARTS,
        _C_GUTTER_AFTER_CONTEXT,
        _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    ):
        sheet.range(f"{col_letter(gutter)}:{col_letter(gutter)}").column_width = 2
    for content, width in (
        (_C_MODEL_CONTEXT_LABEL, _MODEL_CONTEXT_LABEL_WIDTH),
        (_C_MODEL_CONTEXT, _MODEL_CONTEXT_VALUE_WIDTH),
        (_C_SAMPLE_INCLUDE_MATERIALIZED, 14),
    ):
        sheet.range(f"{col_letter(content)}:{col_letter(content)}").column_width = width
    # One outline group per ZONE, not per column: the Model Context zone is the
    # label/value pair and has to collapse as a unit, or its labels would be
    # left stranded beside a collapsed value column.
    for first, last in (
        (_C_MODEL_CONTEXT_LABEL, _C_MODEL_CONTEXT),
        (_C_SAMPLE_INCLUDE_MATERIALIZED, _C_SAMPLE_INCLUDE_MATERIALIZED),
    ):
        sheet.api.Columns(f"{col_letter(first)}:{col_letter(last)}").Group()

    matrix_band = (
        f"{col_letter(_C_DESIGN_MATRIX)}:"
        f"{col_letter(_C_DESIGN_MATRIX + _DESIGN_MATRIX_GROUPED_COLUMNS - 1)}"
    )
    sheet.range(matrix_band).column_width = 12
    sheet.api.Columns(matrix_band).Group()
    # Collapse it. ShowDetail is an ActiveWindow-free property on the range,
    # but it still needs a real outline underneath, so guard it the way every
    # other cosmetic COM call on this sheet is guarded — a workbook that
    # opens with the zone expanded is a nuisance, not a broken build.
    try:
        sheet.api.Columns(matrix_band).ShowDetail = False
    except Exception:  # pylint: disable=broad-except
        pass

    # ── Chart-footprint clearance assertion (Excel only) ────────────────────
    # _LAST_CHART_COLUMN is a conservative bound; this verifies the column
    # past the footprint actually clears the computed chart right edge, so a
    # chart resize that would overlap the context block fails the build.
    #
    # The geometry LOOKUP is best-effort — COM geometry (sheet.range(...).left)
    # is unavailable headless, so that raise is swallowed and the check is
    # skipped there (the conservative constant keeps the layout safe by
    # construction). But the clearance ASSERT itself must NOT be swallowed:
    # wrapping it in the same broad except would make the guard a no-op in
    # Excel, the one place it can actually run. So acquire the geometry under
    # the guard, then assert outside it.
    try:
        chart_right = (
            sheet.range(a1(1, _C_BB)).left + _CHART_RIGHT_OFFSET_PT
        )
        clear_left = sheet.range(a1(1, _LAST_CHART_COLUMN + 1)).left
    except Exception:  # pylint: disable=broad-except — headless / no COM geometry
        chart_right = None
        clear_left = None
    if chart_right is not None and clear_left is not None:
        assert clear_left >= chart_right, (
            f"chart footprint ({chart_right:.0f}pt) overlaps the materialization "
            f"zone (column {col_letter(_LAST_CHART_COLUMN + 1)} left edge "
            f"{clear_left:.0f}pt); raise _LAST_CHART_COLUMN"
        )


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

    sname = REGRESSION_SHEET_NAME

    def _name_ref(local_name: str) -> str:
        return f"='{sname}'!{local_name}"

    def _label_ref(col: int, row: int) -> str:
        return f"='{sname}'!${col_letter(col)}${row}"

    chart_specs = _diagnostic_chart_specs()

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
            # the 4/n or 0.9 threshold get a marker+label. ChartType=xlLine
            # (rather than the chart's own xlColumnClustered) keeps this
            # series off the bar cluster — sharing the category axis
            # without narrowing/shifting the real bars — which makes this
            # a Column+Line combo chart.
            flag_series = chart.SeriesCollection().NewSeries()
            flag_series.XValues = _name_ref("RegChartObsLabel")
            flag_series.Values = _name_ref("RegChartCookDistFlag")
            flag_series.ChartType = _XL_LINE
            flag_series.Name = "Flagged (D > 4/n or D > 0.9)"
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


# ── Public entry point ────────────────────────────────────────────────────────

def write_regression_output_sheet(
    workbook: xw.Book,
    sheet_notes: dict[str, str] | None = None,
    closures: tuple[CatalogFunction, ...] | None = None,
    source_table_ref: str = "=MileageData[#All]",
    spec_profile: SpecDatasetProfile | None = None,
) -> None:
    """Create or refresh the spec-driven Regression sheet in workbook.

    Parameters
    ----------
    sheet_notes : dict[str, str] | None
        Mapping of sheet label → plain-language note text from the
        ``regression_sheet_notes`` key in lambda_functions.json.
        Pass ``None`` to skip annotation (useful for isolated tests).
    closures : tuple[CatalogFunction, ...] | None
        The sheet-scoped constructor functions (scope ``"Regression"``), in
        dependency order. When None, they are loaded from
        ``lambda_functions.json``.
    source_table_ref : str, optional
        The ``RefersTo`` formula for the ``Source_Table`` sheet-scoped name.
        Defaults to ``=MileageData[#All]`` (the shipped default dataset).
        Pass ``=LifeExpectancyData[#All]`` to retarget to Life Expectancy Data.
    spec_profile : SpecDatasetProfile | None, optional
        The dataset's default spec-block contents (variable list plus
        default Role/Include/Type/Sequence values) — sizes SpecTable to
        match the targeted dataset's column count and pre-fills a sensible
        starting model instead of leaving every column an un-flagged
        Predictor. Defaults to the Auto MPG profile
        (``SPEC_DATASET_PROFILES["auto_mpg"]``) when omitted, matching the
        shipped default ``source_table_ref``. Callers that retarget
        ``source_table_ref`` to a different dataset should pass the
        matching entry from ``SPEC_DATASET_PROFILES`` here too — the two
        are independent parameters, not derived from each other, so they
        must be kept in sync by the caller (see build_production.py).
    """

    sheet = next(
        (s for s in workbook.sheets if s.name == REGRESSION_SHEET_NAME), None
    )
    if sheet is None:
        sheet = workbook.sheets.add(
            name=REGRESSION_SHEET_NAME, after=workbook.sheets[-1]
        )

    for idx in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(idx).Delete()
    for idx in range(sheet.api.ChartObjects().Count, 0, -1):
        sheet.api.ChartObjects(idx).Delete()
    sheet.api.Cells.Clear()
    # Cells.Clear does not touch outline levels — drop any grouping from a
    # previous build before the zone groups are re-applied below.
    sheet.api.Cells.ClearOutline()
    safe_activate(sheet)

    # The spec block must run before the names are registered: it creates
    # the structured table (SpecTable), which the Spec_* band names bind
    # to via SpecTable[[#Data],[Column]] references — Excel
    # validates the RefersTo at registration time. The rest of the spec
    # area (headers, feedback, intercept) runs in _write_model_specification
    # below, but the table-creating part needs to come first.
    _write_spec_block(sheet, spec_profile or SPEC_DATASET_PROFILES["auto_mpg"])
    _setup_local_names(sheet, closures, source_table_ref=source_table_ref)

    _write_model_specification(sheet)
    _write_predictor_summary(sheet)
    _write_regression_outputs_header(sheet)
    _write_regression_statistics(sheet)
    _write_diagnostics(sheet)
    _write_model_diagnostic_conditional_formatting(sheet)
    _write_alpha(sheet)
    _write_anova(sheet)
    _write_coefficients(sheet)
    _write_significance_conditional_formatting(sheet)
    _write_unit_space_block(sheet)
    _write_prediction_interval(sheet)
    _write_prediction_inputs(sheet)
    _write_residuals(sheet)
    _annotate_statistical_terms(sheet, sheet_notes or {})
    _write_residual_conditional_formatting(sheet)

    sheet.range(rc(2, _C_S), rc(2, _C_BA)).api.WrapText = True

    # A–O (spec block) widths are owned by write_sheet_model_construction.py
    # so the standalone and shared-block builds can never drift.
    _set_spec_block_column_widths(sheet)

    # Content-column widths, per zone, plus the BB post-zone gutter (last entry).
    # The gap columns are sized from _GAP_COLUMNS below so the layout stays
    # declarative — one width there, not one per hard-coded gap letter.
    for column_letter, width in {
        # Spec feedback (M, N, plus the column-I Verdict overlay):
        # the M and N headers are bold on row 1, the I1 "Verdict" header
        # is bold and shares column I with the Sequence_Period spec rows.
        # The "I" width here is the I2 verdict cell — long message; the
        # widest cell on the sheet.
        "M": 10,        # Δ header / spectrum column 1
        "N": 8,         # Count header / spectrum column 2
        "I": 38,        # Verdict header / switch (overlays Sequence_Period column)

        # Predictor Summary (P–V): level-qualified constructed names + 6 stats.
        "P": 24,        # constructed column names (e.g., "Status[Developed]")
        "Q": 9, "R": 9, "S": 9, "T": 9, "U": 9, "V": 9,  # stats values

        # Regression Outputs (X–AE): regression statistics (X–Y), diagnostics (AA–AB),
        # ANOVA (X–AC), coefficients (X–AE), beta weights (AE).
        # Column roles vary by sub-table; widths below are sized for the widest
        # label/value used in each column.
        "X": 22,        # labels — longest is "Adjusted R Square" (16) or "Status[Developing]" (17)
        "Y": 12,        # stat values, Alpha, df, Coefficients
        "Z": 12,        # SS, Std Error
        "AA": 24,       # diagnostics labels (longest: "BFN Panel Durbin-Watson" = 23) + MS + t Stat
        "AB": 16,       # diagnostics values + "Predicted Variable" section heading + F + P-value
        "AC": 14,       # derived response name (e.g., "Life expectancy") + Significance F + Lower 95%
        "AD": 10,       # Upper 95% values
        "AE": 10,       # Beta Weight values

        # Prediction Outputs (AG–AI): interval box label/values, inputs, training mean.
        "AG": 24,       # section heading, "PREDICTION INTERVAL" / "PREDICTION INPUTS" labels, spilled constructed names
        "AH": 16,       # prediction interval values + prediction input values
        "AI": 14,       # Training Mean header + values spill

        # Residual Output (AN–BA): row identifiers (AN) + 11 diagnostics (AO–AY)
        # + 2 v3.3 unit-space columns (AZ, BA). AN holds Row_Labels() —
        # country/identifier strings like "United States" (13).
        "AN": 16,       # row identifiers (Row_Labels)
        "AO": 9, "AP": 9, "AQ": 12, "AR": 9,
        "AS": 14, "AT": 17, "AU": 14,
        "AV": 12,       # Cook's Distance (Flagged) — chart data-label helper column
        "AW": 10, "AX": 12, "AY": 9,
        "AZ": 12,       # Predicted Y (Original Units) — v3.3
        "BA": 12,       # Residual (Original Units) — v3.3

        # BB is NOT a content column and NOT a zone gap — it is the post-zone
        # gutter that bounds the row-2 header wrap (last content column = BA)
        # and anchors the diagnostic charts (they start at _C_BB). Sized here
        # so it reads as a deliberate margin rather than a default-width column.
        "BB": 15,
    }.items():
        sheet.range(f"{column_letter}:{column_letter}").column_width = width

    # Gap columns: thin (width 2) and — critically — left OUT of the outline
    # groups below, which is what lets the zones on either side collapse
    # independently.
    for gap_col in _GAP_COLUMNS:
        sheet.range(
            f"{col_letter(gap_col)}:{col_letter(gap_col)}"
        ).column_width = 2

    # Zone column groups — one collapsible outline per zone, separated by the
    # ungrouped gap columns above.
    for first_col, last_col in _COLUMN_GROUPS:
        sheet.api.Columns(
            f"{col_letter(first_col)}:{col_letter(last_col)}"
        ).Group()

    # Spec-block optional sub-group (E:N). Nested inside the A:N zone group
    # so the regular-MLR essentials (A: Variable, B: Role, C: Include, D:
    # Type) stay visible by default while the optional columns
    # (Reference Level, Order, Transform, Sequence, Sequence Period,
    # Period In Use, Levels, Reference In Use, plus the M/N spec feedback
    # and the I Verdict overlay) collapse to a single "+" button. One
    # click expands them when a Categorical predictor, a Transform, or a
    # Sequence axis enters the spec. Order matters: this must run AFTER
    # the zone-level group above so Excel assigns it outline level 2
    # underneath the level-1 A:N parent.
    _set_spec_block_optional_outline_group(sheet)

    # Size the sub-header row to its wrapped contents (two lines for the
    # longer residual headers). Must run after the column widths above,
    # since wrap points — and therefore the fitted height — depend on them.
    sheet.api.Rows(2).AutoFit()

    # Chart label formula cells are plain cell writes (no COM chart API), so
    # they're written unconditionally — safe in headless/CI environments and
    # directly unit-testable via RecordingSheet.
    _write_chart_label_cells(sheet)

    # Charts must be positioned after column widths are set so that
    # the chart anchor column's .left reflects the final column layout.
    # Guarded per
    # the documented convention: ChartObjects().Add(...) requires the Excel
    # COM API, which is unavailable in CI/headless environments.
    try:
        _write_diagnostic_charts(sheet)
    except Exception:
        pass

    # §4b materialization zone: the boxed Model Context block read via
    # Fit_Context, the Sample_Include row mask, and the terminal Constructed
    # Design Matrix, at their final far-right positions with gutters between
    # them. Runs after the column widths and zone groups so
    # the chart-footprint clearance assertion sees the final geometry. Plain
    # cell writes + defined-name registration only (the COM-geometry assertion
    # is guarded), so it is headless-safe.
    _write_materialization_zone(sheet, closures)

    # Freeze top 2 rows. Requires an active window, which Excel may refuse to
    # grant in a headless/non-interactive session, so this is best-effort.
    try:
        sheet.activate()
        sheet.range("A3").select()
        win = sheet.api.Application.ActiveWindow
        win.FreezePanes = False
        win.SplitRow = 2
        win.SplitColumn = 0
        win.FreezePanes = True
    except Exception:
        pass

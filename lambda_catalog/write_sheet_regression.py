"""
write_sheet_regression.py
Writes the spec-driven Regression sheet into any target workbook.

Layout (five horizontal zones, each preceded — after the spec block — by a
single ungrouped GAP column so the zones collapse independently; see the
"Column-layout paradigm" note above the _C_* constants):
  Col A–L        — Model Specification: the declarative spec block shared with
                   write_sheet_model_construction (Variable / Role / Include /
                   Type / Reference Level / Order / Transform / Sequence /
                   Sequence Period / Period In Use / Levels / Reference In Use).
                   Row 2 = Intercept control (label A2, Allow_Intercept toggle
                   C2) and the Sequence status line (E1); I1 carries the
                   "Verdict" header and I2 the combined verdict switch
                   (the row-1/row-2 cells of column I are above the spec
                   table (SpecTable), so the verdict overlays the
                   Sequence_Period column's input cells without
                   disturbing the spec rows); M holds the Δ header
                   and N holds the Count header for the
                   Sequence_Delta_Spectrum() spill at M2:N?
  Col O          — thin gap (width 2, ungrouped)
  Col P–V        — Predictor Summary: level-qualified constructed names (P) +
                   Pearson R, Spearman R, Skewness, Kurtosis, GVIF, Tolerance —
                   computed on the CONSTRUCTED design matrix (dummies included);
                   GVIF/Tolerance share one value across a categorical
                   predictor's dummy columns (Generalized VIF)
  Col W          — thin gap (width 2, ungrouped)
  Col X–AE       — Regression Outputs: Predicted Variable readout (AB2:AC2),
                   Statistics (X–Y rows 3–8), Diagnostics (AA–AB rows 3–12;
                   the serial-correlation pair DW/BFN at rows 11–12),
                   Alpha input (Y12), ANOVA Table (rows 13–17, X–AC),
                   Coefficients (rows 19+, X–AE), Beta Weights (AE)
  Col AF         — thin gap (width 2, ungrouped)
  Col AG–AI      — Prediction Outputs: Prediction Interval (AG1:AH14, boxed —
                   point/CI/PI rows 3-11, FE Group selector + ybar_i/T_i
                   readouts rows 12-14), Prediction Inputs (AG16+, one row
                   per constructed column, no Intercept row), Training Mean
                   spill (AI19 — the single Predictor_Columns() evaluation
                   the orange AH prefills INDEX into; owns column AI downward
                   so it can never collide with another spill)
  Col AJ         — thin gap (width 2, ungrouped)
  Col AK–AW      — Residual Output: heading + Row_Labels() identifiers in AK;
                   11 diagnostics columns (AL–AV, the last being the Cook's
                   Distance flagged-point overlay feeding chart data labels),
                   plus AW gutter; spills downward from row 3

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

import math
from pathlib import Path
from typing import Any

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
    MAX_EXCEL_ROW, a1, add_expression_format, bold, bold_row, border_box,
    col_letter, drop_local_name, excel_color, f, format_input, rc,
    safe_activate, section_heading, val,
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
    _FIXED_EFFECTS_COUNT_FORMULA,
    _FIXED_EFFECTS_NAME_FORMULA,
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
    _C_ORDER as _C_SPEC_ORDER,
    _C_REF_IN_USE as _C_SPEC_REF_IN_USE,
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
    _C_TRANSFORM as _C_SPEC_TRANSFORM,
    _C_TYPE as _C_SPEC_TYPE,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _HEADER_ROW as _SPEC_HEADER_ROW,
    _set_sheet_scoped_names as _set_spec_scoped_names,
    _set_spec_block_column_widths,
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
# share one collapse control. The gap columns are O, W, AF, AJ — one before
# each zone that follows the spec block.
#
#   A–L   Model Specification   | M, N spec feedback + I Verdict overlay
#   | O gap | P–V Predictor Summary
#   | W gap | X–AE Regression Outputs | AF gap | AG–AI Prediction Outputs
#   | AJ gap | AK–AV Residual Output
#
# The gap columns and the per-zone content spans below are the single source
# of truth for both column widths and outline grouping (see _ZONES / _GAP_COLUMNS).

# Zone 1: Model Specification — columns A–L are owned by the shared spec-block
# writers in write_sheet_model_construction (imported above); only the section
# heading cell is written here.
_C_A = 1    # spec: Variable labels / A1 zone heading / A2 Intercept label

# Spec feedback (M, N, plus the I Verdict overlay): the delta spectrum
# (Sequence_Delta_Spectrum() spill at M2:N?) sits in M and N. The combined
# verdict switch overlays column I (the Sequence_Period spec column on
# the spec data rows; the row-1/row-2 cells are above the spec table and
# free). I1 holds the "Verdict" header (bold), I2 the priority-ordered
# switch formula (off-grid outranks regularity, both outrank no-natural
# and calendar; red CF outranks yellow via StopIfTrue). The E1 cell
# carries the multi-flag Sequence error status (moved here from H2 when
# the spec data area became a structured table, SpecTable).
_C_M = 13   # spec feedback: Δ header / spectrum spill
_C_N = 14   # spec feedback: Count header / spectrum spill

# Gap before Predictor Summary.
_C_O = 15   # thin gap (ungrouped — splits the spec and predictor outlines)

# Zone 2: Predictor Summary (constructed columns)
_C_P = 16   # constructed column names (level-qualified)
_C_Q = 17   # Pearson R
_C_R = 18   # Spearman R
_C_S = 19   # Skewness
_C_T = 20   # Kurtosis
_C_U = 21   # GVIF
_C_V = 22   # Tolerance

# Gap before Regression Outputs.
_C_W = 23   # thin gap (ungrouped)

# Zone 3: Regression Outputs
_C_X = 24   # labels (stats / ANOVA / coefficients)
_C_Y = 25   # stat values / ANOVA df / coefficient values
_C_Z = 26   # ANOVA SS / coefficient SE
_C_AA = 27  # diagnostics labels / ANOVA MS / coefficient t-stat
_C_AB = 28  # Predicted Variable label (AB2) / diagnostics values / ANOVA F / coefficient p-value
_C_AC = 29  # predicted variable readout (AC2) / ANOVA Sig F / coefficient CI lower
_C_AD = 30  # coefficient CI upper
_C_AE = 31  # Beta Weights

# Gap before Prediction Outputs.
_C_AF = 32  # thin gap (ungrouped)

# Zone 4: Prediction Outputs
_C_AG = 33  # prediction interval labels / prediction input labels
_C_AH = 34  # prediction interval values / prediction input values
_C_AI = 35  # Training Mean — the per-constructed-column means spill (AI13).
            # The spill owns column AI downward, so it can never collide with
            # another spill when the source data or spec changes.

# Gap before Residual Output.
_C_AJ = 36  # thin gap (ungrouped)

# Zone 5: Residual Output
# _C_AK holds Row_Labels() (the identifiers spill); _C_AL onward hold the
# residual-diagnostic columns. Row-2 headers are written by _write_residuals
# (see the note_cells table); row-3 formulas are the source of truth for what
# each column actually contains.
_C_AK = 37  # row identifiers (Row_Labels() spill)
_C_AL = 38  # Y (actual dependent variable)
_C_AM = 39  # Predicted Y
_C_AN = 40  # Residuals
_C_AO = 41  # Hat Diagonal
_C_AP = 42  # Studentized Residuals
_C_AQ = 43  # Cook's Distance
_C_AR = 44  # Normal Scores Ranked
_C_AS = 45  # Studentized Residuals Ranked
_C_AT = 46  # Scale-Location
_C_AU = 47  # PRESS Residual
_C_AV = 48  # Cook's Distance (Flagged) — chart data-label helper column
_C_AW = 49  # non-content gutter column — wrap-text bound for row-2 header strip / diagnostic-chart anchor

# The constructed-column count is spec-dependent (19 on the default WHO spec),
# so bands that v1 sized with the fixed k=18 now cover a generous fixed range.
_PRED_INPUT_FIRST_ROW = 19
_PRED_INPUT_LAST_ROW = 62
_FORMAT_BAND_LAST_ROW = 62

# Content zones as (first_col, last_col) spans — the single source of truth for
# the outline groups. Each pair becomes one collapsible column group; the gap
# columns between them (below) stay ungrouped so the zones collapse
# independently. Zone 1 includes the M/N spec feedback columns and the
# column-I Verdict overlay so the spec block and its feedback share one
# outline (one click to collapse the spec block together — a spec edit
# doesn't need the feedback visible).
_ZONES: tuple[tuple[int, int], ...] = (
    (_C_A, _C_N),                 # A:N  — Model Specification + feedback
    (_C_P, _C_V),                 # P:V  — Predictor Summary
    (_C_X, _C_AE),                # X:AE — Regression Outputs
    (_C_AG, _C_AI),               # AG:AI — Prediction Outputs
    (_C_AK, _C_AV),               # AK:AV — Residual Output
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

# Chart label formula cells — one row per diagnostic chart, well below the
# 2-col x 4-row chart grid's pixel footprint (row_step=320pt starting at row
# 3, ~85 rows at default row height) so nothing ever renders on top of them.
# Columns sit past _C_AW, which stays a literal non-content gutter.
_C_CHART_LABEL_NAME = 50   # AX — human-readable chart name (doc only)
_C_CHART_TITLE = 51        # AY — Chart Title formula
_C_CHART_XLABEL = 52       # AZ — X-Axis Title formula
_C_CHART_YLABEL = 53       # BA — Y-Axis Title formula
_ROW_CHART_LABELS = 95     # first of 7 rows, one per chart in chart_specs order

# ── §4b materialization zone (v3.0 stage two) ─────────────────────────────────
# Per ARCHITECTURE §4b, the Regression sheet carries a band of materialized
# artifacts at its far right — values computed once into a spill range, read by
# the formulas that would otherwise recompute them. Excel does not memoize a name
# whose RefersTo is a formula, so a constructor called inside ~30 engine
# functions runs ~30 times; a materialized cell runs once.
#
# Zone order (increasing width, terminating in the unbounded zone):
#   charts | gutter | Model_Context | gutter | Sample_Include | gutter | (matrix)
#   (4 x 1)                    (n x 1)                       (n x k, stage three)
#
# Model_Context and Sample_Include are one column each and ship EXPANDED; the
# terminal Constructed Design Matrix (stage three) ships collapsed. Gutters are
# width-2 ungrouped separators so each zone collapses independently — the first
# gutter (after the charts) is structural, keeping the floating chart anchors
# out of every collapsible outline group.
#
# The chart footprint needs an explicit bound. _C_AW is the chart ANCHOR, not
# its extent: the seven diagnostic charts are floating objects tiled in a
# _CHART_GRID_COLS x _CHART_GRID_ROWS grid, whose right edge sits
# _CHART_RIGHT_OFFSET_PT points past AW's left edge. _LAST_CHART_COLUMN is a
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
_LAST_CHART_COLUMN = 65   # conservative clear-past-the-footprint bound; asserted

# Bounded materialization columns + their ungrouped gutters. The terminal
# Constructed Design Matrix (stage three) lands at _C_GUTTER_AFTER_SAMPLE_INCLUDE
# + 1 and runs unbounded to the sheet's right edge; nothing may ever sit right
# of it (ARCHITECTURE §4b ordering rule).
_C_GUTTER_AFTER_CHARTS = _LAST_CHART_COLUMN + 1            # 66 — structural
_C_MODEL_CONTEXT = _C_GUTTER_AFTER_CHARTS + 1               # 67 — 4x1 spill
_C_GUTTER_AFTER_CONTEXT = _C_MODEL_CONTEXT + 1               # 68
_C_SAMPLE_INCLUDE_MATERIALIZED = _C_GUTTER_AFTER_CONTEXT + 1  # 69 — n x 1 spill
_C_GUTTER_AFTER_SAMPLE_INCLUDE = _C_SAMPLE_INCLUDE_MATERIALIZED + 1  # 70

# Model_Context is a bounded, fixed-height cache. ROWS(Fit_Context()) is a
# build-time constant asserted in _write_materialization_zone; the materialized
# VSTACK below has exactly this many elements, so the assertion is structural.
_MODEL_CONTEXT_ROWS = 4
# The context's four elements, in versioned public-contract row order
# (append-only, never insert): Has_Intercept, DF_Absorbed, Response_Transform,
# Predictor_Transform. All four are materialized from the spec block in
# _write_materialization_zone: elements 1-2 (the C2 Allow_Intercept toggle and
# Absorbed_Degrees_Of_Freedom()) feed today's engines; elements 3-4 (the
# response and predictor transform summaries) have no engine reader until the
# v3.3 unit-space dispatcher but land now so the row order is fixed.
_MODEL_CONTEXT_ELEMENT_LABELS = (
    "Has_Intercept",
    "DF_Absorbed",
    "Response_Transform",
    "Predictor_Transform",
)
assert len(_MODEL_CONTEXT_ELEMENT_LABELS) == _MODEL_CONTEXT_ROWS

# ── Visual formatting helpers ─────────────────────────────────────────────────


def _input_range(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    sheet.range(rc(r1, c1), rc(r2, c2)).color = _INPUT


# ── Note (cell comment) sizing ─────────────────────────────────────────────────
# Excel's default new-comment box (≈128x74pt) is too small for most of the
# plain-language notes on this sheet, so every note is explicitly sized and
# positioned instead of left at the Excel default. Width/height are guessed
# from the text length (see _note_dimensions); either axis can be overridden
# per note here for manual tuning without touching the sizing heuristic.
# Key is the note's label — the sheet_notes key for statistical-term notes
# (e.g. "Durbin-Watson"), or the human-readable label passed at the four
# Model Specification call sites below (e.g. "Reserved", "Transform").
_NOTE_MIN_WIDTH = 150.0     # points
_NOTE_MAX_WIDTH = 320.0     # points
_NOTE_BASE_WIDTH = 200.0    # width used for a ~80-char note before scaling
_NOTE_CHARS_PER_LINE_PER_POINT = 1.0 / 5.2  # ~5.2pt per character at 8pt Tahoma
_NOTE_LINE_HEIGHT = 12.0    # points per wrapped line
_NOTE_MIN_HEIGHT = 32.0     # points
_NOTE_VERTICAL_PADDING = 10.0  # points added above/below the wrapped text

_NOTE_SIZE_OVERRIDES: dict[str, tuple[float | None, float | None]] = {
    # "Durbin-Watson": (320.0, 170.0),  # example manual override (width, height)
}


def _note_dimensions(label: str, text: str) -> tuple[float, float]:
    """Guess a (width, height) in points that fits `text` without clipping.

    Width grows with text length (clamped to a readable range); height is
    then derived from how many lines that width wraps the text into. A
    per-`label` entry in _NOTE_SIZE_OVERRIDES replaces either axis (or
    both) for hand-tuning notes the heuristic guesses wrong for.
    """
    length = len(text)
    width = min(
        _NOTE_MAX_WIDTH,
        max(_NOTE_MIN_WIDTH, _NOTE_BASE_WIDTH + (length - 80) * 0.35),
    )
    chars_per_line = max(1, int(width * _NOTE_CHARS_PER_LINE_PER_POINT))
    lines = sum(
        max(1, math.ceil(len(paragraph) / chars_per_line))
        for paragraph in text.split("\n")
    )
    height = max(_NOTE_MIN_HEIGHT, lines * _NOTE_LINE_HEIGHT + _NOTE_VERTICAL_PADDING)

    override_width, override_height = _NOTE_SIZE_OVERRIDES.get(label, (None, None))
    return (
        override_width if override_width is not None else width,
        override_height if override_height is not None else height,
    )


def _set_note(
    sheet: xw.Sheet, row: int, col: int, text: str, *, label: str | None = None
) -> None:
    """Replace the cell's note/comment text with a plain-language explanation.

    The note box is sized to fit `text` (see _note_dimensions) and anchored
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
    comment = cell_api.Comment
    comment.Visible = False
    width, height = _note_dimensions(label if label is not None else text, text)
    try:
        comment_shape = comment.Shape
        comment_shape.Width = width
        comment_shape.Height = height
        comment_shape.Left = cell.left + cell.width
        comment_shape.Top = cell.top
    except Exception:  # pylint: disable=broad-exception-caught
        pass


def _annotate_statistical_terms(sheet: xw.Sheet, sheet_notes: dict[str, str]) -> None:
    """Attach plain-language notes to key statistical labels on the sheet."""
    note_cells = [
        (2, _C_Q, "Pearson R"),
        (2, _C_R, "Spearman R"),
        (2, _C_S, "Skewness"),
        (2, _C_T, "Kurtosis"),
        (2, _C_U, "GVIF"),
        (2, _C_V, "Tolerance"),
        (4, _C_X, "Multiple R"),
        (5, _C_X, "R Square"),
        (6, _C_X, "Adjusted R Square"),
        (7, _C_X, "Standard Error"),
        (8, _C_X, "Observations"),
        (4, _C_AA, "PRESS"),
        (5, _C_AA, "PRESS R²"),
        (6, _C_AA, "Mean Leverage"),
        (7, _C_AA, "AIC"),
        (8, _C_AA, "BIC"),
        (9, _C_AA, "AICc"),
        (10, _C_AA, "QQ Correlation"),
        (11, _C_AA, "Durbin-Watson"),
        (12, _C_AA, "BFN Panel Durbin-Watson"),
        (12, _C_X, "Alpha"),
        (14, _C_Y, "df"),
        (14, _C_Z, "SS"),
        (14, _C_AA, "MS"),
        (14, _C_AB, "F"),
        (14, _C_AC, "Significance F"),
        (15, _C_X, "Regression"),
        (16, _C_X, "Residual"),
        (17, _C_X, "Total"),
        (20, _C_Y, "Coefficients"),
        (20, _C_Z, "Std Error"),
        (20, _C_AA, "t Stat"),
        (20, _C_AB, "P-value"),
        (20, _C_AC, "Lower 95%"),
        (20, _C_AD, "Upper 95%"),
        (20, _C_AE, "Beta Weight"),
        (17, _C_AI, "Training Mean"),
        (3, _C_AG, "Point Estimate"),
        (4, _C_AG, "SE (Mean)"),
        (5, _C_AG, "SE (New Obs)"),
        (6, _C_AG, "t Critical"),
        (7, _C_AG, "CI Lower"),
        (8, _C_AG, "CI Upper"),
        (9, _C_AG, "PI Lower"),
        (10, _C_AG, "PI Upper"),
        (11, _C_AG, "Confidence Level"),
        (12, _C_AG, "FE Group"),
        (13, _C_AG, "Group Mean (y)"),
        (14, _C_AG, "Group Count"),
        (2, _C_AL, "Y"),
        (2, _C_AM, "Predicted Y"),
        (2, _C_AN, "Residuals"),
        (2, _C_AO, "Hat Diagonal"),
        (2, _C_AP, "Studentized Residuals"),
        (2, _C_AQ, "Cook's Distance"),
        (2, _C_AR, "Normal Scores Ranked"),
        (2, _C_AS, "Studentized Residuals Ranked"),
        (2, _C_AT, "Scale-Location"),
        (2, _C_AU, "PRESS Residual"),
        (2, _C_AV, "Cook's Distance (Flagged)"),
    ]

    for row, col, key in note_cells:
        note_text = sheet_notes.get(key)
        if note_text is not None:
            _set_note(sheet, row, col, note_text, label=key)


def _write_significance_conditional_formatting(sheet: xw.Sheet) -> None:
    """Flag nonsignificant coefficient and overall-model P-values."""

    coefficient_p_values = f"AB21:AB{MAX_EXCEL_ROW}"
    significance_f = "AC15"

    sheet.range(coefficient_p_values).api.FormatConditions.Delete()
    sheet.range(significance_f).api.FormatConditions.Delete()

    # Individual coefficient P-values above alpha.
    add_expression_format(
        sheet,
        coefficient_p_values,
        "=AND(ISNUMBER(AB21),AB21>$Y$12)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Overall regression P-value above alpha.
    add_expression_format(
        sheet,
        significance_f,
        "=AND(ISNUMBER(AC15),AC15>$Y$12)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_residual_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply diagnostic cutoffs to the residual-output columns."""

    addresses = {
        "hat":                f"AO3:AO{MAX_EXCEL_ROW}",
        "studentized":        f"AP3:AP{MAX_EXCEL_ROW}",
        "cooks":              f"AQ3:AQ{MAX_EXCEL_ROW}",
        "studentized_ranked": f"AS3:AS{MAX_EXCEL_ROW}",
        "scale_location":     f"AT3:AT{MAX_EXCEL_ROW}",
        "press_residual":     f"AU3:AU{MAX_EXCEL_ROW}",
        "cooks_flag":         f"AV3:AV{MAX_EXCEL_ROW}",
    }

    # Remove existing rules so repeated builds do not duplicate them.
    for address in addresses.values():
        sheet.range(address).api.FormatConditions.Delete()

    # ── Hat diagonal ─────────────────────────────────────────────────────────
    # AB6 contains mean leverage, p/n.
    # > 2p/n: light-red fill and dark-red text.
    add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AO3),AO3>2*$AB$6)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # > 3p/n: additionally bold.
    add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AO3),AO3>3*$AB$6)",
        bold_font=True,
    )

    # ── Studentized residuals ────────────────────────────────────────────────
    for column, address in [
        ("AP", addresses["studentized"]),
        ("AS", addresses["studentized_ranked"]),
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

    # ── Cook's distance (and its AV "Flagged" duplicate) ────────────────────
    # Y8 contains the number of observations, n. AV mirrors AQ (NA() below
    # both cutoffs), so the same two rules applied to AV just recolor
    # whatever value AQ already produced there — kept visually consistent
    # with the source column it duplicates.
    for column, address in [
        ("AQ", addresses["cooks"]),
        ("AV", addresses["cooks_flag"]),
    ]:
        # 4/n < D <= 0.9: light-yellow fill and dark-yellow text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({column}3),{column}3>4/$Y$8,{column}3<=0.9)",
            fill=CF_YELLOW_FILL,
            font_color=CF_DARK_YELLOW_TEXT,
        )

        # D > 0.9: light-red fill and dark-red text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({column}3),{column}3>0.9)",
            fill=CF_LIGHT_RED_FILL,
            font_color=CF_DARK_RED_TEXT,
        )

    # ── Scale-Location: SQRT(|Studentized|) ─────────────────────────────────
    # SQRT(2) ≈ 1.414 corresponds to |Studentized| = 2.
    # SQRT(3) ≈ 1.732 corresponds to |Studentized| = 3.
    add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AT3),AT3>1.414,AT3<=1.732)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AT3),AT3>1.732)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS Residual: e_i / (1 - h_i) ─────────────────────────────────────
    # Y7 contains the Standard Error of the regression.
    # |PRESS| > 2*SE: mild concern; > 3*SE: strong concern.
    add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AU3),ABS(AU3)>2*$Y$7,ABS(AU3)<=3*$Y$7)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AU3),ABS(AU3)>3*$Y$7)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_model_diagnostic_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply rule-of-thumb formatting to GVIF, PRESS R², and QQ Correlation."""

    gvif_address = f"U3:U{MAX_EXCEL_ROW}"
    press_r2_address = "AB5"
    qq_corr_address = "AB10"

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
        "=AND(ISNUMBER(U3),U3>5,U3<=10)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # GVIF > 10: strong multicollinearity warning.
    add_expression_format(
        sheet,
        gvif_address,
        "=AND(ISNUMBER(U3),U3>10)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS R² ─────────────────────────────────────────────────────────────
    # Negative PRESS R² means cross-validated predictions perform worse than
    # predicting the outcome mean.
    add_expression_format(
        sheet,
        press_r2_address,
        "=AND(ISNUMBER(AB5),AB5<0)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── QQ Correlation ────────────────────────────────────────────────────────
    # Pearson r of sorted scaled residuals vs. normal quantiles; near 1.0 = normal errors.
    # < 0.98: mild departure (yellow); < 0.95: stronger departure (red).
    add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(AB10),AB10<0.98,AB10>=0.95)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(AB10),AB10<0.95)",
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

    # alpha: confidence level input, lives in Y12
    drop_local_name(sheet, "alpha")
    sheet.api.Names.Add(
        Name="alpha",
        RefersTo=f"={sname}!$Y$12",
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

    # ── Chart data ranges (OFFSET-based, sized to n = $Y$8 observations) ────────
    # These worksheet-scoped names feed chart SERIES formulas as
    # ='Regression'!<Name>, avoiding full-column references that degrade
    # performance and avoiding the unsupported # spill operator in chart formulas.
    # The third element becomes the Name Manager comment so a user browsing
    # names can tell which chart each range feeds.
    for _name, _col_ltr, _comment in [
        ("RegChartQQX", col_letter(_C_AR),
         "Normal Q-Q chart: X values (theoretical quantiles, Normal Scores Ranked)"),
        ("RegChartQQY", col_letter(_C_AS),
         "Normal Q-Q chart: Y values (Studentized Residuals Ranked)"),
        ("RegChartFitY", col_letter(_C_AM),
         "Predicted Y: X values for the Residuals vs. Fitted, Actual vs. Predicted, and Scale-Location charts"),
        ("RegChartResid", col_letter(_C_AN),
         "Residuals vs. Fitted chart: Y values (Residuals)"),
        ("RegChartActY", col_letter(_C_AL),
         "Actual vs. Predicted chart: Y values (Actual Y)"),
        ("RegChartScaleLoc", col_letter(_C_AT),
         "Scale-Location chart: Y values (sqrt of abs Studentized Residual)"),
        ("RegChartCookDist", col_letter(_C_AQ),
         "Cook's Distance chart: bar values"),
        ("RegChartLeverage", col_letter(_C_AO),
         "Studentized Residuals vs. Leverage chart: X values (Hat Diagonal)"),
        ("RegChartStudResid", col_letter(_C_AP),
         "Studentized Residuals vs. Leverage chart: Y values"),
        ("RegChartPRESSResid", col_letter(_C_AU),
         "PRESS Residuals chart: bar values"),
        ("RegChartCookDistFlag", col_letter(_C_AV),
         "Cook's Distance chart: flagged-point overlay for data labels (D > 4/n or D > 0.9)"),
        ("RegChartObsLabel", col_letter(_C_AK),
         "Cook's Distance chart: observation identifier for flagged-point data labels"),
    ]:
        drop_local_name(sheet, _name)
        _nm = sheet.api.Names.Add(
            Name=_name,
            RefersTo=f"=OFFSET('{sname}'!${_col_ltr}$2,1,0,MAX(IFERROR('{sname}'!$Y$8,1),1),1)",
        )
        _nm.Comment = _comment


# ── Section writers ───────────────────────────────────────────────────────────

def _write_model_specification(sheet: xw.Sheet) -> None:
    """Zone A–L: the shared spec block + row-2 Intercept control.

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
    _set_note(
        sheet, _SPEC_HEADER_ROW, _C_SPEC_SEQUENCE_PERIOD, _SEQUENCE_PERIOD_NOTE,
        label="Sequence Period",
    )
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_PERIOD_IN_USE, _PERIOD_IN_USE_NOTE, label="Period In Use")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_LEVELS, _LEVELS_NOTE, label="Levels")
    _set_note(sheet, _SPEC_HEADER_ROW, _C_SPEC_REF_IN_USE, _REF_IN_USE_NOTE, label="Reference In Use")


def _write_predictor_summary(sheet: xw.Sheet) -> None:
    """Zone P–V: EDA stats for the CONSTRUCTED design-matrix columns."""
    section_heading(sheet, 1, _C_P, "PREDICTOR SUMMARY")

    for col, header in zip(
        [_C_P, _C_Q, _C_R, _C_S, _C_T, _C_U, _C_V],
        ["", "Pearson R", "Spearman R", "Skewness", "Kurtosis", "GVIF", "Tolerance"],
    ):
        val(sheet, 2, col, header)
    bold_row(sheet, 2, _C_P, _C_V)

    # Spill anchors at row 3 — each spills once per constructed column.
    # Names come from the constructor twin, so dummies are level-qualified
    # and the stats run on the actual design matrix. GVIF/Tolerance are
    # generalized (Fox & Monette): every dummy column from the same
    # categorical predictor shares one value instead of a separate,
    # coding-dependent number per level.
    f(sheet, 3, _C_P, "=TRANSPOSE(Constructed_Column_Names())")
    f(sheet, 3, _C_Q, "=Pearson_R(Predictor_Columns(),Response_Column(),Sample_Include())")
    f(sheet, 3, _C_R, "=Spearman_R(Predictor_Columns(),Response_Column(),Sample_Include())")
    f(sheet, 3, _C_S, "=Skewness(Predictor_Columns(),Sample_Include())")
    f(sheet, 3, _C_T, "=Kurtosis(Predictor_Columns(),Sample_Include())")
    f(sheet, 3, _C_U, "=GVIF(Predictor_Columns(),Constructed_Column_Names(),Sample_Include())")
    f(sheet, 3, _C_V, "=Generalized_Tolerance(Predictor_Columns(),Constructed_Column_Names(),Sample_Include())")

    sheet.range(
        (rc(3, _C_Q)), (rc(_FORMAT_BAND_LAST_ROW, _C_V))
    ).number_format = "0.00"


def _write_regression_outputs_header(sheet: xw.Sheet) -> None:
    section_heading(sheet, 1, _C_X, "REGRESSION OUTPUTS")
    section_heading(sheet, 2, _C_AB, "Predicted Variable")
    section_heading(sheet, 2, _C_AC, "")
    # Derived response name — the header of the Role=Response spec row.
    f(sheet, 2, _C_AC, f"={_RESPONSE_NAME_FORMULA}")


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    """Cols X–Y, rows 3–8."""
    section_heading(sheet, 3, _C_X, "REGRESSION STATISTICS")
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
        val(sheet, row, _C_X, label)
        f(sheet, row, _C_Y, formula)
    sheet.range(rc(4, _C_Y), rc(7, _C_Y)).number_format = "0.0000"
    sheet.range(rc(8, _C_Y), rc(8, _C_Y)).number_format = "0"
    border_box(sheet, 3, _C_X, 8, _C_Y)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    """Cols AA–AB, rows 3–12."""
    section_heading(sheet, 3, _C_AA, "DIAGNOSTICS")
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
        val(sheet, row, _C_AA, label)
        f(sheet, row, _C_AB, formula)
    sheet.range(rc(4, _C_AB), rc(10, _C_AB)).number_format = "0.0000"

    # Serial-correlation trigger matrix — two fixed cells (AA11/AB11 plain DW,
    # AA12/AB12 BFN panel DW), each SELF-GUARDING: every state a cell can show is
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
    val(sheet, 11, _C_AA, "Durbin-Watson")
    f(
        sheet,
        11,
        _C_AB,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars>0,"n/a — FE active",'
        "Durbin_Watson_By(Design_Columns(),Design_Response(),Sequence_Column(),"
        "Sample_Include())))))",
    )
    sheet.range(rc(11, _C_AB), rc(11, _C_AB)).number_format = "0.000"

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
    # column; its dormant Cluster branch activates at v3.8+ without touching
    # this cell). Interpretation reads like DW (near 2 ⇒ no first-order
    # autocorrelation), but its critical values depend on N and T — surfacing
    # those bounds is a recorded open item, and the standard DW bounds must
    # not be presented next to it (the cell note carries that caveat).
    val(sheet, 12, _C_AA, "BFN Panel Durbin-Watson")
    f(
        sheet,
        12,
        _C_AB,
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
    sheet.range(rc(12, _C_AB), rc(12, _C_AB)).number_format = "0.000"
    border_box(sheet, 3, _C_AA, 12, _C_AB)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at Y12 — controls prediction interval confidence level."""
    val(sheet, 12, _C_X, "Alpha")
    bold(sheet, 12, _C_X)
    val(sheet, 12, _C_Y, 0.05)
    format_input(sheet, 12, _C_Y)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 13–17, cols X–AC."""
    section_heading(sheet, 13, _C_X, "ANOVA TABLE")

    for col, header in zip(
        [_C_X, _C_Y, _C_Z, _C_AA, _C_AB, _C_AC],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        val(sheet, 14, col, header)
    bold_row(sheet, 14, _C_X, _C_AC)

    # SST = SSR + SSE must hold under FE too, so every row reads the SAME
    # fit-time pair (Design_Columns()/Design_Response()) — mixing a raw Total SS against
    # within Regression/Residual SS would break the ANOVA identity.
    val(sheet, 15, _C_X, "Regression")
    f(sheet, 15, _C_Y, "=Regression_Degrees_Of_Freedom(Design_Columns(),Fit_Context())")
    f(sheet, 15, _C_Z, "=SS_Regression(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 15, _C_AA, "=MS_Regression(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 15, _C_AB, "=F_Statistic(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 15, _C_AC, "=F_Statistic_P_Value(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")

    val(sheet, 16, _C_X, "Residual")
    f(sheet, 16, _C_Y, "=Residual_Degrees_Of_Freedom(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 16, _C_Z, "=SS_Residual(Design_Columns(),Design_Response(),Sample_Include())")
    f(sheet, 16, _C_AA, "=MS_Residual(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")

    val(sheet, 17, _C_X, "Total")
    f(sheet, 17, _C_Y, "=Total_Degrees_Of_Freedom(Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 17, _C_Z, "=SS_Total(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")

    sheet.range(rc(15, _C_Y), rc(17, _C_Y)).number_format = "0"
    sheet.range(rc(15, _C_Z), rc(17, _C_Z)).number_format = "0.0"
    sheet.range(rc(15, _C_AA), rc(16, _C_AA)).number_format = "0.0"
    sheet.range(rc(15, _C_AB), rc(15, _C_AB)).number_format = "0.0"
    sheet.range(rc(15, _C_AC), rc(15, _C_AC)).number_format = "0.0E+00"
    border_box(sheet, 13, _C_X, 17, _C_AC)


def _write_coefficients(sheet: xw.Sheet) -> None:
    """Cols X–AE, rows 19+. Spills downward — nothing placed below row 39 in these cols."""
    section_heading(sheet, 19, _C_X, "COEFFICIENTS")

    for col, header in zip(
        [_C_X, _C_Y, _C_Z, _C_AA, _C_AB, _C_AC, _C_AD, _C_AE],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%", "Beta Weight"],
    ):
        val(sheet, 20, col, header)
    bold_row(sheet, 20, _C_X, _C_AE)

    # Spill row labels aligned to the constructed columns (level-qualified via
    # the constructor twin). Zero_Predictors_Selected() branch computes a real
    # intercept-only model (label "Intercept") instead of fabricating a result
    # for an unselected variable; NA() when nothing is fit.
    f(
        sheet,
        21,
        _C_X,
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
    f(sheet, 21, _C_Y,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),Intercept_Only_Point(),NA()),'
       'IF(Allow_Intercept,Coefficients(Design_Columns(),Design_Response(),Sample_Include()),'
       'VSTACK("",Coefficients(Design_Columns(),Design_Response(),Sample_Include()))))')
    f(sheet, 21, _C_Z,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,SE_Coefficients(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()),'
       'VSTACK("",SE_Coefficients(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()))))')
    f(sheet, 21, _C_AA,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_Point()/Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,T_Statistics(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()),'
       'VSTACK("",T_Statistics(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context()))))')
    f(sheet, 21, _C_AB,
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
    f(sheet, 21, _C_AC,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()-T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Lower(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()),'
       'VSTACK("",Confidence_Interval_Lower(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()))))')
    f(sheet, 21, _C_AD,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()+T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Upper(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()),'
       'VSTACK("",Confidence_Interval_Upper(Design_Columns(),Design_Response(),Sample_Include(),0.05,Fit_Context()))))')
    # Beta Weights: k×1 (no intercept row); always prepend blank to align with other columns.
    # No predictor exists to standardize in the zero-predictor branch, so render
    # blank (not an error) when Allow_Intercept is TRUE; NA() when nothing is fit.
    f(sheet, 21, _C_AE,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,"",NA()),'
       'VSTACK("",Beta_Weights(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())))')

    for col in [_C_Y, _C_Z, _C_AA, _C_AC, _C_AD, _C_AE]:
        sheet.range(
            rc(21, col), rc(_FORMAT_BAND_LAST_ROW, col)
        ).number_format = "0.0000"
    sheet.range(
        rc(21, _C_AB), rc(_FORMAT_BAND_LAST_ROW, _C_AB)
    ).number_format = "0.0E+00"


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    """Zone AG1:AH14: boxed prediction interval output, plus the FE group
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
    section_heading(sheet, 1, _C_AG, "PREDICTION OUTPUTS")
    val(sheet, 2, _C_AG, "PREDICTION INTERVAL")
    bold(sheet, 2, _C_AG)
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
        val(sheet, row, _C_AG, label)
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
    # logged), TRANSPOSE'd to match the AH band's column-vector shape, and
    # Ln_Positive is applied elementwise only where it reads "Log". The
    # user always types a raw value in AH (e.g. actual miles), never ln(x).
    f(
        sheet,
        3,
        _C_AH,
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
        f"LET(raw,TAKE($AH${_PRED_INPUT_FIRST_ROW}:$AH${_PRED_INPUT_LAST_ROW},COLUMNS(Predictor_Columns())),"
        "trn,TRANSPOSE(Constructed_Column_Transforms()),"
        'pred_input,IF(trn="Log",Ln_Positive(raw),raw),'
        "Group_Prediction_Interval(Predictor_Columns(),Response_Column(),pred_input,"
        "Prediction_Group_Column(),$AH$12,"
        "Sample_Include(),alpha,Fit_Context())))",
    )
    sheet.range(rc(3, _C_AH), rc(11, _C_AH)).number_format = "0.0000"

    # FE Group selector (row 12): computed-with-override, the same
    # reference-level pattern as the Categorical Reference Level (E) —
    # pre-filled with the alphabetically-first observed group (which is
    # always the "(all)" sentinel itself when no Fixed Effects row is
    # declared, since Prediction_Group_Column() is constant in that state),
    # and directly editable to any other observed group. Red CF flags a
    # typed value that is not among the observed groups.
    val(sheet, 12, _C_AG, "FE Group")
    f(
        sheet,
        12,
        _C_AH,
        "=INDEX(SORT(UNIQUE(FILTER(Prediction_Group_Column(),Sample_Include()))),1,1)",
    )
    format_input(sheet, 12, _C_AH)
    fe_group_cell = f"${col_letter(_C_AH)}$12"
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
    val(sheet, 13, _C_AG, "Group Mean (y)")
    f(
        sheet,
        13,
        _C_AH,
        "=Group_Mean_At(Response_Column(),Prediction_Group_Column(),$AH$12,Sample_Include())",
    )
    val(sheet, 14, _C_AG, "Group Count")
    f(
        sheet,
        14,
        _C_AH,
        "=Group_Count_At(Prediction_Group_Column(),$AH$12,Sample_Include())",
    )
    sheet.range(rc(13, _C_AH), rc(13, _C_AH)).number_format = "0.0000"
    sheet.range(rc(14, _C_AH), rc(14, _C_AH)).number_format = "0"

    border_box(sheet, 1, _C_AG, 14, _C_AH)


def _write_prediction_inputs(sheet: xw.Sheet) -> None:
    """Zone AG16+: one prediction-input row per constructed column.

    No Intercept row here (unlike the pre-v2.1 layout): Group_Prediction_Interval's
    pred_input is exactly COLUMNS(Predictor_Columns()) raw predictor values with no
    intercept slot — group-mean recovery never uses one (the selected
    group's own mean plays that role) — so the row that used to hold it
    would be actively misleading now. Row 18 is a blank spacer between the
    headers and the first predictor row.
    """
    section_heading(sheet, 16, _C_AG, "PREDICTION INPUTS")
    val(sheet, 17, _C_AG, "Predictor")
    val(sheet, 17, _C_AH, "Prediction Value")
    bold_row(sheet, 17, _C_AG, _C_AI)

    # AG19: spill formula — level-qualified names, one per constructed column
    f(sheet, _PRED_INPUT_FIRST_ROW, _C_AG, "=TRANSPOSE(Constructed_Column_Names())")

    # AI19: the Training Mean column — per-column means of the filtered design
    # matrix, computed with a SINGLE Predictor_Columns() evaluation. Predictor_Columns() is
    # a full
    # design-matrix construction on every call (Excel does not cache LAMBDA
    # results), so this spill is the one place the means are computed; the
    # orange prefill cells INDEX into it. The earlier design — every prefill
    # cell calling Predictor_Columns() twice (width guard + column mean) — made the
    # workbook's first full calculation pathological (~20 minutes at the save
    # step). The spill owns column AI downward (the AJ gap and residual zone
    # start beyond it), so a wider dataset or spec can never make it collide
    # with another spill.
    # Degrades to "" on an empty model, which the prefill guard reads as a
    # one-row spill holding a blank.
    #
    # Log columns need the GEOMETRIC mean here, not the arithmetic mean of
    # the already-logged Predictor_Columns() column: the AH prefill cells below just
    # INDEX into this spill, and the row-3 prediction formula applies
    # Ln_Positive to whatever it finds in AH — so if this spill held the
    # log-space arithmetic mean, the default prediction would silently
    # double-log (ln(ln(x)), not merely un-back-transformed). EXP(mean(ln
    # x)) is exact and self-cancelling: Ln_Positive(EXP(mean(ln x))) =
    # mean(ln x), so the default prediction still lands precisely on
    # Predictor_Columns()'s own centroid, unchanged from the pre-Log-transform behavior.
    # Constructed_Column_Transforms() gives the per-column Log/None flag in
    # the same 1xk shape as the BYCOL means row, so the two combine
    # elementwise before the single TRANSPOSE down into the AI column.
    val(sheet, 17, _C_AI, "Training Mean")
    means_anchor = f"$AI${_PRED_INPUT_FIRST_ROW}"
    f(
        sheet,
        _PRED_INPUT_FIRST_ROW,
        _C_AI,
        (
            "=IFERROR(TRANSPOSE(LET(m,BYCOL(FILTER(Predictor_Columns(),Sample_Include()),"
            "LAMBDA(c,AVERAGE(c))),t,Constructed_Column_Transforms(),"
            'IF(t="Log",EXP(m),m))),"")'
        ),
    )

    # AH19:AH62 — the Training Mean of each constructed column, individually
    # overridable. Each row guards on its position against the means-spill
    # height so rows beyond the live constructed width render blank (the
    # width is spec-dependent — 19 on the default WHO spec). Cheap spill
    # references only: no prefill cell may call Predictor_Columns() itself.
    offset = _PRED_INPUT_FIRST_ROW - 1
    for row in range(_PRED_INPUT_FIRST_ROW, _PRED_INPUT_LAST_ROW + 1):
        f(
            sheet,
            row,
            _C_AH,
            (
                f"=IF(ROW()-{offset}<=IFERROR(ROWS({means_anchor}#),0),"
                f"INDEX({means_anchor}#,ROW()-{offset}),"
                '"")'
            ),
        )

    # Orange for all user-editable prediction value cells (the Training Mean
    # column is computed display, not input)
    _input_range(sheet, _PRED_INPUT_FIRST_ROW, _C_AH, _PRED_INPUT_LAST_ROW, _C_AH)
    sheet.range(
        rc(_PRED_INPUT_FIRST_ROW, _C_AH), rc(_PRED_INPUT_LAST_ROW, _C_AH)
    ).number_format = "0.0000"
    # Column-wide: the means spill height is spec-dependent and the column
    # holds nothing else.
    sheet.range(f"{col_letter(_C_AI)}:{col_letter(_C_AI)}").number_format = "0.0000"


def _write_residuals(sheet: xw.Sheet) -> None:
    """Residual diagnostic table — row identifiers + 11 diagnostics columns starting at AL."""
    section_heading(sheet, 1, _C_AK, "RESIDUAL OUTPUT")

    # AK2: static header — Row_Labels() supplies its own per-row content
    # (joined Identifier columns, or positional Obs. n labels).
    val(sheet, 2, _C_AK, "Observation")

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
    # fills the spec feedback block's "FE Variable" cell. Y (_C_AL) gets its
    # own wording instead of "Within": Design_Columns()/Design_Response() only SUBTRACT the
    # group mean, they never divide by a standard deviation, so "Y (Within
    # Country)" would read as a demeaning but "St Devs from Avg." would be
    # outright wrong — "Deviation from <FE> Avg." says exactly what the
    # column holds (still response-scale units, just group-centered).
    response_scale_headers = {_C_AL, _C_AM, _C_AN, _C_AU}
    fe_within_suffix = f'" (Within "&{_FIXED_EFFECTS_NAME_FORMULA}&")"'
    fe_within_log_suffix = f'" (Within "&{_FIXED_EFFECTS_NAME_FORMULA}&", Log)"'
    fe_deviation_suffix = f'" (Deviation from "&{_FIXED_EFFECTS_NAME_FORMULA}&" Avg.)"'
    fe_deviation_log_suffix = (
        f'" (Deviation from "&{_FIXED_EFFECTS_NAME_FORMULA}&" Avg., Log)"'
    )
    for col, header in zip(
        [_C_AL, _C_AM, _C_AN, _C_AO, _C_AP, _C_AQ, _C_AR, _C_AS, _C_AT, _C_AU, _C_AV],
        [
            "Y", "Predicted Y", "Residuals",
            "Hat Diagonal", "Studentized Residuals", "Cook's Distance",
            "Normal Scores Ranked", "Studentized Residuals Ranked",
            "Scale-Location", "PRESS Residual", "Cook's Distance (Flagged)",
        ],
    ):
        if col == _C_AL:
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
    bold_row(sheet, 2, _C_AK, _C_AV)

    # AK3: row labels — the spec-derived Row_Labels() filtered to the sample.
    # Row_Labels() has its own no-Identifier fallback ("Obs. n"), so the only
    # error left to absorb is an all-FALSE mask.
    f(
        sheet, 3, _C_AK,
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
    f(sheet, 3, _C_AL, "=Dependent_Variable(Design_Response(),Sample_Include())")
    f(sheet, 3, _C_AM, "=Predictions(Design_Columns(),Design_Response(),Sample_Include())")
    f(sheet, 3, _C_AN, "=Residuals(Design_Columns(),Design_Response(),Sample_Include())")
    f(sheet, 3, _C_AO, "=Hat_Diagonal(Design_Columns(),Sample_Include())")
    f(sheet, 3, _C_AP, "=Studentized_Residuals(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 3, _C_AQ, "=Cooks_Distance(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    f(sheet, 3, _C_AR, "=SORT(Normal_Scores(Design_Response(),Sample_Include()))")
    f(sheet, 3, _C_AS, "=Studentized_Residuals_Ranked(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())")
    # Scale-Location: SQRT(|Studentized_Residuals|) — horizontal spread should be flat.
    f(
        sheet, 3, _C_AT,
        "=SQRT(ABS(Studentized_Residuals(Design_Columns(),Design_Response(),Sample_Include(),Fit_Context())))",
    )
    # PRESS Residual equals the leave-one-out residual e_i / (1 - h_i).
    f(sheet, 3, _C_AU, "=LOOCV_Residual(Design_Columns(),Design_Response(),Sample_Include())")
    # Cook's Distance (Flagged): NA()'d except where D exceeds the standard
    # influence cutoffs (4/n or 0.9 — algebraically MIN(4/n, 0.9) since these
    # are two lower bounds on the same "flag it" union, for every n). This
    # feeds the Cook's Distance chart's data-label overlay series (see
    # RegChartCookDistFlag in _setup_local_names / _write_diagnostic_charts):
    # NA() points are skipped for both plotting and labeling, so only the
    # flagged points get a label.
    cooks_col = col_letter(_C_AQ)
    f(
        sheet, 3, _C_AV,
        f"=IF({cooks_col}3#>MIN(4/$Y$8,0.9),{cooks_col}3#,NA())",
    )
    # Format every numeric residual-output column — the actual Y (AL) through
    # Cook's Distance (Flagged) (AV). Only the AK identifier column (text:
    # country/Obs. labels) is left unformatted.
    sheet.range(f"{col_letter(_C_AL)}:{col_letter(_C_AV)}").number_format = "0.0000"


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
            '="Residuals vs. Fitted — "&$AC$2',
            '="Fitted Values"', '="Residuals"',
            1, 1,
        ),
        (
            "Normal Q-Q", "scatter",
            _name_ref("RegChartQQX"),
            _name_ref("RegChartQQY"),
            '="Normal Q-Q  (r = "&TEXT($AB$10,"0.000")&")"'
            '&IF($AB$10<0.95,"  — check normality","")',
            '="Theoretical Quantiles"', '="Studentized Residuals"',
            1, 2,
        ),
        (
            "Actual vs. Predicted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartActY"),
            '="Actual vs. Predicted — "&$AC$2&"  (Adj. R² = "&TEXT($Y$6,"0.000")&")"',
            '="Predicted "&$AC$2', '="Actual "&$AC$2',
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
            '="Cook\'s Distance  (flag: D > "&TEXT(MIN(4/$Y$8,0.9),"0.000")&")"',
            '="Observation"', '="Cook\'s Distance"',
            3, 1,
        ),
        (
            "Studentized Residuals vs. Leverage", "scatter",
            _name_ref("RegChartLeverage"),
            _name_ref("RegChartStudResid"),
            '="Studentized Residuals vs. Leverage  (mean leverage = "'
            '&TEXT($AB$6,"0.000")&")"',
            '="Leverage (Hat Diagonal)"', '="Studentized Residuals"',
            3, 2,
        ),
        (
            "PRESS Residuals", "bar",
            None,
            _name_ref("RegChartPRESSResid"),
            '="PRESS Residuals  (PRESS = "&TEXT($AB$4,"#,##0.0000")&")"',
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
    FIXED range holding the materialized 4x1 spill. It is what the ~30
    Regression sheet call sites pass, so they read the actual spec-derived
    context (the C2 Allow_Intercept toggle, the Absorbed_Degrees_Of_Freedom()
    closure, the spec-block transform summaries) rather than the constructor
    default. Splitting the names keeps ``Model_Context`` unshadowed: a single
    sheet-scoped thunk named ``Model_Context`` would make ``Model_Context()``
    in a sheet cell resolve to the materialized values while the same token in
    a carrier's omitted-default resolved to the workbook constructor — the
    invisible shadowing the v3.0 release exists to remove.

    The context is materialized ONCE into a spill cell (Excel does not memoize
    a name whose RefersTo is a formula, so a constructor inside ~30 engine
    calls runs ~30 times); ``Fit_Context`` reads that fixed range, so the ~30
    call sites that pass ``Fit_Context()`` all read the one materialized cell.

    Elements 3-4 (Response_Transform, Predictor_Transform) are populated from
    the spec block now but have no engine reader until the v3.3 unit-space
    dispatcher; the row order is the contract that is expensive to change
    later, so all four rows land together. An error in an unconsumed row is
    contained: ``VSTACK`` returns a 4-row array with the error in one cell, and
    the engines read only elements 1-2 through the accessors, so a bad name
    surfaces as a visible cell error (caught by the headless verifier) without
    shifting a single fitted number.

    The ``Sample_Include`` column is placed at its final §4b position as a
    RESERVED placeholder. Promoting the live ``Sample_Include()`` closure to a
    thunk over a materialized spill is deferred: it needs the dynamic-array
    spill operator (``#``) inside a ``LAMBDA`` defined-name RefersTo, a
    combination not yet used anywhere in this workbook and only verifiable with
    Excel present. A wrong guess would break the row-mask contract that keeps
    every spilled array row-aligned, so the optimization is landed separately
    where it can be Excel-verified, not blind. The live closure is untouched.

    Plain cell writes + defined-name registration only; no chart/COM API, so
    this is exercised in unit tests via ``RecordingSheet`` without Excel. The
    chart-footprint clearance check is the one Excel-only step and is guarded.
    """
    sname = sheet.name
    if closures is None:
        closures = load_catalog_document(_DEFINITIONS_PATH).functions_for_sheet(
            REGRESSION_SHEET_NAME
        )

    # ── Model_Context (4x1) ──────────────────────────────────────────────────
    # Element 1 is the C2 Allow_Intercept toggle (named range). Element 2 is
    # the Absorbed_Degrees_Of_Freedom() Regression closure. Elements 3-4 are
    # the spec-block transform summaries: row 3 the single response transform
    # (mirrors _RESPONSE_LOG_FORMULA's INDEX/XMATCH over the spec rows); row 4
    # the None/Log/Mixed summary over the INCLUDED CONTINUOUS predictors — masked
    # to Continuous so a Categorical dummy's transform never yields a false
    # "Mixed". Both reference worksheet-scoped spec names on this sheet. The
    # VSTACK length IS the build-time constant; the assert below pins it to
    # _MODEL_CONTEXT_ROWS so a future edit that changes the context height
    # fails the build loudly rather than shifting every engine call.
    response_transform = (
        'IFERROR(INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data)),'
        f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data)))),"None")'
    )
    predictor_transform = (
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
    context_elements = (
        "Allow_Intercept",
        "Absorbed_Degrees_Of_Freedom()",
        response_transform,
        predictor_transform,
    )
    assert len(context_elements) == _MODEL_CONTEXT_ROWS
    spill_formula = "=VSTACK(" + ",".join(context_elements) + ")"

    ctx_col = col_letter(_C_MODEL_CONTEXT)
    section_heading(sheet, 1, _C_MODEL_CONTEXT, "Model Context")
    f(sheet, 2, _C_MODEL_CONTEXT, spill_formula)
    # The materialized spill occupies rows 2 .. 1 + _MODEL_CONTEXT_ROWS; the
    # sheet-scoped reader Fit_Context reads that fixed range (no spill operator
    # — the height is a structural constant, so a fixed range is exact and
    # avoids the dynamic-array-in-a-name question entirely). Drop both the
    # legacy "Model_Context" sheet name (left by a pre-split build) and any
    # stale "Fit_Context" before re-adding, so a rebuild never leaves a shadow.
    ctx_ref = f"'{sname}'!${ctx_col}$2:${ctx_col}${1 + _MODEL_CONTEXT_ROWS}"
    drop_local_name(sheet, "Model_Context")
    drop_local_name(sheet, "Fit_Context")
    sheet.api.Names.Add(
        Name="Fit_Context",
        RefersTo=f"=LAMBDA({ctx_ref})",
    )
    # Runtime build assertion: ROWS(Fit_Context()) is the build-time constant.
    # Displays TRUE when the materialized context has exactly _MODEL_CONTEXT_ROWS
    # rows; a verifier (or the user) reads this cell to confirm the invariant.
    f(
        sheet,
        2 + _MODEL_CONTEXT_ROWS,
        _C_MODEL_CONTEXT,
        f"=ROWS(Fit_Context())={_MODEL_CONTEXT_ROWS}",
    )

    # ── Sample_Include (reserved) ────────────────────────────────────────────
    section_heading(sheet, 1, _C_SAMPLE_INCLUDE_MATERIALIZED, "Sample Include")
    val(
        sheet,
        2,
        _C_SAMPLE_INCLUDE_MATERIALIZED,
        "reserved",
    )
    # Document the deferral on the header cell so the reserved column is not
    # mistaken for an oversight.
    try:
        sheet.range(rc(1, _C_SAMPLE_INCLUDE_MATERIALIZED)).api.AddComment(
            "Reserved §4b position. Sample_Include() is materialized here as a "
            "thunk over a spill in a follow-up (Excel-verified), not blind; the "
            "live closure remains the row mask until then."
        )
    except Exception:  # pylint: disable=broad-except
        pass

    # ── Column widths + outline groups ───────────────────────────────────────
    # Each bounded zone is one column and ships EXPANDED (per §4b); the
    # width-2 gutters stay ungrouped so the zones collapse independently, and
    # the first gutter (after the charts) is structural — it keeps the
    # floating chart anchors out of every collapsible outline group.
    for gutter in (
        _C_GUTTER_AFTER_CHARTS,
        _C_GUTTER_AFTER_CONTEXT,
        _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    ):
        sheet.range(f"{col_letter(gutter)}:{col_letter(gutter)}").column_width = 2
    for content in (_C_MODEL_CONTEXT, _C_SAMPLE_INCLUDE_MATERIALIZED):
        sheet.range(
            f"{col_letter(content)}:{col_letter(content)}"
        ).column_width = 14
        sheet.api.Columns(f"{col_letter(content)}:{col_letter(content)}").Group()

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
            sheet.range(a1(1, _C_AW)).left + _CHART_RIGHT_OFFSET_PT
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
    start_left = sheet.range(a1(1, _C_AW)).left
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

    def _set_equal_axis_scale_from_named_ranges(
        x_axis: Any,
        y_axis: Any,
        x_name: str,
        y_name: str,
    ) -> None:
        """Set equal min/max scales on both axes using two sheet-scoped named ranges."""
        common_min = float(sheet.api.Evaluate(f"=MIN('{sname}'!{x_name},'{sname}'!{y_name})"))
        common_max = float(sheet.api.Evaluate(f"=MAX('{sname}'!{x_name},'{sname}'!{y_name})"))

        if common_max <= common_min:
            return

        x_axis.MinimumScale = common_min
        x_axis.MaximumScale = common_max
        y_axis.MinimumScale = common_min
        y_axis.MaximumScale = common_max

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
        y_axis.TickLabels.NumberFormat = "0"

        gridline_mode = gridline_modes.get(key, "none")
        x_axis.HasMajorGridlines = gridline_mode == "both"
        y_axis.HasMajorGridlines = gridline_mode in {"y", "both"}

        if key == "Cook's Distance":
            y_axis.TickLabels.NumberFormat = "0.0E+00"
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
            dls.NumberFormat = "0.0E+00"  # matches the y-axis format
            dls.Position = 0              # xlLabelPositionAbove
        if key == "Studentized Residuals vs. Leverage":
            x_axis.TickLabels.NumberFormat = "0.00"
        if key == "PRESS Residuals":
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone

        if key == "Normal Q-Q":
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "RegChartQQX", "RegChartQQY")
            if x_addr is None:
                raise AssertionError("Normal Q-Q chart requires an x-axis range")
            identity_ref: str = x_addr
            _add_identity_line(chart, identity_ref)
        if key == "Actual vs. Predicted":
            # Axis limits are left at Excel's defaults (not forced equal via
            # _set_equal_axis_scale_from_named_ranges) — the identity line
            # below still reads correctly regardless of scale.
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
    _write_prediction_interval(sheet)
    _write_prediction_inputs(sheet)
    _write_residuals(sheet)
    _annotate_statistical_terms(sheet, sheet_notes or {})
    _write_residual_conditional_formatting(sheet)

    sheet.range(rc(2, _C_P), rc(2, _C_AW)).api.WrapText = True

    # A–L (spec block) widths are owned by write_sheet_model_construction.py
    # so the standalone and shared-block builds can never drift.
    _set_spec_block_column_widths(sheet)

    # Content-column widths, per zone, plus the AW post-zone gutter (last entry).
    # The gap columns (O, W, AF, AJ) are sized from _GAP_COLUMNS below so the
    # layout stays declarative — one width there, not one per hard-coded gap letter.
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

        # Residual Output (AK–AV): row identifiers (AK) + 11 diagnostics (AL–AV).
        # AK holds Row_Labels() — country/identifier strings like "United States" (13).
        "AK": 16,       # row identifiers (Row_Labels)
        "AL": 10, "AM": 10, "AN": 9, "AO": 9, "AP": 9, "AQ": 12, "AR": 9,
        "AS": 14, "AT": 17, "AU": 14,
        "AV": 12,       # Cook's Distance (Flagged) — chart data-label helper column

        # AW is NOT a content column and NOT a zone gap — it is the post-zone
        # gutter that bounds the row-2 header wrap (_C_AW) and anchors the
        # diagnostic charts (they start at _C_AW). Sized here so it reads as
        # a deliberate margin rather than a default-width column.
        "AW": 15,
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

    # Size the sub-header row to its wrapped contents (two lines for the
    # longer residual headers). Must run after the column widths above,
    # since wrap points — and therefore the fitted height — depend on them.
    sheet.api.Rows(2).AutoFit()

    # Chart label formula cells are plain cell writes (no COM chart API), so
    # they're written unconditionally — safe in headless/CI environments and
    # directly unit-testable via RecordingSheet.
    _write_chart_label_cells(sheet)

    # Charts must be positioned after column widths are set so that
    # sheet.range("AW1").left reflects the final column layout. Guarded per
    # the documented convention: ChartObjects().Add(...) requires the Excel
    # COM API, which is unavailable in CI/headless environments.
    try:
        _write_diagnostic_charts(sheet)
    except Exception:
        pass

    # §4b materialization zone: Model_Context (4x1, materialized spill) read via
    # and the reserved Sample_Include column, placed at their final far-right
    # positions with gutters. Runs after the column widths and zone groups so
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

"""
write_sheet_regression.py
Writes the spec-driven Regression sheet into any target workbook.

Layout (five horizontal zones, each preceded — after the spec block — by a
single ungrouped GAP column so the zones collapse independently; see the
"Column-layout paradigm" note above the _C_* constants):
  Col A–O        — Model Specification: the declarative spec block shared with
                   write_spec_block (Variable / Role / Include /
                   Type / Reference Level / Order / Transform / Sequence /
                   Sequence Period / Period In Use / Levels / Reference In Use /
                   Interaction Term / Interaction Operation / Design Columns).
                   Row 2 = Intercept control (label A2, Allow_Intercept toggle
                   C2) and the Sequence status line (E1); I1 carries the
                   "Verdict" header and I2 the combined verdict switch
                   (the row-1/row-2 cells of column I are above the spec
                   block's own rows, so the verdict overlays the
                   Sequence_Period column's input cells without
                   disturbing the spec rows); O2 carries the design-matrix
                   width-guard status and N1/O1 its Σ label and total;
                   P holds the Δ header and Q holds the Count header on
                   the spec block's own header row (3), aligned with it,
                   for the Sequence_Delta_Spectrum() spill from row 4
  Col R          — thin gap (width 2, ungrouped)
  Col S–Y        — Predictor Summary: level-qualified constructed names (S) +
                   Pearson R, Spearman R, Skewness, Kurtosis, GVIF, Tolerance —
                   computed on the CONSTRUCTED design matrix (dummies included);
                   GVIF/Tolerance share one value across a categorical
                   predictor's dummy columns (Generalized VIF)
  Col Z          — thin gap (width 2, ungrouped)
  Col AA–AH      — Regression Outputs: Predicted Variable readout (AE3:AF3),
                   Statistics (AA–AB rows 4–9), Diagnostics (AD–AE rows 4–13;
                   the serial-correlation pair DW/BFN at rows 12–13),
                   Alpha input (AB13), ANOVA Table (rows 14–18, AA–AF),
                   Coefficients (rows 20+, AA–AH), Beta Weights (AH)
  Col AI         — thin gap (width 2, ungrouped)
  Col AJ–AL      — Prediction Outputs: Prediction Interval (AJ1:AK15, boxed —
                   point/CI/PI rows 4-12, FE Group selector + ybar_i/T_i
                   readouts rows 13-15), Prediction Inputs (AJ17+, one row
                   per constructed column, no Intercept row), Training Mean
                   spill (AL20 — the single Predictor_Columns() evaluation
                   the orange AK prefills INDEX into; owns column AL downward
                   so it can never collide with another spill)
  Col AM         — thin gap (width 2, ungrouped)
  Col AN–BA      — Residual Output: heading + Row_Labels() identifiers in AN;
                   11 diagnostics columns (AO–AY, the last being the Cook's
                   Distance flagged-point overlay feeding chart data labels),
                   plus the v3.3 original-units pair AZ/BA; spills
                   downward from row 4
  Col BC–BF      — chart title / axis-label formula cells (below the chart grid)
  Col BQ →       — the ARCHITECTURE §4b materialization band: the Model
                   Context block (a boxed, fixed-size label/value pair, one
                   labelled cell per context element), the materialized
                   Sample_Include row mask, and the terminal Constructed
                   Design Matrix — the last two headed on row 3 and spilling
                   from row 4 — separated by ungrouped gutters. Only the
                   Model Context block is grouped: the two data-dependent
                   zones are left ungrouped and visible, because collapsing
                   a spilled zone hides the cells the engines recalculate
                   through. Nothing may ever be placed right of the
                   design matrix; the Model Formula readout sits on ROW 1 of
                   its zone, above a body that only ever grows rightward.

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
write_spec_block so the two sheets can never drift.
"""
# pylint: disable=too-many-lines
from __future__ import annotations

import xlwings as xw

from .catalog_schema import CatalogFunction, load_catalog_document
from .sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_DARK_YELLOW_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
)
from .sheet_styles import (
    INPUT_COLOR as _INPUT,
)
from .workbook_helpers import (
    MAX_EXCEL_ROW,
    add_expression_format,
    anchor_comment_right_of_cell,
    bold,
    bold_row,
    border_box,
    col_letter,
    drop_local_name,
    f,
    format_input,
    note_dimensions,
    quoted_sheet_name,
    rc,
    reset_column_groups,
    safe_activate,
    section_heading,
    val,
)
from .write_spec_block import (
    _C_DESIGN_COLUMNS as _C_SPEC_DESIGN_COLUMNS,
)
from .write_spec_block import (
    _C_INCLUDE as _C_SPEC_INCLUDE,
)
from .write_spec_block import (
    _C_INTERACTION_OPERATION as _C_SPEC_INTERACTION_OPERATION,
)
from .write_spec_block import (
    _C_INTERACTION_TERM as _C_SPEC_INTERACTION_TERM,
)
from .write_spec_block import (
    _C_LABEL as _C_SPEC_LABEL,
)
from .write_spec_block import (
    _C_LEVELS as _C_SPEC_LEVELS,
)
from .write_spec_block import (
    _C_ORDER as _C_SPEC_ORDER,
)
from .write_spec_block import (
    _C_PERIOD_IN_USE as _C_SPEC_PERIOD_IN_USE,
)
from .write_spec_block import (
    _C_REF_IN_USE as _C_SPEC_REF_IN_USE,
)
from .write_spec_block import (
    _C_REFERENCE as _C_SPEC_REFERENCE,
)
from .write_spec_block import (
    _C_ROLE as _C_SPEC_ROLE,
)
from .write_spec_block import (
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
)
from .write_spec_block import (
    _C_SEQUENCE_PERIOD as _C_SPEC_SEQUENCE_PERIOD,
)
from .write_spec_block import (
    _C_TRANSFORM as _C_SPEC_TRANSFORM,
)
from .write_spec_block import (
    _C_TYPE as _C_SPEC_TYPE,
)
from .write_spec_block import (
    _DEFAULT_TRANSFORM,
    _DESIGN_COLUMNS_NOTE,
    _FIXED_EFFECTS_COUNT_FORMULA,
    _FIXED_EFFECTS_NAME_FORMULA,
    _INCLUDE_NOTE,
    _INTERACTION_OPERATION_NOTE,
    _INTERACTION_TERM_NOTE,
    _LABEL_NOTE,
    _LEVELS_NOTE,
    _PERIOD_IN_USE_NOTE,
    _REF_IN_USE_NOTE,
    _REFERENCE_NOTE,
    _RESERVED_NOTE,
    _RESPONSE_LOG_FORMULA,
    _RESPONSE_NAME_FORMULA,
    _ROLE_NOTE,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    _SEQUENCE_FLAG_COUNT_FORMULA,
    _SEQUENCE_NOTE,
    _SEQUENCE_PERIOD_NOTE,
    _TRANSFORM_LOG,
    _TRANSFORM_LOG_DROP,
    _TRANSFORM_NOTE,
    _TYPE_NOTE,
    SPEC_DATASET_PROFILES,
    SpecDatasetProfile,
    _set_spec_block_column_widths,
    _set_spec_block_optional_outline_group,
    _write_intercept_control,
    _write_sequence_status,
)
from .write_spec_block import (
    _FEEDBACK_STATUS_ROW as _SPEC_STATUS_ROW,
)
from .write_spec_block import (
    _status_cell,
    _write_spec_block,
    _write_spec_feedback,
)
from .write_spec_block import (
    _HEADER_ROW as _SPEC_HEADER_ROW,
)
from .write_spec_block import (
    _set_sheet_scoped_names as _set_spec_scoped_names,
)


# ── Layout constants ────────────────────────────────────────────────────────────
# All column indices, row anchors, zone/gap/width tables, chart constants,
# materialization zone constants, and the model-context element table live in
# regression_layout.py and are re-exported here so existing importers keep
# working. See that module for the full documentation.
from .regression_layout import (  # noqa: F401  — re-exported for importers
    ModelContextElement,
    REGRESSION_SHEET_NAME,
    _A_ADJUSTED_R_SQUARED,
    _A_ALPHA,
    _A_BACK_TRANSFORM_METHOD,
    _A_DESIGN_COLUMNS_TOTAL,
    _A_FE_GROUP,
    _A_MEAN_LEVERAGE,
    _A_OBSERVATIONS,
    _A_PRESS,
    _A_PRESS_R_SQUARED,
    _A_QQ_CORRELATION,
    _A_RESIDUAL_DF,
    _A_RESPONSE_READOUT,
    _A_SIGNIFICANCE_F,
    _A_STANDARD_ERROR,
    _BACK_TRANSFORM_DEFAULT,
    _BACK_TRANSFORM_METHODS,
    _BACK_TRANSFORM_NOTE,
    _CHART_GAP,
    _CHART_GRID_COLS,
    _CHART_GRID_ROWS,
    _CHART_HEIGHT,
    _CHART_RIGHT_OFFSET_PT,
    _CHART_WIDTH,
    _CHART_Y_TICK_FORMAT_DEFAULT,
    _CHART_Y_TICK_FORMATS,
    _CONSTRUCTED_DESIGN_MATRIX_LABEL_WIDTH,
    _COOKS_CUTOFF,
    _C_A,
    _C_AA,
    _C_AB,
    _C_AC,
    _C_AD,
    _C_AE,
    _C_AF,
    _C_AG,
    _C_AH,
    _C_AI,
    _C_AJ,
    _C_AK,
    _C_AL,
    _C_AM,
    _C_AN,
    _C_AO,
    _C_AP,
    _C_AQ,
    _C_AR,
    _C_AS,
    _C_AT,
    _C_AU,
    _C_AV,
    _C_AW,
    _C_AX,
    _C_AY,
    _C_AZ,
    _C_BA,
    _C_BB,
    _C_CHART_LABEL_NAME,
    _C_CHART_TITLE,
    _C_CHART_XLABEL,
    _C_CHART_YLABEL,
    _C_DESIGN_MATRIX,
    _C_DESIGN_MATRIX_NAMES,
    _C_GUTTER_AFTER_CHARTS,
    _C_GUTTER_AFTER_CONTEXT,
    _C_GUTTER_AFTER_SAMPLE_INCLUDE,
    _C_MODEL_CONTEXT,
    _C_MODEL_CONTEXT_LABEL,
    _C_MODEL_FORMULA,
    _C_MODEL_FORMULA_LABEL,
    _C_P,
    _C_Q,
    _C_R,
    _C_S,
    _C_SAMPLE_INCLUDE_MATERIALIZED,
    _C_T,
    _C_U,
    _C_V,
    _C_W,
    _C_X,
    _C_Y,
    _C_Z,
    _COLUMN_GROUPS,
    _COLUMN_WIDTHS,
    _DEFINITIONS_PATH,
    _DESIGN_MATRIX_COLUMN_WIDTH,
    _DESIGN_MATRIX_INTERCEPT_HEADER,
    _DESIGN_MATRIX_MAX_COLUMNS,
    _DESIGN_MATRIX_SIZED_COLUMNS,
    _DESIGN_MATRIX_SOFT_CELLS,
    _DESIGN_MATRIX_SOFT_COLUMNS,
    _FORMAT_BAND_LAST_ROW,
    _GAP_COLUMNS,
    _LAST_CHART_COLUMN,
    _MATERIALIZATION_FIRST_ROW,
    _MATERIALIZATION_HEADER_ROW,
    _MATERIALIZATION_SPILL_ROW,
    _MAX_EXCEL_COLUMN,
    _MODEL_CONTEXT_ELEMENTS,
    _MODEL_CONTEXT_LABEL_WIDTH,
    _MODEL_CONTEXT_LAST_ROW,
    _MODEL_CONTEXT_ROWS,
    _MODEL_CONTEXT_VALUE_WIDTH,
    _MODEL_FORMULA_LABEL_WIDTH,
    _PREDICTION_INPUT_NOTE,
    _PREDICTOR_TRANSFORM_FORMULA,
    _PRED_INPUT_FIRST_ROW,
    _PRED_INPUT_LAST_ROW,
    _RESPONSE_TRANSFORM_FORMULA,
    _ROW_ADJUSTED_R_SQUARED,
    _ROW_ALPHA,
    _ROW_ANOVA_RESIDUAL_DF,
    _ROW_CHART_LABELS,
    _ROW_COEFF_FIRST,
    _ROW_DATA_FIRST,
    _ROW_FE_GROUP,
    _ROW_MEAN_LEVERAGE,
    _ROW_MODEL_CONTEXT_CHECK,
    _ROW_MODEL_FORMULA,
    _ROW_OBSERVATIONS,
    _ROW_PRESS,
    _ROW_PRESS_R_SQUARED,
    _ROW_QQ_CORRELATION,
    _ROW_RESPONSE_READOUT,
    _ROW_SIGNIFICANCE_F,
    _ROW_STANDARD_ERROR,
    _SAMPLE_INCLUDE_HEADER,
    _SAMPLE_INCLUDE_MATERIALIZED_WIDTH,
    _WIDTH_COLUMNS,
    _XL_CATEGORY,
    _XL_COLUMN_CLUSTERED,
    _XL_LINE,
    _XL_VALUE,
    _XL_XY_SCATTER,
    _XL_XY_SCATTER_LINES_NO_MARKERS,
    _ZONES,
    _abs_ref,
    _band,
)
from .regression_charts import (  # noqa: F401  — re-exported for importers
    _diagnostic_chart_specs,
    _write_chart_label_cells,
    _write_diagnostic_charts,
)
from .regression_materialization import (  # noqa: F401  — re-exported for importers
    _write_materialization_zone,
)



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
        (3, _C_T, "Pearson R"),
        (3, _C_U, "Spearman R"),
        (3, _C_V, "Skewness"),
        (3, _C_W, "Kurtosis"),
        (3, _C_X, "GVIF"),
        (3, _C_Y, "Tolerance"),
        (5, _C_AA, "Multiple R"),
        (6, _C_AA, "R Square"),
        (7, _C_AA, "Adjusted R Square"),
        (8, _C_AA, "Standard Error"),
        (9, _C_AA, "Observations"),
        (_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL, "Model Formula"),
        (5, _C_AD, "PRESS"),
        (6, _C_AD, "PRESS R²"),
        (7, _C_AD, "Mean Leverage"),
        (8, _C_AD, "AIC"),
        (9, _C_AD, "BIC"),
        (10, _C_AD, "AICc"),
        (11, _C_AD, "QQ Correlation"),
        (12, _C_AD, "Durbin-Watson"),
        (13, _C_AD, "BFN Panel Durbin-Watson"),
        (13, _C_AA, "Alpha"),
        (15, _C_AB, "df"),
        (15, _C_AC, "SS"),
        (15, _C_AD, "MS"),
        (15, _C_AE, "F"),
        (15, _C_AF, "Significance F"),
        (16, _C_AA, "Regression"),
        (17, _C_AA, "Residual"),
        (18, _C_AA, "Total"),
        (21, _C_AB, "Coefficients"),
        (21, _C_AC, "Std Error"),
        (21, _C_AD, "t Stat"),
        (21, _C_AE, "P-value"),
        (21, _C_AF, "Lower 95%"),
        (21, _C_AG, "Upper 95%"),
        (21, _C_AH, "Beta Weight"),
        (18, _C_AL, "Training Mean"),
        (5, _C_AG, "Back-Transform"),
        (6, _C_AG, "Smearing Factor"),
        (7, _C_AG, "R Square (Unit)"),
        (8, _C_AG, "Adj R Square (Unit)"),
        (9, _C_AG, "RMSE (Unit)"),
        (10, _C_AG, "Response Space"),
        (4, _C_AJ, "Point Estimate"),
        (5, _C_AJ, "SE (Mean)"),
        (6, _C_AJ, "SE (New Obs)"),
        (7, _C_AJ, "t Critical"),
        (8, _C_AJ, "CI Lower"),
        (9, _C_AJ, "CI Upper"),
        (10, _C_AJ, "PI Lower"),
        (11, _C_AJ, "PI Upper"),
        (12, _C_AJ, "Confidence Level"),
        (13, _C_AJ, "FE Group"),
        (14, _C_AJ, "Group Mean (y)"),
        (15, _C_AJ, "Group Count"),
        (4, _C_AL, "Point Estimate (Original Units)"),
        (3, _C_AO, "Y"),
        (3, _C_AP, "Predicted Y"),
        (3, _C_AQ, "Residuals"),
        (3, _C_AR, "Hat Diagonal"),
        (3, _C_AS, "Studentized Residuals"),
        (3, _C_AT, "Cook's Distance"),
        (3, _C_AU, "Normal Scores Ranked"),
        (3, _C_AV, "Studentized Residuals Ranked"),
        (3, _C_AW, "Scale-Location"),
        (3, _C_AX, "PRESS Residual"),
        (3, _C_AY, "Cook's Distance (Flagged)"),
        (3, _C_AZ, "Predicted Y (Original Units)"),
        (3, _C_BA, "Residual (Original Units)"),
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
        r_anchor = f"{column}{_ROW_DATA_FIRST}"
        add_expression_format(
            sheet,
            address,
            (
                f"=AND("
                f"ISNUMBER({r_anchor}),"
                f"ABS({r_anchor})>2,"
                f"ABS({r_anchor})<3"
                f")"
            ),
            fill=CF_YELLOW_FILL,
            font_color=CF_DARK_YELLOW_TEXT,
        )

        # |r| >= 3: light-red fill and dark-red text.
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({r_anchor}),ABS({r_anchor})>=3)",
            fill=CF_LIGHT_RED_FILL,
            font_color=CF_DARK_RED_TEXT,
        )

    # ── Cook's distance (and its "Flagged" duplicate) ───────────────────────
    # ONE tier, not two. The old pair — amber at 4/n, red at 0.9 — graded by
    # two unrelated rules of thumb, and F(0.5, p, n−p) lands between them for
    # most models, so keeping either alongside it would draw a line the cutoff
    # itself does not recognize. The Flagged column mirrors Cook's D (blank
    # below the cutoff), so the same rule applied to it just recolors whatever
    # value the source column already produced there — kept visually
    # consistent with the column it duplicates.
    for column, address in [
        (col_letter(_C_AT), addresses["cooks"]),
        (col_letter(_C_AY), addresses["cooks_flag"]),
    ]:
        cell = f"{column}{_ROW_DATA_FIRST}"
        add_expression_format(
            sheet,
            address,
            f"=AND(ISNUMBER({cell}),{cell}>{_COOKS_CUTOFF})",
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

def _setup_local_names(
    sheet: xw.Sheet,
    closures: tuple[CatalogFunction, ...] | None = None,
    source_table_ref: str = "=MileageData[#All]",
) -> None:
    """Register sheet-scoped names used by every formula on this sheet.

    The spec wiring (Source_Data / Header_Names / Spec_* / Allow_Intercept)
    and the constructor closures (Sample_Include / Response_Column /
    Row_Labels / Predictor_Columns / Constructed_Column_Names) are registered by
    ``_set_sheet_scoped_names`` from write_spec_block; this
    function adds the Regression-only names on top.
    """
    # ALWAYS single-quoted. A sheet name containing a space — which every
    # generated test-model sheet has ("M01 Baseline Categoricals") — makes an
    # unquoted RefersTo an invalid formula, and Excel rejects the whole
    # Names.Add with "There's a problem with this formula". Quoting a name
    # that does not need it is always legal, so carrying the quotes on the
    # variable removes the entire class of bug instead of relying on each
    # call site to remember. write_spec_block's own
    # _set_sheet_scoped_names already did this; this function did not, and
    # four of its seven references were unquoted.
    sname = quoted_sheet_name(sheet.name)

    if closures is None:
        closures = load_catalog_document(_DEFINITIONS_PATH).functions_for_sheet(
            REGRESSION_SHEET_NAME
        )

    _set_spec_scoped_names(
        sheet,
        closures,
        source_table_ref=source_table_ref,
    )

    # Zero_Predictors_Selected(): TRUE when the spec contributes no
    # predictor columns — no included Predictor rows, or every included
    # Categorical degenerate. Predictor_Columns() errors in that state (DROP
    # of the sentinel-only accumulator), so the width probe wraps IFERROR
    # rather than counting Include toggles the way v1 did.
    #
    # It must probe Predictor_Columns(), NOT Fit_Design_Columns(): with the
    # intercept relocated into the constructor, Fit_Design_Columns() returns the
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
    # Fit_Design_Columns() does return a well-formed ones column in this state
    # (the intercept stage runs even when the predictor stage is empty), so
    # the engines could in principle fit it directly. The closed-form bypass
    # is kept because it is what the shipped behaviour was verified against;
    # retiring it is a follow-up, not part of the relocation.
    # Intercept_Only_N counts the included rows via SUMPRODUCT over the
    # computed mask (COUNTIF needs a range reference, and Fit_Sample_Include()
    # is an array) so it never errors, even when the mask has zero TRUE rows —
    # callers guard on its value before invoking the FILTER/STDEV.S-based
    # helpers below.
    #
    # The mask is summed with `--`, NOT N(): Fit_Sample_Include() returns a
    # range/array thunk, and N() of such a thunk collapses to the top-left cell
    # (=1), so SUMPRODUCT(N(Fit_Sample_Include())) would return 1 for any
    # non-empty sample. That makes Intercept_Only_N()=1, which fails every
    # N()>=2 inference branch (SE / t / CI / PI return #N/A) and zeroes
    # Intercept_Only_DF() for an intercept-only model with n>1. `--` coerces a
    # range OR array to a summable 1/0 array — the same fix as Log_Domain_Status
    # (see tests/test_spec_block_writer.py and the amber-fix PR).
    drop_local_name(sheet, "Intercept_Only_N")
    sheet.api.Names.Add(
        Name="Intercept_Only_N",
        RefersTo="=LAMBDA(SUMPRODUCT(--(Fit_Sample_Include())))",
    )

    drop_local_name(sheet, "Intercept_Only_Point")
    sheet.api.Names.Add(
        Name="Intercept_Only_Point",
        RefersTo="=LAMBDA(AVERAGE(FILTER(Response_Column(),Fit_Sample_Include())))",
    )

    drop_local_name(sheet, "Intercept_Only_S")
    sheet.api.Names.Add(
        Name="Intercept_Only_S",
        RefersTo="=LAMBDA(STDEV.S(FILTER(Response_Column(),Fit_Sample_Include())))",
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
         "Cook's Distance chart: flagged-point overlay for data labels (D > F(0.5, p, n-p))"),
        ("RegChartObsLabel", col_letter(_C_AN),
         "Cook's Distance chart: observation identifier — the flagged-point overlay's categories"),
    ]:
        drop_local_name(sheet, _name)
        _nm = sheet.api.Names.Add(
            Name=_name,
            RefersTo=(
                f"=OFFSET({sname}!${_col_ltr}$3,1,0,"
                f"MAX(IFERROR({sname}!{_A_OBSERVATIONS},1),1),1)"
            ),
        )
        _nm.Comment = _comment

    # ── v3.3 Comparison_* names (committed public interface for v3.4 Model
    # Comparison). The unit-space block (AG4:AH10) and the Model Formula
    # readout in the §4b materialization band are the surfaces v3.4 reads
    # from — the readout moved out of AB2, and this name is why that move
    # costs a consumer nothing: v3.4 reads the NAME, never the address.
    # Comparison_Anchor is the
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
            f"={sname}!$AH$7:$AH$9",
            "Unit-space goodness-of-fit triplet (R², adjusted R², RMSE) — feeds the v3.4 Model Comparison headline row",
        ),
        (
            "Comparison_Model_Formula",
            f"={sname}!{_abs_ref(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA)}",
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
    Reference In Use displays) is written by the same ``_write_spec_block``
    / ``_write_spec_feedback`` / ``_set_spec_block_column_widths`` family
    in ``write_spec_block.py`` that built the original standalone sheet,
    so the layout cannot drift. Only the zone heading and the
    reserved-column notes are local.

    The block itself is written earlier in ``write_regression_output_sheet``,
    right after the sheet-scoped names — its four computed columns are spills
    that read those names, which is why the names go first. Here we just
    layer the rows 1-2 band (labels, readouts and the per-column status
    cells), the Intercept control, the design-matrix width guard, and the
    column notes on top.
    """
    section_heading(sheet, 1, _C_A, "MODEL SPECIFICATION")
    _write_spec_feedback(sheet)
    _write_sequence_status(sheet)
    _write_intercept_control(sheet)
    _write_design_matrix_width_guard(sheet)
    # Spec-block notes anchor on the header row (row 3) so the tooltip
    # appears when the user hovers the column heading the notes describe,
    # not the first variable row. All twelve spec-block headers carry a
    # plain-language note; the four (Order, Transform, Sequence, Sequence
    # Period) that double as the shipped spec-feature headers use the
    # longer notes defined in write_spec_block.py.
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
    _set_note(sheet, 18, _C_AJ, _PREDICTION_INPUT_NOTE, label="Predictor")
    # Unit-Space Fit (AG5). The Duan/Naive caveat sits on the control it
    # describes, here with every other note for the reason stated above:
    # AddComment is COM-only, so keeping it out of _write_unit_space_block
    # keeps that writer headless-testable.
    _set_note(sheet, 5, _C_AG, _BACK_TRANSFORM_NOTE, label="Back-Transform")


_WIDTH_GUARD_NOTE = (
    "Whether the design matrix this spec describes will compute in Excel.\n\n"
    "RED — the Gram matrix (X'X) cannot be reliably inverted. Excel's MINVERSE "
    "on X'X squares the condition number of the design matrix, so beyond ~200 "
    "constructed columns the inversion produces all-nan results. This is an "
    "empirical limit, dataset-dependent: 200 orthogonal continuous predictors "
    "may invert fine; 200 collinear dummy columns will not. Reduce a "
    "Categorical predictor's level count, exclude predictors, or group them.\n\n"
    "AMBER — it computes, but it is large enough to be slow: MINVERSE is "
    "cubic in the column count, and every materialized cell recalculates on "
    "any input change. Expect the workbook to lag while you edit the spec.\n\n"
    "Both thresholds are read from the SPEC — the Σ above this cell — never "
    "from the built matrix. A matrix too wide to invert cannot be built in order "
    "to be measured, which is the whole point of checking here first.\n\n"
    "A future QR-decomposition path (Coefficients_QR in the roadmap) would "
    "avoid forming X'X entirely, lifting this limit substantially."
)


def _write_design_matrix_width_guard(sheet: xw.Sheet) -> None:
    """The ARCHITECTURE §4b pre-flight width guard, in the spec-block area.

    Three cells, all in the free row-1/row-2 band above the spec block
    (whose own content starts at the row-3 header):

        N1 = "Σ Design Columns"   (bold label)
        O1 = the total            — Σ(spec column O) plus the intercept
                                    column, i.e. exactly COLUMNS(Fit_Design_Columns())
        O2 = the guard status     — blank while the model is within limits,
                                    a WARNING line at the soft threshold, an
                                    ERROR line at the hard one

    O1 sits directly above the per-row Design Columns audit it totals, so the
    number and its breakdown read as one column, and O2 sits directly under the
    number it is a verdict on. The status lives in its own column like every
    other status in this band; its runway is P2/Q2, which the Δ spectrum
    vacated when it moved down to align with the spec block's own rows.

    Both thresholds are read from the SPEC, never from
    ``COLUMNS(Fit_Design_Columns())`` — a matrix too wide to fit cannot be built
    in order to be measured, which is the failure this guard exists to
    prevent. The hard limit is derived from the layout constants (see
    ``_DESIGN_MATRIX_MAX_COLUMNS``), so moving a zone moves the guard with it.

    The total is a display over displays and feeds no constructor, so it does
    not breach "display derives, never feeds" — see ARCHITECTURE §4.
    """
    total_cell = _abs_ref(1, _C_SPEC_DESIGN_COLUMNS)
    status_cell = _abs_ref(_SPEC_STATUS_ROW, _C_SPEC_DESIGN_COLUMNS)

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

    # The verdict itself is Design_Width_Status(), a sheet-scoped catalog
    # LAMBDA (reunify Part 6.2). The thresholds and the message text moved into
    # lambda_functions.json with it; _DESIGN_MATRIX_MAX_COLUMNS and the soft
    # pair remain the source of truth here — they also size the design-matrix
    # band — and test_sheet_writers pins the catalog body's numbers to them.
    #
    # The LAMBDA recomputes k from the spec rather than reading `total_cell`: a
    # catalog body cannot import the _C_* constants, so keeping the read would
    # mean spelling $O$1 into JSON. The expression is the one this function
    # writes into O1 above, so the number is identical; the cells are simply
    # independent now.
    _status_cell(
        sheet,
        _C_SPEC_DESIGN_COLUMNS,
        "=Design_Width_Status()",
        _WIDTH_GUARD_NOTE,
        label="Width guard",
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
        val(sheet, 3, col, header)
    bold_row(sheet, 3, _C_S, _C_Y)

    # Spill anchors at row 4 — each spills once per constructed column.
    # Names come from the constructor twin, so dummies are level-qualified
    # and the stats run on the actual design matrix. GVIF/Tolerance are
    # generalized (Fox & Monette): every dummy column from the same
    # categorical predictor shares one value instead of a separate,
    # coding-dependent number per level.
    f(sheet, 4, _C_S, "=TRANSPOSE(Constructed_Column_Names())")
    f(sheet, 4, _C_T, "=Pearson_R(Predictor_Columns(),Response_Column(),Fit_Sample_Include())")
    f(sheet, 4, _C_U, "=Spearman_R(Predictor_Columns(),Response_Column(),Fit_Sample_Include())")
    f(sheet, 4, _C_V, "=Skewness(Predictor_Columns(),Fit_Sample_Include())")
    f(sheet, 4, _C_W, "=Kurtosis(Predictor_Columns(),Fit_Sample_Include())")
    f(sheet, 4, _C_X, "=GVIF(Predictor_Columns(),Constructed_Column_Names(),Fit_Sample_Include())")
    f(sheet, 4, _C_Y, "=Generalized_Tolerance(Predictor_Columns(),Constructed_Column_Names(),Fit_Sample_Include())")

    sheet.range(
        (rc(4, _C_T)), (rc(_FORMAT_BAND_LAST_ROW, _C_Y))
    ).number_format = "0.00"


def _write_regression_outputs_header(sheet: xw.Sheet) -> None:
    section_heading(sheet, 1, _C_AA, "REGRESSION OUTPUTS")
    section_heading(sheet, 3, _C_AE, "Predicted Variable")
    section_heading(sheet, 3, _C_AF, "")
    # Derived response name — the header of the Role=Response spec row.
    f(sheet, 3, _C_AF, f"={_RESPONSE_NAME_FORMULA}")

    # The Model Formula readout lives at row 1 of the §4b band's design-matrix
    # zone (see _ROW_MODEL_FORMULA / _write_materialization_zone), not in this
    # zone's row 3. It is a caption rather than a headline statistic, and a
    # long string in row 3 — which wraps and AutoFits — would set the height
    # of the whole header row.


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    """Cols AA–AB, rows 4–9."""
    section_heading(sheet, 4, _C_AA, "REGRESSION STATISTICS")
    # Fit-time X/y (Fit_Design_Columns()/Design_Response()): the response is
    # Response_Column() unchanged with no Fixed Effects row and one-way
    # within-demeaned when one is declared
    # — every statistic below reports the "within" flavor under FE, the same
    # convention panel-regression software (e.g. R's plm) uses. Adjusted R²
    # and Standard Error also carry the absorbed df (element 2 of Model_Context,
    # 0 with no FE row) so their df-dependent penalty/divisor is correct.
    # The Regression Statistics zone was the FIRST fully-migrated zone (v3.2):
    # every cell reads the materialized spills through Fit_Design_Columns() /
    # Fit_Sample_Include() instead of re-running the constructors. PR 1 (#223)
    # spiked exactly two cells — one per spill SHAPE — to confirm
    # `#` resolves inside a defined-name RefersTo in Excel; it does, so the rest
    # of the zone followed. The migration is now COMPLETE: every engine call
    # site on this sheet reads the Fit_ readers (Diagnostics, ANOVA,
    # Coefficients, Smearing/Unit-Space, Prediction/FE-group, Residual Output,
    # Predictor Summary, the serial-correlation triggers, and the
    # n/mean_y/sd_y named ranges all migrated with it).
    #
    # A name resolving to the wrong range does not error — it returns numbers
    # from the wrong rows — so every migrated cell lands where the spec-driven
    # verifier compares it cell-by-cell against NumPy. All five rows here are
    # (Multiple R, R², Adjusted R², SE, Observations), which is what made this
    # zone the safe one to complete first.
    #
    # AB9 keeps Design_Response() as a live constructor beside the materialized
    # Fit_Sample_Include — a reminder that Design_Response() is NOT one of the
    # two materialized spills (it is a constructor the engines still evaluate);
    # only Design_Columns and Sample_Include have spill readers.
    for row, label, formula in [
        (5, "Multiple R",        "=Multiple_R(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (6, "R Square",          "=R_Squared(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (7, "Adjusted R Square", "=Adjusted_R_Squared(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (8, "Standard Error",    "=SE_Regression(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (9, "Observations",      "=Observations(Design_Response(),Fit_Sample_Include())"),
    ]:
        val(sheet, row, _C_AA, label)
        f(sheet, row, _C_AB, formula)
    sheet.range(rc(5, _C_AB), rc(8, _C_AB)).number_format = "0.0000"
    sheet.range(rc(9, _C_AB), rc(9, _C_AB)).number_format = "0"
    border_box(sheet, 4, _C_AA, 9, _C_AB)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    """Cols AD–AE, rows 4–13.

    Rows 5-11 (col AE) read the materialized spills through
    Fit_Design_Columns() / Fit_Sample_Include() — the v3.2 rewiring.
    Design_Response() and Fit_Context() stay live; only the design-matrix and
    row-mask arguments move to the spill readers.
    """
    section_heading(sheet, 4, _C_AD, "DIAGNOSTICS")
    # The zone reads the materialized spills through Fit_Design_Columns() /
    # Fit_Sample_Include() instead of re-running the constructors — the v3.2
    # rewiring, zone by zone. Every row here is compared cell-by-cell against
    # NumPy by the spec-driven verifier (regression_spec_sheet_io.py's
    # scalar_specs: PRESS, PRESS_R2, Mean_Leverage, AIC, BIC, AICc,
    # QQ_Correlation), so a name resolving to the wrong range returns numbers
    # from the wrong rows and the gate catches it rather than silently shipping.
    for row, label, formula in [
        (5,  "PRESS",          "=PRESS(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include())"),
        (
            6,
            "PRESS R²",
            "=1-PRESS(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include())"
            "/SS_Total(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())",
        ),
        (
            7,
            "Mean Leverage",
            "=COLUMNS(Fit_Design_Columns())"
            "/Observations(Design_Response(),Fit_Sample_Include())",
        ),
        (8,  "AIC",            "=AIC(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (9,  "BIC",            "=BIC(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (10, "AICc",           "=AICc(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
        (11, "QQ Correlation", "=QQ_Correlation(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"),
    ]:
        val(sheet, row, _C_AD, label)
        f(sheet, row, _C_AE, formula)
    sheet.range(rc(5, _C_AE), rc(11, _C_AE)).number_format = "0.0000"

    # Serial-correlation trigger matrix — two fixed cells (AD12/AE12 plain DW,
    # AD13/AE13 BFN panel DW), each SELF-GUARDING: every state a cell can show is
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
    val(sheet, 12, _C_AD, "Durbin-Watson")
    f(
        sheet,
        12,
        _C_AE,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars>0,"n/a — FE active",'
        "Durbin_Watson_By(Fit_Design_Columns(),Design_Response(),Sequence_Column(),"
        "Fit_Sample_Include())))))",
    )
    sheet.range(rc(12, _C_AE), rc(12, _C_AE)).number_format = "0.000"

    # BFN panel Durbin-Watson (Bhargava–Franzini–Narendranathan 1982): the
    # within-group form for panels under fixed effects. Both this cell and the
    # plain DW cell above read the fit-time pair
    # (Fit_Design_Columns()/Design_Response()), and both must: BFN's own contract
    # says "the residuals are within-demeaned" — Residuals(X, Y) only produces
    # within residuals when X/Y already ARE the within-transformed pair. The DW
    # cell above only ever fires in the no-FE state, where Design_Response()
    # reduces to Response_Column() and Fit_Design_Columns() to the intercept plus
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
    # column; its dormant Cluster branch activates at v3.5 without touching
    # this cell). Interpretation reads like DW (near 2 ⇒ no first-order
    # autocorrelation), but its critical values depend on N and T — surfacing
    # those bounds is a recorded open item, and the standard DW bounds must
    # not be presented next to it (the cell note carries that caveat).
    val(sheet, 13, _C_AD, "BFN Panel Durbin-Watson")
    f(
        sheet,
        13,
        _C_AE,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars=0,"n/a — no fixed effects",'
        'IF(fe_vars>1,"n/a — multiple FE variables",'
        "BFN_Panel_Durbin_Watson(Fit_Design_Columns(),Design_Response(),"
        "Serial_Correlation_Group(),Sequence_Column(),Base_Period_Delta(),"
        "Fit_Sample_Include()))))))",
    )
    sheet.range(rc(13, _C_AE), rc(13, _C_AE)).number_format = "0.000"
    border_box(sheet, 4, _C_AD, 13, _C_AE)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at AB13 — controls prediction interval confidence level."""
    val(sheet, 13, _C_AA, "Alpha")
    bold(sheet, 13, _C_AA)
    val(sheet, 13, _C_AB, 0.05)
    format_input(sheet, 13, _C_AB)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 14–18, cols X–AC."""
    section_heading(sheet, 14, _C_AA, "ANOVA TABLE")

    for col, header in zip(
        [_C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        val(sheet, 15, col, header)
    bold_row(sheet, 15, _C_AA, _C_AF)

    # SST = SSR + SSE must hold under FE too, so every row reads the SAME
    # fit-time pair (Fit_Design_Columns()/Design_Response()) — mixing a raw Total SS against
    # within Regression/Residual SS would break the ANOVA identity.
    val(sheet, 16, _C_AA, "Regression")
    f(sheet, 16, _C_AB, "=Regression_Degrees_Of_Freedom(Fit_Design_Columns(),Fit_Context())")
    f(sheet, 16, _C_AC, "=SS_Regression(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 16, _C_AD, "=MS_Regression(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 16, _C_AE, "=F_Statistic(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 16, _C_AF, "=F_Statistic_P_Value(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")

    val(sheet, 17, _C_AA, "Residual")
    f(sheet, 17, _C_AB, "=Residual_Degrees_Of_Freedom(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 17, _C_AC, "=SS_Residual(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include())")
    f(sheet, 17, _C_AD, "=MS_Residual(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")

    val(sheet, 18, _C_AA, "Total")
    f(sheet, 18, _C_AB, "=Total_Degrees_Of_Freedom(Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 18, _C_AC, "=SS_Total(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")

    sheet.range(rc(16, _C_AB), rc(18, _C_AB)).number_format = "0"
    sheet.range(rc(16, _C_AC), rc(18, _C_AC)).number_format = "0.0"
    sheet.range(rc(16, _C_AD), rc(17, _C_AD)).number_format = "0.0"
    sheet.range(rc(16, _C_AE), rc(16, _C_AE)).number_format = "0.0"
    sheet.range(rc(16, _C_AF), rc(16, _C_AF)).number_format = "0.0E+00"
    border_box(sheet, 14, _C_AA, 18, _C_AF)


def _write_coefficients(sheet: xw.Sheet) -> None:
    """Cols X–AE, rows 20+. Spills downward — nothing placed below row 40 in these cols."""
    section_heading(sheet, 20, _C_AA, "COEFFICIENTS")

    for col, header in zip(
        [_C_AA, _C_AB, _C_AC, _C_AD, _C_AE, _C_AF, _C_AG, _C_AH],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%", "Beta Weight"],
    ):
        val(sheet, 21, col, header)
    bold_row(sheet, 21, _C_AA, _C_AH)

    # Spill row labels aligned to the constructed columns (level-qualified via
    # the constructor twin). Zero_Predictors_Selected() branch computes a real
    # intercept-only model (label "Intercept") instead of fabricating a result
    # for an unselected variable; NA() when nothing is fit.
    f(
        sheet,
        22,
        _C_AA,
        '=IF(Zero_Predictors_Selected(),'
        'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),"Intercept",NA()),'
        'IF(Allow_Intercept,'
        'VSTACK("Intercept",TRANSPOSE(Constructed_Column_Names())),'
        'VSTACK("",TRANSPOSE(Constructed_Column_Names()))))',
    )

    # Spill anchors at row 22 — pad with blank top row when intercept is disabled;
    # zero-predictor branch uses the closed-form intercept-only statistic, or
    # NA() when there is nothing to fit. The mean (Y) only needs one observation;
    # SE/t/p/CI (Z-AA-AD) need at least two to estimate variance, so they're guarded
    # separately rather than sharing the N>=1 check used for the mean.
    f(sheet, 22, _C_AB,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),Intercept_Only_Point(),NA()),'
       'IF(Allow_Intercept,Coefficients(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include()),'
       'VSTACK("",Coefficients(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include()))))')
    f(sheet, 22, _C_AC,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,SE_Coefficients(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context()),'
       'VSTACK("",SE_Coefficients(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context()))))')
    f(sheet, 22, _C_AD,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_Point()/Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,T_Statistics(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context()),'
       'VSTACK("",T_Statistics(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context()))))')
    f(sheet, 22, _C_AE,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'T.DIST.2T(ABS(Intercept_Only_Point()/Intercept_Only_SE()),Intercept_Only_DF()),NA()),'
       'IF(Allow_Intercept,P_Values(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context()),'
       'VSTACK("",P_Values(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context()))))')
    # Confidence_Interval_Lower/Upper's [Context] sits after [Alpha], and
    # Excel LAMBDA calls cannot skip a middle optional argument — 0.05 is
    # passed explicitly here (matching the function's own internal default
    # bit-for-bit) so [Context] can be reached without changing the
    # pre-existing (Alpha-input-independent) 95% CI behavior.
    f(sheet, 22, _C_AF,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()-T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Lower(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),0.05,Fit_Context()),'
       'VSTACK("",Confidence_Interval_Lower(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),0.05,Fit_Context()))))')
    f(sheet, 22, _C_AG,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()+T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Upper(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),0.05,Fit_Context()),'
       'VSTACK("",Confidence_Interval_Upper(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),0.05,Fit_Context()))))')
    # Beta Weights: k×1 (no intercept row); always prepend blank to align with other columns.
    # No predictor exists to standardize in the zero-predictor branch, so render
    # blank (not an error) when Allow_Intercept is TRUE; NA() when nothing is fit.
    f(sheet, 22, _C_AH,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,"",NA()),'
       'VSTACK("",Beta_Weights(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())))')

    for col in [_C_AB, _C_AC, _C_AD, _C_AF, _C_AG, _C_AH]:
        sheet.range(
            rc(22, col), rc(_FORMAT_BAND_LAST_ROW, col)
        ).number_format = "0.0000"
    sheet.range(
        rc(22, _C_AE), rc(_FORMAT_BAND_LAST_ROW, _C_AE)
    ).number_format = "0.0E+00"


def _write_unit_space_block(sheet: xw.Sheet) -> None:
    """v3.3 unit-space / back-transformation block at AG4:AH10.

    Sits between the Coefficients spill (rows 20+) and the Prediction Outputs
    zone (AJ1+). Pair the catalog's `Unit_Space_*` LAMBDA functions with the
    Back-Transform Method input on row 5 (default "Duan") so the prediction
    column (AL) and the residual-zone original-units columns (AZ, BA) can
    stitch onto a single source. Reads are gated by the response Transform
    string read off Fit_Context(); the Response Space readout on row 10 makes
    the active state visible at a glance.

    Row layout:

    | Row | AG                | AH                                         |
    |-----|-------------------|--------------------------------------------|
    | 4   | "UNIT-SPACE FIT"  | (section heading merged across AG4:AH4)    |
    | 5   | "Back-Transform"  | input: "Duan" / "Naive" (default "Duan")   |
    | 6   | "Smearing Factor" | =Smearing_Factor(Fit_Design_Columns(), ...)    |
    | 7   | "R Square (Unit)" | =Unit_Space_R_Squared(...)                 |
    | 8   | "Adj R Square (Unit)" | =Unit_Space_Adjusted_R_Squared(...)    |
    | 9   | "RMSE (Unit)"     | =Unit_Space_RMSE(...)                      |
    | 10  | "Response Space"  | readout (Fit_Context → "Log"/"None")       |
    """
    section_heading(sheet, 4, _C_AG, "UNIT-SPACE FIT")
    val(sheet, 5, _C_AG, "Back-Transform")
    format_input(sheet, 5, _C_AH)
    # A LITERAL default, never a formula. This cell is an INPUT, and the
    # earlier "=IF(OR($AH$5=..." form read its own address: a circular
    # reference, which Excel resolves to 0 with iterative calculation off,
    # so every consumer below (rows 7-9, AL4, AZ, BA) received an
    # unrecognised method and returned #N/A. Same shape as the Alpha input
    # at AB13 (val(sheet, 13, _C_AB, 0.05)) — a typed value is constrained
    # by the list validation, not by the cell re-deriving itself.
    val(sheet, 5, _C_AH, _BACK_TRANSFORM_DEFAULT)
    # Restrict the input to the two supported methods via list validation.
    try:
        sheet.range(rc(5, _C_AH)).api.Validation.Delete()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    try:
        sheet.range(rc(5, _C_AH)).api.Validation.Add(
            Type=3,  # xlValidateList
            AlertStyle=1,
            # Bare comma-separated items, NOT wrapped in quotes. Excel's
            # xlValidateList takes the list as the raw string "Duan,Naive";
            # the quotes people write around it in VBA are that language's
            # string delimiters, not part of the value. Passing them through
            # COM made them literal, so the dropdown offered `"Duan` and
            # `Naive"` — accepted by the validation and rejected by every
            # consumer, since neither matches a recognised method. Same
            # form as _INCLUDE_VALIDATION_LIST ("TRUE,FALSE") next door.
            Formula1=",".join(_BACK_TRANSFORM_METHODS),
        )
        # IgnoreBlank would let a cleared cell through, and a blank method is
        # not one of the six recognised states — it would silently #N/A the
        # whole block rather than being rejected at entry.
        sheet.range(rc(5, _C_AH)).api.Validation.IgnoreBlank = False
    except Exception:  # pylint: disable=broad-exception-caught
        pass

    for row, label, formula in [
        (
            6,
            "Smearing Factor",
            (
                "=Smearing_Factor(Fit_Design_Columns(),Design_Response(),"
                "Fit_Sample_Include(),Fit_Context())"
            ),
        ),
        (
            7,
            "R Square (Unit)",
            (
                "=Unit_Space_R_Squared(Fit_Design_Columns(),Design_Response(),"
                "Response_Column(),Fit_Sample_Include(),Fit_Context(),"
                f"{_A_BACK_TRANSFORM_METHOD})"
            ),
        ),
        (
            8,
            "Adj R Square (Unit)",
            (
                "=Unit_Space_Adjusted_R_Squared(Fit_Design_Columns(),Design_Response(),"
                "Response_Column(),Fit_Sample_Include(),Fit_Context(),"
                f"{_A_BACK_TRANSFORM_METHOD})"
            ),
        ),
        (
            9,
            "RMSE (Unit)",
            (
                "=Unit_Space_RMSE(Fit_Design_Columns(),Design_Response(),"
                "Response_Column(),Fit_Sample_Include(),Fit_Context(),"
                f"{_A_BACK_TRANSFORM_METHOD})"
            ),
        ),
    ]:
        val(sheet, row, _C_AG, label)
        f(sheet, row, _C_AH, formula)
        sheet.range(rc(row, _C_AH), rc(row, _C_AH)).number_format = "0.0000"

    val(sheet, 10, _C_AG, "Response Space")
    f(
        sheet,
        10,
        _C_AH,
        '=IF(Context_Response_Transform(Fit_Context())="Log",'
        '"Original units (back-transformed)","Same as fit space")',
    )
    border_box(sheet, 4, _C_AG, 10, _C_AH)


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    """Zone AJ1:AK15: boxed prediction interval output, plus the FE group
    selector and its group-mean/count readouts.

    v2.1 Fixed Effects group-mean recovery (docs/DECISIONS.md "FE point
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
    val(sheet, 3, _C_AJ, "PREDICTION INTERVAL")
    bold(sheet, 3, _C_AJ)
    # Sub-headers row 3: AK is "Fit Space" (the log/identity values above),
    # AL is "Original Units" (the back-transformed sibling). Both are
    # unmerged so they read as the column header for the point estimate that
    # follows directly underneath.
    val(sheet, 3, _C_AK, "Fit Space")
    val(sheet, 3, _C_AL, "Original Units")
    bold_row(sheet, 3, _C_AJ, _C_AL)
    for row, label in [
        (4, "Point Estimate"),
        (5, "SE (Mean)"),
        (6, "SE (New Obs)"),
        (7, "t Critical"),
        (8, "CI Lower"),
        (9, "CI Upper"),
        (10, "PI Lower"),
        (11, "PI Upper"),
        (12, "Confidence Level"),
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
        4,
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
        "Fit_Sample_Include(),alpha,Fit_Context())))",
    )
    sheet.range(rc(4, _C_AK), rc(12, _C_AK)).number_format = "0.0000"

    # Original Units column (AL): the back-transformed sibling of AK under a
    # Log response. Point estimate (row 4) honors the Method toggle (Duan =
    # conditional mean via smearing factor; Naive = textbook EXP(ŷ)). CI/PI
    # bounds (rows 8–11) ALWAYS use Naive — bounds are quantiles, so EXP(L)
    # to EXP(U) preserves the right coverage; multiplying by smearing would
    # destroy it. Rows 5–7 (SE and t-critical) have no unit-space counterpart
    # and stay blank.
    f(
        sheet,
        4,
        _C_AL,
        (
            "=Back_Transform_Response("
            f"{_abs_ref(4, _C_AK)},"
            "Fit_Context(),"
            f"{_A_BACK_TRANSFORM_METHOD},"
            "Smearing_Factor(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())"
            ")"
        ),
    )
    for row in (8, 9, 10, 11):
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
    sheet.range(rc(4, _C_AL), rc(4, _C_AL)).number_format = "0.0000"
    sheet.range(rc(8, _C_AL), rc(11, _C_AL)).number_format = "0.0000"

    # FE Group selector (row 13): computed-with-override, the same
    # reference-level pattern as the Categorical Reference Level (E) —
    # pre-filled with the alphabetically-first observed group (which is
    # always the "(all)" sentinel itself when no Fixed Effects row is
    # declared, since Prediction_Group_Column() is constant in that state),
    # and directly editable to any other observed group. Red CF flags a
    # typed value that is not among the observed groups.
    val(sheet, 13, _C_AJ, "FE Group")
    f(
        sheet,
        13,
        _C_AK,
        "=INDEX(SORT(UNIQUE(FILTER(Prediction_Group_Column(),Fit_Sample_Include()))),1,1)",
    )
    format_input(sheet, 13, _C_AK)
    fe_group_cell = f"${col_letter(_C_AK)}$13"
    add_expression_format(
        sheet,
        fe_group_cell,
        f"=ISNA(MATCH({fe_group_cell},Prediction_Group_Column(),0))",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Group Mean (y) / Group Count (rows 14-15): the ȳᵢ / Tᵢ readouts
    # DECISIONS.md calls for, computed on the selected group directly via
    # Group_Mean_At / Group_Count_At — the same primitives
    # Group_Prediction_Interval uses internally, so these never disagree
    # with what the interval above actually used.
    val(sheet, 14, _C_AJ, "Group Mean (y)")
    f(
        sheet,
        14,
        _C_AK,
        f"=Group_Mean_At(Response_Column(),Prediction_Group_Column(),"
        f"{_A_FE_GROUP},Fit_Sample_Include())",
    )
    val(sheet, 15, _C_AJ, "Group Count")
    f(
        sheet,
        15,
        _C_AK,
        f"=Group_Count_At(Prediction_Group_Column(),{_A_FE_GROUP},Fit_Sample_Include())",
    )
    sheet.range(rc(14, _C_AK), rc(14, _C_AK)).number_format = "0.0000"
    sheet.range(rc(15, _C_AK), rc(15, _C_AK)).number_format = "0"

    # The Duan/Naive caveat lives as a note on the Back-Transform label at AG5
    # — see _BACK_TRANSFORM_NOTE. It documents the toggle, so it belongs on
    # the toggle: parked three zones away and below the interval it qualified,
    # it would read as a footnote to the prediction block rather than as an
    # explanation of the control that causes the behaviour. Keeping it off
    # row 16 frees the row and lets this box close on the last cell that
    # actually holds a value.
    border_box(sheet, 1, _C_AJ, 15, _C_AL)


def _write_prediction_inputs(sheet: xw.Sheet) -> None:
    """Zone AJ16+: one prediction-input row per constructed column.

    No Intercept row here: Group_Prediction_Interval's pred_input is exactly
    COLUMNS(Predictor_Columns()) raw predictor values with no intercept slot
    — group-mean recovery never uses one (the selected group's own mean plays
    that role) — so an intercept row would be actively misleading. Row 19 is
    a blank spacer between the headers and the first predictor row.
    """
    section_heading(sheet, 17, _C_AJ, "PREDICTION INPUTS")
    val(sheet, 18, _C_AJ, "Predictor")
    val(sheet, 18, _C_AK, "Prediction Value")
    bold_row(sheet, 18, _C_AJ, _C_AL)
    # The band's header note lives in _write_model_specification alongside
    # every other note: AddComment is a COM-only call, and keeping it out of
    # here is what lets this writer stay exercisable through RecordingSheet.

    # AJ20: spill formula — level-qualified names, one per constructed column
    f(sheet, _PRED_INPUT_FIRST_ROW, _C_AJ, "=TRANSPOSE(Constructed_Column_Names())")

    # AL20: the Training Mean column — per-column means of the filtered design
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
    # INDEX into this spill, and the row-4 prediction formula applies
    # Ln_Positive to whatever it finds in AK — so if this spill held the
    # log-space arithmetic mean, the default prediction would silently
    # double-log (ln(ln(x)), not merely un-back-transformed). EXP(mean(ln
    # x)) is exact and self-cancelling: Ln_Positive(EXP(mean(ln x))) =
    # mean(ln x), so the default prediction still lands precisely on
    # Predictor_Columns()'s own centroid, unchanged from the pre-Log-transform behavior.
    # Constructed_Column_Transforms() gives the per-column Log/None flag in
    # the same 1xk shape as the BYCOL means row, so the two combine
    # elementwise before the single TRANSPOSE down into the AL column.
    val(sheet, 18, _C_AL, "Training Mean")
    means_anchor = _abs_ref(_PRED_INPUT_FIRST_ROW, _C_AL)
    f(
        sheet,
        _PRED_INPUT_FIRST_ROW,
        _C_AL,
        (
            "=IFERROR(TRANSPOSE(LET(m,BYCOL(FILTER(Predictor_Columns(),Fit_Sample_Include()),"
            "LAMBDA(c,AVERAGE(c))),t,Constructed_Column_Transforms(),"
            'IF(t="Log",EXP(m),m))),"")'
        ),
    )

    # AK20:AK63 — the Training Mean of each constructed column, individually
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

    # AN3: static header — Row_Labels() supplies its own per-row content
    # (joined Identifier columns, or positional Obs. n labels).
    val(sheet, 3, _C_AN, "Observation")

    # Every one of these columns is fit off Fit_Design_Columns()/Design_Response(), so once a
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
    # own wording instead of "Within": Fit_Design_Columns()/Design_Response() only SUBTRACT the
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
        f(sheet, 3, col, formula)
    # v3.3 unit-space columns: AZ/BA are the original-units siblings of the
    # fit-space Predicted Y / Residuals columns. They carry no (Log) or
    # (Within …) suffix — they are in original units by construction, so the
    # conditional suffix logic above does not apply to them.
    val(sheet, 3, _C_AZ, "Predicted Y (Original Units)")
    val(sheet, 3, _C_BA, "Residual (Original Units)")
    bold_row(sheet, 3, _C_AN, _C_BA)

    # AN4: row labels — the spec-derived Row_Labels() filtered to the sample.
    # Row_Labels() has its own no-Identifier fallback ("Obs. n"), so the only
    # error left to absorb is an all-FALSE mask.
    f(
        sheet, 4, _C_AN,
        "=IFERROR(FILTER(Row_Labels(),Fit_Sample_Include()),NA())",
    )
    # Spill anchors — each spills n rows downward. Fit-time Fit_Design_Columns()/Design_Response()
    # throughout, INCLUDING the "Y" column (AL): under FE the whole table
    # must read as one internally consistent block — Residuals (AN) is an
    # independently-computed column, not a literal AL-AM subtraction, but a
    # raw "Y" next to a within-fitted "Predicted Y" would make the table look
    # broken (Residuals would not visually match Y - Predicted Y). The actual
    # observed response is still available via Response_Column() elsewhere
    # (e.g. Intercept_Only_*); this table shows the model's own fit space.
    f(sheet, 4, _C_AO, "=Dependent_Variable(Design_Response(),Fit_Sample_Include())")
    f(sheet, 4, _C_AP, "=Predictions(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include())")
    f(sheet, 4, _C_AQ, "=Residuals(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include())")
    f(sheet, 4, _C_AR, "=Hat_Diagonal(Fit_Design_Columns(),Fit_Sample_Include())")
    f(sheet, 4, _C_AS, "=Studentized_Residuals(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 4, _C_AT, "=Cooks_Distance(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    f(sheet, 4, _C_AU, "=SORT(Normal_Scores(Design_Response(),Fit_Sample_Include()))")
    f(sheet, 4, _C_AV, "=Studentized_Residuals_Ranked(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())")
    # Scale-Location: SQRT(|Studentized_Residuals|) — horizontal spread should be flat.
    f(
        sheet, 4, _C_AW,
        "=SQRT(ABS(Studentized_Residuals(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())))",
    )
    # PRESS Residual equals the leave-one-out residual e_i / (1 - h_i).
    f(sheet, 4, _C_AX, "=LOOCV_Residual(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include())")
    # Cook's Distance (Flagged): blank except where D exceeds the influence
    # cutoff, F(0.5, p, n−p) — see _COOKS_CUTOFF for why p comes from $O$1 and
    # not the ANOVA Regression df. This feeds the Cook's Distance chart's
    # data-label overlay series (see RegChartCookDistFlag in _setup_local_names
    # / _write_diagnostic_charts), whose labels read this column through
    # "Value From Cells" — which renders the cell's text verbatim, so the
    # non-flagged branch is "" and not NA(): #N/A would print as a literal
    # "#N/A" on every unflagged point. The empty string costs the overlay its
    # skipped points (Excel plots a ""-returning formula at zero rather than
    # treating it as a gap), which is why the chart leaves Value and Category
    # Name off and lets the range supply the whole label.
    cooks_col = col_letter(_C_AT)
    f(
        sheet, 4, _C_AY,
        f'=IF({cooks_col}4#>{_COOKS_CUTOFF},{cooks_col}4#,"")',
    )
    # AZ: Predicted Y in original units — the unit-space siblings of AP/AR/AQ
    # /AX. Method defaults to Duan smearing when the response is Log, which
    # lifts EXP(ŷ) from a median predictor to a mean predictor; under
    # `None` it reduces to the same Predicted Y as AP. BA: Residual in
    # original units = observed response in original units minus the unit-
    # space predicted. Both carry the Method toggle from `AH5` so flipping
    # Duan → Naive propagates here.
    f(
        sheet, 4, _C_AZ,
        (
            "=Unit_Space_Predictions(Fit_Design_Columns(),Design_Response(),"
            "Response_Column(),Fit_Sample_Include(),Fit_Context(),"
            f"{_A_BACK_TRANSFORM_METHOD})"
        ),
    )
    f(
        sheet, 4, _C_BA,
        (
            "=Unit_Space_Residuals(Fit_Design_Columns(),Design_Response(),"
            "Response_Column(),Fit_Sample_Include(),Fit_Context(),"
            f"{_A_BACK_TRANSFORM_METHOD})"
        ),
    )
    # Format every numeric residual-output column — the actual Y (AO) through
    # Residual (Original Units) (BA). Only the AN identifier column (text:
    # country/Obs. labels) is left unformatted.
    sheet.range(f"{col_letter(_C_AO)}:{col_letter(_C_BA)}").number_format = "0.0000"


# ── Public entry point ────────────────────────────────────────────────────────

def write_regression_output_sheet(
    workbook: xw.Book,
    sheet_notes: dict[str, str] | None = None,
    closures: tuple[CatalogFunction, ...] | None = None,
    source_table_ref: str = "=MileageData[#All]",
    spec_profile: SpecDatasetProfile | None = None,
    sheet_name: str = REGRESSION_SHEET_NAME,
    include_charts: bool = True,
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
        default Role/Include/Type/Sequence values) — pre-fills a sensible
        starting model instead of leaving every column an un-flagged
        Predictor. It decides which rows get shipped DEFAULTS, not how
        many spec rows exist: the block sizes itself from
        ``COLUMNS(Source_Data)``. Defaults to the Auto MPG profile
        (``SPEC_DATASET_PROFILES["auto_mpg"]``) when omitted, matching the
        shipped default ``source_table_ref``. Callers that retarget
        ``source_table_ref`` to a different dataset should pass the
        matching entry from ``SPEC_DATASET_PROFILES`` here too — the two
        are independent parameters, not derived from each other, so they
        must be kept in sync by the caller (see build_production.py).
    sheet_name : str, optional
        Which worksheet to write. Defaults to ``"Regression"`` — the
        production sheet. The one-sheet-per-test-model artifact
        (``build_test_models.py``) passes a per-case name so the same
        writer produces every test sheet; see
        ``lambda_catalog/test_model_sheets.py`` for the naming contract.

        Note this is NOT the same string as the ``"Regression"`` passed to
        ``functions_for_sheet`` in ``_setup_local_names``: that one is the
        ``"scope"`` label in lambda_functions.json identifying the
        constructor closures, which are installed sheet-scoped on whatever
        sheet is being written. Same word, different meaning.
    include_charts : bool, optional
        When False, skip ``_write_diagnostic_charts``. The label formula
        cells are still written (they are plain cell writes and part of
        the sheet's content). Generated test-model sheets turn charts off:
        roughly a dozen COM chart objects per sheet across ~45 sheets is
        the single largest cost in that build, and no oracle reads them —
        the production Regression sheet is where chart wiring is verified.
    """

    sheet = next(
        (s for s in workbook.sheets if s.name == sheet_name), None
    )
    if sheet is None:
        sheet = workbook.sheets.add(
            name=sheet_name, after=workbook.sheets[-1]
        )

    for idx in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(idx).Delete()
    for idx in range(sheet.api.ChartObjects().Count, 0, -1):
        sheet.api.ChartObjects(idx).Delete()
    sheet.api.Cells.Clear()
    # Cells.Clear does not touch outline levels — drop any grouping from a
    # previous build before the zone groups are re-applied below. It also does
    # not UNHIDE: ClearOutline alone removes the outline and leaves the columns
    # a collapsed group had hidden still hidden, with no "+" left to expand
    # them. That is exactly the trap this rebuild has to clear when it lands on
    # a workbook built before the §4b spill zones were ungrouped, so use the
    # shared helper, which clears the outline AND unhides every column. The
    # groups this build wants are re-applied below.
    reset_column_groups(sheet)
    safe_activate(sheet)

    # Names FIRST, then the spec block. The block's four computed columns
    # (Period In Use, Levels, Reference In Use, Design Columns) are spill
    # formulas that read the Spec_* bands, Source_Data, Header_Names and the
    # constructor closures, so every one of those names has to exist before
    # they are written.
    #
    # The bands are dataset-sized dynamic ranges (see _spec_band) that
    # depend on no table, so nothing in the names needs the block to have
    # run. The rest of the spec area (headers, feedback, intercept) still
    # runs in _write_model_specification below.
    _setup_local_names(
        sheet,
        closures,
        source_table_ref=source_table_ref,
    )
    _write_spec_block(sheet, spec_profile or SPEC_DATASET_PROFILES["auto_mpg"])

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

    sheet.range(rc(3, _C_S), rc(3, _C_BA)).api.WrapText = True

    # A–O (spec block) widths are owned by write_spec_block.py
    # so the standalone and shared-block builds can never drift.
    _set_spec_block_column_widths(sheet)

    # Content-column widths, per zone, plus the BB post-zone gutter — every
    # entry keyed on its layout constant (see _COLUMN_WIDTHS). The gap columns
    # are sized from _GAP_COLUMNS below so the layout stays declarative — one
    # width there, not one per hard-coded gap letter.
    for column, width in _COLUMN_WIDTHS:
        letter = col_letter(column)
        sheet.range(f"{letter}:{letter}").column_width = width

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

    # Spec-block optional sub-group (E:Q). Nested inside the A:Q zone group
    # so the regular-MLR essentials (A: Variable, B: Role, C: Include, D:
    # Type) stay visible by default while the optional columns
    # (Reference Level, Order, Transform, Sequence, Sequence Period,
    # Period In Use, Levels, Reference In Use, the M/N interaction pair,
    # the O Design Columns audit, and the P/Q Δ-spectrum feedback)
    # collapse to a single "+" button. One click expands them when a
    # Categorical predictor, a Transform, or a Sequence axis enters the
    # spec. Order matters: this must run AFTER
    # the zone-level group above so Excel assigns it outline level 2
    # underneath the level-1 A:Q parent.
    _set_spec_block_optional_outline_group(sheet)

    # Size the sub-header row to its wrapped contents (two lines for the
    # longer residual headers). Must run after the column widths above,
    # since wrap points — and therefore the fitted height — depend on them.
    sheet.api.Rows(3).AutoFit()

    # Chart label formula cells are plain cell writes (no COM chart API), so
    # they're written unconditionally — safe in headless/CI environments and
    # directly unit-testable via RecordingSheet.
    _write_chart_label_cells(sheet)

    # Charts must be positioned after column widths are set so that
    # the chart anchor column's .left reflects the final column layout.
    # Guarded per
    # the documented convention: ChartObjects().Add(...) requires the Excel
    # COM API, which is unavailable in CI/headless environments.
    if include_charts:
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

    # Freeze top 3 rows: zone labels, the spec-zone status band, and every
    # zone's column headers (row 3 everywhere). Requires an active window,
    # which Excel may refuse to grant in a headless/non-interactive session,
    # so this is best-effort.
    try:
        sheet.activate()
        sheet.range("A4").select()
        win = sheet.api.Application.ActiveWindow
        win.FreezePanes = False
        win.SplitRow = 3
        win.SplitColumn = 0
        win.FreezePanes = True
    except Exception:
        pass

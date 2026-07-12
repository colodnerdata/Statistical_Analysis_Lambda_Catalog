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
                   C2) and the Sequence status line (H2); headers row 3; one
                   spec row per table column on rows 4–26; the Sequence Spacing
                   block (Δ candidate / Δ in use / verdict lines / delta
                   spectrum at K–L) on rows 28–34.
  Col M          — thin gap (width 2, ungrouped)
  Col N–T        — Predictor Summary: level-qualified constructed names (N) +
                   Pearson R, Spearman R, Skewness, Kurtosis, VIF, Tolerance —
                   computed on the CONSTRUCTED design matrix (dummies included)
  Col U          — thin gap (width 2, ungrouped)
  Col V–AC       — Regression Outputs: Predicted Variable readout (Z2:AA2),
                   Statistics (V–W rows 3–8), Diagnostics (Y–Z rows 3–12;
                   the serial-correlation pair DW/BFN at rows 11–12),
                   Alpha input (W12), ANOVA Table (rows 13–17, V–AA),
                   Coefficients (rows 19+, V–AC), Beta Weights (AC)
  Col AD         — thin gap (width 2, ungrouped)
  Col AE–AG      — Prediction Outputs: Prediction Interval (AE1:AF8, boxed),
                   Prediction Inputs (AE10+, one row per constructed column),
                   Training Mean spill (AG13 — the single X_s() evaluation
                   the orange AF prefills INDEX into; owns column AG downward
                   so it can never collide with another spill)
  Col AH         — thin gap (width 2, ungrouped)
  Col AI–AT      — Residual Output: heading + Row_Labels() identifiers in AI;
                   11 diagnostics columns (AJ–AT), spills downward from row 3

The spec block replaces the v1 A–B Model Selection zone (predictor toggles +
Allow_Intercept in B2). Everything the v1 sheet hard-wired is now derived:
    y                          → Response_Column()
    All_Xs / x_s()             → Source_Data / X_s() (the spec-driven
                                 design-matrix constructor)
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
    section_heading, val,
)
from .write_sheet_model_construction import (
    _SEQUENCE_PERIOD_NOTE,
    _FIXED_EFFECTS_COUNT_FORMULA,
    _RESERVED_NOTE,
    _RESPONSE_NAME_FORMULA,
    _SEQUENCE_FLAG_COUNT_FORMULA,
    _SEQUENCE_NOTE,
    _C_SEQUENCE_PERIOD as _C_SPEC_SEQUENCE_PERIOD,
    _C_ORDER as _C_SPEC_ORDER,
    _C_REF_IN_USE as _C_SPEC_REF_IN_USE,
    _C_SEQUENCE as _C_SPEC_SEQUENCE,
    _C_TRANSFORM as _C_SPEC_TRANSFORM,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _set_sheet_scoped_names as _set_spec_scoped_names,
    _set_spec_block_column_widths,
    _write_intercept_control,
    _write_sequence_spacing_block,
    _write_spec_block,
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
# share one collapse control. The gap columns are M, U, AD, AH — one before
# each zone that follows the spec block.
#
#   A–L   Model Specification   | M gap | N–T Predictor Summary | U gap |
#   V–AC  Regression Outputs    | AD gap | AE–AG Prediction Outputs | AH gap |
#   AI–AT Residual Output
#
# The gap columns and the per-zone content spans below are the single source
# of truth for both column widths and outline grouping (see _ZONES / _GAP_COLUMNS).

# Zone 1: Model Specification — columns A–L are owned by the shared spec-block
# writers in write_sheet_model_construction (imported above); only the section
# heading cell is written here.
_C_A = 1    # spec: Variable labels / A1 zone heading / A2 Intercept label

# Gap before Predictor Summary.
_C_M = 13   # thin gap (ungrouped — splits the spec and predictor outlines)

# Zone 2: Predictor Summary (constructed columns)
_C_N = 14   # constructed column names (level-qualified)
_C_O = 15   # Pearson R
_C_P = 16   # Spearman R
_C_Q = 17   # Skewness
_C_R = 18   # Kurtosis
_C_S = 19   # VIF
_C_T = 20   # Tolerance

# Gap before Regression Outputs.
_C_U = 21   # thin gap (ungrouped)

# Zone 3: Regression Outputs
_C_V = 22   # labels (stats / ANOVA / coefficients)
_C_W = 23   # stat values / ANOVA df / coefficient values
_C_X = 24   # ANOVA SS / coefficient SE
_C_Y = 25   # diagnostics labels / ANOVA MS / coefficient t-stat
_C_Z = 26   # Predicted Variable label (Z2) / diagnostics values / ANOVA F / coefficient p-value
_C_AA = 27  # predicted variable readout (AA2) / ANOVA Sig F / coefficient CI lower
_C_AB = 28  # coefficient CI upper
_C_AC = 29  # Beta Weights

# Gap before Prediction Outputs.
_C_AD = 30  # thin gap (ungrouped)

# Zone 4: Prediction Outputs
_C_AE = 31  # prediction interval labels / prediction input labels
_C_AF = 32  # prediction interval values / prediction input values
_C_AG = 33  # Training Mean — the per-constructed-column means spill (AG13).
            # The spill owns column AG downward, so it can never collide with
            # another spill when the source data or spec changes.

# Gap before Residual Output.
_C_AH = 34  # thin gap (ungrouped)

# Zone 5: Residual Output
# _C_AI holds Row_Labels() (the identifiers spill); _C_AJ onward hold the
# residual-diagnostic columns. Row-2 headers are written by _write_residuals
# (see the note_cells table); row-3 formulas are the source of truth for what
# each column actually contains.
_C_AI = 35  # row identifiers (Row_Labels() spill)
_C_AJ = 36  # Y (actual dependent variable)
_C_AK = 37  # Predicted Y
_C_AL = 38  # Residuals
_C_AM = 39  # LOOCV Residual
_C_AN = 40  # Hat Diagonal
_C_AO = 41  # Studentized Residuals
_C_AP = 42  # Cook's Distance
_C_AQ = 43  # Normal Scores Ranked
_C_AR = 44  # Studentized Residuals Ranked
_C_AS = 45  # Scale-Location
_C_AT = 46  # PRESS Residual
_C_AU = 47  # wrap-text bound for row-2 header strip (no own column)

# The constructed-column count is spec-dependent (19 on the default WHO spec),
# so bands that v1 sized with the fixed k=18 now cover a generous fixed range.
_PRED_INPUT_FIRST_ROW = 13
_PRED_INPUT_LAST_ROW = 62
_FORMAT_BAND_LAST_ROW = 62

# Content zones as (first_col, last_col) spans — the single source of truth for
# the outline groups. Each pair becomes one collapsible column group; the gap
# columns between them (below) stay ungrouped so the zones collapse
# independently.
_ZONES: tuple[tuple[int, int], ...] = (
    (_C_A, _C_SPEC_REF_IN_USE),  # A:L  — Model Specification
    (_C_N, _C_T),                # N:T  — Predictor Summary
    (_C_V, _C_AC),               # V:AC — Regression Outputs
    (_C_AE, _C_AG),              # AE:AG — Prediction Outputs
    (_C_AI, _C_AT),              # AI:AT — Residual Output
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
_XL_CATEGORY = 1             # horizontal axis
_XL_VALUE = 2                # vertical axis
_CHART_WIDTH = 310.0         # points
_CHART_HEIGHT = 310.0        # points
_CHART_GAP = 10.0            # gap between charts in points

# ── Visual formatting helpers ─────────────────────────────────────────────────


def _input_range(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    sheet.range(rc(r1, c1), rc(r2, c2)).color = _INPUT


def _set_note(sheet: xw.Sheet, row: int, col: int, text: str) -> None:
    """Replace the cell's note/comment text with a plain-language explanation."""
    cell_api = sheet.range(rc(row, col)).api
    try:
        cell_api.ClearComments()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    cell_api.AddComment(text)
    cell_api.Comment.Visible = False


def _annotate_statistical_terms(sheet: xw.Sheet, sheet_notes: dict[str, str]) -> None:
    """Attach plain-language notes to key statistical labels on the sheet."""
    note_cells = [
        (2, _C_O, "Pearson R"),
        (2, _C_P, "Spearman R"),
        (2, _C_Q, "Skewness"),
        (2, _C_R, "Kurtosis"),
        (2, _C_S, "VIF"),
        (2, _C_T, "Tolerance"),
        (4, _C_V, "Multiple R"),
        (5, _C_V, "R Square"),
        (6, _C_V, "Adjusted R Square"),
        (7, _C_V, "Standard Error"),
        (8, _C_V, "Observations"),
        (4, _C_Y, "PRESS"),
        (5, _C_Y, "PRESS R²"),
        (6, _C_Y, "Mean Leverage"),
        (7, _C_Y, "AIC"),
        (8, _C_Y, "BIC"),
        (9, _C_Y, "AICc"),
        (10, _C_Y, "QQ Correlation"),
        (11, _C_Y, "Durbin-Watson"),
        (12, _C_Y, "BFN Panel Durbin-Watson"),
        (12, _C_V, "Alpha"),
        (14, _C_W, "df"),
        (14, _C_X, "SS"),
        (14, _C_Y, "MS"),
        (14, _C_Z, "F"),
        (14, _C_AA, "Significance F"),
        (15, _C_V, "Regression"),
        (16, _C_V, "Residual"),
        (17, _C_V, "Total"),
        (20, _C_W, "Coefficients"),
        (20, _C_X, "Std Error"),
        (20, _C_Y, "t Stat"),
        (20, _C_Z, "P-value"),
        (20, _C_AA, "Lower 95%"),
        (20, _C_AB, "Upper 95%"),
        (20, _C_AC, "Beta Weight"),
        (3, _C_AE, "Point Estimate"),
        (4, _C_AE, "SE Prediction"),
        (5, _C_AE, "t Critical"),
        (6, _C_AE, "Lower 95%"),
        (7, _C_AE, "Upper 95%"),
        (8, _C_AE, "Confidence Level"),
        (2, _C_AJ, "Y"),
        (2, _C_AK, "Predicted Y"),
        (2, _C_AL, "Residuals"),
        (2, _C_AM, "LOOCV Residual"),
        (2, _C_AN, "Hat Diagonal"),
        (2, _C_AO, "Studentized Residuals"),
        (2, _C_AP, "Cook's Distance"),
        (2, _C_AQ, "Normal Scores Ranked"),
        (2, _C_AR, "Studentized Residuals Ranked"),
        (2, _C_AS, "Scale-Location"),
        (2, _C_AT, "PRESS Residual"),
    ]

    for row, col, key in note_cells:
        note_text = sheet_notes.get(key)
        if note_text is not None:
            _set_note(sheet, row, col, note_text)


def _write_significance_conditional_formatting(sheet: xw.Sheet) -> None:
    """Flag nonsignificant coefficient and overall-model P-values."""

    coefficient_p_values = f"Z21:Z{MAX_EXCEL_ROW}"
    significance_f = "AA15"

    sheet.range(coefficient_p_values).api.FormatConditions.Delete()
    sheet.range(significance_f).api.FormatConditions.Delete()

    # Individual coefficient P-values above alpha.
    add_expression_format(
        sheet,
        coefficient_p_values,
        "=AND(ISNUMBER(Z21),Z21>$W$12)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Overall regression P-value above alpha.
    add_expression_format(
        sheet,
        significance_f,
        "=AND(ISNUMBER(AA15),AA15>$W$12)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_residual_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply diagnostic cutoffs to the residual-output columns."""

    addresses = {
        "hat":                f"AN3:AN{MAX_EXCEL_ROW}",
        "studentized":        f"AO3:AO{MAX_EXCEL_ROW}",
        "cooks":              f"AP3:AP{MAX_EXCEL_ROW}",
        "studentized_ranked": f"AR3:AR{MAX_EXCEL_ROW}",
        "scale_location":     f"AS3:AS{MAX_EXCEL_ROW}",
        "press_residual":     f"AT3:AT{MAX_EXCEL_ROW}",
    }

    # Remove existing rules so repeated builds do not duplicate them.
    for address in addresses.values():
        sheet.range(address).api.FormatConditions.Delete()

    # ── Hat diagonal ─────────────────────────────────────────────────────────
    # Z6 contains mean leverage, p/n.
    # > 2p/n: light-red fill and dark-red text.
    add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AN3),AN3>2*$Z$6)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # > 3p/n: additionally bold.
    add_expression_format(
        sheet,
        addresses["hat"],
        "=AND(ISNUMBER(AN3),AN3>3*$Z$6)",
        bold_font=True,
    )

    # ── Studentized residuals ────────────────────────────────────────────────
    for column, address in [
        ("AO", addresses["studentized"]),
        ("AR", addresses["studentized_ranked"]),
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

    # ── Cook's distance ──────────────────────────────────────────────────────
    # W8 contains the number of observations, n.
    # 4/n < D <= 0.9: light-yellow fill and dark-yellow text.
    add_expression_format(
        sheet,
        addresses["cooks"],
        "=AND(ISNUMBER(AP3),AP3>4/$W$8,AP3<=0.9)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # D > 0.9: light-red fill and dark-red text.
    add_expression_format(
        sheet,
        addresses["cooks"],
        "=AND(ISNUMBER(AP3),AP3>0.9)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── Scale-Location: SQRT(|Studentized|) ─────────────────────────────────
    # SQRT(2) ≈ 1.414 corresponds to |Studentized| = 2.
    # SQRT(3) ≈ 1.732 corresponds to |Studentized| = 3.
    add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AS3),AS3>1.414,AS3<=1.732)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["scale_location"],
        "=AND(ISNUMBER(AS3),AS3>1.732)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS Residual: e_i / (1 - h_i) ─────────────────────────────────────
    # W7 contains the Standard Error of the regression.
    # |PRESS| > 2*SE: mild concern; > 3*SE: strong concern.
    add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AT3),ABS(AT3)>2*$W$7,ABS(AT3)<=3*$W$7)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        addresses["press_residual"],
        "=AND(ISNUMBER(AT3),ABS(AT3)>3*$W$7)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_model_diagnostic_conditional_formatting(sheet: xw.Sheet) -> None:
    """Apply rule-of-thumb formatting to VIF, PRESS R², and QQ Correlation."""

    vif_address = f"S3:S{MAX_EXCEL_ROW}"
    press_r2_address = "Z5"
    qq_corr_address = "Z10"

    # Prevent duplicate rules when rebuilding the sheet.
    sheet.range(vif_address).api.FormatConditions.Delete()
    sheet.range(press_r2_address).api.FormatConditions.Delete()
    sheet.range(qq_corr_address).api.FormatConditions.Delete()

    # ── VIF ─────────────────────────────────────────────────────────────────
    # 5 < VIF <= 10: possible multicollinearity; review.
    add_expression_format(
        sheet,
        vif_address,
        "=AND(ISNUMBER(S3),S3>5,S3<=10)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # VIF > 10: strong multicollinearity warning.
    add_expression_format(
        sheet,
        vif_address,
        "=AND(ISNUMBER(S3),S3>10)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── PRESS R² ─────────────────────────────────────────────────────────────
    # Negative PRESS R² means cross-validated predictions perform worse than
    # predicting the outcome mean.
    add_expression_format(
        sheet,
        press_r2_address,
        "=AND(ISNUMBER(Z5),Z5<0)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── QQ Correlation ────────────────────────────────────────────────────────
    # Pearson r of sorted scaled residuals vs. normal quantiles; near 1.0 = normal errors.
    # < 0.98: mild departure (yellow); < 0.95: stronger departure (red).
    add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(Z10),Z10<0.98,Z10>=0.95)",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )
    add_expression_format(
        sheet,
        qq_corr_address,
        "=AND(ISNUMBER(Z10),Z10<0.95)",
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
    sheet: xw.Sheet, closures: tuple[CatalogFunction, ...] | None = None
) -> None:
    """Register sheet-scoped names used by every formula on this sheet.

    The spec wiring (Source_Data / Header_Names / Spec_* / Allow_Intercept)
    and the constructor closures (Sample_Include / Response_Column /
    Row_Labels / X_s / Constructed_Column_Names) are registered by the shared
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

    _set_spec_scoped_names(sheet, closures)

    # Zero_Predictors_Selected(): TRUE when the spec contributes no design
    # columns — no included Predictor rows, or every included Categorical
    # degenerate. X_s() errors in that state (DROP of the sentinel-only
    # accumulator), so the width probe wraps IFERROR rather than counting
    # Include toggles the way v1 did.
    drop_local_name(sheet, "Zero_Predictors_Selected")
    sheet.api.Names.Add(
        Name="Zero_Predictors_Selected",
        RefersTo="=LAMBDA(IFERROR(COLUMNS(X_s()),0)=0)",
    )

    # alpha: confidence level input, lives in W12
    drop_local_name(sheet, "alpha")
    sheet.api.Names.Add(
        Name="alpha",
        RefersTo=f"={sname}!$W$12",
    )

    # ── Intercept-only closed-form helpers ──────────────────────────────────
    # Used by _write_coefficients and _write_prediction_interval when
    # Zero_Predictors_Selected() is TRUE and Allow_Intercept is TRUE: an
    # intercept-only OLS model (Y = b0 + error) is still statistically
    # well-defined even though X_s() has nothing to construct. Bypasses
    # X_s()/Coefficients()/Prediction_Interval() entirely since Excel cannot
    # represent a valid zero-column array.
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

    # ── Chart data ranges (OFFSET-based, sized to n = $W$8 observations) ────────
    # These worksheet-scoped names feed chart SERIES formulas as
    # ='Regression'!<Name>, avoiding full-column references that degrade
    # performance and avoiding the unsupported # spill operator in chart formulas.
    # The third element becomes the Name Manager comment so a user browsing
    # names can tell which chart each range feeds.
    for _name, _col_ltr, _comment in [
        ("RegChartQQX", col_letter(_C_AQ),
         "Normal Q-Q chart: X values (theoretical quantiles, Normal Scores Ranked)"),
        ("RegChartQQY", col_letter(_C_AR),
         "Normal Q-Q chart: Y values (Studentized Residuals Ranked)"),
        ("RegChartFitY", col_letter(_C_AK),
         "Predicted Y: X values for the Residuals vs. Fitted, Actual vs. Predicted, and Scale-Location charts"),
        ("RegChartResid", col_letter(_C_AL),
         "Residuals vs. Fitted chart: Y values (Residuals)"),
        ("RegChartActY", col_letter(_C_AJ),
         "Actual vs. Predicted chart: Y values (Actual Y)"),
        ("RegChartScaleLoc", col_letter(_C_AS),
         "Scale-Location chart: Y values (sqrt of abs Studentized Residual)"),
        ("RegChartCookDist", col_letter(_C_AP),
         "Cook's Distance chart: bar values"),
        ("RegChartLeverage", col_letter(_C_AN),
         "Studentized Residuals vs. Leverage chart: X values (Hat Diagonal)"),
        ("RegChartStudResid", col_letter(_C_AO),
         "Studentized Residuals vs. Leverage chart: Y values"),
        ("RegChartPRESSResid", col_letter(_C_AT),
         "PRESS Residuals chart: bar values"),
    ]:
        drop_local_name(sheet, _name)
        _nm = sheet.api.Names.Add(
            Name=_name,
            RefersTo=f"=OFFSET('{sname}'!${_col_ltr}$2,1,0,MAX(IFERROR('{sname}'!$W$8,1),1),1)",
        )
        _nm.Comment = _comment


# ── Section writers ───────────────────────────────────────────────────────────

def _write_model_specification(sheet: xw.Sheet) -> None:
    """Zone A–K: the shared spec block + row-2 Intercept control.

    The block itself (headers, defaults, dropdowns, CF, the Levels and
    Reference In Use displays) is written by the same functions that build
    the standalone Model Construction sheet, so the two layouts can never
    drift. Only the zone heading and the reserved-column notes are local.
    """
    section_heading(sheet, 1, _C_A, "MODEL SPECIFICATION")
    _write_spec_block(sheet)
    _write_sequence_spacing_block(sheet)
    _write_intercept_control(sheet)
    _set_note(sheet, _SPEC_FIRST_DATA_ROW, _C_SPEC_ORDER, _RESERVED_NOTE)
    _set_note(sheet, _SPEC_FIRST_DATA_ROW, _C_SPEC_TRANSFORM, _RESERVED_NOTE)
    _set_note(sheet, _SPEC_FIRST_DATA_ROW, _C_SPEC_SEQUENCE, _SEQUENCE_NOTE)
    _set_note(sheet, _SPEC_FIRST_DATA_ROW, _C_SPEC_SEQUENCE_PERIOD, _SEQUENCE_PERIOD_NOTE)


def _write_predictor_summary(sheet: xw.Sheet) -> None:
    """Zone N–T: EDA stats for the CONSTRUCTED design-matrix columns."""
    section_heading(sheet, 1, _C_N, "PREDICTOR SUMMARY")

    for col, header in zip(
        [_C_N, _C_O, _C_P, _C_Q, _C_R, _C_S, _C_T],
        ["", "Pearson R", "Spearman R", "Skewness", "Kurtosis", "VIF", "Tolerance"],
    ):
        val(sheet, 2, col, header)
    bold_row(sheet, 2, _C_N, _C_T)

    # Spill anchors at row 3 — each spills once per constructed column.
    # Names come from the constructor twin, so dummies are level-qualified
    # and the stats run on the actual design matrix (VIF with dummies in).
    f(sheet, 3, _C_N, "=TRANSPOSE(Constructed_Column_Names())")
    f(sheet, 3, _C_O, "=Pearson_R(X_s(),Response_Column(),Sample_Include())")
    f(sheet, 3, _C_P, "=Spearman_R(X_s(),Response_Column(),Sample_Include())")
    f(sheet, 3, _C_Q, "=Skewness(X_s(),Sample_Include())")
    f(sheet, 3, _C_R, "=Kurtosis(X_s(),Sample_Include())")
    f(sheet, 3, _C_S, "=VIF(X_s(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_T, "=Tolerance(X_s(),Allow_Intercept,Sample_Include())")

    sheet.range(
        (rc(3, _C_O)), (rc(_FORMAT_BAND_LAST_ROW, _C_T))
    ).number_format = "0.00"


def _write_regression_outputs_header(sheet: xw.Sheet) -> None:
    section_heading(sheet, 1, _C_V, "REGRESSION OUTPUTS")
    section_heading(sheet, 2, _C_Z, "Predicted Variable")
    section_heading(sheet, 2, _C_AA, "")
    # Derived response name — the header of the Role=Response spec row.
    f(sheet, 2, _C_AA, f"={_RESPONSE_NAME_FORMULA}")


def _write_regression_statistics(sheet: xw.Sheet) -> None:
    """Cols V–W, rows 3–8."""
    section_heading(sheet, 3, _C_V, "REGRESSION STATISTICS")
    for row, label, formula in [
        (4, "Multiple R",        "=Multiple_R(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (5, "R Square",          "=R_Squared(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (6, "Adjusted R Square", "=Adjusted_R_Squared(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (7, "Standard Error",    "=SE_Regression(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (8, "Observations",      "=Observations(Response_Column(),Sample_Include())"),
    ]:
        val(sheet, row, _C_V, label)
        f(sheet, row, _C_W, formula)
    sheet.range(rc(4, _C_W), rc(7, _C_W)).number_format = "0.0000"
    sheet.range(rc(8, _C_W), rc(8, _C_W)).number_format = "0"
    border_box(sheet, 3, _C_V, 8, _C_W)


def _write_diagnostics(sheet: xw.Sheet) -> None:
    """Cols Y–Z, rows 3–12."""
    section_heading(sheet, 3, _C_Y, "DIAGNOSTICS")
    for row, label, formula in [
        (4,  "PRESS",          "=PRESS(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (
            5,
            "PRESS R²",
            "=1-PRESS(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"
            "/SS_Total(Response_Column(),Allow_Intercept,Sample_Include())",
        ),
        (
            6,
            "Mean Leverage",
            "=(Regression_Degrees_Of_Freedom(X_s())+IF(Allow_Intercept,1,0))"
            "/Observations(Response_Column(),Sample_Include())",
        ),
        (7,  "AIC",            "=AIC(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (8,  "BIC",            "=BIC(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (9,  "AICc",           "=AICc(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
        (10, "QQ Correlation", "=QQ_Correlation(X_s(),Response_Column(),Allow_Intercept,Sample_Include())"),
    ]:
        val(sheet, row, _C_Y, label)
        f(sheet, row, _C_Z, formula)
    sheet.range(rc(4, _C_Z), rc(10, _C_Z)).number_format = "0.0000"

    # Serial-correlation trigger matrix — two fixed cells (Y11/Z11 plain DW,
    # Y12/Z12 BFN panel DW), each SELF-GUARDING: every state a cell can show is
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
    # Zero flags → "requires Sequence"; two-plus flags is the spec error the H2
    # status line already reports, and computing on Sequence_Column()'s first
    # match would be ambiguous → "multiple Sequence flags". Under fixed effects
    # the single-series statistic is invalid twice over (within-demeaned
    # residuals; no single panel ordering) → "FE active", and the BFN cell
    # below takes over. With exactly one flag and no FE, Durbin_Watson_By sorts
    # residuals by Sequence_Column() before differencing — all array work
    # internal, scalar out, no spill.
    val(sheet, 11, _C_Y, "Durbin-Watson")
    f(
        sheet,
        11,
        _C_Z,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars>0,"n/a — FE active",'
        "Durbin_Watson_By(X_s(),Response_Column(),Sequence_Column(),"
        "Allow_Intercept,Sample_Include())))))",
    )
    sheet.range(rc(11, _C_Z), rc(11, _C_Z)).number_format = "0.000"

    # BFN panel Durbin-Watson (Bhargava–Franzini–Narendranathan 1982): the
    # within-group form for panels under fixed effects. Differencing is
    # restricted to within-group (group, seq−Δ) pairs via Difference_By inside
    # the LAMBDA, so group seams contribute nothing by construction. Active
    # only when a Sequence axis AND a Fixed Effects variable are declared;
    # multiple FE variables (two-way absorption, out of scope) show a token
    # rather than computing on an arbitrary first match. The group argument is
    # Serial_Correlation_Group() — the grouping-key resolver — NOT the FE
    # column accessor directly: the resolver is the single retargeting point
    # for which dimension partitions residuals (today the Fixed Effects
    # column; its dormant Cluster branch activates at v2.6+ without touching
    # this cell). Interpretation reads like DW (near 2 ⇒ no first-order
    # autocorrelation), but its critical values depend on N and T — surfacing
    # those bounds is a recorded open item, and the standard DW bounds must
    # not be presented next to it (the cell note carries that caveat).
    val(sheet, 12, _C_Y, "BFN Panel Durbin-Watson")
    f(
        sheet,
        12,
        _C_Z,
        f"=LET(seq_flags,{_SEQUENCE_FLAG_COUNT_FORMULA},"
        f"fe_vars,{_FIXED_EFFECTS_COUNT_FORMULA},"
        'IF(seq_flags=0,"n/a — requires Sequence",'
        'IF(seq_flags>1,"n/a — multiple Sequence flags",'
        'IF(fe_vars=0,"n/a — no fixed effects",'
        'IF(fe_vars>1,"n/a — multiple FE variables",'
        "BFN_Panel_Durbin_Watson(X_s(),Response_Column(),"
        "Serial_Correlation_Group(),Sequence_Column(),Base_Period_Delta(),"
        "Allow_Intercept,Sample_Include()))))))",
    )
    sheet.range(rc(12, _C_Z), rc(12, _C_Z)).number_format = "0.000"
    border_box(sheet, 3, _C_Y, 12, _C_Z)


def _write_alpha(sheet: xw.Sheet) -> None:
    """Alpha input cell at W12 — controls prediction interval confidence level."""
    val(sheet, 12, _C_V, "Alpha")
    bold(sheet, 12, _C_V)
    val(sheet, 12, _C_W, 0.05)
    format_input(sheet, 12, _C_W)


def _write_anova(sheet: xw.Sheet) -> None:
    """ANOVA table, rows 13–17, cols V–AA."""
    section_heading(sheet, 13, _C_V, "ANOVA TABLE")

    for col, header in zip(
        [_C_V, _C_W, _C_X, _C_Y, _C_Z, _C_AA],
        ["", "df", "SS", "MS", "F", "Significance F"],
    ):
        val(sheet, 14, col, header)
    bold_row(sheet, 14, _C_V, _C_AA)

    val(sheet, 15, _C_V, "Regression")
    f(sheet, 15, _C_W, "=Regression_Degrees_Of_Freedom(X_s())")
    f(sheet, 15, _C_X, "=SS_Regression(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 15, _C_Y, "=MS_Regression(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 15, _C_Z, "=F_Statistic(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 15, _C_AA, "=F_Statistic_P_Value(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")

    val(sheet, 16, _C_V, "Residual")
    f(sheet, 16, _C_W, "=Residual_Degrees_Of_Freedom(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 16, _C_X, "=SS_Residual(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 16, _C_Y, "=MS_Residual(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")

    val(sheet, 17, _C_V, "Total")
    f(sheet, 17, _C_W, "=Total_Degrees_Of_Freedom(Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 17, _C_X, "=SS_Total(Response_Column(),Allow_Intercept,Sample_Include())")

    sheet.range(rc(15, _C_W), rc(17, _C_W)).number_format = "0"
    sheet.range(rc(15, _C_X), rc(17, _C_X)).number_format = "0.0"
    sheet.range(rc(15, _C_Y), rc(16, _C_Y)).number_format = "0.0"
    sheet.range(rc(15, _C_Z), rc(15, _C_Z)).number_format = "0.0"
    sheet.range(rc(15, _C_AA), rc(15, _C_AA)).number_format = "0.0E+00"
    border_box(sheet, 13, _C_V, 17, _C_AA)


def _write_coefficients(sheet: xw.Sheet) -> None:
    """Cols V–AC, rows 19+. Spills downward — nothing placed below row 39 in these cols."""
    section_heading(sheet, 19, _C_V, "COEFFICIENTS")

    for col, header in zip(
        [_C_V, _C_W, _C_X, _C_Y, _C_Z, _C_AA, _C_AB, _C_AC],
        ["", "Coefficients", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%", "Beta Weight"],
    ):
        val(sheet, 20, col, header)
    bold_row(sheet, 20, _C_V, _C_AC)

    # Spill row labels aligned to the constructed columns (level-qualified via
    # the constructor twin). Zero_Predictors_Selected() branch computes a real
    # intercept-only model (label "Intercept") instead of fabricating a result
    # for an unselected variable; NA() when nothing is fit.
    f(
        sheet,
        21,
        _C_V,
        '=IF(Zero_Predictors_Selected(),'
        'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),"Intercept",NA()),'
        'IF(Allow_Intercept,'
        'VSTACK("Intercept",TRANSPOSE(Constructed_Column_Names())),'
        'VSTACK("",TRANSPOSE(Constructed_Column_Names()))))',
    )

    # Spill anchors at row 21 — pad with blank top row when intercept is disabled;
    # zero-predictor branch uses the closed-form intercept-only statistic, or
    # NA() when there is nothing to fit. The mean (V) only needs one observation;
    # SE/t/p/CI (X-AB) need at least two to estimate variance, so they're guarded
    # separately rather than sharing the N>=1 check used for the mean.
    f(sheet, 21, _C_W,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=1),Intercept_Only_Point(),NA()),'
       'IF(Allow_Intercept,Coefficients(X_s(),Response_Column(),Allow_Intercept,Sample_Include()),'
       'VSTACK("",Coefficients(X_s(),Response_Column(),Allow_Intercept,Sample_Include()))))')
    f(sheet, 21, _C_X,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,SE_Coefficients(X_s(),Response_Column(),Allow_Intercept,Sample_Include()),'
       'VSTACK("",SE_Coefficients(X_s(),Response_Column(),Allow_Intercept,Sample_Include()))))')
    f(sheet, 21, _C_Y,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),Intercept_Only_Point()/Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,T_Statistics(X_s(),Response_Column(),Allow_Intercept,Sample_Include()),'
       'VSTACK("",T_Statistics(X_s(),Response_Column(),Allow_Intercept,Sample_Include()))))')
    f(sheet, 21, _C_Z,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'T.DIST.2T(ABS(Intercept_Only_Point()/Intercept_Only_SE()),Intercept_Only_DF()),NA()),'
       'IF(Allow_Intercept,P_Values(X_s(),Response_Column(),Allow_Intercept,Sample_Include()),'
       'VSTACK("",P_Values(X_s(),Response_Column(),Allow_Intercept,Sample_Include()))))')
    f(sheet, 21, _C_AA,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()-T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Lower(X_s(),Response_Column(),Allow_Intercept,Sample_Include()),'
       'VSTACK("",Confidence_Interval_Lower(X_s(),Response_Column(),Allow_Intercept,Sample_Include()))))')
    f(sheet, 21, _C_AB,
       '=IF(Zero_Predictors_Selected(),'
       'IF(AND(Allow_Intercept,Intercept_Only_N()>=2),'
       'Intercept_Only_Point()+T.INV.2T(alpha,Intercept_Only_DF())*Intercept_Only_SE(),NA()),'
       'IF(Allow_Intercept,Confidence_Interval_Upper(X_s(),Response_Column(),Allow_Intercept,Sample_Include()),'
       'VSTACK("",Confidence_Interval_Upper(X_s(),Response_Column(),Allow_Intercept,Sample_Include()))))')
    # Beta Weights: k×1 (no intercept row); always prepend blank to align with other columns.
    # No predictor exists to standardize in the zero-predictor branch, so render
    # blank (not an error) when Allow_Intercept is TRUE; NA() when nothing is fit.
    f(sheet, 21, _C_AC,
       '=IF(Zero_Predictors_Selected(),'
       'IF(Allow_Intercept,"",NA()),'
       'VSTACK("",Beta_Weights(X_s(),Response_Column(),Allow_Intercept,Sample_Include())))')

    for col in [_C_W, _C_X, _C_Y, _C_AA, _C_AB, _C_AC]:
        sheet.range(
            rc(21, col), rc(_FORMAT_BAND_LAST_ROW, col)
        ).number_format = "0.0000"
    sheet.range(
        rc(21, _C_Z), rc(_FORMAT_BAND_LAST_ROW, _C_Z)
    ).number_format = "0.0E+00"


def _write_prediction_interval(sheet: xw.Sheet) -> None:
    """Zone AE1:AF8: boxed prediction interval output."""
    section_heading(sheet, 1, _C_AE, "PREDICTION OUTPUTS")
    val(sheet, 2, _C_AE, "PREDICTION INTERVAL")
    bold(sheet, 2, _C_AE)
    for row, label in [
        (3, "Point Estimate"),
        (4, "SE Prediction"),
        (5, "t Critical"),
        (6, "Lower 95%"),
        (7, "Upper 95%"),
        (8, "Confidence Level"),
    ]:
        val(sheet, row, _C_AE, label)
    # Zero_Predictors_Selected() branch computes the closed-form single-mean
    # prediction interval instead of feeding a fabricated first-predictor
    # input into Prediction_Interval(); NA() when there is nothing to fit.
    # The live branch takes exactly COLUMNS(X_s()) input rows — the inputs
    # correspond 1:1 to constructed columns, so no Include-filter is needed.
    f(
        sheet,
        3,
        _C_AF,
        "=IF(Zero_Predictors_Selected(),"
        "IF(AND(Allow_Intercept,Intercept_Only_N()>=2),"
        "LET(point,Intercept_Only_Point(),"
        "se_pred,Intercept_Only_S()*SQRT(1+1/Intercept_Only_N()),"
        "t_crit,T.INV.2T(alpha,Intercept_Only_DF()),"
        "VSTACK(point,se_pred,t_crit,point-t_crit*se_pred,point+t_crit*se_pred,1-alpha)),"
        "NA()),"
        f"LET(pred_input,VSTACK($AF$12,"
        f"TAKE($AF${_PRED_INPUT_FIRST_ROW}:$AF${_PRED_INPUT_LAST_ROW},COLUMNS(X_s()))),"
        "Prediction_Interval(X_s(),Response_Column(),pred_input,Allow_Intercept,"
        "Sample_Include(),alpha)))",
    )
    sheet.range(rc(3, _C_AF), rc(8, _C_AF)).number_format = "0.0000"
    border_box(sheet, 1, _C_AE, 8, _C_AF)


def _write_prediction_inputs(sheet: xw.Sheet) -> None:
    """Zone AE10+: one prediction-input row per constructed column."""
    section_heading(sheet, 10, _C_AE, "PREDICTION INPUTS")
    val(sheet, 11, _C_AE, "Predictor")
    val(sheet, 11, _C_AF, "Prediction Value")
    bold_row(sheet, 11, _C_AE, _C_AG)

    # Row 12: intercept (auto-set, still orange to show it's a value)
    val(sheet, 12, _C_AE, "Intercept")
    f(sheet, 12, _C_AF, "=IF(Allow_Intercept,1,0)")
    format_input(sheet, 12, _C_AF)

    # AE13: spill formula — level-qualified names, one per constructed column
    f(sheet, _PRED_INPUT_FIRST_ROW, _C_AE, "=TRANSPOSE(Constructed_Column_Names())")

    # AG13: the Training Mean column — per-column means of the filtered design
    # matrix, computed with a SINGLE X_s() evaluation. X_s() is a full
    # design-matrix construction on every call (Excel does not cache LAMBDA
    # results), so this spill is the one place the means are computed; the
    # orange prefill cells INDEX into it. The earlier design — every prefill
    # cell calling X_s() twice (width guard + column mean) — made the
    # workbook's first full calculation pathological (~20 minutes at the save
    # step). The spill owns column AG downward (the AH gap and residual zone
    # start beyond it), so a wider dataset or spec can never make it collide
    # with another spill.
    # Degrades to "" on an empty model, which the prefill guard reads as a
    # one-row spill holding a blank.
    val(sheet, 11, _C_AG, "Training Mean")
    means_anchor = f"$AG${_PRED_INPUT_FIRST_ROW}"
    f(
        sheet,
        _PRED_INPUT_FIRST_ROW,
        _C_AG,
        (
            "=IFERROR(TRANSPOSE(BYCOL(FILTER(X_s(),Sample_Include()),"
            'LAMBDA(c,AVERAGE(c)))),"")'
        ),
    )

    # AF13:AF62 — the Training Mean of each constructed column, individually
    # overridable. Each row guards on its position against the means-spill
    # height so rows beyond the live constructed width render blank (the
    # width is spec-dependent — 19 on the default WHO spec). Cheap spill
    # references only: no prefill cell may call X_s() itself.
    offset = _PRED_INPUT_FIRST_ROW - 1
    for row in range(_PRED_INPUT_FIRST_ROW, _PRED_INPUT_LAST_ROW + 1):
        f(
            sheet,
            row,
            _C_AF,
            (
                f"=IF(ROW()-{offset}<=IFERROR(ROWS({means_anchor}#),0),"
                f"INDEX({means_anchor}#,ROW()-{offset}),"
                '"")'
            ),
        )

    # Orange for all user-editable prediction value cells (the Training Mean
    # column is computed display, not input)
    _input_range(sheet, _PRED_INPUT_FIRST_ROW, _C_AF, _PRED_INPUT_LAST_ROW, _C_AF)
    sheet.range(
        rc(_PRED_INPUT_FIRST_ROW, _C_AF), rc(_PRED_INPUT_LAST_ROW, _C_AF)
    ).number_format = "0.0000"
    # Column-wide: the means spill height is spec-dependent and the column
    # holds nothing else.
    sheet.range(f"{col_letter(_C_AG)}:{col_letter(_C_AG)}").number_format = "0.0000"


def _write_residuals(sheet: xw.Sheet) -> None:
    """Residual diagnostic table — row identifiers + 11 diagnostics columns starting at _C_AI."""
    section_heading(sheet, 1, _C_AI, "RESIDUAL OUTPUT")

    # AI2: static header — Row_Labels() supplies its own per-row content
    # (joined Identifier columns, or positional Obs. n labels).
    val(sheet, 2, _C_AI, "Observation")

    for col, header in zip(
        [_C_AJ, _C_AK, _C_AL, _C_AM, _C_AN, _C_AO, _C_AP, _C_AQ, _C_AR, _C_AS, _C_AT],
        [
            "Y", "Predicted Y", "Residuals", "LOOCV Residual",
            "Hat Diagonal", "Studentized Residuals", "Cook's Distance",
            "Normal Scores Ranked", "Studentized Residuals Ranked",
            "Scale-Location", "PRESS Residual",
        ],
    ):
        val(sheet, 2, col, header)
    bold_row(sheet, 2, _C_AI, _C_AT)

    # AI3: row labels — the spec-derived Row_Labels() filtered to the sample.
    # Row_Labels() has its own no-Identifier fallback ("Obs. n"), so the only
    # error left to absorb is an all-FALSE mask.
    f(
        sheet, 3, _C_AI,
        "=IFERROR(FILTER(Row_Labels(),Sample_Include()),NA())",
    )
    # Spill anchors — each spills n rows downward
    f(sheet, 3, _C_AJ, "=Dependent_Variable(Response_Column(),Sample_Include())")
    f(sheet, 3, _C_AK, "=Predictions(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_AN, "=Hat_Diagonal(X_s(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_AL, "=Residuals(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_AM, "=LOOCV_Residual(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_AO, "=Studentized_Residuals(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_AP, "=Cooks_Distance(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    f(sheet, 3, _C_AQ, "=SORT(Normal_Scores(Response_Column(),Sample_Include()))")
    f(sheet, 3, _C_AR, "=Studentized_Residuals_Ranked(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    # Scale-Location: SQRT(|Studentized_Residuals|) — horizontal spread should be flat.
    f(
        sheet, 3, _C_AS,
        "=SQRT(ABS(Studentized_Residuals(X_s(),Response_Column(),Allow_Intercept,Sample_Include())))",
    )
    # PRESS Residual equals the leave-one-out residual e_i / (1 - h_i).
    f(sheet, 3, _C_AT, "=LOOCV_Residual(X_s(),Response_Column(),Allow_Intercept,Sample_Include())")
    # Format every numeric residual-output column — the actual Y (AJ) through
    # PRESS (AT). Only the AI identifier column (text: country/Obs. labels) is
    # left unformatted.
    sheet.range(f"{col_letter(_C_AJ)}:{col_letter(_C_AT)}").number_format = "0.0000"


def _write_diagnostic_charts(sheet: xw.Sheet) -> None:  # pylint: disable=too-many-locals,too-many-statements
    """Create 7 pre-built diagnostic charts to the right of the Residual Output section."""
    start_left = sheet.range(a1(1, _C_AT + 1)).left
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

    chart_specs = [
        (
            "Residuals vs. Fitted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartResid"),
            "Fitted Values", "Residuals", 1, 1,
        ),
        (
            "Normal Q-Q", "scatter",
            _name_ref("RegChartQQX"),
            _name_ref("RegChartQQY"),
            "Theoretical Quantiles", "Studentized Residuals", 1, 2,
        ),
        (
            "Actual vs. Predicted", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartActY"),
            "Predicted Y", "Actual Y", 2, 1,
        ),
        (
            "Scale-Location", "scatter",
            _name_ref("RegChartFitY"),
            _name_ref("RegChartScaleLoc"),
            "Fitted Values", "√|Studentized Residual|", 2, 2,
        ),
        (
            "Cook's Distance", "bar",
            None,
            _name_ref("RegChartCookDist"),
            "Observation", "Cook's Distance", 3, 1,
        ),
        (
            "Studentized Residuals vs. Leverage", "scatter",
            _name_ref("RegChartLeverage"),
            _name_ref("RegChartStudResid"),
            "Leverage (Hat Diagonal)", "Studentized Residuals", 3, 2,
        ),
        (
            "PRESS Residuals", "bar",
            None,
            _name_ref("RegChartPRESSResid"),
            "Observation", "PRESS Residual", 4, 1,
        ),
    ]

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

    for title, chart_type, x_addr, y_addr, x_label, y_label, grid_row, grid_col in chart_specs:
        left, top = _pos(grid_row, grid_col)
        co = sheet.api.ChartObjects().Add(left, top, _CHART_WIDTH, _CHART_HEIGHT)
        chart = co.Chart

        chart.ChartType = _XL_XY_SCATTER if chart_type == "scatter" else _XL_COLUMN_CLUSTERED

        sc = chart.SeriesCollection()
        for i in range(sc.Count, 0, -1):
            sc.Item(i).Delete()

        series = chart.SeriesCollection().NewSeries()
        if x_addr is not None:
            series.XValues = x_addr
        series.Values = y_addr
        series.Name = title
        # Bar charts (Cook's Distance, PRESS Residuals) have no markers to resize.
        if chart_type == "scatter":
            series.MarkerSize = 4

        # All charts: Header-style title (bold, 14 pt, light-blue fill).
        chart.HasLegend = False
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        chart.ChartTitle.Font.Bold = True
        chart.ChartTitle.Font.Size = 14
        chart.ChartTitle.Format.Fill.Visible = True
        chart.ChartTitle.Format.Fill.Solid()
        chart.ChartTitle.Format.Fill.ForeColor.RGB = excel_color(_HEADER)

        x_axis = chart.Axes(_XL_CATEGORY)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = x_label
        x_axis.TickLabels.NumberFormat = "0"

        y_axis = chart.Axes(_XL_VALUE)
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_label
        y_axis.TickLabels.NumberFormat = "0"

        gridline_mode = gridline_modes.get(title, "none")
        x_axis.HasMajorGridlines = gridline_mode == "both"
        y_axis.HasMajorGridlines = gridline_mode in {"y", "both"}

        if title == "Cook's Distance":
            y_axis.TickLabels.NumberFormat = "0.0E+00"
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone
        if title == "Studentized Residuals vs. Leverage":
            x_axis.TickLabels.NumberFormat = "0.00"
        if title == "PRESS Residuals":
            x_axis.TickLabelPosition = -4142  # xlTickLabelPositionNone

        if title == "Normal Q-Q":
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "RegChartQQX", "RegChartQQY")
            if x_addr is None:
                raise AssertionError("Normal Q-Q chart requires an x-axis range")
            identity_ref: str = x_addr
            _add_identity_line(chart, identity_ref)
        if title == "Actual vs. Predicted":
            _set_equal_axis_scale_from_named_ranges(x_axis, y_axis, "RegChartFitY", "RegChartActY")
            if x_addr is None:
                raise AssertionError("Actual vs. Predicted chart requires an x-axis range")
            identity_ref = x_addr
            _add_identity_line(chart, identity_ref)


# ── Public entry point ────────────────────────────────────────────────────────

def write_regression_output_sheet(
    workbook: xw.Book,
    sheet_notes: dict[str, str] | None = None,
    closures: tuple[CatalogFunction, ...] | None = None,
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
    sheet.activate()

    _setup_local_names(sheet, closures)

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

    sheet.range(rc(2, _C_N), rc(2, _C_AU)).api.WrapText = True

    # A–L (spec block) widths are owned by write_sheet_model_construction.py
    # so the standalone and shared-block builds can never drift.
    _set_spec_block_column_widths(sheet)

    # Content-column widths, per zone, plus the AU post-zone gutter (last entry).
    # The gap columns (M, U, AD, AH) are sized from _GAP_COLUMNS below so the
    # layout stays declarative — one width there, not one per hard-coded gap letter.
    for column_letter, width in {
        # Predictor Summary (N–T): level-qualified constructed names + 6 stats.
        "N": 24,        # constructed column names (e.g., "Status[Developed]")
        "O": 9, "P": 9, "Q": 9, "R": 9, "S": 9, "T": 9,  # stats values

        # Regression Outputs (V–AC): regression statistics (V–W), diagnostics (Y–Z), ANOVA (V–AA), coefficients (V–AC), beta weights (AC).
        # Column roles vary by sub-table; widths below are sized for the widest label/value used in each column.
        "V": 22,        # labels — longest is "Adjusted R Square" (16) or "Status[Developing]" (17)
        "W": 12,        # stat values, Alpha, df, Coefficients
        "X": 12,        # SS, Std Error
        "Y": 24,        # diagnostics labels (longest: "BFN Panel Durbin-Watson" = 23) + MS + t Stat
        "Z": 16,        # diagnostics values + "Predicted Variable" section heading + F + P-value
        "AA": 14,       # derived response name (e.g., "Life expectancy") + Significance F + Lower 95%
        "AB": 10,       # Upper 95% values
        "AC": 10,       # Beta Weight values

        # Prediction Outputs (AE–AG): interval box label/values, inputs, training mean.
        "AE": 24,       # section heading, "PREDICTION INTERVAL" / "PREDICTION INPUTS" labels, spilled constructed names
        "AF": 16,       # prediction interval values + prediction input values
        "AG": 14,       # Training Mean header + values spill

        # Residual Output (AI–AT): row identifiers (AI) + 11 diagnostics (AJ–AT).
        # AI holds Row_Labels() — country/identifier strings like "United States" (13).
        "AI": 16,       # row identifiers (Row_Labels)
        "AJ": 10, "AK": 10, "AL": 9, "AM": 10, "AN": 9, "AO": 9, "AP": 12, "AQ": 9,
        "AR": 14, "AS": 17, "AT": 14,

        # AU is NOT a content column and NOT a zone gap — it is the post-zone
        # gutter that bounds the row-2 header wrap (_C_AU) and anchors the
        # diagnostic charts (they start at _C_AT + 1). Sized here so it reads as
        # a deliberate margin rather than a default-width column.
        "AU": 15,
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

    # Charts must be positioned after column widths are set so that
    # sheet.range("AU1").left reflects the final column layout.
    _write_diagnostic_charts(sheet)

    # Freeze top 2 rows
    sheet.activate()
    sheet.range("A3").select()
    win = sheet.api.Application.ActiveWindow
    win.FreezePanes = False
    win.SplitRow = 2
    win.SplitColumn = 0
    win.FreezePanes = True

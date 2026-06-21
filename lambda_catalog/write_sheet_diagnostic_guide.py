"""Write the Diagnostic Guide sheet into the target workbook."""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR, SUBHDR_COLOR as _SUBHEADER_COLOR
from .workbook_helpers import get_or_create_sheet, reset_generated_sheet


SHEET_NAME = "Diagnostic Guide"
_COL_A_WIDTH = 28
_COL_B_WIDTH = 22
_COL_C_WIDTH = 22
_COL_D_WIDTH = 46


def _heading(sheet: xw.Sheet, row: int, text: str) -> None:
    cell = sheet.range((row, 1))
    cell.value = text
    cell.api.Font.Bold = True
    sheet.range((row, 1), (row, 4)).color = _HEADER_COLOR


def _subheading(sheet: xw.Sheet, row: int, text: str, cols: int = 4) -> None:
    cell = sheet.range((row, 1))
    cell.value = text
    cell.api.Font.Bold = True
    sheet.range((row, 1), (row, cols)).color = _SUBHEADER_COLOR


def _table_header_row(sheet: xw.Sheet, row: int, headers: list[str]) -> None:
    for col, text in enumerate(headers, start=1):
        cell = sheet.range((row, col))
        cell.value = text
        cell.api.Font.Bold = True
        cell.color = _SUBHEADER_COLOR


def _row(sheet: xw.Sheet, row: int, values: list[str]) -> None:
    for col, text in enumerate(values, start=1):
        cell = sheet.range((row, col))
        cell.value = text
        cell.api.WrapText = True


def write_diagnostic_guide_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Diagnostic Guide reference sheet."""
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    sheet.range("A:A").column_width = _COL_A_WIDTH
    sheet.range("B:B").column_width = _COL_B_WIDTH
    sheet.range("C:C").column_width = _COL_C_WIDTH
    sheet.range("D:D").column_width = _COL_D_WIDTH

    r = 1

    # ── Title ──────────────────────────────────────────────────────────────────
    _heading(sheet, r, "REGRESSION DIAGNOSTIC GUIDE"); r += 1
    sheet.range((r, 1)).value = (
        "Use the charts and flagged cells on the Regression sheet to assess model assumptions. "
        "Work through Tier 1 first; investigate Tier 2 only when a Tier 1 plot raises a concern."
    )
    sheet.range((r, 1)).api.WrapText = True
    r += 2

    # ── Tier 1: Core diagnostic plots ─────────────────────────────────────────
    _subheading(sheet, r, "TIER 1 — Review for every model"); r += 1
    _table_header_row(sheet, r, ["Plot", "X-axis", "Y-axis", "What to look for"]); r += 1

    tier1 = [
        [
            "Residuals vs. Fitted",
            "Predicted Y",
            "Residuals",
            "Random scatter around zero. A curve or funnel shape signals nonlinearity or "
            "heteroscedasticity (non-constant error variance).",
        ],
        [
            "Normal Q-Q",
            "Normal Scores (theoretical)",
            "Studentized Residuals Ranked",
            "Points close to a straight diagonal line. Heavy tails or an S-curve indicate "
            "non-normal errors. Check QQ Correlation (Cell P10): below 0.98 = mild concern, "
            "below 0.95 = stronger concern.",
        ],
        [
            "Actual vs. Predicted",
            "Predicted Y",
            "Y",
            "Points close to a 45° line. A bow or fan shape confirms the same problems "
            "seen in Residuals vs. Fitted but from a different angle.",
        ],
    ]
    for vals in tier1:
        _row(sheet, r, vals); r += 1

    r += 1

    # ── Tier 2: Follow-up plots ────────────────────────────────────────────────
    _subheading(sheet, r, "TIER 2 — Investigate when Tier 1 raises a concern"); r += 1
    _table_header_row(sheet, r, ["Plot", "X-axis", "Y-axis", "What to look for"]); r += 1

    tier2 = [
        [
            "Scale-Location",
            "Predicted Y",
            "Scale-Location\n(√|Studentized Residuals|)",
            "Flat horizontal spread of points = homoscedasticity. An upward trend "
            "confirms heteroscedasticity flagged in Tier 1. Yellow cells: value > √2 ≈ 1.41; "
            "red cells: value > √3 ≈ 1.73.",
        ],
        [
            "Cook's Distance",
            "Observation (bar position)",
            "Cook's Distance",
            "Spikes above 4/n (yellow) or above 0.9 (red) mark observations with "
            "outsized influence on the fitted coefficients. Bars are ordered by "
            "observation number. Inspect those rows for data entry errors or genuine "
            "outliers before removing them. Remove outliers only when you have a "
            "definitive, non-statistical reason to believe the data point is invalid, "
            "comes from a different population, or disproportionately distorts the "
            "analysis. Never filter out data solely because it is an extreme statistical "
            "value, as doing so can introduce bias and erase genuine insights.",
        ],
        [
            "Studentized Residuals vs. Leverage",
            "Hat Diagonal (leverage)",
            "Studentized Residuals",
            "High leverage alone is not a problem. Combined high leverage and large "
            "residual (top-right or bottom-right of the plot) = influential outlier. "
            "Hat > 2p/n is flagged red; Hat > 3p/n is additionally bold.",
        ],
        [
            "PRESS Residuals",
            "Observation (bar position)",
            "PRESS Residual\n(e / (1 − h))",
            "PRESS residuals inflate the ordinary residual by leverage. Bars are ordered "
            "by observation number. Large values (|PRESS| > 2 × SE yellow, > 3 × SE red) "
            "flag observations whose removal would substantially shift the fitted model.",
        ],
    ]
    for vals in tier2:
        _row(sheet, r, vals); r += 1

    r += 1

    # ── Threshold reference table ──────────────────────────────────────────────
    _subheading(sheet, r, "DIAGNOSTIC THRESHOLD REFERENCE"); r += 1
    _table_header_row(
        sheet, r,
        ["Diagnostic", "Location on sheet", "Yellow threshold", "Red threshold"]
    ); r += 1

    thresholds = [
        ["VIF (Variance Inflation Factor)", "Col I, Predictor Summary",
         "VIF > 5  (possible collinearity)", "VIF > 10  (strong collinearity)"],
        ["Tolerance", "Col J, Predictor Summary",
         "Tolerance < 0.2", "Tolerance < 0.1"],
        ["PRESS R²", "Cell P5, Diagnostics",
         "—", "PRESS R² < 0  (worse than predicting Y-mean)"],
        ["QQ Correlation", "Cell P10, Diagnostics",
         "< 0.98  (mild non-normality)", "< 0.95  (clear non-normality)"],
        ["Significance F", "Cell Q15, ANOVA Table",
         "—", "P-value > alpha  (model not significant)"],
        ["Coefficient P-values", "Col P, Coefficients",
         "—", "P-value > alpha  (term not significant)"],
        ["Hat Diagonal (leverage)", "Col AB, Residual Output",
         "—", "h > 2p/n  (high leverage)"],
        ["Studentized Residuals", "Col AC, Residual Output",
         "|r*| > 2  (moderate outlier)", "|r*| ≥ 3  (strong outlier)"],
        ["Cook's Distance", "Col AD, Residual Output",
         "D > 4/n  (review)", "D > 0.9  (high influence)"],
        ["Scale-Location", "Col AG, Residual Output",
         "> √2 ≈ 1.41  (|r*| > 2 equivalent)", "> √3 ≈ 1.73  (|r*| > 3 equivalent)"],
        ["PRESS Residual", "Col AH, Residual Output",
         "|PRESS| > 2 × SE", "|PRESS| > 3 × SE"],
    ]
    for vals in thresholds:
        _row(sheet, r, vals); r += 1

    r += 1

    # ── Interpretation guidance ────────────────────────────────────────────────
    _subheading(sheet, r, "COMMON PATTERNS AND NEXT STEPS", cols=3); r += 1

    guidance = [
        ["Heteroscedasticity\n(funnel residuals)",
         "Residuals vs. Fitted shows fan shape; Scale-Location trends upward.",
         "Consider: log-transforming Y, adding polynomial terms, or fitting a "
         "Weighted Least Squares (WLS) model. WLS is planned for a future version "
         "of this library."],
        ["Nonlinearity\n(curved residuals)",
         "Residuals vs. Fitted shows a curve; Actual vs. Predicted bows.",
         "Add squared or interaction terms to the model. Toggle individual predictors "
         "on/off to isolate which variable is driving the curve."],
        ["Non-normal errors\n(Q-Q deviation)",
         "Q-Q plot shows heavy tails or S-curve; QQ Correlation < 0.98.",
         "With n > 100 moderate departures rarely invalidate inference. For small n, "
         "consider robust standard errors or a bootstrap Confidence Interval approach."],
        ["Influential observations\n(high Cook's D or PRESS)",
         "Cook's D > 4/n or |PRESS| > 2 × SE for one or more rows.",
         "Inspect those rows. Verify data entry. Refit without the observation(s) and "
         "compare coefficients — if they shift substantially, see which ones and "
         "research the reasons why those data points are different. If there's a "
         "significant difference, choose one model, but report the prediction results "
         "of the other regression as a sensitivity analysis."],
        ["Multicollinearity\n(high VIF)",
         "VIF > 10 for one or more predictors.",
         "Remove or combine correlated predictors. Use the Correlation_Matrix "
         "function to identify the pairs. Partial R² shows each predictor's "
         "unique contribution after controlling for the others."],
    ]
    _table_header_row(sheet, r, ["Pattern", "Symptom", "Next step"]); r += 1
    for vals in guidance:
        _row(sheet, r, vals); r += 1

    sheet.autofit("rows")

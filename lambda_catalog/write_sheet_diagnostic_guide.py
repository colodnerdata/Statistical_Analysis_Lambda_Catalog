"""Write the Diagnostic Guide sheet into the target workbook.

This sheet is a fixed, dataset-independent reference, so artifact
builds just copy the already-styled sheet out of ``TEMPLATE_PATH`` (see
``copy_static_sheet``) instead of re-running the row-by-row COM writes on
every build. ``_write_template_sheet`` remains the authored source of the
content; run this module's CLI after editing it to regenerate the template,
then commit the updated ``.xlsx``.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR
from .sheet_styles import SUBHDR_COLOR as _SUBHEADER_COLOR
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    ColumnSpec,
    copy_static_sheet,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
    set_column_widths,
)

SHEET_NAME = "Diagnostic Guide"

_ROOT_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = _ROOT_DIR / "templates" / "static_sheets.xlsx"

_COLUMNS: tuple[ColumnSpec, ...] = (
    ColumnSpec(1, 28, "Plot / Diagnostic / Pattern"),
    ColumnSpec(2, 22, "X-axis / Location / Symptom"),
    ColumnSpec(3, 22, "Y-axis / Yellow threshold / Next step"),
    ColumnSpec(4, 46, "What to look for / Red threshold"),
)
_LAST_COL = _COLUMNS[-1].index


def _heading(sheet: xw.Sheet, row: int, text: str) -> None:
    cell = sheet.range((row, 1))
    cell.value = text
    cell.api.Font.Bold = True
    sheet.range((row, 1), (row, _LAST_COL)).color = _HEADER_COLOR


def _subheading(sheet: xw.Sheet, row: int, text: str, cols: int = _LAST_COL) -> None:
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


def _write_template_sheet(workbook: xw.Book) -> None:
    """(Re)build this sheet's content directly.

    Used only to author ``TEMPLATE_PATH`` — artifact builds never call
    this; they call ``write_diagnostic_guide_sheet``, which copies the sheet
    this function last wrote into the template.
    """
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    set_column_widths(sheet, ((c.index, c.width) for c in _COLUMNS))

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
        ["GVIF (Generalized Variance Inflation Factor)", "Col U, Predictor Summary",
         "GVIF > 5  (possible collinearity)", "GVIF > 10  (strong collinearity)"],
        ["Tolerance", "Col V, Predictor Summary",
         "Tolerance < 0.2", "Tolerance < 0.1"],
        ["PRESS R²", "Cell P5, Diagnostics",
         "—", "PRESS R² < 0  (worse than predicting Y-mean)"],
        ["QQ Correlation", "Cell P10, Diagnostics",
         "< 0.98  (mild non-normality)", "< 0.95  (clear non-normality)"],
        ["Significance F", "Cell Q15, ANOVA Table",
         "—", "P-value > alpha  (model not significant)"],
        # Column letters are the CURRENT layout. They had drifted a whole zone
        # out of date (AB/AC/AD/AG/AH) from before the Residual Output zone
        # moved to AN:BA, which is easy to miss because this sheet is baked
        # into templates/static_sheets.xlsx and no build re-derives it.
        ["Coefficient P-values", "Col AE, Coefficients",
         "—", "P-value > alpha  (term not significant)"],
        ["Hat Diagonal (leverage)", "Col AR, Residual Output",
         "—", "h > 2p/n  (high leverage)"],
        ["Studentized Residuals", "Col AS, Residual Output",
         "|r*| > 2  (moderate outlier)", "|r*| ≥ 3  (strong outlier)"],
        # One tier, not two. F(0.5, p, n-p) is the median of the reference F
        # distribution, so it tracks the model's own dimensionality instead of
        # just its row count — p is the design matrix's column width
        # (intercept included) and n-p the ANOVA residual df.
        ["Cook's Distance", "Col AT, Residual Output",
         "—", "D > F.INV(0.5, p, n-p)  (high influence)"],
        ["Scale-Location", "Col AW, Residual Output",
         "> √2 ≈ 1.41  (|r*| > 2 equivalent)", "> √3 ≈ 1.73  (|r*| > 3 equivalent)"],
        ["PRESS Residual", "Col AX, Residual Output",
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
        ["Multicollinearity\n(high GVIF)",
         "GVIF > 10 for one or more predictors. A categorical predictor's dummy "
         "columns all show the same shared GVIF value — that's the whole "
         "variable's collinearity, not a per-level artifact.",
         "Remove or combine correlated predictors. Use the Correlation_Matrix "
         "function to identify the pairs. Partial R² shows each predictor's "
         "unique contribution after controlling for the others."],
    ]
    _table_header_row(sheet, r, ["Pattern", "Symptom", "Next step"]); r += 1
    for vals in guidance:
        _row(sheet, r, vals); r += 1

    sheet.autofit("rows")


def write_diagnostic_guide_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Diagnostic Guide reference sheet from the static template.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to receive the sheet.
    """
    copy_static_sheet(workbook, TEMPLATE_PATH, SHEET_NAME)


def _main() -> None:
    """Regenerate this sheet inside ``templates/static_sheets.xlsx``.

    Run after editing the content above, then commit the updated template file.
    """
    parser = argparse.ArgumentParser(
        description=(
            f"Rebuild the {SHEET_NAME!r} sheet inside the static template workbook."
        )
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=TEMPLATE_PATH,
        help="Path to the static template workbook (default: templates/static_sheets.xlsx).",
    )
    args = parser.parse_args()

    template_path = args.template.resolve()
    try:
        with xw.App(visible=True, add_book=False) as app:
            workbook, _ = open_or_create_workbook(app, template_path)
            try:
                _write_template_sheet(workbook)
                workbook.save(str(template_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(template_path, "open or save", exc)

    print(f"Template sheet updated: {SHEET_NAME}")
    print(f"Template workbook: {template_path}")


if __name__ == "__main__":
    _main()

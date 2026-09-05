"""Write the Modeling Concepts sheet into the target workbook.

This sheet is a fixed, dataset-independent reference, so artifact
builds just copy the already-styled sheet out of ``TEMPLATE_PATH`` (see
``copy_static_sheet``) instead of re-running the row-by-row COM writes on
every build. ``_write_template_sheet`` remains the authored source of the
content; run ``scripts/rebuild_static_sheets.py`` after editing it to
regenerate the template, then commit the updated ``.xlsx``.

Where the sibling sheets sit: ``Regression Instructions`` is the
operational *how* (point the sheet at data, which column does what) and
``Diagnostic Guide`` is the *what to look for* in the residual charts.
This sheet is the conceptual layer between them: for each modeling
feature — why it exists, the statistical method it enables, and a
concrete use case.
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

SHEET_NAME = "Modeling Concepts"

_ROOT_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = _ROOT_DIR / "templates" / "static_sheets.xlsx"

_COLUMNS: tuple[ColumnSpec, ...] = (
    ColumnSpec(1, 24, "Feature"),
    ColumnSpec(2, 32, "What Problem It Solves"),
    ColumnSpec(3, 36, "Statistical Method"),
    ColumnSpec(4, 46, "Use Case"),
)
_LAST_COL = _COLUMNS[-1].index


def _heading(sheet: xw.Sheet, row: int, text: str) -> None:
    cell = sheet.range((row, 1))
    cell.value = text
    cell.api.Font.Bold = True
    sheet.range((row, 1), (row, _LAST_COL)).color = _HEADER_COLOR


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
    this; they call ``write_modeling_concepts_sheet``, which copies the
    sheet this function last wrote into the template.
    """
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    set_column_widths(sheet, ((c.index, c.width) for c in _COLUMNS))

    r = 1

    # ── Title ──────────────────────────────────────────────────────────────────
    _heading(sheet, r, "MODELING CONCEPTS — THE WHY BEHIND EACH FEATURE"); r += 1
    sheet.range((r, 1)).value = (
        "Each feature on the Regression sheet exists to make a specific statistical "
        "method available without formulas or add-ins. This table explains that link: "
        "what problem the feature solves, the method it enables, and a concrete "
        "situation for using it. For the mechanics of setting a feature up, see the "
        "Regression Instructions sheet; for how to read the fitted model's residual "
        "plots, see the Diagnostic Guide."
    )
    sheet.range((r, 1)).api.WrapText = True
    r += 2

    # ── Feature table ──────────────────────────────────────────────────────────
    _table_header_row(
        sheet, r, ["Feature", "What Problem It Solves", "Statistical Method", "Use Case"]
    )
    r += 1

    features = [
        [
            "Intercept Control\n(C2 toggle)",
            "The intercept is the model's baseline — the expected response at "
            "zero on every predictor. Usually that baseline is meaningful and "
            "should stay; occasionally theory says the response must be zero "
            "when the predictors are.",
            "A model-level on/off toggle (cell C2). With reference-coded "
            "categoricals in the model, turning it OFF is flagged red: each "
            "dummy coefficient redefines from a contrast against the reference "
            "to an absolute level mean, which is usually not the question being "
            "asked. Under Fixed Effects it is flagged because the within "
            "transformation demeans the data, and the intercept of demeaned "
            "data is an uninterpretable artifact.",
            "Theory says cost is zero at zero production, or the instrument "
            "reads zero with no input? Turn the intercept off for a "
            "regression-through-origin fit. Otherwise leave it on and let it "
            "absorb the reference levels' baseline.",
        ],
        [
            "Reference Levels\n(Categorical predictors)",
            "A categorical column cannot enter a regression as text — it has to "
            "become numbers. But one 0/1 column per level PLUS an intercept is "
            "redundant (perfect multicollinearity), so one level must step aside "
            "as the baseline everything else is measured against.",
            "Treatment (dummy) coding: a 0/1 column per retained level, with the "
            "reference level dropped. Each categorical coefficient reads as a "
            "contrast against the reference; the intercept absorbs the reference "
            "level's baseline. The default reference is the first-in-sort-order "
            "level; type another in column E to override. A typed level not "
            "present in the sample is flagged red, never silently fitted.",
            "Origin (US / Europe / Japan) predicting MPG: with US as the "
            "reference, the Europe coefficient is the Europe-vs-US difference in "
            "MPG. Pick the level that makes the contrasts meaningful — an "
            "experiment's control group, or the market standard your question is "
            "'compared to what?'",
        ],
        [
            "Log Transforms\n(Transform = Log / Log (drop ≤ 0))",
            "Many relationships are multiplicative, not additive — effects are "
            "percentages, not absolute amounts. Fitting in log space makes "
            "percentages linear, tames skewed data, and turns elasticity models "
            "into simple slopes.",
            "The fit runs on Ln(y) and/or Ln(x): every statistic (coefficients, "
            "R², residuals, prediction intervals) is in log space. The Unit-Space "
            "Fit block back-transforms predictions and fit statistics to the "
            "original units with Duan smearing — a retransformation correction "
            "for the bias of exp(Ln ỹ) under non-constant variance — next to the "
            "naive exp() form so the two can be compared. Zeros and negatives "
            "have no logarithm: plain Log keeps them in the sample and the fit "
            "fails loudly (#N/A, cell turns red); Log (drop ≤ 0) excludes them "
            "and reports how many.",
            "Price vs. size: a Log response with Log predictors gives an "
            "elasticity — each coefficient is the % change in price for a 1% "
            "change in its predictor. For a skewed response like income, "
            "Log(y) makes the residuals symmetric and the Q-Q plot behave.",
        ],
        [
            "Sample Filtering & Completeness\n(Role = Filter)",
            "The fit should answer a question about a specific population, and "
            "rows with missing values cannot be part of it. The workbook derives "
            "the analysis sample from the spec, so a change of population is a "
            "specification change — not hand-deleting rows on the data sheet.",
            "A per-row mask ANDed together: every Filter column must be TRUE, "
            "the Response must be numeric, and every included Continuous "
            "predictor must be numeric (listwise deletion). Categorical "
            "predictors impose no completeness condition — a blank category is "
            "just not a level. Log (drop ≤ 0) adds its own exclusion layer, and "
            "the excluded-row counts surface in the status cells above the spec "
            "(B2 / G2) rather than staying silent.",
            "Restrict a nationwide demand model to one market segment: add a "
            "derived column that is TRUE only for that segment, set its Role to "
            "Filter, and every statistic is refit to that population. Retarget "
            "the filter later and the sample follows — the data sheet is never "
            "touched.",
        ],
        [
            "Interactions\n(Interaction Term / Operation, M/N)",
            "One predictor's effect can depend on another's value — a discount "
            "that lands differently by market, a drug that scales with dosage. A "
            "purely additive model cannot express that dependency; an "
            "interaction column can.",
            "Pairwise constructed columns in three operations: Product "
            "(symmetric), Difference (antisymmetric), Ratio (asymmetric). Width "
            "follows the operands: Continuous × Continuous adds 1 column, "
            "Continuous × Categorical adds L−1, Categorical × Categorical adds "
            "(L₁−1)(L₂−1) — all counted in the Design Columns audit. Interacting "
            "a variable with itself under Product is the documented quadratic. "
            "An interaction whose main effect is switched off is flagged amber "
            "(marginality); declaring both A×B and B×A under Product or "
            "Difference is flagged red (duplicate column, singular matrix).",
            "Does advertising work equally in every region? Interact Ad Spend "
            "(Continuous) with Region (Categorical): the design gains one "
            "ad-sensitivity slope per region vs. the reference, and each "
            "coefficient answers for its own region. For diminishing returns, "
            "interact x with itself under Product to add x².",
        ],
        [
            "Fixed Effects\n(Role = Fixed Effects)",
            "Groups differ in ways you never measured — plant conditions, firm "
            "culture, country policy — and those stable differences can drive both "
            "the response and the predictors, biasing coefficients. Fixed Effects "
            "absorbs a separate intercept for each group so the group differences "
            "stop competing with the predictors you actually care about.",
            "One-way within transformation: every variable is restated as a "
            "deviation from its group's mean, which removes the group intercepts "
            "from the fit entirely. Algebraically equal to LSDV (a dummy column "
            "per group) without the columns, and the absorbed group degrees of "
            "freedom are credited back into inference.",
            "Modelling salaries across 12 plants: plant-level pay practices drive "
            "both salary and who works where. Set Plant to Role = Fixed Effects and "
            "the coefficients answer the within-plant question — holding the plant "
            "fixed, what does this predictor change? — with no plant dummy columns "
            "cluttering the coefficient table.",
        ],
        [
            "Sequence Effects\n(Sequence flag, column H)",
            "Panel and repeated-measures data violate the independence "
            "assumption — rows within one unit follow each other in time. Lags, "
            "differences, and serial-correlation diagnostics all need to know "
            "which variable defines that ordering, which is what the Sequence "
            "flag declares.",
            "Within-group exact-time matching: Lag_By and Difference_By find each "
            "group's prior period by matching its time value, never by row "
            "position, so a gap in the panel yields #N/A instead of a silently "
            "wrong neighbor. The Base Period Δ (your typed override, or the "
            "computed candidate — the most common within-group spacing) sets the "
            "period step, and the BFN Panel Durbin-Watson tests for serial "
            "correlation within groups.",
            "Yearly observations per country: set Year to Sequence = TRUE and a "
            "year-over-year difference column gives each country's change between "
            "its own consecutive years. The Sequence Spacing verdict above the "
            "spec tells you whether the panel is regular enough for those "
            "features to mean what you think they mean.",
        ],
    ]
    for vals in features:
        _row(sheet, r, vals)
        r += 1

    # ── Planned features ───────────────────────────────────────────────────────
    r += 1
    _heading(
        sheet, r, "PLANNED FEATURES — DECLARED IN THE SHEET, NOT YET BUILT"
    )
    r += 1
    sheet.range((r, 1)).value = (
        "The features below are declared in the sheet's structure — a Role value "
        "in the dropdown, a reserved column — but no formula builds them in this "
        "release. Each row states what problem the feature will solve when it "
        "ships, and what to use in the meantime."
    )
    sheet.range((r, 1)).api.WrapText = True
    r += 2

    _table_header_row(
        sheet, r, ["Feature", "What Problem It Solves", "Statistical Method", "Use Case"]
    )
    r += 1

    planned_features = [
        [
            "Weight — WLS\n(Role = Weight, v3.7)\nPLANNED, future release",
            "Some observations are known to be noisier than others — each row "
            "averages a different number of underlying readings, or a variance "
            "you can quantify at row level. OLS gives the noisy rows the same "
            "vote as the precise ones.",
            "Weighted Least Squares: each row enters the fit scaled by its "
            "known variance, so precise observations dominate. The Weight Role "
            "will name the column carrying that variance.",
            "Today's workaround is the Diagnostic Guide's heteroscedasticity "
            "guidance: a Log transform on the response, or an interaction / "
            "quadratic term when the variance tracks a predictor's mean.",
        ],
        [
            "Cluster\n(Role = Cluster, v3.5)\nPLANNED, future release",
            "Rows collected in groups — plants, firms, countries — share "
            "unmodeled shocks, so nominally-independent standard errors are "
            "too optimistic (they overstate significance).",
            "Clustered-robust variance estimator: the same coefficients, with "
            "standard errors that allow arbitrary correlation within a group. "
            "Partially forward-wired — the Role slot and a dormant branch in "
            "Serial_Correlation_Group() await the estimator.",
            "Survey or panel data where rows within one cluster move together: "
            "name the cluster column and the P-values and confidence intervals "
            "stop assuming every row is independent.",
        ],
        [
            "Time\n(Role = Time, v3.6)\nPLANNED, future release",
            "A panel needs an explicit time index distinct from 'which group a "
            "row belongs to' — lag and difference features must follow time "
            "within a unit, not sheet order.",
            "Time-index designation: the Role names the column holding each "
            "row's time value; Lag_By / Difference_By and the future Time "
            "Series sheet are its consumers. The Role ships first, the sheet "
            "last.",
            "Until then, the Sequence flag (column H) plus its Sequence Period "
            "override already carries the ordering axis for the shipped "
            "lag/difference features.",
        ],
        [
            "Two-way Fixed Effects\n(v3.8)\nPLANNED, future release",
            "Sometimes the stable nuisance is two-dimensional — both the plant "
            "AND the year leave a fixed mark on every observation, and "
            "absorbing only one of them leaves the other confounding the "
            "coefficients.",
            "Absorb_Two_Way_Fixed_Effects: the within transformation applied "
            "along two group axes at once. Today at most ONE Fixed Effects row "
            "is legal — two or more is a red spec error, by design, until this "
            "lands.",
            "Plant-and-year panels: for now, keep one Fixed Effects row and "
            "use a Filter to narrow the sample, or dummy-code the second axis "
            "as a Categorical predictor.",
        ],
        [
            "Order\n(column F)\nPLANNED, future release",
            "Term order in the coefficient table currently follows the source "
            "table's column order — fine for every model so far, but a "
            "hand-authored spec may want terms presented in a chosen sequence.",
            "A reserved, hidden column (width 0): it will control each spec "
            "row's position in the constructed design matrix and the "
            "coefficient table. No formula reads it in this release.",
            "Nothing to do today — the column is hidden and ignored; leave it "
            "blank if you reveal it.",
        ],
    ]
    for vals in planned_features:
        _row(sheet, r, vals)
        r += 1

    sheet.autofit("rows")


def write_modeling_concepts_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Modeling Concepts reference sheet from the static template.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to receive the sheet.
    """
    copy_static_sheet(workbook, TEMPLATE_PATH, SHEET_NAME)


def _main() -> None:
    """Regenerate this sheet inside ``templates/static_sheets.xlsx``.

    Prefer ``python scripts/rebuild_static_sheets.py`` — it regenerates
    every static sheet so two edited modules cannot half-regenerate. Run
    after editing the content above, then commit the updated template
    file.
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

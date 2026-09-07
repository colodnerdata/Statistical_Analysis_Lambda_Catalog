"""Write the Version History sheet into the target workbook."""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR
from .sheet_styles import SUBHDR_COLOR as _SUBHDR_COLOR
from .workbook_helpers import (
    ColumnSpec,
    get_or_create_sheet,
    reset_generated_sheet,
    set_column_widths,
)

SHEET_NAME = "Version History"
_TABLE_HEADER_COLOR = _SUBHDR_COLOR

_COLUMNS: tuple[ColumnSpec, ...] = (
    ColumnSpec(1, 12, "Version"),
    ColumnSpec(2, 16, "Release Date"),
    ColumnSpec(3, 14, "Breaking?"),
    ColumnSpec(4, 90, "Summary of Changes"),
)
_LAST_COL = _COLUMNS[-1].index

# Versions follow the interface-based semantic-versioning convention in
# ROADMAP.md: MAJOR only when a workbook built against the prior version would
# stop working or silently compute something different. Under that definition
# the ladder was renumbered on 2026-07-05 — Univariate is 1.1.0, the
# hardening release is 1.2.0, and Specification-Driven Regression is 2.0.0,
# the one MAJOR.
_VERSIONS = [
    {
        "version": "1.0.0",
        "date": "2026-06-16",
        "breaking": "No",
        "summary": (
            "Initial release — complete OLS/MLR engine with full model-fit and ANOVA "
            "statistics, coefficient inference (SE, t, p, CIs, partial R²/correlation), "
            "multicollinearity diagnostics (VIF, Tolerance, correlation matrix), residual "
            "and influence diagnostics (residuals, studentized residuals, hat diagonal, "
            "Cook's distance, Q-Q machinery, Durbin-Watson), cross-validation (PRESS, "
            "LOOCV), information criteria (AIC, AICc, BIC), distributional exploration "
            "(skewness, kurtosis, Pearson/Spearman), and prediction with confidence "
            "intervals. Ships with Regression sheet, Diagnostic Guide, Regression "
            "Instructions, pre-built diagnostic charts, and this Version History."
        ),
    },
    {
        "version": "1.1.0",
        "date": "2026-06-29",
        "breaking": "No",
        "summary": (
            "Univariate Analysis release. Adds a live Univariate sheet with descriptive "
            "statistics, missing-count handling, Sturges/Scott/Freedman-Diaconis "
            "histogram binning, dynamic histogram charts, lower/upper edge and midpoint "
            "helpers, eight CDF and NLL distribution families (Normal, Lognormal, "
            "Exponential, Weibull, Gamma, Triangular, Beta, BetaPERT), goodness-of-fit "
            "ranking with AIC, BIC, Anderson-Darling, and Kolmogorov-Smirnov, and "
            "two-stage native Data Table grid-search fitting for Weibull, Gamma, and "
            "Beta. Also establishes the shared sheet-style palette and the xlwings COM "
            "chart-writing pattern used by generated worksheets."
        ),
    },
    {
        "version": "1.2.0",
        "date": "2026-07-03",
        "breaking": "No",
        "summary": (
            "Workbook hardening and regression usability update. Adds Name "
            "Manager notes to catalog functions, a predicted-variable readout and row "
            "identifiers on the Regression sheet, LOOCV_Residual as an observation-level "
            "diagnostic, stronger intercept-only and undersized-sample guards, explicit "
            "error handling instead of silent first-predictor fallbacks, a chart data "
            "series for the identity line, and safer production-build retry/RPC handling. "
            "Includes categorical helper groundwork and documentation while leaving the "
            "full specification-driven regression sheet for the v2.0 milestone."
        ),
    },
    {
        "version": "2.0.0",
        "date": "2026-07-05",
        "breaking": "Yes",
        "summary": (
            "Specification-Driven Regression release (MAJOR — existing workbook "
            "inputs change meaning). The Regression sheet's control block is "
            "replaced by a declarative variable-specification block (Role, "
            "Include, Type, Reference Level, plus reserved Order/Transform "
            "columns); X_s is promoted from a column filter to a model-matrix "
            "constructor with sheet-scoped twins (Source_Data, Sample_Include, "
            "Response_Column, Row_Labels, X_s, Constructed_Column_Names) that "
            "dissolve the v1 hard-wired ranges. Categorical predictors are "
            "reference-dropped via the rebuilt Dummy_Levels/Dummy_Code, now "
            "NA()-error-signaling with degenerate or invalid-reference cases "
            "flagged red rather than erroring. Ships the canonical rename pass "
            "and a Model Construction QC analyzer that asserts the default-spec "
            "audit values, the X_s/Constructed_Column_Names twin widths, and "
            "the full-height row-mask contract."
        ),
    },
    {
        "version": "2.1.0",
        "date": "2026-08-02",
        "breaking": "No",
        "summary": (
            "Sequence axis, gap-aware longitudinal features, serial-correlation "
            "diagnostics, and one-way Fixed Effects. The specification block "
            "gains a Sequence flag marking the column that ORDERS the data "
            "(at most one) plus a Sequence Period / Period In Use pair — the "
            "period in use is the computed candidate unless you type an "
            "override. Lag_By and Difference_By read that period and are "
            "gap-aware: a missing prior period returns #N/A rather than "
            "silently reaching back to the previous available row. Durbin-"
            "Watson becomes sequence-aware, and a panel form (BFN) appears "
            "when the data is grouped. Fixed Effects joins the Role dropdown: "
            "the named column enters no design-matrix column of its own — the "
            "response and every predictor are demeaned within its groups "
            "instead, and the absorbed degrees of freedom are carried through "
            "every standard error, t-statistic, p-value, confidence interval, "
            "and information criterion, so the inference matches a "
            "least-squares dummy-variable fit. Prediction gains a group-mean-"
            "recovery form reporting both a mean-response confidence interval "
            "and a new-observation prediction interval. Generalized VIF covers "
            "the multi-level categorical case. A model with no Fixed Effects "
            "Role computes exactly as it did under 2.0.0. Third sample "
            "dataset: Production Lots, an unbalanced learning-curve panel with "
            "a natural grouping column."
        ),
    },
    {
        "version": "2.2.0",
        "date": "2026-08-02",
        "breaking": "No",
        "summary": (
            "Transforms — the reserved Transform column (G) gains its first "
            "real value. Set it to Log on the Response row and/or on "
            "Continuous Predictor rows and the whole fit is computed in log "
            "space: coefficients, R-squared, diagnostics, residuals, and the "
            "prediction interval. The constructed column is relabelled "
            "Ln(name) so the output says what was fitted. Log is disallowed "
            "on a Categorical Predictor and flagged red rather than silently "
            "applied. The Prediction Inputs band takes a raw value and logs it "
            "internally, and its Training Mean default is the geometric mean "
            "for a logged column, so the default prediction is not "
            "double-logged. New function Ln_Positive returns #N/A for a "
            "non-positive value instead of an error that would be mistaken "
            "for a formula fault. Predictions are NOT back-transformed and "
            "stay labelled (Log); unit-space comparability and Duan smearing "
            "are a later release. Transform = None (the default) fits the raw "
            "column, identically to 2.1.0."
        ),
    },
    {
        "version": "3.0.0",
        "date": "2026-08-02",
        "breaking": "Yes",
        "summary": (
            "v3.0: bounded model context, constructor pipeline, the "
            "two-workbook split, and the layout break. The Regression fit "
            "chain is rebuilt around a bounded [Context] argument "
            "(Has_Intercept, DF_Absorbed, response and predictor transforms) "
            "threaded through a single Model_Context constructor with "
            "Context_* field accessors, and Design_Columns / Design_Response "
            "construct the fit-time design matrix (the within estimator's "
            "demean-by-group stage lives here, not in the engines). "
            "Univariate Analysis moves to its own workbook "
            "(Lambda_Library_Univariate.xlsx) so each artifact can set its "
            "own calculation mode; the Regression workbook returns to full "
            "Automatic. The model specification block gains three columns: "
            "Interaction Term (M) and Interaction Operation (N) declare an "
            "interaction between a spec row and another Predictor, and "
            "Design Columns (O) shows how many design-matrix columns each "
            "row contributes, totalled above it as the constructed matrix's "
            "full width. A pre-flight width guard reads that total and flags "
            "a model wide enough to be slow (amber) or too wide for the "
            "sheet (red) — computed from the specification, before any "
            "matrix is built. The sheet's far right gains the Constructed "
            "Design Matrix zone, which terminates the materialization band; "
            "nothing may ever be placed to its right. M and N are reserved: "
            "validated and flagged now, read by no constructor until the "
            "interaction wiring release. WHAT BREAKS: cell ADDRESSES, not "
            "meanings. Columns A-L keep both their letters and their "
            "meanings, so a saved specification survives the upgrade "
            "unchanged and no fitted number moves — but the three new "
            "columns push every zone right of the spec block three columns "
            "over (Alpha moves Y12 to AB12, Prediction Inputs move from "
            "column AH to AK, Residual Output from AK to AN), so any "
            "formula of your own that points at a cell on this sheet needs "
            "re-pointing."
        ),
    },
    {
        "version": "3.1.0",
        "date": "2026-08-03",
        "breaking": "No",
        "summary": (
            "Interaction wiring — the Interaction Term (M) and Interaction "
            "Operation (N) columns 3.0.0 reserved now build real design-matrix "
            "columns. Name another Predictor in M and pick Product, "
            "Difference, or Ratio in N, and that spec row contributes its "
            "interaction alongside its own column. Width follows the two "
            "operands: one column for Continuous x Continuous, one per "
            "retained level for Continuous x Categorical, and the full "
            "product for Categorical x Categorical — the Design Columns "
            "audit (O) now counts them, so a single dropdown that adds 155 "
            "columns says so before anything is built. A row pointing at "
            "itself under Product is the documented way to write a quadratic "
            "term. An operand that is a Predictor with Include = FALSE still "
            "builds (flagged amber — an interaction without its main effect); "
            "an operand that is not a Predictor at all contributes nothing "
            "and stays flagged red. An interaction column is named from the "
            "two operands' own headers joined by its operation's own symbol "
            "(Weight \u00d7 Origin: US, Cost \u00f7 Units), so a header says "
            "what was done as well as what to. NOT "
            "AUTOMATIC: an interaction row in Prediction Inputs is an "
            "independent value, not recomputed from its operand rows — "
            "leave the band at its defaults, or override the interaction "
            "rows to match when you change an operand. A specification with "
            "M and N blank (the default) computes exactly as it did under "
            "3.0.0."
        ),
    },
    {
        "version": "3.3.0",
        "date": "2026-08-03",
        "breaking": "No",
        "summary": (
            "Transforms remainder — unit-space fit, Duan back-transformation, "
            "and the model-formula label. The Regression sheet gains a "
            "Unit-Space Fit block at AG3:AH9 with a per-cell Duan / Naive "
            "back-transformation toggle (AH4, default Duan), a smearing "
            "factor (AH5), R² / Adj R² / RMSE in original units (AH6:AH8), "
            "and a response-space readout (AH9). A new Original Units column "
            "(AL) in the Prediction Outputs block carries the back-"
            "transformed point estimate (Duan by default, Naive on toggle) "
            "and the four CI/PI bounds (always Naive — never smeared — "
            "because multiplying a quantile by the smearing factor would "
            "destroy coverage). The Residual Output zone gains two "
            "Original-Units columns (AZ, BA): Predicted Y and Residual in "
            "original units, dispatched on the AH4 toggle. The new Model "
            "Formula cell at AA2:AB2 renders the saved spec as a readable "
            "string (response name with Ln()-wrap on Log, intercept 1 with "
            "the no-intercept flag, the constructed column names — which "
            "already emit Ln(...) per logged predictor, level-qualified "
            "dummies, and left × right interaction names — and a '| FE' "
            "suffix when a Fixed Effects row is present). Three sheet-"
            "scoped named ranges (Comparison_Anchor, Comparison_Headline_GoF, "
            "Comparison_Model_Formula) are the v3.4 Model Comparison sheet's "
            "reading surface. Eight new catalog functions under the "
            "Back-Transformation subcategory (Smearing_Factor, "
            "Back_Transform_Response, Unit_Space_Predictions, "
            "Unit_Space_Observed, Unit_Space_Residuals, Unit_Space_R_Squared, "
            "Unit_Space_Adjusted_R_Squared, Unit_Space_RMSE) each SWITCH "
            "on the (response_transform, predictor_transform) pair read off "
            "Fit_Context() and return NA() outside the six recognised "
            "pairs. Reduction invariant: with Transform = None throughout, "
            "Unit_Space_R_Squared equals R_Squared to fp precision, "
            "Unit_Space_Adjusted_R_Squared equals Adjusted_R_Squared, and "
            "Unit_Space_RMSE equals SE_Regression — the same non-breaking "
            "backward property v2.2 established. The default WHO and "
            "Standard UN Population Practice demo (Transform = None "
            "everywhere) shows no change. WHAT MOVES: nothing in the spec "
            "block. The two new Residual Output columns push the chart "
            "anchor and the entire §4b materialization band right by two "
            "columns (the chart anchor letter BP at column 68 is preserved); "
            "the Regression sheet's residual zone is now AN-BA instead of "
            "AN-AY. Library version moves from 3.1.0 to 3.3.0; Univariate "
            "stays at 2.0.0 — the unit-space machinery is Regression-only."
        ),
    },
    {
        "version": "3.4.0",
        "date": "2026-09-06",
        "breaking": "No",
        "summary": (
            "Unit-space LOOCV — cross-validation brought out of fit space. "
            "v3.3 shipped the unit-space fit (R², Adj R², RMSE in original "
            "units) but left cross-validation in the space the model was "
            "fitted in: LOOCV_Residual, LOOCV_Prediction, PRESS, and the hat "
            "diagonal all reported in log units on a Log-response model. This "
            "release closes that gap. The Regression sheet gains a "
            "CROSS-VALIDATED FIT sub-block at AG11:AH14: LOOCV RMSE (Unit) "
            "and LOOCV MAE (Unit) in original units (rows 12–13), and a "
            "Smearing Treatment readout (row 14) that names how the smearing "
            "factor was obtained. The Residual Output zone gains a fourth "
            "original-units column (BB) carrying the LOOCV residual per "
            "observation. Four new catalog functions under the "
            "Back-Transformation subcategory (Unit_Space_LOOCV_Residual, "
            "Unit_Space_LOOCV_RMSE, Unit_Space_LOOCV_MAE, Smearing_Treatment) "
            "SWITCH on the same (response_transform, predictor_transform) "
            "pair as the v3.3 family and return NA() outside the six "
            "recognised pairs. LOOCV RMSE and MAE use a divisor of n, not "
            "df_residual — every leave-one-out prediction is genuinely "
            "out-of-sample, so no degrees of freedom are consumed and the "
            "pair reads consistently. Under Duan the smearing factor is "
            "estimated on the full sample including the held-out row, a real "
            "but small optimism the Smearing Treatment cell states plainly "
            "('Full-sample Duan (approx.)') rather than hides; fold-specific "
            "Duan is the fourth SWITCH arm the function is structured to "
            "receive. Reduction invariant extended, not broken: with "
            "Transform = None throughout, Unit_Space_LOOCV_Residual equals "
            "LOOCV_Residual and Unit_Space_LOOCV_RMSE equals sqrt(PRESS/n) "
            "to fp precision, including under Fixed Effects. The AG9 in-"
            "sample scalar is relabelled 'SE Regression (Unit)' (formula "
            "unchanged — it divides by df_residual by design, so the label "
            "stops promising ÷ n while LOOCV RMSE one row below delivers "
            "it). New test-model case P08 (Production Lots, Log response, "
            "quadratic self-interaction) supplies the heavy-tailed leverage "
            "corner where LOOCV departs materially from the in-sample fit. "
            "Under Fixed Effects the leave-one-out leverage is the "
            "equivalent LSDV design's, not the within design's: the "
            "within hat diagonal omits the absorbed group effects and "
            "never re-estimates the held-out row's group mean, which "
            "understates the error. The correction is h_lsdv = h - 1/n "
            "+ 1/n_g, exact because the within and group-mean "
            "projections are orthogonal, and it is pinned against "
            "explicit n-fold LSDV refits. A model that absorbs fixed "
            "effects without supplying the group column returns #N/A "
            "rather than the conditional number. "
            "The default WHO demo (Transform = None) shows no change. "
            "Catalog grows from 152 to 157 functions (the fifth is "
            "Group_Size, the per-row group size the FE branch needs). "
            "The planned train is renumbered by one from Model "
            "Comparison on: it moves to v3.5, Cluster Role to v3.6, "
            "and so on through Time Series at v3.13."
        ),
    },
]


def write_version_history_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Version History sheet.

    One workbook, one lineage. The version history records the single
    shipped workbook; the pre-split Univariate releases are narrated in
    ``docs/ROADMAP.md`` § *Univariate releases*. The Univariate history is
    part of this lineage, so nothing in the changelog below is conditional.

    Parameters
    ----------
    workbook : xw.Book
        The target workbook.
    """
    versions = _VERSIONS

    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    set_column_widths(sheet, ((c.index, c.width) for c in _COLUMNS))

    cell = sheet.range("A1")
    cell.value = "VERSION HISTORY"
    cell.api.Font.Bold = True
    sheet.range((1, 1), (1, _LAST_COL)).color = _HEADER_COLOR

    for column in _COLUMNS:
        cell = sheet.range((2, column.index))
        cell.value = column.header
        cell.api.Font.Bold = True
        cell.color = _TABLE_HEADER_COLOR

    for row_offset, entry in enumerate(versions):
        row = 3 + row_offset
        sheet.range((row, 1)).value = entry["version"]
        sheet.range((row, 2)).value = entry["date"]
        sheet.range((row, 3)).value = entry["breaking"]
        cell = sheet.range((row, _LAST_COL))
        cell.value = entry["summary"]
        cell.api.WrapText = True
        sheet.range((row, 1), (row, _LAST_COL)).api.EntireRow.AutoFit()

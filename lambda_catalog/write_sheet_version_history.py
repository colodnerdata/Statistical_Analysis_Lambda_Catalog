"""Write the Version History sheet into the target workbook."""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR, SUBHDR_COLOR as _SUBHDR_COLOR
from .workbook_helpers import ColumnSpec, get_or_create_sheet, reset_generated_sheet, set_column_widths


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
# the ladder was renumbered on 2026-07-05 — Univariate (previously released as
# 2.0.0) is 1.1.0, the hardening release (previously 2.1.0) is 1.2.0, and
# Specification-Driven Regression (previously 3.0.0) is 2.0.0, the one MAJOR.
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
]


# The Univariate artifact's own version lineage. It starts at 1.0.0 with the
# v3.0 split (Univariate Analysis becoming its own workbook) — see DECISIONS.md
# § v3.0 "Univariate becomes its own workbook" and the two-version scheme in
# README.md ("Univariate Workbook 1.0.0 · Function Library 3.0.0"). The
# pre-split Univariate history (1.1.0's "Univariate Analysis release" onward)
# belongs to the Regression workbook's lineage, which carried the sheet until
# 3.0.0; this artifact's history begins at the split.
_UNIVARIATE_VERSIONS = [
    {
        "version": "1.0.0",
        "date": "2026-08-02",
        "breaking": "No",
        "summary": (
            "Initial release of the standalone Univariate workbook. Ships "
            "Univariate Analysis in its own file so its six two-input Data "
            "Tables (Weibull, Gamma, Beta across two stages; 2,400 NLL "
            "evaluations per recalculation) can run in full Automatic "
            "calculation — distribution-fit results are never stale pending a "
            "manual Ctrl+Alt+F9. Carries the complete 126-function LAMBDA "
            "library, the Life Expectancy Data sheet the Univariate data zone "
            "reads, and this Version History. The sheet content is unchanged "
            "from the Univariate Analysis that shipped inside Lambda_Library.xlsx "
            "through 3.0.0; only the packaging (own workbook, own calc mode) is new."
        ),
    },
    {
        "version": "2.0.0",
        "date": "2026-08-03",
        "breaking": "Yes",
        "summary": (
            "The grid shrink, Weibull and Gamma half. Both fits now profile "
            "their scale/rate parameter out in closed form (Weibull "
            "λ(k) = ((1/n)·Σ x^k)^(1/k); Gamma β(α) = α/mean(x)) "
            "and search a 20-point Profile NLL column per stage, replacing "
            "four 20x20 grid blocks — 40 negative-log-likelihood evaluations "
            "per fit instead of 800. Each Stage 1 brackets a closed-form "
            "starting value (Weibull from a probability-plot regression, "
            "Gamma from Minka's approximation) at one third to three times "
            "that value, so the search no longer depends on a guessed range; "
            "the two NLL heatmaps are replaced by profile-NLL line charts. "
            "Profiling is still exact MLE — the profile maximizer is the "
            "joint maximizer — and on the shipped Life Expectancy sample "
            "both fits land closer to the true optimum than the grids they "
            "replace (Gamma by 6.8 NLL, worth about 13.7 of AIC). "
            "The fitting band is also rearranged: each distribution now has "
            "its own column zone with that fit's two stages stacked inside "
            "it, rather than one wide band per stage with the three "
            "distributions stacked across them. A fit reads as a single "
            "self-contained block, the two profile-NLL curves sit side by "
            "side, and the Q-Q plot data moves left of the fits. The band is "
            "26 rows shorter and the sheet ends at column DD instead of DF. "
            "Each profile-NLL chart moves out of the chart band to sit "
            "directly under its own fit zone (BP33, BZ33), one clear row "
            "below the curve it draws, and now plots both stages rather than "
            "just the wide Stage 1 bracket — Stage 2 carries + markers, so "
            "the narrow region the search actually resolved is visible "
            "instead of being a couple of pixels of the Stage 1 line. "
            "BREAKING for this workbook, on two counts. The Scale (λ) and "
            "Rate (β) rows lose their Min, Max, and Step Size input cells — "
            "those parameters are solved, not searched — so saved bounds for "
            "them no longer mean anything. And every cell in the fitting and "
            "Q-Q band changes address, so anything pointing into that band "
            "from outside the sheet needs updating; formulas inside the sheet "
            "and all named ranges were moved with it. The Shape Min/Max cells "
            "remain editable overrides, now formula-defaulted from the start "
            "value. Beta is unchanged apart from its new location: it keeps "
            "the artifact's only two two-input Data Tables. No function "
            "changed; the library version does not move."
        ),
    },
]


def write_version_history_sheet(workbook: xw.Book, *, artifact: str = "regression") -> None:
    """Create or refresh the Version History sheet.

    Parameters
    ----------
    workbook : xw.Book
        The target workbook.
    artifact : str, optional
        Which artifact's version lineage to write: ``"regression"`` (default)
        for the Regression workbook (Lambda_Library.xlsx) or ``"univariate"``
        for the standalone Univariate workbook (Lambda_Library_Univariate.xlsx).
        The two artifacts have independent version lines under the two-version
        scheme (shared function-library version + per-workbook version); see
        README.md "Versions".
    """
    if artifact == "univariate":
        versions = _UNIVARIATE_VERSIONS
    elif artifact == "regression":
        versions = _VERSIONS
    else:
        raise ValueError(
            f"artifact must be 'regression' or 'univariate', got {artifact!r}"
        )

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

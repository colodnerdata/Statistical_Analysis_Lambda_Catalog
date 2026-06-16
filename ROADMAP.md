# Lambda Library — Roadmap

A native-Excel statistical and regression library (LAMBDA-based, no VBA) intended to
replace and expand upon the Excel Analysis ToolPak. Every release ships **functions plus
a pre-built worksheet** that drives them — the worksheet is a first-class deliverable, not
an afterthought.

Design philosophy: live recalculation, formula transparency, auditability, and no Solver /
no VBA wherever it can be avoided. The goal is that any result can be interrogated by
clicking the cell.

---

## Versioning & Release Conventions

Semantic versioning, `MAJOR.MINOR.PATCH`:

- **MAJOR** — breaking changes to function signatures or sheet structure.
- **MINOR** — new functions or new sheet capabilities.
- **PATCH** — bug fixes, formula corrections, documentation edits.

Each release maintains a changelog. For distribution to non-git users (e.g. cost
estimators), a "Version History" sheet inside the workbook mirrors the changelog so the
history travels with the file.

---

## v1.0 — Multivariate (OLS / MLR)

The complete OLS package and the first stable release. This is the multivariate capstone;
later versions fill in the conceptual foundations beneath it and then climb past OLS.

**Status:** all gate items complete — ready to tag v1.0.0.

**Gate items (completed):**

- ✅ Diagnostic guide sheet — Tier 1 and Tier 2 plot specifications, threshold reference
  table, and "Common Patterns & Next Steps" interpretation guidance.
- ✅ Helper columns on the Regression sheet:
  - Scale-Location: `=SQRT(ABS([Studentized Residuals]))`
  - PRESS residual index: `=[Residuals]/(1-[Hat Diagonal])`
- ✅ Standardized coefficients (Beta weights), added to the Coefficients section.
- ✅ Conditional formatting to highlight diagnostics that fall out of bounds for the given
  alpha or rule-of-thumb threshold (VIF, PRESS R², QQ Correlation, hat diagonal,
  studentized residuals, Cook's distance, Scale-Location, PRESS residual,
  coefficient p-values, Significance F).
- ✅ Prediction-with-filtering validation — `Prediction_Interval` and prediction inputs
  confirmed correct when independent variables are toggled; `Ind_Var_Filter`, optional
  `Coefficient_Name_Col` filter, and `X_new` construction all validated.
- ✅ QC flow — `analyze_regression_sheet.py` computes Python reference values and compares
  against sheet output across six configurations (sparse / medium / full predictor sets ×
  intercept / no-intercept).

**Gate items (all complete — ready to tag v1.0.0):**

- ✅ Default diagnostic charts — placed to the right of the Residual Output section.
  The four Tier 1 and Tier 2 charts that the Diagnostic Guide references should be
  pre-built in the sheet so users see them immediately without manual chart creation:
  Residuals vs. Fitted, Normal Q-Q, Actual vs. Predicted, Scale-Location,
  Cook's Distance, Leverage vs. Studentized, and PRESS Residuals. Charts should
  recalculate live with the spill data from the residual columns.
- ✅ Version History sheet — a workbook-level sheet (per the versioning convention stated
  above) that mirrors the changelog so version history travels with the distributed file.
  Entries should include version number, release date, and summary of changes. Written
  by `write_sheet_version_history.py`, called from `build_production.py`.
- ✅ Conditional formatting on Prediction Inputs variable names — strikethrough on the
  name in column U for any predictor whose toggle in column B is FALSE. Rule must cover
  the full `All_Xs` range dynamically so it remains correct when the user swaps to a
  dataset with more or fewer independent variables. The formula should reference the
  `Ind_Var_Filter` named range (trimmed to `n` rows via `TAKE`) so no hard-coded row
  count is needed.

**Additional items completed beyond the original gate list:**

- ✅ Regression Instructions sheet — step-by-step guide for adapting the sheet to a new
  dataset, including Name Manager updates and table setup.
- ✅ Number formatting applied uniformly across all output zones: 2 dp for predictor
  summary; 4 dp for statistics, diagnostics, coefficients, prediction interval/inputs, and
  residual columns; scientific notation (2 sig digits) for p-values and Significance F;
  1 dp for SS/MS/F; integers for df and Observations.
- ✅ Word wrap on header row 2 across all zones.
- ✅ Prediction interval refactored to a single spill formula
  (`=Prediction_Interval(...)` in V3) rather than six individual `INDEX(...)` calls.
- ✅ Diagnostic guide heading and subheading highlights span the correct column widths
  (4 columns for Tier 1, Tier 2, and Diagnostic Threshold sections; 3 for Common
  Patterns & Next Steps).

**Engine (all implemented):** full model-fit and ANOVA statistics, coefficient
inference (SE, t, p, CIs, partial R²/correlation), multicollinearity (VIF, Tolerance,
correlation matrix), residual and influence diagnostics (residuals, studentized residuals,
hat diagonal, Cook's distance, Q-Q machinery, Durbin-Watson), cross-validation (PRESS,
LOOCV), information criteria (AIC, AICc, BIC), distributional exploration (skewness,
kurtosis, Pearson/Spearman), and prediction. Filter argument supports stratified OLS
natively.

---

## v2.0 — Univariate

The foundation layer. A single input column drives three coordinated sections, telling one
story: **describe** the data, **visualize** its shape, then **formalize** that shape with a
fitted distribution. The skewness value from the descriptive section motivates the
distribution choice in the fitting section.

### 1. Descriptive statistics

Mean, median, mode, standard deviation, variance, min, max, range, skewness, kurtosis,
count, and **missing count** (surfaced prominently — honest handling of dirty real-world
data rather than silently dropping blanks).

### 2. Histogram binning

Three **separate** bin tables, side by side, one per bin-selection method:

- Sturges
- Scott
- Freedman-Diaconis

Ragged bottoms are expected and correct (methods disagree on bin count). Bin labels stored
as **upper edges** (half-open `(lower, upper]` intervals), matching Excel's native
`FREQUENCY` and histogram-chart convention. Each method gets its own column chart
(gap width 0) built from the computed bin table — not the native histogram chart type,
since the point is to demonstrate the library's own bin logic.

### 3. Distribution fitting

Fit a panel of candidate distributions and rank them in a single comparison table.

**Candidate distributions (8):**

| Distribution | Parameters | Fitting approach |
|---|---|---|
| Normal | mean, sd | Closed-form MLE (sample mean, variance) |
| Lognormal | mean, sd of logs | Closed-form — normal MLE on ln(x) |
| Exponential | rate | Closed-form MLE (1/mean) |
| Weibull | shape, scale | Grid-search MLE (two-input Data Table, multi-stage) |
| Gamma | shape, rate | Grid-search MLE (two-input Data Table, multi-stage) |
| Triangular | min, mode, max | Direct min/mode/max; true MLE non-differentiable at the mode |
| Beta | α, β (on [0,1]) | Grid-search MLE (two-input Data Table); requires rescaling to [0,1] |
| BetaPERT | min, mode, max | Closed-form by construction (PERT reparameterization of Beta) |

**The MLE challenge and its resolution:** the wall is at the two-parameter shape
distributions (Weibull, Gamma, Beta) where a parameter cannot be isolated algebraically,
plus the separate oddity of Triangular's non-differentiable likelihood. The no-Solver,
no-VBA resolution is **grid search over the parameter space using native two-input Data
Tables**, minimizing negative log-likelihood. This is genuine MLE — the grid is a
zeroth-order optimizer (no derivatives) standing in for Newton-Raphson — so the library can
legitimately claim MLE without Solver. The reframing: the wall was never "MLE without
Solver"; it was "MLE in closed form." Grid search clears the no-Solver bar for the entire
two-parameter likelihood class.

**Grid-search mechanics (Weibull, Gamma, Beta):**

- Row header = parameter 1 candidates; column header = parameter 2 candidates (min, max,
  bin count). One parameter → one-input Data Table; two → two-input Data Table.
- Corner cell evaluates NLL; the Data Table fills the grid across all parameter
  combinations with full live recalc.
- `MIN` of the grid gives the best NLL; a flatten-and-locate step (e.g. `TOCOL`, then map
  the matched position back to row/column via integer division and modulo against grid
  width) recovers the fitted parameters. Prototype this 2D argmin lookup in isolation — it
  is the one genuinely tricky cell.

**Multi-stage refinement (2, optionally 3 stages):** each stage is structurally identical
— same dimensions, same NLL corner, same argmin lookup — differing only in the min/max
bounds it sweeps. Each stage centers on the previous stage's located minimum and zooms in.
Adding or removing a stage is a copy/delete of an identical block, not a redesign. Geometric
precision gain at linear cost: a 20×20 grid refined twice reaches ~20³ effective steps from
~1200 evaluations. Two stages suffice for practical fitting (GoF ranking rarely changes past
3–4 sig figs); a third is available for principle or ill-conditioned data.

*Critical bounds rule:* the next stage's range must extend **one current-grid step on either
side** of the located cell — zoom into the landed cell plus its neighbors, not a fixed
shrink factor. A fixed shrink can drop a near-boundary true minimum outside the refined
window and converge to the wrong point.

**Conditional formatting on the grid (heatmap):** color-scale the NLL grid to visualize the
likelihood basin, confirm the minimum sits in the interior, and expose overflow/error
regions at a glance. Keeps the machinery legible rather than opaque — consistent with the
library's transparency philosophy.

**Two guards:**

- *Minimum on a grid boundary* — if the located parameter is the first/last row or column,
  the true optimum is likely outside the swept range and refinement will chase it off the
  edge. Conditionally format the located cell when it lands on a border to flag "widen
  stage-1 bounds."
- *Undefined / infinite NLL* — Beta on [0,1] blows up if data sits exactly at 0 or 1;
  Gamma/Weibull overflow as shape approaches 0. Wrap the NLL cell in `IFERROR` mapping bad
  combinations to a large finite sentinel, so `MIN` and the argmin lookup stay robust and
  the heatmap cleanly shows the overflow zones.

Triangular remains the exception: its likelihood is non-differentiable at the mode, so it is
fit by direct min/mode/max estimation rather than grid-search MLE.

**Fit comparison table** — one row per distribution, best fit highlighted, columns:

- AIC
- BIC
- Anderson-Darling statistic
- Kolmogorov-Smirnov statistic
- Q-Q correlation (reuses existing regression Q-Q machinery)

Plus a per-distribution Q-Q plot for visual fit assessment.

**Implementation note:** Triangular, Beta, and BetaPERT are bounded-support; the others are
unbounded or semi-bounded. AIC/BIC compare cleanly across all (likelihood-based), but
Anderson-Darling and K-S depend on the fitted CDF and behave differently at support edges
for the bounded distributions — handle edge behavior deliberately to avoid misleading GoF
values.

**Open design decision:** method of moments vs. MLE as the default fitting path for the
closed-form-friendly distributions. Method of moments keeps everything live and
formula-transparent; resolve before building.

---

## v3.0 — Resampling & Simulation

Bootstrap confidence intervals and Monte Carlo simulation. Validated as worthwhile
differentiators by their presence in Pyrcz's Excel demos and squarely in cost-estimation
territory (three-point estimates, MCS, risk analysis). These do not depend on the
two-sample or ANOVA work, so they come early. Bootstrap and Monte Carlo pair naturally and
may share a single sheet.

---

## v4.0+ — Future (sequence TBD)

Deliberately left loose. Candidate milestones, roughly in conceptual order:

- **Bivariate / two-sample** — t-tests (one-sample, two-sample equal variance, Welch
  unequal variance, paired), F-test for variance equality feeding a recommendation on which
  t-test to use, and Covariance to complement the existing Correlation. (Pyrcz's
  "difference in means" / "difference in variances" demos map here.)
- **Multi-group means (ANOVA)** — one-way ANOVA, implemented as regression on group
  dummies, reusing the existing SS/MS/F machinery. A natural hinge showing ANOVA *is*
  regression.
- **Time series** — Moving Average, Exponential Smoothing.
- **Weighted regression (WLS)** — its own milestone. Extends the OLS engine to non-constant
  error variance; closes the loop opened by v1's Scale-Location diagnostic (v1 detects
  heteroscedasticity, WLS addresses it). Likely threads a weights argument through the core
  regression functions with WLS-aware diagnostics.
- **Fourier analysis** — to be added later.
- **Decision analysis** — possible long-tail addition (loss functions), cost/risk oriented.

---

## Analysis ToolPak Parity Reference

The ToolPak ships 19 tools. Tracking which are covered, planned, or intentionally skipped.

**Covered or exceeded (v1):** Regression (with diagnostics, influence measures, cross-validation, information criteria, and prediction), Correlation, partial descriptive stats.

**Planned:** Descriptive Statistics + Histogram + Rank/Percentile (v2); t-tests, F-test,
Covariance (future two-sample); one-way ANOVA (future); Moving Average + Exponential
Smoothing (future time series).

**Intentionally skipped:**

- **z-Test (two-sample for means)** — assumes known population variance; rarely applicable.
- **Fourier Analysis** — engineering/signal domain; out of scope (may add later).
- **Two-factor ANOVA** — complexity vs. demand.
- **Random Number Generation / Sampling** — largely redundant with native Excel functions.

**Why the library exists (ToolPak flaws it fixes):** ToolPak output is static (pasted
values that never update when inputs change), opaque (no formula trace), one-sheet-at-a-time
with manual reruns, locked behind a modal dialog, and diagnostically dated (no VIF, Cook's
distance, leverage, studentized residuals, PRESS, AIC/BIC, or cross-validation). The Lambda
Library is live, transparent, auditable, reusable, and diagnostically modern.

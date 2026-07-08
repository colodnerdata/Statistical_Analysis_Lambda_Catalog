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

- **MAJOR** — a breaking change to the library's **public interface**, defined below.
- **MINOR** — new functions, new sheets, or new sheet capabilities that do **not**
  break the public interface. A large additive feature is still a MINOR.
- **PATCH** — bug fixes, formula corrections, documentation edits.

**What the "public interface" (API) is for this library.** Unlike a code library
whose API is a set of function signatures, this library's API is **the user's inputs
to the workbook**: the layout and semantics of the input cells and control blocks a
user fills in, the sheet-scoped and workbook-scoped named ranges a user may reference
in their own formulas, and the meaning of an existing specification already saved in a
user's file. A release is **breaking (MAJOR)** when a workbook a user built against the
prior version would, on adopting the new version, either stop working or silently
compute something different — e.g. a control block moving or changing meaning, a named
range disappearing or being repurposed, or `x_s()` changing what it returns for the
same inputs. A release is **non-breaking (MINOR)** when every specification valid under
the prior version produces the same result, and the new capability is reached only by
new inputs the user opts into. Canonical function *names* are part of the API only to
the extent a user types them into cells; internal implementation is not.

This is why, for example, Univariate is a MINOR (a wholly new sheet; no existing input
changes meaning) while Specification-Driven Regression is a MAJOR (it changes what
`x_s()` returns and restructures the Regression sheet's control block — existing
formulas and saved specs change behavior).

Each release maintains a changelog. For distribution to non-git users (e.g. cost
estimators), a "Version History" sheet inside the workbook mirrors the changelog so the
history travels with the file. That sheet carries an explicit **`Breaking?` (yes/no)**
column so a workbook user gets the one signal the version number is *for* — "do my
existing inputs still work?" — without the number also having to convey "how big is
this release." Deliverable size is described in the changelog prose; breakage is the
flag.

**Version ladder (current plan):**

| Version | Milestone | Breaking? | Status |
|---|---|---|---|
| v1.0 | Multivariate OLS / MLR | — (baseline) | Shipped |
| v1.1 | Univariate (descriptives, histograms, distribution fitting) | No | **Shipped 2026-06-29** (workbook 1.1.0; renumbered from 2.0.0). MoM-vs-MLE resolved: MLE throughout. New sheet, no existing input changes meaning. PDF functions dropped as unnecessary — the histogram tables already compute per-bin probabilities as CDF deltas between bin boundaries. The two post-release leftovers (per-distribution Q-Q plots and combo-chart overlay lines built on those CDF-delta columns) are now implemented and ship with the next workbook build |
| v1.2 | Workbook hardening & regression usability (Name Manager notes, identity-line data series, intercept-only and undersized-sample guards, LOOCV_Residual, build retry/RPC handling) | No | **Shipped 2026-07-03** (workbook 1.2.0; renumbered from 2.1.0) |
| v2.0 | Specification-Driven Regression (roles: Continuous / Categorical) | **Yes** | **Shipped 2026-07-05** (workbook 2.0.0; renumbered from 3.0.0) — MAJOR. Changed `x_s()` return semantics and restructured the Regression control block; includes the canonical rename pass. Shipped with `Transform` as a reserved placeholder column as planned; users transform their own variables via extra input-table columns in the interim |
| v2.1 | Sequence axis + gap-aware longitudinal + serial-correlation diagnostics + Fixed Effects (Role axis, one-way only) | No | In progress — pending Sequence auto-detection/override fix (`Sequence Period` (I) / `Period In Use` (J) rename, spill-collision guard) and FE engine work; no release scheduled. Bundles the Sequence/BFN/QC chain with the FE engine as a single release so users never see "FE is in the dropdown but the engine is forthcoming" |
| v2.2 | Transforms (Response / Predictor Log, unit-space comparability) + the standalone Data Transformation function library | No | Planned — MINOR. Wires the reserved spec column G and ships the user-callable transform functions (Center, Zscore, Winsorize, Lag_By, …). Completes the Regression sheet as a fully functional deliverable |
| v2.3 | Model Comparison Sheet | No | Planned — MINOR, a *nice-to-have*. Read-only across finished Regression sheets; ships after Transforms so its comparisons are unit-space-honest from day one |
| v2.4 | Resampling & Simulation (bootstrap, Monte Carlo) | No | Planned — MINOR. Pre-drawn random table (`Bootstrap_Random_Draws` named range) indexed at use time; non-volatile by design (every recalc reproduces the same draw). The QA build seeds the table from the same SHA-derived seed as `analysis_cache.py` |
| v2.5 | Bivariate / two-sample (one-sample t, two-sample t [equal-var / Welch / paired], F-test, Covariance) | No | Claimed — next MINOR after v2.4. F-test feeds a recommendation cell that selects the t-test variant; Covariance complements the existing `Correlation_Matrix` |
| v2.6 | `Weight` Role (WLS) | No | Claimed — after v2.5. User-supplied weights as the first stage; variance-driver-derived weights and FGLS as v2.6+ follow-ons. Engine signature addition with a default-uniform `[Weights]` argument (default-uniform → identical to OLS, the v2.1 `[DF_Absorbed]` precedent) |
| v2.7+ | Two-way FE, `Cluster` and `Time` Roles, Time series, ANOVA, Fourier, Decision | mixed | Unordered (deliberate — see Future section). Two-way FE has forward wiring from the v2.1 FE engine; `Cluster` has forward wiring from `Serial_Correlation_Group()`'s dormant branch; the rest are design-not-started |

**Ladder rationale.** Under the interface definition above, only one planned milestone
breaks user inputs — Specification-Driven Regression — so it alone takes the next major
number (v2.0). Everything after it is additive and opt-in, forming a v2.x train
directly analogous to Python's 3.x line: one breaking 3.0 followed by years of large
but non-breaking minors (async/await, pattern matching) that never forced a new major.
The next MAJOR is reserved for the next genuine interface break, whenever that is.

Univariate shipped **before** Specification-Driven Regression despite the lower version
gap, as planned: its engine was already implemented (all 8 CDF/NLL distribution
functions, the grid-search optimizer, all three binning methods and their `Bin_*`
helpers, and all four goodness-of-fit measures) and its sheet writer was wired into
`build_production.py`. Specification-Driven Regression was greenfield by comparison,
so the near-finished milestone shipped first.

**Fixed Effects breakage flag (v2.1) — RESOLVED as non-breaking.** The absorbed-df
correction is threaded as an **optional `[DF_Absorbed]` argument defaulting to 0**,
leaving the no-FE df path untouched, so a model with no Fixed-Effects Role behaves
identically to v2.0. FE therefore stays a MINOR at v2.1. (The alternative that would
have made it MAJOR — changing the default behavior of an existing engine function, e.g.
altering the shared df formula's default for all models — is explicitly not the chosen
implementation. `y_s()` is introduced as a new function, not a replacement wired into
existing no-FE call sites.) Full mechanics under v2.1 Resolved decisions.

---

## Naming Convention

Canonical function names use **Title_Case_With_Underscores**, fully spelled out — no
abbreviations (e.g. `Absorb_Two_Way_Fixed_Effects`, not `ABSORB_2FE`). This is a
deliberate departure from Excel's all-caps native functions: a mixed-case name in a
formula bar is immediately recognizable as library code, not a built-in. A reviewer
should be able to scan a nested formula and tell at a glance which calls are library
functions versus native Excel (`SUM`, `FILTER`, `XLOOKUP`) without cross-referencing a
function list.

Rules:

- Full English words only. Spell out what an abbreviation would have stood for
  (`Two_Way`, not `2WAY`; `Fixed_Effects`, not `FE`).
- Numerals appear only when the numeral is itself the statistical quantity (e.g. a
  literal lag of `2`, not a stand-in for "two-way").
- Underscores separate words; no camelCase.
- One canonical name, one LAMBDA, one place it can be wrong.

**Existing names predate this convention.** Functions shipped in v1.0 (`Observations`,
`DF_Regression`, `R_squared`, etc.) are already Title_Case-ish but were not written
against an explicit rule. No retroactive rename is planned for v1.0 names unless a
broader breaking-change pass is already underway — folding a pure cosmetic rename into
an unrelated MAJOR bump avoids spending two breaking changes on what could be one.
**Decision made and executed: the rename pass shipped inside the v2.0 MAJOR bump**
(2026-07-05) — no pre-convention names remain in `lambda_functions.json`.

**Sign-off record (v2.0 rename pass, 2026-07-03).** The rename table's DECIDE rows
were resolved as follows, and these exceptions are now part of the convention:

- **Retained initialisms/abbreviations** — words in their own right, never expanded:
  `AIC`, `AICc`, `BIC`, `VIF`, `PRESS`, `CDF`, `NLL`, `LOOCV`, `PERT`, `R`
  (Pearson/Spearman R), `QQ` (as in `QQ_Correlation`), `GoF` (the `GoF_*` family),
  and the classical ANOVA-table shorthands **`MS`, `SS`, `SE`** (`MS_Residual`,
  `SS_Total`, `SE_Regression`, `SE_Coefficients`, etc. keep their names).
- `CDF_BetaPERT` / `NLL_BetaPERT` — kept (no underscore inserted).
- `P_Value_F` → `F_Statistic_P_Value` (pairs with `F_Stat` → `F_Statistic`).
- `Grid_Argmin` → `Grid_Argument_Minimum` (spelled out per the convention).

### Alias layer (future, optional)

A separate, optional layer of short, ALL-CAPS aliases may be added in a later pass for
power-user typing speed (e.g. `ABSORB2FE` as an alias for
`Absorb_Two_Way_Fixed_Effects`). Aliases are thin wrappers — each alias LAMBDA's entire
body is a call to the canonical function, with no independent logic:

```excel
ABSORB2FE = LAMBDA(x, group1, group2, [include], [passes],
    Absorb_Two_Way_Fixed_Effects(x, group1, group2, include, passes)
)
```

This keeps a single source of truth: if the canonical implementation changes, every
alias inherits the fix automatically. Aliases are never the documented or taught form —
they exist purely as optional shortcuts and should be introduced only after the
canonical library is stable, to avoid maintaining two names for a function that's still
under active revision.

---

## Function Categories

Every catalog entry carries a `category` and a `subcategory`, used to drive filtering on
the `LAMBDA_functions` sheet. This taxonomy is **purely functional — it does not encode
version.** Version is a property of the library's release history (tracked in the
changelog and Version History sheet); category is a property of what the function *does*,
and a function's category should not change just because it shipped in a later release.

Subcategories are scoped *within* a category — each category defines its own
subcategory list rather than sharing one flat list across categories, so a category can
grow its own subdivisions independently as it fills up.

| Category | Subcategories |
|---|---|
| **Model Construction** | MLR Core · Coefficient Inference · Prediction · Specification & Design Matrix · Transforms & Unit-Space |
| **Diagnostics** | Residual · Influence & Leverage · Multicollinearity · Cross-Validation · Information Criteria |
| **Data Transformation** | Sample Construction & Diagnostics · Location & Scale · Group & Panel · Categorical & Model Construction · Longitudinal & Panel-Time |
| **Distribution Fitting** | Descriptive · Histogram Binning · Parameter Estimation · Goodness-of-Fit |
| **Resampling & Simulation** | Bootstrap · Monte Carlo |
| **Model Comparison** | Spec String & Registry *(provisional — promote from a subcategory only once it holds 2+ functions; see v2.3)* |

*Change from prior revision:* **Specification & Design Matrix** added under Model
Construction to house the v2.0 constructor functions (see below); **Transforms &
Unit-Space** added under the same category for the v2.2 unit-space fit statistics. The
Data Transformation subcategories are unchanged; those functions now serve double duty
as constructor internals and standalone user-callable transforms. **Model Comparison**
is listed provisionally as a top-level category for the v2.3 sheet's single function —
per the "categories describe what a function does" rule it earns top-level status once
it holds 2+ functions; until then it may equally live as a Model Construction
subcategory.

This table is the source of truth for the controlled vocabulary; `category` and
`subcategory` values in `lambda_functions.json` should be drawn only from this list,
not invented ad hoc per function.

---

## v1.0 — Multivariate (OLS / MLR) — SHIPPED

The complete OLS package and the first stable release.

**Engine:** full model-fit and ANOVA statistics, coefficient inference (SE, t, p, CIs,
partial R²/correlation), multicollinearity (VIF, Tolerance, correlation matrix),
residual and influence diagnostics (residuals, studentized residuals, hat diagonal,
Cook's distance, Q-Q machinery, Durbin-Watson), cross-validation (PRESS, LOOCV),
information criteria (AIC, AICc, BIC), distributional exploration (skewness, kurtosis,
Pearson/Spearman), and prediction. `Include` argument supports stratified OLS natively.

**Sheet:** five-zone Regression sheet (model inputs · predictor summary · regression
outputs · prediction · residual output), Diagnostic Guide sheet, Regression Instructions
sheet, Version History sheet, seven pre-built diagnostic charts, conditional formatting
on out-of-bounds diagnostics, QC via `analyze_regression_sheet.py` across six
configurations (sparse/medium/full predictor sets × intercept/no-intercept).

(The full v1.0 gate history is in the git history of this file rather than restated here; the bullets above are the still-relevant surface of the release.)

## v1.1 — Univariate — SHIPPED 2026-06-29

*(Renumbered from v2.0 to v1.1 — a new sheet that changes no existing input is a MINOR
under the interface definition in Versioning & Release Conventions. Released in the
workbook Version History as 1.1.0.)*

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

**Grid-search mechanics (implemented for Weibull; reusable for Gamma and Beta):**

- Shape is the column parameter and Scale is the row parameter. Their visible Input cells
  are the `RowInput` and `ColumnInput` substitution cells for Excel's two-input Data Table.
- The corner cell evaluates NLL; column candidates sit immediately above the body and row
  candidates immediately to its left. The sheet-scoped grid name covers the body only.
- `Grid_Argument_Minimum(grid)` returns `minimum | 1-based row | 1-based column` horizontally and
  resolves ties to the first row-major occurrence.
- `Grid_Search_Optimum(grid)` returns the best column parameter followed by the best row
  parameter vertically. It requires the physical axes to remain adjacent to the body.
- **Rows/Columns** records the generated number of values per axis. A value of 20 creates a
  20×20 grid (400 evaluations); editing the worksheet value does not resize the Data Table.
- Both endpoints are included, so `Step Size = (Max-Min)/(Rows/Columns-1)`. Axis `SEQUENCE`
  formulas consume the visible Step Size cells rather than repeating that calculation.

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

Plus a per-distribution Q-Q plot for visual fit assessment. *(Built post-release:
the per-distribution Q-Q plots and the histogram distribution-overlay series — the
two v1.1 deliverables that did not ship on 2026-06-29 — are now implemented in
`write_sheet_univariate.py` and ship with the next workbook build. The Q-Q data zone
(cols CW–DF) holds Hazen plotting positions, the sorted sample, and eight
theoretical-quantile columns; eight XY scatter charts with identity-line data series
stack under the histogram charts.)*

**Histogram overlays — PDF functions dropped as unnecessary.** The original plan
called for `PDF_*` LAMBDAs evaluated at bin midpoints to draw fitted curves over the
histograms. That is redundant: each histogram block already computes, for all 8
distributions, the per-bin probability as the **CDF delta between the bin
boundaries** — `CDF(upper edge) − CDF(lower edge)` via the existing `CDF_*`
functions. That delta is the bin's probability mass (the PDF integrated exactly over
the bin, which is more faithful to the histogram than a midpoint PDF evaluation), so no PDF
functions are needed and none will be implemented.

The overlay is instead delivered (implemented, ships with the next workbook build)
by converting each histogram chart into a **combo chart**: the existing column
series (counts, gap width 0) plus one line series per distribution sourced from the
histogram table's CDF-delta columns — smoothed lines (`Smooth = True`) with **no
markers** so the fitted shapes read as curves. The probabilities-vs-counts axis
question was settled as **expected counts on the shared count axis**: each line
series references a `UV_<method>_<Dist>_Expected` named formula that multiplies the
OFFSET-sized CDF-delta column by the Count stat cell (probability × n), rather than
introducing a secondary axis.

**Open question — dynamically hiding the worst-fit / error columns.** Ideally the
overlay would show only the credible fits: columns for distributions with the worst
GoF ranking or N/A errors would be hidden, and since Excel charts plot only visible
cells by default (`PlotVisibleOnly`), hidden columns would drop out of the combo
chart automatically — the desired behavior. It is not yet known whether column
hiding can be driven dynamically from cell values without VBA (manual hiding works
but is a hand step; data-driven hiding may be VBA-only, which the library forbids).
A candidate no-VBA alternative: have the series formula emit `NA()` across a
suppressed distribution's column, since line charts skip `#N/A` points — same chart
effect, no hiding. To be investigated at implementation (tracked in TODOs.md).

**Implementation note:** Triangular, Beta, and BetaPERT are bounded-support; the others are
unbounded or semi-bounded. AIC/BIC compare cleanly across all (likelihood-based), but
Anderson-Darling and K-S depend on the fitted CDF and behave differently at support edges
for the bounded distributions — handle edge behavior deliberately to avoid misleading GoF
values.

**Design decision — RESOLVED: MLE throughout** (see `write_sheet_univariate.py`).
Closed-form MLE for Normal, Lognormal, and Exponential; grid-search MLE for the
two-parameter shape family (Weibull, Gamma, Beta); direct min/mode/max estimation for
Triangular and BetaPERT (non-differentiable / by-construction cases). Method of moments
was not adopted as a default path anywhere.

---

## v1.2 — Workbook hardening & regression usability — SHIPPED 2026-07-03

*(Released in the workbook Version History as 1.2.0; renumbered from 2.1.0. A
non-breaking MINOR consolidating the v1.0 / v1.1 surface into a workbook that holds up
under hand-use by cost estimators and other non-git users. The technical core is
unchanged; every existing v1.0 / v1.1 input computes identically. Five additions:)*

**1. Name Manager notes on every worksheet-scoped name.** Every worksheet-scoped
named range (constructor closures, `RegChart*` chart feeds, helper formulas) now
carries a `Comment` in the workbook's Name Manager describing what the name is for
and which sheet zone owns it. The motivation: a non-git user opening
`Lambda_Library.xlsx` in Excel has no way to read `ROADMAP.md` or
`write_sheet_regression.py` — the only documentation they can reach is the workbook
itself. The Name Manager `Comments` column is the in-workbook index. The loop in
`_setup_local_names` in `write_sheet_regression.py` is the single source of the
`Comment` strings; the same pattern applies to the other sheet writers.

**2. Identity-line data series on the diagnostic charts (replacing the failed
`Shapes.AddLine` approach).** The QQ, Actual-vs-Predicted, and Studentized-vs-Leverage
charts all need a `y = x` reference line. v1.0 attempted to draw it with
`chart.Shapes.AddLine(...)` — a position-in-fixed-plot-area-pixels approach that
silently went wrong when the chart was resized, moved, or its axis scaling changed.
v1.2 replaces every reference line with a real data series pointing both
`XValues` and `Values` at the same named range (`RegChartQQX` for the QQ chart,
`RegChartFitY` for the Actual-vs-Predicted chart, the corresponding `RegChart*`
range for the leverage chart), with `ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS`
and `Name = "Identity"`. The series then auto-rescales with the chart, since
its pixel coordinates are *not* fixed. The "Never draw reference lines as
shapes" guide in `CLAUDE.md` / `AGENTS.md` captures the rationale; the helper
`_add_identity_line` in `write_sheet_regression.py` is the implementation.

**3. Intercept-only and undersized-sample guards.** Two narrow edge cases that v1.0
returned silently-wrong numbers for. With the v1.0 input layout, an intercept-only
model (every `Include` FALSE, intercept TRUE) and a model with n ≤ k + 1
observations both computed reasonable-looking OLS output by accident — k = 0
trivially; the undersized case produced a singular `Gram_Inverse` and the engine
fell through to a near-`#NUM!` result. v1.2 adds explicit guards in
`analyze_regression_sheet.py` (and the corresponding LAMBDA-side checks in
`regression_shared.py`) that surface a visible `n/a — undersized sample` or
`n/a — intercept-only` token in the affected output cells rather than returning
silently-wrong numbers. The status block on the Regression sheet echoes the same
token. Visible failure is the design philosophy.

**4. `LOOCV_Residual` (new function, pair to `LOOCV_Prediction`).** v1.0 shipped
`LOOCV_Prediction(X_s, Y, n, [Allow_Intercept], [Include])` — the leave-one-out
predicted value for a single observation. v1.2 adds `LOOCV_Residual(X_s, Y,
[Allow_Intercept], [Include])` (catalog entry in `lambda_functions.json`) returning the leave-one-out residual vector `eᵢ / (1 − hᵢ)` in a single call, which makes `PRESS = SUMSQ(LOOCV_Residual(...))` an obvious idiom and
unifies the LOOCV machinery under the Sherman-Morrison-Woodbury update that
`LOOCV_Prediction` was already using internally.

**5. Build-phase retry / RPC handling.** Excel's COM automation occasionally
returns `RPC server unavailable` mid-`Calculate` — a transient that surfaced
during the v1.1 build under heavy Data Table evaluation. v1.2 introduces
`_retry_on_open` in `build_production.py` and splits `main()` into two phases:
phase 1 (sheet writes, no retry — minutes long, the only phase
that benefits from a clean Excel instance) and phase 2 (recalculate + save,
short, the most likely phase to hit the RPC failure). The retry-on-open
behavior is applied to phase 2 only, so a transient failure there doesn't
restart the multi-minute sheet-writing phase. The same pattern lives in
`build_qc.py`.

**Resolved decisions (v1.2):**

- *Name Manager note scope — RESOLVED: every worksheet-scoped name gets a note.*
  The alternative (notes only on names a non-git user is likely to encounter) was
  rejected because the line between "likely" and "unlikely" is too subjective to
  keep stable across releases. The blanket rule is the only one that scales.
- *Identity-line technique — RESOLVED: real data series, not `Shapes.AddLine`.*
  See point 2 above; the `Shapes.AddLine` approach was tried and abandoned in
  the v1.0 → v1.1 cycle.
- *Undersized-sample failure mode — RESOLVED: visible `n/a` token, not silent
  numbers.* Consistent with the library's "visible failure" philosophy for
  categorical degeneracy, invalid reference levels, and the
  blank-categorical caveat (see v2.0).

---

## v2.0 — Specification-Driven Regression — SHIPPED 2026-07-05

*(Released in the workbook Version History as 2.0.0; renumbered from 3.0.0. Shipped
across the Model Construction PRs plus the Regression-sheet spec-driven changeover
(PR #92). The canonical rename pass shipped inside this MAJOR as planned. The reserved
`Order`/`Transform` columns shipped unread by any formula, confirmed by construction.)*

**The central idea:** factor (categorical) and panel (fixed-effects) regression are not
new estimators — they are OLS on a transformed design matrix. Rather than telling the
user to manipulate their dataset into MLR form by hand, the sheet's control block
becomes a **declarative model specification spanning the entire source table**, and
`x_s()` is promoted from a column filter into a **model-matrix constructor** that reads
the spec and emits the numeric design matrix. Because every engine function already
consumes `x_s()`, the entire engine inherits the new capability without a signature
change.

The specification dissolves all three of the v1 sheet's hard-wired range names into
declarations: `y` (hard-wired) → the row whose Role is Response; `All_Xs` (hard-wired
span) → the table reference; `Regression_Sample_Include` (hard-wired to `[Full_Data]`)
→ the column(s) whose Role is Filter. After v2.0 there is nothing left to hard-wire —
the spec block plus one source-table reference *is* the model.

> **Supersession notes (recorded per the open-decisions convention):**
>
> 1. *Separate Factor Regression and Panel Regression sheets* with on-sheet staging
>    bands — replaced by the one spec-driven sheet; factor and panel become documented
>    walkthroughs in the Regression Instructions sheet.
> 2. *The WLS-as-optional-`[Weights]`-argument vs. parallel-function-set decision* —
>    re-litigated as a **`Weight` value on the Role axis** (see *Future roles*). The
>    dedicated WLS Regression sheet plan is likewise superseded.
> 3. *The single-axis "Predictor Type" design* (one column holding Continuous /
>    Categorical / Fixed Effects) — superseded by the **two-axis taxonomy** below.
>    Fixed Effects was never a predictor type: it contributes no columns and no
>    coefficients. It moves to the Role axis (v2.1), the Type axis becomes permanently
>    Continuous/Categorical, and all future growth happens on Role.

### The taxonomy: two model axes plus a structural axis

| Axis | Values | Future values | Meaning |
|---|---|---|---|
| **Variable Role** | Response · Predictor · Identifier · Filter · Omit | Fixed Effects, Weight, Time (v2.1+) | What the column *is* in the model |
| **Predictor Type** | Continuous · Categorical | *(closed — never grows)* | How a Predictor *enters the design matrix* — meaningful only when Role = Predictor |
| **Sequence** *(structural, post-v2.0)* | TRUE · blank | *(flag — never grows)* | Which column *orders* the data, for lag/difference/serial-correlation features |

Literature precedent: this is essentially the tidymodels `recipes` role system
(outcome / predictor / ID) plus Stata's factor-variable notation on the Type axis.

**The Sequence structural axis** is deliberately NOT a Role value and NOT a
Predictor Type — it annotates *structure*, orthogonal to what a column is in
the model: a column can be Role = Predictor, Type = Continuous AND
Sequence = TRUE simultaneously (the escalation-index case), and an
Identifier like Year is a typical sequence axis. Cardinality is
**zero-or-one**: zero flags is a valid spec (non-panel data); two-plus is a
visible status-line error (same pattern as exactly-one-Response, with a >1
threshold). No constructor formula reads the flag yet — it exists so the
lag/difference/serial-correlation features land on a declared axis. Not to
be confused with the reserved Order column (F), which is term-ordering.

**Role semantics (v2.0):**

- **Response** — the dependent variable. Derived, not hard-wired: `y` becomes "the
  column whose Role is Response" (`XMATCH` over the Role range). Swapping the
  dependent variable is a dropdown change — cost vs. log-cost vs. cost-per-unit
  against the same table is a workflow feature, not plumbing. Validation: **exactly
  one** Response (status-block error at zero or two-plus).
- **Predictor** — candidate model input; the only Role for which Include, Type,
  Reference, and Levels are meaningful (cascading relevance, below).
- **Identifier** — labeling columns (Country, Year): available to the residual-output
  zone for row labeling; never enter the design matrix.
- **Filter** — supplies the row mask. **Zero-or-more, ANDed:** no Filter columns →
  all rows included; multiple → logical AND, which is declarative stratification
  (e.g. `Full_Data` AND a hand-built "Developing only" flag). AND-composition is a
  documented semantic, so it passes the visible-failure test. Filter columns hold
  boolean / 0-1 values; blanks are FALSE.
- **Omit** — dataset semantics, set once: "never a candidate for anything" (helper
  columns, notes). Distinct from Include, which is *model-iteration state* flipped
  constantly ("candidate, currently out of this run"). Collapsing them would destroy
  the record of what a column *is* every time a toggle experiment runs. Cost of
  keeping both: one dropdown value.

**The effective row mask** composes two layers: *automatic role-aware completeness*
(numeric required for Continuous Predictors and the Response; non-blank for
Categorical Predictors; ignored for Identifier/Filter/Omit) **AND** *declared Filter
columns*. This settles the architecture of the former role-aware-completeness open
item by construction; only the auto-completeness LAMBDA remains to build, and the
hard-coded `Data_Completeness(...[Life expectancy]:[Schooling])` span dies with it.

### The specification block (columns A–K)

The spec spans **every column of the source table**, one row per column:

| Col | Contents | UX |
|---|---|---|
| A | Variable name | Spills from the table's header row |
| B | **Role** | Dropdown: `Response` · `Predictor` · `Identifier` · `Filter` · `Omit`. Pre-filled by the build — no blank cells. Role precedes Include because it is the larger declaration. |
| C | Include toggle | Orange input; meaningful only when Role = Predictor |
| D | **Predictor Type** | Dropdown: `Continuous` · `Categorical`; meaningful only when Role = Predictor; pre-filled `Continuous` |
| E | **Reference Level** | Orange input, meaningful only for Categorical Predictors. Blank = **first level in sort order** (confirmed default, matching R). CF: red when the entered level does not exist in the analysis sample. |
| F | **Order** *(reserved, not implemented v2.0)* | Input, integer. Will control user-specified ordering of Identifier columns in the row-label text-join; v2.0 always joins in table order. Present now so the layout absorbs the feature without a future column insertion. Cell comment marks it reserved; no validation yet (no fixed domain). |
| G | **Transform** *(reserved, not implemented v2.0)* | Input, dropdown. Will apply a transform (e.g. `Log`) to a Continuous Response or Predictor. v2.0 dropdown list is `None` only; build pre-fills `None`. Cell comment marks it reserved. |
| H | **Sequence** *(structural axis, post-v2.0)* | Orange input flag, dropdown `TRUE`/blank. The shipped default pre-flags **Year** `TRUE` (the WHO panel's ordering axis; every other row blank) so the Sequence machinery is live at T0; on a non-panel dataset leave it blank. Marks **at most one** variable as the ordering axis. Status line at H2: red error at two-plus flags (zero is valid); per-cell red CF points at the offending rows. Read by the validation layer, by the sequence-spacing layer (`Sequence_Deltas`, `Base_Period_Delta`) since the base-period release, and — since the DW-gate release — by the serial-correlation accessor `Sequence_Column` (which feeds the gated `Durbin_Watson_By` diagnostic cell). No design-matrix constructor consumes it: Sequence orders the data, it never enters the model matrix. |
| I | **Base Period Δ** *(live — base-period release; Sequence companion)* | **Computed-with-override**, the reference-level pattern: pre-filled with a formula showing `Base_Period_Delta_Candidate()` (MODE of within-group consecutive spacings, MIN fallback when no spacing repeats) on the Sequence-flagged row, blank elsewhere; typing a number overrides. Read only by the base-period layer — the `Base_Period_Delta()` accessor (the omitted-`[delta]` default of `Lag_By`/`Difference_By`) and the Sequence Spacing block's Δ-in-use display (yellow CF when overridden). The block (rows 28–34 under the spec) also shows the delta spectrum and verdicts: Regularity (any spacing ≠ Δ), Off-grid (spacing not a whole multiple of Δ), the no-natural-base-period override prompt, and calendar-signature guidance (~28–31/90–92/365–366 clusters → recommend an integer period index upstream; day counts are never quantized to a scalar Δ). |
| J | **Levels** | **Computed display**: distinct level count over the mask-included rows, shown only for Categorical Predictors. Live against stratification. CF: **red when L ≤ 1 while included** (contributes L−1 = 0 columns). Large L needs no flag — the visible count is the warning. |
| K | **Reference In Use** | **Computed display**: the reference level the constructor will actually drop, surfaced even when defaulted. *(Deviation from the earlier draft of this table, recorded: the shipped v2.0 sheet implemented this display here instead of the optional Design Columns audit column; the Σ(design columns) = COLUMNS(x_s()) audit lives in the status strip's `k` cell, and the gap column right of the spec block still visually reserves a future Design Columns slot.)* |

**Reserved-column policy (F, G).** Neither reserved column is read by any
formula — confirmed by construction, not by convention: `X_s()`,
`Constructed_Column_Names()`, `Row_Labels()`, and `Sample_Include()` must not
reference `Spec_Order` or `Spec_Transform` (and may not reference
`Spec_Sequence` or `Spec_Base_Period_Delta` either — those names are consumed
only by the zero-or-one validation and the base-period layer, never by a
constructor). The columns exist purely so the *sheet layout* absorbs the
future feature now; wiring them in a later release is additive (a formula
change), not a second column-insertion breaking the sheet a second time —
exactly how column I went live in the base-period release.

**Cascading relevance:** C–G and J–K gray out (conditional formatting)
whenever Role ≠ Predictor — the same pattern as
Reference-only-for-Categorical, applied one level up. H–I key on the
**Sequence flag itself**, not on Role: they gray on every row that is not the
sequence axis, because Sequence is structural and Role-independent.

**Display derives, never feeds.** Columns J and K must not be inputs to the
constructor. Both the cells and `x_s()` call the same mask-aware primitive
(`Dummy_Levels`); the cells display what the constructor will do. Letting the engine
read a display column would make it load-bearing. One source of truth is the
*function*. (This is also what settled the `Dummy_Code` design: the level-vector
split is required — it makes display and constructor provably consistent, and gives
prediction the training level set.)

### Release scoping: v2.0 vs v2.1

**v2.0 ships the five Roles above with Continuous + Categorical Predictors — `x_s()`
alone, the Response consumed raw.** Anything that demeans or constructs a transformed
response is **v2.1**: the Fixed Effects Role, `y_s()`, absorbed df and its plumbing,
FE prediction. v2.0's constructor is purely column-local; the one global operation —
demeaning — gets its own release. Identifier's v2.1 dividend: Country and Year are
already declared, so panel setup is changing two Role dropdowns.

### Constructor design (`x_s()`) — decisions settled

1. **Spec-order assembly (settled).** `x_s()` iterates spec rows top to bottom and
   stacks each contribution in that order, so the coefficient table reads in spec
   order. Mechanism: `REDUCE` over spec-row indices, `HSTACK` from a full-height
   sentinel seed column, `DROP` the sentinel at the end. Iteration predicate: Role =
   Predictor AND Include = TRUE.
2. **Level-vector split (settled).** *Determine the levels* (mask-aware
   `SORT(UNIQUE(...))` via `Dummy_Levels`) is separate from *encode against a level
   vector* (the broadcast `--(category = retained_levels)`). Training and prediction
   call the same encoder with the same training level vector — an unseen level at
   prediction time encodes as all-zeros, the correct behavior.
3. **Intercept coupling (settled — flag, don't switch).** v2.0 supports reference
   coding (L−1 columns), correct with the intercept on. No-intercept with a
   Categorical included is *flagged* — CF on `Allow_Intercept` plus a cell comment —
   never silently switched to one-hot. Full one-hot is deferred (v2.0.x candidate).

**The `x_s()` row-mask contract:** the constructor reads the effective mask *only* to
fix Categorical level sets; it always emits **full-height** columns and leaves row
filtering to the engine functions, exactly as v1 does. Filtering rows inside `x_s()`
would double-filter against the engine's own mask application.

**Derived names (all late-bound zero-argument LAMBDAs, replacing hard-wired ranges):**
the Response column (`XMATCH("Response", Role range)` → `INDEX` into the source
table), the effective mask (auto-completeness AND Filter columns), and
`Constructed_Column_Names` (the structural twin of `x_s()` — same iteration, same
skip conditions, so names and columns cannot disagree; QC asserts their widths
match). Source-table indirection: structured references cannot be parameterized by
another name without volatile `INDIRECT`, so a single sheet-scoped `Source_Data`
name wrapping the table reference is the retargeting point — dataset changeover is
a one-name edit.

### Degrees of freedom — automatic, not manual *(v2.1)*

Absorbed df is computable from the spec: Σ over Fixed Effects variables of
(group count − 1). Without the correction, coefficients on demeaned data are correct
but **every SE, t, p, CI, MS_Residual, and information criterion is wrong**
(df_residual = n − k − Σ(Gᵢ − 1), alongside the intercept adjustment).

> **Open decision — df plumbing:** optional `[DF_Absorbed]` argument threaded through
> the inference chain vs. a wrapper set. Current lean: the optional argument,
> consistent with one-source-of-truth and the rejection of parallel function sets.

### Model Spec status block — the transparency price, paid visibly

Construction inside a LAMBDA means the user can no longer *see* the design matrix by
scrolling. The replacement is a fixed-height block answering "what model did I
actually specify?":

- **Response in effect** (derived), and the Response-count validation: exactly one,
  error at zero or two-plus
- Constructed column count k — reconcilable against the Σ(design columns)
  audit identity
- Level-qualified constructed column names
- **Included row count** after the effective mask (auto-completeness AND Filters),
  with the active Filter columns listed
- **Degenerate Categorical Predictors** (L ≤ 1 in the analysis sample): listed by
  name; the constructor contributes zero columns so the rest of the model still
  computes — visible degradation (red F cell + status line), not a hard error, not
  silent omission
- References in effect for each Categorical (surfaced even when defaulted)
- Intercept status; *(v2.1)* active FE variables, group counts, absorbed df
- **Error state** — a single visible cell flagging illegal specs: zero or multiple
  Response rows; invalid reference level; `Continuous` on a non-numeric column;
  no-intercept with a Categorical included (CF flag + cell comment, not a hard
  error); *(v2.1)* more than one FE variable (two-way absorption is a post-v2.1
  milestone)

**Intercept × FE interaction (v2.1) — flag, don't force.** An intercept on demeaned
data estimates ≈ 0 and wastes a df — not catastrophic — so silently forcing
`Allow_Intercept` FALSE would be silent reinterpretation. Flag red and instruct.

### Zone changes on the Regression sheet (breaking — hence MAJOR)

- **Spec block:** A–B → A–K (Role, Include, Type, Reference, Order, Transform,
  Sequence, Base Period Δ, Levels, Reference In Use — Order/Transform/Base
  Period Δ reserved, unread by any formula; Sequence added post-v2.0 as the
  structural axis), spanning all table columns, with cascading-relevance CF.
- **Predictor summary changes referent:** Pearson/Spearman/Skewness/Kurtosis/VIF/
  Tolerance run on the **constructed** columns — more correct anyway (VIF on the
  actual design matrix, dummies included). Block height = constructed width.
- **Coefficient zone:** names become level-qualified via the constructor twin; table
  height is a computed property of the spec.
- **Prediction inputs:** raw values (GDP = 5000, Status = "Developing"), Categorical
  inputs validated against `Dummy_Levels`, encoded through the same code path as
  training.
- **Residual output:** structure unchanged; Identifier columns become available as
  row labels. *(v2.1: residuals are within-model residuals under FE — Diagnostic
  Guide gains a paragraph.)*

**Layout principle check (fixed-width left, fixed-height top):** the A–K spec block
is fixed-width; the status block fixed-height; the constructed matrix never appears
on the sheet. Note the estimator choice is load-bearing: LSDV would spill a
data-dependent number of columns; the within transformation (v2.1) replaces them
with the same k columns transformed.

### New catalog functions (Specification & Design Matrix subcategory)

Working list; canonical names per the Naming Convention:

- Sheet-scoped constructors: `x_s()`, `Constructed_Column_Names()`, the derived
  Response and effective-mask LAMBDAs, and *(v2.1)* `y_s()`
- `Role_Aware_Complete_Cases(...)` — the automatic completeness layer of the
  effective mask (successor to the hard-coded `Full_Data` formula)
- `Absorbed_Degrees_Of_Freedom(...)` *(v2.1)* — Σ(Gᵢ − 1) from the spec
- `Dummy_Column(category, level, [include])` — one indicator column per explicit
  call; complements `Dummy_Code`
- Spec-validation helpers backing the status-block error states

### Demonstration walkthroughs (Instructions sheet, WHO data)

- **Factor walkthrough:** Status as a Categorical Predictor — the pooled vs.
  by-Status coefficient flip (the documented Simpson's Paradox risk), with the
  stratified comparison driven by a Filter column.
- **Panel walkthrough** *(v2.1)*: Country (and optionally Year) flipped from
  Identifier to Fixed Effects — directly addressing the panel structure already
  identified as violating OLS independence.

**Superseded, not reused.** `Dummy_Levels` and `Dummy_Code` already exist as catalog
functions (added independently of this sheet's build) but are not yet referenced by
any sheet writer. **Decision made: drop both and rebuild from scratch** rather than
amend or defensively wrap — see the resolved decision below. `X_s()` depends on the
rebuilt `Dummy_Levels`; `Dummy_Code` is rebuilt to the same standard for standalone
free-form use even though `X_s()` does not call it directly (it encodes inline via
broadcast, per the level-vector split).

### Resolved decisions *(v2.1)*

1. **df plumbing — RESOLVED: optional argument.** The absorbed-df correction is an
   optional `[DF_Absorbed]` argument defaulting to 0, threaded through the df /
   MS-residual / t-critical chain. Because it defaults to 0, every no-FE model
   computes exactly as under v2.0 — this is what keeps Fixed Effects a non-breaking
   MINOR (see the ladder's breakage flag, now resolved to "No"). It is an *engine
   signature addition*, not a sheet restructure: existing formulas that omit the
   argument are unaffected.

2. **FE point prediction — RESOLVED: group-mean recovery form.** Rather than
   materialize the G group intercepts the within estimator discards, prediction uses
   the algebraic identity α̂ᵢ = ȳᵢ − x̄ᵢ′β̂ substituted back in, giving

   > ŷ = ȳᵢ + (x_new − x̄ᵢ)′β̂

   — the selected group's mean response, adjusted for how far the new covariates sit
   from that group's mean covariates. Requires only three new group-keyed summaries
   for the *one* selected group: ȳᵢ (`AVERAGEIFS` on the response), x̄ᵢ (`AVERAGEIFS`
   per predictor), Tᵢ (`COUNTIFS`) — all respecting the same Include/Filter mask as
   the fit. Group selection is a **data-validation dropdown sourced from the observed
   level list** (the spilled `Dummy_Levels` output via a `#` spill reference / named
   range), which also enforces the hard constraint that α̂ᵢ exists only for groups the
   model actually saw. A degenerate G = 1 (one "group" = whole sample) collapses this
   exactly to the ordinary v1.0 prediction — see "Redesign impact" below, this is the
   key to building it once.

3. **Prediction interval — RESOLVED: surface BOTH mean-CI and new-observation-PI.**
   Same center, differing by one variance term:

   > Var(mean) = σ²/Tᵢ + (x_new − x̄ᵢ)′ V_β (x_new − x̄ᵢ)
   > Var(new)  = σ²·(1 + 1/Tᵢ) + (x_new − x̄ᵢ)′ V_β (x_new − x̄ᵢ)

   with σ² = MS-residual on absorbed df (N − G − K), V_β the coefficient covariance
   already computed for inference, and the interval ŷ ± t(N−G−K)·√Var. The quadratic
   form reuses the existing v1.0 prediction-leverage machinery, fed the *deviation*
   (x_new − x̄ᵢ) instead of x_new. The interval is **group-specific in both center and
   width**: σ² and V_β are pooled, but Tᵢ (via 1/Tᵢ) and x̄ᵢ (via the deviation) change
   with the selected group — so changing the group dropdown re-computes uncertainty,
   not just the point estimate. Present three lines: point · mean-response CI (low/high)
   · new-observation PI (low/high). *Sanity check for the test plan:* predicting at the
   group's own centroid (x_new = x̄ᵢ) kills the quadratic term and the mean-CI collapses
   to t·√(σ²/Tᵢ), the standard error of ȳᵢ.

4. **Prediction input location — RESOLVED: on the regression sheet, in place.** Inputs
   stay local to each model sheet (making each sheet self-contained and able to predict
   without the Model Comparison sheet), with validation-list dropdowns for categorical
   predictors and the FE group. This does not foreclose v2.3: whether those local cells
   later become `XLOOKUP` formulas pointing at the shared Model-Comparison inputs is a
   v2.3-only decision and is not forced by building the local version now.

5. **Scope — RESOLVED: one-way FE only, iid errors, existing groups only.** v2.1 is a
   one-way release in its entirety, not just for prediction: a spec with two or more
   Fixed Effects variables is a visible status-block error, and the two-way machinery
   (`Absorb_Two_Way_Fixed_Effects`, `Demean_Two_Way_Balanced`,
   `Fixed_Effects_Convergence_Check`) is deferred to its own post-v2.1 milestone (see
   v2.5+). The ȳᵢ recovery is valid for a single grouping dimension. Two-way FE
   (Country × Year) does not recover intercepts as a simple group mean and is
   **explicitly out of v2.1 scope** — flagged so the clean formula is not silently
   misapplied. The
   interval assumes homoskedastic, non-serially-correlated errors (the classic FE
   assumption); clustered/robust SEs are out of v2.1 — ship the iid interval with a
   documented caveat rather than implying robustness. Overlaps the Durbin-Watson-under-FE
   item below.

### Open items (recorded, not resolved)

1. **Durbin-Watson under FE** *(v2.1)* — **RESOLVED by the BFN release as "second
   cell + mutual gating"** (neither relabel nor suppress). (Related to the iid caveat
   above.) The DW-gate release had already removed the physical-row-order hazard
   (`Durbin_Watson_By` sorts by `Sequence_Column()` before differencing); the BFN
   release adds the group-aware panel statistic: `BFN_Panel_Durbin_Watson`
   (Bhargava–Franzini–Narendranathan 1982), BFN = Σᵢ Σₜ₌₂ (û(i,t) − û(i,t−Δ))² ÷
   Σᵢ Σₜ û(i,t)², with the numerator's differencing restricted to within-group
   (group, seq−Δ) pairs via `Difference_By` (one source of truth — first periods and
   panel gaps contribute no term, seams cannot manufacture correlation, and the
   statistic is invariant to permuting group blocks). Two fixed diagnostic cells,
   each self-guarding: no Sequence → both `n/a — requires Sequence`; Sequence + no
   FE → DW active, BFN `n/a — no fixed effects`; Sequence + FE → BFN active, DW
   `n/a — FE active`. FE detection counts Role="Fixed Effects" spec rows — forward
   wiring that activates when the v2.1 FE role lands in the dropdown. **Remaining
   open sub-item: BFN critical values.** The cell ships with an interpretation
   caveat only (near 2 ⇒ no first-order autocorrelation in the within residuals);
   its significance bounds are N,T-dependent (Bhargava et al. 1982 tables) and the
   standard DW bounds must not be presented next to it — surfacing proper BFN
   bounds is the recorded open item.
2. **Categorical × FE prediction encoding** *(v2.1)* — moved to the v2.1
   Pending block (item #11 in `TODOs.md`). When non-FE categorical predictors
   coexist with fixed effects, x_new and x̄ᵢ must be formed in the *constructed*
   design-matrix space (dummies encoded through the same `Dummy_Code` path
   `x_s()` uses), not raw input space. Shares machinery with v2.0 categorical
   prediction, so largely subsumed — recorded so the encoding step is not
   forgotten.
3. **Dataset bundling** (carried forward) — bundled WHO vs. data-agnostic. The
   `Source_Data` indirection reduces changeover to a one-name edit, which
   strengthens the data-agnostic option; the named-range rename it gated is
   partially subsumed by the v2.0 name changeover itself.
4. **Auto-completeness implementation — RESOLVED by construction, one caveat open.**
   `Sample_Include()` shipped with the role-aware completeness layer built in (every
   Filter column truthy AND the Response numeric AND every included Continuous
   Predictor numeric), so no separate `Role_Aware_Complete_Cases` function was needed.
   Remaining gap (the **blank-categorical caveat**): Categorical Predictors impose no
   non-blank condition; a blank category value encodes as all-zero dummies. Tracked
   in `TODOs.md` (v2.0 leftovers) and verified in
   `HUMAN_TEST_PLAN_v3_model_construction.md` "Known caveat to verify and accept (or
   escalate)" — the decision is to accept the caveat as documented v2.0 behavior
   (interim: declare a completeness column as a Filter).

### Regression sheet redesign impact — read this before scheduling

The one question worth being precise about: **the current (v1.0) Regression sheet is
redesigned exactly once, at v2.0.** That is the only breaking restructure — control
block → declarative spec block, and `x_s()` from column-filter to model-matrix
constructor. None of the v2.1/v2.2 decisions above independently forces a *second*
redesign; each is either an additive section on the already-restructured v2.0 sheet or
a backward-compatible engine signature addition:

| Decision | Touches the sheet? | Redesign or additive? |
|---|---|---|
| `[DF_Absorbed]` optional arg (v2.1) | One new status-block cell (absorbed df), wired only when FE active | **Additive** — engine signature change, default-0 leaves no-FE models identical |
| FE point prediction (v2.1) | Group dropdown + ȳᵢ/x̄ᵢ/Tᵢ cells in the prediction zone | **Additive** *if the v2.0 prediction zone is built in the general form* — see below |
| Both CI and PI (v2.1) | Two interval lines instead of one in the prediction zone | **Additive** — and applies to plain OLS too (see the one v1.0-relevant note below) |
| Transforms / unit-space (v2.2) | New unit-space block + `Transform` column G already reserved in v2.0 | **Additive** — column G shipped reserved precisely so this needs no restructure |

**The one forethought item that prevents a second redesign.** Because the degenerate
G = 1 case of ŷ = ȳᵢ + (x_new − x̄ᵢ)′β̂ collapses exactly to the ordinary OLS prediction
(with 1/Tᵢ = 1/N recovering the textbook 1 + 1/n + leverage term), the **v2.0 prediction
zone should be built in this general group-mean form from the start**, with the whole
sample as a single implicit group. Then v2.1 Fixed Effects is literally just "let G > 1"
— the group dropdown and `AVERAGEIFS` keys activate, but the formula structure is
untouched. Building v2.0's prediction in the naive x_new′β̂-only form would force tearing
it up at v2.1; building it in the general form makes FE prediction a genuine addition.
This is the single most important sequencing note for the v2.0 implementer.

> **Post-ship note:** v2.0 shipped the prediction zone in the standard
> `Prediction_Interval` form rather than the general group-mean form, and kept the
> single-interval output rather than a CI + PI pair. Both items move to the v2.1 work
> list (see `TODOs.md` under v2.1) — v2.1 rebuilds the prediction zone instead of
> merely activating it. The arithmetic consequence is unchanged; the cost is one
> rebuild at v2.1.

---

## v2.1 — Sequence, fixed effects, and the forward-wiring chain — IN PROGRESS

The 2.1 milestone bundles three coherent pieces that share the Sequence axis
and the FE Role: the Sequence/Base Period/longitudinal/serial-correlation chain
that the v2.0 work record says is "the v2.1 work that doesn't need the FE
engine proper" (now substantially complete in the repo), the FE engine
proper (`y_s`, `[DF_Absorbed]`, `Demean_By`, `Group_Mean`,
`Absorbed_Degrees_Of_Freedom`, `Is_Balanced_Panel`), and the sheet work that
activates the engine (FE Role dropdown, status-block validation, CI+PI
prediction layout, FE group dropdown, BFN cell flips active when FE is set).
Two-way FE remains a post-2.1 milestone (see v2.5+).

The 2.1.0 release ships as **a single release**, not as a sequence of preview
builds — the FE Role dropdown, the CI+PI prediction layout, and the FE
activation on the sheet are all gated to ship with the engine, so users never
see a workbook that says "FE is in the dropdown but the engine is forthcoming."

### Shipped (since v2.0.0)

- ~~Sequence structural axis (spec column H) and reserved Base Period Δ (column I).~~ **DONE (PR #101).** Spec block grew A–I to A–K; zero-or-one validation at H2; read only by the validation layer and the base-period layer.
- ~~Gap-aware `Difference_By` / `Lag_By` and the Base Period Δ layer.~~ **DONE (PR #102).** Within-group differencing via exact `(group, seq−Δ)` lookups; NA() at first periods and panel gaps; Sequence Spacing block (rows 28–34); parser supports zero-parameter LAMBDAs so `Base_Period_Delta()` survives the workbook.xml sync.
- ~~Sequence-aware Durbin-Watson.~~ **DONE (PR #103).** `Durbin_Watson_By(X_s, Y, seq, [Allow_Intercept], [Include])`; gated cell at Regression X11/Y11.
- ~~BFN panel Durbin-Watson.~~ **DONE (PR #105).** `BFN_Panel_Durbin_Watson` (Bhargava–Franzini–Narendranathan 1982) with the Sequence/FE trigger matrix; cell at X12/Y12.
- ~~Grouping-key resolver.~~ **DONE (PR #106).** `Serial_Correlation_Group()` SWITCH with the dormant `Cluster` branch (the v2.6+ Cluster role is a resolver-only edit).
- ~~v1.1 leftovers — histogram distribution overlays and per-distribution Q-Q plots.~~ **DONE (PRs #96, #97, #99, #100).** Eight per-distribution Q-Q charts with closed-form quantile inverses; histogram combo charts with overlay lines on the shared count axis; BetaPERT singularity fix; 4×2 Q-Q layout.
- ~~`--skip-univariate` CLI option.~~ **DONE (PR #98).** Build ergonomics for the slow grid-search step.
- ~~Spec-driven QC refactor.~~ **DONE (PR #103).** `analyze_regression_spec.py` and `test_regression_spec_qc.py` replace the legacy MLR smoke-test-sheet path; `Full_Data` ships as Omit in the default spec; Year flagged as Sequence, Population as Omit demonstrator.

### Pending (in ship order; #1 prerequisite, #2/#3/#9 gated to #5)

The full ship order lives in `TODOs.md` under the v2.1 section. The high-level
shape: the Sequence fix lands first (it's a real sheet-writer change with
significant testing, not a one-liner), then the FE Role dropdown + status
block validation, then the CI+PI prediction layout, then the engine
(`Demean_By` / `Group_Mean` / `Is_Balanced_Panel` / `Absorbed_Degrees_Of_Freedom`
/ `y_s` / `[DF_Absorbed]`), then the FE group dropdown + ȳᵢ/x̄ᵢ/Tᵢ cells +
BFN cell flips active. The follow-on polish (BFN critical values, Categorical
× FE prediction encoding, residual relabel + Diagnostic Guide) ships with
2.1.0 if there's room.

**#1 — Sequence axis auto-detection and override (prerequisite).** The
override mechanic has a spill-collision risk: the spec block reads its own H/I
cells, and a source table wider than the shipped WHO sample could let the
override spill overrun an input band. Fix: rename column I to
**`Sequence Period`**, add column J **`Period In Use`** following the
Reference Level / Reference In Use pattern (displays the candidate if no
override, the typed value if the override is non-blank), bound every read of
the H/I/J band by `COLUMNS(Source_Data)` (the spill-placement principle from
`CLAUDE.md`), add yellow CF on `Period In Use` when overridden, red CF when
off-grid, and per-row CF on the override cell. Update the Sequence Spacing
block (rows 28–34), the spec layout constants, and the QC analyzers.

**#2 — FE Role dropdown + status-block validation (gated).** Sheet-only, no
engine change. `Fixed Effects` in the Role axis; status-block cells for
"active FE variable," "group count," and "absorbed df" — the last two show
`n/a — engine forthcoming` until the engine lands. Visible error at 2+ FE
variables; intercept × FE red flag; CF bands update.

**#3 — Surface BOTH intervals in adjacent cells of the prediction outputs
section (gated).** Three lines: point · CI low/high · PI low/high. The PI
half-width is `√(σ²·(1+1/T) + q)` and the CI half-width is `√(σ²/T + q)`;
same center, same x_new, same t-critical. Sheet layout only — the math lives
in the engine. The FE engine drops in group-keyed inputs at the activation
step without restructuring this layout.

**#4–#8 — Engine work.** `Demean_By` / `Group_Mean` (constructor internals
and user-callable transforms) and `Is_Balanced_Panel` (one-way diagnostic)
ship together; `Absorbed_Degrees_Of_Freedom(spec)`; `y_s()` (the
demeaned-Response constructor, new function, not a replacement); and the
optional `[DF_Absorbed]` argument (default 0) threaded through df /
MS-residual / t-critical. Because `[DF_Absorbed]` defaults to 0, no-FE
models compute identically to v2.0 — this is what keeps v2.1 non-breaking
under the interface definition. **Significant testing on the engine PR:**
assert bit-equality of every existing engine test with and without the
argument, plus FE-active cases for the full inferential chain (SE, t, p,
CIs, AIC/BIC).

**#9 — FE group selection dropdown + ȳᵢ / x̄ᵢ / Tᵢ cells + BFN cell flips
active (gated).** AVERAGEIFS/COUNTIFS respecting the Include/Filter mask;
the prediction outputs section from #3 activates the group-mean form
(ȳᵢ + (x_new − x̄ᵢ)′β̂) with group-keyed inputs; BFN cell (X12/Y12) flips
from `n/a — no fixed effects` to active when FE is set. **#2, #3, and #9
ship as one release with the engine.**

**#10–#12 — Follow-on polish (ships with 2.1.0 if there's room).** BFN
critical values (N,T-dependent per Bhargava et al. 1982 tables — do NOT
present standard DW bounds next to the BFN cell); Categorical × FE
prediction encoding (x_new and x̄ᵢ formed in constructed design-matrix
space, UI wire to encode through `Dummy_Code`); and the residual-output
relabel plus a Diagnostic Guide paragraph on residuals under FE.

---

## Future roles (Role-axis values, v2.1+)

The Variable Role axis is extensible by design — these are **Role values, not new
mechanisms**, and the Predictor Type axis never grows:

- **`Fixed Effects`** *(v2.1)* — the panel role: enters no column; the entire design
  matrix and the Response are demeaned by its groups (one FE variable → `Demean_By`;
  two-plus → visible error — two-way absorption via `Absorb_Two_Way_Fixed_Effects` is
  its own post-v2.1 milestone, see v2.5+). Brings `y_s()`, absorbed df, and the
  within-model output relabeling.
- **`Weight`** — WLS. Supersedes the standalone WLS milestone and its
  argument-threading debate. Three-stage scope carried forward: user-supplied
  weights → variance-driver-derived weights → FGLS.
- **`Time`** — designates the time index, feeding `Lag_By` / `Difference_By` and
  future serial-correlation diagnostics.
- **`Cluster`** *(later candidate)* — clustered standard errors, the honest endgame
  for panel inference. Substantial engine work; out of scope until after FE ships.
  Forward wiring already in place: `Serial_Correlation_Group()` — the grouping-key
  resolver the BFN panel DW consumes instead of referencing the FE column directly —
  carries a dormant, unreachable Cluster branch in its SWITCH (returns the
  `RESERVED — v2.6+` token; the reserved-spec-column pattern), so supplying the
  grouping key from a Cluster role for pooled-panel diagnostics without absorption
  is a resolver-only edit.

Cardinality constraints when added: at most one `Weight`, at most one `Time`; the
status block validates (same pattern as exactly-one-Response).

---

## Data Transformation

Cross-cutting infrastructure as a *taxonomy*: these functions serve **double duty** —
internals of the spec-driven constructor, and standalone user-callable transforms for
free-form work on the data sheet. Tracked as its own catalog category, separate from
the version ladder (see *Function Categories*).

**Delivery, however, is pinned to the ladder** (none of these standalone functions is
built yet): the user-callable transform library ships at **v2.2** alongside the
column-G wiring, with two exceptions — `Demean_By` and `Group_Mean` ship at **v2.1**
as Fixed-Effects internals, and the two-way functions
(`Absorb_Two_Way_Fixed_Effects`, `Demean_Two_Way_Balanced`,
`Fixed_Effects_Convergence_Check`) follow the **two-way FE milestone (post-v2.1)**.

Row-aligned transforms accept an optional `include` mask (`1`/`TRUE` keeps the row,
`0`/`FALSE` excludes it). When `include` is omitted, those functions construct a default
mask from the required inputs. Excluded rows return `""` so spilled arrays stay aligned
with the source rows — `""` was chosen deliberately over `NA()` because it round-trips
cleanly through `ISNUMBER`-based keep logic elsewhere in the library (`ISNUMBER("")` is
`FALSE`, so a downstream mask built on `ISNUMBER` correctly re-excludes it) without
erroring inside `SUM`/`AVERAGE` the way a propagated error value would.

**Note on naming:** these functions are new, so they're written directly against the
Naming Convention above (no abbreviations) rather than needing a retrofit.

### Sample Construction & Diagnostics — *subcategory*

- `Numeric_Complete_Cases(data)` — listwise-deletion sample mask; `1` when every value in
  a row is numeric.
- `Is_Balanced_Panel(group, time, [include])` — `TRUE` only when every included group has
  exactly one observation for every included time period.
- `Fixed_Effects_Convergence_Check(x, group1, group2, [include])` — largest absolute
  remaining group mean across two fixed-effect dimensions; near zero indicates
  convergence after `Absorb_Two_Way_Fixed_Effects`. Surfaced in the status block
  whenever two Fixed Effects variables are active *(two-way FE milestone, post-v2.1)*.

### Location & Scale — *subcategory*

- `Center(x, [include])` — grand-mean centering, \(x_i - \bar{x}\).
- `Zscore(x, [include])` — standardization via `STDEV.S`, \((x_i - \bar{x}) / s_x\).
- `Minmax_Scale(x, [include])` — scales to \([0, 1]\).
- `Winsorize(x, [lower_p], [upper_p], [include])` — caps values outside selected
  percentiles (default 1st/99th). Remains an explicit modeling decision, never an
  automatic preprocessing step.
- `Ln_Positive(x, [include])` — natural log, restricted to strictly positive numeric
  values; returns `""` rather than a worksheet error for zero, negative, or non-numeric
  input.

### Group & Panel — *subcategory*

- `Group_Mean(x, group, [include])` — matching group mean on every included row.
- `Demean_By(x, group, [include])` — one-way within transformation, \(x_{ig} - \bar{x}_g\).
  **v2.1 constructor internal** for a single Fixed Effects variable.
- `Zscore_By(x, group, [include])` — within-group standardization.
- `Decompose_By(x, group, [include])` — returns the between-group mean and within-group
  deviation as two columns, exposing \(x_{ig} = \bar{x}_g + (x_{ig} - \bar{x}_g)\).
- `Demean_Two_Way_Balanced(x, group1, group2, [include])` — direct two-way demeaning,
  exact only for a balanced panel; check with `Is_Balanced_Panel` first. *(Two-way FE
  milestone, post-v2.1.)*
- `Absorb_Two_Way_Fixed_Effects(x, group1, group2, [include], [passes])` — iterative
  alternating-projection demeaning for unbalanced panels. Convergence is not verified
  internally; always pair with `Fixed_Effects_Convergence_Check`. **Constructor
  internal for two Fixed Effects variables at the two-way FE milestone (post-v2.1).**

### Categorical & Model Construction — *subcategory*

- `Dummy_Levels(category, [reference], [include])` — **rebuilt for v2.0** (an earlier
  version existed as a catalog function with string-based error returns; dropped and
  replaced rather than amended — see the v2.0 rebuild PR). Signals failure via a real
  Excel error (`NA()`), never a descriptive string, so every downstream
  `IFERROR`/`ISNA` guard works without special-casing. Retained categorical levels as
  a horizontal header row; backs the v2.0 prediction-input validation lists and is
  the hard dependency of `X_s()`'s level-vector split.
- `Dummy_Code(category, [reference], [include])` — **rebuilt for v2.0** alongside
  `Dummy_Levels`, calling it internally for level determination (one source of truth,
  same NA()-based error contract). Dummy-coded matrix. Use a reference
  level (treatment coding) when the design includes an intercept; full one-hot coding
  plus an intercept causes perfect multicollinearity. **Reference-level validation
  (confirming the requested reference actually exists in the included sample) is
  required at implementation, not deferred** — an invalid reference must error
  (`NA()`), never silently fail to drop a column and reintroduce the exact
  collinearity the function exists to prevent. Standalone; `X_s()` does not call it
  directly (it encodes inline via broadcast) but is held to the same standard.
  **v2.0 constructor internal** for Categorical roles, via `Dummy_Levels`.
- `Dummy_Column(category, level, [include])` — single indicator column per explicit
  call (see v2.0 catalog additions).
- `Interact(x1, x2)` — elementwise product \(x_1 x_2\); broadcasts across dummy-coded
  matrices to produce one interaction column per retained level.
- `Model_Matrix(X, [add_intercept])` — optionally prepends an intercept column.
  Intentionally not variadic — predictors are assembled explicitly with `HSTACK` so the
  specification stays visible and auditable.

### Longitudinal & Panel-Time — *subcategory*

**Shipped early, in the base-period release** (ahead of the v2.2 bundle), to
the gap-aware semantics recorded in the v2.1 design record §1:

- `Lag_By(x, group, seq, [delta], [include])` — prior-period value within the
  same group, keyed on `group`/`seq` **by exact time value, not physical row
  order** (exact-match lookup of `(group, seq − Δ)` pairs, never OFFSET/row
  arithmetic). A gap — `seq − Δ` absent within the group — returns `NA()`,
  never the previous available row.
- `Difference_By(x, group, seq, [delta], [include])` — within-group time
  difference \(\Delta x_{it} = x_{it} - x_{i,t-\Delta}\), delegating the pair
  lookup to `Lag_By`. Each group's first period returns `NA()` — never a
  fabricated 0 that would enter a design matrix silently.
- `Base_Period_Delta()` — the Δ in effect: spec column I on the
  Sequence-flagged row (computed candidate or typed override). The omitted-
  `[delta]` default of both functions above — Δ is never a silent 1; with no
  declared axis they return `NA()` everywhere (visible failure). Companion
  sheet-scoped closures `Sequence_Deltas` / `Base_Period_Delta_Candidate` /
  `Sequence_Delta_Spectrum` drive the spec's Sequence Spacing block.

Exception to the row-aligned `""` convention, recorded: an *included* row
whose difference is incomputable (first period, gap, non-numeric value)
returns `NA()`, not `""` — it is a visible incomputable observation, not an
excluded row. `""` remains the excluded-row return. Verification:
`tests/test_difference_by_verification.py` (WHO exact counts plus the
punched-out-year and calendar-date synthetic cases); human test plan T17–T19.

---

## v2.2 — Transforms & Unit-Space Comparability

**Why this is its own milestone, not a Model Comparison sub-feature:** the Model
Comparison sheet is only an honest comparison tool if the numbers it lines up mean the
same thing. An R² computed on `Ln(Life expectancy)` and an R² computed on raw
`Life expectancy` are not the same quantity, and putting them in adjacent cells of a
comparison table without correction is exactly the kind of silent misfiring the
library's design philosophy exists to prevent. Transforms is the release that makes
cross-model comparison trustworthy, not just possible — and it is what completes the
Regression sheet as a fully functional deliverable, which is why it precedes the
Model Comparison convenience layer.

**This is also what finally wires the reserved spec-block column G** (`Transform`),
present since v2.0 but explicitly unread by every v2.0 formula (`X_s()`,
`Constructed_Column_Names()`, `Row_Labels()`, `Sample_Include()`). v2.0's
"reserved, not implemented" comment on that column was written for this release.

**Why this is a MINOR (non-breaking).** The `Transform` column defaults to `None`, and
the existing OLS MLR formula names remain the unit-space defaults — a `None`-transform
model produces output identical to v2.1. New behavior is reached only by a user setting
a `Transform` value they previously could not; no existing specification changes
meaning. This satisfies the non-breaking criterion in Versioning & Release Conventions.

### Scope (v2.2.0 — Log only, plus the standalone transform library)

- `Transform` dropdown gains one real value: `Log`, in addition to `None`. Applicable to
  a Continuous Response and/or Continuous Predictors.
- **The standalone Data Transformation function library ships in this release** (see
  the *Data Transformation* section for the full specs): `Center`, `Zscore`,
  `Minmax_Scale`, `Winsorize`, `Ln_Positive`, `Zscore_By`, `Decompose_By`,
  `Numeric_Complete_Cases`, `Dummy_Column`, `Interact`, `Model_Matrix`.
  (`Demean_By`/`Group_Mean` arrive at v2.1 as FE internals; `Lag_By`/
  `Difference_By`/`Base_Period_Delta` shipped early in the base-period release;
  the two-way functions follow the two-way FE milestone.)
- **OLS MLR formula names stay the unit-space defaults** — `R_Squared`, `AICc`, etc. are
  not renamed or forked; they continue to describe whatever space the model was actually
  fit in. New functions are added *alongside* them for the unit-space (back-transformed)
  view.

### Unit-space fit statistics — naming

**RESOLVED: dispatcher function, with arguments in `(model, response_transform, predictor_transform)` order.**

```
Unit_Space_R_Squared(model, response_transform, predictor_transform)
Unit_Space_Adjusted_R_Squared(model, response_transform, predictor_transform)
Unit_Space_RMSE(model, response_transform, predictor_transform)
```

The dispatcher internally `SWITCH`es on the `(response_transform, predictor_transform)` pair.
For Log at v2.2.0, the four branches are: `(None, None)`, `(Log, None)`, `(None, Log)`,
`(Log, Log)` — plus an `else` branch that returns `NA()` for any unrecognised transform
value, consistent with the library's `NA()`-based error-signaling convention.

The argument order is deliberate. `model` first (the heavy argument — the design matrix,
the response, the coefficients), then the two transform descriptors in `response →
predictor` order, matching the spec-block's column-G reading order
(`Spec_Transform` on the Response row first, then on the Predictor rows). When the
Regression sheet's unit-space block builds the dispatcher call, the arguments map
1:1 to that reading order; no reordering at the sheet layer.

**Why dispatcher, not per-combination names.** A per-combination scheme
(`Log_Level_Unit_Space_R_Squared`, `Level_Log_Unit_Space_R_Squared`, etc.) scales as
O(N²) per statistic and O(N² × M) across the M statistics with unit-space
counterparts — and the rename pain only gets worse with each new transform
(square-root, Box-Cox, …). The dispatcher scales as O(N²) in SWITCH branches
*regardless of how many statistics get a unit-space counterpart*: adding a new
unit-space statistic is one new dispatcher, not N² new named functions. The
SWITCH body is auditable in one place per statistic, and the spec block's
existing `SWITCH(Spec_Transform, …)` pattern on column G is the precedent.

**Naming-style departure, recorded.** The dispatcher is a bigger departure from
the per-statistic-per-shape naming style used everywhere else in the catalog.
Worth flagging because every other function in the library follows the
"one canonical name, one LAMBDA" rule; the dispatcher is the first deliberate
exception, and the exception is justified by the combinatorial blow-up the
exception avoids. Pattern: when a family is closed-form (one shape per name),
use the per-shape style; when a family is combinatorial in its inputs (N
transforms × M statistics), use a dispatcher. Future combinatorially-named
families (e.g. a future `Cross_Product_Of_Transforms_*` if one ever exists)
follow the same exception.

**Statistics with a unit-space counterpart (v2.2.0):** R², Adjusted R², RMSE.
**Statistics flagged but deferred:** AIC, AICc, BIC — likelihood-based; comparing
them across differently-transformed responses is a separate, harder question
(likelihood depends on the Jacobian of the transformation, and the "right"
comparison is on the original response's likelihood evaluated at the
back-transformed prediction, not the transformed response's likelihood).
Tracked as an open item — *not* shipped at v2.2.0 even via a dispatcher,
because the answer is not yet a single SWITCH branch.

### Unit-space section on the Regression sheet

A new fixed-height block, alongside the existing GoF stats, containing SWITCH formulas
that read the Response/Predictor `Transform` values out of spec column G and select the
matching unit-space lambda — so the sheet always surfaces one "headline" comparable
statistic regardless of which transform is active, and that is the value the Model
Comparison sheet's GoF table references (not the raw-space statistic) once this ships.

### Prediction under transforms

**RESOLVED: ship Duan's smearing estimator as the default, with a per-cell toggle for naive `EXP()`.**

- Predicted values need back-transformation (e.g. `EXP()` for a logged response) before
  they are comparable in the Model Comparison sheet's prediction results table.
- **Why Duan, not naive:** naive exponentiation of a log-linear prediction is a biased
  estimator of the conditional mean response (Jensen's inequality: E[exp(X)] ≠ exp(E[X])
  whenever X has variance). Duan (1983) smearing — ŷ_smeared = ŷ_log · mean(exp(residuals)) —
  is unbiased under the iid-residual assumption and adds one extra `AVERAGE(EXP(residuals))`
  cell. For a cost-estimator audience the naive-vs-Duan difference is not a footnote; it
  is a number-vs-number question they will see and ask about.
- **Default = Duan; per-cell toggle for naive:** the prediction outputs section exposes
  a single input cell (e.g. `Back_Transform_Method` at a known coordinate) with values
  `Duan` (default) and `Naive` (the documented-biased alternative). The toggle is a
  per-sheet user preference — a user who wants the textbook `EXP()` form for a
  homework-style check can flip it; the default ships at Duan for honest reporting.
  Form: `IF(method="Naive", EXP(ŷ_log), EXP(ŷ_log) * AVERAGE(EXP(residuals)))`.
- **Caveat row on the sheet:** regardless of which is selected, the prediction outputs
  block carries a one-line caveat under the value:
  *Duan = Duan (1983) smearing; Naive = textbook EXP(ŷ), biased.*
  Visible from day one so the choice is auditable, not implicit.
- **Reference:** Duan, N. (1983). "Smearing Estimate: A Nonparametric Retransformation
  Method." *Journal of the American Statistical Association*, 78(383), 605–610.

### Sequencing note

Because Transforms (v2.2) ships **before** Model Comparison (v2.3), the "headline"
unit-space GoF cell each Regression sheet exposes is already correct by the time the
comparison sheet is built — Model Comparison simply points its GoF table at that cell and
inherits unit-space-honest comparisons for free, with no transitional window in which
non-comparable numbers get lined up. The completion of the core deliverable removes a
correctness hazard from the later convenience layer before it can ever surface.

---

## v2.3 — Model Comparison Sheet

**The central idea:** every v2.0 Regression sheet already exposes a fixed-height,
fixed-position **Model Spec status block** (response in effect, constructed column
count, level-qualified names, included row count, error state). That block is an
interface, not just a display — the Model Comparison sheet is what happens when a second
sheet is allowed to *read* it. No new modeling capability is added; this is purely a
cross-sheet aggregation and navigation layer, which is why it is a MINOR: it is
read-only against finished Regression sheets and changes no existing input.

### `Model_Formula_String` (name RESOLVED)

**RESOLVED name: `Model_Formula_String(anchor_cell)`.** Considered alternatives
(`Regression_Model_Spec_String`, `Regression_Spec_Label`) and rejected them:
`Regression_*` prefixes couple the function name to a particular sheet kind,
which is wrong because the same string format is meaningful for any model
sheet (and the same anchor-validation pattern would extend to a future
GLM sheet, time-series sheet, etc.). `Spec_Label` is vague — could be any
text label. `Model_Formula_String` is self-describing (a model, in
formula-string form) and matches R's "model formula" terminology.

A workbook-scoped LAMBDA that reads a Regression sheet's spec block and returns a
human-readable model formula string, e.g. `Life expectancy ~ GDP, Schooling, Status`.

- **Simplest case:** `y & " ~ " & TEXTJOIN(", ", x_s_names, TRUE)`, where `y` and
  `x_s_names` are the spec-block Column A values for the Response row and every included
  Predictor row (Role = Predictor, Include = TRUE) — Continuous *and* Categorical alike.
  This mirrors R's model-formula convention: a formula lists the variable name
  regardless of type, it does not expand categorical levels inline. No special-casing is
  needed to include categoricals in the first version of this function.
- **Validation:** before building the string, confirm the target is actually a Regression
  sheet by checking the spec-block header signature (the same kind of check the
  constructor functions already assume) — if the header does not match, return `NA()`
  rather than a malformed string, consistent with the library's `NA()`-based
  error-signaling convention (`Dummy_Levels`/`Dummy_Code`).

**Argument type — RESOLVED: anchor cell, with the rationale recorded for the
next person to revisit it.** The considered alternatives were:

1. **Sheet name (text)** — `Model_Formula_String("Life Expectancy Model")`.
   Requires `INDIRECT` to reach an arbitrary sheet's cells, which is volatile and
   breaks on sheet rename — the same class of problem `Source_Data` was built to avoid
   for the table reference.
2. **Cell reference (anchor cell)** — `Model_Formula_String(Sheet2!$A$1)`,
   where the passed reference is a fixed anchor cell *inside* the target sheet's spec
   block. Every other cell the function needs is reached by `OFFSET`/`INDEX` relative
   to that one reference — no `INDIRECT`, not volatile, and it keeps the same "one
   retargeting point" pattern `Source_Data` already established.

The choice is option 2 (anchor cell). More consistent with the project's existing
aversion to `INDIRECT`, and it costs the user only one extra click (pointing at a cell
instead of typing a sheet name) when registering a model on the Comparison sheet.

### Sheet layout

| Zone | Contents |
|---|---|
| **Model registry** | One row per registered Regression sheet: Col A = hyperlink, display text = `Model_Formula_String(...)`, link target = a fixed anchor cell inside that sheet's status block (to the right of the spec block, per the "focus on the outputs, not the inputs" framing) |
| **GoF table** | R², Adjusted R², AIC, AICc, BIC, PRESS, LOOCV, F-statistic, F p-value, n, k — each a cross-sheet reference to a fixed status-block cell on the corresponding Regression sheet. References the **unit-space headline** statistic (shipped in v2.2 Transforms), so a logged model and a level model line up as comparable quantities by construction |
| **Shared prediction inputs** | Two columns: spec name, value. A single shared set of "what-if" inputs, entered once |
| **Prediction results table** | One row per model: hyperlink to that sheet's prediction output cell, predicted value, prediction interval bounds |

**Data flow for prediction inputs (settled direction):** the Model Comparison sheet is
the *source*; individual Regression sheets pull from it via `XLOOKUP` keyed on spec name,
rather than the reverse. This lets one shared "what-if" scenario (e.g. GDP = 5000,
Schooling = 12) drive every registered model's prediction row simultaneously — the actual
point of an apples-to-apples comparison sheet.

> **Open decision — mismatched predictor sets.** Shared prediction inputs assume a common
> variable universe across the models being compared. A registered model that does not
> include a given spec name (e.g. one model uses `Status`, another does not) needs a
> defined fallback — `XLOOKUP`'s `[if_not_found]` argument returning a visible `""`/`NA()`
> rather than a `#N/A` propagating into the prediction is the likely answer, but the exact
> behavior should be decided before building.

### Interface contract — RESOLVED: `Comparison_Anchor` (and its siblings) ship at v2.3 as a public-interface commitment

**RESOLVED.** v2.3 introduces three sheet-scoped named ranges per Regression sheet
that become the public interface for the Model Comparison sheet's formulas:

| Named range | Points at | Used by |
|---|---|---|
| `Comparison_Anchor` | A single anchor cell inside the status block (e.g. the Response-in-effect cell) | `Model_Formula_String`'s first argument; the model-registry hyperlink target on the Comparison sheet |
| `Comparison_Headline_GoF` | The v2.2 unit-space headline cells (R², Adjusted R², RMSE — all three) | The GoF table on the Comparison sheet; references unit-space-honest values by construction |
| `Comparison_Prediction_Output` | The center cell of the v2.1 prediction outputs (point · CI low/high · PI low/high) | The prediction results table on the Comparison sheet |

**Why named ranges, not raw coordinates.** A name gets re-pointed in one place if the
status block ever shifts rows in a future release; every downstream reference inherits
the new coordinate automatically. Raw coordinates would require finding and updating
every consumer at every layout change, with high risk of silent breakage.

**The public-interface commitment.** Per the Versioning definition above, these three
named ranges are part of the library's public interface the moment they ship at v2.3:
their existence, scope (sheet-scoped), and meaning (which status-block concept each
points at) become a versioning commitment. A future release may rename or repoint
them, but cannot silently remove or repurpose them without going MAJOR. The
changelog entry for v2.3.0 must name them explicitly so the commitment is discoverable.

**One anchor cell, not one per output.** `Comparison_Anchor` is a single cell inside
the status block; `Comparison_Headline_GoF` and `Comparison_Prediction_Output` are
separate named ranges because they live in different blocks (GoF block vs. prediction
block). The three together give the Comparison sheet everything it needs: anchor for
the spec string and the registry hyperlink, GoF cell for the comparison table,
prediction cell for the prediction-results table.

### New catalog surface

- New subcategory: **Model Comparison**, under either **Model Construction** (it is
  another way of reading a model) or as its own top-level category — open, since right
  now it is a single function, but the interface-contract helpers above could grow the
  list. Lean toward a new top-level category once there are 2+ functions in it, per the
  existing "categories describe what a function does" rule.
- `Model_Formula_String(anchor_cell)` — name RESOLVED. Reads a Regression sheet's spec
  block via a single anchor-cell reference (no `INDIRECT`) and returns the
  `y ~ x1, x2, x3` formula string. Header-signature validation; `NA()` on non-Regression
  targets.
- Possible follow-on: a small validation helper confirming an anchor cell actually points
  at a v2.0-shaped status block, reused by both the spec-string function and any future
  comparison-sheet helper.

---

## v2.4 — Resampling & Simulation

Bootstrap confidence intervals and Monte Carlo simulation. Validated as worthwhile
differentiators by their presence in Pyrcz's Excel demos and squarely in cost-estimation
territory (three-point estimates, MCS, risk analysis). These do not depend on the
two-sample or ANOVA work, so they come early. Bootstrap and Monte Carlo pair naturally and
may share a single sheet.

**No-volatile constraint, RESOLVED.** `RANDARRAY()`-based resampling is rejected.
Every recalc would re-draw a new bootstrap sample and silently re-compute every
resampled statistic; the result is the opposite of the library's
"live recalculation, formula transparency, auditability" philosophy. The cost
estimator who sees a 90% CI of (4.2, 5.7) one moment and (4.0, 5.9) the next, with
no record of which sample produced which, has not been given a tool — they have
been given a number that changes meaning under their feet.

**Resolved mechanism: pre-drawn random table, indexed.** A single sheet-scoped
named range `Bootstrap_Random_Draws` holds a fixed table of uniformly-distributed
random numbers, pre-drawn once at build time and visible in the workbook (auditable,
reproducible). The bootstrap loop indexes into this table via
`INDEX(Bootstrap_Random_Draws, MOD(SEQUENCE(n_resamples), ROWS(Bootstrap_Random_Draws))+1)`,
giving a resample-index sequence that wraps cleanly through the table. The
random number seed is the same SHA-derived seed the QC build already uses
(see `analysis_cache.py`), so the draw is deterministic across builds. The
table is sized once for the library's default n_resamples (e.g. 1,000 draws) and
is rebuilt (not re-randomised at use time) only when a new build is generated.

**Consequence for `Bootstrap_CI`:** the function is non-volatile — same inputs,
same output, every recalc. Reproducibility wins; "fresh randomness" loses, but
for a workbook-bound tool the reproducibility is the more important property.
Users who want a new draw open the workbook in Excel, hit `F9` to force a
recalculation, and *the draw does not change* (it was pre-drawn at build). To
get a new sample they regenerate the workbook via `build_production.py`. This
is deliberate, not a limitation — it is the same trade-off the spec block made
when it decided the source-data table is a stable reference, not a live query.

**Monte Carlo draws (`MC_Percentile`, `PERT_Sample`)** use the same pre-drawn
table mechanism, with the per-distribution transform applied at INDEX time.

---

## v2.5+ — Future (sequence TBD; first two claimed)

The v2.5+ bucket previously listed seven candidates with no order. Two are
now claimed as the immediate successors to v2.4; the rest are deliberately
unordered pending real demand from a user (a single maintainer should not
pre-order things nobody is asking for yet).

### v2.5 — Bivariate / Two-sample *(claimed, next MINOR after v2.4)*

- **TODO: `T_Test_OneSample(data, mu0, alpha, [include])`** — test statistic, p-value, CI.
- **TODO: `T_Test_TwoSample(data1, data2, alpha, equal_var, [include1], [include2])`** —
  equal-variance, Welch unequal-variance, and paired variants via the `equal_var` flag
  (paired is a separate code path the flag does not cover; design before implementation
  — a 3-way flag or a separate `paired` boolean is the open question).
- **TODO: `F_Test_Variance(data1, data2, alpha, [include1], [include2])`** — feeds a
  recommendation cell that selects the appropriate t-test variant (equal-variance vs.
  Welch). The recommendation is a visible single cell, not silent.
- **TODO: `Covariance_Matrix(data, [include])`** — complement to the existing
  `Correlation_Matrix`. Sample covariance, not population; consistent with the existing
  catalog's sample-statistic convention.
- **TODO: Sheet layout** — inputs, test selector, F-test assumption check, output
  (test statistic, df, p-value, CI, effect size). `write_sheet_two_sample.py`.

**Why this next.** Two-sample tests are a discrete "tests of means" add that
maps directly to the cost-estimator audience (a/b comparisons, baseline-vs-proposed
frequently surface in the same workbooks that already use the v1.0 regression).
Independent of the Role-axis work; does not block or get blocked by Weight/Cluster.

### v2.6 — `Weight` Role (WLS) *(claimed, after v2.5)*

- **TODO: Implement the `Weight` Role** (at most one, per the cardinality rule that
  Response, Time, and Weight all share). Status-block validation identical to
  exactly-one-Response.
- **TODO: Thread weights through the engine** per the Role-axis design: weighted
  least squares with the same model-matrix constructor the v2.0 spec block already
  uses, with a single optional `[Weights]` argument (default uniform, i.e. OLS) on
  the inferential chain. The `[DF_Absorbed]` precedent (default 0, no-FE models
  identical) is the exact pattern to follow — default-uniform weights means every
  existing OLS call computes identically.
- **TODO: Update the Diagnostic Guide** to describe which diagnostics change
  interpretation under WLS. (WLS closes the loop opened by v1's Scale-Location
  diagnostic, which is the classic visual prompt that the user *might* want
  WLS.)

**Three-stage scope carried forward** (per the *Future roles* section): at v2.6
the user-supplied-weight stage is what ships. Variance-driver-derived weights
(`Var(x)` form) and full FGLS are v2.6+ follow-ons, not part of the first MINOR.

**Why this second.** After v2.5 (tests of means) the natural next Role-axis
addition is `Weight`, because it pairs with the existing Scale-Location
diagnostic the user is already looking at. Two-way FE (v2.7 or later) and
`Cluster` (v2.7 or later) build on the v2.1/v2.2 framework; `Time` is partly
forward-wired via the Sequence axis but the rest of its semantics still need
design work.

### Unordered (v2.7+ candidates, no claim)

The following are real candidate work but deliberately unordered pending
demand. The two above (Two-sample, `Weight`) come first; these do not have a
user-pressing-for-them signal yet, and a single maintainer should not
pre-order work that may not be the next thing actually needed.

- **Two-way Fixed Effects** — the trio (`Absorb_Two_Way_Fixed_Effects`,
  `Demean_Two_Way_Balanced`, `Fixed_Effects_Convergence_Check`), the two-way
  `Is_Balanced_Panel` check, lifting the v2.1 one-FE-variable status-block
  error, and the two-way prediction question. The v2.1 FE engine is the
  prerequisite — does not start until v2.1.0 ships.
- **Multi-group means (ANOVA)** — one-way ANOVA as regression on group
  dummies (post-v2.0, a single Categorical predictor *is* an ANOVA — but the
  framing, the post-hoc comparisons (Tukey HSD or Bonferroni), and the
  dedicated sheet are new work).
- **`Cluster` Role (clustered SEs)** — the post-FE inference endgame. Has
  forward wiring in `Serial_Correlation_Group()`'s dormant branch (PR #106),
  so the resolver side is partial; the engine side (cluster-robust variance
  estimator) is not.
- **`Time` Role (time-index designation)** — partially forward-wired via the
  v2.1 Sequence axis. The full `Time` Role adds the time-index semantics
  (e.g. for the future time-series sheet, for the Lag_By / Difference_By
  cross-sheet calls) and depends on design work for how `Time` and
  `Sequence` interact (a column can be both? one or the other?).
- **Time series** — `Moving_Average`, `Exponential_Smoothing`; depends on
  the `Time` Role landing.
- **Fourier analysis** — long-tail. Out of scope; the *ToolPak Parity
  Reference* notes it is "intentionally skipped" and a later
  addition-by-demand decision, not a planned milestone.
- **Decision analysis** — long-tail (loss functions, cost/risk oriented).
  Not on the planning horizon.

---

## Analysis ToolPak Parity Reference

The ToolPak ships 19 tools. Tracking which are covered, planned, or intentionally skipped.

**Covered or exceeded (v1):** Regression (with diagnostics, influence measures,
cross-validation, information criteria, and prediction), Correlation, partial descriptive
stats; Descriptive Statistics + Histogram (v1.1). **v2.0 exceeded further:** categorical
predictors via the declarative spec, which the ToolPak has never offered; fixed-effects
panel regression follows at v2.1.

**Planned:** Rank/Percentile; t-tests, F-test,
Covariance (future two-sample); one-way ANOVA (future); Moving Average + Exponential
Smoothing (future time series).

**Intentionally skipped:**

- **z-Test (two-sample for means)** — assumes known population variance; rarely applicable.
- **Fourier Analysis** — engineering/signal domain; out of scope (may add later).
- **Two-factor ANOVA** — complexity vs. demand. *(Note: post-v2.0, two-way fixed effects
  via the spec covers adjacent territory; revisit whether this skip still holds.)*
- **Random Number Generation / Sampling** — largely redundant with native Excel functions.

**Why the library exists (ToolPak flaws it fixes):** ToolPak output is static (pasted
values that never update when inputs change), opaque (no formula trace), one-sheet-at-a-time
with manual reruns, locked behind a modal dialog, and diagnostically dated (no VIF, Cook's
distance, leverage, studentized residuals, PRESS, AIC/BIC, or cross-validation). The Lambda
Library is live, transparent, auditable, reusable, and diagnostically modern.

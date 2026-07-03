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

**Version ladder (current plan):**

| Version | Milestone | Status |
|---|---|---|
| v1.0 | Multivariate OLS / MLR | Shipped |
| v2.0 | Univariate (descriptives, histograms, distribution fitting) | Near complete — engine and sheet built; MoM-vs-MLE decision and QC outstanding |
| v3.0 | Specification-Driven Regression (roles: Continuous / Categorical / Fixed Effects) | Planned — MAJOR (breaks Regression sheet structure and `x_s()` semantics) |
| v4.0 | Resampling & Simulation (bootstrap, Monte Carlo) | Planned |
| v5.0+ | Two-sample, ANOVA, time series, additional roles (Weight, Cluster, Time) | Open |

Univariate keeps the v2.0 slot: its engine is already implemented (all 8 CDF/NLL
distribution functions, the grid-search optimizer, all three binning methods and their
`Bin_*` helpers, and all four goodness-of-fit measures) and its sheet writer is wired
into `build_production.py`. Specification-Driven Regression is greenfield by comparison,
so the near-finished milestone ships first.

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
**The v3.0 MAJOR bump is that opportunity** — decide during v3.0 planning whether to
fold the rename pass in or explicitly decline it.

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
| **Model Construction** | MLR Core · Coefficient Inference · Prediction · Specification & Design Matrix |
| **Diagnostics** | Residual · Influence & Leverage · Multicollinearity · Cross-Validation · Information Criteria |
| **Data Transformation** | Sample Construction & Diagnostics · Location & Scale · Group & Panel · Categorical & Model Construction · Longitudinal & Panel-Time |
| **Distribution Fitting** | Descriptive · Histogram Binning · Parameter Estimation · Goodness-of-Fit |
| **Resampling & Simulation** | Bootstrap · Monte Carlo |

*Change from prior revision:* **Specification & Design Matrix** added under Model
Construction to house the v3.0 constructor functions (see below). The Data
Transformation subcategories are unchanged; those functions now serve double duty as
constructor internals and standalone user-callable transforms.

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

(Full v1.0 gate history retained in git history of this file.)

## v2.0 — Univariate

*(Retains its original v2.0 slot — see version ladder. Engine and sheet already built.)*

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
- `Grid_Argmin(grid)` returns `minimum | 1-based row | 1-based column` horizontally and
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

## v3.0 — Specification-Driven Regression

**The central idea:** factor (categorical) and panel (fixed-effects) regression are not
new estimators — they are OLS on a transformed design matrix. Rather than telling the
user to manipulate their dataset into MLR form by hand, or building per-workflow staging
sheets, the Regression sheet's control block becomes a **declarative model
specification**, and `x_s()` is promoted from a column filter into a **model-matrix
constructor** that reads the spec and emits the numeric design matrix. Because every
engine function already consumes `x_s()`, the entire engine inherits factor and panel
capability without a single signature change.

> **Supersession note (recorded per the open-decisions convention):** this design
> replaces two earlier plans —
>
> 1. *Separate Factor Regression and Panel Regression sheets* with on-sheet staging
>    bands. The role column makes one spec-driven sheet handle all three cases; separate
>    sheets would be three copies of one mechanism. Factor and panel become documented
>    walkthroughs in the Regression Instructions sheet (see *Demonstration walkthroughs*
>    below), not sheets.
> 2. *The WLS-as-optional-`[Weights]`-argument vs. parallel-function-set decision.*
>    Re-litigated as a **`Weight` role value on the specification** — see *Future
>    roles*, below. The dedicated WLS Regression sheet plan is likewise superseded;
>    WLS-specific output relabeling happens on the one Regression sheet, keyed off the
>    active roles.

### The specification block (columns A–D)

| Col | Contents | UX |
|---|---|---|
| A | Predictor label | Unchanged from v1 |
| B | Include toggle (TRUE/FALSE) | Unchanged from v1 (orange input) |
| C | **Predictor Type** | Data-validation dropdown: `Continuous` · `Categorical` · `Fixed Effects`. **Pre-filled to `Continuous` by the build** — no cell is ever blank. A blank cell silently defaulting would be silent reinterpretation; a visible pre-filled default the user can change is not. |
| D | **Reference Level** | Orange input, meaningful only for Categorical rows. Blank = default reference, which is **the first level in sort order** (confirmed default — deterministic and explainable in one sentence, matching R's convention). Conditional formatting: grayed out when C is not `Categorical`; red when the entered level does not exist in the included sample. |

`All_Xs` keeps its current meaning — the raw, contiguous predictor range — but is now
allowed to contain text columns (e.g. Status, Country), since nothing numeric is
demanded of it until construction time.

### Role semantics

Named for **treatment, not measurement type** — the literature distinction that matters.
Status and Country are both categorical variables; what differs is their role in the
model. (The closest existing precedent is Stata's factor-variable notation: `c.gdp`,
`i.status`, `absorb(country)`. This is the spreadsheet-native version of that varlist.)

| Role | Constructor behavior | Coefficients |
|---|---|---|
| `Continuous` | Column passes through numerically, untouched | One, reported |
| `Categorical` | Dummy-coded against the reference level; one column per non-reference level | One per non-reference level, reported with level-qualified names ("Status: Developing") |
| `Fixed Effects` | Enters **no** column — instead the entire constructed block *and Y* are demeaned by its groups | None, deliberately |

**The asymmetry to keep in view:** Continuous and Categorical are column-local
operations; Fixed Effects is global. Marking GDP continuous affects one output column;
marking Country FE transforms every column of the design matrix and the response.

### Constructor pipeline (`x_s()` and `y_s()`)

1. Partition included spec rows by role.
2. Build the raw numeric block — Continuous columns as-is, Categorical columns
   dummy-coded via the Data Transformation primitives.
3. If any FE rows exist, demean the entire block **and Y** by those groups: one FE
   variable → `Demean_By`; two → `Absorb_Two_Way_Fixed_Effects`; three or more →
   visible error (no engine support).
4. Emit the design matrix as `x_s()` and the transformed response as **`y_s()`** — a
   new constructed named LAMBDA, parallel to `x_s()`. The moment FE is a role, the
   dependent variable is a constructed object too; raw Y is no longer what the engine
   should see.

One code path constructs both the training block and `X_new` for prediction — one
source of truth for encoding.

### Degrees of freedom — automatic, not manual

Absorbed df is computable from the spec: Σ over FE variables of (group count − 1). The
df correction that would have been a manual input cell under the staging-sheet design
becomes automatic and auditable. Without it, coefficients on demeaned data are correct
but **every SE, t, p, CI, MS_Residual, and information criterion is wrong** — the
absorbed effects consumed df the engine doesn't know about
(df_residual = n − k − Σ(Gᵢ − 1) for absorbed FE, alongside the intercept adjustment).

> **Open decision — df plumbing:** how absorbed df reaches the inference functions.
> Options: (a) thread an optional `[DF_Absorbed]` argument through the inference chain
> (`DF_Residual`, `MS_Residual`, `SE_Regression`, `SE_Coefficients`, `T_Stats`,
> `P_Values`, `CI_Lower`/`CI_Upper`, `F_Stat`, `P_Value_F`, `AIC`/`AICc`/`BIC`,
> `Prediction_Interval`, studentized/Cook's chain) — additive, so MINOR-compatible in
> principle, though it lands inside a MAJOR anyway; or (b) a small set of wrapper
> functions. Current lean: **(a)**, consistent with the one-source-of-truth principle
> and the rejection of parallel function sets.

### Model Spec status block — the transparency price, paid visibly

Construction inside a LAMBDA means the user can no longer *see* the design matrix by
scrolling. The replacement is a fixed-height block at the top of the outputs zone
answering "what model did I actually specify?":

- Constructed column count (k of the actual des
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
| v1.1 | Univariate (descriptives, histograms, distribution fitting) | No | **Shipped 2026-06-29** (workbook 1.1.0; renumbered from 2.0.0). MoM-vs-MLE resolved: MLE throughout. New sheet, no existing input changes meaning. Per-distribution Q-Q plots and PDF overlay curves outstanding (TODOs.md) |
| v1.2 | Workbook hardening & regression usability (Name Manager notes, identity-line data series, intercept-only and undersized-sample guards, LOOCV_Residual, build retry/RPC handling) | No | **Shipped 2026-07-03** (workbook 1.2.0; renumbered from 2.1.0) |
| v2.0 | Specification-Driven Regression (roles: Continuous / Categorical) | **Yes** | **Shipped 2026-07-05** (workbook 2.0.0; renumbered from 3.0.0) — MAJOR. Changed `x_s()` return semantics and restructured the Regression control block; includes the canonical rename pass. Shipped with `Transform` as a reserved placeholder column as planned; users transform their own variables via extra input-table columns in the interim |
| v2.1 | Fixed Effects (Role axis) — **one-way only** | No | Planned — panel regression, `y_s()`, absorbed df. One FE variable only; two-way absorption is its own post-v2.1 milestone (see v2.5+). Non-breaking: the absorbed-df correction is an optional `[DF_Absorbed]` argument defaulting to 0 (decision recorded under v2.1), so no-FE models are unchanged |
| v2.2 | Transforms (Response / Predictor Log, unit-space comparability) + the standalone Data Transformation function library | No | Planned — MINOR. Wires the reserved spec column G and ships the user-callable transform functions (Center, Zscore, Winsorize, Lag_By, …). Completes the Regression sheet as a fully functional deliverable |
| v2.3 | Model Comparison Sheet | No | Planned — MINOR, a *nice-to-have*. Read-only across finished Regression sheets; ships after Transforms so its comparisons are unit-space-honest from day one |
| v2.4 | Resampling & Simulation (bootstrap, Monte Carlo) | No | Planned |
| v2.5+ | Two-sample, ANOVA, time series, additional roles (Weight, Cluster, Time) | mixed | Open. Mostly additive minors; the **next MAJOR (v3.0)** is claimed by whichever of these next breaks the public interface |

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

(Full v1.0 gate history retained in git history of this file.)

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

Plus a per-distribution Q-Q plot for visual fit assessment. *(Not yet built — the
per-distribution Q-Q plots and the PDF histogram-overlay curves are the two v1.1
deliverables that did not ship; both are tracked in TODOs.md.)*

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

### The two-axis taxonomy

| Axis | v2.0 values | Future values | Meaning |
|---|---|---|---|
| **Variable Role** | Response · Predictor · Identifier · Filter · Omit | Fixed Effects, Weight, Time (v2.1+) | What the column *is* in the model |
| **Predictor Type** | Continuous · Categorical | *(closed — never grows)* | How a Predictor *enters the design matrix* — meaningful only when Role = Predictor |

Literature precedent: this is essentially the tidymodels `recipes` role system
(outcome / predictor / ID) plus Stata's factor-variable notation on the Type axis.

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

### The specification block (columns A–H)

The spec spans **every column of the source table**, one row per column:

| Col | Contents | UX |
|---|---|---|
| A | Variable name | Spills from the table's header row |
| B | **Role** | Dropdown: `Response` · `Predictor` · `Identifier` · `Filter` · `Omit`. Pre-filled by the build — no blank cells. Role precedes Include because it is the larger declaration. |
| C | Include toggle | Orange input; meaningful only when Role = Predictor |
| D | **Predictor Type** | Dropdown: `Continuous` · `Categorical`; meaningful only when Role = Predictor; pre-filled `Continuous` |
| E | **Reference Level** | Orange input, meaningful only for Categorical Predictors. Blank = **first level in sort order** (confirmed default, matching R). CF: red when the entered level does not exist in the analysis sample. |
| F | **Order** *(reserved, not implemented v2.0)* | Input, integer. Will control user-specified ordering of Identifier columns in the row-label text-join; v2.0 always joins in table order. Present now so the layout absorbs the feature without a future column insertion. Cell comment on row 3 marks it reserved; no validation yet (no fixed domain). |
| G | **Transform** *(reserved, not implemented v2.0)* | Input, dropdown. Will apply a transform (e.g. `Log`) to a Continuous Response or Predictor. v2.0 dropdown list is `None` only; build pre-fills `None`. Cell comment on row 3 marks it reserved. |
| H | **Levels** | **Computed display**: distinct level count over the mask-included rows, shown only for Categorical Predictors. Live against stratification. CF: **red when L ≤ 1 while included** (contributes L−1 = 0 columns). Large L needs no flag — the visible count is the warning. |
| I *(optional)* | **Design Columns** | Computed: Continuous Predictor → 1, Categorical Predictor → L−1, everything else → 0. Audit identity **ΣI = COLUMNS(x_s())**. Lean: include it. |

**Reserved-column policy (F, G).** Neither column is read by any v2.0 formula —
confirmed by construction, not by convention: `X_s()`, `Constructed_Column_Names()`,
`Row_Labels()`, and `Sample_Include()` must not reference `Spec_Order` or
`Spec_Transform`. The columns exist purely so the *sheet layout* absorbs the future
feature now; wiring them in a later release is additive (a formula change), not a
second column-insertion breaking the sheet a second time.

**Cascading relevance:** C–H gray out (conditional formatting) whenever Role ≠
Predictor — the same pattern as Reference-only-for-Categorical, applied one level up.

**Display derives, never feeds.** Columns H and I must not be inputs to the
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
- Constructed column count k — reconcilable against the ΣI audit identity
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

- **Spec block:** A–B → A–H (Role, Include, Type, Reference, Order, Transform,
  Levels — Order/Transform reserved, unread by any v2.0 formula), spanning all
  table columns, with cascading-relevance CF.
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

**Layout principle check (fixed-width left, fixed-height top):** the A–H spec block
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

1. **Durbin-Watson under FE** *(v2.1)* — relabel, caveat, or suppress. (Related to the
   iid caveat above.)
2. **Categorical × FE prediction encoding** *(v2.1)* — when non-FE categorical
   predictors coexist with fixed effects, x_new and x̄ᵢ must be formed in the
   *constructed* design-matrix space (dummies encoded through the same `Dummy_Code`
   path `x_s()` uses), not raw input space. The arithmetic above is unchanged (it is
   general in design-matrix coordinates); what remains is wiring the prediction-input
   UI to encode the chosen level through `Dummy_Code` before it reaches the formula.
   Shares machinery with v2.0 categorical prediction, so largely subsumed — recorded
   so the encoding step is not forgotten.
3. **Dataset bundling** (carried forward) — bundled WHO vs. data-agnostic. The
   `Source_Data` indirection reduces changeover to a one-name edit, which
   strengthens the data-agnostic option; the named-range rename it gated is
   partially subsumed by the v2.0 name changeover itself.
4. **Auto-completeness implementation — RESOLVED by construction, one caveat open.**
   `Sample_Include()` shipped with the role-aware completeness layer built in (every
   Filter column truthy AND the Response numeric AND every included Continuous
   Predictor numeric), so no separate `Role_Aware_Complete_Cases` function was needed.
   Remaining gap: **Categorical Predictors impose no non-blank condition yet** (known
   caveat, recorded in the human test plan and TODOs.md). Interim: a completeness
   column declared as a Filter (the `Full_Data` pattern) covers the gap manually.

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

> **Status note (post-ship correction):** this forethought item was **not** taken up —
> v2.0 shipped the prediction zone in the standard `Prediction_Interval` form with a
> single interval and Training-Mean input defaults (`write_sheet_regression.py`,
> Prediction Outputs zone), not the general group-mean form. The prediction-zone
> restructure (group-mean form, ȳᵢ/x̄ᵢ/Tᵢ keys, group dropdown) therefore moves onto
> the v2.1 work list, contrary to the "prevents a second redesign" intent above. The
> arithmetic consequence is unchanged; the cost is that v2.1 rebuilds the prediction
> zone instead of merely activating it.

**The one thing that could touch v1.0 (optional).** Surfacing both a mean-response CI
and a new-observation PI is not FE-specific — it is a standard OLS distinction. If v1.0
currently shows only one interval, adding the second is an additive change to the v1.0
prediction zone that could be backported now, or simply introduced at v2.0 when the
prediction zone is rebuilt in the general form anyway. *(Post-ship note: v2.0 kept the
single interval, so the CI + PI pair lands at v2.1 together with the group-mean
restructure above.)*

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

- `Lag_By(x, group, time, [periods], [include])` — prior-period value within the same
  group, keyed on `group`/`time`, not on physical row order.
- `Difference_By(x, group, time, [periods], [include])` — within-group time difference,
  \(\Delta_k x_{it} = x_{it} - x_{i,t-k}\).

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
  `Minmax_Scale`, `Winsorize`, `Ln_Positive`, `Zscore_By`, `Decompose_By`, `Lag_By`,
  `Difference_By`, `Numeric_Complete_Cases`, `Dummy_Column`, `Interact`,
  `Model_Matrix`. (`Demean_By`/`Group_Mean` arrive at v2.1 as FE internals; the
  two-way functions follow the two-way FE milestone.)
- **OLS MLR formula names stay the unit-space defaults** — `R_Squared`, `AICc`, etc. are
  not renamed or forked; they continue to describe whatever space the model was actually
  fit in. New functions are added *alongside* them for the unit-space (back-transformed)
  view.

### Unit-space fit statistics — naming

Proposed pattern: `{Response_Transform}_{Predictor_Transform}_Unit_Space_{Statistic}`,
covering the four combinations a single Log transform produces:

| Response | Predictor(s) | Function (example: R²) |
|---|---|---|
| Level | Level | *(no new function — existing `R_Squared` already is the unit-space value)* |
| Log | Level | `Log_Level_Unit_Space_R_Squared` |
| Level | Log | `Level_Log_Unit_Space_R_Squared` |
| Log | Log | `Log_Log_Unit_Space_R_Squared` |

> **Open decision — one LAMBDA per combination vs. a dispatch function.** Four
> combinations for one transform is manageable, but this is combinatorial: N transform
> types produce roughly N² named functions once a second transform (e.g. square-root,
> Box-Cox) is added. That tension sits directly against the "one canonical name, one
> LAMBDA" principle only if each combination stays a separate function — an alternative
> is a single dispatcher, e.g.
> `Unit_Space_R_Square(model, response_transform, predictor_transform)`, that internally
> `SWITCH`es on the transform arguments. The dispatcher is more scalable and still
> auditable (the SWITCH branches are visible in one place) but is a bigger departure
> from the current per-statistic-per-shape naming style used everywhere else in the
> catalog. Worth resolving before v2.2 implementation starts, since it sets the pattern
> for every future transform.

Per-statistic, this multiplies across whichever fit statistics get a unit-space
counterpart (R², Adjusted R², RMSE at minimum — AIC/BIC are likelihood-based and
comparing them across differently-transformed responses is a separate, harder question
worth flagging rather than solving here).

### Unit-space section on the Regression sheet

A new fixed-height block, alongside the existing GoF stats, containing SWITCH formulas
that read the Response/Predictor `Transform` values out of spec column G and select the
matching unit-space lambda — so the sheet always surfaces one "headline" comparable
statistic regardless of which transform is active, and that is the value the Model
Comparison sheet's GoF table references (not the raw-space statistic) once this ships.

### Prediction under transforms

- Predicted values and intervals need back-transformation (e.g. `EXP()` for a logged
  response) before they are comparable in the Model Comparison sheet's prediction
  results table.
- **Flagging, not resolving:** naive exponentiation of a log-linear prediction is a
  known biased estimator of the mean (the smearing-estimator problem). Whether to apply
  a correction (e.g. Duan's smearing estimator) or ship the naive back-transform with a
  documented caveat is a real statistical decision, not an implementation detail — record
  it as an open item rather than deciding it implicitly by whichever version of `EXP()`
  gets typed first.

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

### `Regression_Model_Spec_String` (name open — see below)

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

> **Open decision — argument type: sheet name vs. cell reference.** Two options:
>
> 1. **Sheet name (text)** — `Regression_Model_Spec_String("Life Expectancy Model")`.
>    Requires `INDIRECT` to reach an arbitrary sheet's cells, which is volatile and
>    breaks on sheet rename — the same class of problem `Source_Data` was built to avoid
>    for the table reference.
> 2. **Cell reference (anchor cell)** — `Regression_Model_Spec_String(Sheet2!$A$1)`,
>    where the passed reference is a fixed anchor cell *inside* the target sheet's spec
>    block. Every other cell the function needs is reached by `OFFSET`/`INDEX` relative
>    to that one reference — no `INDIRECT`, not volatile, and it keeps the same "one
>    retargeting point" pattern `Source_Data` already established.
>
> **Lean: option 2.** More consistent with the project's existing aversion to `INDIRECT`,
> and it costs the user only one extra click (pointing at a cell instead of typing a
> sheet name) when registering a model on the Comparison sheet.

**Function name — also open.** `Regression_Model_Spec_String` follows the naming
convention correctly (full words, Title_Case). Alternatives worth considering before
locking it in: `Regression_Spec_Label`, `Model_Formula_String`. Whichever is chosen, it
is the first function in a new subcategory (below), so get the name right once.

### Sheet layout

| Zone | Contents |
|---|---|
| **Model registry** | One row per registered Regression sheet: Col A = hyperlink, display text = `Regression_Model_Spec_String(...)`, link target = a fixed anchor cell inside that sheet's status block (to the right of the spec block, per the "focus on the outputs, not the inputs" framing) |
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

### Interface contract (formalizes what v2.0 built implicitly)

The v2.0 status block was designed to be fixed-height and fixed-position, but v2.0 never
had to promise that layout to *another sheet's formulas*. Model Comparison is the first
consumer that does. Recommend formalizing a small set of sheet-scoped named ranges per
Regression sheet (e.g. a `Comparison_Anchor` cell) purely so the Model Comparison sheet's
formulas reference a name, not a raw coordinate — if the status block ever shifts rows in
a future release, one name gets re-pointed instead of every downstream reference breaking
silently. (Note: promoting these anchors to named ranges makes them part of the public
interface per the definition above, so their stability becomes a versioning commitment —
worth doing deliberately.)

### New catalog surface

- New subcategory: **Model Comparison**, under either **Model Construction** (it is
  another way of reading a model) or as its own top-level category — open, since right
  now it is a single function, but the interface-contract helpers above could grow the
  list. Lean toward a new top-level category once there are 2+ functions in it, per the
  existing "categories describe what a function does" rule.
- `Regression_Model_Spec_String(anchor_cell)` (name pending)
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

---

## v2.5+ — Future (sequence TBD)

Deliberately left loose. Candidate milestones, roughly in conceptual order:

- **Two-way Fixed Effects** — the first candidate after v2.1: the `Absorb_Two_Way_Fixed_Effects`
  / `Demean_Two_Way_Balanced` / `Fixed_Effects_Convergence_Check` trio, the two-way
  `Is_Balanced_Panel` check, lifting the v2.1 one-FE-variable status-block error, and
  the (harder) two-way prediction question. Deliberately deferred until the one-way
  Fixed Effects framework is finished.
- **Bivariate / two-sample** — t-tests (one-sample, two-sample equal variance, Welch
  unequal variance, paired), F-test for variance equality feeding a recommendation on which
  t-test to use, and Covariance to complement the existing Correlation. (Pyrcz's
  "difference in means" / "difference in variances" demos map here.)
- **Multi-group means (ANOVA)** — one-way ANOVA, implemented as regression on group
  dummies, reusing the existing SS/MS/F machinery. A natural hinge showing ANOVA *is*
  regression — and, post-v2.0, expressible as a spec with one Categorical predictor.
- **Future specification roles** — `Weight` (WLS), `Cluster` (clustered SEs), `Time`
  (see *Future roles*, above).
- **Time series** — Moving Average, Exponential Smoothing.
- **Fourier analysis** — to be added later.
- **Decision analysis** — possible long-tail addition (loss functions), cost/risk oriented.

*(Removed from this list relative to the prior revision: the standalone WLS milestone
and dedicated WLS Regression sheet — superseded by the `Weight` role; see v2.0
supersession note.)*

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

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
**Decision made: the rename pass ships inside the v3.0 MAJOR bump** — the full
rename table lives in the v3.0 implementation TODO (AUTOCODER_TODO document).

**Sign-off record (v3.0 rename pass, 2026-07-03).** The rename table's DECIDE rows
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
user to manipulate their dataset into MLR form by hand, the sheet's control block
becomes a **declarative model specification spanning the entire source table**, and
`x_s()` is promoted from a column filter into a **model-matrix constructor** that reads
the spec and emits the numeric design matrix. Because every engine function already
consumes `x_s()`, the entire engine inherits the new capability without a signature
change.

The specification dissolves all three of the v1 sheet's hard-wired range names into
declarations: `y` (hard-wired) → the row whose Role is Response; `All_Xs` (hard-wired
span) → the table reference; `Regression_Sample_Include` (hard-wired to `[Full_Data]`)
→ the column(s) whose Role is Filter. After v3.0 there is nothing left to hard-wire —
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
>    coefficients. It moves to the Role axis (v3.1), the Type axis becomes permanently
>    Continuous/Categorical, and all future growth happens on Role.

### The two-axis taxonomy

| Axis | v3.0 values | Future values | Meaning |
|---|---|---|---|
| **Variable Role** | Response · Predictor · Identifier · Filter · Omit | Fixed Effects, Weight, Time (v3.1+) | What the column *is* in the model |
| **Predictor Type** | Continuous · Categorical | *(closed — never grows)* | How a Predictor *enters the design matrix* — meaningful only when Role = Predictor |

Literature precedent: this is essentially the tidymodels `recipes` role system
(outcome / predictor / ID) plus Stata's factor-variable notation on the Type axis.

**Role semantics (v3.0):**

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

### The specification block (columns A–F)

The spec spans **every column of the source table**, one row per column:

| Col | Contents | UX |
|---|---|---|
| A | Variable name | Spills from the table's header row |
| B | **Role** | Dropdown: `Response` · `Predictor` · `Identifier` · `Filter` · `Omit`. Pre-filled by the build — no blank cells. Role precedes Include because it is the larger declaration. |
| C | Include toggle | Orange input; meaningful only when Role = Predictor |
| D | **Predictor Type** | Dropdown: `Continuous` · `Categorical`; meaningful only when Role = Predictor; pre-filled `Continuous` |
| E | **Reference Level** | Orange input, meaningful only for Categorical Predictors. Blank = **first level in sort order** (confirmed default, matching R). CF: red when the entered level does not exist in the analysis sample. |
| F | **Levels** | **Computed display**: distinct level count over the mask-included rows, shown only for Categorical Predictors. Live against stratification. CF: **red when L ≤ 1 while included** (contributes L−1 = 0 columns). Large L needs no flag — the visible count is the warning. |
| G *(optional)* | **Design Columns** | Computed: Continuous Predictor → 1, Categorical Predictor → L−1, everything else → 0. Audit identity **ΣG = COLUMNS(x_s())**. Lean: include it. |

**Cascading relevance:** C–F gray out (conditional formatting) whenever Role ≠
Predictor — the same pattern as Reference-only-for-Categorical, applied one level up.

**Display derives, never feeds.** Columns F and G must not be inputs to the
constructor. Both the cells and `x_s()` call the same mask-aware primitive
(`Dummy_Levels`); the cells display what the constructor will do. Letting the engine
read a display column would make it load-bearing. One source of truth is the
*function*. (This is also what settled the `Dummy_Code` design: the level-vector
split is required — it makes display and constructor provably consistent, and gives
prediction the training level set.)

### Release scoping: v3.0 vs v3.1

**v3.0 ships the five Roles above with Continuous + Categorical Predictors — `x_s()`
alone, the Response consumed raw.** Anything that demeans or constructs a transformed
response is **v3.1**: the Fixed Effects Role, `y_s()`, absorbed df and its plumbing,
FE prediction. v3.0's constructor is purely column-local; the one global operation —
demeaning — gets its own release. Identifier's v3.1 dividend: Country and Year are
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
3. **Intercept coupling (settled — flag, don't switch).** v3.0 supports reference
   coding (L−1 columns), correct with the intercept on. No-intercept with a
   Categorical included is *flagged* — CF on `Allow_Intercept` plus a cell comment —
   never silently switched to one-hot. Full one-hot is deferred (v3.0.x candidate).

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

### Degrees of freedom — automatic, not manual *(v3.1)*

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
- Constructed column count k — reconcilable against the ΣG audit identity
- Level-qualified constructed column names
- **Included row count** after the effective mask (auto-completeness AND Filters),
  with the active Filter columns listed
- **Degenerate Categorical Predictors** (L ≤ 1 in the analysis sample): listed by
  name; the constructor contributes zero columns so the rest of the model still
  computes — visible degradation (red F cell + status line), not a hard error, not
  silent omission
- References in effect for each Categorical (surfaced even when defaulted)
- Intercept status; *(v3.1)* active FE variables, group counts, absorbed df
- **Error state** — a single visible cell flagging illegal specs: zero or multiple
  Response rows; invalid reference level; `Continuous` on a non-numeric column;
  no-intercept with a Categorical included (CF flag + cell comment, not a hard
  error); *(v3.1)* more than two FE variables

**Intercept × FE interaction (v3.1) — flag, don't force.** An intercept on demeaned
data estimates ≈ 0 and wastes a df — not catastrophic — so silently forcing
`Allow_Intercept` FALSE would be silent reinterpretation. Flag red and instruct.

### Zone changes on the Regression sheet (breaking — hence MAJOR)

- **Spec block:** A–B → A–F (Role, Include, Type, Reference, Levels), spanning all
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
  row labels. *(v3.1: residuals are within-model residuals under FE — Diagnostic
  Guide gains a paragraph.)*

**Layout principle check (fixed-width left, fixed-height top):** the A–F spec block
is fixed-width; the status block fixed-height; the constructed matrix never appears
on the sheet. Note the estimator choice is load-bearing: LSDV would spill a
data-dependent number of columns; the within transformation (v3.1) replaces them
with the same k columns transformed.

### New catalog functions (Specification & Design Matrix subcategory)

Working list; canonical names per the Naming Convention:

- Sheet-scoped constructors: `x_s()`, `Constructed_Column_Names()`, the derived
  Response and effective-mask LAMBDAs, and *(v3.1)* `y_s()`
- `Role_Aware_Complete_Cases(...)` — the automatic completeness layer of the
  effective mask (successor to the hard-coded `Full_Data` formula)
- `Absorbed_Degrees_Of_Freedom(...)` *(v3.1)* — Σ(Gᵢ − 1) from the spec
- `Dummy_Column(category, level, [include])` — one indicator column per explicit
  call; complements `Dummy_Code`
- Spec-validation helpers backing the status-block error states

### Demonstration walkthroughs (Instructions sheet, WHO data)

- **Factor walkthrough:** Status as a Categorical Predictor — the pooled vs.
  by-Status coefficient flip (the documented Simpson's Paradox risk), with the
  stratified comparison driven by a Filter column.
- **Panel walkthrough** *(v3.1)*: Country (and optionally Year) flipped from
  Identifier to Fixed Effects — directly addressing the panel structure already
  identified as violating OLS independence.

### Open items (recorded, not resolved)

1. **df plumbing** *(v3.1)* — optional argument vs. wrappers (lean: optional
   argument).
2. **FE prediction** *(v3.1)* — recovering group effects (α̂ᵢ = ȳᵢ − x̄ᵢ′β̂) implies a
   group-selection dropdown and recovery machinery. Fallback if deferred: display
   "within-model prediction (group effect not included)" with a visible caveat.
3. **Durbin-Watson under FE** *(v3.1)* — relabel, caveat, or suppress.
4. **Dataset bundling** (carried forward) — bundled WHO vs. data-agnostic. The
   `Source_Data` indirection reduces changeover to a one-name edit, which
   strengthens the data-agnostic option; the named-range rename it gated is
   partially subsumed by the v3.0 name changeover itself.
5. **Auto-completeness implementation** — the architecture is settled (effective
   mask = role-aware completeness AND Filters); `Role_Aware_Complete_Cases`
   itself remains to build. Interim: a completeness column declared as a Filter
   (the `Full_Data` pattern) is a valid manual substitute.

---

## Future roles (Role-axis values, v3.1+)

The Variable Role axis is extensible by design — these are **Role values, not new
mechanisms**, and the Predictor Type axis never grows:

- **`Fixed Effects`** *(v3.1)* — the panel role: enters no column; the entire design
  matrix and the Response are demeaned by its groups (one FE variable → `Demean_By`;
  two → `Absorb_Two_Way_Fixed_Effects`; three+ → visible error). Brings `y_s()`,
  absorbed df, and the within-model output relabeling.
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

Cross-cutting infrastructure, not tied to a single version. These functions serve
**double duty**: internals of the v3.0 constructor, and standalone user-callable
transforms for free-form work on the data sheet. Tracked as its own catalog, separate
from the version ladder (see *Function Categories*).

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
  whenever two Fixed Effects variables are active (v3.1).

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
  **v3.1 constructor internal** for a single Fixed Effects variable.
- `Zscore_By(x, group, [include])` — within-group standardization.
- `Decompose_By(x, group, [include])` — returns the between-group mean and within-group
  deviation as two columns, exposing \(x_{ig} = \bar{x}_g + (x_{ig} - \bar{x}_g)\).
- `Demean_Two_Way_Balanced(x, group1, group2, [include])` — direct two-way demeaning,
  exact only for a balanced panel; check with `Is_Balanced_Panel` first.
- `Absorb_Two_Way_Fixed_Effects(x, group1, group2, [include], [passes])` — iterative
  alternating-projection demeaning for unbalanced panels. Convergence is not verified
  internally; always pair with `Fixed_Effects_Convergence_Check`. **v3.1 constructor
  internal** for two Fixed Effects variables.

### Categorical & Model Construction — *subcategory*

- `Dummy_Levels(category, [reference], [include])` — retained categorical levels as a
  horizontal header row. Backs the v3.0 prediction-input validation lists.
- `Dummy_Code(category, [reference], [include])` — dummy-coded matrix. Use a reference
  level (treatment coding) when the design includes an intercept; full one-hot coding
  plus an intercept causes perfect multicollinearity. **Reference-level validation
  (confirming the requested reference actually exists in the included sample) is
  required at implementation, not deferred** — an invalid reference silently fails to
  drop a column and reintroduces the exact collinearity the function exists to prevent.
  **v3.0 constructor internal** for Categorical roles.
- `Dummy_Column(category, level, [include])` — single indicator column per explicit
  call (see v3.0 catalog additions).
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

## v4.0 — Resampling & Simulation



Bootstrap confidence intervals and Monte Carlo simulation. Validated as worthwhile
differentiators by their presence in Pyrcz's Excel demos and squarely in cost-estimation
territory (three-point estimates, MCS, risk analysis). These do not depend on the
two-sample or ANOVA work, so they come early. Bootstrap and Monte Carlo pair naturally and
may share a single sheet.

---

## v5.0+ — Future (sequence TBD)

Deliberately left loose. Candidate milestones, roughly in conceptual order:

- **Bivariate / two-sample** — t-tests (one-sample, two-sample equal variance, Welch
  unequal variance, paired), F-test for variance equality feeding a recommendation on which
  t-test to use, and Covariance to complement the existing Correlation. (Pyrcz's
  "difference in means" / "difference in variances" demos map here.)
- **Multi-group means (ANOVA)** — one-way ANOVA, implemented as regression on group
  dummies, reusing the existing SS/MS/F machinery. A natural hinge showing ANOVA *is*
  regression — and, post-v3.0, expressible as a spec with one Categorical predictor.
- **Future specification roles** — `Weight` (WLS), `Cluster` (clustered SEs), `Time`
  (see *Future roles*, above).
- **Time series** — Moving Average, Exponential Smoothing.
- **Fourier analysis** — to be added later.
- **Decision analysis** — possible long-tail addition (loss functions), cost/risk oriented.

*(Removed from this list relative to the prior revision: the standalone WLS milestone
and dedicated WLS Regression sheet — superseded by the `Weight` role; see v3.0
supersession note.)*

---

## Analysis ToolPak Parity Reference

The ToolPak ships 19 tools. Tracking which are covered, planned, or intentionally skipped.

**Covered or exceeded (v1):** Regression (with diagnostics, influence measures,
cross-validation, information criteria, and prediction), Correlation, partial descriptive
stats. **v3.0 will exceed further:** categorical predictors and fixed-effects panel
regression, which the ToolPak has never offered.

**Planned:** Descriptive Statistics + Histogram + Rank/Percentile (v2.0); t-tests, F-test,
Covariance (future two-sample); one-way ANOVA (future); Moving Average + Exponential
Smoothing (future time series).

**Intentionally skipped:**

- **z-Test (two-sample for means)** — assumes known population variance; rarely applicable.
- **Fourier Analysis** — engineering/signal domain; out of scope (may add later).
- **Two-factor ANOVA** — complexity vs. demand. *(Note: post-v3.0, two-way fixed effects
  via the spec covers adjacent territory; revisit whether this skip still holds.)*
- **Random Number Generation / Sampling** — largely redundant with native Excel functions.

**Why the library exists (ToolPak flaws it fixes):** ToolPak output is static (pasted
values that never update when inputs change), opaque (no formula trace), one-sheet-at-a-time
with manual reruns, locked behind a modal dialog, and diagnostically dated (no VIF, Cook's
distance, leverage, studentized residuals, PRESS, AIC/BIC, or cross-validation). The Lambda
Library is live, transparent, auditable, reusable, and diagnostically modern.

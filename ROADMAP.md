# Lambda Library — Roadmap

A native-Excel statistical and regression library (LAMBDA-based, no VBA) intended to
replace and expand upon the Excel Analysis ToolPak. Every release ships **functions plus
a pre-built worksheet** that drives them — the worksheet is a first-class deliverable, not
an afterthought.

Design philosophy: live recalculation, formula transparency, auditability, and no Solver /
no VBA wherever it can be avoided. The goal is that any result can be interrogated by
clicking the cell.

**Documentation map.** This file is the version plan only — the ladder, what's shipped,
and what's next. The "why" behind any design choice lives in
[DECISIONS.md](DECISIONS.md). The between-versions rules (naming convention, function
categories, the Role / Type / Sequence taxonomy, the spec block, the data-transformation
taxonomy) live in [ARCHITECTURE.md](ARCHITECTURE.md). Active work lives in
[TODOs.md](TODOs.md). A reader who needs the rationale behind a milestone's
design should follow the cross-link from the milestone's bullet to the corresponding
section in DECISIONS.md.

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
| v1.1 | Univariate (descriptives, histograms, distribution fitting) | No | **Shipped 2026-06-29** (workbook 1.1.0; renumbered from 2.0.0). MoM-vs-MLE resolved: MLE throughout. New sheet, no existing input changes meaning. PDF functions dropped as unnecessary — the histogram tables already compute per-bin probabilities as CDF deltas between bin boundaries. The two post-release leftovers (per-distribution Q-Q plots and combo-chart overlay lines built on those CDF-delta columns) shipped with the next workbook build |
| v1.2 | Workbook hardening & regression usability (Name Manager notes, identity-line data series, intercept-only and undersized-sample guards, LOOCV_Residual, build retry/RPC handling) | No | **Shipped 2026-07-03** (workbook 1.2.0; renumbered from 2.1.0) |
| v2.0 | Specification-Driven Regression (roles: Continuous / Categorical) | **Yes** | **Shipped 2026-07-05** (workbook 2.0.0; renumbered from 3.0.0) — MAJOR. Changed `x_s()` return semantics and restructured the Regression control block; includes the canonical rename pass. Shipped with `Transform` as a reserved placeholder column as planned; users transform their own variables via extra input-table columns in the interim |
| v2.1 | Sequence axis + gap-aware longitudinal + serial-correlation diagnostics + Fixed Effects (Role axis, one-way only) | No | In progress — the major Sequence / BFN / FE chain is substantially complete; pending the Sequence Period split (the v2.1 #1 prerequisite, the FE Role dropdown + status-block validation, the CI+PI prediction layout, and the FE engine work) |
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
gap, as planned: its engine was already implemented and its sheet writer was wired into
`build_production.py`. Specification-Driven Regression was greenfield by comparison,
so the near-finished milestone shipped first.

**Fixed Effects breakage flag (v2.1) — RESOLVED as non-breaking.** The absorbed-df
correction is threaded as an **optional `[DF_Absorbed]` argument defaulting to 0**,
leaving the no-FE df path untouched, so a model with no Fixed-Effects Role behaves
identically to v2.0. FE therefore stays a MINOR at v2.1. The full argument-threading
rationale is in [DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

---

## v1.0 — Multivariate (OLS / MLR) — SHIPPED

The complete OLS package and the first stable release: full model-fit and ANOVA
statistics, coefficient inference, multicollinearity, residual and influence
diagnostics, cross-validation, information criteria, distributional exploration, and
prediction. `Include` argument supports stratified OLS natively. Five-zone Regression
sheet (model inputs · predictor summary · regression outputs · prediction · residual
output), Diagnostic Guide sheet, Regression Instructions sheet, Version History sheet,
seven pre-built diagnostic charts, conditional formatting on out-of-bounds
diagnostics, QC via `analyze_regression_sheet.py` across six configurations
(sparse/medium/full predictor sets × intercept/no-intercept).

The full v1.0 gate history is in the git history of this file rather than restated
here. The release predates the open-decisions convention, so DECISIONS.md has no v1.0
section.

---

## v1.1 — Univariate — SHIPPED 2026-06-29

*(Renumbered from v2.0 to v1.1 — a new sheet that changes no existing input is a MINOR
under the interface definition in Versioning & Release Conventions. Released in the
workbook Version History as 1.1.0.)*

The foundation layer. A single input column drives three coordinated sections:
**describe** the data, **visualize** its shape, then **formalize** that shape with a
fitted distribution. The skewness value from the descriptive section motivates the
distribution choice in the fitting section.

- **Descriptive statistics** — mean, median, mode, sd, variance, min/max/range,
  skewness, kurtosis, count, and a prominent missing count.
- **Histogram binning** — three methods side-by-side: Sturges, Scott,
  Freedman-Diaconis. Each method gets its own column chart (gap width 0) built
  from the computed bin table.
- **Distribution fitting** — eight candidates (Normal, Lognormal, Exponential,
  Weibull, Gamma, Triangular, Beta, BetaPERT) ranked in a single comparison
  table. Closed-form MLE where possible; grid-search MLE over native two-input
  Data Tables for the two-parameter shape family. Per-distribution Q-Q plots
  and histogram distribution overlays (post-release v1.1 leftovers, shipped
  with the next workbook build).
- **The MLE-via-grid reframing** — the wall was never "MLE without Solver"; it
  was "MLE in closed form." Grid search clears the no-Solver bar for the
  two-parameter likelihood class.
- **PDF functions dropped as unnecessary** — the histogram tables already
  compute per-bin probabilities as `CDF(upper edge) − CDF(lower edge)`,
  the PDF integrated exactly over the bin. The histogram overlay is
  delivered via combo charts with the CDF-delta columns as line series,
  not via `PDF_*` LAMBDAs at bin midpoints.

Design rationale and resolved decisions: [DECISIONS.md § v1.1](DECISIONS.md#v11--univariate).

---

## v1.2 — Workbook hardening & regression usability — SHIPPED 2026-07-03

*(Released in the workbook Version History as 1.2.0; renumbered from 2.1.0. A
non-breaking MINOR consolidating the v1.0 / v1.1 surface into a workbook that holds up
under hand-use by cost estimators and other non-git users. The technical core is
unchanged; every existing v1.0 / v1.1 input computes identically.)*

- **Name Manager notes on every worksheet-scoped name** — every worksheet-scoped
  named range (constructor closures, `RegChart*` chart feeds, helper formulas)
  carries a `Comment` describing what the name is for and which sheet zone
  owns it. The Name Manager `Comments` column is the in-workbook index for
  non-git users.
- **Identity-line data series on the diagnostic charts** — the QQ,
  Actual-vs-Predicted, and Studentized-vs-Leverage charts need a `y = x`
  reference line. Real data series pointing both `XValues` and `Values` at
  the same named range, with `ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS`.
  The `Shapes.AddLine` approach was tried in v1.0 and abandoned — the
  pixel-coordinate positioning silently went wrong on resize.
- **Intercept-only and undersized-sample guards** — visible `n/a` tokens in
  the affected cells rather than silently-wrong numbers. Consistent with
  the library's "visible failure" philosophy.
- **`LOOCV_Residual` (new function, pair to `LOOCV_Prediction`)** — the
  leave-one-out residual vector `eᵢ / (1 − hᵢ)` in a single call. Makes
  `PRESS = SUMSQ(LOOCV_Residual(...))` an obvious idiom and unifies the
  LOOCV machinery under the Sherman-Morrison-Woodbury update.
- **Build-phase retry / RPC handling** — Excel's COM automation occasionally
  returns `RPC server unavailable` mid-`Calculate`. `_retry_on_open` in
  `build_production.py`; `main()` is split into two phases (sheet writes,
  no retry; recalculate + save, retry) so a transient failure in the
  short phase doesn't restart the multi-minute sheet-writing phase.

Design rationale and resolved decisions: [DECISIONS.md § v1.2](DECISIONS.md#v12--workbook-hardening).

---

## v2.0 — Specification-Driven Regression — SHIPPED 2026-07-05

*(Released in the workbook Version History as 2.0.0; renumbered from 3.0.0. The
canonical rename pass shipped inside this MAJOR as planned. The reserved
`Order`/`Transform` columns shipped unread by any formula, confirmed by construction.)*

The central idea: factor (categorical) and panel (fixed-effects) regression are not
new estimators — they are OLS on a transformed design matrix. Rather than telling
the user to manipulate their dataset into MLR form by hand, the sheet's control
block becomes a **declarative model specification spanning the entire source table**,
and `x_s()` is promoted from a column filter into a **model-matrix constructor**
that reads the spec and emits the numeric design matrix. Because every engine
function already consumes `x_s()`, the entire engine inherits the new capability
without a signature change.

The specification dissolves all three of the v1 sheet's hard-wired range names into
declarations: `y` (hard-wired) → the row whose Role is Response; `All_Xs`
(hard-wired span) → the table reference; `Regression_Sample_Include`
(hard-wired to `[Full_Data]`) → the column(s) whose Role is Filter. After v2.0
there is nothing left to hard-wire — the spec block plus one source-table
reference *is* the model.

- **Two-axis taxonomy (Role / Type) plus a structural Sequence axis** —
  Variable Role (Response / Predictor / Identifier / Filter / Omit),
  Predictor Type (Continuous / Categorical — closed, never grows), and
  Sequence (a structural flag, never grows). The full taxonomy and the
  cardinality rules are in
  [ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence).
- **Spec block A–L on the Regression sheet** — every column of the source
  table, one row per column. Cascading-relevance CF grays out cells
  irrelevant to the column's Role. The full A–L layout, the
  reserved-column policy, and the "Display derives, never feeds" rule
  are in [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-al).
- **Spec-order assembly for `x_s()`** with the **level-vector split** for
  Categorical Predictors — training and prediction both call the same
  encoder with the same training level vector. Reference-level validation
  via `NA()`, not silent fallback.
- **Canonical rename pass** — every LAMBDA function renamed to the
  Title_Case_With_Underscores convention. The retained initialisms
  (`AIC`, `AICc`, `BIC`, `VIF`, `PRESS`, `CDF`, `NLL`, `LOOCV`, `PERT`,
  `R`, `QQ`, `GoF`, `MS`, `SS`, `SE`) and the per-pass sign-offs are in
  [ARCHITECTURE.md § 1](ARCHITECTURE.md#1-naming-convention).
- **One breaking restructure, never a second** — the spec-block
  changeover at v2.0 is the only MAJOR change to the Regression sheet.
  All v2.1+ additions are either additive sections on the
  already-restructured sheet or backward-compatible engine signature
  additions.
- **Supersessions** — separate Factor / Panel Regression sheets →
  one spec-driven sheet; WLS parallel function set → `Weight` Role
  axis value; single-axis "Predictor Type" → two-axis taxonomy.
  Full supersession record in
  [DECISIONS.md § Supersession log](DECISIONS.md#supersession-log).

Design rationale and resolved decisions: [DECISIONS.md § v2.0](DECISIONS.md#v20--specification-driven-regression).

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
Two-way FE remains a post-2.1 milestone (see v2.7+).

The 2.1.0 release ships as **a single release**, not as a sequence of preview
builds — the FE Role dropdown, the CI+PI prediction layout, and the FE
activation on the sheet are all gated to ship with the engine, so users never
see a workbook that says "FE is in the dropdown but the engine is forthcoming."

**Shipped since v2.0.0:**

- Sequence structural axis (spec column H) and reserved Base Period Δ (column I). **(PR #101)**
- Gap-aware `Difference_By` / `Lag_By` and the Base Period Δ layer, with the Sequence Spacing block. **(PR #102)**
- Sequence-aware Durbin-Watson. **(PR #103)**
- BFN panel Durbin-Watson. **(PR #105)**
- Grouping-key resolver (`Serial_Correlation_Group()` with the dormant Cluster branch). **(PR #106)**
- v1.1 leftovers — histogram distribution overlays and per-distribution Q-Q plots. **(PRs #96, #97, #99, #100)**
- `--skip-univariate` CLI option. **(PR #98)**
- Spec-driven QC refactor (`analyze_regression_spec.py` and `test_regression_spec_qc.py`). **(PR #103)**
- Durbin-Watson under FE — second cell + mutual gating (BFN + resolver releases). **(PRs #105, #106)**

**Pending (in ship order; #1 prerequisite, #2/#3/#9 gated to #5):**

The full ship order is in [TODOs.md § v2.1](TODOs.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects-in-progress).
High-level shape: the Sequence fix lands first (the v2.1 #1 — column I
split into typed override + candidate-with-override display, with
spill-collision guard), then the FE Role dropdown + status-block
validation, then the CI+PI prediction layout, then the engine
(`Demean_By` / `Group_Mean` / `Is_Balanced_Panel` /
`Absorbed_Degrees_Of_Freedom` / `y_s` / `[DF_Absorbed]`), then the FE
group dropdown + ȳᵢ/x̄ᵢ/Tᵢ cells + BFN cell flips active. Follow-on
polish (BFN critical values, Categorical × FE prediction encoding,
residual relabel + Diagnostic Guide) ships with 2.1.0 if there's room.

Design rationale and resolved decisions: [DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

---

## v2.2 — Transforms & Unit-Space Comparability — PLANNED

Wires the reserved spec-block column G (`Transform`) and ships the standalone Data
Transformation function library. The `Transform` dropdown gains one real value:
`Log`, applicable to a Continuous Response and/or Continuous Predictors. With
`Transform = None` (the default) the model computes identically to v2.1 — this
is what keeps the release non-breaking.

**Why this is its own milestone, not a Model Comparison sub-feature:** the Model
Comparison sheet is only an honest comparison tool if the numbers it lines up mean the
same thing. An R² computed on `Ln(Life expectancy)` and an R² computed on raw
`Life expectancy` are not the same quantity, and putting them in adjacent cells of a
comparison table without correction is exactly the kind of silent misfiring the
library's design philosophy exists to prevent. Transforms is the release that makes
cross-model comparison trustworthy, not just possible — and it is what completes the
Regression sheet as a fully functional deliverable, which is why it precedes the
Model Comparison convenience layer.

- **Unit-space dispatcher, RESOLVED** — `Unit_Space_R_Squared(model,
  response_transform, predictor_transform)` with argument order
  model-then-response-then-predictor (matches the spec block's
  column-G reading order). One canonical name per statistic, internal
  `SWITCH` on the transform pair. The dispatcher is the first
  deliberate departure from "one canonical name, one LAMBDA" —
  justified by the combinatorial blow-up the exception avoids.
- **Prediction back-transformation, RESOLVED** — Duan's smearing
  estimator as the default, with a per-cell `Back_Transform_Method`
  toggle (`Duan` default | `Naive`). Naive is biased (Jensen's
  inequality); Duan is unbiased under iid residuals. Caveat row
  visible on the sheet.
- **Statistics with a unit-space counterpart:** R², Adjusted R², RMSE.
  AIC / AICc / BIC deferred (likelihood depends on the Jacobian of
  the transformation; the "right" comparison is on the original
  response's likelihood, not the transformed one's).
- **Standalone transform library ships in this release** —
  `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Ln_Positive`,
  `Zscore_By`, `Decompose_By`, `Numeric_Complete_Cases`,
  `Dummy_Column`, `Interact`, `Model_Matrix`. The full taxonomy
  and the `""`-vs-`NA()` row-alignment convention are in
  [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).

Design rationale and resolved decisions: [DECISIONS.md § v2.2](DECISIONS.md#v22--transforms--unit-space-comparability).

---

## v2.3 — Model Comparison Sheet — PLANNED

Every v2.0 Regression sheet already exposes a fixed-height, fixed-position **Model
Spec status block** (response in effect, constructed column count, level-qualified
names, included row count, error state). That block is an interface, not just a
display — the Model Comparison sheet is what happens when a second sheet is allowed
to *read* it. No new modeling capability is added; this is purely a cross-sheet
aggregation and navigation layer, which is why it is a MINOR.

- **Model registry** — one row per registered Regression sheet, hyperlink
  + display text from `Model_Formula_String(anchor_cell)`, link
  target a fixed anchor cell inside the spec block.
- **GoF table** — R², Adjusted R², AIC, AICc, BIC, PRESS, LOOCV,
  F-statistic, F p-value, n, k. References the **unit-space
  headline** statistic from v2.2 — a logged model and a level
  model line up as comparable quantities by construction.
- **Shared prediction inputs** — the Comparison sheet is the source;
  individual Regression sheets pull from it via `XLOOKUP` keyed on
  spec name, so one shared "what-if" scenario drives every
  registered model simultaneously.
- **Interface contract, RESOLVED** — three sheet-scoped named ranges
  per Regression sheet (`Comparison_Anchor`,
  `Comparison_Headline_GoF`, `Comparison_Prediction_Output`) become
  part of the library's public interface the moment they ship.
  The changelog entry for v2.3.0 must name them explicitly so
  the commitment is discoverable.

Design rationale and resolved decisions: [DECISIONS.md § v2.3](DECISIONS.md#v23--model-comparison-sheet).

---

## v2.4 — Resampling & Simulation — PLANNED

Bootstrap confidence intervals and Monte Carlo simulation. Validated as worthwhile
differentiators by their presence in Pyrcz's Excel demos and squarely in cost-estimation
territory (three-point estimates, MCS, risk analysis). These do not depend on the
two-sample or ANOVA work, so they come early. Bootstrap and Monte Carlo pair naturally
and may share a single sheet.

- **`Bootstrap_Random_Draws` table** — sheet-scoped named range
  holding a pre-drawn uniformly-distributed table, indexed at use
  time. Seeded from the same SHA-derived seed the QC build already
  uses (`analysis_cache.py`). Non-volatile by design.
- **`RANDARRAY()` rejected** — silently re-drawing per recalc is
  the opposite of the library's auditability philosophy. A cost
  estimator who sees a 90% CI of (4.2, 5.7) one moment and
  (4.0, 5.9) the next, with no record of which sample produced
  which, has not been given a tool.
- **Functions** — `Bootstrap_CI(data, stat_lambda, n_resamples,
  alpha, [include])`, `MC_Percentile(dist_params, n_samples,
  percentile)`, `PERT_Sample(min, mode, max, n_samples)`.

Design rationale and resolved decisions: [DECISIONS.md § v2.4](DECISIONS.md#v24--resampling--simulation).

---

## v2.5+ — Future (sequence TBD; first two claimed)

The v2.5+ bucket previously listed seven candidates with no order. Two are
now claimed as the immediate successors to v2.4; the rest are deliberately
unordered pending real demand from a user (a single maintainer should not
pre-order things nobody is asking for yet).

### v2.5 — Bivariate / Two-sample *(claimed, next MINOR after v2.4)*

`T_Test_OneSample`, `T_Test_TwoSample` (equal-variance / Welch / paired
variants — the 3-way flag or separate `paired` boolean is the open design
question, see [DECISIONS.md § v2.5](DECISIONS.md#v25--claimed)),
`F_Test_Variance` (feeds a recommendation cell that selects the t-test
variant), `Covariance_Matrix` (complement to the existing
`Correlation_Matrix`). Dedicated sheet layout with test selector and
F-test assumption check.

### v2.6 — `Weight` Role (WLS) *(claimed, after v2.5)*

A `Weight` value on the Role axis (see
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)
for the cardinality rule), with a single optional `[Weights]` argument
(default uniform, i.e. OLS) on the inferential chain. The `[DF_Absorbed]`
precedent (default 0, no-FE models identical) is the exact pattern to
follow. Three-stage scope: user-supplied weights → variance-driver-derived
weights → FGLS. v2.6 ships the first stage only.

### Unordered (v2.7+ candidates, no claim)

Two-way FE (the trio `Absorb_Two_Way_Fixed_Effects`,
`Demean_Two_Way_Balanced`, `Fixed_Effects_Convergence_Check`, the two-way
`Is_Balanced_Panel` check, lifting the v2.1 one-FE-variable status-block
error, and the two-way prediction question); multi-group means (ANOVA,
with Tukey HSD or Bonferroni post-hoc comparisons); `Cluster` Role
(clustered-robust V_β; partial forward wiring from
`Serial_Correlation_Group()`'s dormant branch); `Time` Role (partially
forward-wired via the Sequence axis; the full `Time` / `Sequence`
interaction is an open design question); Time series (`Moving_Average`,
`Exponential_Smoothing`); Fourier analysis and Decision analysis
(long-tail, out of planning horizon).

A user-pressing-for-them signal would reorder these; absent that, they
stay in this unordered bucket.

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

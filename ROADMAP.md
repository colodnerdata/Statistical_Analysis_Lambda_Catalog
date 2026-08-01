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

### Two numbers, once the build emits two workbooks

The definition above says "the user's inputs to **the** workbook" — singular. From
the Univariate split that is no longer true, and one number cannot answer the
question for two artifacts with entirely different input surfaces. The scheme:

| Number | Covers | Moves when |
|---|---|---|
| **Library version** | The shared function catalog — all 126 LAMBDA definitions, identical in both workbooks | A function is added, renamed, or changes what it returns |
| **Workbook version** *(one per artifact)* | That artifact's sheets, input cells, control blocks, and sheet-scoped names | That workbook's input surface changes |

**Why split this way.** Both workbooks carry the identical complete function
library, so a function change is genuinely a shared event and should move one
number. The input surfaces share nothing, so a Univariate layout change must not
move the number a Regression user reads as the answer to "do my existing inputs
still work?"

**The `Breaking?` flag attaches to the workbook version, not the library version.**
It answers a question about a user's saved inputs, and inputs are a property of a
workbook's sheets. A library-version bump that adds a function breaks nothing.

**How they display.** Each workbook's Version History sheet shows both, with its
own workbook version as the headline and the library version beside it:

```
Regression Workbook 3.0.0   ·   Function Library 3.0.0
Univariate Workbook 1.0.0   ·   Function Library 3.0.0
```

**One changelog serves both.** Entries stay in a single chronological list, each
tagged with which artifact's version it moves (or `Library` for a shared function
change). A reader filters by their artifact; a maintainer does not keep two files
in sync. The Version History sheet in each workbook renders the entries tagged
`Library` plus those tagged with that workbook.

**Worked example — the first two uses.** The Univariate split moves neither
workbook's major: it is packaging only, and every specification valid before it
produces the same result after. The grid shrink that follows is MAJOR for the
Univariate workbook version alone, because its Scale Min/Max/Step input cells
change meaning; the Regression workbook version does not move, and neither does
the library version unless a catalog function changes with it.

Rationale in
[DECISIONS.md § v3.0 versioning](DECISIONS.md#versioning-across-two-artifacts).

**Version ladder (current plan):**

| Version | Milestone | Breaking? | Status |
|---|---|---|---|
| v1.0 | Multivariate OLS / MLR | — (baseline) | Shipped |
| v1.1 | Univariate (descriptives, histograms, distribution fitting) | No | **Shipped 2026-06-29** (workbook 1.1.0; renumbered from 2.0.0). MoM-vs-MLE resolved: MLE throughout. New sheet, no existing input changes meaning. PDF functions dropped as unnecessary — the histogram tables already compute per-bin probabilities as CDF deltas between bin boundaries. The two post-release leftovers (per-distribution Q-Q plots and combo-chart overlay lines built on those CDF-delta columns) shipped with the next workbook build |
| v1.2 | Workbook hardening & regression usability (Name Manager notes, identity-line data series, intercept-only and undersized-sample guards, LOOCV_Residual, build retry/RPC handling) | No | **Shipped 2026-07-03** (workbook 1.2.0; renumbered from 2.1.0) |
| v2.0 | Specification-Driven Regression (roles: Continuous / Categorical) | **Yes** | **Shipped 2026-07-05** (workbook 2.0.0; renumbered from 3.0.0) — MAJOR. Changed `x_s()` return semantics and restructured the Regression control block; includes the canonical rename pass. Shipped with `Transform` as a reserved placeholder column as planned; users transform their own variables via extra input-table columns in the interim |
| v2.1 | Sequence axis + gap-aware longitudinal + serial-correlation diagnostics + Fixed Effects (Role axis, one-way only) + Generalized VIF | No | **Built and verified** — every TODOs #1–#10 item is DONE, verified against a live build (0 mismatches across all 12 spec-driven QC cases). `Design_Response` and `Design_Columns` (shipped at v2.1 as `y_s` / `X_s_Within`; renamed by the v3.0 constructor pipeline), `Absorbed_Degrees_Of_Freedom`, `Group_Prediction_Interval`, `GVIF`, and `Generalized_Tolerance` are all in `lambda_functions.json`. Awaiting only the human sign-off run of `HUMAN_TEST_PLAN_v21_regression_fixed_effects.md` and the 2.1.0 Version History entry, plus DEFERRED follow-on polish |
| v2.2 | Transforms (Response / Predictor Log, unit-space comparability) + the standalone Data Transformation function library | No | Partially delivered — MINOR. Column-G `Log` wiring shipped (`Response_Column()`/`X_s()`/`Constructed_Column_Names()`/`Constructed_Column_Transforms()`, the Prediction Inputs auto-log step, `Ln_Positive`); the unit-space dispatcher, Duan back-transformation, and the rest of the standalone transform library (Center, Zscore, Winsorize, …) remain open |
| v2.3 | Model Comparison Sheet | No | Planned — MINOR, a *nice-to-have*. Read-only across finished Regression sheets; ships after Transforms so its comparisons are unit-space-honest from day one |
| v2.4 | Resampling & Simulation (bootstrap, Monte Carlo) | No | Planned — MINOR. Pre-drawn random table (`Bootstrap_Random_Draws` named range) indexed at use time; non-volatile by design (every recalc reproduces the same draw). The QC build seeds the table from the same SHA-derived seed as `analysis_cache.py` |
| v2.5 | Bivariate / two-sample (one-sample t, two-sample t [equal-var / Welch / paired], F-test, Covariance) | No | Claimed — next MINOR after v2.4. F-test feeds a recommendation cell that selects the t-test variant; Covariance complements the existing `Correlation_Matrix` |
| v2.6 | `Weight` Role (WLS) | No | Claimed — after v2.5. User-supplied weights as the first stage; variance-driver-derived weights and FGLS as v2.6+ follow-ons. The `Weight` Role, its cardinality rule, and the three-stage scope stand; the **implementation mechanism changed at v3.0** — √w scaling in the constructor, not a threaded `[Weights]` argument |
| v2.7+ | Two-way FE, `Cluster` and `Time` Roles, Time series, ANOVA, Fourier, Decision | mixed | Unordered (deliberate — see Future section). Two-way FE has forward wiring from the v2.1 FE engine; `Cluster` has forward wiring from `Serial_Correlation_Group()`'s dormant branch; the rest are design-not-started |
| **v3.0** | **The engine-interface release** — bounded `Model_Context`, intercept relocation, the constructor pipeline, interaction spec columns, the materialization zone | **Yes** | In progress — MAJOR, and the second (and intended last) breaking restructure of the Regression sheet. Scope is **resolved**: one release delivered in three stages, of which stage 1 (constructor pipeline + intercept relocation) is code complete but has not yet cleared its spec-driven QC gate. See the milestone entry below |
| v3.x | Univariate as its own workbook; then the grid shrink | No / **Yes** (Univariate workbook only) | The split is packaging-only and non-breaking for both artifacts. The grid shrink that follows is MAJOR **for the Univariate workbook version only** and does not move the Regression workbook version |

**Ladder rationale.** Under the interface definition above, exactly two milestones
break user inputs. Specification-Driven Regression took v2.0; everything after it
was additive and opt-in, forming a v2.x train directly analogous to Python's 3.x
line — one breaking 3.0 followed by years of large but non-breaking minors
(async/await, pattern matching) that never forced a new major.

**v3.0 is the second break, and it is not a failure of that plan.** The v2.0 record
says "one breaking restructure, never a second," and the reasoning behind it still
holds: no single v2.x feature justified another. What accumulated instead was the
*sum* of correctly-classified additive changes — 24 functions carrying
`[DF_Absorbed]`, 48 carrying `[Allow_Intercept]`, two constructor names for one
pipeline, and no representation for interactions at all. "Additive" is the property
that makes a change a MINOR; it was never evidence that the interface could absorb
it indefinitely. v3.0 spends one break to unwind that, and the discipline it
replaces the old rule with is stated in the milestone entry below.

The next MAJOR after v3.0 is reserved for the next genuine interface break,
whenever that is.

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
- **Spec block A–L on the Regression sheet** *(A–N from v3.0, which appends the
  two interaction columns)* — every column of the source
  table, one row per column. Cascading-relevance CF grays out cells
  irrelevant to the column's Role. The full A–L layout, the
  reserved-column policy, and the "Display derives, never feeds" rule
  are in [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-an).
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

## v2.1 — Sequence, fixed effects, and the forward-wiring chain — BUILT, AWAITING SIGN-OFF

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
- Second sample dataset — Auto MPG as the **Mileage Data** sheet, and the default `Source_Table` retarget to it, demonstrating the one-name-edit dataset changeover. **(PRs #123, #125, #126, #127)**
- Generalized VIF for the multi-level-categorical case. **(PR #124)**
- Sequence Period (column I) / Period In Use (column J) split — TODOs #1. **(PRs #111, #112, #114, #129)**
- FE engine proper: `Demean_By`/`Group_Mean` primitives, the fit-time demeaned `y_s()`/`X_s_Within()` pair, `[DF_Absorbed]` threaded through 23 engine functions, the FE Role dropdown + status-block validation + intercept×FE red flag, and the group-mean-recovery Prediction Outputs rebuild (`Group_Prediction_Interval`) — TODOs #2–#9. **(PR #128, with the array-shaped `Group_Mean`/BYCOL fix in PR #130)**
- FE-aware residual-output headers (conditional on the FE Role being active, not just on a label). **(PRs #131, #132)**
- Third sample dataset — **Production Lots**, a small unbalanced learning-curve panel with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year); the `--regression-dataset {auto_mpg,life_expectancy,production_lots}` build flag; extended the spec-driven QC oracle to compute and verify Fixed Effects (group/within-transform support); a guard against degenerate `df_residual` when FE absorbs too many degrees of freedom. **(PR #133)**
- QC-oracle rebuild for the shipped 9-value CI+PI/group-mean-recovery Prediction Interval shape (`regression_shared.RegressionPredictionInterval`, `analyze_regression_sheet.py`, `analysis_cache.py` schema v16, `tools/inspect_regression_sheet.py`), closing the 58 automated mismatches the pre-v2.1 6-value oracle was reporting; a `selected_group` validation guard so an unrecognized group name errors loudly instead of dividing by zero — TODOs #10. **(PRs #134, #136)**
- Build-time sheet tab order and colors; Auto MPG `Origin` values decoded to region labels (US/Europe/Asia) in the source dataset. **(pre-#128 commits)**
- `safe_activate()` / `safe_freeze_top_row()` guards so a headless/no-focus Excel session (no interactive desktop, focus denied by the OS) cannot abort the build when a sheet writer activates its sheet or freezes its header row. **(PR #135)**

**Pending — follow-on polish, human sign-off, and the changelog entry:**

Every numbered TODOs #1–#10 item is DONE and verified against a live build
(0 mismatches across all 12 spec-driven QC cases). What remains before the
2.1.0 Version History entry:

- **Follow-on polish** (ships with 2.1.0 if there's room, otherwise slips to
  a 2.1.x patch): BFN critical values (**DEFERRED** — N,T-dependent bounds),
  Categorical × FE prediction encoding (**DEFERRED** — encode `x_new`/`x̄ᵢ`
  through `Dummy_Code` before the FE formula), and a residual-output
  relabel + Diagnostic Guide paragraph on residuals under FE
  (documentation-only). Full list in
  [TODOs.md § v2.1 follow-on polish](TODOs.md#follow-on-polish-ships-with-210-if-theres-room).
- **Human sign-off** — execute
  `HUMAN_TEST_PLAN_v21_regression_fixed_effects.md` (T0–T4) end-to-end in
  Excel and record a PASS, the same gate `HUMAN_TEST_PLAN_v20_model_construction.md`
  passed for the spec block at v2.0/v2.1 #1.
- **The Version History entry** — write the 2.1.0 row once the above lands.

Design rationale and resolved decisions: [DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

---

## v2.2 — Transforms & Unit-Space Comparability — PARTIALLY DELIVERED

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

**Shipped:**

- **Column-G Log wiring** — `Response_Column()` and `X_s()` apply
  `Ln_Positive` in place when a Response or Continuous Predictor row
  declares `Log`; `Constructed_Column_Names()` relabels the column
  `Ln(name)`; the new structural twin `Constructed_Column_Transforms()`
  carries the per-constructed-column Log/None flag the Prediction Inputs
  band needs (a Categorical Predictor's dummy columns always read
  `None` — Log is disallowed there and flagged red, never silently
  applied). The v2.1 FE-demeaning wrappers (`y_s()`/`X_s_Within()`) needed
  no changes — they compose transform-then-demean automatically. The
  Prediction Inputs band takes a raw value and auto-logs it internally;
  its Training Mean spill emits the geometric mean for a logged column to
  avoid double-logging the default prediction. New catalog function
  `Ln_Positive(x, [include])`. Verified against a real learning-curve
  model (Production Lots, raw columns with `transform="Log"`) matching
  the pre-existing precomputed-log-column case to floating-point
  precision. Full design rationale in
  [DECISIONS.md § v2.2 Transform column wiring](DECISIONS.md#v22--transforms--unit-space-comparability).

**Still open (this pass deliberately excluded them — no back-transformation or
cross-model comparability yet, only a correctly-fitted log-space model):**

- **Unit-space dispatcher, RESOLVED (design only, not implemented)** — `Unit_Space_R_Squared(model,
  response_transform, predictor_transform)` with argument order
  model-then-response-then-predictor (matches the spec block's
  column-G reading order). One canonical name per statistic, internal
  `SWITCH` on the transform pair. The dispatcher is the first
  deliberate departure from "one canonical name, one LAMBDA" —
  justified by the combinatorial blow-up the exception avoids.
- **Prediction back-transformation, RESOLVED (design only, not implemented)** — Duan's smearing
  estimator as the default, with a per-cell `Back_Transform_Method`
  toggle (`Duan` default | `Naive`). Naive is biased (Jensen's
  inequality); Duan is unbiased under iid residuals. Caveat row
  visible on the sheet. Until this ships, in-sample "Predicted Y" and
  the prediction outputs are labelled `(Log)` rather than back-transformed.
- **Statistics with a unit-space counterpart:** R², Adjusted R², RMSE.
  AIC / AICc / BIC deferred (likelihood depends on the Jacobian of
  the transformation; the "right" comparison is on the original
  response's likelihood, not the transformed one's).
- **Standalone transform library, remainder** —
  `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`,
  `Zscore_By`, `Decompose_By`, `Numeric_Complete_Cases`,
  `Dummy_Column`, `Interact`, `Model_Matrix` (`Ln_Positive` shipped
  early with the column-G wiring above). The full taxonomy
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
for the cardinality rule). Three-stage scope: user-supplied weights →
variance-driver-derived weights → FGLS. v2.6 ships the first stage only.

**The mechanism changed at v3.0; the feature did not.** This milestone was
planned as "a single optional `[Weights]` argument (default uniform), following
the `[DF_Absorbed]` precedent." That threading is superseded — with the intercept
owned by the constructor, √w scaling of the design matrix and response yields the
exact WLS estimator, standard errors, leverage, and Cook's distance, because the
intercept column correctly becomes √w rather than remaining ones. **WLS becomes a
constructor concern rather than an engine argument.** Everything else carries
forward unchanged: the `Weight` Role, its cardinality rule, the status-block
validation, and the three-stage scope. Weights are still declared in the spec
block.

One trap this avoids, recorded because it is the kind that ships silently:
`DEVSQ(√w ⊙ y)` is *not* the weighted total sum of squares — it centres on
mean(√w·y) rather than ȳ_w — so a naive "scale everything by √w" implementation
would leave SS_Total, and therefore R², wrong under WLS with no error anywhere.
The v3.0 projection form of `SS_Total` is correct by construction. See
[DECISIONS.md § v3.0 SS_Total](DECISIONS.md#ss_total-redefined-as-the-intercept-only-residual-sum-of-squares).

**Sequencing note.** If v2.6 ships before v3.0, it needs the `[Weights]` argument
after all, and v3.0 then unwinds it across the same ~24 functions. The review's
first sequencing implication applies directly: any decision to change the
mechanism is cheaper before v2.6 than after.

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

## v3.0 — The engine-interface release — PLANNED

The second and intended-last breaking restructure. It responds to
[REVIEW.md](REVIEW.md), whose findings share one shape: each decision was correct
in isolation and the cost is in the sum. Every design question below is
**resolved** in
[DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline);
only the *scope* is open.

- **Bounded `Model_Context`** — engine signatures collapse from
  `(X_s, Y, [Allow_Intercept], [Include], [DF_Absorbed])` to
  `(X, Y, [Include], [Context])`. Exactly four elements (`Has_Intercept`,
  `DF_Absorbed`, `Response_Transform`, `Predictor_Transform`), materialized once
  into a spill range. `[Include]` is a permanent floor, not a transitional
  state — an n×1 row mask would break boundedness. Resolves F1.
- **Intercept relocation** — the intercept column moves into the constructor and
  `Design_Matrix` stops synthesizing it. `SS_Total` is redefined as the
  intercept-only residual sum of squares, which collapses the ones / absent / √w
  cases into one projection formula.
- **The constructor pipeline** — `Design_Columns()` / `Design_Response()` /
  `Predictor_Columns()` replace the `X_s()` / `X_s_Within()` name fork, applying
  declared stages in the fixed order `encode → transform → demean → intercept →
  weight`. Resolves F2.
- **Interaction spec columns** — M (Interaction Term) and N (Interaction
  Operation), with a closed operation vocabulary carrying a symmetry attribute.
  Resolves F6.
- **The materialization zone** — `Model_Context`, `Sample_Include`, and the
  Constructed Design Matrix at the far right, in increasing width, terminating in
  the unbounded zone. Resolves F3.
- **Two artifacts and two version numbers** — see the Versioning section above.
  Resolves F4 and F8.

**What replaces "one breaking restructure, never a second."** That rule failed
because it constrained the *count* of breaks without constraining what could
accumulate between them. The v3.0 replacement constrains the accumulation
directly, in two rules that live in ARCHITECTURE because they bind every future
feature, not just this release:

1. **Properties of a fit travel in the context block, never as new optional
   arguments.** The reserved-slot pattern explicitly no longer applies to
   argument lists ([ARCHITECTURE.md § 7](ARCHITECTURE.md#7-reserved-spec-column-pattern-general)).
2. **Nothing may be placed right of the Constructed Design Matrix**, and
   materialized zones run in increasing width
   ([ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone)).

### Scope — **RESOLVED**

Which of the five pieces ship together was the one open question. They are
interdependent: the bounded context requires intercept relocation, which requires
the pipeline order; the materialization zone depends on the constructor pipeline
being settled first, or two variants of a soon-to-change architecture get
materialized; and the interaction columns and the audit column both touch the
spec-block layout. That argued for one release, against the general principle of
small increments.

**Resolved: one release, delivered in three pull requests.** The recommendation
below stands as the release contents; the counter-argument it records — that v3.0
becomes a large release hard to verify in one pass — is answered by splitting the
*delivery* rather than the release, so no second layout break is spent.

| Stage | Contents | Status |
|---|---|---|
| **1** | Constructor pipeline · intercept relocation | Code complete — **QC gate outstanding** |
| 2 | `Model_Context` · `Sample_Include` materialized · `[Has_Intercept]`/`[DF_Absorbed]` collapse into `[Context]` | Planned |
| 3 | Interaction spec columns M/N (reserved) · Design Columns audit column · Constructed Design Matrix zone + width guard · version bump | Planned |

The order is forced by the same dependencies listed above. Stage one carries a
verification property the others do not — **no number moves**, so the spec-driven
QC pass must report zero mismatches across all twelve cases — which is why it goes
first despite touching the most functions. The version number does not move until
stage three; v3.0 is not reached part-way through. See
[DECISIONS.md § v3.0 ships in three stages](DECISIONS.md#v30-ships-in-three-stages).

**Release contents — §3/§4/§5 fully, and §6/§7 as layout only:**

| Release | Contents | Break |
|---|---|---|
| **v3.0** | `Model_Context` · intercept relocation · constructor pipeline · **interaction columns reserved-and-unwired** · **materialization zone established + Design Columns audit column built** | MAJOR |
| v3.1 | Interaction wiring — the constructor actually builds the columns | MINOR |
| v3.2 | Full materialization of the design matrix | MINOR |

**Justification.** REVIEW.md's own sequencing note observes that F3 and F6 "all
want the same breaking change — resolving them separately spends three layout
breaks where one would do." The two interaction columns and the audit column are
*insertions* that shift every column to their right; that is the irreversible
part. The wiring of each is a formula change against a column that already
exists — precisely the reserved-column pattern, and exactly how column G went live
at v2.2. This spends one signature break and one layout break together, satisfies
the materialization zone's dependency on the pipeline, and leaves genuinely
additive work for the minors.

**What this scope does *not* cost.** F5 read as though the layout work would have
to be done twice — "the spec block is implemented twice; a layout change touches
both writers." It is not, and it does not. `write_sheet_regression.py` imports
the spec-block writers from `write_sheet_model_construction.py` and calls them,
so the two interaction columns, the audit column, and the materialization zone
each land in **one** writer. That single-implementation structure is part of what
makes this scope affordable, and it is why F5 does not appear in the release
above. See
[DECISIONS.md § v3.0 spec block](DECISIONS.md#the-spec-block-is-implemented-once-not-twice).

The counter-argument weighed and answered: v3.0 is a large release that is hard to
verify in one pass, and the human test plan for it is substantial. The earlier
draft of this entry offered splitting §6 and §7's layout work into a v3.0 and v3.1
pair, at the cost of a second layout break. The three-stage delivery above gets the
same reviewability without paying that cost — each stage is a reviewable diff with
its own verification gate, and the layout insertions still land exactly once.

**What stage one actually cost, against the estimates in this entry.** Two numbers
moved and both are recorded in DECISIONS: `Has_Intercept` survives in **thirteen**
functions rather than the estimated seven (the R²/sums-of-squares chain needs it as
an identifier, which the estimate did not anticipate), and `R_Squared` turned out
to be a **third** LINEST `const` site — the one that would have failed silently,
since LINEST reports the uncentered R² under `const = FALSE`. Neither changes the
design; both are why stage one is the one with the zero-mismatch gate.

---

## v3.x — The Univariate split, then the grid shrink — PLANNED

Two releases, deliberately not bundled.

**The split** moves Univariate Analysis into its own workbook. Both artifacts
carry the complete 126-function library — there is no bundling, no dependency
closure, and no per-artifact function subsetting; they differ only in which sheets
they contain. It is **non-breaking for both**.

The reason is a live correctness bug, not tidiness. `build_production.py` ships
the workbook in `XL_CALCULATION_SEMIAUTOMATIC` — Automatic except Data Tables —
forced by the Univariate sheet's six two-input Data Tables (2,400 NLL evaluations
per full recalculation). So **Univariate fit results are stale until the user
presses Ctrl+Alt+F9**: the flagship distribution-fitting sheet displays a previous
answer with no indication it has done so, which is the exact silent wrongness the
library's visible-failure philosophy exists to prevent. Splitting lets each
artifact set its own calculation mode, and the Regression workbook returns to full
Automatic. `--skip-univariate` and `--skip-data-table-calculations` already exist;
formalizing them into two build targets is most of the mechanism.

**The grid shrink** follows as a separate release of the Univariate artifact, and
is **MAJOR for that workbook's version only**. Weibull and Gamma collapse to
one-dimensional searches by profiling out the scale/rate parameter in closed form;
Beta stays two-dimensional but gets a method-of-moments start and a smaller grid.
Total evaluations fall from ~2,400 to ~370. Profiling is still genuine MLE — the
profile maximizer is the joint maximizer — so this extends the v1.1 MLE-via-grid
reframing rather than replacing it. The two-dimensional NLL heatmap becomes a
profile-NLL line chart for Weibull and Gamma, which is an upgrade in legibility:
the basin, the interior minimum, and any boundary hit are more visible in a line
chart than in a one-row colour strip.

It is MAJOR because the Scale Min/Max/Step input cells change or disappear, so a
user's saved bounds stop meaning anything. That is a workbook-interface break, and
it does not move the Regression workbook version.

Design rationale: [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).

---

## A note on the "v3.0" label in the codebase

The Specification-Driven Regression changeover was planned as v3.0 and renumbered
to **v2.0** before release, under the interface definition above. The old label
survives in comments and docstrings across `write_sheet_model_construction.py`,
`analyze_regression_spec_block.py`, `build_production.py`, and three test modules,
where "v3.0" means the spec-block changeover.

**v3.0 now means the engine-interface release.** The two are unrelated, and the
collision is live. `write_sheet_model_construction.py`'s docstring and the human
test plan filename (`HUMAN_TEST_PLAN_v20_model_construction.md`) are corrected;
the remaining comment references are tracked as a cleanup item in
[TODOs.md](TODOs.md). They are comments only — no executable logic reads the
label.

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

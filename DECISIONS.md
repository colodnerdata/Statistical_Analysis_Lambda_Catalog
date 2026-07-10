# Decisions

Resolved design decisions with their rationale, indexed by the version that
resolved them. Each entry is self-contained: the question, the resolution,
and the rationale in one paragraph. The version plan lives in
[ROADMAP.md](ROADMAP.md); foundational patterns live in
[ARCHITECTURE.md](ARCHITECTURE.md); active work lives in
[TODOs.md](TODOs.md). This file is the "why" record.

**Reading order for someone asking "why is the library shaped this way?"**
Open the version of interest, find the decision entry, follow the
cross-link to ARCHITECTURE for the foundational pattern it instantiates.

**Conventions used in this file**

- **RESOLVED** — a decision was made, the rationale is recorded here, and
  the resolution is in the codebase.
- **SUPERSEDED** — a decision was made and later replaced by a later
  decision. Cross-link to the superseding decision.
- **DEFERRED** — the question is open and intentionally held for a
  specific future version. The full deferral record (why, what changes
  when resolved) lives here; the action item (resolve it then) lives in
  [TODOs.md](TODOs.md).

---

## v1.0

The v1.0 release predates the open-decisions convention; the design record
lives in the git history of [ROADMAP.md](ROADMAP.md) rather than in this
log. No entries.

---

## v1.1 — Univariate

### MLE-via-grid reframing

**Question:** can the library claim MLE for the two-parameter shape
distributions (Weibull, Gamma, Beta) without Solver or VBA?

**Resolution:** grid search over the parameter space using native two-input
Data Tables, minimizing negative log-likelihood. This is genuine MLE — the
grid is a zeroth-order optimizer (no derivatives) standing in for
Newton-Raphson.

**Rationale:** the no-Solver, no-VBA resolution reframes the wall. The
wall was never "MLE without Solver"; it was "MLE in closed form." Grid
search clears the no-Solver bar for the entire two-parameter likelihood
class. The Triangular / BetaPERT cases are the separate exception:
Triangular's likelihood is non-differentiable at the mode, so it is fit
by direct min/mode/max estimation; BetaPERT is closed-form by
construction (PERT reparameterization of Beta).

### Closed-form MLE vs. grid-search MLE by distribution

**Question:** which fitting approach per distribution?

**Resolution:**

- **Closed-form MLE:** Normal (sample mean, variance), Lognormal (normal
  MLE on ln(x)), Exponential (1/mean), Triangular (min/mode/max — true
  MLE non-differentiable at the mode), BetaPERT (by construction).
- **Grid-search MLE:** Weibull (shape, scale), Gamma (shape, rate), Beta
  (α, β on [0,1]; requires rescaling).

### MLE throughout

**Question:** method of moments vs. MLE as the default fitting approach.

**Resolution:** MLE throughout.

**Rationale:** consistent with the library's transparency philosophy. The
two-stage grid-search refinement is the only practical complication;
closed-form where possible, grid-search where not, direct min/mode/max
where the likelihood is non-differentiable.

### Histogram overlays via CDF deltas (PDF functions dropped)

**Question:** how to draw fitted curves over the histograms.

**Resolution:** convert each histogram chart into a combo chart: the
existing column series (counts, gap width 0) plus one line series per
distribution sourced from the histogram table's CDF-delta columns. The
expected-counts question is settled as **expected counts on the shared
count axis**: each line series references a `UV_<method>_<Dist>_Expected`
named formula that multiplies the OFFSET-sized CDF-delta column by the
Count stat cell (probability × n), rather than introducing a secondary
axis.

**Rationale:** the histogram tables already compute, for all 8
distributions, the per-bin probability as the **CDF delta between the bin
boundaries** — `CDF(upper edge) − CDF(lower edge)` via the existing
`CDF_*` functions. That delta is the bin's probability mass (the PDF
integrated exactly over the bin, which is more faithful to the histogram
than a midpoint PDF evaluation), so no PDF functions are needed and none
will be implemented. The overlay is delivered by the combo chart instead.

### Edge behavior for bounded-support distributions

**Question:** how should Anderson-Darling and K-S handle bounded-support
distributions (Triangular, Beta, BetaPERT)?

**Resolution:** handle edge behavior deliberately. AIC/BIC compare cleanly
across all distributions (likelihood-based), but AD and K-S depend on the
fitted CDF and behave differently at support edges for the bounded
distributions.

**Rationale:** not enough to call out at the design table; the
implementation must guard against misleading GoF values at the support
edges.

### Two grid guards

**Question:** what happens when the located parameter sits on a grid
boundary, or when NLL is undefined / infinite for a parameter combination?

**Resolution:**

- *Minimum on a grid boundary* — conditionally format the located cell
  when it lands on a border to flag "widen stage-1 bounds." The next
  stage's range extends one current-grid step on either side of the
  located cell — zoom into the landed cell plus its neighbors, not a
  fixed shrink factor (a fixed shrink can drop a near-boundary true
  minimum outside the refined window and converge to the wrong point).
- *Undefined / infinite NLL* — wrap the NLL cell in `IFERROR` mapping
  bad combinations to a large finite sentinel, so `MIN` and the argmin
  lookup stay robust and the heatmap cleanly shows the overflow zones
  (Beta on [0,1] blows up if data sits exactly at 0 or 1; Gamma/Weibull
  overflow as shape approaches 0).

---

## v1.2 — Workbook hardening

### Name Manager note scope

**Question:** which worksheet-scoped named ranges get Name Manager
`Comment` notes?

**Resolution:** every worksheet-scoped name gets a note.

**Rationale:** the alternative (notes only on names a non-git user is
likely to encounter) was rejected because the line between "likely" and
"unlikely" is too subjective to keep stable across releases. The blanket
rule is the only one that scales. The motivation: a non-git user opening
`Lambda_Library.xlsx` in Excel has no way to read ROADMAP.md or
`write_sheet_regression.py` — the only documentation they can reach is
the workbook itself. The Name Manager `Comments` column is the
in-workbook index. The loop in `_setup_local_names` in
`write_sheet_regression.py` is the single source of the `Comment`
strings; the same pattern applies to the other sheet writers.

### Identity-line via real data series (not `Shapes.AddLine`)

**Question:** how to draw a `y = x` reference line on the QQ,
Actual-vs-Predicted, and Studentized-vs-Leverage charts.

**Resolution:** add a real data series pointing both `XValues` and
`Values` at the same named range (`RegChartQQX` for the QQ chart,
`RegChartFitY` for the Actual-vs-Predicted chart, the corresponding
`RegChart*` range for the leverage chart), with
`ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS` and `Name = "Identity"`.

**Rationale:** v1.0 attempted to draw the reference line with
`chart.Shapes.AddLine(...)` — a position-in-fixed-plot-area-pixels
approach that silently went wrong when the chart was resized, moved, or
its axis scaling changed. The data-series approach auto-rescales with the
chart, since its pixel coordinates are *not* fixed. The "Never draw
reference lines as shapes" guide in CLAUDE.md / AGENTS.md captures the
rationale; the helper `_add_identity_line` in
`write_sheet_regression.py` is the implementation.

### Undersized-sample failure mode

**Question:** what should the engine return for an intercept-only model
or a model with n ≤ k + 1 observations?

**Resolution:** surface a visible `n/a — undersized sample` or
`n/a — intercept-only` token in the affected output cells, with the same
token echoed in the status block. Visible failure, not silent numbers.

**Rationale:** consistent with the library's "visible failure"
philosophy for categorical degeneracy, invalid reference levels, and the
blank-categorical caveat. The undersized case produced a singular
`Gram_Inverse` and the engine fell through to a near-`#NUM!` result; the
intercept-only case computed reasonable-looking OLS output by accident.
Both were silently wrong. The fix lives in `analyze_regression_sheet.py`
with corresponding LAMBDA-side checks in `regression_shared.py`.

---

## v2.0 — Specification-Driven Regression

### Two-axis taxonomy (Role vs. Type)

**Question:** how should the model specification describe a column's
place in the model?

**Resolution:** two orthogonal axes plus a structural axis — **Variable
Role** (Response / Predictor / Identifier / Filter / Omit, with FE /
Weight / Time on deck for v2.1+), **Predictor Type** (Continuous /
Categorical — closed, never grows), and **Sequence** (a structural flag
post-v2.0, never grows).

**Rationale:** the single-axis "Predictor Type" design (one column
holding Continuous / Categorical / Fixed Effects) was superseded. Fixed
Effects was never a predictor type: it contributes no columns and no
coefficients. It moves to the Role axis (v2.1), the Type axis becomes
permanently Continuous/Categorical, and all future growth happens on
Role. The full taxonomy and cardinality rules live in
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence).

### Spec-order assembly for `x_s()`

**Question:** in what order should `x_s()` stack the included columns?

**Resolution:** top to bottom through the spec rows, via `REDUCE` over
spec-row indices with `HSTACK` from a full-height sentinel seed column
and `DROP` the sentinel at the end. Iteration predicate: Role = Predictor
AND Include = TRUE.

**Rationale:** the coefficient table reads in spec order, so the
spec's row order is the audit order. The sentinel + REDUCE pattern is
necessary because the included-column count is data-dependent (degenerate
categoricals contribute zero columns) — a fixed-arity constructor cannot
emerge until the iteration is done.

### Level-vector split (Dummy_Levels vs. Dummy_Code)

**Question:** how should categorical encoding handle the
training-vs-prediction asymmetry?

**Resolution:** separate *determine the levels* (mask-aware
`SORT(UNIQUE(...))` via `Dummy_Levels`) from *encode against a level
vector* (the broadcast `--(category = retained_levels)`). Training and
prediction call the same encoder with the same training level vector.

**Rationale:** an unseen level at prediction time encodes as all-zeros,
the correct behavior. The split also makes the spec block's Levels and
Reference In Use displays provably consistent with the constructor (they
all call the same `Dummy_Levels` primitive — the
"Display derives, never feeds" rule, see below).

### Intercept coupling (flag, don't switch to one-hot)

**Question:** how should the engine handle no-intercept models with a
Categorical Predictor included?

**Resolution:** flag red and instruct, never silently switch to one-hot
coding. Full one-hot is deferred (v2.0.x candidate).

**Rationale:** v2.0 supports reference coding (L−1 columns), correct
with the intercept on. No-intercept with a Categorical included is a
modeling error (the L−1 columns plus the omitted intercept fit L−1
parameters, but L intercepts are needed to fit the data) — silently
fixing it by switching to one-hot would silently change the meaning of
the model. The conditional format on `Allow_Intercept` plus a cell
comment make the error visible.

### `x_s()` row-mask contract (filter inside the engine, not the constructor)

**Question:** should `x_s()` apply the row mask internally?

**Resolution:** the constructor reads the effective mask *only* to fix
Categorical level sets; it always emits **full-height** columns and
leaves row filtering to the engine functions.

**Rationale:** filtering rows inside `x_s()` would double-filter against
the engine's own mask application. The contract keeps the constructor
column-local (a property of the spec, not the sample) and makes the
mask's two layers (auto-completeness AND declared Filter) live in one
place — the engine — where they belong.

### `Dummy_Levels` / `Dummy_Code` rebuild from scratch

**Question:** `Dummy_Levels` and `Dummy_Code` already exist as catalog
functions (added independently of this sheet's build) but are not yet
referenced by any sheet writer. Amend them or rebuild?

**Resolution:** drop both and rebuild from scratch.

**Rationale:** the v2.0 design's reference-level validation requirement
(confirming the requested reference actually exists in the included
sample) is required at implementation, not deferred — an invalid
reference must error (`NA()`), never silently fail to drop a column and
reintroduce the exact collinearity the function exists to prevent. The
pre-existing versions used string-based error returns, which break the
`IFERROR`/`ISNA` guard pattern the rest of the library relies on.
Rebuilding is cheaper than defensively wrapping.

### Auto-completeness — RESOLVED by construction, one caveat open

**Question:** how should the effective row mask handle role-aware
completeness (numeric required for Continuous, non-blank for
Categorical)?

**Resolution:** `Sample_Include()` shipped with the role-aware
completeness layer built in (every Filter column truthy AND the Response
numeric AND every included Continuous Predictor numeric), so no separate
`Role_Aware_Complete_Cases` function was needed. The hard-coded
`Data_Completeness(...[Life expectancy]:[Schooling])` span dies with it.

**Remaining gap (the blank-categorical caveat):** Categorical Predictors
impose no non-blank condition; a blank category value encodes as
all-zero dummies (indistinguishable from the reference level). Recorded
as `DEFERRED` in
[TODOs.md § v2.0](TODOs.md#v2.0--specification-driven-regression-shipped-leftovers);
interim workaround is a completeness column declared as a Filter.

### Spec-validation semantics (NA() everywhere)

**Question:** how should spec validation surface errors?

**Resolution:** `NA()` in the affected output cell. Never a descriptive
string; never a silent fallback. The status block aggregates signals
into a single visible Error State cell.

**Rationale:** the `NA()`-based error-signaling convention used by
`Dummy_Levels` / `Dummy_Code` lets every downstream `IFERROR`/`ISNA`
guard work without special-casing. Descriptive strings would require
every consumer to know the specific error vocabulary; `NA()` is the
universal "this didn't compute" signal that the rest of the Excel
ecosystem already understands.

### v1.0 → v2.0 spec changeover mechanic (one breaking change, never a second)

**Question:** how should the redesign to a spec-driven control block
affect the existing v1.0 Regression sheet?

**Resolution:** the v1.0 Regression sheet is redesigned exactly once, at
v2.0. That is the only breaking restructure — control block →
declarative spec block, and `x_s()` from column-filter to model-matrix
constructor. None of the v2.1/v2.2 decisions independently forces a
*second* redesign; each is either an additive section on the already-
restructured v2.0 sheet or a backward-compatible engine signature
addition.

**The one forethought item that prevents a second redesign.** The
degenerate G = 1 case of the FE group-mean prediction form
`ȳᵢ + (x_new − x̄ᵢ)′β̂` collapses exactly to the ordinary OLS prediction
(with 1/Tᵢ = 1/N recovering the textbook 1 + 1/n + leverage term), so
the **v2.0 prediction zone should be built in this general group-mean
form from the start**, with the whole sample as a single implicit group.
Then v2.1 Fixed Effects is literally just "let G > 1" — the group
dropdown and `AVERAGEIFS` keys activate, but the formula structure is
untouched. Building v2.0's prediction in the naive x_new′β̂-only form
would force tearing it up at v2.1; building it in the general form makes
FE prediction a genuine addition.

**Post-ship note:** v2.0 shipped the prediction zone in the standard
`Prediction_Interval` form rather than the general group-mean form, and
kept the single-interval output rather than a CI + PI pair. Both items
move to the v2.1 work list — v2.1 rebuilds the prediction zone instead
of merely activating it. The arithmetic consequence is unchanged; the
cost is one rebuild at v2.1.

### `Source_Data` indirection (volatile-INDIRECT avoidance)

**Question:** how should the spec reference the source data table when
structured references cannot be parameterized by another name without
volatile `INDIRECT`?

**Resolution:** a single sheet-scoped `Source_Data` name wrapping the
table reference is the retargeting point. Dataset changeover is a
one-name edit.

**Rationale:** the project has an explicit aversion to `INDIRECT` (it is
volatile, breaks on sheet rename, and silently re-evaluates). The
single-name indirection gives the same convenience (one edit to retarget
the model) without the volatility cost. The same pattern extends to
`Model_Formula_String(anchor_cell)` at v2.3 — the argument is an anchor
cell, not a sheet name, for the same reason.

### Supersession notes from the v1 planning

**Question:** what happens to the planned Separate Factor Regression and
Panel Regression sheets; the WLS-as-optional-`[Weights]`-argument vs.
parallel-function-set debate; and the single-axis "Predictor Type" design?

**Resolution:**

1. *Separate Factor Regression and Panel Regression sheets* with
   on-sheet staging bands — SUPERSEDED by the one spec-driven sheet;
   factor and panel become documented walkthroughs in the Regression
   Instructions sheet.
2. *The WLS-as-optional-`[Weights]`-argument vs. parallel-function-set
   decision* — SUPERSEDED as a **`Weight` value on the Role axis** (see
   [ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)
   and v2.6 below). The dedicated WLS Regression sheet plan is likewise
   superseded.
3. *The single-axis "Predictor Type" design* — SUPERSEDED by the
   two-axis taxonomy (this section's first entry).

---

## v2.1 — Sequence, gap-aware longitudinal, serial-correlation diagnostics, fixed effects

The full v2.1 ship list is in
[TODOs.md § v2.1](TODOs.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects-in-progress).
This file records the *decisions* — what was resolved and why.

### df plumbing — optional `[DF_Absorbed]` argument

**Question:** how should absorbed df be threaded through the inference
chain (SE, t, p, CI, MS-Residual, AIC/BIC)?

**Resolution:** optional `[DF_Absorbed]` argument defaulting to 0,
threaded through the df / MS-residual / t-critical chain.

**Rationale:** because the argument defaults to 0, every no-FE model
computes exactly as under v2.0 — this is what keeps Fixed Effects a
non-breaking MINOR. It is an *engine signature addition*, not a sheet
restructure: existing formulas that omit the argument are unaffected.
The alternative (a wrapper set) was rejected as parallel-function-set
maintenance, the same anti-pattern that collapsed the WLS-`[weights]`
debate into a Role-axis value.

### FE point prediction — group-mean recovery form

**Question:** how should prediction work under a one-way FE model when
the within estimator discards the G group intercepts the LSDV estimator
would have materialized?

**Resolution:** the algebraic identity
\(\hat{\alpha}_i = \bar{y}_i - \bar{x}_i' \hat{\beta}\) substituted
back in, giving:

> \(\hat{y} = \bar{y}_i + (x_{\text{new}} - \bar{x}_i)' \hat{\beta}\)

— the selected group's mean response, adjusted for how far the new
covariates sit from that group's mean covariates. Requires only three
new group-keyed summaries for the *one* selected group: \(\bar{y}_i\)
(`AVERAGEIFS` on the response), \(\bar{x}_i\) (`AVERAGEIFS` per
predictor), \(T_i\) (`COUNTIFS`) — all respecting the same
Include/Filter mask as the fit. Group selection is a data-validation
dropdown sourced from the observed level list (the spilled
`Dummy_Levels` output via a `#` spill reference / named range), which
also enforces the hard constraint that \(\hat{\alpha}_i\) exists only
for groups the model actually saw. A degenerate G = 1 (one "group" =
whole sample) collapses this exactly to the ordinary v1.0 prediction —
the key to building it once.

### Prediction interval — surface BOTH mean-CI and new-observation-PI

**Question:** what should the prediction zone show — a mean-CI, a PI,
or both?

**Resolution:** both, in adjacent cells. Same center, differing by one
variance term:

> Var(mean) = σ²/Tᵢ + (x_new − x̄ᵢ)' V_β (x_new − x̄ᵢ)
> Var(new)  = σ²·(1 + 1/Tᵢ) + (x_new − x̄ᵢ)' V_β (x_new − x̄ᵢ)

with σ² = MS-residual on absorbed df (N − G − K), V_β the coefficient
covariance already computed for inference, and the interval
\(\hat{y} \pm t(N-G-K) \cdot \sqrt{\text{Var}}\). The quadratic form
reuses the existing v1.0 prediction-leverage machinery, fed the
*deviation* (x_new − x̄ᵢ) instead of x_new. The interval is
**group-specific in both center and width**: σ² and V_β are pooled, but
Tᵢ (via 1/Tᵢ) and x̄ᵢ (via the deviation) change with the selected
group — so changing the group dropdown re-computes uncertainty, not
just the point estimate. Three lines: point · mean-response CI
(low/high) · new-observation PI (low/high).

**Sanity check for the test plan:** predicting at the group's own
centroid (x_new = x̄ᵢ) kills the quadratic term and the mean-CI
collapses to \(t \cdot \sqrt{\sigma^2 / T_i}\), the standard error of
\(\bar{y}_i\).

### Prediction input location — local on the regression sheet

**Question:** should prediction inputs live on the Regression sheet
itself, or only on the v2.3 Model Comparison sheet?

**Resolution:** on the regression sheet, in place. Inputs stay local to
each model sheet (making each sheet self-contained and able to predict
without the Model Comparison sheet), with validation-list dropdowns
for categorical predictors and the FE group. This does not foreclose
v2.3: whether those local cells later become `XLOOKUP` formulas
pointing at the shared Model-Comparison inputs is a v2.3-only decision
and is not forced by building the local version now.

### Scope — one-way FE only, iid errors, existing groups only

**Question:** what scope does v2.1's FE support cover?

**Resolution:** one-way FE only, iid errors, existing groups only.

**Rationale:**

- **One-way only.** A spec with two or more Fixed Effects variables is
  a visible status-block error; the two-way machinery
  (`Absorb_Two_Way_Fixed_Effects`, `Demean_Two_Way_Balanced`,
  `Fixed_Effects_Convergence_Check`) is deferred to its own post-v2.1
  milestone (see v2.7+).
- **Iid errors.** The interval assumes homoskedastic, non-serially-
  correlated errors (the classic FE assumption). Clustered/robust SEs
  are out of v2.1 — ship the iid interval with a documented caveat
  rather than implying robustness. Overlaps the Durbin-Watson-under-FE
  item below.
- **Existing groups only.** The ȳᵢ recovery is valid for a single
  grouping dimension. Two-way FE (Country × Year) does not recover
  intercepts as a simple group mean and is **explicitly out of v2.1
  scope** — flagged so the clean formula is not silently misapplied.

### Sequence structural axis (post-v2.0)

**Question:** where should the "ordering" semantic for
lag/difference/serial-correlation features live?

**Resolution:** a new structural axis on the spec block (column H),
distinct from the Variable Role axis.

**Rationale:** Sequence is structural, not a Role — a column can be
Role = Predictor, Type = Continuous AND Sequence = TRUE simultaneously.
Sequence never enters the design matrix. Cardinality is zero-or-one;
two-plus is a visible status-line error. Cardinality rules live in
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence).

### Sequence Period / Period In Use split (v2.1 #1)

**Question:** the v2.0.0 Base Period Δ cell (column I) was a
computed-with-override single cell. With source tables wider than the
shipped WHO sample, the spec block reads its own H/I cells and a
longer table could let the override spill overrun an input band. How to
fix the spill-collision risk without breaking the override pattern?

**Resolution:** rename column I to **`Sequence Period`** (the typed
override input) and add column J **`Period In Use`** following the
Reference Level / Reference In Use pattern (displays the typed
override if non-blank, otherwise the candidate). Bound every read of
the H/I/J band by `COLUMNS(Source_Data)` (the spill-placement principle
from CLAUDE.md). Override flagging lives on the Sequence Spacing
block's verdict lines only — the J spec-block cell stays plain, so the
spec reads top-to-bottom as a clean declaration. The named range moves
from `Spec_Base_Period_Delta` to `Spec_Sequence_Period`; every reader
(the `Base_Period_Delta()` accessor, the Sequence Spacing block's
Δ-in-use cell) is updated to read the new name. The Sequence Spacing
block (rows 28–34) and the QC analyzers are updated to match.

**Rationale:** the same reference-level pattern as the Categorical
Reference (E) / Reference In Use (L) pair. Override flagging on the
verdict lines (Regularity, Off-grid, no-natural-base-period prompt,
calendar-signature guidance) keeps the spec block clean while still
surfacing override prompts visibly to the user.

### Reserved-spec-column pattern (worked examples)

**Question:** how should features that span multiple versions be
absorbed without forcing a second column-insertion breaking the sheet?

**Resolution:** introduce the column in the sheet layout, ship it as a
placeholder that no formula reads, and absorb the future feature
additively (formula change) instead of with a second column-insertion.
Worked examples: v2.0 shipped `Order` (F) and `Transform` (G) as
reserved columns, unread by any formula; the v2.0 → v2.1 Sequence
release added column H live; the v2.1 #1 release split the override
mechanic into columns I and J. The function-side equivalent is a
dormant branch in a `SWITCH` returning a `RESERVED — vN+` token; the
v2.6+ `Cluster` branch in `Serial_Correlation_Group()` is the worked
example. The general pattern is in
[ARCHITECTURE.md § 7](ARCHITECTURE.md#7-reserved-spec-column-pattern-general).

### Durbin-Watson under FE — second cell + mutual gating

**Question:** how should the existing Durbin-Watson cell behave when
FE is active?

**Resolution:** ship a second cell, `BFN_Panel_Durbin_Watson`
(Bhargava–Franzini–Narendranathan 1982), with mutual gating against
the existing DW cell. Each self-guards: no Sequence → both
`n/a — requires Sequence`; Sequence + no FE → DW active, BFN
`n/a — no fixed effects`; Sequence + FE → BFN active, DW
`n/a — FE active`. FE detection counts Role="Fixed Effects" spec rows —
forward wiring that activates when the v2.1 FE role lands in the
dropdown.

**Rationale:** neither relabeling nor suppressing the existing DW cell
preserves its meaning. The BFN statistic differs from DW in two ways:
(1) the numerator's differencing is restricted to within-group
`(group, seq−Δ)` pairs via `Difference_By` (one source of truth — first
periods and panel gaps contribute no term, seams cannot manufacture
correlation, and the statistic is invariant to permuting group blocks);
(2) it is sign-aware for the within residuals, not the pooled OLS
residuals. Two cells, each true to its own statistic, is the
defensible answer.

### BFN critical values — DEFERRED

**Question:** what critical values should the BFN cell display?

**Resolution:** DEFERRED. The cell ships with an interpretation caveat
only (near 2 ⇒ no first-order autocorrelation in the within residuals).

**Rationale:** BFN significance bounds are N,T-dependent (Bhargava et
al. 1982 tables), and the standard DW bounds must not be presented
next to it. Surfacing proper BFN bounds is the recorded open item.
Tracked in [TODOs.md § v2.1](TODOs.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects-in-progress).

### Categorical × FE prediction encoding — DEFERRED

**Question:** how should prediction inputs be encoded when a
non-FE Categorical Predictor coexists with Fixed Effects?

**Resolution:** DEFERRED to v2.1 polish. When non-FE categorical
predictors coexist with fixed effects, x_new and x̄ᵢ must be formed in
the *constructed* design-matrix space (dummies encoded through the
same `Dummy_Code` path `x_s()` uses), not raw input space. Largely
subsumed by v2.0 categorical prediction; recorded so the encoding
step is not forgotten. Tracked in
[TODOs.md § v2.1](TODOs.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects-in-progress).

---

## v2.2 — Transforms & Unit-Space Comparability

### Unit-space dispatcher — `(model, response_transform, predictor_transform)`

**Question:** how should the unit-space (back-transformed) fit
statistics be named and structured for the four possible
(response_transform, predictor_transform) combinations (Log at v2.2.0)?

**Resolution:** a single dispatcher per statistic, internal `SWITCH` on
the `(response_transform, predictor_transform)` pair. `NA()` on
unrecognised values. Argument order: `model` first, then
`response_transform` then `predictor_transform` — matches the
spec-block's column-G reading order (`Spec_Transform` on the Response
row first, then on the Predictor rows), so the Regression sheet's
unit-space block can build the call with a 1:1 argument map. Functions:

```
Unit_Space_R_Squared(model, response_transform, predictor_transform)
Unit_Space_Adjusted_R_Squared(model, response_transform, predictor_transform)
Unit_Space_RMSE(model, response_transform, predictor_transform)
```

**Rationale — why dispatcher, not per-combination names.** A
per-combination scheme (`Log_Level_Unit_Space_R_Squared`,
`Level_Log_Unit_Space_R_Squared`, etc.) scales as O(N²) per statistic
and O(N² × M) across the M statistics with unit-space counterparts —
and the rename pain only gets worse with each new transform
(square-root, Box-Cox, …). The dispatcher scales as O(N²) in SWITCH
branches *regardless of how many statistics get a unit-space
counterpart*: adding a new unit-space statistic is one new dispatcher,
not N² new named functions. The SWITCH body is auditable in one place
per statistic, and the spec block's existing `SWITCH(Spec_Transform, …)`
pattern on column G is the precedent.

### Naming-style departure recorded

The dispatcher is a bigger departure from the per-statistic-per-shape
naming style used everywhere else in the catalog. The exception is
justified by the combinatorial blow-up the exception avoids. Pattern:
when a family is closed-form (one shape per name), use the per-shape
style; when a family is combinatorial in its inputs (N transforms × M
statistics), use a dispatcher. Future combinatorially-named families
(e.g. a future `Cross_Product_Of_Transforms_*` if one ever exists)
follow the same exception. See
[ARCHITECTURE.md § 1 "Naming-style departures"](ARCHITECTURE.md#1-naming-convention).

### Likelihood-based statistics (AIC / AICc / BIC) — DEFERRED

**Question:** do AIC, AICc, BIC get unit-space counterparts at v2.2.0?

**Resolution:** DEFERRED. They are flagged but not shipped at v2.2.0
even via a dispatcher, because the answer is not yet a single SWITCH
branch.

**Rationale:** comparing them across differently-transformed responses
is a separate, harder question (likelihood depends on the Jacobian of
the transformation, and the "right" comparison is on the original
response's likelihood evaluated at the back-transformed prediction,
not the transformed response's likelihood).

### Prediction back-transformation — Duan's smearing, per-cell naive toggle

**Question:** how should predicted values be back-transformed when the
Response is logged?

**Resolution:** Duan (1983) smearing estimator as the default
(\(\hat{y}_{\text{smeared}} = \hat{y}_{\text{log}} \cdot
\text{mean}(\exp(\text{residuals}))\)), with a per-cell
`Back_Transform_Method` toggle (`Duan` default | `Naive` = textbook
`EXP()`). A caveat row is visible on the sheet: *Duan = Duan (1983)
smearing; Naive = textbook EXP(ŷ), biased.*

**Rationale:** naive exponentiation of a log-linear prediction is a
biased estimator of the conditional mean response (Jensen's
inequality: \(E[\exp(X)] \ne \exp(E[X])\) whenever X has variance).
Duan smearing is unbiased under the iid-residual assumption and adds
one extra `AVERAGE(EXP(residuals))` cell. For a cost-estimator audience
the naive-vs-Duan difference is not a footnote; it is a
number-vs-number question they will see and ask about. Defaulting to
Duan ships the honest estimator; the toggle exists for users who want
the textbook form for a homework-style check. Reference: Duan, N.
(1983). "Smearing Estimate: A Nonparametric Retransformation
Method." *Journal of the American Statistical Association*, 78(383),
605–610.

---

## v2.3 — Model Comparison Sheet

### `Model_Formula_String(anchor_cell)` — name resolution

**Question:** what should the spec-string function be named?

**Resolution:** `Model_Formula_String(anchor_cell)`.

**Rationale:** considered alternatives
(`Regression_Model_Spec_String`, `Regression_Spec_Label`) and rejected
them. `Regression_*` prefixes couple the function name to a particular
sheet kind, which is wrong because the same string format is
meaningful for any model sheet (and the same anchor-validation pattern
would extend to a future GLM sheet, time-series sheet, etc.).
`Spec_Label` is vague — could be any text label. `Model_Formula_String`
is self-describing (a model, in formula-string form) and matches R's
"model formula" terminology.

### Argument type — anchor cell (not sheet-name text)

**Question:** should the first argument be a sheet name (text) or a
cell reference (anchor cell)?

**Resolution:** anchor cell. `Model_Formula_String(Sheet2!$A$1)`, where
the passed reference is a fixed anchor cell *inside* the target
sheet's spec block. Every other cell the function needs is reached by
`OFFSET`/`INDEX` relative to that one reference — no `INDIRECT`, not
volatile, and it keeps the same "one retargeting point" pattern
`Source_Data` already established.

**Rationale:** sheet name would require `INDIRECT` to reach an
arbitrary sheet's cells, which is volatile and breaks on sheet rename
— the same class of problem `Source_Data` was built to avoid. The
choice costs the user only one extra click (pointing at a cell
instead of typing a sheet name) when registering a model on the
Comparison sheet. Consistent with the project's `INDIRECT`-avoidance
stance.

### `Comparison_Anchor` / `Comparison_Headline_GoF` / `Comparison_Prediction_Output` — public-interface commitment

**Question:** how should the Comparison sheet reference cells on the
Regression sheet across version updates?

**Resolution:** three sheet-scoped named ranges per Regression sheet
that become the public interface for the Model Comparison sheet's
formulas:

| Named range | Points at | Used by |
|---|---|---|
| `Comparison_Anchor` | A single anchor cell inside the status block (e.g. the Response-in-effect cell) | `Model_Formula_String`'s first argument; the model-registry hyperlink target on the Comparison sheet |
| `Comparison_Headline_GoF` | The v2.2 unit-space headline cells (R², Adjusted R², RMSE — all three) | The GoF table on the Comparison sheet; references unit-space-honest values by construction |
| `Comparison_Prediction_Output` | The center cell of the v2.1 prediction outputs (point · CI low/high · PI low/high) | The prediction results table on the Comparison sheet |

**Rationale:** named ranges, not raw coordinates. A name gets
re-pointed in one place if the status block ever shifts rows in a
future release; every downstream reference inherits the new coordinate
automatically. Raw coordinates would require finding and updating
every consumer at every layout change, with high risk of silent
breakage. One anchor cell, not one per output — `Comparison_Anchor` is
a single cell inside the status block;
`Comparison_Headline_GoF` and `Comparison_Prediction_Output` are
separate named ranges because they live in different blocks.

**The public-interface commitment.** Per the Versioning definition in
[ROADMAP.md](ROADMAP.md), these three named ranges are part of the
library's public interface the moment they ship at v2.3: their
existence, scope (sheet-scoped), and meaning (which status-block
concept each points at) become a versioning commitment. A future
release may rename or repoint them, but cannot silently remove or
repurpose them without going MAJOR. The changelog entry for v2.3.0
must name them explicitly so the commitment is discoverable.

---

## v2.4 — Resampling & Simulation

### No-volatile constraint — pre-drawn `Bootstrap_Random_Draws` table

**Question:** how should the bootstrap and Monte Carlo functions
source their randomness?

**Resolution:** a single sheet-scoped named range
`Bootstrap_Random_Draws` holds a fixed table of uniformly-distributed
random numbers, pre-drawn once at build time and visible in the
workbook (auditable, reproducible). The bootstrap loop indexes into
this table via
`INDEX(Bootstrap_Random_Draws, MOD(SEQUENCE(n_resamples), ROWS(Bootstrap_Random_Draws))+1)`,
giving a resample-index sequence that wraps cleanly through the
table. The random number seed is the same SHA-derived seed the QC
build already uses (`analysis_cache.py`), so the draw is deterministic
across builds. The table is sized once for the library's default
n_resamples (e.g. 1,000 draws) and is rebuilt (not re-randomised at
use time) only when a new build is generated.

`RANDARRAY()` is rejected. Every recalc would re-draw a new bootstrap
sample and silently re-compute every resampled statistic; the result
is the opposite of the library's "live recalculation, formula
transparency, auditability" philosophy. The cost estimator who sees a
90% CI of (4.2, 5.7) one moment and (4.0, 5.9) the next, with no
record of which sample produced which, has not been given a tool —
they have been given a number that changes meaning under their feet.

**Consequence for `Bootstrap_CI`:** the function is non-volatile —
same inputs, same output, every recalc. Reproducibility wins; "fresh
randomness" loses, but for a workbook-bound tool the reproducibility
is the more important property. Users who want a new draw open the
workbook in Excel, hit `F9` to force a recalculation, and *the draw
does not change* (it was pre-drawn at build). To get a new sample
they regenerate the workbook via `build_production.py`. This is
deliberate, not a limitation — it is the same trade-off the spec
block made when it decided the source-data table is a stable
reference, not a live query. Monte Carlo draws (`MC_Percentile`,
`PERT_Sample`) use the same pre-drawn table mechanism, with the
per-distribution transform applied at INDEX time.

---

## v2.5+ — Claimed

### v2.5 — Two-sample test selector (OPEN)

**Question:** how should `T_Test_TwoSample` select among equal-var,
Welch, and paired variants?

**Resolution:** OPEN. The `equal_var` flag covers the equal-var and
Welch cases; the paired case is a separate code path the flag does
not cover. A 3-way flag or a separate `paired` boolean is the open
question, not yet resolved. Tracked in
[TODOs.md § v2.5](TODOs.md#v25--bivariate--two-sample-claimed-next-minor-after-v24).

### v2.6 — WLS: `Weight` Role, default-uniform `[Weights]` argument

**Question:** how should weighted least squares be added to the
engine?

**Resolution:** a `Weight` value on the Role axis (see
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)
for the cardinality rule), with a single optional `[Weights]` argument
(default uniform, i.e. OLS) on the inferential chain. The
`[DF_Absorbed]` precedent (default 0, no-FE models identical) is the
exact pattern to follow — default-uniform weights means every existing
OLS call computes identically.

**Three-stage scope carried forward.** v2.6 ships the
user-supplied-weight stage only. Variance-driver-derived weights
(`Var(x)` form) and full FGLS are v2.6+ follow-ons, not part of the
first MINOR.

**Rationale — why a Role value, not a parallel function set.** The
WLS-as-optional-`[Weights]`-argument-vs-parallel-function-set debate
was resolved as a `Weight` Role value rather than as a parallel
function set. The Role-axis design keeps the spec block as the single
place where every column's role in the model is declared; adding WLS
is a one-cell spec change (flip the Role to `Weight`), not a
parallel-call pattern. The same anti-pattern collapse that settled
the FE Role question.

---

## Aliases

A separate, optional layer of short, ALL-CAPS aliases may be added in
a later pass for power-user typing speed. Aliases are thin wrappers —
each alias LAMBDA's entire body is a call to the canonical function,
with no independent logic:

```excel
ABSORB2FE = LAMBDA(x, group1, group2, [include], [passes],
    Absorb_Two_Way_Fixed_Effects(x, group1, group2, include, passes)
)
```

This keeps a single source of truth: if the canonical implementation
changes, every alias inherits the fix automatically. Aliases are never
the documented or taught form — they exist purely as optional
shortcuts and should be introduced only after the canonical library
is stable, to avoid maintaining two names for a function that's still
under active revision.

The full alias table is a future-implementation record (not active
work), held here until the canonical library stabilizes enough to make
alias maintenance cheap.

### Regression — scalar outputs

| Alias | Canonical |
|---|---|
| `R2` | `R_Squared` |
| `ADJ_R2` | `Adjusted_R_Squared` |
| `MULT_R` | `Multiple_R` |
| `SE_REG` | `SE_Regression` |
| `DW` | `Durbin_Watson` |
| `QQ_CORR` | `QQ_Correlation` |
| `DFR` | `Regression_Degrees_Of_Freedom` |
| `DFE` | `Residual_Degrees_Of_Freedom` |
| `DFT` | `Total_Degrees_Of_Freedom` |
| `SSR` | `SS_Regression` |
| `SSE` | `SS_Residual` |
| `SST` | `SS_Total` |
| `MSR` | `MS_Regression` |
| `MSE` | `MS_Residual` |
| `P_VAL_F` | `F_Statistic_P_Value` |

### Regression — coefficient vectors

| Alias | Canonical |
|---|---|
| `COEF` | `Coefficients` |
| `SE_COEF` | `SE_Coefficients` |
| `TSTAT` | `T_Statistics` |
| `PVALS` | `P_Values` |
| `CI_LOW` | `Confidence_Interval_Lower` |
| `CI_UP` | `Confidence_Interval_Upper` |
| `P_R2` | `Partial_R_Squared` |
| `P_COR` | `Partial_Correlation` |
| `BETA_W` | `Beta_Weights` |

### Regression — observation vectors

| Alias | Canonical |
|---|---|
| `PRED` | `Predictions` |
| `RESID` | `Residuals` |
| `STDR` | `Studentized_Residuals` |
| `LEV` | `Hat_Diagonal` |
| `COOK_D` | `Cooks_Distance` |
| `LOOCV` | `LOOCV_Prediction` |
| `PI` | `Prediction_Interval` |

### Regression — utilities

| Alias | Canonical |
|---|---|
| `COMPLETE` | `Complete_Cases_Filter` |
| `CORMAT` | `Correlation_Matrix` |
| `DESIGN` | `Design_Matrix` |

### Univariate — descriptive

| Alias | Canonical |
|---|---|
| `DSTAT` | `Descriptive_Statistics` |
| `NMISS` | `Missing_Count` |

### Univariate — histogram binning

| Alias | Canonical |
|---|---|
| `NBINS` | `Number_Of_Histogram_Bins` |
| `EDGES` | `Bin_Edges` |
| `UEDGES` | `Upper_Bin_Edges` |
| `LEDGES` | `Bin_Lower_Edges` |
| `BIN_MIDS` | `Bin_Midpoints` |
| `BIN_FREQS` | `Bin_Counts` |

### Univariate — goodness-of-fit

| Alias | Canonical |
|---|---|
| `GOF_AD` | `GoF_Anderson_Darling` |
| `GOF_KS` | `GoF_Kolmogorov_Smirnov` |

### Grid-search helpers

| Alias | Canonical |
|---|---|
| `GS_MIN` | `Grid_Argument_Minimum` |
| `GS_OPT` | `Grid_Search_Optimum` |

---

## Supersession log

Decisions that were made and later replaced by a later decision. The
superseding decision lives at its version's section above; this log
just records what was replaced, when, and by what.

- **Separate Factor / Panel Regression sheets** (v1 planning) →
  SUPERSEDED at v2.0 by the one spec-driven sheet. Factor and panel
  become documented walkthroughs in the Regression Instructions
  sheet.
- **WLS as a parallel function set** (v1 planning) → SUPERSEDED at
  v2.0 by a `Weight` value on the Role axis (decided at v2.0, shipped
  at v2.6).
- **Single-axis "Predictor Type"** (v1 planning) → SUPERSEDED at v2.0
  by the two-axis taxonomy (Role vs. Type); Fixed Effects moves to
  the Role axis (v2.1).
- **PDF LAMBDAs for histogram overlays** (v1.1 planning) → SUPERSEDED
  at v1.1 by the combo-chart-from-CDF-deltas approach. PDF functions
  dropped as unnecessary; the histogram tables already compute per-bin
  probabilities as `CDF(upper edge) − CDF(lower edge)`.
- **`Shapes.AddLine` for reference lines** (v1.0) → SUPERSEDED at
  v1.2 by a real data series. The pixel-coordinate approach silently
  went wrong when the chart was resized or axis scaling changed.
- **String-based error returns in `Dummy_Levels` / `Dummy_Code`** (v1)
  → SUPERSEDED at v2.0 by `NA()`-based error returns. The string
  approach broke the `IFERROR`/`ISNA` guard pattern.
- **Computed-with-override single cell for Base Period Δ** (v2.0.0
  spec block, column I) → SUPERSEDED at v2.1 #1 by the
  Sequence Period (I) / Period In Use (J) split. Same reference-level
  pattern as the Categorical Reference (E) / Reference In Use (L)
  pair; resolves the spill-collision risk for source tables wider
  than the shipped WHO sample.
- **v1.0 standard prediction zone in x_new′β̂-only form** → recorded
  at v2.0 as needing the general group-mean form to absorb v2.1 FE
  additively. v2.0 shipped the standard form anyway; v2.1 rebuilds
  the prediction zone in the general form (the rebuild is the
  cost paid for the v2.0 short-cut).

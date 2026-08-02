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
- **REJECTED** — an alternative was considered at a decision point and
  not taken. Recorded rather than deleted, so the same option is not
  re-proposed without the reason it lost.
- **WITHDRAWN** — a proposal made during planning for a version and
  retracted before that version resolved, usually because a later
  decision in the same pass removed its basis.
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
as `DEFERRED` in
[TODOs.md § v2.0](TODOs.md#v20--specification-driven-regression-shipped-leftovers);
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

**Demonstrated by a shipped second table:** the workbook now ships a
second sample dataset — the Auto MPG data as the `MileageData` table on
the **Mileage Data** sheet — alongside Life Expectancy Data. It exists to
make the one-name retarget concrete: repoint `Source_Table` at
`MileageData[#All]` in Name Manager and the whole spec block re-populates
from the new columns, with no data of the user's own required. This is
the reason both `build_production.py` and `build_qc.py` write the two
datasets as the default build output, and why the QC verifier checks each
sheet's `Full_Data` completeness column independently.

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
mechanic into columns I and J; the v2.2 Log wiring (see
[DECISIONS.md § v2.2 Transform column wiring](#v22--transforms--unit-space-comparability))
wired column G itself, the last of the two v2.0 reserved columns to go
live — `Order` (F) remains reserved. The function-side equivalent is a
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

### Transform column wiring — Log on Response and Continuous Predictors

**Question:** how should spec column G (Transform) actually be read once
it stops being reserved — where does the `Log` transform get applied, and
what happens to the FE-demeaning wrappers, the Prediction Inputs band, and
Categorical Predictors?

**Resolution, in five parts:**

1. **Modify `Response_Column()` and `X_s()` in place, rather than adding
   `y_s()`-style wrappers.** Every existing consumer of `Response_Column()`
   (the intercept-only fit, Pearson/Spearman R, `Durbin_Watson_By`,
   `Group_Prediction_Interval`, `Group_Mean_At`, and `y_s()` itself) wants
   log-space data the instant Log is declared — there is no consumer that
   needs the untransformed value the way FE demeaning needed the
   untransformed response (the zero-predictor `Intercept_Only_*` branch).
   `X_s()`'s Continuous branch gets the same treatment; the Categorical
   branch never reads `Spec_Transform` at all. A new twin,
   `Constructed_Column_Transforms()`, gives the per-constructed-column
   Log/None flag in the same "structural twin of `X_s`" pattern
   `Constructed_Column_Names()` already uses — needed because a
   Categorical Predictor contributes a variable number of dummy columns,
   so a spec-row-indexed flag cannot align with `X_s()`'s output; every
   dummy column reads `None` unconditionally.
2. **Transform-then-demean composition order, unchanged code.** Because
   (1) puts the Log transform inside `Response_Column()`/`X_s()`
   themselves, `y_s()`/`X_s_Within()` (the v2.1 FE-demeaning wrappers)
   need zero code changes — they automatically demean already-logged
   values, which is the algebraically correct order for a log-linear
   fixed-effects model (`ln(y_it) − mean_g(ln y)`; demeaning before
   logging would take logs of negative numbers).
3. **No `Type` gate on the Response row.** Type (column D) is itself
   hidden-in-place on Response rows by the spec block's own cascading
   conditional formatting (Predictor-only), so gating Log's application
   on an invisible cell would be a trap. `Ln_Positive`'s own `NA()`-on-
   non-numeric behavior is the only guard needed — see
   [DECISIONS.md § v1.2 undersized-sample failure mode](#v12--workbook-hardening)
   for the same "visible failure over silent one" precedent applied here.
4. **Categorical Predictors: disallowed and flagged, not silently
   ignored.** `Log` on a Categorical Predictor is computationally inert
   (the Categorical branch of `X_s()`/`Constructed_Column_Transforms()`
   never reads the flag) AND visibly wrong on the sheet — a red
   conditional-format flag on column G, the same "flag red and instruct,
   never silently switch" precedent the v2.0 Intercept×Categorical
   decision established (see
   [DECISIONS.md § v2.0 intercept coupling](#v20--specification-driven-regression)).
   Silent inertness alone was rejected: a user who mistakenly sets Log on
   a Categorical row deserves a visible correction signal, not a fit that
   quietly ignores their spec.
5. **Prediction Inputs band: raw value in, auto-logged internally.** The
   user always types the real-world value (e.g. actual miles), never
   ln(x) — the same convention a Categorical predictor already uses (a
   raw level string, not a pre-encoded dummy vector) at prediction time.
   This forced one non-obvious fix: the AH19:AH62 prefill cells `INDEX`
   into the AI19 Training Mean spill, which is computed from `X_s()` and
   therefore already log-space for a logged column — feeding that
   straight through `Ln_Positive` a second time at prediction time would
   silently double-log the default (`ln(ln(x))`), not merely leave it
   unback-transformed. Fix: AI19 emits the **geometric mean**
   (`EXP(mean(ln x))`) for a logged column instead of the arithmetic mean
   of the already-logged values — exact and self-cancelling
   (`Ln_Positive(EXP(mean(ln x))) = mean(ln x)`), so the default
   prediction still lands precisely on `X_s()`'s own centroid, unchanged
   from the pre-Log-transform behavior. The Python QC oracle mirrors this
   exact split: `build_spec_design` logs the response/predictor values
   directly, and `tools/inspect_regression_sheet.py`'s harness
   `EXP`s a logged column's mean before writing it into AH, so the sheet
   and the oracle agree on which space each input cell holds.

**Explicit non-goals of this pass:** the unit-space GoF dispatcher
(`Unit_Space_R_Squared` etc., below) and Duan's-smearing back-transformed
predictions (below) are both separately resolved decisions, not
implemented by this wiring — the model fits correctly in log space
end-to-end, but in-sample "Predicted Y" and the prediction outputs are
labelled `(Log)` rather than back-transformed to the response's original
units. Both remain tracked in
[TODOs.md § v3.3](TODOs.md#v33--transforms-remainder).

**Verification:** `tests/test_ln_positive_verification.py` (the
primitive, pure-Python mirror + implementation-shape assertions);
`tests/test_transform_threading.py` — the acceptance test cross-checks a
new spec-driven QC case (`production_lots_log_transform`: `Cumulative_Units`
and `Unit_Cost_BY` with `transform="Log"` declared) against the
pre-existing `production_lots_fixed_effects` case, which points at
`production_lots.xlsx`'s precomputed `"log Cum Units"`/`"log Unit Cost"`
columns — a genuine Crawford/Wright learning-curve model
(\(\ln(\text{unit cost}) = a + b \cdot \ln(\text{cumulative units})\)),
composed with Fixed Effects. The two designs and every downstream
statistic (coefficients, R², SE, residuals, the full CI+PI prediction
block) agree to floating-point precision, confirming the Log wiring
reproduces exactly what the precomputed-column workaround already
delivered. Non-breaking by construction: default `Transform="None"`
produces identical results to before this change (verified against the
full existing spec-case suite, zero edits needed to any case's expected
numbers).

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

### Static template drift — `rebuild_static_sheets.py`

**Question:** the "Regression Instructions" sheet's how-to text (`_ROWS`
in `write_sheet_regression_instructions.py`) needed a correction once
column G's Transform wiring shipped (it still called Transform an unread
placeholder). Fixing the Python source alone turned out not to be
enough — `build_production.py` / `build_qc.py` never execute `_ROWS`;
they only copy the already-baked sheet out of
`templates/static_sheets.xlsx` (see CONTRIBUTING.md → "Static reference
sheets"). Regenerating that template was, until now, a per-sheet manual
step (`python -m lambda_catalog.write_sheet_regression_instructions`,
`python -m lambda_catalog.write_sheet_diagnostic_guide`) with nothing
enforcing it ever ran. This was not a one-off miss: the same gap shipped
stale text once before, when the Fixed Effects role was added to the Role
bullet list and the template regeneration had to be done as a separate
follow-up commit. How should this class of drift be prevented going
forward?

**Resolution:** add `rebuild_static_sheets.py`, a root-level script that
opens `templates/static_sheets.xlsx` once, calls every static sheet's own
`_write_template_sheet(workbook)` — the identical function each
module's individual CLI already calls, so it cannot drift from what
running those CLIs would produce — and saves once. This is now the
standard command for regenerating the template; the per-module CLIs
remain for regenerating a single sheet in isolation while debugging, not
as the primary path. See CONTRIBUTING.md → "Static reference sheets" for
the exact commands.

**Rationale — one command instead of N remembered ones.** The failure
mode both times was the same shape: edit one static sheet's content,
regenerate the template, but either forget the regenerate step
entirely or (with two sheets that changed together) regenerate only one
of them. A single script that rebuilds every static sheet in one Excel
session removes the "which CLI do I need to run" judgment call — there
is one command, and it is always complete. This does not change the
production/QC build's behavior at all: `write_regression_instructions_sheet`
/ `write_diagnostic_guide_sheet` still only ever copy from the template
(see the performance rationale in CONTRIBUTING.md — rebuilding hundreds
of styled cells with COM calls on every build for text that never
changes per-dataset is wasted work), and `rebuild_static_sheets.py` still
requires a real Excel COM engine, same as the CLIs it supersedes as the
default entry point. A test-based drift guard (comparing the template's
baked-in text against `_ROWS` in CI) was considered and rejected for this
pass — it would need Excel-independent parsing of the template's cell
text, which is a reasonable follow-up but a separate piece of work from
fixing the actual command-proliferation problem.

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
[TODOs.md § v3.6](TODOs.md#v36--bivariate--two-sample-claimed-next-minor-after-v35).

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

## v3.0 — Two artifacts, a bounded model context, and the constructor pipeline

The v3.0 decisions respond to [REVIEW.md](REVIEW.md), a standing architecture
review whose findings share one shape: each individual decision was correct,
argued well, and recorded properly, and the cost is in the sum. Findings F1,
F2, F3, F4, F6, and F8 are resolved by the decisions below and struck from that
file. **F5** is struck too, but on different grounds — it was already fixed in
the code when the review was written, and the entry below records the correction
rather than a decision. **F7** (documentation drift) remains open there.

The v3.0 *scope* — which of these ship together — is the one open question, and
it lives in [ROADMAP.md](ROADMAP.md), in the v3.0 milestone entry.
Everything below is resolved.

### Univariate becomes its own workbook

**Question:** the shipped workbook leaves Excel in
`XL_CALCULATION_SEMIAUTOMATIC` — Automatic except Data Tables. Should the build
keep emitting one artifact?

**Resolution:** the build emits **two workbooks**. Univariate Analysis moves to
its own artifact; the Regression workbook keeps every other sheet. **Both
workbooks carry the complete function library** — all 126 catalog functions ship
in both Name Managers. There is no bundling, no dependency closure, and no
per-artifact function subsetting; the workbooks differ only in which sheets they
contain. Splitting lets each artifact set its own calculation mode, and the
Regression workbook returns to full Automatic.

**Rationale:** the semiautomatic mode is forced by the Univariate sheet's six
two-input Data Tables (Weibull, Gamma, Beta × two stages, 20×20 each —
2,400 NLL evaluations per full recalculation). Two consequences shipped with it.
First, every Regression user receives a non-default calculation mode as a side
effect of a sheet they may never open. Second, and far more serious, **Univariate
fit results are stale until the user presses Ctrl+Alt+F9** — the flagship
distribution-fitting sheet silently displays a previous answer. That is a direct
violation of the live-recalculation and visible-failure philosophy, in the one
place the philosophy exists to prevent it. Keeping the two sheets in one file
means one calculation mode has to be wrong for one of them.

**Breakage class: non-breaking, for both artifacts.** The split is packaging
only — no formula, no input cell, and no named range changes meaning. Per the
public-interface definition in [ROADMAP.md](ROADMAP.md), every specification
valid before the split produces the same result after it.

**Mechanism — shipped.** The two build targets now exist:
`build_production.py` emits the Regression artifact (`Lambda_Library.xlsx`,
Regression-only, full Automatic) and `build_univariate.py` emits the Univariate
artifact (`Lambda_Library_Univariate.xlsx`, full Automatic including Data
Tables). Shared build scaffolding lives in `lambda_catalog/build_common.py`,
and the verifier carries a `skip_regression` mode so a Univariate-only workbook
is checked without its absent Regression / Mileage / Production Lots sheets.
Resolves REVIEW.md F4.

### The grid shrink ships as a later release of the Univariate artifact

**Question:** the 2,400-evaluation grid is what forced the calculation mode.
Should shrinking it be bundled with the split?

**Resolution:** no. The shrink ships as a **subsequent release of the Univariate
artifact**, after the split. Breakage class: **MAJOR for the Univariate workbook
version only** — the Scale Min/Max/Step input cells change or disappear, so a
user's saved bounds stop meaning anything. It does not move the Regression
workbook version.

**Weibull and Gamma collapse to one-dimensional searches** by profiling out the
scale/rate parameter in closed form.

Gamma — with β̂(α) = α / x̄:

> ℓ_p(α) = n·α·ln(α/x̄) − n·lnΓ(α) + (α−1)·Σln xᵢ − n·α

Two fixed sample statistics (x̄ and Σln xᵢ) are computed once; only
`GAMMALN.PRECISE` is α-dependent. Starting value from Minka's approximation:
with s = ln(AM/GM), α₀ ≈ (3 − s + √((s−3)² + 24s)) / (12s).

Weibull — with λ̂(k) = ((1/n)·Σxᵢᵏ)^(1/k):

> ℓ_p(k) = n·ln k + (k−1)·Σln xᵢ − n·ln((1/n)·Σxᵢᵏ) − n

Σxᵢᵏ is k-dependent and recomputes per grid point; only Σln xᵢ is fixed.
Starting value from the probability-plot regression of ln(−ln(1−F̂)) on ln x,
with F̂ from the existing `Rank_Fraction()`.

**Beta remains two-dimensional** — both conditional MLEs involve digamma — but
gets a method-of-moments start: on the rescaled data with mean m and variance v,
α₀ = m·(m(1−m)/v − 1) and β₀ = (1−m)·α₀/m, allowing a smaller grid (~12×12).

Total evaluations fall from ~2,400 to ~370.

**Profiling is still genuine MLE.** The profile maximizer *is* the joint
maximizer, so the "MLE without Solver" claim strengthens rather than weakens:
closed form for one parameter, zeroth-order search for the other. This **extends**
the [v1.1 MLE-via-grid reframing](#v11--univariate) rather than replacing it —
that entry said the wall was never "MLE without Solver" but "MLE in closed form,"
and profiling moves one more parameter back across that line.

### Profile-NLL line charts replace the Weibull and Gamma heatmaps

**Question:** the two-dimensional NLL heatmap is the visible artifact of the
grid. What replaces it once Weibull and Gamma search one parameter?

**Resolution:** a **profile-NLL line chart** plotting NLL against the searched
parameter. Beta keeps its heatmap. Both existing grid guards — the boundary-hit
flag and the `IFERROR` sentinel for undefined NLL — carry forward unchanged.

**Rationale:** this is an upgrade in legibility, not a downgrade. The basin, the
interior minimum, and any boundary-hit are all more directly visible in a line
chart than in a one-row color strip, which is what a 1-D search would reduce the
heatmap to. The two guards from
[§ v1.1 two grid guards](#v11--univariate) are what make a boundary hit and an
undefined-likelihood region visible at all, and neither depends on the chart type.

### `Model_Context` — a bounded, materialized cache of spec-derived scalars

**Question:** `[DF_Absorbed]` is carried by 24 functions and `[Allow_Intercept]`
by 48. v2.6 would have added `[Weights]` and v2.7+ a Cluster key. Excel has no
keyword arguments, so every addition converts to positional comma-counting — the
worst current sheet formula is 695 characters ending in a nine-argument call. How
should the engine receive properties of the fit?

**Resolution:** a bounded, fixed-height array holding **only the minimum each
engine function needs that it cannot derive from its own arguments**. It is
materialized once into a spill range on the Regression sheet, and a sheet-scoped
reader `Fit_Context()` *reads that materialized range* rather than recomputing —
the same pattern as `Base_Period_Delta()` reading spec column J.

**Contents — exactly four elements:**

| Element | Consumed by |
|---|---|
| `Has_Intercept` | `SS_Total` (identifies the intercept column), total-df and regression-df display |
| `DF_Absorbed` | the fixed-effects df correction |
| `Response_Transform` | the unit-space dispatcher family |
| `Predictor_Transform` | the unit-space dispatcher family — a `None`/`Log`/`Mixed` summary; per-column detail already lives in `Constructed_Column_Transforms()` |

**Explicitly excluded, with reasons:**

- `Is_Weighted` — no engine needs it. Under the intercept relocation below, the
  weighting is carried by the design matrix itself (column 1 is √w); `SS_Total`
  reads the column, not a flag.
- `n`, `k` — derivable as `ROWS`/`COLUMNS` of arrays the caller already holds.
- Model formula string, headline GoF, error state — **status-block** content, not
  engine plumbing. See the status-block relationship below.

**It is built by the spec block.** Every element is a pure function of the spec
block plus the source data: `Has_Intercept` from the C2 intercept toggle,
`DF_Absorbed` from `Absorbed_Degrees_Of_Freedom()` reading `Spec_Role`, both
transforms from `Spec_Transform`.

**The boundary — why this does not contradict "display derives, never feeds."**
The context block introduces **no new state**. It is a materialized cache of a
pure function of state that already exists. No user types into it. It belongs to
the same family as the computed-display cells J, K, and L in
[ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-ao) — cached for
performance rather than shown for display. Recorded explicitly so it is not later
cited as precedent for putting genuine user input into a computed block: a cell
whose value a user can change is an input, wherever it sits, and inputs belong in
the spec block.

**Signature change.** Engine signatures collapse from
`(X_s, Y, [Allow_Intercept], [Include], [DF_Absorbed])` to
`(X, Y, [Include], [Context])`. A `Model_Context(...)` constructor with its own
optional arguments serves free-form callers outside the sheet.

**`[Include]` is a permanent floor, not a transitional state.** It cannot be
absorbed into the context: the row mask is n×1 and would break boundedness, and
engine functions are workbook-scoped so they cannot default to a sheet-scoped
closure. Four arguments, down from five, with both trailing arguments being
zero-argument closures at the sheet call site.

**Boundedness is the invariant.** `ROWS(Model_Context())` is a build-time
constant and should be asserted in the build.

**Row order is a versioned public contract.** Append only, never insert — the
same discipline as reserved spec columns F and G.

**Two names, by design (v3.0 stage two).** `Model_Context(...)` is the
workbook-scoped constructor — the one definition of the default context, and
what every carrier's omitted-`[Context]` default routes through. `Fit_Context()`
is the sheet-scoped reader — a zero-arg thunk over the fixed materialized range,
and what the ~30 Regression sheet call sites pass so they read the actual
spec-derived context rather than the constructor default. Splitting the names
keeps `Model_Context` unshadowed: a single sheet-scoped thunk named
`Model_Context` would make `Model_Context()` in a sheet cell resolve to the
materialized values while the same token in a carrier's omitted-default
resolved to the workbook constructor — the invisible shadowing the v3.0 release
exists to remove.

**The row order is enforced in one place.** Four workbook-scoped accessors —
`Context_Has_Intercept`, `Context_DF_Absorbed`,
`Context_Response_Transform`, `Context_Predictor_Transform`, each
`=LAMBDA(Context, INDEX(Context, N))` for N = 1..4 — are the only context
reads. Carriers route `Context_Has_Intercept(context_arg)` /
`Context_DF_Absorbed(context_arg)`, never a bare `INDEX(context_arg, N)`, so a
future row insertion changes one accessor instead of 32 hard-coded positional
indices.

**All four rows land together.** Elements 1-2 (the C2 `Allow_Intercept`
toggle, `Absorbed_Degrees_Of_Freedom()`) feed today's engines; elements 3-4
(the response transform, and the `None`/`Log`/`Mixed` summary over the included
Continuous predictors) have no engine reader until the v3.3 unit-space
dispatcher. They are populated from the spec block now, not left as `"None"`,
because the row order is the contract that is expensive to change later. An
error in an unconsumed row is contained: the engines read only elements 1-2
through the accessors, so a bad spec name surfaces as a visible cell error
(caught by the headless verifier) without shifting a fitted number.

**Performance rationale.** Names whose Refers To is a formula are re-evaluated at
every use site; Excel does not memoize them. A materialized cell is computed once.
The recomputation eliminated is not small: `Absorbed_Degrees_Of_Freedom()` calls
`Dummy_Levels(Fixed_Effects_Column(), "", Sample_Include())` — a UNIQUE/SORT over
the included sample — at all 24 current `[DF_Absorbed]` call sites.

Resolves REVIEW.md F1.

### The context is NOT the Model Comparison interface

**Question:** the v2.3 Model Comparison sheet needs a fixed-position interface on
each Regression sheet. Is the context block that interface?

**Resolution:** no. Two distinct objects, and an earlier draft that proposed
collapsing part of v2.3 into the context block is **withdrawn** by the minimality
decision above.

- **`Model_Context`** — engine plumbing. Minimal, bounded, order-versioned.
- **The status block** — display and cross-sheet interface. Already fixed-height
  and fixed-position, and already what the v2.3 anchor design was written against.

The v2.3 `Comparison_Anchor` / `Comparison_Headline_GoF` /
`Comparison_Prediction_Output` design stands unchanged (see
[§ v2.3](#v23--model-comparison-sheet)). One relationship to record: status-block
cells that overlap the context (`Has_Intercept`, absorbed df) should **read the
materialized context** rather than recompute — "display derives, never feeds,"
applied one layer up.

### Row filtering stays in the engine, never in the constructor

**Question:** moving the row mask into the constructor would remove `[Include]`
from every signature and apply the mask once per recalculation instead of roughly
thirty times. Should it move?

**Resolution:** rejected. Row filtering stays in the engine.

**Rationale, in three parts:**

1. It breaks the `x_s()` full-height row-mask contract
   ([§ v2.0](#v20--specification-driven-regression)), which exists precisely to
   prevent double-filtering against the engine's own mask application.
2. Full-height output is what keeps every spilled array **row-aligned with the
   source table**. `Row_Labels()`, the residual-output zone, and the materialized
   design matrix all read across against the source rows. A filtered constructor
   would silently break that alignment.
3. The materialized design matrix is specified full-height for the same reason; a
   filtered constructor would contradict it.

### `Sample_Include()` is materialized instead

**Question:** the performance argument for filtering in the constructor was real
even though the mechanism was wrong. How is it recovered?

**Resolution:** materialize the mask. `Sample_Include()` is, like the context, a
pure function of the spec block plus the source data. It cannot go *inside* the
context (n×1 breaks boundedness), but it becomes its own materialized full-height
spill range with a name, and `Sample_Include()` reads the range rather than
recomputing.

**Rationale:** this eliminates roughly thirty recomputations of the role-aware
completeness logic — which tests numeric-ness across the response and every
included Continuous Predictor over every row — while leaving the contract
untouched. The Model Construction sheet already spills `Sample_Include()`
full-height at Q4, so the pattern is proven; this promotes it to production on the
Regression sheet. It is placed immediately left of the design-matrix terminal
zone, per the zone layout below.

### Intercept relocation into the constructor

**Question:** `[Allow_Intercept]` is the largest optional-argument carrier in the
catalog. Can the intercept stop being an engine argument?

**Resolution:** yes. The intercept column moves out of `Design_Matrix` and into
the design-matrix constructor. `Design_Matrix` stops synthesizing it.

**What it costs — the honest count.** 48 functions carry `[Allow_Intercept]`.
**37 are pure pass-through; 11 branch on the flag** — `Total_Degrees_Of_Freedom`,
`SS_Total`, `Design_Matrix`, `AIC`, `BIC`, `AICc`, `Coefficients`,
`SE_Coefficients`, `Beta_Weights`, `Prediction_Interval`, and
`Group_Prediction_Interval`. `Has_Intercept` therefore does not disappear; it
survives in roughly seven places, as an identifier rather than an arithmetic
switch: identify column 1 for `SS_Total`, `n − Has_Intercept` for total df,
`COLUMNS(X) − Has_Intercept` for the ANOVA regression-df display convention, the
coefficient-vector `DROP` in `Beta_Weights` and `Group_Prediction_Interval`, and
the `x_new` alignment in `Prediction_Interval`.

`Regression_Degrees_Of_Freedom(X_s)` is `COLUMNS(X_s)` and takes the predictor
matrix, so it never counted the intercept and needs no change.

**Rationale:** the relocation is still worth doing at eleven functions. It is what
makes the `Model_Context` collapse possible, and it converts WLS from an engine
argument threaded through 24 more functions into a constructor concern — see the
v2.6 supersession below. The alternative is the accretion trajectory REVIEW.md F1
describes, which has no stopping rule.

### `SS_Total` redefined as the intercept-only residual sum of squares

**Question:** with the intercept a column rather than a flag, what does
`SS_Total` compute?

**Resolution:** the residual sum of squares from the intercept-only model — the
projection of y off whatever the intercept column actually is:

> SS_Total = ‖y‖² − (c′y)² / (c′c)

This collapses three cases into one formula:

| Intercept column c | Result | Equals |
|---|---|---|
| ones | Σy² − (Σy)²/n | `DEVSQ(y)` |
| absent | ‖y‖² | `SUMSQ(y)` |
| √w | Σwy² − (Σwy)²/Σw | Σw(y − ȳ_w)² |

The decomposition SS_Total = SS_Regression + SS_Residual continues to hold in all
three cases.

**The WLS trap, recorded with the algebra.** `DEVSQ(√w ⊙ y)` is **not** the
weighted total sum of squares — it centers on mean(√w·y) rather than on ȳ_w. A
naive "scale everything by √w" implementation would leave SS_Total, and therefore
R², silently wrong under WLS, with no error anywhere and a plausible-looking
number in the cell. The projection form above is correct by construction. This is
exactly the class of silent misfiring the library exists to prevent, which is why
it is recorded here rather than in a code comment.

### The LINEST `const` trap

**Question:** `Coefficients` and `SE_Coefficients` both call
`LINEST(FILTER(Y, filt), FILTER(X_s, filt), allow_arg, …)`. What happens to that
call when the design matrix arrives with the intercept already in it?

**Resolution:** `const` must be hard-`FALSE` at both sites, and the coefficient
reassembly simplifies accordingly.

**Rationale — this is the sharpest edge in the relocation.** `allow_arg` is not
merely passed along in these two functions: it *is* LINEST's third argument,
`const`, which tells Excel whether to fit its own intercept. If the constructor
prepends a column of ones and `const` is left `TRUE`, Excel fits a second
intercept on top of it — two perfectly collinear terms, an exactly singular Gram
matrix, and a result that is wrong rather than absent.

Two consequences follow. `Coefficients` currently unwinds LINEST's
reverse-order output with `INDEX(ls, 1, k + 1)` for the intercept plus
`CHOOSECOLS(ls, SEQUENCE(1, k, k, -1))` for the slopes; with `const = FALSE` and
the intercept already a column of `X`, LINEST returns exactly `COLUMNS(X)`
coefficients and the branch collapses to the reversal alone. `SE_Coefficients`'
`naive_df` rescaling changes with it, because LINEST computes n − k under
`const = FALSE` where it computed n − k − 1 before — which is the correct df once
the intercept is counted in k.

The same arithmetic reaches the information criteria: `AIC`, `BIC`, and `AICc`
compute `p = Regression_Degrees_Of_Freedom(X_s) + IF(allow_arg, 1, 0) + absorbed_arg`,
and the `+ IF(allow_arg, 1, 0)` term must go once `COLUMNS(X)` counts the
intercept itself.

Recorded beside the WLS trap above because it is the same failure class: a change
that looks like a pass-through, is not, and fails silently in a plausible
direction.

### Pipeline order is a hard constraint

**Question:** in what order do the construction stages apply?

**Resolution:** one fixed order, recorded in
[ARCHITECTURE.md § 4a](ARCHITECTURE.md#4a-the-constructor-pipeline):

> encode → transform → demean → intercept → weight

**Rationale:** a column of ones demeaned by group is a column of zeros, giving a
singular Gram matrix. The current code is safe *by accident* — `Design_Matrix`
prepends the intercept to the already-demeaned `X_s_Within()` output, so the
ordering is enforced by which function calls which. Once the constructor owns the
intercept, that accident disappears and the ordering must be stated, or someone
will rediscover it as a singular-matrix bug. The transform-then-demean half of the
order was already settled at
[v2.2](#v22--transforms--unit-space-comparability) for a different reason —
demeaning before logging would take logs of negative numbers — so this entry
extends an existing constraint rather than introducing a new one.

### Weighted fixed effects — out of scope, not overlooked

**Question:** under WLS with Fixed Effects, should the within transformation
demean using weighted group means?

**Resolution:** yes, it should — and it is **not part of v3.0**. Recorded here so
a later WLS + FE release does not inherit an unstated assumption that plain group
means are correct under weighting.

### One constructor pipeline replaces the constructor name fork

**Question:** `X_s()` and `X_s_Within()` are two names for stages of one pipeline,
and the sheet must currently know which each call site wants — fit statistics take
`X_s_Within()` while `GVIF`, `Generalized_Tolerance`, `Pearson_R`, `Spearman_R`,
`Skewness`, and `Kurtosis` take `X_s()`. The distinction is statistically correct
and entirely invisible in the names, and nothing enforces it. Adding weighting and
two-way absorption would produce a cross product of variants. What replaces the
fork?

**Resolution:** one constructor applying the declared stages in the fixed order
above, plus one explicit escape hatch returning pre-demeaning columns for the
predictor-summary zone that legitimately wants them:

| Name | Replaces | Stages applied |
|---|---|---|
| `Design_Columns()` | `X_s_Within()` | encode → transform → demean → intercept → weight |
| `Design_Response()` | `y_s()` | transform → demean → weight |
| `Predictor_Columns()` | `X_s()` | encode → transform only — the escape hatch |
| `Response_Column()` | *(unchanged)* | transform only |

**Rationale for the names.** The problem being fixed is that the distinction is
invisible, so the names have to carry it. `Design_Columns()` is what enters the
model; `Predictor_Columns()` is the predictors as columns, before the model-fitting
stages. Call sites read correctly without cross-referencing —
`R_Squared(Design_Columns(), Design_Response(), …)` against
`GVIF(Predictor_Columns())` — and the "design columns" vocabulary is shared with
the Design Columns audit column and the Constructed Design Matrix zone below, so
one term covers all three. The names satisfy the
[§ 1 naming convention](ARCHITECTURE.md#1-naming-convention): full English words,
Title_Case_With_Underscores, no abbreviations.

Resolves REVIEW.md F2.

### Interactions are declared with two spec columns

**Question:** the spec block is one row per source column, and an interaction term
is not a column. It fits neither declared axis — Predictor Type is permanently
closed, and Role describes what a column *is*. Where do interactions live?

**Resolution:** two new spec-block columns — an **Interaction Term** column
(naming the other operand) and an **Interaction Operation** column — both
defaulting to none/blank.

**Built at v3.0 stage three, as M and N**, appended to the block so A–L keep
their addresses. They ship reserved: the dropdowns, the marginality flag, and the
reciprocal-declaration flag are live, but `Spec_Interaction_Term` and
`Spec_Interaction_Operation` are read by no defined name and no cell formula
until the wiring release. Conditional-formatting expressions are neither, which
is what lets the flags be useful while the bands stay unread — the validation
that catches a typo is worth having from the day the column exists, not the day
the constructor learns to build it.

**Rejected alternative: a second spec section below the per-column block.** The
source table's column count is variable, so anything positioned below the
per-column block has no fixed address. The spec block auto-extends as a real Excel
Table (`SpecTable`); a second section beneath it would be pushed down by any table
that grows, and every formula referencing it would need a dynamic offset. Recorded
in the supersession log.

**Rationale for resolving it now, ahead of the feature.** REVIEW.md F6 notes this
is the only finding that is irreversible if deferred: every other pending feature
can be absorbed additively, but an interaction mechanism cannot be retrofitted the
way column G was, because it is not a per-column property. The representation
decision is therefore worth making before the next layout touch, independent of
when interactions are actually implemented — which is exactly what the
reserved-spec-column policy exists to enable. Resolves REVIEW.md F6.

**One correction to F6's premise.** That finding states `Interact(x1, x2)` "exists
as a standalone catalog function." It does not — `Interact`, `Model_Matrix`, and
`Dummy_Column` are all *specified* in
[ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy) and listed
as v2.2 work items in [TODOs.md](TODOs.md), but none is in
`lambda_functions.json`. The finding's conclusion is unaffected; its evidence is
corrected here.

### Interaction operation vocabulary — closed, with a symmetry attribute

**Question:** which operations may an interaction declare, and what happens when
both operands declare each other?

**Resolution:** the operation list is **closed**, in the same sense as Predictor
Type, and each operation carries a symmetry attribute that determines whether a
reciprocal declaration is legitimate:

| Operation | Symmetry | Reciprocal declaration (B on A as well as A on B) |
|---|---|---|
| Product | Symmetric | Produces a duplicate column → singular Gram. **Flag red.** |
| Difference | Antisymmetric | Produces the exact negative → equally collinear. **Flag red.** |
| Ratio | Asymmetric | Legitimate and distinct. **Allowed.** |

**Rationale:** flag rather than silently deduplicate — the same "flag and instruct,
never silently switch" precedent as the
[v2.0 intercept × categorical case](#v20--specification-driven-regression) and the
[v2.2 Log-on-Categorical case](#v22--transforms--unit-space-comparability).
Silently dropping the duplicate would make the constructed design matrix disagree
with the declared spec, which is the one thing the spec block exists to prevent.

### Self-interaction is allowed and documented

**Question:** an Interaction Term pointing at its own row — accident or feature?

**Resolution:** allowed and documented. With Operation = `Product` it yields x².
This is the documented way to get a quadratic term.

**Rationale:** it is useful in cost work, unambiguous in meaning, and falls out of
the layout anyway. Allowing it explicitly is better than leaving it as an
accidental edge case whose behavior nobody has decided.

### Operand Role and Include semantics — four cases

**Question:** what happens when an interaction's target is not an included
Predictor?

**Resolution:**

| Case | Behavior |
|---|---|
| Target Role is Omit, Response, Filter, Identifier, or Fixed Effects | **Error.** Only a Predictor can be an operand. |
| Declaring row has Include = FALSE | Contributes nothing; the interaction is excluded. Consistent with cascading relevance. |
| Target is a Predictor with Include = FALSE | **Allowed, flagged amber.** |
| Both included | Normal. |

**Rationale for the amber case:** an interaction without its main effect is a
marginality violation — usually a specification error, but occasionally deliberate.
Blocking it would be the library deciding a modeling question on the user's behalf,
which is not its role; flagging it surfaces the issue without overriding the
declaration. Amber rather than red because, unlike the reciprocal-Product case, the
resulting model is estimable.

### Two-way interactions only

**Question:** should nested or three-way interactions be expressible?

**Resolution:** no. One operand per spec row; (A×B)×C is not expressible.

**Rationale:** this is the right limit for a spreadsheet library — three-way
interactions on a categorical pair already produce column counts no one can audit
by eye, and the spec block's readability is the feature. Recorded as a **deliberate
decision** rather than left as an emergent property of the one-operand-per-row
layout, so a future reader does not mistake the limit for an oversight.

### Interactions make the Design Columns audit column required

**Question:** [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-ao)
notes that the gap column right of the spec block "visually reserves a future
Design Columns slot." Does interaction support change its status?

**Resolution:** it must now be built. The reserved slot becomes a required column.

**Rationale — width.** Continuous × Categorical broadcasts to L−1 columns;
Categorical × Categorical gives (L₁−1)(L₂−1). Status × Country on the WHO data is
155 columns from a single spec row. **Interactions, not main effects, are where the
design matrix explodes.** The audit column is the only place a user can see that
one dropdown added 155 columns, and it supplies the pre-flight width number the
guard below needs.

**`Constructed_Column_Names()` needs an interaction convention.** R's colon form —
`GDP:Schooling`, level-qualified as `GDP:StatusDeveloping`. The constructor twin
must stay structurally identical to the constructor, per the existing rule that
names and columns cannot disagree.

### Materialize the design matrix as the terminal zone

**Question:** the constructor is called inside every engine function, so the
design matrix is rebuilt roughly thirty times per recalculation and is never
visible. Should it be materialized on the sheet?

**Resolution:** yes. Materialize the constructed design matrix at the far right of
the Regression sheet, **full-height**, preserving the `x_s()` row-mask contract —
the engine keeps doing the row filtering.

**Persistent benefits, beyond one-time performance:**

1. It restores the visible design matrix that v2.0 gave up; the status block was
   the consolation prize for construction moving inside a LAMBDA.
2. The Model Construction sheet's V/W zones already prove the pattern; this
   promotes a QC feature to production.
3. v2.4 bootstrap resamples the design matrix per draw, and reconstruction per
   draw is not viable.
4. Two-way FE absorption is iterative (alternating projections with a `[passes]`
   argument); recomputing it inside every engine call is prohibitive.

**Record the cost honestly.** Materialization is a tradeoff, not a pure win. On the
WHO data with Country as a Categorical Predictor, the design matrix is roughly
2,938 × 156 ≈ 458,000 live cells that recalculate on any input change, and the used
range and file size grow accordingly. This is still far cheaper than reconstructing
the matrix inside thirty engine calls, but it is not free.

**Dependency:** the constructor pipeline must be resolved first, or two variants of
a soon-to-change architecture get materialized.

### The width guard — pre-flight, two thresholds

**Question:** the materialized zone's width is unbounded and one dropdown away.
What stops it running off the sheet?

**Resolution:** a pre-flight check computed from the **Σ Design Columns audit
total**, not from `COLUMNS(Design_Columns())`, with two thresholds.

- **Hard error** at the sheet-width bound: `16,384 − (last_chart_column + 5)`,
  where the five columns are three gutters plus the `Model_Context` and
  `Sample_Include` columns. Computed from the layout constants, never hard-coded.
  Surfaced as a spec-block-area error flag and in the status block's error state.
- **Soft warning** at **k = 200 constructed columns, or 500,000 materialized cells
  (n × k), whichever trips first.**

**Rationale for pre-flight.** Constructing a 16,000-column array in order to
discover it does not fit is precisely the failure being prevented. The audit column
gives the number before the constructor runs.

**Rationale for the soft thresholds.** `Gram_Inverse` is O(k³) in `MMULT`, so the
practical wall is in the hundreds, not thousands — a model that reaches 16k columns
has been unusable for a long time already. The two numbers are calibrated against
the largest sane shipped example: WHO with Country as a Categorical Predictor is
k = 156 and ≈ 458,000 cells, deliberately just under both, so the worked example
does not trip its own warning while anything materially larger does. The
cell-count trigger exists because materialized-cell count, not just k, is now part
of the cost.

### Materialization zone layout

**Question:** where exactly do the materialized artifacts sit, and in what order?

**Resolution:** all materialized artifacts live at the far right of the Regression
sheet, each in its **own outline group separated by a thin ungrouped gutter
column**, following the sheet's established pattern. The group begins after the
diagnostic-chart columns:

```
… existing zones … │ charts │ gutter │ Model_Context │ gutter │ Sample_Include │ gutter │ Constructed Design Matrix →
                                        (4 × 1)                  (n × 1)                  (n × k, unbounded)
```

**Naming.** The terminal zone is the **Constructed Design Matrix**, not the "Model
Construction" zone — `Model Construction` is already a sheet name
(`write_sheet_model_construction.py`) and the two must stay distinguishable.

**Ordering rule:** the materialized zones run in **increasing width and terminate
in the unbounded zone**. This single rule covers both the "nothing may ever be
placed to the right of the Constructed Design Matrix" commitment and the question
of where any future bounded materialization goes. It lives in
[ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone) as a pattern a
new feature must honor. It also supersedes REVIEW.md F3's framing: the sheet now
has an explicit terminal boundary and a stated rule for what may be added, which
is the eviction mechanism that finding said was missing.

**Collapse behavior differs by zone.** Three separate groups exist precisely so
they collapse independently. `Model_Context` and `Sample_Include` are one column
each and ship **expanded**; the Constructed Design Matrix ships **collapsed by
default**, because an unbounded-width zone that cannot be collapsed is a scrolling
hazard.

**The first gutter is structural, not cosmetic.** Charts anchored over columns
inside a collapsed outline group get squashed. The gutter after the chart columns
is what keeps the diagnostic-chart anchors outside every collapsible group.

**The chart footprint needs an explicit bound.** `_C_AW` is currently the chart
*anchor* column, not the chart *extent*: the seven diagnostic charts are floating
objects tiled in a 4×2 grid roughly 640 points wide from AW's left edge, and four
further content columns (AX–BA) carry the chart title and axis-label formula cells.
Nothing records where that footprint ends. Introduce a named constant for the last
chart column with a build assertion that no chart extends past it. Without it, a
chart resize silently overlaps the context block, and the zone start column cannot
be computed reliably — which the width guard above depends on.

**All zones share a first data row.** Read-across is the point — the mask value
beside its design-matrix row, both aligned to the source table rows, with the
gutters as visual separators. Assert the shared start row in the build rather than
leaving it to layout constants.

**Build details for implementation:**

- Column widths cannot be set per-column across an unbounded zone. Set a generous
  fixed block of narrow columns in a single range call (the univariate grid's
  width-6 pattern is the precedent) and let the remainder take default width.
- Each zone needs a name over its spill anchor (`Regression!$XX$n#`).
- The `Model_Context` row order remains append-only.
- Gutter columns must remain **ungrouped**, or the groups merge.

### Materialization lands in two steps — `Model_Context` now, `Sample_Include` deferred

**Question:** the layout above shows both `Model_Context` and `Sample_Include`
materialized. Does stage two land both?

**Resolution:** no. `Model_Context` materializes in stage two; `Sample_Include`
is placed at its final §4b position as a **reserved placeholder**, and
promoting the live `Sample_Include()` closure to a thunk over a materialized
spill is a separate, Excel-verified follow-up.

**Rationale:** the two artifacts have different risk profiles.

- `Model_Context` is **bounded** — `ROWS(Fit_Context())` is a 4-row
  build-time constant. The sheet-scoped reader `Fit_Context` reads a **fixed
  range** (`$BO$2:$BO$5`), so the `#` spill operator never enters a `LAMBDA`
  defined-name `RefersTo`. That combination is the only unproven one in this
  workbook, and the fixed-range read sidesteps it entirely. The materialization
  is safe to land blind.
- `Sample_Include` is **unbounded** — `n × 1`, sized to the source table. A
  thunk over it requires `#` inside the `RefersTo`, the combination not used
  anywhere else. A wrong guess breaks the row-mask contract that keeps every
  spilled array row-aligned with its design-matrix row, and that breakage is
  only catchable by the spec-driven Excel gate — the headless suite does not
  exercise it. `Sample_Include` is also a pure performance optimization: the
  collapse into `[Context]` does not depend on it, the live closure already
  works, and nothing about stage three's layout needs it materialized. So it
  lands where it can be Excel-verified, not blind.

**Consequence for stage three, as it played out:** the `Sample_Include` column
already occupied its final §4b position, so stage three added the Constructed
Design Matrix zone and its width guard behind the existing reserved column with
no relocation. The placeholder is labelled "reserved" on the sheet and carries a
cell comment documenting the deferral so it is not mistaken for an oversight —
and stage three gave the terminal zone exactly the same treatment for the same
reason, so the band now carries two reserved positions and one live spill. The
promotion lands at v3.2 alongside the design matrix's own materialization, since
both need the same Excel-verified `#`-inside-a-`LAMBDA` answer.

### Versioning across two artifacts

**Question:** [ROADMAP.md](ROADMAP.md) defines the public interface as "the user's
inputs to the workbook" — singular. Two emitted workbooks break that definition.
What replaces it?

**Resolution:** a **single library version** covering the shared function catalog,
plus a **per-workbook version** covering each artifact's sheets and input surface.

**Rationale:** both workbooks carry the identical complete function library, so a
function change is genuinely a shared event and should move one number. The input
surfaces differ entirely, so a Univariate layout change must not move the number a
Regression user reads as the answer to "do my existing inputs still work?"

**The `Breaking?` flag attaches to the workbook version, not the library version.**
It answers a question about a user's saved inputs, and inputs are a property of the
workbook's sheets. A library-version bump that adds a function breaks nothing.

Consequences, recorded so the two-number scheme is unambiguous at its first two
uses: the Univariate split is **non-breaking for both artifacts**. On the
Regression side the split ships *bundled into the 3.0.0 release* — the 3.0.0
MAJOR marks the architectural milestone (the bounded `Model_Context`, the
constructor pipeline, and the split itself), not a public-interface break; every
specification valid before 3.0.0 produces the same result after it, so the split
alone would have moved no version. On the Univariate side the split *is* the
1.0.0 initial release, the artifact's first existence. The grid shrink is
**MAJOR for the Univariate workbook version only** and does not move the
Regression workbook version. The full display and changelog conventions are in
[ROADMAP.md § Versioning](ROADMAP.md#versioning--release-conventions). Resolves
REVIEW.md F8.

### `PRESS` correctly omits `[DF_Absorbed]`

**Question:** `PRESS` does not carry `[DF_Absorbed]` but `QQ_Correlation` does.
The two sit in adjacent zones of the same sheet and the asymmetry is not legible
from the signatures. Is it correct?

**Resolution:** correct, and the reason is mechanical rather than a judgment call.
`PRESS` is `SUMSQ(LOOCV_Residual(…))` — a sum of squared leave-one-out residuals,
each of which is `eᵢ / (1 − hᵢ)`. Neither the residual nor the leverage depends on
a degrees-of-freedom count, so there is no term for absorbed df to enter.
`QQ_Correlation` calls `Scaled_Residuals_Ranked`, which divides by an estimate of
σ computed on residual df — so absorbed df changes its value.

The rule this generalizes to: **a statistic needs `[DF_Absorbed]` exactly when it
divides by a residual-df-based variance estimate.** Recorded in each function's
JSON `notes` field so the asymmetry is legible from the catalog sheet without
reading both formulas. Resolves the REVIEW.md Minor item.

### `AIC` and `GoF_AIC` are deliberately distinct

**Question:** `write_sheet_univariate.py` references both `AIC`/`BIC` and
`GoF_AIC`/`GoF_BIC`. Is that a naming collision?

**Resolution:** deliberate, and there is no collision in the code. The Univariate
sheet **only ever calls** `GoF_AIC(nll, k)` and `GoF_BIC(nll, k, n)`; the strings
`"AIC"` and `"BIC"` that also appear there are column-header labels in the
fit-comparison table, not function calls.

**Rationale:** the two families take different arguments because they answer
different questions. `AIC(X_s, Y, [Allow_Intercept], [Include], [DF_Absorbed])`
computes a regression information criterion from the residual sum of squares, and
its df argument counts model parameters including absorbed fixed effects.
`GoF_AIC(nll, k)` takes an already-computed negative log-likelihood and a
parameter count, because a distribution fit has no design matrix and no absorbed
df. Collapsing them would mean one name whose argument list changes meaning by
context, which is exactly the ambiguity the naming convention exists to prevent.

### The spec block is implemented once, not twice

**Question:** REVIEW.md F5 observes that `write_sheet_regression.py` (1,862
lines) and `write_sheet_model_construction.py` (1,512 lines) "each implement a
spec block," so "a layout change touches both writers." With v3.0 adding two
spec columns, the Design Columns audit column, and the materialization zone,
does that double cost need paying — or unwinding — as part of this release?

**Resolution:** neither. The premise is false. There is **one** implementation.
`write_sheet_regression.py` imports the spec-block writers from
`write_sheet_model_construction.py` and calls them:

```python
from .write_sheet_model_construction import (
    _set_sheet_scoped_names as _set_spec_scoped_names,
    _set_spec_block_column_widths,
    _write_intercept_control,
    _write_spec_block,
    _write_spec_feedback,
    # … plus every _C_* column constant and formula string
)
```

`write_sheet_regression.py`'s own module docstring states the intent: *"the
spec-block writers are imported from write_sheet_model_construction so the two
sheets can never drift."* Separately, the Model Construction **sheet** is
deleted by both builds — `_delete_sheet_if_present(workbook, "Model
Construction")` in `build_production.py` and `build_qc.py` — so only one spec
block ships at all.

**Consequence for v3.0 scope:** the interaction columns, the audit column, and
the materialization zone each land in **one** writer. F5 is not a cost of this
release, and the shared-import structure is part of what makes the recommended
scope affordable.

**Rationale for recording a non-decision.** Nothing was decided here — the fix
predates the review that reported the problem. It is recorded because the wrong
version was load-bearing twice: F5 was triaged as medium and "expensive after
v2.3," and an earlier draft of this v3.0 pass argued the release made it *worse*
on the same reasoning. Both inferred coupling from two large files without
reading the import. Writing down that the coupling does not exist is what stops
a third round.

**What remains is a naming problem.** `write_sheet_model_construction.py` no
longer writes a shipped sheet; it is the spec-block component library the
Regression sheet is built from, and it still carries the name of a sheet both
builds delete. Renaming it — and dropping the unreachable
`write_model_construction_sheet()` / `main()` standalone-CLI path — is tracked
in [TODOs.md](TODOs.md) as cosmetic follow-up. Deliberately **not** dropped:
`_write_audit_row` and `_write_filtered_zones`, which are the working reference
implementations of the Design Columns audit column and the V/W filtered-display
pattern that this release promotes to production.

### The standalone user-callable layer is documented, not orphaned

**Question:** 30 catalog functions are called by no sheet writer. With the full
library in both workbooks this raises no packaging question, but do they look
abandoned?

**Resolution:** they are the **standalone user-callable layer** and the
`LAMBDA_functions` catalog sheet documents them as such, rather than leaving them
to read as orphans.

**Rationale:** the count is a property of the sheets, not of the functions. A
function like `Correlation_Matrix`, `Lag_By`, or `Descriptive_Statistics` exists so
a user can call it in their own cell on their own data — that is the library's
primary purpose, and the pre-built sheets are demonstrations of it, not the whole
of it. `ARCHITECTURE.md § 5` already records that the Data Transformation family
serves "double duty" as constructor internals and standalone transforms; this
extends the same framing to the rest. (The count is a text-match of function names
against `write_sheet_*.py` and moves as sheets change; it is illustrative, not a
tracked invariant.)

### v3.0 shipped in stages; the layout break lands last

**Question:** the scope entry in [ROADMAP.md](ROADMAP.md) settled *what* v3.0
contains but not how it lands. All of it in one change is a diff nobody can review
against a workbook nobody can rebuild in CI.

**Resolution:** the release landed in four reviewable pull requests, in dependency
order, all under **v3.0.0** (2026-08-02). Staging the work did not stage the
version: the four changes answer one question together, so they carry one number.

| Stage | Contents | Status |
|---|---|---|
| 1 | The constructor pipeline and the intercept relocation | **Shipped** (#148) |
| 2 | `Model_Context` — `[Has_Intercept]` and `[DF_Absorbed]` collapse into `[Context]`; the two-name split (`Model_Context` constructor / `Fit_Context` reader) keeps `Model_Context` unshadowed; four `Context_*` accessors make the row order a contract enforced in one place; all four context rows materialized (1-2 feed the engines, 3-4 populated from the spec block for v3.3); `Sample_Include` placed at its final §4b position as a reserved placeholder (thunk materialization deferred to an Excel-verified follow-up) | **Shipped** (#150) |
| + split | Univariate Analysis becomes its own workbook; the Regression workbook returns to full Automatic | **Shipped** (#151) |
| 3 | Layout — interaction spec columns M/N reserved, the Design Columns audit column and its pre-flight width guard, the Constructed Design Matrix zone | **Shipped** |

**Rationale:** the order is forced by the dependencies the scope entry already
names — the bounded context requires the intercept relocation, which requires the
pipeline order, and the materialization zone requires the constructor to be
settled. Stage one also has a verification property the others do not: **no number
moves.** The relocation changes where the intercept is created, not what is
fitted, so the spec-driven QC pass must report zero mismatches across all twelve
cases. Any mismatch is a bug rather than an expected delta, which is the cleanest
gate available for the stage that touches the most functions.

Stage three went last for the same reason and gained a second benefit from it.
With the numbers pinned by stages 1-2, the layout stage's own verification gate
becomes unambiguous: it moves columns and nothing else, so any mismatch it
produced could only have come from the layout. It passed.

### The v3.0 break is an ADDRESS break, not a meaning break

**Question:** 3.0.0's `Breaking?` flag. Stages 1-2 and the split are non-breaking
by construction; stage three moves columns. Does that make the release breaking,
and if so, breaking in what sense?

**Resolution: yes, and the distinction is worth stating rather than collapsing.**
The flag is **Yes**, because a workbook built against 2.0.0 has formulas that stop
pointing where they used to. But two very different things could have earned that
flag, and only the milder one happened.

A **meaning break** is what 2.0.0 was: the same cell still exists, still holds the
user's value, and now means something else. Nothing detects it; the model just
quietly computes a different answer. An **address break** is what 3.0.0 is: no
cell changes meaning, and no fitted number moves — cells are simply somewhere
else. A formula pointing at a moved cell reads the wrong thing *visibly*, or
`#REF!`s outright.

Stage three keeps the break in the milder category deliberately, by **appending**
its three columns rather than inserting them. A–L keep both their letters and
their meanings, so a saved specification — the thing users actually invest in —
survives untouched. The cost is that two *inputs* (M/N) now sit to the right of
the J/K/L computed displays, reading slightly against the block's
inputs-then-displays order. That was judged the cheaper of the two: the
alternative shifts eight columns to preserve a reading convention. What does move
is every zone right of the spec block, three columns over — the Alpha input from
Y12 to AB12, the Prediction Inputs band from AH to AK, the Residual Output from AK
to AN. A user who only fills in the spec block notices nothing; a user with their
own formulas against this sheet has to re-point them, and the Version History entry
names the moved anchors so they can.

**The break was made once, on purpose.** REVIEW.md's sequencing note observed that
the interaction columns, the audit column, and the materialization zone "all want
the same breaking change — resolving them separately spends three layout breaks
where one would do." All three landed in the single stage-three change, and each
ships **reserved**: M/N are validated and flagged but read by no constructor, and
the Constructed Design Matrix zone is positioned, bounded, and guarded but not yet
filled. Wiring them (v3.1, v3.2) is then a formula change against columns that
already exist, which is exactly the reserved-column pattern column G went live
under at v2.2 — additive MINOR work with no second break behind it.

### `R_Squared` is the third LINEST `const` site — and the one that fails silently

**Question:** the LINEST `const` trap above names `Coefficients` and
`SE_Coefficients`. Is that the whole exposure?

**Resolution:** no. `R_Squared` is a third call site, and it is the most dangerous
of the three. It read `INDEX(LINEST(Y, X, allow_arg, TRUE), 3, 1)` — LINEST's own
R². **Under `const = FALSE`, that cell holds the *uncentered* R²**, computed
against Σy² rather than DEVSQ(y). Flipping the flag and leaving the read in place
would have silently changed R², Adjusted R², Multiple R, and the F-statistic for
every model with an intercept — no error, no `#REF!`, just different numbers.

`R_Squared` is therefore derived from the sums of squares instead, which also
breaks the cycle that made the old definitions self-referential (`SS_Residual` was
`SS_Total × (1 − R²)` while `R²` came from LINEST):

```
SS_Residual(X, Y, [Include])           = INDEX(LINEST(y, X, FALSE, TRUE), 5, 2)
SS_Total(X, Y, [Has_Intercept], [Inc]) = ‖y‖² − (c′y)²/(c′c)
R_Squared                              = 1 − SS_Residual / SS_Total
SS_Regression                          = SS_Total − SS_Residual
```

LINEST's fifth output row holds Σe² regardless of `const`, so `SS_Residual` can
stand on its own. The result is value-preserving in both states: with an intercept
the denominator is DEVSQ(y) and the expression reproduces LINEST's centered R²;
without one it is SUMSQ(y) and reproduces the uncentered R² LINEST already
returned. **Recorded because it was missed by the design pass**, and it belongs
beside the WLS `DEVSQ` trap and the `const` trap as the same failure class: a
change that looks like a pass-through, is not, and fails in a plausible direction.

### The honest count: 48 → 13, not 48 → 7

**Question:** the intercept relocation entry above estimates that `Has_Intercept`
"survives in roughly seven places." What is the built number?

**Resolution:** **thirteen**, and the difference is instructive rather than a
miscount of the eleven branchers.

`Adjusted_R_Squared`, `Beta_Weights`, `F_Statistic`, `F_Statistic_P_Value`,
`Group_Prediction_Interval`, `MS_Regression`, `Multiple_R`, `Prediction_Interval`,
`R_Squared`, `Regression_Degrees_Of_Freedom`, `SS_Regression`, `SS_Total`, and
`Total_Degrees_Of_Freedom`.

Two things moved in opposite directions. Three of the original eleven branchers
**stopped** needing it: `Coefficients` and `SE_Coefficients` collapse to a plain
reversal under `const = FALSE`, and `Design_Matrix` loses the argument outright
once it stops synthesizing the column — it reduces to the row filter. `AIC`,
`BIC`, and `AICc` also shed it, because `COLUMNS(X)` now counts the intercept and
`p` needs no separate term.

But the whole R²/SS chain **acquired** it, which the estimate did not anticipate:
`SS_Total` needs to know which column to project off, and everything computed from
it inherits the need. That is the real content of the correction — the estimate
counted the functions that branch on an intercept *arithmetically* and missed the
ones that need it as an *identifier*, which is exactly the role the entry above
says the flag now plays.

This does not weaken the case for the context block; it strengthens it. Thirteen
functions threading a positional flag is precisely the accretion stage two
eliminates, and all thirteen collapse into `[Context]` by the same mechanical
substitution.

### `Regression_Degrees_Of_Freedom` takes the design matrix and a flag

**Question:** the relocation entry states this function "takes the predictor matrix,
so it never counted the intercept and needs no change." Does that survive contact
with the sheet?

**Resolution:** no — it becomes `Regression_Degrees_Of_Freedom(X, [Has_Intercept])`
returning `COLUMNS(X) − N(has)`.

**Rationale:** the premise assumed the caller has a predictor matrix to hand. At
the fit sites it does not: the sheet holds `Design_Columns()`, and every internal
caller — `MS_Regression`, `F_Statistic`, `F_Statistic_P_Value` — receives the
design matrix. Keeping the old signature would mean passing `Predictor_Columns()`
alongside `X` to functions that otherwise need only one matrix, which adds an
argument to avoid adding an argument. Subtracting at each call site instead would
scatter the ANOVA convention across five formulas.

Corollary worth stating: `Residual_Degrees_Of_Freedom` moves the *other* way and
sheds its flag entirely, becoming `n − COLUMNS(X) − absorbed`. Both changes come
from the same fact — `COLUMNS(X)` is now the fitted parameter count — and the two
together are why the df arithmetic gets simpler rather than more conditional.

### `Predictors`, not `Predictor_Columns`, as the parameter name

**Question:** the escape-hatch functions take pre-intercept columns. What is that
parameter called?

**Resolution:** `Predictors`. Not `Predictor_Columns` — that is the name of the
sheet-scoped constructor closure, and a LAMBDA parameter of the same name would
shadow it inside the function body. Legal in Excel, confusing to read, and exactly
the kind of invisible distinction the constructor rename exists to eliminate.

The split is load-bearing beyond readability: the QC test-sheet harness renders its
formulas from each function's *declared argument names*, so `X` resolves to the
intercept-bearing design matrix and `Predictors` to the bare predictor block with
no per-function special case. `VIF`, `Tolerance`, `GVIF`, `Generalized_Tolerance`,
and `Correlation_Matrix` take `Predictors`; a constant column in a correlation
matrix would make it singular.

`VIF` gained one internal consequence: its auxiliary regressions genuinely need an
intercept, and `R_Squared` no longer synthesizes one, so `VIF` builds a ones column
onto each sub-matrix itself.

### `Group_Prediction_Interval` takes predictor columns, not a design matrix

**Question:** this function demeans its input internally. Which matrix should it
receive once the constructor owns the intercept?

**Resolution:** `Predictors` — the un-demeaned predictor columns, exactly what the
sheet passed before — plus `[Has_Intercept]`. It builds the intercept itself, after
demeaning.

**Rationale:** it is the one function that must *not* receive `Design_Columns()`.
That constructor has already demeaned when Fixed Effects are active, and this
function demeans again by the prediction group; handing it the design matrix would
double-demean, and the intercept column it strips would have to be re-added anyway.
Taking predictor columns and applying `demean → intercept` internally is the same
stage order [ARCHITECTURE.md § 4a](ARCHITECTURE.md#4a-the-constructor-pipeline)
fixes for the constructor, applied for the same reason: a ones column demeaned by
group is a column of zeros.

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
- **Optional-argument accretion as the standard extension mechanism**
  (v2.1 `[DF_Absorbed]`, generalized by ARCHITECTURE § 7 to "a
  LAMBDA's argument list or its internal SWITCH") → SUPERSEDED at
  v3.0 by the bounded `Model_Context` block. The v2.1 decision
  remains correct on its own terms — default 0, no-FE models
  identical, MINOR instead of MAJOR — but each addition was
  individually non-breaking, which is exactly why the sequence had
  no stopping rule. The reserved-slot pattern survives for *sheet
  columns* and for dormant `SWITCH` branches; it no longer applies
  to engine argument lists.
- **`[Allow_Intercept]` as an engine argument** (v1.0) → SUPERSEDED
  at v3.0 by the intercept column moving into the design-matrix
  constructor. `Design_Matrix` stops synthesizing the column;
  `Has_Intercept` survives in the context block as an identifier,
  not an arithmetic switch.
- **`[Weights]` as a threaded engine argument** (v2.6 planning) →
  SUPERSEDED at v3.0, **narrowly**, by WLS as a constructor concern:
  with the intercept in the constructor, √w scaling of the design
  matrix and response yields the exact WLS estimator, standard
  errors, leverage, and Cook's distance, because the intercept
  column correctly becomes √w rather than remaining ones. Only the
  *implementation mechanism* is superseded. The `Weight` value on
  the Role axis, its cardinality rule, the status-block validation,
  and the three-stage scope (user-supplied weights →
  variance-driver-derived weights → FGLS) all carry forward
  unchanged — weights remain declared in the spec block, exactly as
  recorded at v2.0 and v2.6.
- **The `X_s()` / `X_s_Within()` constructor name fork** (v2.1) →
  SUPERSEDED at v3.0 by one pipeline constructor
  (`Design_Columns()` / `Design_Response()`) plus one named escape
  hatch (`Predictor_Columns()`). The fork was statistically correct
  and invisible in the names, with nothing enforcing the
  correct-call-site rule.
- **A second spec section below the per-column block, for
  interactions** (v3.0 planning) → REJECTED at v3.0 in favor of two
  new per-column spec columns. The source table's column count is
  variable, so anything positioned below the per-column block has no
  fixed address — the spec block auto-extends as a real Excel Table,
  and a section beneath it would be displaced by any table that
  grows.
- **Row filtering moved into the constructor** (v3.0 planning) →
  REJECTED at v3.0. It would have removed `[Include]` from every
  signature and applied the mask once instead of ~30 times, but it
  breaks the `x_s()` full-height row-mask contract and the
  source-table row alignment that `Row_Labels()`, the residual-output
  zone, and the materialized design matrix all depend on. The
  performance is recovered instead by materializing
  `Sample_Include()` as its own full-height spill range.
- **`Model_Context` as the Model Comparison interface** (v3.0
  planning) → WITHDRAWN at v3.0 by the minimality decision. An
  earlier draft proposed collapsing part of v2.3 into the context
  block; the context is engine plumbing and the status block remains
  the display and cross-sheet interface. The v2.3
  `Comparison_Anchor` / `Comparison_Headline_GoF` /
  `Comparison_Prediction_Output` design stands unchanged.
- **A single version number for a single workbook** (v1.0 →
  v2.x) → SUPERSEDED at v3.0 by one library version plus a
  per-workbook version, once the build began emitting two artifacts.
  The `Breaking?` flag moves to the workbook version, where the
  question it answers actually lives.

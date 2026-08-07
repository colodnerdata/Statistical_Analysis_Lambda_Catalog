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
impose no non-blank condition; a blank category value encodes as all-zero
dummies, indistinguishable from the reference level. Tracked as an `OPEN`
item in [TODOs.md § v2.0](TODOs.md#v20-leftovers);
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
the reason `build_production.py` writes the two
datasets as the default build output, and why the spec-driven Excel verifier checks each
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
[TODOs.md § v2.1](TODOs.md#v21-leftovers--follow-on-polish).
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
Tracked in [TODOs.md § v2.1](TODOs.md#v21-leftovers--follow-on-polish).

### Categorical × FE prediction encoding — DEFERRED

**Question:** how should prediction inputs be encoded when a
non-FE Categorical Predictor coexists with Fixed Effects?

**Resolution:** DEFERRED to v2.1 polish. When non-FE categorical
predictors coexist with fixed effects, x_new and x̄ᵢ must be formed in
the *constructed* design-matrix space (dummies encoded through the
same `Dummy_Code` path `x_s()` uses), not raw input space. Largely
subsumed by v2.0 categorical prediction; recorded so the encoding
step is not forgotten. Tracked in
[TODOs.md § v2.1](TODOs.md#v21-leftovers--follow-on-polish).

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
units. Both were tracked as v3.3 and have since shipped —
[ROADMAP.md § v3.3](ROADMAP.md#v33--transforms-remainder--shipped-dispatcher--duan--model-formula-label).

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
enough — `build_production.py` / `build_univariate.py` never execute `_ROWS`;
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
artifact build's behavior at all: `write_regression_instructions_sheet`
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
table. The random number seed is a SHA-256 digest of the source CSV,
so the draw is deterministic across builds and reproducible from the
data alone. The table is sized once for the library's default
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
[TODOs.md § v3.10](TODOs.md#v310--bivariate--two-sample).

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

The v3.0 decisions respond to the 2026 architecture review, a standing
audit whose findings share one shape: each individual decision was correct,
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
workbooks carry the complete function library** — all 131 catalog functions ship
in both Name Managers. There is no bundling, no dependency closure, and no
per-artifact function subsetting; the workbooks differ only in which sheets they
contain. Splitting lets each artifact set its own calculation mode, and the
Regression workbook returns to full Automatic.

**Rationale:** the semiautomatic mode is forced by any workbook that contains
a Data Table. Even after Weibull and Gamma were demoted to static formula
grids (a hardening change shipped with v3.0, not separately logged), Beta's
two-stage grid search still uses two two-input Data Tables (20×20 each).
Two consequences would ship with a combined workbook. First, every Regression
user receives a non-default calculation mode as a side effect of a sheet they
may never open. Second, and far more serious, **Univariate fit results are
stale until the user presses Ctrl+Alt+F9** — the flagship distribution-fitting
sheet silently displays a previous answer. That is a direct violation of the
live-recalculation and visible-failure philosophy, in the one place the
philosophy exists to prevent it. Keeping the two sheets in one file means
one calculation mode has to be wrong for one of them.

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

**Status:** the Weibull and Gamma half shipped as **Univariate 2.0.0** — see
[§ Univariate 2.0.0](#univariate-200--the-grid-shrink-weibull-and-gamma-half)
for the three implementation questions this entry did not answer. Beta's
method-of-moments start and ~12×12 grid are still open, so the total is ~880
evaluations rather than the ~370 below.

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

**Shipped at Univariate 2.0.0.** Two XY-scatter-with-lines charts, each anchored
directly under its own fit zone (BP33, BZ33 — one clear row below the bodies,
one zone wide) rather than in the chart band with the histograms and Q-Q plots.
The curve and the column it is drawn from then read together; from the far chart
band a reader had to scroll between them.

**Both stages are plotted, not just Stage 1.** An earlier draft charted the wide
Stage 1 bracket alone, on the reasoning that Stage 2 is a narrow refinement a
line chart adds nothing to. That has it backwards: Stage 2 re-samples 20 points
across ±1 Stage 1 step around the winner, so it is the region the search
actually resolved and the one that fixes the reported parameter — and at Stage 1
scale it is a couple of pixels wide, which is exactly why it needs its own
series. Stage 2 carries `+` markers so it stays distinguishable where it
overlaps the Stage 1 curve. Each series reads OFFSET-based
`UV_Profile_<dist>_<S1|S2>_<Axis|NLL>` names sized by *that stage's* Grid Points
cell, so changing either stage's point count resizes only its own series.

The green→yellow→red colour scale carries over onto the profile column as well,
so the body reads at a glance without the chart. The boundary-hit flag now sits
on the searched parameter's `Best` cell only: the profiled-out partner is solved,
not searched, so it has no grid edge to land on.

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
per-column block has no fixed address. A second section beneath it would be
pushed down whenever the block grew, and every formula referencing it would need
a dynamic offset. Recorded in the supersession log. *(The reasoning has only got
stronger since: the block no longer grows by auto-extending a `SpecTable`
ListObject when the user types, it sizes itself from `COLUMNS(Source_Data)` on
every retarget, and its bands and spills now reach row 16000 — so columns B–O
below the block are reserved outright.)*

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

> **The colon is SUPERSEDED at v3.1** by one symbol per operation
> (`GDP × Schooling`). R's `:` works because R's interaction is a single
> operation; this library has three, so a shared separator names the operands
> without naming what was done to them. See
> [§ v3.1](#one-symbol-per-operation-not-a-shared-colon).

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
Construction" zone — `Model Construction` was already a sheet name
(`write_sheet_model_construction.py`, retired at v2.0) and the two had to
stay distinguishable.

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

> **SUPERSEDED** by v3.4+ *The spilled §4b zones are no longer grouped or
> collapsed*. Only `Model_Context` is grouped now; the two zones that hold
> spills are ungrouped and expanded, because a collapsed group over a spill
> range leaves the array stale and the model refits on it.

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
`write_spec_block.py` (formerly `write_sheet_model_construction.py`, renamed
2026-08-06) and calls them:

```python
from .write_spec_block import (
    _set_sheet_scoped_names as _set_spec_scoped_names,
    _set_spec_block_column_widths,
    _write_intercept_control,
    _write_spec_block,
    _write_spec_feedback,
    # … plus every _C_* column constant and formula string
)
```

`write_sheet_regression.py`'s own module docstring states the intent: *"the
spec-block writers are imported from write_spec_block so the two
sheets can never drift."* Separately, the Model Construction **sheet** is
deleted by both builds — `_delete_sheet_if_present(workbook, "Model
Construction")` in `build_production.py` and `build_univariate.py` — so only one spec
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

**Naming problem resolved 2026-08-06.** `write_sheet_model_construction.py`
was renamed to `write_spec_block.py`, and the unreachable
`write_model_construction_sheet()` / `main()` standalone-CLI path and the
`SHEET_NAME` constant were dropped — see git history for the
implementation.
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

### The model context is individual cells, not a `VSTACK` spill

**Question:** stage two materialized the context as one `VSTACK` formula spilling
into four cells. `Fit_Context` reads the fixed range those cells occupy. Is a
spill the right shape for it?

**Resolution:** no — SUPERSEDED. The block is four independent formula cells, one
per context element, each labelled in the column to its left, under a section
heading and inside a border box. `_MODEL_CONTEXT_ELEMENTS` carries the contract
name, the displayed label, and the formula for each element in one record, and is
the single source of the row order, the labels, and the height.

**Rationale:** the spill was buying nothing and costing correctness.

- **Nothing gained.** A spill exists to size output to data. This output is not
  data-sized: `ROWS(Fit_Context())` is a build-time constant, and the fixed-range
  read was already relying on that. Four cells produce the same range.
- **Correctness lost.** One formula producing four cells is a single dependency
  node. Any spec-block edit makes Excel vacate and re-spill the whole block, and
  while it is vacated the fixed range behind `Fit_Context()` holds nothing — so
  the ~30 engine call sites that read the context can observe it mid-flight. That
  is a genuine race, not a cosmetic one, and it is entirely a consequence of
  coupling four independent scalars into one node. Independent cells recalculate
  independently and are never vacated.
- **Readability gained.** Four anonymous spilled cells showed values with no
  indication of which element was which. Labelled rows show what the sheet
  computed, and the block is a fixed-size table, so it gets the heading + border
  box every other fixed-size block on the sheet gets (Regression Statistics,
  Diagnostics, Prediction Interval).

The v3.0 rationale for a *fixed range* read — no `#` inside a `LAMBDA`
`RefersTo` — is untouched and is in fact strengthened: it no longer depends on a
spill landing where it was told to.

One consequence is worth naming. With a spill, an error in any element poisoned
the block as a unit and was impossible to miss; with independent cells a broken
spec name errors in exactly one of them and leaves the other three looking
correct. So the block carries a `Context OK` row directly beneath it, inside the
box, reporting both the height invariant and that no element errored. The old
`=ROWS(Fit_Context())=4` cell was tautological against a fixed range — it checked
the range's shape, which the build had already fixed — so the error half is the
only part of that check that was ever load-bearing.

**Generalized in [ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone):**
materialize a bounded, fixed-height artifact as cells; reserve spills for the
data-dependent zones (`Sample_Include`, the design matrix), whose height genuinely
follows the source table.

Layout consequence: the zone is now two columns (labels then values), grouped as
a pair so collapsing it never strands the labels beside a hidden value column.
Every zone right of it shifts by one, and the design-matrix hard-error threshold
moves from `16,384 − (last_chart_column + 5)` to `+ 6`. Both are derived from the
layout constants, so the shift is mechanical.

---

## v3.1 — Interaction wiring

The release that consumes the M/N pair v3.0 stage 3 reserved. The
representation, the operation vocabulary and its symmetry attribute, the four
operand Role/Include cases, the two-way limit, and the self-interaction rule
were all settled at v3.0 and are unchanged — see
[§ v3.0 interactions](#interactions-are-declared-with-two-spec-columns) and the
four entries that follow it. Recorded below are the questions the
*implementation* had to answer that the representation decision did not.

### An interaction's columns follow their own spec row

**Question:** the constructor walks the spec in table order, emitting each
included Predictor's columns. Where do a row's interaction columns go — beside
that row's own block, or appended after every main effect?

**Resolution:** immediately after the declaring row's own block.

**Rationale.** The Design Columns audit is a **per-row** display, so a per-row
emission is what makes its number mean what it says: column O reads `k(row) +
k(row)×k(operand)`, and the columns those two terms describe are adjacent in the
matrix. Appending interactions at the end would be equally k-correct and would
break that correspondence — the audit would report a count for a row whose
columns are elsewhere. It also preserves the one ordering property every zone
right of the spec block already relies on: the constructed columns are in spec
order, so the predictor-summary and residual zones can be read against the spec
block row by row.

### An unusable operand degrades to the main effect; it does not error

**Question:** DECISIONS records a non-Predictor operand as an "Error". Does the
constructor refuse to build the matrix, or skip the interaction?

**Resolution:** skip the interaction, keep the main effect, leave the red flag
on the cell.

**Rationale.** This is the established precedent for an invalid spec entry on
this sheet, not a new policy: an invalid Reference Level makes `Dummy_Levels`
return `#N/A`, and the constructor's `acc` passthrough drops that variable while
the E cell shows red. Erroring the whole design matrix on a mistyped operand
name would take the entire sheet — every fit statistic, diagnostic, and chart —
down for one bad cell, and would hide the *other* red flags a user needs to see
to fix it. The audit column is where the consequence surfaces: the row reports
its main-effect count only, so "I declared an interaction and the count did not
move" is visible in the same glance that shows the red cell.

The same reasoning covers a **degenerate operand** (a Categorical whose masked
levels collapse to one). It contributes nothing, and `k(row) × 0 = 0` falls out
of the audit arithmetic with no special case.

### Interaction headers compose the library's own column names

**Question:** v3.0 specified "R's colon form — `GDP:Schooling`, level-qualified
as `GDP:StatusDeveloping`". R names dummy columns `StatusDeveloping`; this
library names them `Status: Developing`. Which wins in the composed name?

**Resolution:** compose **this library's own** constructed names.

**Rationale.** The v3.0 example is R's *output*, cited to fix the separator, not
to import R's dummy-naming. Composing the library's own names keeps the property
that matters: an interaction header always decomposes back into two headers that
appear elsewhere in the same strip, so a user reading a coefficient can find both
operands. Adopting R's dummy form would have meant `Constructed_Column_Names()`
emitting one spelling for a main effect and a different spelling for the same
column inside an interaction. The composed form is noisier than R's; that is the
cost of the names being traceable, and it is the right trade for a sheet whose
whole premise is that a result can be interrogated by clicking through it.

### One symbol per operation, not a shared colon

**Question:** v3.0 fixed the separator as a colon, from R. Does that survive
three operations?

**Resolution:** no. Each operation renders its own operator — ` × ` for Product,
` − ` for Difference, ` ÷ ` for Ratio — with ` ? ` for anything else.
`_INTERACTION_HEADER_SYMBOLS` in `write_spec_block.py` is the
single source, and `test_interaction_header_symbols_match_the_catalog_formula`
pins it to the `SWITCH` inside `Constructed_Column_Names()`.

**Rationale.** R's `:` is unambiguous *in R*, where interaction is one
operation. Here it is not: `Weight:Displacement` could be a product, a
difference, or a ratio, and those are three different models. A header that
names the operands but not what was done to them fails the same test the
`Ln(name)` relabel passes — the output has to say what was fitted.

The colon was also **doubly** ambiguous, which is what made this worth fixing
rather than tolerating. A level-qualified categorical name already contains
`": "`, so `Weight:Status: Developing` reads as one name with two colons and no
indication which is the join.

`−` is U+2212 MINUS SIGN, not a hyphen: a hyphen is a legal character in a
source column name, so `Unit-Cost - Weight` would be unparseable by eye. The
symbols are spaced because operand names contain spaces.

**Why ` ? ` rather than nothing for an unrecognized operation.** The header
strip must stay exactly as wide as the design matrix — the twin invariant — so
an unrecognized operation still needs *a* header. Pairing a visibly wrong header
with the `NA()` column `Predictor_Columns()` emits for the same input makes the
failure legible from the strip alone.

### The Python mirror matches Excel's comparison semantics, not Python's

**Question:** `mate()` resolves the operand with `XMATCH` and the operation
dispatches on `SWITCH`. The Python oracle used `==` for both. Is that the same
thing?

**Resolution:** no — both Excel functions compare text **case-insensitively**,
and the mirror now case-folds to match. Case folding only: `XMATCH` is neither
accent-insensitive nor whitespace-trimming, so the mirror is neither.

**Rationale.** The failure this prevents is worse than either behaviour alone:
with a case-sensitive mirror, a user who pastes `weight` where the header reads
`Weight` gets an interaction column on the sheet and an oracle that predicts
none — so the QC pass reports a mismatch on a *correctly built* matrix, and the
oracle is wrong about the thing it exists to describe. **A mirror's job is to
reproduce Excel's semantics, including the ones Python does not share.** Any
future mirror of a formula that compares text has the same obligation;
`_retained_levels` (mirroring `Dummy_Levels`) and `_compute_mask` are the other
two places this rule bears on.

Found in review of the v3.1 wiring, not by a test — which is why it is recorded
as a rule rather than a fix.

### `Ratio` is an explicit `SWITCH` case, so the default can be `NA()`

**Question:** `SWITCH`'s trailing argument is its *default*. Writing the three
operations as `SWITCH(o, "Product", …, "Difference", …, <ratio>)` is one
character shorter than naming `Ratio`. Does it matter?

**Resolution:** name `Ratio` explicitly and make the default `NA()`.

**Rationale.** With `Ratio` as the fallthrough, *every* unrecognized value —
not just the three on the dropdown — silently computes a ratio. That is
reachable: Excel's data validation does not block a paste, so a value the
dropdown would refuse can still land in N. Silently computing a ratio for a
value the user did not choose is precisely the "silently switch" failure the
closed-vocabulary decision exists to prevent, and it would be invisible — the
column builds, the fit succeeds, and the number is for a model nobody
specified. `NA()` fails visibly instead, and the Python mirror raises for the
same input, so neither side guesses.

### Prediction Inputs does not recompute interaction rows

**Question:** the Prediction Inputs band writes one overridable value per
constructed column, defaulting to that column's training mean. An interaction
column is *derived* from two others. Should the band recompute it when an
operand row changes?

**Resolution:** no. The interaction row is an independent input like every other
row, and the band's header note says so explicitly.

**Rationale.** Recomputing would mean one user input silently rewriting another,
which is the "flag and instruct, never silently switch" line this library holds
everywhere else. It is also not obviously *correct* — a user exploring a
scenario may legitimately want to hold an interaction at its training mean while
moving an operand. What the band must not do is leave the inconsistency
undiscoverable, hence the note.

The default state is self-consistent without any of this: leave every row at its
Training Mean and the prediction sits on the design matrix's own centroid,
interaction columns included. The inconsistency only arises from a partial
override, which is exactly what the note describes.

**Deferred, not rejected:** a band that knows which constructed columns are
derived and from which operands could offer a derive-on-change toggle. That
needs a fourth structural twin carrying each column's provenance, which is real
scope and no part of what v3.1 set out to do.

### Workbook scope belongs to the catalog

**Question:** the v3.0 split left each artifact carrying the *other* one's
named ranges at workbook scope — twelve `RegChart*` entries reading
`OFFSET(#REF!,…)` in the Univariate workbook, forty-two `UV_*` entries in the
Regression workbook, alongside twenty-one LAMBDA names the catalog retired
releases ago. `sync_workbook_names` already stripped workbook-scoped residue,
but only when the body *was* an error literal, and none of these were: they
wrap the `#REF!` inside an `OFFSET(...)`. What is the rule that catches all of
them?

**Resolution:** the catalog owns workbook scope outright. After
`sync_workbook_names` runs, the only workbook-scoped `<definedName>` entries
are the catalog's workbook functions plus Excel's reserved `_xlnm.*` names;
every other workbook-scoped entry is dropped without inspecting its body.
Sheet-scoped entries are never touched.

**Rationale.** Every range the sheet writers create is created through
`sheet.api.Names.Add` and is therefore sheet-scoped — that was already true of
`RegChart*` and `UV_*` when the residue was found, which is why the *live*
names in both artifacts were correct and only the stale copies were wrong.
That makes "workbook-scoped and not in the catalog" an exact characterization
of residue, where any body-shaped test is a guess about what the last broken
build happened to leave behind. The rule also subsumes the two narrower ones it
replaces (broken-body, duplicate-of-sheet-scoped) and needs no maintenance as
name families come and go.

**Enforcement:** `TestRealWorkbookNameScope` in
`tests/test_workbook_invariants.py` asserts it against both committed
artifacts on every commit — pure zipfile, no Excel, so it runs in CI.

### The Univariate artifact does not carry `Base_Period_Delta`

**Question:** `Base_Period_Delta()` is workbook-scoped (workbook-scoped
callers like `Difference_By` and `BFN_Panel_Durbin_Watson` fall back to it when
their `[delta]` argument is omitted, and a workbook-scoped name cannot resolve
a sheet-scoped one unqualified), but its body reads the Regression sheet:
`COLUMNS('Regression'!Source_Data)`, `'Regression'!Spec_Sequence`,
`'Regression'!Spec_Sequence_Period`. The Univariate artifact has no Regression
sheet. What happens to it there?

**Resolution:** it is not written into that artifact. `sync_workbook_names`
skips any definition naming a worksheet the target workbook does not have, and
reports the skip in the build summary. The v3.0 "both workbooks carry the
complete function library" rule stands with this one narrow exception, which
the sync derives rather than a hand-maintained per-artifact list.

**Rationale.** Excel does not leave a reference to a missing sheet unresolved.
It rebinds it to an external workbook — `Regression!Source_Data` becomes
`[1]!Source_Data` — writes an `xlExternalLinkPath/xlPathMissing` external link
part, and prompts about broken links on every open of the shipped file. That is
what the Univariate artifact was doing. The alternatives are worse: making the
name sheet-scoped breaks its workbook-scoped callers in the *Regression*
artifact, where it works correctly today, and shipping a deliberately broken
name to keep an inventory count intact trades a real defect for a cosmetic one.
`Base_Period_Delta()` reads the Regression spec block; in a workbook with no
spec block there is nothing for it to mean.

**Consequence:** in the Univariate workbook, `Difference_By` /
`BFN_Panel_Durbin_Watson` / `Lag_By` called with `[delta]` omitted return
`#NAME?` rather than a delta. No Univariate cell calls them, and the panel
diagnostics they serve are Regression-sheet features.

---

## Univariate 2.0.0 — the grid shrink, Weibull and Gamma half

The design settled at v3.0 — see
[§ the grid shrink ships as a later release](#the-grid-shrink-ships-as-a-later-release-of-the-univariate-artifact)
and [§ profile-NLL line charts](#profile-nll-line-charts-replace-the-weibull-and-gamma-heatmaps).
Weibull and Gamma now profile their scale / rate parameter out in closed form
and search a 20-point profile-NLL column per stage, replacing four 20×20 grid
blocks (`_write_profile_stage` in `write_sheet_univariate.py`). **Beta is
unchanged** — it keeps its two two-input Data Tables, and its method-of-moments
start and ~12×12 grid are still open. Recorded below are the three questions the
implementation had to answer that the design decision did not.

### The profiled-out parameter is substituted into the catalog NLL, not into an analytic profile log-likelihood

**Question:** the v3.0 entry wrote out ℓ_p(k) and ℓ_p(α) in closed form, with
their fixed sample statistics (x̄, Σln xᵢ) computed once. Should each body cell
evaluate that expression?

**Resolution:** no. Each body cell calls the same catalog NLL LAMBDA the
fitting table reports — `NLL_Weibull(UV_Data, k, λ̂(k), UV_Include)` — with only
the *partner* parameter substituted in closed form. The two forms are
algebraically identical; this one is what makes the curve's minimum and the
fitting table's NLL **the same number by construction** rather than by two
derivations agreeing. The analytic form's advantage was hoisting the fixed
statistics out of the inner loop, and at 20 evaluations per stage instead of
400 that is not a cost worth buying correctness risk with.

Consequence: the profiled-out parameter appears exactly twice per stage, both
generated from one helper (`_weibull_profile_scale`, `_gamma_profile_rate`) —
in the body formula and in the `Best` cell that reports it. There is no third
place for the two to drift apart.

**The sample is bound once per body cell**, as `LET(x, FILTER(UV_Data,
UV_Include), …)`, and passed to both the partner helper and the NLL call. Two
full-range `FILTER`s per cell is what the naive form costs — the partner
re-filters the 2,000-row input range and the catalog NLL LAMBDA filters its
`data` argument again internally — and at ~2,900 included rows that is
comparable to the likelihood evaluation the cell exists to perform, so it would
have given back a real fraction of the shrink. Both helpers therefore take the
sample as an expression (defaulting to `FILTER(UV_Data,UV_Include)` for the
standalone `Best` cells, which evaluate once and need no binding), and the NLL
call omits its optional `[filter]`: `UV_Include` *is* `ISNUMBER(...)`, so `x` is
already the included numeric sample and the LAMBDA's `ISNUMBER(x)` default
filters nothing. The body carries an explicit `IFERROR(…, 1E+15)` because `LET`
propagates an error in `x` past the sentinel inside the NLL LAMBDAs, which is
where the 2-D bodies get theirs.

### The Weibull start builds its own Hazen positions, not `Rank_Fraction`

**Question:** the v3.0 entry specified the Weibull starting value as the
probability-plot regression of ln(−ln(1−F̂)) on ln x, "with F̂ from the existing
`Rank_Fraction()`."

**Resolution:** the formula instead builds the plotting positions the way Zone
6's `P` column already does — `SORT` once, then `(SEQUENCE(n)-0.5)/n` — for two
independent reasons.

*Correctness:* `Rank_Fraction` returns i/n, which is exactly 1 at the sample
maximum, so ln(−ln(1−F̂)) is an error there. The raw form would have thrown on
**every** sample, not an unlucky one, falling back to the constant k₀ = 2 and
quietly disabling the start it was supposed to provide. The Hazen adjustment
(i−0.5)/n is what makes the expression evaluable at all, and it is the
convention Zone 6 and the Regression sheet's Normal Q-Q machinery already use,
so the start is consistent with the plots a user compares it against.

*Cost:* `Rank_Fraction` is `BYROW` over `SUMPRODUCT` — O(n²), or ~8.6M
comparisons on the shipped 2,928-row sample, in a cell that recalculates on
every data change. The sorted form is O(n log n) and is already being computed
one zone over.

The whole expression still carries an `IFERROR` fallback (k₀ = 2) for a sample
that cannot support the regression at all.

### Stage 1's bounds are formulas in an input-coloured cell

**Question:** the searched parameter's Min/Max were user-typed numeric
defaults. With a closed-form start available, should they become derived?

**Resolution:** both. Stage 1's Min and Max hold formulas bracketing the start
(`start/3`, `start*3`) and **stay input-coloured**: the shipped default needs no
user judgement about a plausible range, and a user whose sample trips the
boundary guard can still type a literal over either cell. This is the one place
in the search a user has a decision to make, so removing the cell entirely — the
other reading of "the bounds become derived" — would have left a red boundary
flag with no remedy.

The bracket is wide relative to the starts' accuracy. Across simulated Weibull
shapes 0.4–30 and Gamma shapes 0.3–200 at n = 60 and n = 2000, the two stages
never landed on a boundary and never finished more than 0.09 NLL above the
scipy MLE. On the shipped Life Expectancy sample both fits came out *better*
than the grids they replaced: Weibull 10608.12 vs. 10608.22, and Gamma 10854.86
vs. **10861.68** — the old 2-D Gamma bracket ([0.5, 100] × [0.001, 2] in 20
steps) was too coarse to resolve α ≈ 48.5, and had been shipping a fit 6.8 NLL
units off the optimum, worth ~13.7 of AIC. The shrink is not a
speed-for-accuracy trade here; it is both.

### The 1-D stage cannot use `Grid_Search_Optimum`

**Question:** the 2-D stages read their result with
`Grid_Search_Optimum(body)`. Does it carry over?

**Resolution:** no, and this is a silent-wrong-answer trap rather than an error.
`Grid_Search_Optimum` returns `VSTACK(OFFSET(grid,-1,col−1,1,1),
OFFSET(grid,row−1,-1,1,1))` — the cell above the body and the cell to its left.
On a single-column body the first of those is the `Profile NLL` **header text**,
not a parameter value, and the call returns a two-element stack whose first
element is a label. The 1-D stage instead reads
`INDEX(<axis name>, INDEX(Grid_Argument_Minimum(<body>),1,2))`;
`Grid_Argument_Minimum` is dimension-agnostic and needs no change.

`Grid_Search_Optimum`'s catalog entry is unchanged and correct — it documents
itself as reading a two-input Data Table body. Nothing new is added to the
catalog: the shrink is entirely a sheet-writer change, so the function-library
version does not move.

### The fit band is distribution-major, and the fit zones go last

**Question:** the search band was **stage-major** — Stage 1 in one 21-column
band, Stage 2 in a second, with the three distributions stacked vertically
inside them. Does that survive the shrink?

**Resolution:** no. The band is now **distribution-major**: one column zone per
fit, sized to what that fit needs, with its two stages stacked inside it. The
band runs Q-Q data → Weibull (9 cols) → Gamma (9) → Beta (21).

**Rationale.** Stage-major was right when all three fits were 20×20 grids of
identical width. The shrink broke both of its premises at once. The two profile
fits now need 9 of each band's 21 columns, so twelve columns per band per fit
sat empty; and a single distribution's two stages ended up 22 columns apart, so
you could not see a fit's Stage 1 and Stage 2 without scrolling past an
unrelated distribution. Distribution-major restores the thing a reader actually
wants adjacent — one fit's two stages — and lets each zone be the width it
needs. The band is 26 rows shorter and the sheet ends at DD instead of DF.

Within a profile zone the two stages' **control blocks stack vertically** while
their **bodies sit side by side** on shared rows, each under one half of the
control block. That is not decoration: the two profile-NLL curves are the wide
search and its refinement of the same parameter, so reading them against each
other is the point, and sharing the body rows is what makes that possible.

**The fit zones come last, and must stay last.** They are the only zones in the
band whose width is a tunable — `_N_GRID` and `_N_PROFILE` set it, and Beta's
still-open shrink to a ~12×12 grid will change its width by nine columns. With
them at the end, a resize displaces nothing; with the Q-Q data after them, every
Q-Q column and named range would move on an ordinary grid-size change. This is
the same rule as the Regression sheet's "nothing may ever be placed to the right
of the design-matrix zone," for the same reason.

**Consequence — `_final_grid_best_refs` is no longer one formula.** Under
stage-major all three fits shared a Stage 2 Best column, so the fitting table
read them with a single helper parameterised only by row. With one zone per fit,
each distribution has its own Best column *and* its own Stage 2 row (5 rows down
for a profile fit, 26 for Beta). The helper now reads `_STAGE2_ANCHORS`, a dict
the profile specs, the Beta call, and `_dist_rows` all share — so the fitting
table cannot reference a block the search writers did not produce. Giving up the
shared column was the accepted cost of per-fit sizing; routing both through one
dict is what keeps it from becoming a drift risk.

Zone geometry derives from one ordered `_BAND_ZONES` table
(`_derive_band_columns()`), so reordering the band means reordering that list
and nothing else. No column letter in the band is hard-coded.

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

**No alias table is kept here.** One was — 46 proposed short names against
their canonical functions — and it was deleted rather than maintained: every
row is a second name to rename in step with the first, for a layer nobody has
committed to building. The naming convention above is enough to regenerate the
whole set mechanically whenever the canonical library does stabilize, and
regenerating it then is cheaper than keeping it correct until then.

---

## Supersession log

Decisions that were made and later replaced by a later decision. The
superseding decision lives at its version's section above; this log
just records what was replaced, when, and by what.

- **The colon as the interaction-header separator** (v3.0) → SUPERSEDED at
  v3.1 by one symbol per operation (` × ` / ` − ` / ` ÷ `). A shared separator
  named the operands without naming the operation, and collided with the `": "`
  already inside a level-qualified categorical name.
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
- **Weibull and Gamma NLL as static 20×20 formula grids** (v3.0) →
  SUPERSEDED at Univariate 2.0.0 by the 1-D profile search. The formula
  grid was a holding position — it removed the Data Table object while
  the search stayed two-dimensional, so it still paid 400 evaluations
  per stage for a surface one of whose axes has a closed-form solution.
  Profiling the scale/rate parameter out drops each stage to 20
  evaluations and, on the shipped sample, lands closer to the MLE than
  the grid it replaces.
- **`Model_Context` materialized as a single `VSTACK` spill** (v3.0
  stage two) → SUPERSEDED by the decomposed block: one labelled cell
  per context element, headed and border-boxed. A spill sizes output
  to data, and this output is a build-time constant, so it gained
  nothing; what it cost was a single dependency node that Excel
  vacates and re-spills on every spec-block edit, leaving the fixed
  range behind `Fit_Context()` transiently blank for all ~30 engine
  call sites. The fixed-range read itself (no `#` inside a `LAMBDA`
  `RefersTo`) is unchanged.
- **The Model Formula cell at `AA2:AB2`** (v3.3) → SUPERSEDED by the
  labelled readout on row 1 of the §4b band's terminal Constructed Design
  Matrix zone, holding `=Model_Formula()`. Row 2 of the Regression Outputs zone wraps and is
  then AutoFitted, so the sheet's longest string set the height of the
  whole header row; and an inline 300-character concatenation in one cell
  documented nothing on the LAMBDA_functions sheet. `Comparison_Model_Formula`
  is why the move cost its v3.4 consumer nothing — that surface is a NAME.
- **Content-column widths keyed on literal column letters** (pre-v3.0) →
  SUPERSEDED by `_COLUMN_WIDTHS`, keyed on the `_C_*` layout constants with
  a coverage assertion. The letter keys survived the v3.0 layout break
  unchanged and had been sizing the wrong columns ever since.

---

## v3.3 — Transforms remainder (unit-space dispatch + Duan back-transformation + model formula label)

Closes the unit-space gap that v2.2 left open: the v2.2 column-G `Log` wiring
delivered a model that **fits** correctly in log space end to end, but stops
there. The `Y`, `Predicted Y`, residual columns, and the whole Prediction
Outputs block stay in log space, labelled `(Log)`; an R² computed on `Ln(y)`
is not comparable with one computed on raw `y`. v3.3 closes that gap and,
alongside it, ships the model-formula label — needed to head the unit-space
block and read again by v3.4 Model Comparison.

Six amendments to the v2.2/v2.3 design record:

1. **The transform pair comes from `[Context]`, not from two positional
   arguments.** The v2.2 resolution wrote `Unit_Space_R_Squared(model,
   response_transform, predictor_transform)`; that predates the v3.0 context
   array, which already carries both. Passing them again would be a second
   source of truth. Signature becomes `Unit_Space_R_Squared(X, Y, Y_Full,
   [Include], [Context], [Method])`, dispatching on
   `Context_Response_Transform` / `Context_Predictor_Transform`. These are
   the reader elements 3–4 of `Fit_Context()` were reserved for.

2. **`Y_Full` is `Response_Column()`** — transformed but *not* demeaned —
   passed alongside `Y` = `Design_Response()` (transformed *and* demeaned).
   Their difference on the filtered sample is exactly the level the within
   transformation removed, so the full-space fitted value is
   `Predictions(X,Y,Include) + shift`. This is what makes Fixed
   Effects + Log work without a `Group_Mean` call, an FE-detection branch
   or a new closure — the level shift is zero whenever no FE row is
   declared.

   **The shift is gated on a Log response:** `shift = IF(rt="Log",
   FILTER(Y_Full,Include) − FILTER(Y,Include), 0)`. It exists only so that
   `EXP()` under FE exponentiates a predicted log response rather than a
   group deviation, and nothing is exponentiated under `None`. Applying it
   unconditionally silently converts the within-flavoured statistics into
   total ones, so `Unit_Space_R_Squared` stops equalling `R_Squared` on an
   FE model with no transform — the reduction invariant below, broken by a
   number that still looks plausible.

2a. **`Y_Full` is NOT the observed response in original units, and the
   distinction is load-bearing.** It is fit space: `ln(y)` under a Log
   Response row. Every unit-space statistic subtracts an observed value from
   a back-transformed prediction, so reading the observed side straight off
   `Y_Full` puts `ln(y)` and a response-unit prediction in one subtraction.
   On a synthetic log-linear fit that returns **R² = −148.66 where the true
   original-scale R² is 0.693** — a wrong number, not an error, in the cell
   `Comparison_Headline_GoF` publishes to v3.4.

   Resolved with a dedicated accessor, **`Unit_Space_Observed(Y, Y_Full,
   [Include], [Context])`**, so the choice is made once rather than restated
   at each of the three goodness-of-fit call sites. Both cases fall out of one
   expression: `Back_Transform_Response(Dependent_Variable(Y, Include) +
   shift, Context, "Naive", 1)`, reusing the same gated `shift`. Under `Log`,
   `y_level` is `Y_Full` and the back-transform returns raw `y`; under `None`
   the shift vanishes, the back-transform is a pass-through, and the observed
   side stays the within-demeaned column the ordinary statistics use — which
   is what keeps the reduction invariant true under Fixed Effects.

   **The observed side is always `Naive`, never smeared.** Duan's factor
   lifts a *prediction* from the conditional median to the conditional mean.
   An observation is neither, and smearing one corrupts `SSE` and `SST`
   alike.

3. **The pair-SWITCH lives in `Unit_Space_Predictions` /
   `Back_Transform_Response`, not repeated in each GoF name.** The
   transform pair changes the *arithmetic* only at the back-transformation
   step; R², Adjusted R² and RMSE are ordinary per-statistic names over
   the back-transformed fitted values. Confines the combinatorial dispatch
   to one place and keeps the `Unit_Space_*` names the ARCHITECTURE § 1
   departure records.

4. **The recognised pairs are `{None, Log} × {None, Log, Mixed}` — six, not
   four.** The v2.2 text says "four possible combinations", written before
   `_PREDICTOR_TRANSFORM_FORMULA` gained its `Mixed` value. Without a
   `Mixed` branch a spec with one logged and one unlogged predictor falls
   through to `NA()`. Anything outside the six is `NA()`. The predictor
   half of the pair is **inert** for these three statistics — a
   response-unit statistic cannot depend on predictor units — so it is
   carried as *validation*, not as a computation input.

5. **CI/PI bounds are back-transformed with `EXP` alone, never smeared.** A
   monotone transform preserves coverage:
   `P(L ≤ ln y ≤ U) = P(e^L ≤ y ≤ e^U)`. Multiplying both bounds by the
   smearing factor would destroy it. Consequence the caveat row must state:
   under `Duan` the point estimate (conditional **mean**) does not sit at
   the centre of the interval, whose ends bracket the conditional
   **median**.

6. **Under FE + Log the unit-space statistics are *total*, not *within*.**
   `exp()` of a within deviation predicts nothing, so the back-transformed
   fitted value necessarily carries the group effect and `SST_unit` is
   taken about the grand mean of raw `y`. Every other statistic on the
   sheet reports the within flavour; this one cannot. State it on the
   block's cell note. With `Transform = None` the dispatchers return the
   ordinary statistic verbatim, so the within convention is untouched
   there.

### Catalog

Eight new functions in `lambda_functions.json` under subcategory
`Back-Transformation` (a new subcategory of `Model Construction`):

- `Smearing_Factor(X, Y, [Include], [Context])` — scalar. `1` when the
  response transform is `None`; `AVERAGE(EXP(Residuals(X,Y,Include)))` when
  `Log`; `NA()` otherwise. Returning `1` rather than `NA()` on `None` is
  what lets the sheet's cells stay uniform.
- `Back_Transform_Response(Values, [Context], [Method], [Smearing])` —
  elementwise. `None` → `Values` unchanged. `Log` + `Duan` → `EXP(Values) *
  Smearing`. `Log` + `Naive` → `EXP(Values)`. Unrecognised transform or
  method → `NA()`. `[Method]` defaults `"Duan"`, `[Smearing]` defaults `1`.
- `Unit_Space_Predictions(X, Y, Y_Full, [Include], [Context], [Method])` —
  n×1. `SWITCH` on the six recognised `(response, predictor)` pairs; each
  branch back-transforms `Predictions(X,Y,Include) + shift`, with
  `shift = IF(rt="Log", FILTER(Y_Full,Include) − FILTER(Y,Include), 0)`.
  Computes its own smearing factor.
- `Unit_Space_Observed(Y, Y_Full, [Include], [Context])` — n×1. The observed
  response read in the **same space** `Unit_Space_Predictions` returns:
  `Back_Transform_Response(Dependent_Variable(Y,Include) + shift, Context,
  "Naive", 1)`. Raw `y` under `Log`; the within-demeaned fit-space column
  under `None`. The single reason the three GoF names cannot disagree about
  which space their observed side is in.
- `Unit_Space_Residuals(X, Y, Y_Full, [Include], [Context], [Method])` — n×1.
  `Unit_Space_Observed` minus `Unit_Space_Predictions`.
- `Unit_Space_R_Squared(X, Y, Y_Full, [Include], [Context], [Method])` —
  `1 − SSE_unit/SST_unit`; `SST_unit` about the mean with an intercept
  (`Context_Has_Intercept`), about zero without — mirroring `SS_Total`.
  May go negative; that is honest, do not clamp.
- `Unit_Space_Adjusted_R_Squared(...)` — `1 − (1 − R²_unit) ·
  Total_Degrees_Of_Freedom / Residual_Degrees_Of_Freedom`, reusing the
  existing df functions so `Context_DF_Absorbed` is honoured.
- `Unit_Space_RMSE(...)` — `SQRT(SSE_unit / Residual_Degrees_Of_Freedom(...))`
  — the same divisor `SE_Regression` uses, so the `None` case reduces to it
  exactly.

**Reduction invariant (acceptance criterion):** with `Transform = None`
everywhere, `Unit_Space_R_Squared ≡ R_Squared`, `Unit_Space_Adjusted_R_Squared
≡ Adjusted_R_Squared`, `Unit_Space_RMSE ≡ SE_Regression`, and the two new
residual columns equal `Predicted Y` / `Residuals`. Same non-breaking
property v2.2 established.

**It has to hold WITH Fixed Effects too**, and that is the version worth
testing: with no FE the level shift is zero and every branch is trivially
inert, so a no-FE-only invariant test passes against code that is wrong.
`production_lots_fixed_effects` is the real case — FE declared, no transform —
and both the gated shift (2) and `Unit_Space_Observed`'s `None` branch (2a)
exist to keep it within-flavoured.

**A mirror test cannot verify this family on its own.** The pure-Python
mirrors and the catalog LAMBDAs share the same author and the same reading of
what `Y_Full` means, so a wrong shared assumption produces a green suite and a
wrong workbook — which is exactly what happened. Two kinds of test are
therefore load-bearing here and must be kept: one that computes the expected
`R²_unit` **straight from the raw response** without touching a mirror, and
one that recomputes the QC oracle's unit-space block from its own already-
verified residual columns. Neither shares a derivation path with the code it
checks.

### Sheet additions (`lambda_catalog/write_sheet_regression.py`)

- **Unit-space block at `AG3:AH9`**: section heading on row 3; Back-Transform
  Method input on row 4 (default `"Duan"`, list validation against
  `Duan,Naive`); Smearing Factor, R² (Unit), Adj R² (Unit), RMSE (Unit) on
  rows 5–8; Response Space readout on row 9. `border_box(3, AG, 9, AH)`.
- **Original Units column in `AL` (Prediction Outputs)**: AK2 sub-header
  `"Fit Space"`, AL2 sub-header `"Original Units"`. `AL3` point estimate via
  `Back_Transform_Response(AK3, Fit_Context(), $AH$4, Smearing_Factor(...))`.
  `AL4:AL6` blank (no SE/t-critical counterpart). `AL7:AL10` CI/PI bounds via
  `Back_Transform_Response(AK{row}, Fit_Context(), "Naive", 1)`. Caveat row
  at `AJ15:AL15` (merged, wrapped) explaining the asymmetric placement.
- **Model Formula cell at `AA2:AB2`** *(SUPERSEDED at v3.3.x — the readout
  moved to row 1 of the §4b band's design-matrix zone and the assembly
  became the sheet-scoped `Model_Formula()` closure; see § Regression sheet
  layout repair)*: AA2 bold label, AB2 the assembled
  string. Built from `_RESPONSE_NAME_FORMULA` (which already emits
  `Ln(name)` when Log), `Allow_Intercept`, `Constructed_Column_Names()`,
  and the FE-name suffix gated by the Fixed Effects count. The mixed
  Log/None predictor case renders correctly with no extra work because
  `Constructed_Column_Names()` already emits `Ln(name)` per logged
  predictor, level-qualified dummy names, and `left × right` interaction
  names.
- **Residual Output zone extended to `AN:BA`**: AZ (`Predicted Y (Original
  Units)`) and BA (`Residual (Original Units)`) added as content columns.
  Headers carry NO `(Log)` / `(Within ...)` suffix — they are in original
  units by construction. Row 3 formulas call `Unit_Space_Predictions` /
  `Unit_Space_Residuals` with the Method toggle from `AH4`. Chart anchor
  shifts from `_C_AZ` to `_C_BB` (so the chart letter BP — the end of the
  chart-anchor constant `_LAST_CHART_COLUMN = _C_BB + 14` — is preserved).
  The `_C_AZ`/`_C_BA`/`_C_BB` constants and the chart-label columns
  (`_C_CHART_LABEL_NAME` through `_C_CHART_YLABEL`) all derive from the
  anchor, not from literal letters.
- **Sheet-scoped `Comparison_*` named ranges** registered in
  `_setup_local_names`: `Comparison_Anchor` → `$AF$2` (response-name
  readout); `Comparison_Headline_GoF` → `$AH$6:$AH$8` (the three
  unit-space GoF statistics); `Comparison_Model_Formula` → `$AB$2` (the
  assembled model formula string; retargeted at v3.3.x when the readout moved,
  which is the point of naming the surface). All three are the v3.4 Model Comparison
  sheet's reading surface — the public-interface commitment this milestone
  ships.

### QC oracle (`lambda_catalog/regression_shared.py` + `analyze_regression_spec.py`)

`RegressionUnitSpace` dataclass holds the unit-space scalars and vectors
mirrored against the workbook. The oracle computes the smearing factor, R²,
Adjusted R², RMSE, predictions, residuals, and the model formula string. The
Regression spec case list grows by three — `production_lots_log_no_fe`,
`production_lots_log_mixed_predictors`, `production_lots_log_predictor_only`
— covering the FE+Log, Log+Mixed, and (None, Log) pairs respectively. The
cache schema version bumps to 17.

### Tests

- `tests/test_unit_space_dispatch.py` (new) — pure-Python mirror of each
  catalog function cross-checked against a NumPy OLS reference; reduction
  invariant and `Y_Full` level-shift assertions; catalog implementation-shape
  checks (SWITCH on the six pairs, `EXP(Values) * smear_arg` under Duan, no
  `(Log)` / `(Within)` leak in the new residual-output headers).
- `tests/test_sheet_writers.py` (extended) — RecordingSheet tests pin the
  unit-space block (`AG3:AH9`), the `AZ`/`BA` residual columns, the Model
  Formula cell, and the `Comparison_*` named ranges.
- `tests/test_workbook_invariants.py` — unchanged; the new workbook-scoped
  catalog names and sheet-scoped names are picked up by the existing name-
  scope checks.
- `tests/test_catalog_schema.py` and `tests/test_lambda_catalog_plain_language.py`
  — pick up the 7 new entries automatically.
---

## v3.3.x — Regression sheet layout repair

Three defects found by reading the built Regression sheet against its own
layout constants, all with the same shape: something that had to move when the
zones moved did not.

### The content-column widths were still keyed on pre-layout-break letters

**Question:** the Predictor Summary, Regression Outputs, Prediction Outputs and
Residual Output zones render at visibly wrong widths — a name column too narrow
to show a level-qualified name, a stats column too wide, a whole zone at Excel's
default. Where does the width come from?

**RESOLVED** — from a dict of literal column LETTERS in
`write_regression_output_sheet`. The v3.0 layout break (three columns appended to
the spec block, every zone right of it shifted three columns over) moved the
zones and left that dict untouched, so from v3.0 on the widths were applied
three columns to the left of where they were meant: `S` (constructed column
names, wanting 24) took a stats column's 9, `X`/`Y` (GVIF/Tolerance) took the
22/12 meant for the Regression Outputs labels, `AD` (diagnostics labels, longest
`"BFN Panel Durbin-Watson"`) took 10, and `AJ`–`AL` (the entire Prediction
Outputs zone) matched nothing in the dict at all and rendered at the default
8.43. Nothing failed, because a wrong column letter is still a valid column —
the same silent-wrong-answer mode ARCHITECTURE's "never spell an A1 address into
a formula string" rule exists to prevent, in the one table that had been left
outside it.

The fix is the rule the rest of the sheet already follows: `_COLUMN_WIDTHS` is a
tuple of `(column constant, width)` pairs, with module-level assertions that
every zone content column is sized exactly once, that no gap column is sized (the
`_GAP_COLUMNS` loop owns those), and that nothing outside the zones is sized
except the two deliberate cases — column `I`, the Regression-only Verdict
overlay on a spec-block column, and `BB`, the post-zone chart gutter. The next
shift fails at import.

**REJECTED — re-derive the letters and keep the dict.** It is a smaller diff and
would have been correct on the day. It also leaves the next layout change with
the same trap, and this trap had already been shipped once without anyone
noticing, which is the argument against a fix that cannot fail loudly.

### The Model Formula readout moved out of the Regression Outputs header

**Question:** the assembled `<response> ~ 1 + <predictors> [| <FE>]` string
shipped at `AA2:AB2` (v3.3). Row 2 of that zone has `WrapText` set across
`S2:BA2` and is then `AutoFit`-ed, so the longest string on the sheet — in a
12-wide column — dictates the height of the sheet's entire header row, pushing
every zone's data down the screen. Where should a caption live?

**RESOLVED** — on **row 1 of the ARCHITECTURE §4b band's terminal Constructed
Design Matrix zone**, right of that zone's own heading: header two columns over,
readout three columns past the header, both derived from the zone anchor, with
`WrapText` explicitly FALSE. It is a caption, not a headline statistic, and the
band past the charts is where the sheet's other read-only, machine-consumed
surfaces already live (`Fit_Context`, the `Sample_Include` mask, the design
matrix itself).

Row 1 of that zone is the specific choice, and it is chosen for what the zone
cannot do to it: the design matrix's names spill on `_MATERIALIZATION_HEADER_ROW`
and its values on `_MATERIALIZATION_SPILL_ROW`, and both grow *rightward* from
the anchor — never up — so row 1 stays empty no matter how wide the matrix gets.
That is also why this does not breach the zone-ordering rule. Nothing is placed
to the RIGHT of the terminal zone; the caption is placed ABOVE its body, inside
the zone's own columns, where an ordinary modelling choice cannot displace it.

What that buys is the display the earlier placement could not give. Under the
Model Context block the value overflowed only as far as the next gutter, so a
long formula read in full in the formula bar and nowhere else; on an empty row 1
it overflows across as many columns as the string needs. The three-column gap
between header and readout is load-bearing for the same reason in miniature —
`"Model Formula"` is wider than one 12-wide design-matrix column, so a readout
immediately beside the header would clip it.

**Accepted cost:** the design-matrix zone ships COLLAPSED (an unbounded-width
zone that cannot be collapsed is a scrolling hazard), and a collapsed outline
hides its columns including row 1 — so the caption, like the zone's own heading,
is not visible until the zone is expanded. `Comparison_Model_Formula` reads it
regardless; hidden columns still calculate.

> **SUPERSEDED** by v3.4+ *The spilled §4b zones are no longer grouped or
> collapsed*. The zone ships expanded, so this cost is not paid: the caption is
> visible. `Comparison_Model_Formula` reads it by name either way.

`Comparison_Model_Formula` is what makes the move free. The v3.4 reading surface
is a sheet-scoped NAME, and its `RefersTo` is now built from the layout
constants (`_abs_ref(_ROW_MODEL_FORMULA, _C_MODEL_FORMULA)`) rather than the
literal `$AB$2` it shipped with — the address was the last hardcoded A1 string in
`_setup_local_names`.

### The assembly became a catalog LAMBDA — sheet-scoped, not workbook-scoped

**Question:** the readout's formula was a ~300-character inline concatenation
built by string-formatting four Python constants together at build time. Extract
it into a catalog LAMBDA, and at what scope?

**RESOLVED** — `Model_Formula`, `scope: "Regression"`, under the existing
`Sheet-Scoped Constructors` subcategory; the cell holds `=Model_Formula()`. The
catalog is where an expression of this size is documented (it gets a row on the
LAMBDA_functions sheet with its `yields`, description and plain-language summary
like every other function), and the four assembly rules — Log-wrapped response,
`"1 + "`/`"0 + "` intercept prefix, `TEXTJOIN` over `Constructed_Column_Names()`,
FE suffix gated on the Fixed Effects count — stop being a build-time Python
concatenation nobody can read in Excel.

**REJECTED — workbook scope.** The body reads `Spec_Role`, `Spec_Transform`,
`Header_Names`, `Allow_Intercept` and `Constructed_Column_Names()`, every one of
which is sheet-scoped, so a workbook-scoped definition would resolve against
whichever sheet is literally named `Regression` — wrong in a workbook with 47
Regression-shaped sheets (the test-model artifact) and `#NAME?` in one with none.
This is exactly the `Base_Period_Delta` case CLAUDE.md § *Workbook scope belongs
to the catalog* records, and the same resolution applies: sheet-scoped, with
unqualified names that resolve against the calling sheet.

No new test-model case is warranted. The corner is not a new modelling corner —
every existing spec case and guard case already asserts the readout's text
against `_build_model_formula`, an independent Python mirror, so a broken
`Model_Formula()` fails all of them at once. What the oracle chain gained is
that it now covers the catalog body rather than an inline cell formula.

---

## v3.4+ — Ladder ordering and the test-model suite

### The post-v3.3 ladder: Regression work first, then test-suite growth

**Question:** in what order should the remaining planned milestones ship?
Through v3.3 the order was inherited from the original v2.x feature train,
renumbered but never re-argued, with five candidates sitting in an
unordered "v3.8+" bucket.

**Resolution:** RESOLVED — every milestone from v3.4 on is sequenced by two
keys, in this order:

1. **All remaining Regression work ships first.** A milestone that extends
   the Regression sheet, its spec block, or its engine precedes either
   milestone that opens a *new* analysis surface. Two-sample and Resampling
   are the only two of the latter, and they go last as a block.
2. **Within the Regression track, by how much the test-model suite has to
   grow** — additive first, per-model multipliers next, axis-wideners last;
   within a tier, the most commonly used feature first.

| Track | Tier | Milestones |
|---|---|---|
| Regression | additive | v3.4 Model Comparison |
| Regression | near-additive | v3.5 `Cluster` · v3.6 `Time` / time series |
| Regression | ~2× | v3.7 WLS · v3.8 Two-way FE |
| Regression | ~10× axis-widener | v3.9 standalone transform library |
| New surface | additive | v3.10 Two-sample · v3.11 Resampling |

Three milestones changed number and three left the unordered bucket:
`Cluster` → **v3.5**, `Time` / time series → **v3.6**, two-way FE →
**v3.8** (all promoted); the standalone transform library out of the v3.3
remainder to **v3.9**; Two-sample v3.6 → **v3.10** and Resampling v3.5 →
**v3.11**. WLS holds **v3.7**, the number it was claimed under, but reaches
it as the first ~2× item in the Regression track rather than by inheritance.
v3.3 keeps its number for the half that shipped. What was left unordered —
ANOVA, Fourier, decision analysis — stays unordered as **v3.12+**, because
nothing about their test cost sequences them either.

**Rationale, key 2.** The suite is a covering array over the implemented
feature axes, so a feature's cost is not the code it adds but the *cross* it
forces: a new `Transform` value multiplies the response × predictor dispatch
table that every model is scored against, while a new variance estimator only
varies a handful of existing models. Ordering by that number keeps the suite
growing linearly for as long as possible and lands the multiplicative lifts on
the most mature harness. Two side effects: `Time` moves early, and its
calendar-dated dataset closes the **one Section-1 coverage gap that exists
today** (the Sequence calendar-signature verdict, untestable because no wired
dataset carries real dates); and WLS sits behind two cheap milestones without
losing anything, since its own sequencing constraint — ship after v3.0 so √w
scaling is the first implementation rather than a rewrite — is already
satisfied.

**Rationale, key 1, and why it outranks key 2.** Test cost is the right
tiebreaker *within* one artifact and the wrong primary key across two. Every
Regression-track milestone extends a surface that already exists and is
verified by the harness that already exists — a spec column, an engine change,
more cases in the same oracle. Two-sample and Resampling each need a new sheet
writer, a new layout, and a verification path sharing nothing with
`calculate_regression_spec_case`. Interleaving them means carrying two
half-built analysis surfaces at once and leaving the artifact users actually
have feature-incomplete for longer while effort goes elsewhere. The deferral
costs no rework: neither milestone depends on any Regression milestone, and
none depends on them, so both cost the same whenever they are built.

**The inversion is deliberate.** v3.10 and v3.11 are *cheaper* to test than
four of the milestones ahead of them and still ship last. That is key 1
overriding key 2, recorded explicitly so a future reader does not "correct"
the ladder back to pure test-scale order.

**REJECTED — ordering by user-facing value alone.** It is the tie-breaker
*within* a tier, not a key of its own. Value-first ordering is what put the
~10× transform library at v3.3, immediately after the milestone with the most
axes to cross.

**REJECTED — pure test-scale ordering across both tracks.** That was the
first form of this decision, and it interleaved the two non-Regression
milestones at v3.5 and v3.6, ahead of every multiplier. It optimizes the
suite's growth curve at the cost of the artifact's completeness, which is the
wrong trade for a tool with one user waiting on the Regression workbook.

**Not frozen.** The tool is single-user and pre-release. A user pressing for a
milestone reorders it; the rule is that
[docs/MODEL_TESTING_ASSETS.md](MODEL_TESTING_ASSETS.md) § 2 is edited
first and the [ROADMAP.md](ROADMAP.md#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
ladder second, so the two never disagree about why the order is what it is.

### The test-model plan is a document, not a test file

**Question:** where does the plan for the regression test-model suite live —
in `tests/`, as a docstring or a list of pending cases, or as prose?

**Resolution:** RESOLVED — as prose, in
[docs/MODEL_TESTING_ASSETS.md](MODEL_TESTING_ASSETS.md), and the code
holds only cases that actually run. The plan carries things a test file
cannot: the coverage matrix (which feature corner each model is there for),
the ~15 configurations not yet declared anywhere, the datasets future
milestones will need and what each one buys, and the covering-array rule that
bounds the suite's size. A skipped or commented-out case in `tests/` would
express none of that and would rot silently.

The half that *is* enforced in code stays enforced: `_EXPECTED_CASE_NAMES` in
`tests/test_regression_spec_qc.py` is an ordered list asserted against
`build_regression_spec_cases()`, so no case can be added, renamed, reordered,
or dropped without a test failure. The document says what should exist; the
pinned list says what does.

## v3.4+ — Test-model oracles and the one-sheet-per-model framework

### Guard states get their own case type, not a flag on `RegressionSpecCase`

**Question:** the § 1.4 guard-rail configurations are not fittable models.
Should they be `RegressionSpecCase` entries with an "expect failure" flag, or
a separate type?

**Resolution:** RESOLVED — a separate `GuardStateCase` in
`lambda_catalog/analyze_regression_guard_states.py`. The two assert disjoint
things. A fittable case's oracle is ~200 numbers from a fit;
`calculate_regression_spec_case` raises on most guard specs *by design*,
because a QC case must describe a legal, fully-computable model. What a guard
case asserts is status text, the per-row Design Columns audit, the Model
Formula string, and which conditional-formatting rules fire. Bolting that
onto `RegressionSpecCase` would mean a dataclass where most fields are
meaningless for half its instances, and an oracle function with a top-level
branch that shares nothing below it.

### Conditional formatting is asserted as a predicate, never as a colour

**Question:** how does a guard case verify that a cell is flagged red or
amber — read `Range.DisplayFormat.Interior.Color` over COM, or recompute the
rule?

**Resolution:** RESOLVED — recompute the predicate. `GuardFlag` records that
a rule *fires*, derived in Python from the same condition the CF expression
encodes. Reading the colour back would only re-report what Excel already
decided from the rule that is under test: if someone changed the CF
expression, Excel would faithfully apply the *new* rule and the colour check
would pass. Recomputing the predicate independently is what makes a silent CF
change fail. It is also the only form that runs headlessly, so the guard
oracles are covered in CI while the sheet comparison is not.

### One sheet per test model, in a third non-shipped artifact

**Question:** the suite pushes every case through the single `Regression`
sheet in turn. Should each case get its own sheet, and if so, in which
workbook?

**Resolution:** RESOLVED — one sheet per case, in a new
`Lambda_Library_TestModels.xlsx` built by `build_test_models.py` and
gitignored.

The single-sheet harness has three costs that are all consequences of reuse.
A case exists only as a log line, so a failure gives a number with nothing to
open. Every case must defensively re-set every input in case the previous one
left something behind — which is precisely why `source_table_ref` and
`prediction_group` are non-Optional and rewritten on every iteration. And the
~150–250 COM writes plus a recalculation per case make the loop serial and
slow. With one sheet per case the verifier only reads: no writing, no
per-case recalculation, no state to leak, and a failing case is a tab.

A third artifact rather than the QC or production workbook: ~48 heavy sheets
have no business in a shipped file, and folding them into the artifact-specific verify builds would
make every QC run pay for them. It is a fixture regenerated from the case
registries on demand, so it is not committed.

REJECTED — a lean engine-only sheet (spec block plus scalar zones, no
residual band). It would drop the per-observation residual comparison, which
is the largest part of the oracle by count and the part most likely to catch
a masking or ordering bug. Charts are dropped instead: a dozen COM chart
objects per sheet is the single biggest cost in the build, and no oracle
reads one — chart wiring is verified once, on the production sheet.

### Sheet names state the concept, not the variables

**Question:** what does a generated sheet's tab say?

**Resolution:** RESOLVED — `<PlanID> <Concept>`: `M05 Log-Log NA Masking`,
never `MPG ~ Ln(Weight) + Ln(HP)`. Excel allows 31 characters, which cannot
hold a model formula for any case worth testing, and truncating one produces
a tab that is both unreadable and ambiguous. The formula is also the least
useful thing to put there — the sheet exists to exercise one corner, that
corner is what a reader needs, and the variables are one click away in the
spec block. The plan ID prefix ties each tab to a row in
docs/MODEL_TESTING_ASSETS.md, so a case can be renamed without orphaning the
mapping. The contract is enforced at registry-build time
(`lambda_catalog/test_model_sheets.py`), so an illegal or duplicated name
fails in a millisecond-long unit test instead of at sheet 30 of a
multi-minute Excel build.

### Three planned models did not survive contact with the data

**Question:** what happens when a model the plan specifies turns out not to
be fittable?

**Resolution:** RESOLVED — implement the corner the plan wanted, by the
nearest configuration that actually fits, and record the deviation in the
plan document rather than silently substituting.

Three cases needed it. The plan's M9 (`C(Cylinders) × C(Origin)`) has a
sparse cross-tabulation — Cylinders=3 occurs only in Asia, 5 only in Europe,
8 only in the US — so two of its eight product columns are identically zero
and the Gram matrix is singular; Excel answers `#NUM!` where NumPy quietly
returns a minimum-norm solution, so the two sides cannot be compared at all.
`Model Year × Origin` populates all 39 cells (minimum cell count 2, condition
number ~116) and is a genuine saturated design. The plan's M10
(`Displacement + Horsepower + Displacement − Horsepower`) is exactly
collinear by construction. And L7's arithmetic (193 countries → 192 dummies +
8 predictors = 200) ignores missingness: at most 183 countries survive the
mask, so `C(Year)` is added to clear the width-guard threshold, pinned by a
test so a future change cannot silently drop the case back under 200.

### L6 contradicts the shipped mask — flagged, not fixed

**Question:** § 1.2 says a Log transform on a column with true zeros makes
the row "drop out of the mask". It does not. Fix `Sample_Include`, or record
the actual behaviour?

**Resolution:** RESOLVED for now — record it. `Sample_Include` tests
`ISNUMBER(col)` on the Response and the included Continuous Predictors, and
`ISNUMBER(0)` is TRUE; there is no Log-positivity term anywhere in it. So
Schooling's 28 zero rows stay in the sample, `Ln_Positive` returns `#N/A` for
each, and the `#N/A` propagates into every downstream statistic. L6 ships as
a guard state asserting that propagation.

DEFERRED — whether to add a positivity term to `Sample_Include`. It would
make the plan's description true and is arguably the better behaviour, since
the Response's own non-positive guard effectively achieves it by raising. But
it changes a constructor every model in the workbook depends on, and making
that change as a side effect of writing an oracle would be exactly the kind
of silent scope creep this file exists to prevent.

### `_build_model_formula` was never compared against the cell it mirrors

**Question:** the oracle's model-formula string and the sheet's AB2 cell were
written independently and never string-compared. Which is right?

**Resolution:** RESOLVED — the cell. `_build_model_formula` now mirrors AB2
character for character, and the guard cases compare it. It had three
divergences, all invisible because nothing read it: it omitted the intercept
term entirely when the intercept was OFF (AB2 writes `"0 + "`), it treated
the intercept marker as the first element of the join rather than a prefix
(so an empty model rendered `"MPG ~ 1"` instead of AB2's `"MPG ~ 1 + "`), and
it appended `"1"` as a fallback for an empty predictor list. The same pass
found `Base_Period_Delta_Candidate`'s mirror using `statistics.mode`, which —
unlike Excel's `MODE.SNGL` — never errors on a non-repeating input, making
both the fall-back-to-MIN branch and the "no natural base period" verdict
unreachable in the oracle.

### The first live Excel run settled three things the headless suite could not

**Question:** a legacy combined verifier run reported 22,898 mismatches and
`build_test_models.py` died on its first sheet. Which were real?

**Resolution:** RESOLVED — all of them, and none was a false alarm. Recorded
together because they share a cause: 878 headless tests cannot see what
Excel does with a sheet name, a 205-column design, or an accented string.

1. **Sheet names with spaces were never quoted.** `_setup_local_names` built
   `=M01 Baseline Categoricals!$AB$12` and Excel rejected the whole
   `Names.Add`. Four of that function's seven references were unquoted,
   invisible for the life of the project because the only sheet it ever
   wrote was named `Regression`. `sname` now carries the quotes — quoting a
   name that does not need it is always legal, so one quoted form removes
   the class of bug rather than relying on each site to remember. Two
   existing tests had pinned the unquoted form; they were pinning the bug.

2. **A 205-column design cannot be fitted, so L07 is a guard state.** At
   k = 205 the workbook returns `nan` for every engine output — 22,886 of
   the 22,898 mismatches were this one case comparing real numbers against
   nothing. That is the condition the width guard exists to warn about, so
   the case now asserts the `M2` WARNING and the visible degradation instead
   of numbers the sheet cannot produce. The general rule, which will come up
   again: **a numeric oracle for a model the sheet cannot compute is
   comparing against nothing** — when a case's whole point is a limit, the
   limit is the assertion.

3. **Categorical levels sorted by code point, not by collation.** `Côte
   d'Ivoire` files after `Czechia` in Python but between `Costa Rica` and
   `Croatia` in Excel, which shifts an entire dummy block by one position.
   Nothing errors: every column keeps a valid name and a valid 0/1 pattern,
   and every per-predictor statistic is silently paired with the wrong
   header. A pre-existing bug in `_retained_levels`, first reachable when a
   QC case used a column with non-ASCII levels.

   REJECTED — PyICU for exact collation. It is the only way to match Windows
   collation across all scripts, but it adds a binary dependency to a project
   whose only ones are the scientific stack, to serve one accented string in
   one dataset. `level_sort_key` strips combining marks and casefolds
   instead, which reproduces the locale order for Latin scripts, and the
   limitation is documented at the function and in MODEL_TESTING_ASSETS §1b.

**And one tolerance decision.** L05's `Population` coefficient disagreed with
Excel in the 6th significant digit on its t-statistic and p-value. Population
spans 34 to 1.3e9 against predictors of order 1–100, so the normal equations
are badly scaled and the disagreement is on the one coefficient
indistinguishable from zero (t = −0.367, p = 0.71). `T_Statistics` and
`P_Values` join the scale-free comparison set alongside the sums of squares,
which already compare as 3 significant digits for the same reason. This
widens the unit, never the tolerance — both sides are divided by the same
factor, so a genuinely wrong number still fails. Dropping `Population` was
rejected: L05 exists to give the *shipped* profile an oracle, so changing its
spec would stop it testing what users actually get.

### A spec block must have exactly one row per Source_Table column

**Question:** the first successful build of the test-model workbook produced
74,065 mismatches, with entire cases reading `None`. Why, and what enforces
it in future?

**Resolution:** RESOLVED — the fixture column `Is_USA`, added to
`MileageData` for M15's benefit, widened that table to 13 columns while every
other Auto MPG case still wrote a 12-row spec block. Every constructor opens
`n_c = COLUMNS(Source_Data)` and then indexes the `Spec_*` bands at `1..n_c`,
so `INDEX(rl, 13)` on a 12-row band runs off the end: the row mask errors and
every engine cell downstream reads as an error. M15 was the one Auto MPG
sheet that worked, because its spec happens to declare the thirteenth column.

Two things were wrong, and only one of them was the fixture. The **shape
invariant** — one spec row per source column — was never written down or
checked anywhere, even though every constructor depends on it. And the
fixture list existed in **two places**: the driver knew to add the column to
the data sheet, and the spec writer did not know it existed. Either alone
would have been survivable; together they produced a build that succeeded,
looked right, and was wrong in a way visible only to a verify run.

The fix states the invariant and enforces it.
`write_sheet_test_model.pad_spec_to_source_table` appends an `Omit` row for
every source column a case does not name, and raises if the result is still
the wrong width. `Omit` is the correct filler rather than a convenient one:
it contributes no design column and imposes no mask condition, so a padded
spec fits exactly the same model — which is precisely why the Python oracle
can remain ignorant of fixture columns. `FIXTURE_COLUMNS` is now declared
once and read by both the data-sheet writer and the spec block.

REJECTED — giving M15 its own copy of the Mileage data sheet so the shared
table stays clean. It isolates the fixture, but it leaves the shape invariant
unstated and unchecked, so the next fixture column reintroduces the same
class of failure. Padding fixes the general case; a private data sheet fixes
one instance of it.

### `Base_Period_Delta` is sheet-scoped, not workbook-scoped-and-sheet-qualified

**Question:** the accessor's body hardcoded `'Regression'!Source_Data` /
`Spec_Sequence` / `Spec_Sequence_Period` at workbook scope. In a workbook
with 47 Regression-shaped sheets, which sheet's Δ is "the" Δ?

**Resolution:** RESOLVED — each sheet's own. `Base_Period_Delta` moves to
`"scope": "Regression"` with unqualified spec references. An unqualified name
resolves against the sheet the calling formula lives on, so one definition
per Regression-shaped sheet gives each its own Δ, and the sheet-qualified
form disappears from the catalog entirely.

The old form was wrong in both directions at once. In a workbook with
several such sheets, every one of them read whichever sheet was literally
named `Regression`. In a workbook with none — the test-model artifact — the
build correctly skipped the function rather than let Excel rebind it to an
external workbook, which left `#NAME?` at every call site: the BFN Panel
Durbin-Watson cell on all 47 sheets, and any omitted-delta `Lag_By` /
`Difference_By`.

**The narrow cost, recorded so it is not a surprise.** An omitted-delta
`Lag_By`/`Difference_By` evaluated on a sheet with no spec block — a data
sheet, say — now returns `#NAME?` where it previously borrowed the
Regression sheet's Δ. That is the honest answer: such a sheet declares no
sequence axis, so it has no base period to default to. The alternative,
silently reaching across to another sheet's spec block, is the behaviour that
made this wrong on a multi-sheet workbook in the first place.

**This requires rebuilding the committed artifact.** `Lambda_Library.xlsx`
carries the name at workbook scope from its last build, and
`sync_workbook_names` only makes workbook scope match the catalog — the
sheet-scoped replacement is installed by the sheet writer, which needs Excel.
Until `build_production.py` is re-run and the artifact committed,
`test_regression_workbook_scope_belongs_to_the_catalog` reports the stale
name as residue, which is exactly what it is.

### Auto MPG ships no Sequence axis — `Model Year` was never one

**Question:** the shipped T0 spec flagged `Model Year` as
`Sequence = TRUE`, and every Auto MPG QC case inherited the flag from it.
Is that a correct description of the dataset?

**Resolution:** RESOLVED — no. `_DEFAULT_SEQUENCE_VARIABLES` becomes empty,
and the flag is removed from every Auto MPG spec case.

Auto MPG is cross-sectional. Each row is a distinct car model observed once;
no unit is repeated across periods, so there is no axis to order along. The
Sequence flag is not a formatting preference — it activates the Base Period Δ
candidate, the Sequence Spacing block, and the gated Durbin-Watson
diagnostic, all of which presuppose a panel. What the flag actually bought
here was a Δ candidate nobody can interpret and a DW statistic computed over
an arbitrary row order. It did not even light up the spacing verdict:
`Sequence_Deltas` groups by the **Identifier** columns, and the shipped
Identifier (`Car Name`) is very nearly unique, so every group is a singleton,
there are no within-group consecutive pairs, and the verdict cell was
unconditionally blank the whole time. A default that asserts panel structure
the data does not have is worse than no default.

**What keeps its flag, and why.** The two datasets that *are* panels keep
theirs: Life Expectancy (`Year`, country × year) and Production Lots
(`Fiscal_Year`, facility × fiscal year). Both ship it in their
`SpecDatasetProfile`, both are reachable through `--regression-dataset`, so
the Sequence layer is still demonstrated by default — on data where it means
something. On Auto MPG the layer now self-reports `n/a — requires Sequence`
until a user types TRUE into column H, which is the honest state.

**Two guard cases still declare it, explicitly.** G3 (two flags → the H2
cardinality error) and M16 (typed period override) test the flag's
*mechanics*, not the data: H2 counts flags, and the override path resolves
the flagged row positionally. Both are dataset-independent and unreachable
without a flag present, so each now states its own rather than inheriting
one, with a comment saying it is wiring and not a claim. Auto MPG's evenly
spaced integer years remain a clean substrate for M16's candidate Δ = 1
against a typed Δ = 2.

**REJECTED — keeping the default so the flagship artifact still demos the
feature.** That was the only real argument for the old behaviour, and it does
not survive the observation above: demonstrating the serial-correlation layer
on data with no serial structure teaches the wrong lesson, and the two panel
profiles demonstrate it correctly at no cost.

Pinned by `test_sequence_is_flagged_only_on_datasets_that_have_an_ordering_axis`
(no fittable Auto MPG case flags anything; every Production Lots case flags
`Fiscal_Year` and every Life Expectancy case `Year`) and
`test_only_the_two_mechanics_cases_flag_sequence_on_auto_mpg` (G3 and M16 are
the entire exception list). **This changes the shipped spec block**, so
`Lambda_Library.xlsx` needs the same rebuild the `Base_Period_Delta` scope
change above already requires.

### The pre-derived/transform-axis pairing runs twice — with FE and without

**Question:** Production Lots ships both raw columns (`Cumulative_Units`,
`Unit_Cost_BY`) and pre-derived log columns (`log Cum Units`, `log Unit
Cost`) that are exact logs of them. P1/P2 already fit the same model by both
routes under Fixed Effects. Is one such pair enough?

**Resolution:** RESOLVED — no. A second pair, **P3 (pre-derived) / P3b
(transform axis)**, fits the same model without Fixed Effects. The existing
`production_lots_log_no_fe` case becomes P3b; the new
`production_lots_derived_log_no_fe` is P3.

One pair leaves the two axes entangled. The mechanisms reach the design
matrix by different code paths — one *reads* a column, the other *computes*
one — and composing either with FE demeaning is a third path again. With only
the FE pair, a transform-axis regression can hide behind the demeaning, or a
demeaning regression behind the transform. P3/P3b has no Fixed Effects, so
the transform axis is the only thing between the CSV and the design matrix
and a disagreement can only be the transform wiring.

It is also the cheapest strong oracle the suite can buy. Because the shipped
log columns are exact logs, the pair must agree **bit-for-bit** on the design
matrix and response vector (`np.array_equal`, not `allclose` — the default
`rtol=1e-05` would swallow a real regression) and to floating point on every
downstream statistic, with neither side reading the workbook.
`test_no_fe_pair_agrees_the_same_way_the_fe_pair_does` asserts it, mirroring
the P1/P2 assertion next to it.

**What the pair may NOT agree on**, pinned so it cannot silently collapse
into two copies of one spec: `constructed_column_names` ("log Cum Units" vs
"Ln(Cumulative_Units)"), `constructed_column_transforms` (`None` vs `Log`),
and the response display name. Those are the mechanism showing through the
label, not a disagreement about the fit, so the cross-check compares numerics
only.

**Ordering and naming are part of the decision.** Each pair is registered
adjacently so the two land on adjacent worksheets, and the sheet names state
the *route* — `P03 Power Law Derived Cols` against `P03b Power Law Transform
Axis` — because the route is the only thing that differs between the tabs and
the entire reason both exist. The old name, `P03 Power Law No FE`, described
what the pair has in common rather than what separates them.

**REJECTED — pre-derived twins for P4 and P5 as well.** That would give all
four transform-dispatch branches a matched pair, but the extra two prove
nothing the P3/P3b pair does not: P4's `(Log, Mixed)` and P5's `(None, Log)`
already have their dispatch branches covered, and a twin would only re-assert
that a log column equals a logged column, which one no-FE pair settles. The
suite is a covering array; three cases for one claim is a full cross. Suite
size goes 32 -> 33 fittable rather than 32 -> 35.

**REJECTED — a case for the unused `log experience` column.** No case
references it. The two-predictor shape is covered by P4 using raw
`Experience_Stock`, and swapping in the logged column would make that model
`(Log, Log)` rather than the `(Log, Mixed)` dispatch P4 exists for — a
different model, not a twin.

### The Back-Transform caveat becomes a note on its own control

**Question:** the Duan/Naive caveat shipped as a merged, wrapped text row at
`AJ15:AL15` — the last row of the Prediction Outputs zone. Is that where it
belongs?

**Resolution:** RESOLVED — no. It is now a cell note on the **Back-Transform
label at `AG4`**, and row 15 is gone; the zone's border box closes at row 14,
the last row that holds a value.

The text explains what the `AH4` toggle does and why the point estimate sits
off-centre in its interval. Three zones away and below the interval it
qualifies, it read as a footnote to the prediction block — something to
notice after the fact — rather than as documentation of the control that
causes the behaviour. On the toggle's own label it is where a user looks
when deciding which method to pick, which is the moment the explanation is
worth anything. It also matches how every other explanatory text on this
sheet is delivered: a note on the header of the thing it describes.

The note lives in `_write_model_specification` with the other `_set_note`
calls, not in `_write_unit_space_block`. `AddComment` is COM-only, and
keeping it out of the zone writers is what lets them stay exercisable through
`RecordingSheet`.

### The Back-Transform dropdown offered `"Duan` and `Naive"`

**Question:** the `AH4` list validation was built as
`Formula1=f'"{",".join(_BACK_TRANSFORM_METHODS)}"'`. Why did the dropdown
show quote characters?

**Resolution:** RESOLVED — because they were real. Excel's `xlValidateList`
takes its items as a bare comma-separated string, `Duan,Naive`. The quotes
that appear around it in VBA examples are that language's string delimiters,
not part of the value; wrapping the Python string in literal `"` characters
passed them through COM as data. The dropdown offered `"Duan` and `Naive"`,
the validation accepted either, and every consumer — `Unit_Space_R_Squared`,
`Unit_Space_RMSE`, the `AL3` point estimate, the `AZ`/`BA` residual columns —
matched neither against a recognised method. Fixed to `",".join(...)`, the
same form as `_INCLUDE_VALIDATION_LIST` (`"TRUE,FALSE"`) in the spec block,
which was correct all along.

**Why no test caught it, which is the more useful half.** `RecordingValidation.Add`
in `tests/recording_sheet.py` required an `Operator` keyword. Excel treats
`Operator` as optional and a list validation does not need one, so the
Back-Transform call omitted it — and the recorder raised `TypeError`, which
the writer's own `except Exception: pass` around the validation block
swallowed. The rule was silently never recorded, so no assertion could see
it. `Operator` is now optional in the recorder, and
`test_write_unit_space_block_writes_section_input_and_three_gof_cells`
asserts the parsed dropdown items equal `_BACK_TRANSFORM_METHODS`. Confirmed
it fails on the old form, reporting `dropdown offers ['"Duan', 'Naive"']`.

The general lesson is about the guard, not the quotes: a broad `except` around
a COM call also silences the test double, so a writer whose only verification
runs through `RecordingSheet` has no coverage at all inside such a block.
Where a mock's signature is stricter than the API it stands in for, the guard
turns that mismatch into silence.

### Spec-case oracles are not cached — measured, not assumed

**Question:** the QC oracles were to gain the disk-cache treatment
`analysis_cache.py` already gives the legacy MLR/regression-sheet configs, so
they are not recomputed on every run.

**Resolution:** REJECTED, on measurement. Computing **every** oracle in the
suite — all 33 fittable `RegressionSpecCase` results plus all 16
`GuardStateCase` results, across all three datasets — takes **1.71 s**.
*(Since settled the other way too: `analysis_cache.py` was deleted outright
once the legacy configs it served were the only thing still reading it.)* The
slowest single case is 0.15 s (L03, at 2938-row scale); the whole Life
Expectancy block is about half the total.

That is not worth a cache. The one the legacy path uses costs a schema
version to maintain (already at 19, bumped six times for field changes), a
serialize/deserialize pair per dataclass, a fingerprint check, and a class of
staleness bug where a code change that is not a schema change silently serves
a wrong expected value. Trading that for 1.7 s is a bad trade, and the risk
lands precisely on the numbers the whole suite exists to be sure about.

For scale: the artifact this feeds, `build_test_models.py`, takes **~82
minutes** to write and recalculate its 48 sheets. The oracle phase is under
0.04 % of it. Even the doubled cost in a `--verify` run — the build computes
every oracle, then the verifier computes them again — is 3.4 s.

**What the measurement does NOT say.** The legacy cached path costs about the
same (1.65 s), so this is not a claim that the existing cache was a mistake;
it is a claim that a second one buys nothing today. If a future milestone
adds a genuinely expensive oracle — a bootstrap or permutation case under
v3.11 Resampling, say, where the cost is in resamples rather than a single
fit — revisit it for that case rather than for the suite.

### BFN panel Durbin-Watson joins the compared scalars

**Question:** the sheet's `AE12` BFN cell was compared by neither harness and
had no oracle field. Was that deliberate?

**Resolution:** RESOLVED — no; nothing in this file recorded a reason, and the
gap left a real hole. `RegressionSummary` gains
`bfn_panel_durbin_watson`, `analyze_regression_sheet` computes it, and
`regression_spec_sheet_io` compares it at `ROW_BFN_PANEL_DW`.

The hole was specific and worse than "one cell unchecked". `durbin_watson` is
set to NaN whenever Fixed Effects are declared — correctly, because `AE11`
then reads `n/a — FE active` — so on a Fixed Effects sheet the suite verified
**no serial-correlation diagnostic at all**. The one cell holding a number
was the one nobody read.

**The oracle mirrors the sheet's gating rather than always computing.** The
two cells are mutually exclusive and at most one is ever a number, so the
oracle NaNs whichever one the sheet renders as text — otherwise it would
offer a value for a cell displaying a string, and the mismatch would be
blamed on the workbook. `test_dw_and_bfn_are_never_both_live_in_the_registry`
asserts that across every fittable case, and that plain DW is live exactly
when no FE row is declared.

**Making the cell live required a typed Sequence Period, which is the
interesting part.** `Base_Period_Delta()` is the *override* accessor — it
reads the typed value in spec column I and returns `#N/A` when blank, never a
silent 1 — and BFN passes it as Δ. No fittable case typed one, so BFN would
have been `#N/A` everywhere and every comparison vacuously true.
`RegressionSpecCase` gains `sequence_period`, and **P01/P02** declare 1:
Production Lots is an annual panel, so it is a true statement about the data
rather than wiring for its own sake, and the pair's cross-check now extends
to BFN (they agree at 0.9854876217402373). **L08 deliberately leaves it
untyped**, keeping one registered case on the honest `#N/A` path.
`test_only_cases_that_need_a_period_declare_one` pins that a case typing a
period has both a Sequence axis and Fixed Effects, so the field cannot spread
by copy-paste the way the Sequence flag once did.

**The oracle is checked against an independent value.**
`tests/test_bfn_panel_durbin_watson_verification.py` already agreed on
0.6362023311147436 for `Life expectancy ~ GDP + Schooling | Country` by two
paths sharing no implementation (statsmodels LSDV residuals with an explicit
per-group loop; a within-estimator fed through the `Difference_By` mirror).
The new oracle is a third path and agrees to within 1e-12 relative.

**Not to the last bit, and the first revision of this entry was wrong to say
so.** The assertion originally demanded exact equality; it passed on one
machine and failed in CI at 3 ULP (`…433` against `…436`). The three paths
reach the statistic through *different fits* — LSDV with per-country dummies
on one side, the within-estimator on the other — so the residual vectors they
sum differ in the last bits, by an amount that depends on which BLAS the
runner links. Bit-exactness is a legitimate claim only where the arithmetic
is genuinely identical, as in the P01/P02 and P03/P03b pairs, where
`np.array_equal` compares a stored log column against a computed one. Here it
was a coincidence of build being asserted as an invariant. The tolerance is
still four orders tighter than the six-decimal first-differing-digit rule the
workbook comparison uses, so any drift that could reach a QC result fails
here first.

**Two test-double defects surfaced on the way, both the same shape.**
`RecordingSheet.range` keyed its store on the raw argument tuple, so
`range(row, col)` and `range((row, col))` were different slots for one cell —
xlwings means the same cell by both, so a writer using one spelling was
invisible to an assertion using the other. That is why the typed period
appeared unwritten. Normalised. Together with the `Validation.Add(Operator=…)`
mismatch fixed alongside it, the pattern is worth naming: **where the double's
contract is narrower than the API it stands in for, the writer's own
`except Exception: pass` turns the mismatch into silence rather than a
failure.**

---

## v3.4+ — The spec block sizes itself; `SpecTable` removed

**RESOLVED — the `Spec_*` bands and the four computed spec columns derive their
height from `COLUMNS(Source_Data)`, and the `SpecTable` ListObject is gone.**

The Regression Instructions sheet has always told the user the dataset
changeover is a one-name edit: update `Source_Table` and "the header row, the
data body, and the variable list in the MODEL SPECIFICATION block all update
automatically." Two of those three were true. The variable-name column is a
`=TRANSPOSE(Header_Names)` spill and resized correctly; the spec rows under it
did not.

`SpecTable` was a ListObject created at build time over `B3:O{4 + N - 1}`, where
`N = len(profile.variables)` — `B3:O15` in the shipped artifact. The `Spec_*`
band names bound to it as `SpecTable[[#Data],[Role]]` structured references, so
the bands were exactly as tall as the table. Retarget to `LifeExpectancyData`
and column A spilled 23 names above a 12-row spec: `TAKE(band, 23)` returned 12
rows, because `TAKE` does not pad, and `INDEX(rl, 23)` ran off the end into
`#REF!` — which propagated through every constructor to every engine cell. The
same failure mode `effective_variables` already records from the 74,065-mismatch
build, reached by a different route.

**Why the table could not stay.** Excel exposes no formula-driven way to resize
a ListObject; only VBA or a user typing into the row below its bottom edge
extends one, and this workbook is macro-free by design. So any fix that kept the
table could only pick a generous fixed ceiling and hope no dataset exceeded it.
The self-sizing alternative requires the computed columns to be dynamic arrays,
and **a spill cannot live inside a ListObject** — J/K/L sit between the input
columns I and M, and O sits after N, so there is no arrangement that keeps the
table and makes them spills without reshuffling columns. That reshuffle is
exactly what was refused when M/N were appended rather than inserted (v3.1),
and the reasons have not changed.

**What replaced it.**

* `_spec_band` builds every band as `=TAKE($X$4:$X$16000,MAX(1,COLUMNS(Source_Data)))`.
  `TAKE` rather than `OFFSET` for the same reason `Source_Data` and
  `Header_Names` use it — non-volatile, so it is not re-evaluated on every Data
  Table substitution pass. `MAX(1,…)` keeps the name resolvable mid-retarget,
  when a zero-row `TAKE` would error into every dependent name.
* The four computed columns are one spill each at `_FIRST_DATA_ROW`, written
  through `Formula2`. `MAP(SEQUENCE(nc),LAMBDA(i,…))` rather than `BYROW`,
  because the bodies need the column *index* to reach `INDEX(Source_Data,0,i)`
  and `BYROW` passes a row's values, not its position. That index also retires
  the `ROW()-_ROW_TO_COL_OFFSET` arithmetic, so a formula no longer depends on
  which row it occupies.
* The input band's `INPUT_COLOR` fill moved from per-row `format_input` calls to
  one lowest-priority CF rule on the same predicate. The per-row version painted
  only the build-time profile's rows, so a retarget left its new rows functional
  but unpainted — the same build-time pinning as the bands, in the styling layer.

**Three things this did not have to touch, and the reason it was tractable at
all.** `lambda_functions.json` needs no edit: all 45 functional `Spec_*` reads
were already `TAKE`-trimmed to `COLUMNS(Source_Data)`, so their trim became an
idempotent no-op over a band that is now the right length by construction. No
conditional-formatting expression changed: every one already used relative A1
references (`=$B4<>"Predictor (x)"`), never structured references. And every CF
and `Validation` range already ran to 16000, so the rows a retarget reveals were
already covered — that ceiling is now shared with the bands as
`_SPEC_BAND_LAST_ROW`, precisely so the three cannot disagree about how far the
block may grow.

**A build-order inversion falls out of it.** `_setup_local_names` now runs
*before* `_write_spec_block`. The dependency reversed direction: the spills
reference the bands, `Source_Data`, `Header_Names` and the constructor closures,
where the bands used to reference a table the block had to create first (Excel
validates a name's `RefersTo` at `Names.Add` time). The comment at that call
site records the reversal, because the old ordering looks arbitrary without it.

**Two constraints this creates.** Nothing may ever be written below the spec
block in columns B–O: the bands reach row 16000 and the spills need clear space
beneath them, so stray content is a `#SPILL!` error rather than a quiet
truncation. And a `SpecDatasetProfile` now governs which rows arrive with
shipped *defaults*, never how many rows exist — a distinction worth keeping
straight, since the parameter's name suggests otherwise.

**Removing the table also removed machinery that existed only to serve it.**
ListObject names are workbook-scoped, so the one-sheet-per-test-model artifact
needed a unique name per sheet (`SpecTable_M05`), threaded through
`_write_spec_block` → `_create_spec_table` → `_set_sheet_scoped_names` and
generated by `test_model_sheets.spec_table_name` — five call sites across three
writers, for 47 sheets, all deleted.

**The regression that guards it (G14 / L10).** Every existing test-model case
derives its shell's profile from its own `source_table_ref`, so the two can
never disagree and the retarget path is never exercised.
`GuardStateCase.shell_profile_key` creates the disagreement deliberately: L10
builds its block with the Auto MPG profile and points `Source_Table` at
Life Expectancy. The case earns its sheet by where its evidence sits rather than
by the model it fits — `Schooling` contributes design columns from spec index
21, sheet row 25, ten rows past the old table's bottom edge — and a test pins
that property so an edit moving the predictors into the first 12 rows fails
instead of silently testing nothing.

**A stale name was swept on the way.** `Spec_Base_Period_Delta`
(`Regression!$I$4:$I$15989`) is residue from the rename to
`Spec_Sequence_Period` and is still in the shipped artifact:
`sync_workbook_names` only sweeps **workbook**-scoped residue, so a sheet-scoped
name outlives the code that created it indefinitely. `_RETIRED_LOCAL_NAMES` now
drops it on every build. Worth generalizing — a sheet-scoped name the writers
stopped creating has no other sweeper.

---

## v3.4+ — The spilled §4b zones are no longer grouped or collapsed

### Collapsing a zone that holds a spill stops the model recalculating

**Question:** the Regression sheet's §4b materialization band shipped with all
three content zones grouped and collapsed — `Model Context`, the
`Sample_Include` row mask, and the terminal Constructed Design Matrix. The band
is secondary reading surface and the terminal zone's width is one dropdown away
from hundreds of columns, so collapsing it read as free. Is it?

**RESOLVED — not for the two zones that hold spills.** `Sample_Include` and the
Constructed Design Matrix are full-height dynamic arrays, and a collapsed
outline group over a spill range is the configuration in which Excel fails to
recompute the model: the hidden columns keep the stale arrays, and every engine
that reads across them refits on stale values. The failure is silent in the
worst way — the numbers are all present, all plausible, and all computed from a
matrix that no longer matches the spec block. Both zones are now written
**ungrouped and expanded**, and `_write_materialization_zone` carries a
do-not-re-add note on the removed `Group()` / `ShowDetail` calls.

**Model Context keeps its group.** It is a fixed-height block of *individual
cells* — that is the v3.0 decision that made it cells rather than a `VSTACK`
spill, taken for the torn-context race — so there is no spill to leave stale
and hiding it costs nothing. It stays grouped as the label/value **pair** so
its labels never strand beside a collapsed value column, and it remains the
band's only collapsed zone.

**Accepted cost: the scrolling hazard is back**, and it is the exact cost the
original decision was buying off. An expanded terminal zone whose width follows
the design matrix means a wide model leaves a long ride to the right of the
sheet. That is a nuisance; refitting on stale values is a wrong answer. The
§4b ordering rule (nothing may ever be placed to the right of the design
matrix) is what keeps the expanded zone from displacing anything else, so the
cost stays contained to scrolling.

**What survived the removal.** The terminal zone still gets an explicit column
width across a bounded band — `_DESIGN_MATRIX_SIZED_COLUMNS`, sized to the soft
column threshold past which the width guard has already fired. That constant
and `_DESIGN_MATRIX_COLUMN_WIDTH` were named `_DESIGN_MATRIX_GROUPED_*` when
the band was also the outline group; they are renamed, because nothing about
them is grouped now.

**Supersedes** the v3.0 *Materialization zone layout* entry's "Collapse
behavior differs by zone" paragraph (which had `Model_Context` and
`Sample_Include` shipping expanded and the design matrix collapsed — the code
had since collapsed all three) and the v3.3 *Model Formula readout* entry's
accepted cost, which noted the caption on row 1 of the terminal zone was hidden
until the zone was expanded. It no longer is: the zone ships expanded, so the
caption is visible, and `Comparison_Model_Formula` reads the cell by name
either way.

**REJECTED — keep the group, ship it expanded.** An outline whose collapse
button is one click away from a silently wrong fit is a trap with a label on
it, not a feature. Removing the group removes the click.

## v3.x+ — Beta's grid search becomes a Full_Factorial spill; the Data-Table driver for the workbook split retires

**Question:** the v3.0 split ([§ Univariate becomes its own workbook](#univariate-becomes-its-own-workbook)) was forced by Excel's "Automatic except Data Tables" mode — Beta's two-stage grid search used two two-input Data Tables, and a combined workbook would either stale the Univariate fits or impose semiautomatic mode on every Regression user. The pending "grid shrink" entry ([§ The grid shrink ships as a later release](#the-grid-shrink-ships-as-a-later-release-of-the-univariate-artifact)) left "Beta's method-of-moments start and ~12×12 grid" open. With Beta now reworked onto a `Full_Factorial` dynamic-array spill, the Univariate artifact contains no Data Table at all. Does the split still earn its keep?

**Resolution:** the mechanism changed; the split question is left open. Beta's grid search is now two dynamic-array spills per stage — a `Full_Factorial(N, mins, maxs)` grid of N²×2 (`Alpha | Beta`) and a `BYROW` NLL column that reads it through the `#` operator — laid out two stages side by side in a 6-column zone (BY:CD), sized live by an in-sheet N cell (default N=10). Both artifacts therefore ship in full Automatic, and the original driver — "Automatic except Data Tables" is the only mode that can ship a workbook containing a Data Table — is retired. **Whether to re-merge the two workbooks is undecided and deliberately not settled here.** This entry records only that the reason the split originally existed no longer applies; anyone revisiting the question starts from a blank slate rather than from a rationale written to defend the status quo.

**Supersedes** the v3.0 *Univariate becomes its own workbook* entry's Data-Table rationale paragraph and closes the "Beta's method-of-moments start and ~12×12 grid are still open" tail of the *grid shrink* entry. The CLAUDE.md / CONTRIBUTING.md / README.md narratives describe the current mechanism and no longer argue for or against the split.

**Breakage class: MAJOR for the Univariate workbook version only.** The three fit zones relayout from 9/9/21 cols (stacked stages, Data Table) to 4/4/6 cols (stages side by side, vertical field-list control block, profile-NLL charts re-anchored above the body at rows 13–30, body rows 33+). The Beta Scale/Shape Min/Max/Step input cells the old Data Table exposed are replaced by α/β Min/Max/Step cells feeding `Full_Factorial`'s `VSTACK(mins)/VSTACK(maxs)`, with the NLL as its own `BYROW` column beside the grid, and the old `UV_BETA_S1`/`UV_BETA_S2` ranges are replaced by `UV_BETA_S{1,2}_{Alpha,Beta,NLL}` OFFSET ranges sized by the live N cell. A user's saved Beta bounds carry over in spirit but not in cell address. No Regression-side change.

## v3.x+ — the profile fits read their axis spill, so Grid Points is live in all three

**Question:** Beta's NLL column reads its `Full_Factorial` grid through the `#` operator, so
raising Beta's Grid Points resizes both spills together. The Weibull and Gamma profile fits
were left as they were: their axis is a live `Full_Factorial(N, Min, Max)` spill, but the NLL
beside it was 20 individual per-row formulas, each hard-pointing at its own axis cell
(`$BP$33` … `$BP$52`), and `UV_WB_S1` / `UV_GAMMA_S1` and their `_Axis` partners were fixed
A1 ranges. Raise Grid Points on a profile fit and the axis grew while the NLL column, the
recovery formulas, and the charts all stayed at 20. Should the profile fits get the same
treatment, or should Grid Points be frozen there instead?

**Resolution:** the same treatment. Each stage's NLL is now one spill —
`=LET(x,FILTER(UV_Data,UV_Include),BYROW($<axis>$#,LAMBDA(r,LET(p,INDEX(r,1,1),IFERROR(<NLL at p>,1E+15)))))`
— and all four names per zone become OFFSET ranges sized `MAX(IFERROR(<N cell>,1),1)`, the
un-squared form of what `UV_BETA_S*` already uses. Every fit in the artifact now has one
shape: a `Full_Factorial` spill beside a `BYROW` column that reads it, both following one live
cell. Three details in the new formula are load-bearing and should not be simplified away:

* **`INDEX(r,1,1)`** — `BYROW` hands the callback a 1×1 array, and passing that straight into
  `NLL_Weibull` or the profiled-out closed form broadcasts instead of evaluating.
* **`IFERROR` inside the `LAMBDA`** — outside it, one non-evaluable trial value collapses the
  whole column to `1E+15` instead of costing its own row.
* **`x` bound once per stage** — the per-row form paid a full-range `FILTER` for every point.

**Grid Points gets a floor of 2, applied three times.** The Step cell is `(Max-Min)/(N-1)`, so
N=1 is a `#DIV/0!` and a one-point grid searches nothing — `Full_Factorial` itself tolerates
N=1 via its `MAX(1,N-1)` divisor, but the sheet built around it does not. The floor lives in
`build_common.MIN_GRID_POINTS` and is enforced by `positive_grid_size` (the `--beta-grid-size`
argparse type on both build scripts), by a whole-number Stop `Validation` on each editable
Grid Points cell, and by a red conditional format on the same cells. Both in-sheet guards are
wanted: the Validation catches typing, and the CF catches a paste, which bypasses Validation
entirely.

**Accepted cost.** Number formats, the colour scale, and the border box are still painted over
the *default*-size window (`_PROFILE_BODY_CF_ROWS_CAP`, `_BETA_BODY_CF_ROWS_CAP`), so rows a
live N-increase adds beyond it are unshaded until the next build. Cosmetic — the named ranges
and every recovery formula track the real height.

---

## v3.3.y — Cook's screening, the Log domain, and the spec-block header band

Four changes to the Regression sheet, all requested by the repo owner.

### Cook's Distance screens on `F(0.5, p, n−p)`, not `4/n`

**Question:** the influence cutoff was `MIN(4/n, 0.9)` — two rules of thumb
collapsed into one lower bound. `4/n` is a function of the row count alone and
says nothing about the model; `0.9` is a bare constant. What should the sheet
actually compare Cook's D against?

**RESOLVED** — the median of the reference F distribution,
`F.INV(0.5, p, n−p)`, written once as `_COOKS_CUTOFF` in
`write_sheet_regression.py` and reused by the `AT`/`AY` conditional formats,
the `AY` flag column and the chart title, so the three cannot disagree about
what "flagged" means.

**The numerator df is `p`, not the ANOVA Regression df.** `$AB$15`
(`Regression_Degrees_Of_Freedom`) is `COLUMNS(X) − has_intercept`, i.e. `p−1`.
`Cooks_Distance` divides by `COLUMNS(Design_Matrix(X, Include))` — the
intercept column included — so `p` is what makes the statistic and its
reference distribution the same quantity, and `p + (n−p) = n` rather than
`n−1`. `$O$1` (the Σ Design Columns total) already holds exactly that, which
is why the cutoff reads it instead of the ANOVA table.

**One tier, not two.** The old pair graded amber at `4/n` and red at `0.9`.
`F(0.5, p, n−p)` lands between them for most models, so keeping either
alongside it would draw a line the cutoff itself does not recognize.

**The `IFERROR` is load-bearing.** Under `Zero_Predictors_Selected()` the
design collapses, `F.INV` sees a zero df and returns `#NUM!`. `NA()` makes
every comparison fail closed — nothing flagged — where a raw `#NUM!` would
propagate into the flag column and light the whole band.

### Two Log transforms — strict and drop-non-positive

**Question:** a `Log` transform on a column containing zeros or negatives kills
the fit. `Ln_Positive` returns `#N/A` on an included non-positive row and
`Sample_Include()` had no positivity term, so the `#N/A` propagated through
`Predictor_Columns()` into every statistic on the sheet. Life Expectancy's
`Schooling` has 28 true zeros and demonstrated it. This was recorded as an open
question in `_LN_ZERO_GUARD_NOTE` and in MODEL_TESTING_ASSETS § 1.2: should
`Sample_Include` grow a positivity term?

**RESOLVED — yes, but as a SECOND dropdown token rather than by changing what
`Log` means.**

* `Log` behaves exactly as before: the rows stay, `Ln_Positive` returns
  `#N/A`, the sample does not narrow. What is new is that it is no longer
  silent about it — the Transform cell turns red, and the `G2` status line
  names the variable, its count of non-positive rows *in the sample*, and the
  token that would exclude them.
* `Log (drop ≤ 0)` adds the positivity term to `Sample_Include` for its own
  columns only, and reports the excluded-row count in amber at `G2`.

**Why not simply make `Log` filter.** Dropping rows changes the sample being
fitted, which changes the model. A workbook that quietly did that to a spec the
user had already written would be answering a modeling question on their
behalf — the same objection that settled Intercept × Categorical and
Categorical × Log, and the same resolution: flag red and instruct, never
silently switch. It would also have contradicted `Ln_Positive`'s recorded
`NA()`-exception convention (`""` means "not in the sample", `#N/A` means "in
the sample and genuinely undefined"), which the two-token split preserves
exactly — under the filtering token those rows arrive with `include = 0` and
get `""`, which is the same contract, not an exception to it.

**Why this is not the ~10× axis-widener MODEL_TESTING_ASSETS § 2 warns about.**
Both tokens build the identical `Ln(x)` column and
`Constructed_Column_Transforms()` reports `"Log"` for both, so the
`(response_transform, predictor_transform)` unit-space dispatcher gains no new
combination and the Duan / back-transformation family needs no change at all.
The test cost is two cases, not a multiplier: L06 keeps the strict half
(unchanged behaviour, plus the new red flag) and L11 is the fittable filtering
half.

**`Sample_Include` gained an optional argument, not a twin.**
`Sample_Include(FALSE)` is the mask before the positivity layer; the difference
between the two populations is the excluded-row count. One predicate, evaluated
twice — the display and the constructor read the same closure, as
ARCHITECTURE § "display derives, never feeds" requires, and there is no second
copy to drift.

**The token string lives in two places that no import can bridge** — the Python
constant `_TRANSFORM_LOG_DROP` and the catalog bodies in
`lambda_functions.json`, which are JSON string data.
`test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform` pins
the two spellings together across all six readers, so a rename cannot
half-land. Writing that test is what caught `Model_Formula` still testing
`= "Log"` alone, which would have printed the raw response name instead of
`Ln(name)` under the new token.

### The rows 1–2 band above the spec block has one grammar

**Question:** the band had accreted. The Fixed Effects cardinality error sat at
`B1`, the Sequence error at `E1` (above Reference Level, a column with nothing
to do with sequencing), the width guard at `M2` (above Interaction Term, ditto)
relying on `N2:O2` as overflow, the FE readouts printed a literal `"n/a"` on
every non-panel model, and `_write_sequence_status` — which writes `H2`, the
right cell — had stopped being called at all, so the repo carried two copies of
the Sequence message and shipped the wrong one. Where does the next status go?

**RESOLVED — row 1 is labels, row 2 is the control/value/status, and every
status sits in the spec column it is about.** Role cardinality at `B2`, the Log
domain at `G2`, Sequence cardinality at `H2` (the dead writer revived, and now
the only copy), the spacing verdict at `I2`, the width guard at `O2` directly
under the Σ total it is a verdict on.

**Status cells carry no row-1 label.** They are blank whenever the spec is
legal, so a permanent label would caption nothing most of the time; the message
names its own subject when it appears. Only readouts are labelled.

**The accepted cost is runway.** One status per column means a long message has
nowhere to overflow. Every status cell is therefore `WrapText` with a short
imperative message and a hover Note carrying the full guidance, and row 2 is
left on automatic height — one line while the spec is legal, growing the moment
a message fires, which makes an error more prominent rather than less.

**Inactive readouts hide white-on-white** (`_hide_when`) instead of printing
`"n/a"` — the font-matches-fill idiom the spec rows already use for cascading
relevance, applied one band up. The Δ spectrum moved from `P1/Q1` + `P2` down
to `P3/Q3` + `P4`, aligning its header with the spec block's header row and its
body with the spec data rows; that also freed `P2:Q2` as the only overflow
runway anything on row 2 has, which is what `O2` uses.

**Known consequence:** `E`–`O` is a collapsed-by-default outline sub-group, so
the Log, Sequence, verdict and width-guard statuses are hidden until the user
expands it. That is not a regression (the old `E1`/`M2` placements were inside
the same group) and it is coherent under the new grammar: a status shares the
visibility of the control it is about, and a user who has set a Log token has
necessarily expanded the group to do so.

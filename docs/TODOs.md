# TODOs

Open work only. Nothing on this page is finished — if an item is here, it is
still to do.

Where the rest lives: resolved design decisions with their rationale are in
[DECISIONS.md](DECISIONS.md), foundational patterns in
[ARCHITECTURE.md](ARCHITECTURE.md), the version plan and the shipped-release
narratives in [ROADMAP.md](ROADMAP.md), the test-model suite each milestone has
to grow in [docs/MODEL_TESTING_ASSETS.md](MODEL_TESTING_ASSETS.md).
Completed items are removed from this file rather than accumulated here; each
milestone section below names the PRs that closed its shipped work, so the
detail is one `git show` away.

Every entry cross-links the DECISIONS entry that explains *why* — TODOs holds
only *what to do*.

## How to read an entry

Each pickable item carries a one-line tag: **status · size · Excel**.

| Field | Values |
|---|---|
| Status | `READY` — specified enough to start · `OPEN` — needs a decision first · `DEFERRED` — deliberately held, don't start without reopening the decision |
| Size | `S` — one small commit · `M` — a focused session · `L` — multi-session |
| Excel | `no Excel` — headless, CI-verifiable · `needs Excel` — cannot be finished or verified without Microsoft Office on a developer machine (see [CONTRIBUTING.md](../CONTRIBUTING.md) → *Verifying builds*) |

A `DEFERRED` item carries the status field **only**. It is not available to
pick up, so a size and an Excel flag would be invented rather than estimated —
whoever reopens the deferral scopes it then. Everything with a bare `DEFERRED`
tag is deliberate, not an entry someone forgot to finish tagging.

---

## Pick something to work on

### Start here

One self-contained item:

1. [Wire the calendar-dated monthly series](#test-model-suite) — one dataset config closes the single axis the coverage matrix still lists as uncovered. `M · no Excel`

### Ready now — no Excel required

| Task | Size | Milestone |
|---|---|---|
| [Wire the calendar-dated monthly series](#test-model-suite) | M | Test suite |
| [`Model_Formula_String` LAMBDA](#v34--model-comparison-sheet) | M | v3.4 |
| [`Cluster` Role — clustered-robust V_β](#v35--cluster-role-clustered-ses) | L | v3.5 |
| [Location & Scale transforms — `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`](#v39--standalone-data-transformation-library) | M | v3.9 |
| [Group & Panel transforms — `Zscore_By`, `Decompose_By`](#v39--standalone-data-transformation-library) | M | v3.9 |
| [Two-sample tests — `T_Test_OneSample`, `F_Test_Variance`, `Covariance_Matrix`](#v310--bivariate--two-sample) | L | v3.10 |
| [`Bootstrap_CI` / `MC_Percentile` / `PERT_Sample`](#v311--resampling--simulation) | L | v3.11 |
| [`Autocorrelation` / `Bartlett_Bands` / `Partial_Autocorrelation`](#v312--time-series-analysis-sheet) | M | v3.12 |
| [`Ljung_Box_Q` / `Box_Pierce_Q`](#v312--time-series-analysis-sheet) | S | v3.12 |
| [`ADF_Statistic` + critical-value table](#v312--time-series-analysis-sheet) | L | v3.12 |
| [`KPSS_Statistic` + critical-value table](#v312--time-series-analysis-sheet) | M | v3.12 |
| [`Classical_Decomposition`](#v312--time-series-analysis-sheet) | M | v3.12 |
| [`Moving_Average` / `Exponential_Smoothing` / MAE / MAPE](#v312--time-series-analysis-sheet) | M | v3.12 |

The catalog-function items above are headless by construction: a LAMBDA lands in
`lambda_functions.json` with a Python mirror and tests, and only the *sheet* that
surfaces it needs Excel.

Backlog and version-independent rows come first; the milestone rows below them are
in ladder order. The ladder sorts on two keys — all remaining Regression work
first, then, within that, how much the
[regression test-model suite](MODEL_TESTING_ASSETS.md) has to grow to cover
each milestone. That is why the three flat-cost non-Regression milestones (v3.10,
v3.11, v3.12) sit below four more expensive ones, and why the two v3.9 rows are
last in the Regression track — they are also listed in their own working order.

### Needs Excel on a developer machine

| Task | Size | Milestone |
|---|---|---|
| [Re-examine the intercept-only closed-form bypass](#v30-leftovers) | M | v3.0 |
| [Diagnostic-chart reference lines](#v1x--regression-sheet) | M | v1.x |
| [Suppress worst-fit distributions from the combo charts](#v11-leftovers--univariate-sheet-writer) | M | v1.1 |
| [Model Comparison sheet layout](#v34--model-comparison-sheet) | L | v3.4 |
| [Two-sample sheet layout](#v310--bivariate--two-sample) | M | v3.10 |
| [Simulation sheet layout](#v311--resampling--simulation) | M | v3.11 |
| [Time-series sheet (`write_sheet_time_series.py`)](#v312--time-series-analysis-sheet) | L | v3.12 |
| [Fix `warn_if_workbook_open`'s buffered-prompt deadlock](#build-tooling--found-by-the-first-real-poe-verify-run) | S | — |
| [Stop leaking Excel instances under `parallel`](#build-tooling--found-by-the-first-real-poe-verify-run) | S | — |
| [Clean the four `#VALUE!` cells out of the Regression artifact](#build-tooling--found-by-the-first-real-poe-verify-run) | S | — |

### Blocked on a decision

| Question | Milestone |
|---|---|
| [Blank-categorical caveat in `Sample_Include()`](#v20-leftovers) | v2.0 |
| [Mismatched-predictor-set fallback for the Comparison sheet](#v34--model-comparison-sheet) | v3.4 |
| [Can a column be both `Sequence` and `Time`](#v36--time-role--lagdifference-semantics) | v3.6 |
| [Two-sample selector — 3-way flag or separate `paired` boolean](#v310--bivariate--two-sample) | v3.10 |
| [Does the Time Series sheet declare its own series, or read a `Time` Role](#v312--time-series-analysis-sheet) | v3.12 |
| [Is the differenced series a materialized column or a closure](#v312--time-series-analysis-sheet) | v3.12 |

Deliberately held, not available to pick up: [BFN critical
values](#v21-leftovers--follow-on-polish), [Categorical × FE prediction
encoding](#v21-leftovers--follow-on-polish).

---

## v1.x — Regression sheet

- **READY · M · needs Excel** — Add reference lines to the Cook's Distance,
  PRESS Residuals, and Leverage vs. Studentized charts. Format thematically
  similar to the conditional formatting in the table (yellow = mild, red =
  strong). Use minimalist helper columns (2 anchor points using max/min of the
  relevant threshold; place them beneath the chart area). Same
  identity-line-via-real-data-series approach as the v1.2 diagnostic chart
  pattern — see
  [DECISIONS.md § v1.2 identity-line](DECISIONS.md#v12--workbook-hardening).
  Cook's now has ONE threshold, `_COOKS_CUTOFF` (`F.INV(0.5, p, n-p)`), so its
  reference line is a single series rather than the yellow/red pair the other
  two charts want.

## v1.1 leftovers — Univariate sheet writer

Shipped 2026-06-29; see [ROADMAP.md](ROADMAP.md#v11--univariate--shipped-2026-06-29).

- **READY · M · needs Excel** — Investigate suppressing worst-fit / N/A-error
  distributions from the combo charts. Best outcome would be dynamically hiding
  those columns — hidden columns drop out of charts automatically
  (`PlotVisibleOnly` default) — but it is unclear whether column hiding can be
  driven from cell values without VBA (manual hiding works; data-driven hiding
  may be VBA-only, which the library forbids). No-VBA fallback to evaluate: emit
  `NA()` across a suppressed distribution's column, since line charts skip
  `#N/A` points — same chart effect without hiding. See
  [DECISIONS.md § v1.1 histogram overlays](DECISIONS.md#v11--univariate) for the
  combo-chart design.

- **READY · L · no Excel** — Add support for more distribution families:
  Bernoulli, Binomial, Geometric, Negative Binomial, Hypergeometric, Poisson,
  Uniform, Chi-Square, Student-t. (The engine half is headless; surfacing each
  new fit on the sheet needs Excel.)

## v2.0 leftovers

Specification-Driven Regression shipped 2026-07-05; the human test plan
(T0–T16) was executed, signed off PASS, and retired — its cases live on in
`tests/test_analyze_model_construction.py` and
`tests/test_difference_by_verification.py`. What remains is one open decision.

- **OPEN · M · needs Excel** — Resolve the blank-categorical caveat.
  `Sample_Include()`'s role-aware completeness layer requires numeric Response
  and numeric included Continuous Predictors, but Categorical Predictors impose
  no non-blank condition; a blank category value encodes as all-zero dummies
  (indistinguishable from the reference level). Verify against a live build —
  declare a Categorical Predictor whose column has a blank value on at least one
  otherwise-complete row, confirm the row is included and encodes as the
  reference level, and check whether the fitted coefficients shift — then record
  the decision: accept as documented behavior, or extend `Sample_Include()` with
  a non-blank condition for included Categorical Predictors. Interim workaround:
  a completeness column declared as a Filter. See
  [DECISIONS.md § v2.0 auto-completeness](DECISIONS.md#v20--specification-driven-regression).

## v2.1 leftovers — follow-on polish

The Sequence axis, gap-aware longitudinal layer, serial-correlation diagnostics,
and one-way Fixed Effects all shipped inside the 3.0.0 artifact (TODOs #1–#10,
verified at 0 mismatches across all 12 spec-driven QC cases). Design rationale:
[DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).
Two-way FE is deliberately deferred until this framework is finished — see
[v3.8](#v38--two-way-fixed-effects).

- **DEFERRED** — **BFN critical values**. N,T-dependent bounds per Bhargava et
  al. 1982 tables; do NOT present standard DW bounds next to the BFN cell. The
  deferral record (why N,T-dependent, why not standard DW) is in
  [DECISIONS.md § v2.1 BFN critical values](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

- **DEFERRED** — **Categorical × FE prediction encoding**. x_new and x̄ᵢ formed
  in constructed design-matrix space; UI wire to encode through `Dummy_Code`
  before reaching the FE formula. Largely subsumed by v2.0 categorical
  prediction; recorded so the encoding step is not forgotten. The deferral record
  is in
  [DECISIONS.md § v2.1 Categorical × FE](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

## v3.0 leftovers

The engine-interface release shipped 2026-08-02 in three stages plus the
Univariate split (later reunified): stage 1 constructor pipeline + intercept
relocation (#148), stage 2 `Model_Context` (#150), stage 3 layout (#152), plus
polish (#153, #154). Every stage cleared the spec-driven Excel gate at 0
mismatches across all 12 QC cases. Narrative in
[ROADMAP.md](ROADMAP.md#v30--the-engine-interface-release--shipped-2026-08-02);
rationale in
[DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).

- **READY · M · needs Excel** — Re-examine the intercept-only closed-form bypass
  in `write_sheet_regression.py` → `_setup_local_names` (`Intercept_Only_N` /
  `_Point` / `_SE` / `_S` / `_DF`). `Design_Columns()` now returns a well-formed
  ones column in the zero-predictor state, so the engines could fit it directly
  and the bypass may be removable. Kept for stage 1 because the shipped
  behaviour was verified against it; retire it only with a QC pass behind it.

## v3.2 — Full materialization of the design matrix

The stage-3 follow-on. The terminal Constructed Design Matrix zone, its collapse
behaviour, and its width guard shipped at v3.0; the spills that fill it — plus
`Sample_Include()`'s own zone — shipped at #161, each headed on
`_MATERIALIZATION_HEADER_ROW` and spilling from `_MATERIALIZATION_SPILL_ROW`,
full height and row-aligned with the source table. See
[ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone). What remains
is rewiring the readers, which is where the performance win actually is.

The rewiring proceeded zone by zone against the cell-by-cell spec verifier — a
name resolving to the wrong range returns numbers from the wrong rows rather
than erroring, so a sweep is unsafe. The `#`-inside-a-`LAMBDA`-`RefersTo`
combination the readers use (`Fit_Design_Columns` / `Fit_Sample_Include`,
registered by `_add_spill_reader` in `regression_materialization.py`) is
Excel-verified — settled on #223's spike — so the rest were mechanical
retargets, each landed Excel-verified.

Migrated:
- **Regression Statistics** (col AB, rows 4-8) — #224.
- **Diagnostics** (col AE, rows 4-10) — PR 3 of N.
- **ANOVA, Coefficients, Smearing / Unit-Space, Prediction Interval / FE-group,
  Residual Output (cols AO-AX), Predictor Summary (row 3), the
  serial-correlation trigger cells (rows 11-12), and the `n` / `mean_y` / `sd_y`
  named ranges** — the remaining zones, migrated together and Excel-verified.
  Every engine call site on the sheet now reads `Fit_Design_Columns()` /
  `Fit_Sample_Include()`; no bare `Design_Columns()` / `Sample_Include()`
  constructor call survives outside the spill-source cells that PRODUCE the
  arrays. Those cells now call the `*_Calc` computational leaves
  (`=Sample_Include_Calc()` / `=Design_Columns_Calc()`) rather than the public
  names, which is what lets the public names themselves be readers over the
  spills (see the promotion note below) without the producing cell becoming
  self-referential.

- **DONE · L · needs Excel** — Point the readers at the spills. Complete: the
  recalculate time fell from the 21.8s spike baseline to 10.5s, and the
  cell-by-cell spec verifier (NumPy) passes across every migrated zone.

The catalog LAMBDA bodies that *called* the constructors are rewired too:
`Response_Column`, `Predictor_Columns`, `Constructed_Column_Names`,
`Constructed_Column_Transforms`, `Absorbed_Degrees_Of_Freedom`,
`Design_Response`, `Design_Columns`, and `Log_Domain_Status` now call
`Fit_Sample_Include()` for the default mask. This is cycle-safe because the
computational leaves `Sample_Include_Calc` / `Design_Columns_Calc` are leaves —
their bodies depend only on `Source_Data` and the spec arrays, so the spill a
leaf produces can never depend back on a caller. No JSON lambda calls
`Design_Columns()` (the engines receive the matrix as an argument), so there
were no `Design_Columns()` swaps to make in the catalog. `Sample_Include(FALSE)`
(the pre-positivity mask `Log_Domain_Status` differences against the default)
now delegates to `Sample_Include_Calc(FALSE)` — the materialized spill is the
default mask only, and `FALSE` expresses an argument it cannot.

- **DONE · M · needs Excel** — Promote the `Sample_Include` / `Design_Columns`
  *names* from live closures to readers over their materialized spills. The
  self-reference that had kept this deferred (the spill cell WAS
  `=Sample_Include()`, so pointing the name at its own spill made the producing
  cell self-referential) is broken by a `_Calc` split: the existing REDUCE
  bodies move verbatim into private `Sample_Include_Calc` /
  `Design_Columns_Calc` leaves, the spill-source cells call those leaves, and
  the public `Sample_Include` / `Design_Columns` names become thin readers
  (`=LAMBDA([apply_log_domain],LET(use_log,IF(ISOMITTED(apply_log_domain),TRUE,apply_log_domain),IF(use_log,Fit_Sample_Include(),Sample_Include_Calc(FALSE))))`
  and `=LAMBDA(Fit_Design_Columns())`). Cosmetic, not a perf step — every call
  site already read the spills, so the win was banked at the rewiring; this
  just makes the names say what they do. Side-benefit: the spec-block K
  (Levels) and O (Design Columns audit) computed columns call bare
  `Sample_Include()`, so they now resolve through the reader to the spill and
  stop recomputing the REDUCE. `Log_Domain_Status`'s `Sample_Include(FALSE)`
  is unchanged in effect (it recomputes via the `_Calc` leaf — the
  pre-positivity mask is deliberately NOT materialized). See
  [DECISIONS.md § materialization in two steps](DECISIONS.md#materialization-lands-in-two-steps--model_context-now-sample_include-deferred).

## v3.4 — Model Comparison Sheet

Planned as v2.3; moved after v3.0 when the feature train was resequenced — see
[ROADMAP.md](ROADMAP.md#v34--model-comparison-sheet--planned). Comes after v3.3
because an R² computed on `Ln(y)` is not comparable with one computed on raw `y`.

**Test assets — additive (~1×).** No new data: M1, L2, and P2 already supply ≥3
registered models with shared prediction inputs. Add one mismatched-predictor-set
pair (M1 vs M14) for the `XLOOKUP [if_not_found]` question below. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **READY · M · no Excel** — Implement the `Model_Formula_String` LAMBDA with
  header-signature validation (`NA()` on non-Regression targets). The name
  resolution and the anchor-cell argument-type rationale are in
  [DECISIONS.md § v2.3 Model_Formula_String](DECISIONS.md#v23--model-comparison-sheet).

- **READY · L · needs Excel** — Sheet layout: model registry (hyperlinks), GoF
  table referencing the v3.3 unit-space headline cells, shared prediction inputs
  (Comparison sheet is the source; Regression sheets pull via XLOOKUP),
  prediction results table. The data-flow direction
  (Comparison-as-source-via-XLOOKUP) is in
  [DECISIONS.md § v2.3 prediction inputs](DECISIONS.md#v23--model-comparison-sheet).

- **OPEN · S · no Excel** — Decide the mismatched-predictor-set fallback
  (XLOOKUP `[if_not_found]`). See the open-decision note in
  [DECISIONS.md § v2.3 Model Comparison Sheet](DECISIONS.md#v23--model-comparison-sheet).

## v3.5 — `Cluster` Role (clustered SEs)

Planned as v2.7+, promoted out of the unordered bucket by the ladder reordering —
it is Regression work, and a variance-estimator variant over a few existing models
is the cheapest thing left to cover in that track. See
[ROADMAP.md](ROADMAP.md#v35--cluster-role-clustered-robust-ses--planned).

**Test assets — near-additive, no new data to start.** Production Lots' three
facilities are the initial within-group-correlated fixture, and deliberately few —
three clusters is what exercises the small-cluster warning path. `Grunfeld` arrives
with [v3.8](#v38--two-way-fixed-effects) and supplies 10–11 proper clusters. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **READY · L · no Excel** — Implement the `Cluster` Role (at most one) —
  clustered-robust variance estimator. Has partial forward wiring from
  `Serial_Correlation_Group()`'s dormant Cluster branch (PR #106), so the
  resolver side is partial; the engine side (cluster-robust V_β) is not.

- **READY · S · needs Excel** — Lift the v2.1 `n/a — engine forthcoming` token on
  the BFN cell when Cluster is active (the BFN formula already uses
  `Serial_Correlation_Group()` as its resolver, so the wiring is partial).

## v3.6 — `Time` Role + lag/difference semantics

Planned as v2.7+, promoted out of the unordered bucket by the ladder reordering.
See [ROADMAP.md](ROADMAP.md#v36--time-role--lagdifference-semantics--planned).

**The sheet half moved to [v3.12](#v312--time-series-analysis-sheet).** This
milestone was carrying both an engine change and a new worksheet;
`Moving_Average`, `Exponential_Smoothing` and `write_sheet_time_series.py` now
live there. What remains here is Regression-track work plus the dataset.

**Test assets — near-additive, and the one item that closes a coverage gap
existing today.** Wiring a calendar-dated monthly series (~144 rows,
AirPassengers-shaped, with a real date column) makes the Sequence
**calendar-signature verdict** testable — the single uncovered axis in the
Section-1 coverage matrix, because no wired dataset carries real dates. That test
can be written as soon as the dataset lands, ahead of the Role itself. See
[docs/MODEL_TESTING_ASSETS.md § 1.5](MODEL_TESTING_ASSETS.md#15-coverage-matrix)
and [§ 3](MODEL_TESTING_ASSETS.md#section-3--supplemental-datasets-kept-minimal).

- **OPEN · L · no Excel** — Design and implement the `Time` Role. Partially
  forward-wired via the v2.1 Sequence axis, but the full `Time` Role adds
  time-index semantics (for the v3.12 time-series sheet, for cross-sheet
  `Lag_By` / `Difference_By` calls). The open question: can a column be both
  `Sequence` and `Time`, or are they mutually exclusive?

## v3.7 — `Weight` Role (WLS)

Planned as v2.6 and claimed as v3.7; it keeps that number, but now as the first
~2× item in the Regression track — the ladder reordering put `Cluster` and `Time`
ahead of it and Two-sample and Resampling behind it. See
[ROADMAP.md](ROADMAP.md#v37--weight-role-wls--planned). The standalone WLS milestone and its
`[weights]`-argument-vs-parallel-function-set debate are superseded by a
**`Weight` value on the Role axis** (see
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)).
Three-stage scope carried forward: user-supplied weights →
variance-driver-derived weights → FGLS. This milestone ships the first stage
only. The default-uniform → OLS pattern (the "non-breaking MINOR" guarantee) is
in [DECISIONS.md § v2.6 WLS](DECISIONS.md#v26--wls-weight-role-default-uniform-weights-argument).

- **READY · M · needs Excel** — Implement the `Weight` Role (at most one, per the
  cardinality rule that Response, Time, and Weight share; status-block validation
  identical to exactly-one-Response).

- **READY · L · no Excel** — Thread weights through the engine per the Role-axis
  design: a single optional `[Weights]` argument (default uniform) on the
  inferential chain. Default-uniform means every existing OLS call computes
  identically — the v2.1 `[DF_Absorbed]` precedent (default 0 → identical no-FE
  model) is the exact pattern to follow.

- **READY · S · needs Excel** — Update the Diagnostic Guide to describe which
  diagnostics change interpretation under WLS. (WLS closes the loop opened by
  v1's Scale-Location diagnostic.) Static template sheet — regenerate via
  `rebuild_static_sheets.py`.

**Test assets — ~2× over a representative subset.** A dataset with a natural
weight column (R/MASS `Insurance`, 64 rows, or a grouped-mean aggregation of an
existing one). Plan **weighted variants of ~6 representative models — one per
dispatch-pair family — not the whole suite**; that bound is what keeps this at ~2×.
Assert the recorded trap as an oracle: `DEVSQ(√w ⊙ y)` ≠ weighted SST. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

## v3.8 — Two-way Fixed Effects

Planned as v2.7+, promoted out of the unordered bucket by the ladder reordering.
See [ROADMAP.md](ROADMAP.md#v38--two-way-fixed-effects--planned).

**Test assets — ~2× over the FE family.** R `Grunfeld` (200 rows, 10 firms × 20
years) plus an **unbalanced variant** with rows deleted, to exercise
`Is_Balanced_Panel` and the convergence check; the existing FE family is re-run
two-way. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **READY · L · no Excel** — Implement `Absorb_Two_Way_Fixed_Effects(x, group1,
  group2, [include], [passes])` (alternating-projection demeaning for unbalanced
  panels).

- **READY · M · no Excel** — Implement `Demean_Two_Way_Balanced(x, group1,
  group2, [include])` and the two-way `Is_Balanced_Panel` check.

- **READY · M · no Excel** — Implement `Fixed_Effects_Convergence_Check(x,
  group1, group2, [include])`; surface in the status block whenever two FE
  variables are active.

- **OPEN · M · needs Excel** — Lift the v2.1 one-FE-variable status-block error;
  resolve the two-way prediction question (group intercepts are not recoverable
  as simple group means). The one-way-scope rationale is in
  [DECISIONS.md § v2.1 scope](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

## v3.9 — Standalone Data Transformation library

Planned as the second half of v2.2, then carried as the v3.3 remainder; moved to
the end of the **Regression track** by the ladder reordering — see
[ROADMAP.md](ROADMAP.md#v39--standalone-data-transformation-library--partially-delivered).
Full specs in
[ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).

**Shipped so far — the additive helpers (#234/#235):** `Numeric_Complete_Cases`,
`Dummy_Column`, `Interact`, `Model_Matrix`, each a standalone catalog LAMBDA with
a pure-Python mirror (catalog 148 → 151); the construction trio also carries a
workbook-backed Excel COM test (`tests/test_categorical_model_construction_excel.py`,
gated on `RUN_EXCEL_INTEGRATION=1`). They widen no axis.

**Test assets — the ~10× axis-widener, and no new data.** Every new `Transform`
value widens the predictor-transform axis that today holds {None, Log}, and each
widening multiplies the response × predictor dispatch table. **Work the two items
below in the order given** — predictor-side Location & Scale transforms before the
Group & Panel pair — and hold any response-side extension for last, because a response
transform also multiplies the back-transformation / unit-space semantics. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **READY · M · no Excel** — Location & Scale: `Center`, `Zscore`,
  `Minmax_Scale`, `Winsorize`. (`Ln_Positive` shipped early, alongside the
  Transform column-G wiring, rather than waiting for the rest of this bundle.)

- **READY · M · no Excel** — Group & Panel: `Zscore_By`, `Decompose_By`.
  (`Demean_By` / `Group_Mean` shipped at v2.1; the two-way functions follow the
  two-way FE milestone.)

## v3.10 — Bivariate / Two-sample

Planned as v2.5, claimed as v3.6, briefly held at v3.5. It sits here because the
ladder reordering ships all remaining Regression work first — this is the first
milestone that opens a **new analysis surface** rather than extending the
Regression sheet. It still precedes Resampling: both are flat-cost to test, and
two-sample is the parity gap a user hits first. See
[ROADMAP.md](ROADMAP.md#v310--bivariate--two-sample--planned).

Nothing here got harder by waiting — this milestone depends on no Regression
milestone and none depends on it.

**Test assets — additive.** Two small new datasets: a two-group set (R
`ToothGrowth`, 60 rows, or the in-repo `Status` split of Life Expectancy) and a
**paired** set (R `sleep`, 20 rows). Cases: equal-variance t, Welch t, paired t,
and the F-test of variances feeding the selector cell. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order)
and [§ 3](MODEL_TESTING_ASSETS.md#section-3--supplemental-datasets-kept-minimal)
for the wiring cost of a new dataset.

- **READY · M · no Excel** — Implement `T_Test_OneSample(data, mu0, alpha,
  [include])` → test statistic, p-value, CI.

- **OPEN · M · no Excel** — Implement `T_Test_TwoSample(data1, data2, alpha,
  equal_var, [include1], [include2])` — equal-variance, Welch unequal-variance,
  and paired variants. The open question: paired is a separate code path the
  `equal_var` flag does not cover — 3-way flag or separate `paired` boolean? See
  [DECISIONS.md § v2.5 two-sample selector](DECISIONS.md#v25--claimed).

- **READY · M · no Excel** — Implement `F_Test_Variance(data1, data2, alpha,
  [include1], [include2])` — output feeds a recommendation cell that selects the
  appropriate t-test variant.

- **READY · S · no Excel** — Implement `Covariance_Matrix(data, [include])` —
  sample covariance (consistent with the existing catalog's sample-statistic
  convention); complement to the existing `Correlation_Matrix`.

- **READY · M · needs Excel** — Design the two-sample sheet layout: inputs, test
  selector, F-test assumption check, output (test statistic, df, p-value, CI,
  effect size). Implement `write_sheet_two_sample.py`.

## v3.11 — Resampling & Simulation

Planned as v2.4, claimed as v3.5, briefly held at v3.6. The second non-Regression
milestone, behind Two-sample — see
[ROADMAP.md](ROADMAP.md#v311--resampling--simulation--planned).

**Test assets — additive, no new data.** The seeded pre-drawn
`Bootstrap_Random_Draws` table *is* the asset; Production Lots (n = 51) is the
natural small-n bootstrap target (slope CI on the `production_lots_log_no_fe`
case), and PERT/MC cases need only parameter cells. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **READY · M · no Excel** — **The pre-drawn random table** (design RESOLVED, in
  answer to the no-volatile constraint). A single sheet-scoped named range
  `Bootstrap_Random_Draws` holds a uniformly-distributed random table pre-drawn
  once at build time, seeded from a SHA-256 digest of the source CSV.
  `Bootstrap_CI` indexes via
  `INDEX(Bootstrap_Random_Draws, MOD(SEQUENCE(n_resamples), ROWS(Bootstrap_Random_Draws))+1)`.
  Same inputs → same output, every recalc. `RANDARRAY()` rejected. Full rationale
  (auditability vs. fresh randomness) in
  [DECISIONS.md § v2.4 no-volatile constraint](DECISIONS.md#v24--resampling--simulation).
  To get a new draw, regenerate the workbook via `build_production.py`
  (deliberate, not a limitation).

- **READY · L · no Excel** — Implement `Bootstrap_CI(data, stat_lambda,
  n_resamples, alpha, [include])` — bootstrap confidence interval for an
  arbitrary statistic passed as a LAMBDA. Uses the pre-drawn table above.

- **READY · M · no Excel** — Implement `MC_Percentile(dist_params, n_samples,
  percentile)` — Monte Carlo draw from a fitted distribution; complements the
  Univariate fitting. Uses the same pre-drawn table.

- **READY · M · no Excel** — Implement `PERT_Sample(min, mode, max, n_samples)` —
  BetaPERT sampling for cost/schedule risk analysis. Uses the same pre-drawn
  table.

- **READY · M · needs Excel** — Design the sheet layout (bootstrap section +
  Monte Carlo section; may share one sheet). Implement
  `write_sheet_simulation.py`.

## v3.12 — Time Series Analysis sheet

The third and last new analysis surface, and it absorbs the sheet half of v3.6.
See [ROADMAP.md](ROADMAP.md#v312--time-series-analysis-sheet--planned).

**Test assets — additive, no new data.** The calendar-dated monthly series v3.6
wires is the only dataset needed: in levels it is the non-stationary seasonal
case, log-differenced at `d=1, D=1, m=12` the stationary counterpart, so one
series gives both verdicts of every test. Oracles are `statsmodels.tsa`
(`acf`, `pacf`, `acorr_ljungbox`, `adfuller`, `kpss`, `seasonal_decompose`)
through a `TimeSeriesSpecCase` registry mirroring `RegressionSpecCase`. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 9](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **BLOCKED · — · no Excel** — Decide whether the sheet declares its own series
  or consumes a Regression sheet's `Time` Role. Leaning independent; cross-sheet
  reading is the v3.4 Comparison sheet's job. Everything else here waits on the
  answer only for where the series comes from, not for what it is.

- **BLOCKED · — · no Excel** — Decide whether the differenced series is a
  materialized column or a per-consumer constructor closure. The ARCHITECTURE
  §4b lesson (one spill everything reads) argues for the column; ACF, PACF, both
  tests and the decomposition all want the identical vector.

- **READY · M · no Excel** — Implement `Autocorrelation(data, max_lag,
  [include])` returning lag / r / SE, and `Bartlett_Bands` for the widening
  confidence band.

- **READY · M · no Excel** — Implement `Partial_Autocorrelation` via the
  Durbin–Levinson recursion — **not** a Yule–Walker solve, which would pull in
  `Gram_Inverse` and a singularity path for no gain.

- **READY · S · no Excel** — Implement `Ljung_Box_Q(acf, n, h, [df_fitted])` and
  `Box_Pierce_Q`. The `df_fitted` argument is what makes the same function
  correct on fitted residuals (`h − p − q`) as on a raw series (`h`).

- **READY · L · no Excel** — Implement `ADF_Statistic(data, lags, spec)` as a
  regression of Δy on y(t−1) plus lagged differences, routed through the
  existing OLS engine, with `ADF_Critical_Value(n, spec, alpha)` as a lookup
  table (MacKinnon). Non-standard null distribution — the table is the point.

- **READY · M · no Excel** — Implement `KPSS_Statistic(data, lags, spec)` and
  `KPSS_Critical_Value(n, spec, alpha)`. Ships with ADF, not after it: opposite
  nulls, and the four reject/fail-to-reject combinations are the actual reading.

- **READY · M · no Excel** — Implement `Classical_Decomposition` (additive and
  multiplicative) returning trend / seasonal / remainder as one spill; trend by
  centered `Moving_Average` at period `m`, seasonal indices as de-trended period
  means.

- **READY · M · no Excel** — Implement `Moving_Average(data, window, [include])`.
  *Moved from v3.6.*

- **READY · M · no Excel** — Implement `Exponential_Smoothing(data,
  alpha_smooth, [include])` — note: use `alpha_smooth` to distinguish from the
  significance-level `alpha`. *Moved from v3.6.*

- **READY · S · no Excel** — Implement `MAE` / `MAPE` forecast-error metrics.
  `RMSE` already exists in unit-space form (`Unit_Space_RMSE`); check whether it
  generalizes before adding a third.

- **READY · L · needs Excel** — Implement `write_sheet_time_series.py`: the
  series spec block (Time / Value / Group / `Log` / `d` / `D` / `m`), the
  ACF and PACF tables and charts with bands as **real data series, never
  shapes**, the test blocks showing statistic + critical value + verdict side by
  side, the decomposition block, and the forecast block with an
  actual-vs-smoothed chart. Register the `TSChart`-prefixed sheet-scoped named
  ranges in the sheet's own `_setup_local_names`, following the `RegChart`
  precedent. *Moved from v3.6.*

- **OPEN · S · no Excel** — Decide whether `Ljung_Box_Q` also belongs on the
  Regression sheet's diagnostic band beside `Durbin_Watson`. DW tests lag 1
  only; a portmanteau over h lags is strictly more informative, and `df_fitted`
  already exists to make it correct there.

## v3.13+ — Unordered candidates (no claim)

What is left after the ladder reordering gave the other candidates numbers, and
after the Time Series sheet took v3.12. Nothing about their test cost sequences
them: ANOVA-as-regression needs only `warpbreaks` plus the existing categorical
machinery, and the other two are design-not-started. A user pressing for one of
these would reorder it; absent that signal, a single maintainer should not
pre-order work that may not be the next thing actually needed.

### Multi-group means (ANOVA)

- **READY · L · no Excel** — Implement one-way ANOVA as regression on group
  dummies, reusing the existing SS/MS/F machinery. Frame explicitly as "ANOVA is
  regression" — group means, SS decomposition, and F-test should match the MLR
  output exactly.

- **READY · M · no Excel** — Add post-hoc comparisons (Tukey HSD or Bonferroni)
  as an optional output section.

### Long-tail (out of planning horizon)

- **Fourier analysis** — the *ToolPak Parity Reference* notes it is
  "intentionally skipped" and a later addition-by-demand decision, not a planned
  milestone.
- **Decision analysis** — loss functions, cost/risk oriented. Not on the planning
  horizon.

---

## Test-model suite

Version-independent; the plan of record is
[docs/MODEL_TESTING_ASSETS.md](MODEL_TESTING_ASSETS.md). Every model in Section 1
now has an oracle: 35 fittable `RegressionSpecCase` entries in
`lambda_catalog/analyze_regression_spec.py` (two of them heavy — built only under
`--include-heavy`) and 17 `GuardStateCase` entries in
`lambda_catalog/analyze_regression_guard_states.py`, each pinned in
`_EXPECTED_CASE_NAMES` / `_EXPECTED_GUARD_NAMES` and materialized as a worksheet.
The covering-array rule is that a case earns its place by covering something no
other case does — so **the way to find work here is the
[§ 1.5 coverage matrix](MODEL_TESTING_ASSETS.md#15-coverage-matrix), not this
page**. A corner with no case named against it is the open work; there is
currently one.

- **READY · M · no Excel** — Wire a **calendar-dated monthly series**
  (~144 rows, AirPassengers-shaped, with a real date column). This closes the one
  axis § 1.5 lists as uncovered: the Sequence **calendar-signature verdict**
  (~28–31 / ~90–92 / ~365–366-day spacing clusters), untestable today because no
  wired dataset carries real dates. One `CsvDatasetConfig` + one
  `SpecDatasetProfile` + a registry entry. It is pulled out of the DEFERRED entry
  below because it is the only one of those datasets that closes a gap existing
  *now* rather than arriving with a milestone; the case can be written as soon as
  the data lands, ahead of the [v3.6](#v36--time-role--lagdifference-semantics) `Time` Role
  it also serves.

- **DEFERRED** — Wire the remaining supplemental datasets (`warpbreaks`,
  `Grunfeld`, `Insurance`, `sleep`, `ToothGrowth`). Same one-config-plus-profile
  shape as above, but each lands with the milestone that needs it — see
  [§ 3 Timing](MODEL_TESTING_ASSETS.md#timing).

## Build tooling — found by the first real `poe verify` run

Version-independent; not tied to a milestone. These came out of the first
developer-machine `poe verify` after the concurrency change, on 2026-08-06 —
the run [CONTRIBUTING.md](../CONTRIBUTING.md) asks for and no CI can perform.
The transcripts are in [excel-only-runs/](../excel-only-runs).

The concurrency itself worked: multiple Excel instances built the workbook
side by side for ~84 minutes with no contention over `templates/static_sheets.xlsx`,
and the verifiers passed (Univariate `Verify: passed`; test-models 48/48 `ok`).
What follows is what the run exposed around it.

- **READY · S · needs Excel** — **`warn_if_workbook_open` deadlocks under
  `poe verify`.** Its prompt is an `input()` call, but the warning above it goes
  to **stderr with `flush=True`** while the prompt goes to **stdout**, which
  `output_mode = "buffer"` holds until the task ends. So a locked workbook
  prints a warning, then blocks forever on a question the user cannot see; the
  prompt text only appeared when Ctrl+C flushed the buffer. It reads as a hang
  with no Excel process consuming CPU. The function already returns early when
  `not sys.stdin.isatty()`; it needs the same treatment when **stdout is not a
  live terminal**, letting the reactive `_retry_on_open` catch a genuine lock at
  save time instead. Verifiable without Excel — the probe is injectable, and
  `tests/test_build_common.py` already drives the prompt loop with a stub.

- **READY · S · needs Excel** — **Excel instances leak under `parallel`.**
  Three `EXCEL.EXE` processes survived the run at 0% CPU, after the two
  completed drivers should have quit theirs. `_quit_app_quietly` is a bare
  `try/except: pass`, so a failed quit is invisible. The cost is not the idle
  process: an orphan can hold the workbook and leave a `~$Lambda_Library.xlsx`
  sidecar, which is what the *next* run's pre-flight probe trips on — plausibly
  how this run acquired the stale lock that triggered the item above. At minimum
  the swallowed exception should be reported; better, the quit should be
  verified.

- **READY · S · needs Excel** — **Four `#VALUE!` cells ship in the Regression
  artifact.** `Mileage Data` J159, K159, J355, K355 hold literal `#VALUE!`
  cached values. They are copied faithfully from
  `sample_data/auto_mpg_data.csv`, which carries pre-split `Make` / `Model?`
  columns: the two rows whose `Car Name` is `subaru` have no space to split on,
  so the spreadsheet that produced the CSV wrote `#VALUE!` into both. The
  documented exception in `tests/test_workbook_invariants.py` covers the `#N/A`
  that `Difference_By` / `Lag_By` legitimately return at gap rows — this is
  neither of those. Fix the two CSV rows, rebuild, commit the artifact. Present
  on `main` today; the always-on structural screen checks workbook structure
  (names, scopes, charts), not cached cell values, so it does not catch this.

## Documentation

Version-independent; not tied to a milestone.

The three mechanical drift checks of
[CONTRIBUTING.md § Documentation drift](../CONTRIBUTING.md#documentation-drift)
(review finding F7 — documentation drift is measurable) are all built:

- **Check 1 (link targets) and check 2 (cross-document anchors)** —
  `tests/test_doc_links.py`.
- **Check 3, count half** — `tests/test_doc_catalog_counts.py`, which found
  four stale numbers on its first run (139/139/131 for a 140-entry catalog,
  and 17 for 18 sheet-scoped closures).
- **Check 3, name half** — `tests/test_doc_function_names.py`. Every function
  name the docs *call* — written as `Name(...)`, in the catalog's naming
  convention, inside an inline span or a fenced block — resolves to a catalog
  entry or one of five pinned exclusion lists (native Excel functions, planned
  names, the v3.2 spill readers, retired names the shipped-changelog prose
  cites in rename history, doc shorthand). The exclusion lists are the hard
  part, so each is guarded: a planned name must still appear in ROADMAP or
  TODOs, a dead entry fails, and no entry may shadow a catalog function. The
  blind spots (bare backticked names, CamelCase single words like `Interact`)
  are deliberate and derived from the catalog by a test, so they cannot grow
  silently.

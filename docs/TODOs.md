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

Three self-contained items, in ascending cost:

1. [Retire the stale `v3.0` label](#v20-leftovers) — comments only, no executable logic reads it. `S · no Excel`
2. [Build the cross-document anchor check](#documentation) — ~40 lines of `re` plus a pytest case, runs in the existing Linux CI job. `S · no Excel`
3. [Add the P6 LSDV ↔ FE equivalence case](#test-model-suite) — one spec builder over data already wired, and the strongest cheap cross-oracle in the suite. `S · no Excel`

### Ready now — no Excel required

| Task | Size | Milestone |
|---|---|---|
| [Retire the stale `v3.0` label in comments](#v20-leftovers) | S | v2.0 |
| [Cross-document anchor check](#documentation) | S | Documentation |
| [LSDV ↔ FE equivalence case (P6)](#test-model-suite) | S | Test suite |
| [Auto MPG coverage gaps — `Difference`, `Ratio`, Cat×Cat, period override](#test-model-suite) | M | Test suite |
| [Wire Life Expectancy into the regression QC suite](#test-model-suite) | M | Test suite |
| [Rename `write_sheet_model_construction.py` → `write_spec_block.py`](#v20-leftovers) | M | v2.0 |
| [`Model_Formula_String` LAMBDA](#v34--model-comparison-sheet) | M | v3.4 |
| [`Cluster` Role — clustered-robust V_β](#v35--cluster-role-clustered-ses) | L | v3.5 |
| [`Moving_Average` / `Exponential_Smoothing`](#v36--time-role--time-series) | M | v3.6 |
| [`Numeric_Complete_Cases`](#v39--standalone-data-transformation-library) | S | v3.9 |
| [`Dummy_Column`, `Interact`, `Model_Matrix`](#v39--standalone-data-transformation-library) | M | v3.9 |
| [Location & Scale transforms — `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`](#v39--standalone-data-transformation-library) | M | v3.9 |
| [Group & Panel transforms — `Zscore_By`, `Decompose_By`](#v39--standalone-data-transformation-library) | M | v3.9 |
| [Two-sample tests — `T_Test_OneSample`, `F_Test_Variance`, `Covariance_Matrix`](#v310--bivariate--two-sample) | L | v3.10 |
| [`Bootstrap_CI` / `MC_Percentile` / `PERT_Sample`](#v311--resampling--simulation) | L | v3.11 |

The catalog-function items above are headless by construction: a LAMBDA lands in
`lambda_functions.json` with a Python mirror and tests, and only the *sheet* that
surfaces it needs Excel.

Backlog and version-independent rows come first; the milestone rows below them are
in ladder order. The ladder sorts on two keys — all remaining Regression work
first, then, within that, how much the
[regression test-model suite](MODEL_TESTING_ASSETS.md) has to grow to cover
each milestone. That is why the two flat-cost non-Regression milestones (v3.10,
v3.11) sit below four more expensive ones, and why the four v3.9 rows are last in
the Regression track — they are also listed in their own working order.

### Needs Excel on a developer machine

| Task | Size | Milestone |
|---|---|---|
| [Rebuild and commit the Regression artifact](#v31-leftovers) | S | v3.1 |
| [Promote `Sample_Include()` to a thunk over its spill](#v32--full-materialization-of-the-design-matrix) | M | v3.2 |
| [Point the engine call sites at the materialized spills](#v32--full-materialization-of-the-design-matrix) | L | v3.2 |
| [Re-examine the intercept-only closed-form bypass](#v30-leftovers) | M | v3.0 |
| [Diagnostic-chart reference lines](#v1x--regression-sheet) | M | v1.x |
| [Suppress worst-fit distributions from the combo charts](#v11-leftovers--univariate-sheet-writer) | M | v1.1 |
| [Beta: method-of-moments start + ~12×12 grid](#univariate-21--the-beta-half-of-the-grid-shrink) | L | Univariate 2.1 |
| [Relabel within-model residual outputs + Diagnostic Guide paragraph](#v21-leftovers--follow-on-polish) | S | v2.1 |
| [Model Comparison sheet layout](#v34--model-comparison-sheet) | L | v3.4 |
| [Time-series sheet (`write_sheet_time_series.py`)](#v36--time-role--time-series) | L | v3.6 |
| [Two-sample sheet layout](#v310--bivariate--two-sample) | M | v3.10 |
| [Simulation sheet layout](#v311--resampling--simulation) | M | v3.11 |

The Diagnostic Guide item also needs Excel for a second reason: that sheet is
baked into `templates/static_sheets.xlsx` and only regenerates through
`python scripts/rebuild_static_sheets.py`.

### Blocked on a decision

| Question | Milestone |
|---|---|
| [Blank-categorical caveat in `Sample_Include()`](#v20-leftovers) | v2.0 |
| [Mismatched-predictor-set fallback for the Comparison sheet](#v34--model-comparison-sheet) | v3.4 |
| [Can a column be both `Sequence` and `Time`](#v36--time-role--time-series) | v3.6 |
| [Two-sample selector — 3-way flag or separate `paired` boolean](#v310--bivariate--two-sample) | v3.10 |
| [How the guard-state configurations (G1–G13) are exercised](#test-model-suite) | Test suite |

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

## Univariate 2.1 — the Beta half of the grid shrink

The Weibull and Gamma half shipped as Univariate 2.0.0 (#160): both fits profile
their scale / rate parameter out in closed form and search a 20-point profile-NLL
column per stage, and the rebuilt artifact is committed —
`TestShippedUnivariateLayout` passes against it. Beta was deliberately out of
that scope and still runs the artifact's only two Data Tables at 20×20, so the
total is ~880 evaluations rather than the ~370 the shrink was costed at.

- **READY · L · needs Excel** — **Give Beta a method-of-moments start and a
  ~12×12 grid.** On the rescaled data with mean m and variance v:
  α₀ = m·(m(1−m)/v − 1), β₀ = (1−m)·α₀/m. Bracket both axes around that start
  the way `_write_profile_stage` brackets its 1-D start (`_PROFILE_BRACKET`),
  then drop `_N_GRID` from 20 to ~12. Beta stays two-dimensional — both of its
  conditional MLEs involve digamma, so neither parameter profiles out — and
  keeps `_write_grid_stage`, its two Data Tables, its heatmap, and both boundary
  rules. Breakage class: **MAJOR for the Univariate workbook version** if the
  Alpha/Beta Min/Max cells stop being plain typed inputs, the same call made for
  the Weibull/Gamma bounds at 2.0.0. See
  [DECISIONS.md § the grid shrink](DECISIONS.md#the-grid-shrink-ships-as-a-later-release-of-the-univariate-artifact)
  for the estimator and
  [DECISIONS.md § Univariate 2.0.0](DECISIONS.md#univariate-200--the-grid-shrink-weibull-and-gamma-half)
  for how the 1-D half resolved the equivalent questions.

## v2.0 leftovers

Specification-Driven Regression shipped 2026-07-05; the human test plan
(T0–T16) was executed, signed off PASS, and retired — its cases live on in
`tests/test_analyze_model_construction.py` and
`tests/test_difference_by_verification.py`. What remains is one open decision
and two cleanups.

- **READY · M · no Excel** — **Rename `write_sheet_model_construction.py` to
  match what it is.** It no longer writes a shipped sheet — both
  `build_production.py` and `build_univariate.py` call
  `_delete_sheet_if_present(workbook, "Model Construction")`. What it actually is
  now is the **spec-block component library**: `write_sheet_regression.py`
  imports `_write_spec_block`, `_write_spec_feedback`, `_write_intercept_control`,
  `_set_sheet_scoped_names`, `_set_spec_block_column_widths`, and every `_C_*`
  column constant from it. Proposed: rename to `write_spec_block.py` and update
  the importers (`write_sheet_regression.py`, `analyze_model_construction.py`,
  `analyze_regression_spec.py`, `analyze_regression_spec_block.py`,
  `tools/inspect_regression_sheet.py`, and five test modules). Mechanical;
  changes no behavior.

  Drop at the same time: `write_model_construction_sheet()`, `main()`, and
  `SHEET_NAME` — the standalone-CLI path, unreachable from any build.

  **Keep** `_write_audit_row` and `_write_filtered_zones`. They are the working
  reference implementations of the Design Columns audit column (now required —
  [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-ao)) and the V/W
  filtered-display pattern
  ([ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone)). Their
  `RecordingSheet` coverage in `tests/test_model_construction_writer.py` is the
  only test for that behavior. Note that promoting them into the Regression
  writer turned out **not** to be needed — the audit is one spill written by the
  shared `_write_spec_block`, so the Regression sheet inherits it by already
  calling that writer. (It was a per-row calculated column inside `SpecTable`
  when this was written; the block is table-free now, but the inheritance
  argument is unchanged.)

  Context: this is what remains of REVIEW.md F5 after the finding itself was
  struck as never-true — see
  [DECISIONS.md § v3.0 spec block](DECISIONS.md#the-spec-block-is-implemented-once-not-twice).

- **READY · S · no Excel** — **Retire the stale `v3.0` label for the spec-block
  changeover.** This changeover was planned as v3.0 and renumbered to v2.0 before
  release; v3.0 now means the engine-interface release (see
  [ROADMAP.md](ROADMAP.md)), so the old label is a live collision. The docstring
  in `write_sheet_model_construction.py` is corrected and the human test plan
  that carried the old label in its filename is retired. Still carrying the old
  label in **comments only** — no executable logic reads it:
  `write_sheet_model_construction.py` (several inline comments),
  `analyze_regression_spec_block.py` (module docstring), `build_production.py`
  (one comment), and `tests/test_dummy_functions.py`,
  `tests/test_model_construction_writer.py`, `tests/test_catalog_schema.py`.
  Deliberately left out of the v3.0 documentation pass, which was
  documentation-only and would otherwise have touched three test modules; do it
  as its own small commit.

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

- **READY · S · needs Excel** — **Relabel within-model residual outputs + add a
  Diagnostic Guide paragraph on residuals under FE.** Documentation-only, but the
  Diagnostic Guide is a static template sheet: edit
  `write_sheet_diagnostic_guide.py`, then run `python scripts/rebuild_static_sheets.py`
  and commit the regenerated `templates/static_sheets.xlsx` alongside it, or the
  change reaches no build.

## v3.0 leftovers

The engine-interface release shipped 2026-08-02 in three stages plus the
two-artifact split: stage 1 constructor pipeline + intercept relocation (#148),
stage 2 `Model_Context` (#150), stage 3 layout (#152), plus polish (#153, #154).
Every stage cleared the spec-driven Excel gate at 0 mismatches across all 12 QC
cases. Narrative in [ROADMAP.md](ROADMAP.md#v30--the-engine-interface-release--shipped-2026-08-02);
rationale in
[DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).

- **READY · M · needs Excel** — Re-examine the intercept-only closed-form bypass
  in `write_sheet_regression.py` → `_setup_local_names` (`Intercept_Only_N` /
  `_Point` / `_SE` / `_S` / `_DF`). `Design_Columns()` now returns a well-formed
  ones column in the zero-predictor state, so the engines could fit it directly
  and the bypass may be removable. Kept for stage 1 because the shipped
  behaviour was verified against it; retire it only with a QC pass behind it.

## v3.1 leftovers

Interaction wiring shipped 2026-08-03 (#156): spec columns M/N are read by
`Predictor_Columns()` and its two twins, the Design Columns audit counts the
interaction columns, and the Excel gate cleared. The same release gave the
catalog sole ownership of workbook scope (#159), clearing the v3.0 split's
cross-artifact name residue. See
[ROADMAP.md](ROADMAP.md#v31--interaction-wiring--shipped-2026-08-03) and
[DECISIONS.md § v3.1](DECISIONS.md#v31--interaction-wiring).

- **READY · S · needs Excel** — **Rebuild and commit the Regression artifact.**
  The committed `Lambda_Library.xlsx` predates v3.2: `BU2` and `BW2` still read
  the `"reserved"` placeholders where the `Sample_Include` and Constructed Design
  Matrix spills now belong. Its Version History and interaction wiring are
  current (2.1.0 / 2.2.0 / 3.1.0 entries are all present), so this is the
  materialization work and nothing else. The Univariate artifact is already
  rebuilt and committed (#160) — this item is Regression-only now.

  A second, cosmetic reason has since joined it: the Normal Q-Q chart no longer
  forces equal axis limits, so the shipped chart still carries the old scaling
  until a rebuild. No input cell, named range, or fitted number moves, so no
  version number moves with it.

## v3.2 — Full materialization of the design matrix

The stage-3 follow-on. The terminal Constructed Design Matrix zone, its collapse
behaviour, and its width guard shipped at v3.0; the spills that fill it — plus
`Sample_Include()`'s own zone — shipped at #161, each headed on
`_MATERIALIZATION_HEADER_ROW` and spilling from `_MATERIALIZATION_SPILL_ROW`,
full height and row-aligned with the source table. See
[ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone). What remains
is rewiring the readers, which is where the performance win actually is.

- **READY · L · needs Excel** — Point the readers at the spills. Surfacing the
  values did not rewire anything: `Design_Columns()` is still a live closure
  evaluated at each of its ~30 engine call sites, so the performance win the zone
  exists for is not yet banked. Same `#`-inside-a-`LAMBDA`-`RefersTo` question as
  the `Sample_Include` promotion below, and the same answer — it lands
  Excel-verified, not blind.

- **READY · M · needs Excel** — Promote `Sample_Include()` from a live closure to
  a thunk over its materialized spill. Deferred out of v3.0 stage 2 for a reason
  that still holds: it needs the dynamic-array spill operator (`#`) inside a
  `LAMBDA` defined-name `RefersTo`, a combination used nowhere else in this
  workbook and verifiable only with Excel present. A wrong guess breaks the
  row-mask contract that keeps every spilled array row-aligned. The live closure
  is untouched and remains the row mask until then. See
  [DECISIONS.md § materialization in two steps](DECISIONS.md#materialization-lands-in-two-steps--model_context-now-sample_include-deferred).

## v3.3 — Transforms remainder

Planned as the second half of v2.2; moved after v3.0 when the feature train was
resequenced — see [ROADMAP.md § v3.3](ROADMAP.md#v33--transforms-remainder--shipped-dispatcher--duan--model-formula-label).
**The standalone transform library is no longer part of this milestone** — it now
ships as [v3.9](#v39--standalone-data-transformation-library), at the end of the
Regression track, because it is that track's most expensive item to test.
The column-G `Log` wiring already shipped at v2.2: `Response_Column()` / `X_s()`
read column G, `Constructed_Column_Names()` / `Constructed_Column_Transforms()`
relabel and carry the per-column Log/None flag, the Prediction Inputs band
auto-logs, and `Ln_Positive` is in the catalog
([DECISIONS.md § v2.2](DECISIONS.md#v22--transforms--unit-space-comparability)).
The longitudinal transforms `Lag_By` / `Difference_By` also shipped early, with
gap-aware t−Δ semantics — verification in
`tests/test_difference_by_verification.py`.

- **RESOLVED · SHIPPED** — **Unit-space dispatcher functions** (design
  RESOLVED, code SHIPPED): eight catalog functions under the
  `Back-Transformation` subcategory — `Smearing_Factor`,
  `Back_Transform_Response`, `Unit_Space_Predictions`,
  `Unit_Space_Observed`, `Unit_Space_Residuals`, `Unit_Space_R_Squared`,
  `Unit_Space_Adjusted_R_Squared`, `Unit_Space_RMSE`. The dispatcher pattern
  (one canonical name per statistic, internal `SWITCH` on the transform pair)
  is the documented naming-style-departure pattern, in
  [DECISIONS.md § v3.3](DECISIONS.md#v33--transforms-remainder-unit-space-dispatch--duan-back-transformation--model-formula-label)
  and [ARCHITECTURE.md § 1 "Naming-style departures"](ARCHITECTURE.md#1-naming-convention).
  The transform pair is read off `Fit_Context()` (elements 3–4 the v3.0
  context array reserved for just this) rather than passed as positional
  arguments. SWITCH on the six recognised `(response, predictor)` pairs and
  `NA()` outside.

- **RESOLVED · SHIPPED** — Unit-space section on the Regression sheet:
  `AG3:AH9` block with the Back-Transform Method input (row 4), Smearing
  Factor (row 5), R² / Adj R² / RMSE in original units (rows 6–8), and the
  Response Space readout (row 9). The Original Units column (AL) in the
  Prediction Outputs block carries the back-transformed point estimate
  (Duan by default, Naive on toggle) and the four CI/PI bounds (always
  Naive, never smeared). The two new residual columns (AZ, BA) hold the
  back-transformed predicted and residual series. The `Comparison_*`
  sheet-scoped named ranges (`Comparison_Anchor`, `Comparison_Headline_GoF`,
  `Comparison_Model_Formula`) are the v3.4 Model Comparison sheet's reading
  surface.

- **RESOLVED · SHIPPED** — **Prediction back-transformation** (design
  RESOLVED, code SHIPPED): Duan's smearing estimator as the default, with a
  per-cell `Back_Transform_Method` toggle (`Duan` default | `Naive`).
  Caveat visible on the sheet as a note on the Back-Transform label at `AG4`: *Duan = Duan (1983) smearing;
  Naive = textbook EXP(ŷ), biased.* The resolution is in
  [DECISIONS.md § v3.3](DECISIONS.md#v33--transforms-remainder-unit-space-dispatch--duan-back-transformation--model-formula-label).

- **RESOLVED · SHIPPED** — Model Formula label: the sheet-scoped
  `Model_Formula()` catalog closure, rendered by a labelled readout on
  row 1 of the terminal Constructed Design Matrix zone, right of that zone's heading (header two columns over, readout three columns past the header). It shipped inline
  at `AA2:AB2` and moved off that cell afterwards — row 2 of the Regression
  Outputs zone wraps and AutoFits, so the sheet's longest string set the
  height of the whole header row. Built from the response-name lookup (which
  already emits `Ln(name)` when Log), `Allow_Intercept`,
  `Constructed_Column_Names()`, and the FE-name suffix gated by the Fixed
  Effects count.

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

## v3.6 — `Time` Role + time series

Planned as v2.7+, promoted out of the unordered bucket by the ladder reordering.
See [ROADMAP.md](ROADMAP.md#v36--time-role--time-series--planned).

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
  time-index semantics (for the future time-series sheet, for cross-sheet
  `Lag_By` / `Difference_By` calls). The open question: can a column be both
  `Sequence` and `Time`, or are they mutually exclusive?

- **READY · M · no Excel** — Implement `Moving_Average(data, window, [include])`.

- **READY · M · no Excel** — Implement `Exponential_Smoothing(data,
  alpha_smooth, [include])` — note: use `alpha_smooth` to distinguish from the
  significance-level `alpha`.

- **READY · L · needs Excel** — Implement `write_sheet_time_series.py` with
  forecast output, error metrics (MAE, RMSE, MAPE), and an actual vs. smoothed
  series chart.

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
[ROADMAP.md](ROADMAP.md#v39--standalone-data-transformation-library--planned).
Full specs in
[ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).

**Test assets — the ~10× axis-widener, and no new data.** Every new `Transform`
value widens the predictor-transform axis that today holds {None, Log}, and each
widening multiplies the response × predictor dispatch table. **Work the four items
below in the order given** — helpers first (they widen nothing), predictor-side
transforms next, and any response-side extension last, because a response
transform also multiplies the back-transformation / unit-space semantics. See
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

- **READY · S · no Excel** — Sample construction: `Numeric_Complete_Cases`.

- **READY · M · no Excel** — Categorical & model construction: `Dummy_Column`,
  `Interact`, `Model_Matrix`.

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

## v3.12+ — Unordered candidates (no claim)

What is left after the ladder reordering gave the other candidates numbers.
Nothing about their test cost sequences them: ANOVA-as-regression needs only
`warpbreaks` plus the existing categorical machinery, and the other two are
design-not-started. A user pressing for one of these would reorder it; absent that
signal, a single maintainer should not pre-order work that may not be the next
thing actually needed.

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
[docs/MODEL_TESTING_ASSETS.md](MODEL_TESTING_ASSETS.md). Section 1 of that
document lists ~15 model configurations tagged **new** — declared nowhere in the
code yet — against the three wired datasets. Each is a `RegressionSpecCase` in
`lambda_catalog/analyze_regression_spec.py` plus a pinned expectation in
`tests/test_regression_spec_qc.py`; the covering-array rule is that a case earns
its place by covering something no other case does.

- **READY · M · no Excel** — Add the **Auto MPG** gaps: the first `Difference`
  case (M10), the first `Ratio` case with its zero-denominator `#N/A` path and
  legal reciprocal pair (M11), the Cat×Cat full-product width (M9), the
  (Log, Log) pair on a dataset other than Production Lots (M5), and the typed
  Sequence-period override (M16).

- **READY · M · no Excel** — Wire the **Life Expectancy** dataset into the
  regression QC suite. It currently has *zero* spec cases, so every dispatch pair
  it would cover — (None, Mixed), (Log, None) with Duan *and* Naive, the
  `Ln_Positive` zero guard, the k = 200 width-guard soft warning, and
  high-cardinality Fixed Effects at 193 groups — is untested (L1–L9).

- **READY · S · no Excel** — Add **P6**, the LSDV ↔ within-estimator
  equivalence case: `production_lots_log_transform` with `Facility` declared as a
  Categorical Predictor instead of Fixed Effects must produce identical slopes and
  fit. The strongest cheap cross-oracle available, and it needs no new data.

- **OPEN · M · no Excel** — Decide how the ~10 **guard-state configurations**
  (G1–G13 — zero/duplicate Response rows, duplicate Sequence flags, Transform on a
  Categorical, unrecognized interaction operations) are exercised. They are not
  fittable models, so they do not fit the `RegressionSpecCase` oracle shape; they
  verify status lines and conditional formatting instead.

- **DEFERRED** — Wire the supplemental datasets (`warpbreaks`, the calendar-dated
  monthly series, `Grunfeld`, `Insurance`, `sleep`, `ToothGrowth`). Each one is
  one `CsvDatasetConfig` + one `SpecDatasetProfile` + a registry entry, and each
  lands with the milestone that needs it — see
  [§ 3 Timing](MODEL_TESTING_ASSETS.md#timing). The calendar-dated series is
  the exception worth pulling forward: it closes a coverage gap that exists today
  (see [v3.6](#v36--time-role--time-series)).

## Documentation

Version-independent; not tied to a milestone.

- **READY · S · no Excel** — **Build one of the two mechanical drift checks**
  proposed in
  [CONTRIBUTING.md § Documentation drift](../CONTRIBUTING.md#documentation-drift-proposed-check--not-yet-implemented)
  and tracked as review finding F7 (documentation drift is measurable),
  the one finding still open. The cheaper and higher-yield of the two is the
  **cross-document anchor check**: every `](FILE.md#anchor)` and `](#anchor)`
  resolves to a real heading in the target file. Roughly 40 lines of `re` plus a
  pytest case, no Excel, runs in the existing Linux CI job. The 2026-08-03 review
  ran it by hand (zero unresolved targets) and then deleted three documents —
  exactly the change that breaks anchors with no error anywhere. The second check
  (function names in docs resolving against `lambda_functions.json`) would have
  caught the stale `X_s` references and the 126-vs-131 count drift that same
  review found by hand.

# TODOs

Active work only. Resolved design decisions with their rationale live in
[DECISIONS.md](DECISIONS.md); foundational patterns live in
[ARCHITECTURE.md](ARCHITECTURE.md); the version plan lives in
[ROADMAP.md](ROADMAP.md). Every entry below has a cross-link to the
relevant DECISIONS entry for context on *why* — TODOs only holds
*what to do*.

**Conventions**

- `**OPEN**` — a question awaiting a decision. The full question and
  options live in the linked DECISIONS entry.
- `**DEFERRED**` — a question intentionally held for a specific
  future version. See DECISIONS for the deferral record.
- Items are listed in version order, then by ship order within a
  version.

---

## v1.x — Regression sheet

- TODO: Add reference lines to the Cook's Distance, PRESS Residuals, and
  Leverage vs. Studentized charts. Format thematically similar to the
  conditional formatting in the table (yellow = mild, red = strong). Use
  minimalist helper columns (2 anchor points using max/min of the
  relevant threshold; place them beneath the chart area). Same
  identity-line-via-real-data-series approach as the v1.2 diagnostic
  chart pattern — see [DECISIONS.md § v1.2 identity-line](DECISIONS.md#v12--workbook-hardening).

## v1.1 — Univariate (shipped; leftovers)

### Sheet writer (`write_sheet_univariate.py`)

- TODO: Investigate suppressing worst-fit / N/A-error distributions from
  the combo charts. Best outcome would be dynamically hiding those
  columns — hidden columns drop out of charts automatically
  (`PlotVisibleOnly` default) — but it is unclear whether column hiding
  can be driven from cell values without VBA (manual hiding works;
  data-driven hiding may be VBA-only, which the library forbids).
  No-VBA fallback to evaluate: emit `NA()` across a suppressed
  distribution's column, since line charts skip `#N/A` points — same
  chart effect without hiding. See
  [DECISIONS.md § v1.1 histogram overlays](DECISIONS.md#v11--univariate)
  for the combo-chart design.

### Additional distributions (long-term)

- TODO: Add support for more distribution families: Bernoulli, Binomial,
  Geometric, Negative Binomial, Hypergeometric, Poisson, Uniform,
  Chi-Square, Student-t.

## v2.0 — Specification-Driven Regression (shipped; leftovers)

Human test plan fully executed and signed off PASS 2026-07-05 (T0–T16).
One open decision remains from it:

- TODO: Resolve the blank-categorical caveat — `Sample_Include()`'s
  role-aware completeness layer requires numeric Response and numeric
  included Continuous Predictors, but Categorical Predictors impose no
  non-blank condition; a blank category value encodes as all-zero
  dummies (indistinguishable from the reference level). Run the caveat
  verification step in `HUMAN_TEST_PLAN_v3_model_construction.md` and
  record the decision: accept as documented behavior, or extend
  `Sample_Include()` with a non-blank condition for included
  Categorical Predictors. Interim workaround: a completeness column
  declared as a Filter. **OPEN** — see
  [DECISIONS.md § v2.0 auto-completeness](DECISIONS.md#v20--specification-driven-regression).

## v2.1 — Sequence, gap-aware longitudinal, serial-correlation diagnostics, fixed effects (in progress)

Two-way FE is deliberately deferred until this framework is finished — see
the v2.7+ section. Items below are in the locked ship order: the
Sequence fix is the prerequisite for the 2.1.0 entry; the FE Role
dropdown, the CI+PI prediction layout, and the engine are gated to ship
as a single release so users never see "FE is in the dropdown but the
engine is forthcoming."

### Pending (in ship order; #1 prerequisite, #2/#3/#9 gated to #5)

- TODO: **#1 — Sequence axis auto-detection and override** (renames
  column I to **`Sequence Period`** (the typed override input), adds
  column J **`Period In Use`** following the Reference Level /
  Reference In Use pattern). The current override mechanic has a
  spill-collision risk for source tables wider than the shipped WHO
  sample: the spec block reads its own H/I cells, and a longer table
  could let the override spill overrun an input band. Fix: relocate
  the override spill and bound every read of the H/I/J band by
  `COLUMNS(Source_Data)` (the spill-placement principle from
  CLAUDE.md). **Note: the Sequence Spacing block (rows 28–34), which
  previously hosted the override-flagging verdict lines, has since been
  removed — there is currently no on-sheet display of override status.**
  Update the spec layout constants, the named-range rename
  (`Spec_Base_Period_Delta` → `Spec_Sequence_Period`), and the QC
  analyzers. **Significant testing. Resolve before writing the 2.1.0
  Version History entry.**
  The full design rationale (spill-collision risk, the
  reference-level pattern parallel, override-flagging location) is in
  [DECISIONS.md § v2.1 #1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

- DONE: **#2 — FE Role dropdown + status-block validation.**
  `Fixed Effects` is in the Role axis (`_ROLE_VALIDATION_LIST`,
  `write_sheet_model_construction.py`); status-block cells for the active
  FE variable, group count, and absorbed df are live (no more "engine
  forthcoming" token — the engine backs them); a B1 cardinality error
  fires at 2+ FE rows (same pattern as the Sequence E1 check); the
  intercept × FE red flag is on the C2 toggle. The Role-axis design
  (cardinality, what FE contributes) is in
  [ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence).
  Tests: `tests/test_model_construction_writer.py`.

- DONE: **#3 — Surface BOTH intervals in adjacent cells of the
  prediction outputs section.** Nine rows: point · SE (Mean) · SE (New
  Obs) · t Critical · CI Lower/Upper · PI Lower/Upper · Confidence
  Level, via `Group_Prediction_Interval` (`lambda_functions.json`),
  wired at `write_sheet_regression.py::_write_prediction_interval`. The
  full math (the two variance terms, the group-specific width) is in
  [DECISIONS.md § v2.1 prediction interval](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).
  Tests: `tests/test_group_prediction_interval.py` (verified against an
  explicit LSDV `get_prediction()` reference and, for the no-FE case,
  bit-identical to the pre-v2.1 `Prediction_Interval()` numbers).

- DONE: **#4 — `Demean_By(x, group, [include])` and
  `Group_Mean(x, group, [include])`** (constructor internals, also
  user-callable transforms). The taxonomy and the v2.1 ship schedule
  are in [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).
  Tests: `tests/test_group_panel_transforms.py`.

- DONE: **#5 — `Is_Balanced_Panel(group, time, [include])`** —
  one-way/panel diagnostic; shipped with `Demean_By` (shares the "valid
  group set" primitive). Tests: `tests/test_group_panel_transforms.py`.

- DONE: **#6 — `Absorbed_Degrees_Of_Freedom()`** — G−1 via
  `Dummy_Levels` on the FE column, a Regression-sheet closure (not a
  standalone `(spec)`-argument function — reads `Spec_Role`/`Source_Data`
  directly, the same pattern as `Fixed_Effects_Column()`).

- DONE: **#7 — `y_s()`** and its predictor-side sibling `X_s_Within()`
  — new sheet-scoped closures, not replacements wired into existing
  no-FE call sites (`Response_Column()`/`X_s()` keep their other raw
  consumers — Predictor Summary, `Constructed_Column_Names()` labeling,
  `Intercept_Only_*`). Tests: `tests/test_within_estimator.py`.

- DONE: **#8 — `[DF_Absorbed]` argument (default 0) threaded through
  df / MS-residual / t-critical**, plus AIC/BIC/AICc's parameter count,
  across 23 engine functions. `SE_Coefficients` needed an exact
  rescaling (`SQRT(naive_df/true_df)`) rather than a direct threading,
  since Excel's `LINEST` always computes its own df. Bit-equality of
  every existing no-FE case confirmed by construction (default 0 ⇒
  rescale factor 1 / subtraction no-op) and numerically; FE-active SE/t/
  p/CI/AIC match an independent `statsmodels` LSDV fit. The default-0 →
  identical no-FE pattern (the "non-breaking MINOR" guarantee) is in
  [DECISIONS.md § v2.1 df plumbing](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).
  Tests: `tests/test_df_absorbed_threading.py`.

- DONE: **#9 — FE group selection + ȳᵢ / Tᵢ cells** (x̄ᵢ is computed
  internally by `Group_Prediction_Interval`/`xbar_i` rather than also
  surfaced per-predictor on the sheet — column AJ is the fixed
  Prediction-Outputs/Residual-Output gap column, so there was no fourth
  column available without breaking the outline-grouping architecture;
  a deliberate scope trim, not an oversight). `Group_Mean_At`/
  `Group_Count_At`/`Prediction_Group_Column` back both the group-mean
  form and the visible Group Mean (y) / Group Count readouts; the BFN
  cell already flips from `n/a — no fixed effects` to active (this was
  forward-wired before v2.1, confirmed working once #2 made the Role
  selectable). The group-mean-recovery form is in
  [DECISIONS.md § v2.1 FE point prediction](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

**Remaining before the 2.1.0 Version History entry:** the QC-oracle chain
(`regression_shared.RegressionPredictionInterval`, `analyze_regression_sheet.py`,
`analysis_cache.py`, `tools/inspect_regression_sheet.py`) still models the
pre-v2.1 6-value Prediction Interval shape, not the new 9-value CI+PI/
group-mean-recovery form — flagged in `tools/inspect_regression_sheet.py`
with a code comment. Not run in CI regardless (requires desktop Excel), but
should be updated before relying on it for a `--verify` pass on an FE-active
model. A human test plan covering T0 (pooled baseline) through T4 (degenerate
FE variable) on the WHO Life Expectancy panel is at
`HUMAN_TEST_PLAN_v21_regression_fixed_effects.md`.

### Follow-on polish (ships with 2.1.0 if there's room)

- TODO: **BFN critical values**. **DEFERRED** — N,T-dependent bounds
  per Bhargava et al. 1982 tables; do NOT present standard DW bounds
  next to the BFN cell. The deferral record (why N,T-dependent, why
  not standard DW) is in
  [DECISIONS.md § v2.1 BFN critical values](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

- TODO: **Categorical × FE prediction encoding**. **DEFERRED** — x_new
  and x̄ᵢ formed in constructed design-matrix space; UI wire to encode
  through `Dummy_Code` before reaching the FE formula. Largely
  subsumed by v2.0 categorical prediction; recorded so the encoding
  step is not forgotten. The deferral record is in
  [DECISIONS.md § v2.1 Categorical × FE](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

- TODO: **Relabel within-model residual outputs + Diagnostic Guide
  paragraph on residuals under FE**. Documentation-only.

## v2.2 — Transforms & the standalone transform library

### Transform wiring (spec column G)

- TODO: `Transform` dropdown gains `Log`; wire `X_s()` /
  `Constructed_Column_Names()` / prediction to read column G. The
  reserved-column policy (column G shipped as a placeholder, unread by
  any formula) is in
  [ARCHITECTURE.md § 4 "Reserved-column policy"](ARCHITECTURE.md#4-the-model-spec-block-al).

- TODO: **Unit-space dispatcher function, RESOLVED:**
  `Unit_Space_R_Squared(model, response_transform, predictor_transform)`,
  `Unit_Space_Adjusted_R_Squared(...)`, `Unit_Space_RMSE(...)`. The
  dispatcher pattern (one canonical name per statistic, internal
  `SWITCH` on the transform pair) is the documented
  naming-style-departure pattern, in
  [DECISIONS.md § v2.2 unit-space dispatcher](DECISIONS.md#v22--transforms--unit-space-comparability)
  and [ARCHITECTURE.md § 1 "Naming-style departures"](ARCHITECTURE.md#1-naming-convention).

- TODO: Unit-space section on the Regression sheet — SWITCH on column
  G, one headline comparable statistic (the cell v2.3 Model Comparison
  will reference).

- TODO: **Prediction back-transformation, RESOLVED:** Duan's smearing
  estimator as the default, with a per-cell `Back_Transform_Method`
  toggle (`Duan` default | `Naive`). Caveat row visible on the sheet:
  *Duan = Duan (1983) smearing; Naive = textbook EXP(ŷ), biased.* The
  resolution (Duan as default, the per-cell toggle, the caveat row)
  is in
  [DECISIONS.md § v2.2 prediction back-transformation](DECISIONS.md#v22--transforms--unit-space-comparability).

### Standalone Data Transformation functions (specs in ARCHITECTURE.md)

- TODO: Location & Scale — `Center`, `Zscore`, `Minmax_Scale`,
  `Winsorize`, `Ln_Positive`. The full specs are in
  [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).

- TODO: Group & Panel — `Zscore_By`, `Decompose_By`
  (`Demean_By`/`Group_Mean` arrive at v2.1; two-way functions follow
  the two-way FE milestone).

- ~~Longitudinal — `Lag_By`, `Difference_By`~~ — **DONE (shipped early,
  base-period release)** with the gap-aware t−Δ semantics: exact-match
  lookup of (group, seq−Δ) pairs, `NA()` at first periods and gaps,
  `[delta]` defaulting to the spec's Period In Use cell via
  `Base_Period_Delta()` (never a silent 1). The same release wired
  spec column I (typed override → Sequence Period) and J
  (candidate-with-override display → Period In Use) plus the Sequence
  Spacing block (delta spectrum, Regularity/Off-grid flags,
  calendar-signature guidance). Verification:
  `tests/test_difference_by_verification.py`; human test plan T17–T19.
  The shipped semantics and the `NA()` exception are in
  [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy)
  and
  [DECISIONS.md § v2.1 base-period layer](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

- TODO: Sample construction — `Numeric_Complete_Cases`.

- TODO: Categorical & model construction — `Dummy_Column`, `Interact`,
  `Model_Matrix`.

## v2.3 — Model Comparison Sheet

- TODO: Implement the `Model_Formula_String` LAMBDA with
  header-signature validation (`NA()` on non-Regression targets). The
  name resolution and the anchor-cell argument-type rationale are in
  [DECISIONS.md § v2.3 Model_Formula_String](DECISIONS.md#v23--model-comparison-sheet).

- TODO: Sheet layout — model registry (hyperlinks), GoF table
  referencing the v2.2 unit-space headline cells, shared prediction
  inputs (Comparison sheet is the source; Regression sheets pull via
  XLOOKUP), prediction results table. The data-flow direction
  (Comparison-as-source-via-XLOOKUP) is in
  [DECISIONS.md § v2.3 prediction inputs](DECISIONS.md#v23--model-comparison-sheet).

- TODO: Decide the mismatched-predictor-set fallback (XLOOKUP
  `[if_not_found]`). **OPEN** — see the open-decision note in
  [DECISIONS.md § v2.3 Model Comparison Sheet](DECISIONS.md#v23--model-comparison-sheet).

## v2.4 — Resampling & Simulation

- TODO: **No-volatile constraint, RESOLVED: pre-drawn random table.** A
  single sheet-scoped named range `Bootstrap_Random_Draws` holds a
  uniformly-distributed random table pre-drawn once at build time,
  seeded from the same SHA-derived seed the QC build already uses
  (`analysis_cache.py`). `Bootstrap_CI` indexes via
  `INDEX(Bootstrap_Random_Draws, MOD(SEQUENCE(n_resamples), ROWS(Bootstrap_Random_Draws))+1)`.
  Same inputs → same output, every recalc. `RANDARRAY()` rejected. The
  full rationale (auditability vs. fresh randomness, the
  reproducibility trade-off) is in
  [DECISIONS.md § v2.4 no-volatile constraint](DECISIONS.md#v24--resampling--simulation).
  To get a new draw, regenerate the workbook via `build_production.py`
  (deliberate, not a limitation).

- TODO: Implement `Bootstrap_CI(data, stat_lambda, n_resamples,
  alpha, [include])` — bootstrap confidence interval for an
  arbitrary statistic passed as a LAMBDA. Uses the pre-drawn table
  above.

- TODO: Implement `MC_Percentile(dist_params, n_samples, percentile)` —
  Monte Carlo draw from a fitted distribution; complements v2.0
  fitting. Uses the same pre-drawn table.

- TODO: Implement `PERT_Sample(min, mode, max, n_samples)` — BetaPERT
  sampling for cost/schedule risk analysis. Uses the same pre-drawn
  table.

- TODO: Design sheet layout (bootstrap section + Monte Carlo section;
  may share one sheet). Implement `write_sheet_simulation.py`.

## v2.5+ — Future (sequence TBD; first two claimed)

The v2.5+ bucket previously had seven candidates with no order.
Two-sample tests are now v2.5 (next MINOR after v2.4) and the `Weight`
Role is v2.6 (after v2.5). The rest are deliberately unordered pending
actual user demand — a single maintainer should not pre-order work that
may not be the next thing actually needed.

### v2.5 — Bivariate / Two-sample *(claimed, next MINOR after v2.4)*

- TODO: Implement `T_Test_OneSample(data, mu0, alpha, [include])` →
  test statistic, p-value, CI.

- TODO: Implement `T_Test_TwoSample(data1, data2, alpha, equal_var,
  [include1], [include2])` — equal-variance, Welch unequal-variance,
  and paired variants. **OPEN design question:** paired is a separate
  code path the `equal_var` flag does not cover — 3-way flag or
  separate `paired` boolean? See
  [DECISIONS.md § v2.5 two-sample selector](DECISIONS.md#v25--claimed).

- TODO: Implement `F_Test_Variance(data1, data2, alpha, [include1],
  [include2])` — output feeds a recommendation cell that selects the
  appropriate t-test variant.

- TODO: Implement `Covariance_Matrix(data, [include])` — sample
  covariance (consistent with the existing catalog's sample-statistic
  convention); complement to the existing `Correlation_Matrix`.

- TODO: Design two-sample sheet layout: inputs, test selector, F-test
  assumption check, output (test statistic, df, p-value, CI, effect
  size). Implement `write_sheet_two_sample.py`.

### v2.6 — `Weight` Role (WLS) *(claimed, after v2.5)*

The standalone WLS milestone and its `[weights]`-argument-vs-parallel-
function-set debate are superseded by a **`Weight` value on the Role
axis** (see
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)).
Three-stage scope carried forward: user-supplied weights →
variance-driver-derived weights → FGLS. v2.6 ships the first stage
only. The default-uniform → OLS pattern (the
"non-breaking MINOR" guarantee) is in
[DECISIONS.md § v2.6 WLS](DECISIONS.md#v25--claimed).

- TODO: Implement the `Weight` Role (at most one, per the cardinality
  rule that Response, Time, and Weight share; status-block validation
  identical to exactly-one-Response).

- TODO: Thread weights through the engine per the Role-axis design: a
  single optional `[Weights]` argument (default uniform) on the
  inferential chain. Default-uniform means every existing OLS call
  computes identically — the v2.1 `[DF_Absorbed]` precedent (default
  0 → identical no-FE model) is the exact pattern to follow.

- TODO: Update the Diagnostic Guide to describe which diagnostics
  change interpretation under WLS. (WLS closes the loop opened by
  v1's Scale-Location diagnostic.)

### v2.7+ — Unordered candidates (no claim)

The following are real candidate work but deliberately unordered. Two-way
FE and `Cluster` have partial forward wiring (from v2.1 FE and the
`Serial_Correlation_Group()` resolver); the rest are design-not-started.
A user-pressing-for-them signal would reorder these; absent that, they
stay in this unordered bucket.

#### Two-way Fixed Effects

- TODO: Implement `Absorb_Two_Way_Fixed_Effects(x, group1, group2,
  [include], [passes])` (alternating-projection demeaning for
  unbalanced panels).

- TODO: Implement `Demean_Two_Way_Balanced(x, group1, group2, [include])`
  and the two-way `Is_Balanced_Panel` check.

- TODO: Implement `Fixed_Effects_Convergence_Check(x, group1, group2,
  [include])`; surface in the status block whenever two FE variables
  are active.

- TODO: Lift the v2.1 one-FE-variable status-block error; resolve the
  two-way prediction question (group intercepts are not recoverable
  as simple group means). The one-way-scope rationale is in
  [DECISIONS.md § v2.1 scope](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

#### Multi-group means (ANOVA)

- TODO: Implement one-way ANOVA as regression on group dummies,
  reusing the existing SS/MS/F machinery. Frame explicitly as "ANOVA
  is regression" — group means, SS decomposition, and F-test should
  match the MLR output exactly.

- TODO: Add post-hoc comparisons (Tukey HSD or Bonferroni) as an
  optional output section.

#### `Cluster` Role (clustered SEs)

- TODO: Implement the `Cluster` Role (at most one) — clustered-robust
  variance estimator. Has partial forward wiring from
  `Serial_Correlation_Group()`'s dormant Cluster branch (PR #106), so
  the resolver side is partial; the engine side (cluster-robust V_β)
  is not.

- TODO: Lift the v2.1 `n/a — engine forthcoming` token on the BFN
  cell when Cluster is active (the BFN formula already uses
  `Serial_Correlation_Group()` as its resolver, so the wiring is
  partial).

#### `Time` Role (time-index designation)

- TODO: Design and implement the `Time` Role. Partially forward-wired
  via the v2.1 Sequence axis, but the full `Time` Role adds
  time-index semantics (for the future time-series sheet, for
  cross-sheet `Lag_By`/`Difference_By` calls). **OPEN design
  question:** can a column be both `Sequence` and `Time`, or are
  they mutually exclusive?

#### Time series

- TODO: Implement `Moving_Average(data, window, [include])`.

- TODO: Implement `Exponential_Smoothing(data, alpha_smooth, [include])`
  — note: use `alpha_smooth` to distinguish from the
  significance-level `alpha`.

- TODO: Implement `write_sheet_time_series.py` with forecast output,
  error metrics (MAE, RMSE, MAPE), and an actual vs. smoothed series
  chart.

#### Long-tail (out of planning horizon)

- **Fourier analysis** — long-tail; the *ToolPak Parity Reference*
  notes it is "intentionally skipped" and a later
  addition-by-demand decision, not a planned milestone.
- **Decision analysis** — long-tail (loss functions, cost/risk
  oriented). Not on the planning horizon.

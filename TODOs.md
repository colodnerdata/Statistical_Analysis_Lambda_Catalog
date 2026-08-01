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
One open decision remains from it, plus one numbering-cleanup item:

- TODO: **Rename `write_sheet_model_construction.py` to match what it is.**
  It no longer writes a shipped sheet — both `build_production.py` and
  `build_qc.py` call `_delete_sheet_if_present(workbook, "Model
  Construction")`. What it actually is now is the **spec-block component
  library**: `write_sheet_regression.py` imports `_write_spec_block`,
  `_write_spec_feedback`, `_write_intercept_control`,
  `_set_sheet_scoped_names`, `_set_spec_block_column_widths`, and every
  `_C_*` column constant from it. Proposed: rename to `write_spec_block.py`
  and update the importers (`write_sheet_regression.py`,
  `analyze_model_construction.py`, `analyze_regression_spec.py`,
  `analyze_regression_spec_block.py`, `tools/inspect_regression_sheet.py`,
  and five test modules). Mechanical; changes no behavior.

  Drop at the same time: `write_model_construction_sheet()`, `main()`, and
  `SHEET_NAME` — the standalone-CLI path, unreachable from any build.

  **Keep** `_write_audit_row` and `_write_filtered_zones`. They are the
  working reference implementations of the Design Columns audit column
  (now required — [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-an))
  and the V/W filtered-display pattern the Constructed Design Matrix
  promotes to production
  ([ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone)).
  Promote them into the Regression writer as v3.0 builds those; do not
  delete and rewrite. Their `RecordingSheet` coverage in
  `tests/test_model_construction_writer.py` is the only test for that
  behavior.

  Context: this is what remains of REVIEW.md F5 after the finding itself
  was struck as never-true — see
  [DECISIONS.md § v3.0 spec block](DECISIONS.md#the-spec-block-is-implemented-once-not-twice).

- TODO: **Retire the stale `v3.0` label for the spec-block changeover.**
  This changeover was planned as v3.0 and renumbered to v2.0 before
  release; v3.0 now means the engine-interface release (see
  [ROADMAP.md](ROADMAP.md)), so the old label is a live collision. The
  docstring in `write_sheet_model_construction.py` and the human test
  plan filename are corrected. Still carrying the old label in
  **comments only** — no executable logic reads it:
  `write_sheet_model_construction.py` (several inline comments),
  `analyze_regression_spec_block.py` (module docstring),
  `build_production.py` (one comment), and
  `tests/test_dummy_functions.py`, `tests/test_model_construction_writer.py`,
  `tests/test_catalog_schema.py`. Deliberately left out of the v3.0
  documentation pass, which was documentation-only and would otherwise
  have touched three test modules; do it as its own small commit.

- TODO: Resolve the blank-categorical caveat — `Sample_Include()`'s
  role-aware completeness layer requires numeric Response and numeric
  included Continuous Predictors, but Categorical Predictors impose no
  non-blank condition; a blank category value encodes as all-zero
  dummies (indistinguishable from the reference level). Run the caveat
  verification step in `HUMAN_TEST_PLAN_v20_model_construction.md` and
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

- DONE: **#1 — Sequence axis auto-detection and override.** Column I
  is **`Sequence Period`** (the typed override input), column J is
  **`Period In Use`** (computed-with-override display), following the
  Reference Level / Reference In Use pattern. Verified directly: no
  stale `Spec_Base_Period_Delta` reference remains anywhere in the
  codebase (fully renamed to `Spec_Sequence_Period`), the spec layout
  constants (`_C_SEQUENCE_PERIOD`, `_C_PERIOD_IN_USE`) are wired
  consistently, and the feature is live in shipped workbooks (confirmed
  against a real built `Regression` sheet).
  Two sub-items resolved differently than originally envisioned, not
  left undone: the spill-collision risk is structurally moot — `Spec_Sequence`/
  `Spec_Sequence_Period`/`Spec_Period_In_Use` are `SpecTable[[#Data],[...]]`
  structured references, which auto-bound to the live table rows (the
  same auto-extend behavior ARCHITECTURE.md § 4 documents for the whole
  spec block), rather than needing a manual `TAKE(...,COLUMNS(Source_Data))`
  bound — and "update the QC analyzers" turned out to mean the
  RecordingSheet unit-test layer (`test_spec_ranges_cover_the_standard_input_band`,
  `test_spec_feedback_writes_delta_count_verdict_with_priority_cf` in
  `tests/test_model_construction_writer.py`), not the xlwings-based
  `analyze_regression_spec.py` oracle, which never modeled this feature
  and doesn't need to. The on-sheet override-status display the Sequence
  Spacing block used to carry now lives in the I2 combined Verdict cell
  (off-grid/regularity/no-natural-base-period messages driven by
  `Spec_Period_In_Use` vs. `Sequence_Deltas()`) — a data-quality check
  that serves the same "is the declared Δ trustworthy" need, not a
  literal "you typed an override" flag.
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

- DONE: **#10 — QC-oracle rebuild for the v2.1 Prediction Interval shape.**
  `regression_shared.RegressionPredictionInterval`, `analyze_regression_sheet.py`,
  `analysis_cache.py` (schema v16), and `tools/inspect_regression_sheet.py`
  moved from the pre-v2.1 6-value single-CI shape to the shipped 9-value
  CI+PI/group-mean-recovery form (point, se_mean, se_new, t_critical,
  ci_lower, ci_upper, pi_lower, pi_upper, confidence_level) plus
  `group_mean`/`group_count` for the AH13/AH14 readouts, via
  `Group_Prediction_Interval`'s own formula (group-mean recovery: refit on
  the group-demeaned pair — even a no-FE case demeans by one constant
  `"(all)"` group — then `ybar_i + (x_new - xbar_i)'beta`). This closed the
  58 automated `Regression/prediction_interval` mismatches `build_qc.py
  --verify` was reporting (confirmed pre-existing/independent of the
  Production Lots/Fixed Effects QC work, PR #133, via a clean `origin/main`
  baseline run before starting this fix).
  `RegressionSpecCase` gained a `prediction_group: str | None` field (`None`
  → the sheet's own default, the alphabetically-first observed group);
  `tools/inspect_regression_sheet.py::_apply_spec_case` writes the
  resolved group into the sheet's own `$AH$12` cell — positioned above the
  variable-size Prediction Inputs band (rows 19+), so this only ever
  overwrites a fixed cell, never something that needs to grow. One bug
  caught along the way: `beta`/`sigma` can only be shortcut through the
  main fit's coefficients/SE when an intercept is present (centering
  doesn't move an intercept-included fit's slopes or residuals by FWL,
  but does for a through-the-origin no-intercept fit) — first pass
  reused the shortcut unconditionally and broke exactly the three
  `*_no_intercept` cases; fixed by always refitting on the group-demeaned
  pair rather than assuming an equivalence that only sometimes holds.
  Verified against two independent sources before landing (bit-exact to
  every T1 reference number in `HUMAN_TEST_PLAN_v21_regression_fixed_effects.md`
  and an independent statsmodels LSDV cross-check), then a full spec-case
  sweep against a live Regression sheet: 0 mismatches across all 12 cases
  (132 `prediction_interval` rows), 0 elsewhere. Tests:
  `tests/test_group_prediction_interval.py` (the reference math),
  `tests/test_independent_verification.py`, `tests/test_qc_configs.py`.

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

- ~~`Transform` dropdown gains `Log`; wire `X_s()` /
  `Constructed_Column_Names()` / prediction to read column G.~~ —
  **DONE.** Column G's dropdown is `None`/`Log`; `Log` is read by
  `Response_Column()` (Response row) and `X_s()`'s Continuous branch
  (Predictor rows), both modified in place rather than wrapped, so every
  existing consumer (including the v2.1 FE wrappers `y_s()`/`X_s_Within()`,
  unchanged) inherits log-space data automatically. `Constructed_Column_Names()`
  relabels a logged column `Ln(name)`; a new structural twin,
  `Constructed_Column_Transforms()`, gives the per-constructed-column
  Log/None flag the Prediction Inputs band needs (a Categorical
  Predictor's dummy columns always read `None`). Log is disallowed —
  flagged red, not silently ignored — on Categorical Predictors. The
  Prediction Inputs band takes a raw value and auto-logs it internally
  (never a typed ln(x)); the Training Mean spill emits the geometric mean
  for a logged column to avoid double-logging the default. Residual-output
  headers and the audit-strip response name gain a `(Log)` suffix when
  active. New catalog function `Ln_Positive(x, [include])` backs all of
  this (`NA()` on an included non-positive/non-numeric value, `""` on an
  excluded row). Scope explicitly excludes the unit-space dispatcher and
  Duan back-transformation below — those remain open. Verified against a
  real learning-curve model: a new QC case
  (`production_lots_log_transform`, raw `Cumulative_Units`/`Unit_Cost_BY`
  with `transform="Log"`) matches the pre-existing
  `production_lots_fixed_effects` case (precomputed log columns) to
  floating-point precision — `tests/test_transform_threading.py`,
  `tests/test_ln_positive_verification.py`. Full design rationale in
  [DECISIONS.md § v2.2 Transform column wiring](DECISIONS.md#v22--transforms--unit-space-comparability).

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
  `Winsorize`. The full specs are in
  [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).
  (`Ln_Positive` shipped early, alongside the Transform column-G wiring
  above, rather than waiting for the rest of this bundle.)

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

## v3.0 — The engine-interface release (in progress)

Delivered in three stages, in dependency order. Stage 1 is built; the
scope decision and the stage table live in
[ROADMAP.md](ROADMAP.md) and
[DECISIONS.md § v3.0 ships in three stages](DECISIONS.md#v30-ships-in-three-stages).

### Stage 1 — constructor pipeline + intercept relocation — **BUILT**

- DONE: `X_s()` / `X_s_Within()` / `y_s()` renamed to
  `Predictor_Columns()` / `Design_Columns()` / `Design_Response()`, with
  the `encode → transform → demean → intercept → weight` stage order
  made explicit in `Design_Columns()`. The weight stage is declared and
  inert until v2.6.
- DONE: the intercept moved into the constructor. `Design_Matrix` stops
  synthesizing it, LINEST runs `const = FALSE` at all three call sites,
  `SS_Total` became the intercept-only residual sum of squares, and
  `[Allow_Intercept]` left all 48 signatures — 13 now carry
  `[Has_Intercept]` as an identifier.
- DONE: the QC test-sheet writers render formulas from declared argument
  names (`make_test_sheet.build_call`) rather than positional lists.

- TODO: **Run the spec-driven verifier on a machine with Excel** —
  `python build_production.py --verify --no-launch
  --skip-data-table-calculations --skip-univariate`. Stage 1 must report
  **0 mismatches across all 12 QC cases**: it changes where the intercept
  is created, not what is fitted, so any mismatch is a bug rather than an
  expected delta. Not runnable in CI (no Office on the GitHub-hosted
  runner) — see [CONTRIBUTING.md](CONTRIBUTING.md) → *Verifying builds*.

- TODO: Re-examine the intercept-only closed-form bypass in
  `write_sheet_regression.py` → `_setup_local_names`
  (`Intercept_Only_N` / `_Point` / `_SE` / `_S` / `_DF`).
  `Design_Columns()` now returns a well-formed ones column in the
  zero-predictor state, so the engines could fit it directly and the
  bypass may be removable. Kept for stage 1 because the shipped
  behaviour was verified against it; retire it only with a QC pass
  behind it.

### Stage 2 — `Model_Context` — PLANNED

- TODO: Collapse `[Has_Intercept]` and `[DF_Absorbed]` into a single
  `[Context]` argument across the 13 + 24 carriers; materialize
  `Model_Context` and `Sample_Include` as spill ranges on the Regression
  sheet, with a build assertion that `ROWS(Model_Context())` is a
  build-time constant.

### Stage 3 — layout — PLANNED

- TODO: Insert spec columns M (Interaction Term) and N (Interaction
  Operation), reserved-and-unwired; build the Design Columns audit
  column; establish the Constructed Design Matrix zone and its two-
  threshold width guard; move the version number to 3.0.

- TODO: Promote `_write_audit_row` and `_write_filtered_zones` from
  `write_sheet_model_construction.py` into the Regression writer rather
  than rewriting them — they are the working reference implementations
  of the audit column and the filtered-display pattern.

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

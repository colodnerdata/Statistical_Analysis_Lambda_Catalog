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

## Univariate artifact 2.1 — the Beta half of the grid shrink

The Weibull and Gamma half shipped as Univariate 2.0.0: both fits profile
their scale / rate parameter out in closed form and search a 20-point
profile-NLL column per stage. Beta was deliberately out of that scope and
still runs the artifact's only two Data Tables at 20×20, so the total is
~880 evaluations rather than the ~370 the shrink was costed at.

- TODO: **Give Beta a method-of-moments start and a ~12×12 grid.** On the
  rescaled data with mean m and variance v: α₀ = m·(m(1−m)/v − 1),
  β₀ = (1−m)·α₀/m. Bracket both axes around that start the way
  `_write_profile_stage` brackets its 1-D start (`_PROFILE_BRACKET`), then
  drop `_N_GRID` from 20 to ~12. Beta stays two-dimensional — both of its
  conditional MLEs involve digamma, so neither parameter profiles out —
  and keeps `_write_grid_stage`, its two Data Tables, its heatmap, and both
  boundary rules. Breakage class: **MAJOR for the Univariate workbook
  version** if the Alpha/Beta Min/Max cells stop being plain typed inputs,
  the same call made for the Weibull/Gamma bounds at 2.0.0. See
  [DECISIONS.md § the grid shrink](DECISIONS.md#the-grid-shrink-ships-as-a-later-release-of-the-univariate-artifact)
  for the estimator and
  [DECISIONS.md § Univariate 2.0.0](DECISIONS.md#univariate-200--the-grid-shrink-weibull-and-gamma-half)
  for how the 1-D half resolved the equivalent questions.

- TODO: **Rebuild and commit `Lambda_Library_Univariate.xlsx`.** The 2.0.0
  writer changes have not been rendered into the committed artifact — that
  needs a machine with Microsoft Excel (`python build_univariate.py
  --verify --no-launch`). Until then the committed workbook still carries
  the four 20×20 Weibull/Gamma formula grids, their `UV_WB_*` / `UV_GAMMA_*`
  bodies at the old 21-column extents, and no `UV_Profile_*` names or
  profile charts. The headless invariant tests pass against it because they
  check package structure, not sheet layout.

## v2.0 — Specification-Driven Regression (shipped; leftovers)

The human test plan for this milestone was fully executed and signed off PASS
2026-07-05 (T0–T16), and has since been retired — its cases live on in
`tests/test_analyze_model_construction.py` and
`tests/test_difference_by_verification.py`. One open decision remains from it,
plus one numbering-cleanup item:

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
  (now required — [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-ao))
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
  docstring in `write_sheet_model_construction.py` is corrected, and the
  human test plan that carried the old label in its filename is retired.
  Still carrying the old label in
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
  dummies (indistinguishable from the reference level). Verify against a live
  build — declare a Categorical Predictor whose column has a blank value on at
  least one otherwise-complete row, confirm the row is included and encodes as
  the reference level, and check whether the fitted coefficients shift — then
  record the decision: accept as documented behavior, or extend
  `Sample_Include()` with a non-blank condition for included
  Categorical Predictors. Interim workaround: a completeness column
  declared as a Filter. **OPEN** — see
  [DECISIONS.md § v2.0 auto-completeness](DECISIONS.md#v20--specification-driven-regression).

## v2.1 — Sequence, gap-aware longitudinal, serial-correlation diagnostics, fixed effects (shipped within 3.0.0; leftovers)

The engine and sheet work is done and shipped — it reached users inside the
3.0.0 artifact, never as its own release build, which is why the 2.1.0
Version History entry is still missing (see § Documentation at the end of
this file). Two-way FE is deliberately deferred until this framework is
finished — see the v3.8+ section. What remains here is DEFERRED polish.

### Done — engine and sheet work (TODOs #1–#10)

All ten items are DONE and verified against a live build (0 mismatches
across all 12 spec-driven QC cases). Design rationale for each lives in
[DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects);
verification evidence lives in the named test modules and git history. In
ship order: #1 Sequence Period / Period In Use split, #2 FE Role dropdown
+ status-block validation, #3 CI+PI prediction layout, #4 `Demean_By` /
`Group_Mean`, #5 `Is_Balanced_Panel`, #6 `Absorbed_Degrees_Of_Freedom`, #7
`y_s` / `X_s_Within` (renamed at v3.0 to `Design_Response` /
`Design_Columns`), #8 `[DF_Absorbed]` threading across 23 engine
functions, #9 FE group selection + group-mean readouts, #10 QC-oracle
rebuild for the 9-value Prediction Interval shape.

### Follow-on polish (deferred)

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

## v2.2 — Transforms & the standalone transform library (shipped; remainder is v3.3)

### Transform wiring (spec column G)

Column-G `Log` wiring shipped at v2.2 — `Response_Column()` / `X_s()`
read column G, `Constructed_Column_Names()` / `Constructed_Column_Transforms()`
relabel and carry the per-column Log/None flag, the Prediction Inputs band
auto-logs, and `Ln_Positive` is in the catalog. Verified against a real
learning-curve model; design rationale in
[DECISIONS.md § v2.2 Transform column wiring](DECISIONS.md#v22--transforms--unit-space-comparability).
The unfinished remainder (unit-space dispatcher, prediction
back-transformation, standalone transform library) moved to v3.3 below.

## v3.0 — The engine-interface release (shipped 2026-08-02)

Delivered in three stages plus the two-artifact split, in dependency order, all
under one version number. Every stage has cleared its gate; the scope decision
and the stage table live in
[ROADMAP.md](ROADMAP.md) and
[DECISIONS.md § v3.0 shipped in stages](DECISIONS.md#v30-shipped-in-stages-the-layout-break-lands-last).

### Stage 1 — constructor pipeline + intercept relocation — BUILT AND VERIFIED

The code is written, the headless layers pass, and the spec-driven Excel gate
has reported **0 mismatches across all 12 QC cases** on a developer machine —
the standard the v2.1 row uses. Stage 1 shipped merged as #148.

- DONE: Stage 1 — the constructor pipeline + intercept relocation
  landed (merged as #148): the `X_s()`/`X_s_Within()`/`y_s()` →
  `Predictor_Columns()`/`Design_Columns()`/`Design_Response()` rename with
  the `encode → transform → demean → intercept → weight` order made
  explicit; the intercept moved into the constructor (`Design_Matrix` no
  longer synthesizes it, LINEST runs `const = FALSE` at all three sites,
  `SS_Total` became the intercept-only residual SS, `[Allow_Intercept]`
  left all 48 signatures — 13 now carry `[Has_Intercept]` as an
  identifier); and the QC test-sheet writers render formulas from declared
  argument names (`make_test_sheet.build_call`). Design rationale in
  [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).

- DONE: **Spec-driven verifier passed on a machine with Excel** —
  `python build_production.py --verify --no-launch` reported 0 mismatches
  across all 12 QC cases. Stage 1 changes where the intercept is created,
  not what is fitted, so any mismatch would have been a bug rather than an
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

### Stage 2 — `Model_Context` — BUILT AND VERIFIED

The code is written, the headless layers pass, and the spec-driven Excel gate
below has been run and passed on a developer machine — the same standard stage 1
cleared. Stage 2 shipped merged as #150.

- DONE: Engine — the 32 carriers (13 `[Has_Intercept]`-only + 19
  `[DF_Absorbed]`-only + 5 dual) collapsed `[Has_Intercept]` and
  `[DF_Absorbed]` into a single trailing `[Context]` argument. Each
  carrier's LET binds `context_arg, IF(ISOMITTED(Context), Model_Context(),
  Context)` once and reads `has_arg, Context_Has_Intercept(context_arg)` /
  `absorbed_arg, Context_DF_Absorbed(context_arg)` — never a bare
  `INDEX(context_arg, N)`, so the context row order is a contract enforced
  in one place (the accessors), not 32 hard-coded positional indices. The
  omitted-`[Context]` default routes through the `Model_Context()`
  constructor (one definition of the default), not a per-carrier inline
  `VSTACK`. Inter-carrier calls drop the two tokens and append
  `context_arg`. The workbook `Model_Context` constructor (category
  *Model Construction*, subcategory *Context Constructor*, scope defaults
  *workbook*) builds the 4×1 array `[Has_Intercept, DF_Absorbed,
  Response_Transform, Predictor_Transform]`. `Design_Columns` (the only
  Regression-scoped closure that reads the context) reads
  `Context_Has_Intercept(Fit_Context())=1` (was
  `INDEX(Model_Context(),1)=1`, which after the unshadowing below would
  have resolved to the always-TRUE constructor — a correctness fix that is
  a direct consequence of the two-name split). Pinned by
  `test_intercept_relocation` and `test_df_absorbed_threading`.

- DONE: Accessors — four workbook-scoped one-line functions
  `Context_Has_Intercept` / `Context_DF_Absorbed` /
  `Context_Response_Transform` / `Context_Predictor_Transform`, each
  `=LAMBDA(Context, INDEX(Context, N))` for N = 1..4 (category *Model
  Construction*, subcategory *Context Accessor*). Every context read goes
  through them; a future row insertion changes one accessor, not 32
  indices. The catalog is now 131 functions. Pinned by
  `test_context_accessors_index_rows_one_through_four` and
  `test_no_carrier_reads_the_context_with_a_bare_positional_index`.

- DONE: Sheet — the two-name split. `Model_Context(...)` stays the
  workbook-scoped constructor (the omitted-`[Context]` default and the MLR
  test-sheet path); `Fit_Context()` is the SHEET-scoped reader — a zero-arg
  thunk over the FIXED range holding the materialized context block. The ~30
  Regression sheet call sites pass `Fit_Context()` so they read the actual
  spec-derived context, not the constructor default. Splitting the names
  keeps `Model_Context` unshadowed: a single sheet-scoped thunk named
  `Model_Context` would make `Model_Context()` in a sheet cell resolve to
  the materialized values while the same token in a carrier's
  omitted-default resolved to the workbook constructor — the invisible
  shadowing the v3.0 release exists to remove. All four rows are
  materialized: elements 1-2 (the C2 `Allow_Intercept` toggle and the
  `Absorbed_Degrees_Of_Freedom()` closure) feed today's engines; elements
  3-4 (the response transform, and the None/Log/Mixed summary over the
  included Continuous predictors) have no engine reader until the v3.3
  unit-space dispatcher but land now so the row order is fixed. No `#`
  inside the LAMBDA `RefersTo` — the height is a structural constant, so a
  fixed range sidesteps the dynamic-array-in-a-name question entirely. An
  error in an unconsumed row 3/4 is contained (the engines read only
  elements 1-2 through the accessors). Pinned by
  `test_materialization_zone_materializes_model_context`.
  *Amended in v3.0 polish:* the block is written as one labelled cell per
  element rather than a single `VSTACK` spill — a spill sizes output to
  data and this height is a build-time constant, so it gained nothing while
  making the four cells one dependency node that Excel vacates and
  re-spills on every spec-block edit (blanking the range behind
  `Fit_Context()` mid-recalc). `_MODEL_CONTEXT_ELEMENTS` is now the single
  source of the row order, labels, and height; the old
  `=ROWS(Fit_Context())=_MODEL_CONTEXT_ROWS` guard cell becomes a
  `Context OK` row that checks the height *and* that no element errored.
  See [DECISIONS.md § v3.0](DECISIONS.md#the-model-context-is-individual-cells-not-a-vstack-spill).

- DONE: QC harness — all three MLR test-sheet writers
  (`write_sheet_mlr_{scalar,observation,vector_outputs}_test.py`) thread
  `context` through their `build_call` `reference_map` via the workbook
  `Model_Context` constructor (the scalar sheet carries the per-row
  `[@[Has_Intercept]]`; the observation/vector sheets carry a per-section
  `TRUE`/`FALSE` literal). Without this, `build_call` raises `KeyError` on
  the carriers' new trailing `[Context]` argument — a regression the headless
  suite did not catch (it only exercised non-carrier functions). Pinned by
  the `test_*_formula_threads_context_*` tests.

- DONE: Numeric + contract coverage — the relocated chain reproduces the
  pre-relocation numbers through the context-accessor path (200 datasets,
  both intercept states — packaging the flag into a 4×1 context and
  reading it via `INDEX(Context,1)` moves nothing); and the FE correction
  matches LSDV when `DF_Absorbed` is routed through element 2 of the same
  context array. Plus the contract assertions: only `Model_Context`
  declares `Has_Intercept`/`DF_Absorbed`; no formula anywhere contains a
  bare `INDEX(context_arg,`; `ROWS(Model_Context())` is 4. Pinned by
  `test_the_relocated_chain_is_unchanged_through_the_context_accessor_path`,
  `test_df_absorbed_routed_through_a_context_array_still_matches_lsdv`,
  `test_only_model_context_declares_df_absorbed`, and
  `test_model_context_constructor_is_a_four_row_vstack`.

- DEFERRED: Promote `Sample_Include()` from a live closure to a thunk over a
  materialized spill. The column sits at its final §4b position — as a
  RESERVED placeholder — with the Constructed Design Matrix zone behind it
  since stage 3. The thunk materialization needs
  the dynamic-array spill operator (`#`) inside a `LAMBDA` defined-name
  `RefersTo`, a combination not used anywhere in this workbook and only
  verifiable with Excel present. A wrong guess would break the row-mask
  contract that keeps every spilled array row-aligned, so it lands as a
  separate Excel-verified follow-up, not blind. The live closure is
  untouched and remains the row mask until then.

- DONE: **Ran the spec-driven verifier on a machine with Excel** —
  `python build_production.py --verify --no-launch` passed. Stage 2 had to
  report **0 mismatches across all 12 QC cases**: the collapse is behaviour-
  preserving (`context_arg` carries exactly the two scalars the dropped
  arguments carried), so any mismatch would have been a bug, not an expected
  delta. Not runnable in CI (no Office on the GitHub-hosted runner) — see
  [CONTRIBUTING.md](CONTRIBUTING.md) → *Verifying builds*.

### Stage 3 — layout — DONE

- DONE: Spec columns **M (Interaction Term)** and **N (Interaction
  Operation)**, appended rather than inserted so A–L keep both their
  addresses and their meanings and a saved spec survives the upgrade.
  Reserved-and-unwired: M's dropdown sources the variable-name spill
  (`=$A$4#`, which Excel stores as `ANCHORARRAY`, so the list resizes with
  the dataset and needs no volatile OFFSET), N's is the closed
  `Product | Difference | Ratio` axis, and three conditional-format rules
  are live — red on M when the named operand is not a Predictor (or not a
  variable at all), amber on M when it IS a Predictor but is excluded (the
  marginality violation: flagged, never blocked), and red on N on a
  reciprocal declaration under a symmetric or antisymmetric operation
  (Ratio excluded, self-reference excluded — that is the documented
  quadratic term). The bands `Spec_Interaction_Term` /
  `Spec_Interaction_Operation` are read by no defined name and no cell
  formula. Pinned by `test_interaction_bands_are_declared_but_read_by_no_constructor`
  and `test_interaction_flags_key_on_the_named_operands_own_spec_row`.

- DONE: Spec column **O (Design Columns)** — the per-row audit. Mirrors
  `Predictor_Columns()`'s own iteration predicate, Continuous/Categorical
  split, reference normalization, and degenerate skip, so the two agree by
  construction rather than by coincidence; deliberately NOT derived from
  the K (Levels) display, which would make one display depend on another.
  Interactions are not counted while M/N are unwired — an audit that
  anticipated columns the constructor does not build would be reporting a
  matrix that does not exist. Pinned by
  `test_design_columns_audit_mirrors_the_constructors_own_skip_rules` and
  `test_design_columns_audit_is_read_only_by_the_width_guard`.

- DONE: The **pre-flight width guard** — total at O1 (Σ column O plus the
  intercept, i.e. exactly `COLUMNS(Design_Columns())`), status at M2. Hard
  error at `16,384 − (last_chart_column + 5)`, derived from the layout
  constants rather than hard-coded, so moving a zone moves the limit with
  it; soft warning at k = 200 constructed columns or 500,000 materialized
  cells, whichever trips first. Both read the SPEC, never
  `COLUMNS(Design_Columns())` — a matrix too wide to fit cannot be built in
  order to be measured, which is the failure the guard exists to prevent.
  All three branches evaluated against live Excel. Pinned by
  `test_width_guard_reads_the_spec_not_the_constructed_matrix`.

- DONE: The **Constructed Design Matrix zone**, terminating the §4b band,
  collapsed by default. Established as a RESERVED position — the same
  treatment `Sample_Include` got at stage 2, and for the same reason: the
  position, the collapse behaviour, and the guard are what a later release
  cannot add without moving columns a second time; the spill that fills it
  is a formula change against a column that already exists (v3.2). Pinned
  by `test_design_matrix_zone_ships_collapsed_and_the_others_expanded`.

- DONE: Every zone right of the spec block shifted three columns, and every
  hard-coded A1 address in a formula string went with it — the CF
  expressions, the OFFSET-based chart named ranges, the chart title
  formulas, the prediction-input band, and the `alpha` name are now BUILT
  from the `_C_*` constants via `_abs_ref` / `_band` and the `_A_*` anchors.
  Spelling those letters out by hand is what makes a column insertion a
  silent-wrong-answer bug rather than a build failure: the formula still
  parses, it just reads a different cell. `tools/inspect_regression_sheet.py`
  likewise now IMPORTS the column map instead of keeping its own copy "to
  match" it.

- DONE: Verified against live Excel — `python build_production.py --verify
  --no-launch` passed, and the built workbook's cached values confirm the
  audit column (16 predictor columns on the shipped spec, 17 with the
  intercept), the relocated feedback block, and the three §4b zone headers.

- NOT DONE (and not needed): promote `_write_audit_row` /
  `_write_filtered_zones` from `write_sheet_model_construction.py` into the
  Regression writer. The earlier plan assumed the audit column would be
  built by copying the Model Construction sheet's audit-strip machinery.
  It was not: the audit is a per-row calculated column INSIDE `SpecTable`,
  written by the shared `_write_spec_block`, so the Regression sheet
  inherits it by already calling that writer. The row-1 audit strip and the
  filtered-display zones remain Model-Construction-only, which is correct —
  the Regression sheet displays the same facts through its own zones.

## v3.1 — Interaction wiring — SHIPPED

Spec columns M/N are wired: `Predictor_Columns()` and its two twins build the
declared interaction columns, and the Design Columns audit counts them. See
[ROADMAP.md](ROADMAP.md#v31--interaction-wiring--shipped-2026-08-03) and
[DECISIONS.md § v3.1](DECISIONS.md#v31--interaction-wiring).

- DONE: `Predictor_Columns()` reads `Spec_Interaction_Term` /
  `Spec_Interaction_Operation` through a LET-bound `mate()` and combines the
  declaring row's block with the operand's — both from the same `blk()` — under
  the closed `Product | Difference | Ratio` vocabulary. `Ratio` returns `NA()`
  on a zero denominator.

- DONE: `Constructed_Column_Names()` and `Constructed_Column_Transforms()`
  extended in the same edit, gating interactions identically so the three stay
  width-identical. Interaction headers join the operands' own constructed names
  with the operation's own symbol (` × ` / ` − ` / ` ÷ `, ` ? ` for anything
  else); interaction columns always read Transform `None`.

- DONE: The column O audit gained its `k(row) × k(operand)` term off the same
  per-row width helper both operands use, so audit and constructor agree by
  construction.

- DONE: `test_interaction_bands_are_declared_but_read_by_no_constructor`
  retired and replaced by
  `test_interaction_bands_are_read_by_the_three_constructor_twins`, which pins
  that exactly those three read the bands — `Sample_Include()`, `Row_Labels()`,
  and `Response_Column()` must not.

- DONE: Three QC cases added to the spec-driven oracle covering all three width
  regimes, plus `tests/test_interaction_wiring.py` for the semantics.

- DONE: **Spec-driven verifier run on a machine with Excel.**
  `uv run python build_production.py --verify --no-launch` was run and the
  v3.1 interaction-wiring checks cleared for the Regression workbook scope.
  The Univariate-sheet-missing error in that output is fixed: `build_qc`'s
  Univariate stage fell through to the check when `skip_univariate=True` and
  the sheet was absent — the Regression artifact's normal state since the
  split — reporting `[Univariate] sheet is missing` as a QC failure and
  exiting 1 on a good workbook. `_univariate_verification_action` now returns
  skip / warn / check explicitly.

- DONE: **Workbook-scope cleanup after the v3.0 split.** `sync_workbook_names`
  now owns workbook scope outright — anything workbook-scoped that is not a
  catalog function or an `_xlnm.*` reserved name is dropped on every build —
  and skips a catalog definition whose body names a worksheet the target
  artifact lacks. That clears the split's cross-artifact residue (twelve
  `RegChart*` `OFFSET(#REF!,…)` entries in the Univariate workbook,
  forty-two `UV_*` entries plus twenty-one retired LAMBDA names in the
  Regression workbook) and the broken external link `Base_Period_Delta`
  produced in the Univariate artifact. Both committed artifacts were cleaned
  in place with `tools/resync_workbook_names.py` (name manager only — the
  rebuild below is still outstanding), and
  `TestRealWorkbookNameScope` guards it in CI. See
  [DECISIONS.md § v3.1](DECISIONS.md#workbook-scope-belongs-to-the-catalog).

- TODO: **Rebuild and commit both artifacts.** The 2.1.0 / 2.2.0 / 3.1.0
  Version History entries and the new constructor formulas only reach users
  through a rebuild.

## v3.2 — Full materialization of the design matrix

The other stage-3 follow-on. The terminal Constructed Design Matrix zone, its
collapse behaviour, and its width guard all shipped; what is missing is the
spill that fills it.

- TODO: Materialize `Design_Columns()` into the reserved terminal zone.
  Position, collapse state, and guard are already established, so this is a
  formula change against a column that exists — the same reserved-position
  treatment `Sample_Include` got. See
  [ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone).

- TODO: Promote `Sample_Include()` from a live closure to a thunk over a
  materialized spill — deferred out of v3.0 stage 2 (see the DEFERRED entry
  in the stage 2 list above). Needs the dynamic-array spill operator (`#`)
  inside a `LAMBDA` defined-name `RefersTo`, a combination used nowhere else
  in this workbook and verifiable only with Excel present. A wrong guess
  breaks the row-mask contract that keeps every spilled array row-aligned,
  so it lands Excel-verified, not blind.

## v3.3 — Transforms remainder

Planned as the second half of v2.2; moved after v3.0 when the feature train
was resequenced — see [ROADMAP.md](ROADMAP.md). The column-G `Log` wiring
already shipped at v2.2.

- TODO: **Unit-space dispatcher function, RESOLVED:**
  `Unit_Space_R_Squared(model, response_transform, predictor_transform)`,
  `Unit_Space_Adjusted_R_Squared(...)`, `Unit_Space_RMSE(...)`. The
  dispatcher pattern (one canonical name per statistic, internal
  `SWITCH` on the transform pair) is the documented
  naming-style-departure pattern, in
  [DECISIONS.md § v2.2 unit-space dispatcher](DECISIONS.md#v22--transforms--unit-space-comparability)
  and [ARCHITECTURE.md § 1 "Naming-style departures"](ARCHITECTURE.md#1-naming-convention).

- TODO: Unit-space section on the Regression sheet — SWITCH on column
  G, one headline comparable statistic (the cell v3.4 Model Comparison
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
  `tests/test_difference_by_verification.py`, which carries the T17–T19
  cases from the since-retired human test plan.
  The shipped semantics and the `NA()` exception are in
  [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy)
  and
  [DECISIONS.md § v2.1 base-period layer](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

- TODO: Sample construction — `Numeric_Complete_Cases`.

- TODO: Categorical & model construction — `Dummy_Column`, `Interact`,
  `Model_Matrix`.

## v3.4 — Model Comparison Sheet

*Planned as v2.3; moved after v3.0 when the feature train was resequenced —
see [ROADMAP.md](ROADMAP.md).*

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

## v3.5 — Resampling & Simulation

*Planned as v2.4; moved after v3.0 when the feature train was resequenced —
see [ROADMAP.md](ROADMAP.md).*

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

## v3.6+ — Future (sequence TBD; first two claimed)

*Planned as v2.5+; moved after v3.0 when the feature train was resequenced —
see [ROADMAP.md](ROADMAP.md).*

This bucket previously had seven candidates with no order.
Two-sample tests are now v3.6 (next MINOR after v3.5) and the `Weight`
Role is v3.7 (after v3.6). The rest are deliberately unordered pending
actual user demand — a single maintainer should not pre-order work that
may not be the next thing actually needed.

### v3.6 — Bivariate / Two-sample *(claimed, next MINOR after v3.5)*

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

### v3.7 — `Weight` Role (WLS) *(claimed, after v3.6)*

The standalone WLS milestone and its `[weights]`-argument-vs-parallel-
function-set debate are superseded by a **`Weight` value on the Role
axis** (see
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)).
Three-stage scope carried forward: user-supplied weights →
variance-driver-derived weights → FGLS. This milestone ships the first stage
only. The default-uniform → OLS pattern (the
"non-breaking MINOR" guarantee) is in
[DECISIONS.md § v2.6 WLS](DECISIONS.md#v26--wls-weight-role-default-uniform-weights-argument).

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

### v3.8+ — Unordered candidates (no claim)

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

---

## Documentation

Version-independent items. Not tied to a milestone; both stay open until done.

- DONE: **The missing 2.1.0 and 2.2.0 Version History entries are written.**
  `_VERSIONS` in `lambda_catalog/write_sheet_version_history.py` ran
  1.0.0 → 1.1.0 → 1.2.0 → 2.0.0 → 3.0.0, so Fixed Effects, the Sequence
  axis, the gap-aware longitudinal layer, GVIF, and the column-G Log
  transform all reached users with nothing in the workbook's own changelog
  describing them. Both entries are non-breaking and both are
  Regression-workbook events. **They only reach users on a rebuild** — see
  the rebuild item under v3.1.

- TODO: **Build one of the two mechanical drift checks** proposed in
  [CONTRIBUTING.md § Documentation drift](CONTRIBUTING.md#documentation-drift-proposed-check--not-yet-implemented)
  and tracked as [REVIEW.md F7](REVIEW.md#f7--documentation-drift-is-measurable),
  the one finding still open. The cheaper and higher-yield of the two is the
  **cross-document anchor check**: every `](FILE.md#anchor)` and `](#anchor)`
  resolves to a real heading in the target file. Roughly 40 lines of `re`
  plus a pytest case, no Excel, runs in the existing Linux CI job. The
  2026-08-03 review ran it by hand (zero unresolved targets) and then
  deleted three documents — exactly the change that breaks anchors with no
  error anywhere. The second check (function names in docs resolving
  against `lambda_functions.json`) would have caught the stale `X_s`
  references and the 126-vs-131 count drift that same review found by hand.

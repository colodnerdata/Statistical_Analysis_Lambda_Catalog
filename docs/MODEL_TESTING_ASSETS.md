# Model Testing Assets

The plan of record for the **regression test-model suite** — the model configurations the QC
harness covers, the datasets they need, and the order in which future features force the suite to
grow.

Every model in Section 1 now has an oracle. Fittable models are `RegressionSpecCase` entries in
`lambda_catalog/analyze_regression_spec.py` (`SpecVariable` rows built with `_spec_var(...)`,
per-case `source_table_ref` retargeting of `Source_Table`, statsmodels/NumPy expectations via
`calculate_regression_spec_case`); the § 1.4 guard-rail configurations are `GuardStateCase`
entries in `lambda_catalog/analyze_regression_guard_states.py`, which assert status text, the
Design Columns audit and CF predicates rather than fit statistics. Each case is also materialized
as its own worksheet — see [Section 1b](#section-1b--one-worksheet-per-test-model).

**Where this sits in the documentation.** [CONTRIBUTING.md](../CONTRIBUTING.md#the-regression-test-model-suite)
describes how a case is added and verified; this file decides *which* cases exist and *why*.
Section 2's ordering is the source of the [ROADMAP.md](ROADMAP.md#versioning--release-conventions)
version ladder from v3.4 onward — the ladder is sequenced by the test-suite growth each milestone
forces, so reordering a milestone means reordering Section 2 first.

**Coverage philosophy.** The suite is a covering array, not a full factorial: every implemented
corner case is exercised by at least one model, and every model earns its place by covering
something no other model does. Target size: ~25–30 fittable models plus ~10 guard-state
configurations. Full crosses (every transform × every interaction × every role…) are explicitly
out of scope — they blow up sheet count without adding information.

Status column legend: **existing** = already a QC case (name given, pinned in
`tests/test_regression_spec_qc.py`); **new** = not yet declared anywhere.

---

## Section 1 — Test models for implemented features

Notation: `C(x)` = Categorical, `Ln(x)` = `Transform = Log`, `| G` = Fixed Effects on G,
`×` / `−` / `÷` = interaction operations, `(resp, pred)` = the transform dispatch pair from
`Fit_Context()`. Intercept is ON unless noted.

Section 1 is ordered by **what the case demonstrates**, from lowest modelling complexity to
highest. Dataset names appear only where they matter to the demonstration; the ordering does not
preserve legacy Auto MPG / Life Expectancy / Production Lots blocks or old tab aliases.

### 1.1 Fittable model cases, ordered by demonstration complexity

| ID | Demonstration | Model / configuration | Covers | Case |
|---|---|---|---|---|
| M01 | Baseline T0 | `MPG ~ Horsepower + Weight + C(Model Year) + C(Origin)` | shipped baseline; categoricals + continuous; fit/unit-space identity | `baseline_t0_intercept` |
| M02 | All continuous | `MPG ~` all five continuous predictors | plain OLS feature-order/full-data filter | `all_continuous_intercept` |
| M03 | Continuous subset | curated continuous subset | `Include = FALSE` candidate rows | `continuous_subset_intercept` |
| M04 | All continuous, no intercept | M02 with intercept OFF | no-intercept continuous OLS | `all_continuous_no_intercept` |
| M05 | Continuous subset, no intercept | M03 with intercept OFF | no-intercept plus excluded candidates | `continuous_subset_no_intercept` |
| M06 | Baseline categoricals, no intercept | M01 with intercept OFF | red advisory for intercept-off categorical models | `baseline_t0_no_intercept` |
| M07 | Default categorical reference | Origin reference blank | first-in-sort-order reference | `origin_default_reference` |
| M08 | Explicit categorical reference | Origin reference = `Europe` | typed reference override | `origin_explicit_reference` |
| M09 | Invalid categorical reference | Origin reference = `99` | invalid reference degrades to zero columns but remains fittable | `origin_invalid_reference` |
| M10 | Binary categorical reference | Life Status reference = `Developing` | binary categorical retained dummy | `life_status_reference` |
| M11 | Mixed categoricals | `MPG ~ Displacement + Horsepower + Weight + C(Model Year) + C(Origin)` | multi-level categorical alongside continuous predictors | `categorical_mixed_predictors` |
| M12 | Categorical-only design | `MPG ~ C(Model Year) + C(Origin)` | categorical design with no continuous predictors | `categorical_only_design` |
| M13 | Filter-degenerate categorical | M01 + `Is_USA` filter | filter collapses Origin to one sampled level | `filter_degenerate_categorical` |
| M14 | Mileage log-log missingness | `Ln(MPG) ~ Ln(Weight) + Ln(Horsepower)` | (Log, Log) and NA masking on Auto MPG | `mileage_log_log_missingness` |
| M15 | Linear-log mixed | `Life expectancy ~ Ln(Population) + Ln(GDP) + Alcohol + C(Status)` | (None, Mixed) transform dispatch; heavy missingness | `life_linear_log_mixed` |
| M16 | Log response, Duan | `Ln(Life expectancy) ~ Adult Mortality + Schooling + C(Status)` | smearing factor/original-units outputs | `life_log_response_duan` |
| M17 | Log response, Naive | M16 with Back-Transform = Naive | naive back-transform branch | `life_log_response_naive` |
| M18 | Elasticity log-log | `Ln(Life expectancy) ~ Ln(GDP) + Ln(Population)` | large-sample (Log, Log) | `life_elasticity_log_log` |
| M19 | Power law from derived logs | `log Unit Cost ~ log Cum Units` | pre-derived log-column route, no FE | `production_power_law_derived_no_fe` |
| M20 | Power law from transform axis | `Ln(Unit_Cost_BY) ~ Ln(Cumulative_Units)` | transform-axis route, no FE; M19 twin | `production_power_law_transform_no_fe` |
| M21 | Log response with mixed predictors | Production Lots mixed logged/unlogged predictors | (Log, Mixed) branch | `production_log_mixed_predictors` |
| M22 | Log predictor only | `Unit_Cost_BY ~ Ln(Cumulative_Units)` | (None, Log) branch | `production_log_predictor_only` |
| M23 | Continuous product | continuous × continuous `Product` | one-column product interaction | `interaction_continuous_product` |
| M24 | Quadratic self-product | Horsepower × Horsepower | self-product/x² declaration | `interaction_quadratic_self_product` |
| M25 | Continuous × categorical | Weight × C(Origin) | L−1 broadcast interaction columns | `interaction_continuous_by_categorical` |
| M26 | Categorical × categorical | C(Model Year) × C(Origin) | full-product width regime | `interaction_categorical_cross` |
| M27 | Difference interaction | Displacement − Acceleration | antisymmetric `Difference`; U+2212 header | `interaction_difference` |
| M28 | Ratio reciprocal pair | Weight ÷ Horsepower and Horsepower ÷ Weight | asymmetric ratio pair; zero-denominator path | `interaction_ratio_reciprocal` |
| M29 | Fixed effects from derived logs | `log Unit Cost ~ log Cum Units | Facility` | FE, filter, group prediction, typed annual sequence period | `production_fixed_effects_derived` |
| M30 | Fixed effects from transform axis | `Ln(Unit_Cost_BY) ~ Ln(Cumulative_Units) | Facility` | FE transform-axis twin of M29 | `production_fixed_effects_transform` |
| M31 | LSDV equivalence | M30 with Facility as categorical predictor | LSDV ↔ within-estimator slope/residual equivalence | `production_lsdv_equivalence` |
| M32 | High-cardinality FE | `Life expectancy ~ Schooling + Adult Mortality | Country` | 173 surviving FE groups; panel verdicts at scale | `life_country_fixed_effects` (**heavy**) |
| M33 | Numerical stress profile | Life Expectancy kitchen-sink profile | k=19 ill-conditioned precision floor | `life_full_profile` (**heavy**) |

### 1.2 Guard-rail / error-state configurations

Not fittable models — these verify status lines, conditional formatting, model-formula degradation,
Design Columns audits, and graceful refusal paths. They remain after the fittable cases and are
ordered from the simplest validation failure to the highest structural complexity.

| ID | Configuration | Expected behavior | Case |
|---|---|---|---|
| G01 | zero `Response (y)` rows | response readout degrades to `"(none)"`; sample grows because the mask loses its response term | `guard_no_response` |
| G01b | no included `Predictor (x)` rows | `Predictor_Columns` errors; consumers degrade to `"(empty model)"` | `guard_empty_model` |
| G02 | two `Response (y)` rows | response count flags red; `XMATCH` still resolves the first | `guard_two_responses` |
| G03 | two `Sequence = TRUE` flags | E1/H2 multiple-Sequence error and red H cells | `guard_two_sequence_flags` |
| G04 | two `Fixed Effects` rows | B1 cardinality error and red CF | `guard_two_fixed_effects` |
| G05 | Fixed Effects row + Intercept `TRUE` | red intercept-toggle advisory for a demeaned design | `guard_fixed_effects_with_intercept` |
| G06 | Intercept `FALSE` + included categorical | red intercept-toggle advisory; model still fits elsewhere as M06 | `guard_intercept_off_with_categorical` |
| G07 | `Transform = Log` on a Categorical predictor | red transform cell; dummies encode without silently applying `Ln` | `guard_log_on_categorical` |
| G08 | Interaction Term names a non-Predictor row | red operand cell; row contributes main effect only | `guard_interaction_bad_operand` |
| G09 | reciprocal `Product` declarations | red symmetric-operation flags on both rows | `guard_reciprocal_product` |
| G10 | interaction operand has `Include = FALSE` | amber marginality warning; columns still build | `guard_excluded_operand` |
| G11 | unrecognized Interaction Operation pasted past validation | visible `" ? "` header and `NA()` design column | `guard_unknown_interaction_operation` |
| G12 | `Ln(Schooling)` where Schooling contains true zeros | zero rows stay in the mask and `Ln_Positive` propagates `#N/A` | `guard_ln_zero_propagation` |
| G13 | Life Expectancy width soft warning (k = 205) | M2 WARNING and visible engine degradation instead of plausible wrong numbers | `guard_width_guard_warning` |
| G14 | spec block built narrow then retargeted wider | `Spec_*` bands/input fill resize to the wider `Source_Table` | `guard_spec_block_retarget_widens` |
| G15 | typed Sequence Period override on Auto MPG | Period In Use follows typed Δ = 2 and verdict re-evaluates against it | `guard_sequence_period_override` |
| G16 | irregular Production Lots panel spacing | gapped facilities under typed Δ = 1 produce the yellow Regularity verdict | `guard_irregular_panel_spacing` |

**Auto MPG carries no Sequence axis.** The dataset is cross-sectional: every fittable Auto MPG
case leaves `Sequence` blank. Only the guard mechanics cases (`guard_two_sequence_flags` and
`guard_sequence_period_override`) deliberately set Sequence flags on Auto MPG, because they test
status-cell mechanics rather than panel diagnostics. Production Lots and Life Expectancy carry the
substantive panel/sequence coverage.

**Transform pairs remain paired by concept even when they use different datasets.** M19/M20 and
M29/M30 are the pre-derived-column vs. transform-axis twins; M16/M17 are the Duan/Naive
back-transform twins. Their adjacency is intentional and replaces the old source-dataset grouping.

## Section 1b — One worksheet per test model

Every case in Section 1 is materialized as its own Regression-shaped worksheet in
**`Lambda_Library_TestModels.xlsx`**, built by `build_test_models.py`. The workbook
is gitignored: it is a QC fixture regenerated from the case registries on demand,
not a shipped artifact.

**Why it exists.** The original harness pushes each case through the single
`Regression` sheet in turn. That works, but a case then exists only as a log line
— a failure says "expected 0.79596, got 0.79601" with nothing to open — and every
case has to defensively re-set every input in case the previous one left something
behind (which is why `source_table_ref` and `prediction_group` are non-optional and
rewritten on every iteration). With one sheet per case the verifier only **reads**:
no writing, no per-case recalculation, no state to leak, and a failing case is a
tab you can open.

**Sheet names state the concept, not the variables.** `M14 Log Log Missingness`,
never `MPG ~ Ln(Weight) + Ln(HP)`. Excel allows 31 characters, which cannot hold a
model formula, and the formula is the least interesting thing about a test case
anyway — the sheet exists to exercise one corner, and that corner is what the tab
should say. The variables are one click away in the spec block. The `<PlanID>
<Concept>` shape ties each tab back to a row in this document: `M05` is this
document's M5, `G03` its G3.

The contract lives in `lambda_catalog/test_model_sheets.py` and is enforced at
registry-build time, so an illegal or duplicated name fails in a millisecond-long
unit test rather than partway through a multi-minute Excel build:

| Rule | Enforced by |
|---|---|
| 1–31 chars, none of `[ ] : * ? / \`, no leading/trailing apostrophe or space, not `History` | `validate_sheet_name` |
| `<PlanID> <Concept>`, plan ID matching `[MLPG]\d\d[a-z]?` | `validate_sheet_name` |
| Unique across model **and** guard cases, case-insensitively | `assert_sheet_names_unique` |
| One `SpecTable_<PlanID>` ListObject per sheet | `spec_table_name` |

That last row is not a style choice. Excel ListObject names are **workbook**-scoped,
so a second sheet naming its table `SpecTable` is an error from `ListObjects.Add`,
not a silent rename — `spec_table_name` is threaded through `_write_spec_block` →
`_create_spec_table` and `_set_sheet_scoped_names` so the table and the `Spec_*`
band names that bind to it can never disagree.

**What the writer reuses.** Nothing reimplements the Regression sheet.
`write_regression_output_sheet` gained three defaulted parameters — `sheet_name`,
`include_charts`, `spec_table_name` — so the production build is unchanged, and
`lambda_catalog/write_sheet_test_model.py` calls it with a per-case identity. The
spec is then applied through `lambda_catalog/regression_spec_sheet_io.py`, which
is also what the legacy single-sheet verifier uses: if the builder and the verifier
disagreed about what a case *is*, the sheets would be verifying something other
than what the QC harness fits.

Charts are **off** on generated sheets. Roughly a dozen COM chart objects per sheet
across ~48 sheets is the single largest cost in the build, and no oracle reads one;
chart wiring is verified once, on the production Regression sheet.

**Two cases are opt-in.** M33 (Numeric Stress Profile, k = 19, n = 2117) and M32
(173 Fixed Effects groups) carry `heavy=True` and are skipped unless
`--include-heavy` or an explicit `--cases M33` / `--cases M32` is given. M32
is gated on sheet-build cost; M33 is gated on the statsmodels-vs-Excel
floating-point floor at fdd = 5/6 that both implementations agree on (it
lives here as a deliberate showcase for the floor, not as a defect). Their
Python oracles always run in the unit suite — only the sheet build is gated.
G13 was the third candidate until the live run showed the workbook cannot
fit a 205-column design at all — it is a guard state now, and guard sheets
are cheap.

```
python scripts/build_test_models.py                        # 48 sheets (31 non-heavy models + 17 guards)
python scripts/build_test_models.py --include-heavy        # 50, adding M33 and M32
python scripts/build_test_models.py --cases M26,G09        # just those two
python scripts/build_test_models.py --verify --no-launch   # build, check, exit 1 on drift
poe verify-test-models                             # the same, verbose
```

**Every run archives its own transcript** to `excel-only-runs/<script> <flags>.log`
— stdout *and* stderr, flushed per line, with the traceback written into the file
before the streams are restored. That directory is gitignored rather than
committed: this check needs Excel, so the transcript only exists on the machine
that ran it. The directory name reflects what produced it (Excel-required
verifier builds — the GitHub-hosted Linux CI cannot run them) rather than where
(a developer's box). `--log PATH` overrides the destination.

**`--verbose` names each sheet before writing it**, not after. A ~46-sheet run
through COM takes minutes, and the two questions a watcher has are "is it still
moving?" and "which sheet is it stuck on?" — both need the current sheet on screen,
flushed, while the work is happening. An interrupted run therefore leaves the
offending case named with no duration after it. The verify pass reports per-sheet
mismatch counts and prints **per-case totals before the failure list**: the first
live run of the sibling verifier buried 12 real mismatches under 22,886 from a
single case, and that summary line is what would have said so on sight.

### Three bugs the live runs found

None was reachable headlessly, which is the argument for running the deep check
before merging anything that touches the writers or wires a new dataset. Each is
now pinned by a unit test that fails against the old code.

**Sheet names with spaces were never quoted.** `_setup_local_names` built
`=M01 Baseline Categoricals!$AB$12`, which is not a valid formula, and Excel
rejected the whole `Names.Add`. Four of that function's seven references were
unquoted — invisible for the life of the project because the only sheet it ever
wrote was named `Regression`, a single word. `sname` now carries the quotes, so no
call site can forget, and a test builds every name-registering writer against a
spaced sheet name.

**A fixture column widened the source table past the spec block.** `Is_USA`
exists for M13, which declares it as a Filter. The legacy verifier adds it and
deletes it again around that one case; this workbook keeps it, which makes it
part of *every* Auto MPG case's source table. Every constructor opens
`n_c = COLUMNS(Source_Data)` and then indexes the `Spec_*` bands at `1..n_c`, so
a 12-row spec block against a 13-column table makes `INDEX(rl, n_c)` run off the
end — the row mask errors and every engine cell downstream reads as an error.
The build succeeds, the sheet looks right, and it surfaces only as a verify run
where an entire case reads `None`. M13 was the one Auto MPG sheet that worked,
because its spec happens to declare the thirteenth column.

The rule is now explicit and enforced: **a spec block must have exactly one row
per Source_Table column.** `write_sheet_test_model.pad_spec_to_source_table`
appends an `Omit` row for every source column a case does not name — `Omit`
contributes no design column and imposes no mask condition, so a padded spec
fits exactly the same model, which is why the Python oracle can stay ignorant of
fixture columns entirely. The fixture list is declared once, in
`FIXTURE_COLUMNS`, and read by both the data-sheet writer and the spec block:
two copies of it is what let the table grow a column the spec block did not know
about.

**Categorical levels sorted by code point, not by collation.** `Côte d'Ivoire`
files after `Czechia` in Python (`ô` = U+00F4 > `z`) but between `Costa Rica` and
`Croatia` in Excel. On a categorical predictor that one-position difference shifts
the *whole* dummy block: every column keeps a valid name and a valid 0/1 pattern,
nothing errors, and every per-predictor statistic is silently paired with the wrong
header. `level_sort_key` in `analyze_model_construction.py` now strips combining
marks and casefolds before comparing, which reproduces the locale order for Latin
scripts. It is an **approximation** — real ICU collation has per-locale rules it
does not model (Scandinavian `å` after `z`, Hungarian digraphs, non-Latin scripts)
— and exact parity would need PyICU. If a dataset with such levels is ever wired
in, that is the first place to look when a categorical's statistics come out
permuted.

Verification needs Excel and so does not run in CI, exactly like the other two deep
checks. Everything else about the framework — the naming contract, both registries'
coverage, the per-sheet parameterization, and every oracle in Section 1 — is
asserted headlessly by `tests/test_test_model_sheets.py`,
`tests/test_regression_guard_states.py` and `tests/test_regression_spec_qc.py`.

---

## Section 2 — Assets for roadmap features, in ladder order

**Ordering principle — two keys.** The ladder from v3.4 on sorts first by **track**, then by
**test-scale multiplier**:

1. **Regression work ships first.** Every milestone that extends the Regression sheet, its spec
   block, or its engine comes before either milestone that opens a *new* analysis surface. The
   two exceptions to the Regression track — Two-sample and Resampling — go last as a block, even
   though both are flat-cost to test.
2. **Within the Regression track, cheapest-to-cover first.** Additive features first (each adds a
   fixed number of cases), per-model multipliers next, axis-widening multipliers last — so the
   suite grows linearly for as long as possible and the biggest lift lands when the harness is at
   its most mature. Within a tier, the most valuable / most commonly used feature goes first.

Key 1 outranks key 2 deliberately. Test cost is the right tiebreaker inside one artifact and the
wrong primary key across two: a Regression milestone is a spec column, an engine change, and more
cases in the oracle that already exists, while Two-sample and Resampling each need a new sheet
writer, a new layout, and a verification path sharing nothing with
`calculate_regression_spec_case`. Interleaving them means two half-built analysis surfaces at
once, and it leaves the artifact users actually have feature-incomplete for longer. The deferral
costs no rework — neither milestone depends on any Regression milestone, and none depends on them.

**This ordering is the roadmap's ordering.** The `Ships as` column below is the version ladder
from v3.4 on: ROADMAP.md was renumbered to follow this table rather than the other way round.
See [ROADMAP.md § Ladder order](ROADMAP.md#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
and [DECISIONS.md § ladder ordering](DECISIONS.md#the-post-v33-ladder-regression-work-first-then-test-suite-growth).

Two framing notes:

- One multiplicative feature — the Log transform and its response × predictor dispatch pairs — is
  *already integrated*; its cost is absorbed into Section 1's covering array. That is a fact about
  the current suite, **not a constraint on the ordering logic**: the scale effects below score each
  future feature purely by the marginal growth it forces, independent of what happens to be in
  already.
- The tool is single-user and pre-release; nothing here is frozen. Existing QC cases, dataset
  profiles, and this ordering itself can be backtracked and reshaped as iteration reveals better
  structure. The order is a rework-minimizing default, not a commitment.

### The Regression track

| # | Roadmap item | Ships as | Scale effect | Test assets needed |
|---|---|---|---|---|
| 1 | **Model Comparison sheet** | **v3.4** *(unchanged)* | additive (~1×) — reads existing models | ≥3 registered models with shared prediction inputs — Section 1 already supplies them (e.g. M1, L2, P2). Add **one mismatched-predictor-set pair** (e.g. M1 vs M14) to exercise the `XLOOKUP [if_not_found]` open question. |
| 2 | **`Cluster` role** | **v3.5** *(was unordered v3.8+)* | near-additive — a variance-estimator variant on a few models | Within-group correlated data: Production Lots facilities suffice initially (3 clusters — deliberately few, to test the small-cluster warning path); `Grunfeld` (item 5) later provides 10–11 proper clusters. |
| 3 | **`Time` role / time series** | **v3.6** *(was unordered v3.8+)* | near-additive — **and unlocks a today-gap** | A real **calendar-dated monthly series** (~144 rows, AirPassengers-shaped, with an actual date column). No wired dataset has dates; this asset also enables the Sequence **calendar-signature verdict** test in Section 1 immediately, before the Time role ships. Also serves `Moving_Average` / `Exponential_Smoothing` cases. |
| 4 | **WLS `Weight` role** | **v3.7** *(unchanged number, new reasoning)* | ~2× over a representative subset | Grouped/heteroskedastic data with a natural weight column: R/MASS `Insurance` (64 rows, claims with exposure `Holders`) or a grouped-mean aggregation of an existing dataset. Plan **weighted variants of ~6 representative Section-1 models** (one per dispatch-pair family), not the whole suite. Include the recorded trap as an oracle assertion: `DEVSQ(√w ⊙ y)` ≠ weighted SST. |
| 5 | **Two-way Fixed Effects** | **v3.8** *(was unordered v3.8+)* | ~2× over the FE family | A balanced two-factor panel: R `Grunfeld` (200 rows, 10 firms × 20 years) plus an **unbalanced variant** (rows deleted) to exercise `Is_Balanced_Panel` and the convergence check. Re-run the FE family (P1/P2/L8 analogues) two-way. |
| 6 | **Standalone transform library** | **v3.9** *(was the v3.3 remainder)* | the **~10× axis-widener — last in the track** | Each new Transform value (`Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Zscore_By`, `Decompose_By`) widens the predictor-transform axis that currently holds {None, Log}, and every widening multiplies the response × predictor dispatch table (six recognized pairs today). No new data needed — existing datasets cover all of them. Sequencing *within* the item: (a) the additive helpers first (`Numeric_Complete_Cases`, `Dummy_Column`, `Interact`, `Model_Matrix` — standalone LAMBDAs, fixed test count); (b) predictor-side location/scale transforms next (each adds pairs but not back-transform semantics); (c) **any response-side extension last** — a response transform also multiplies the back-transformation / unit-space semantics (what is the smearing analogue for Zscore⁻¹?), which is the single most expensive kind of growth this project has. |

### Then the new analysis surfaces

| # | Roadmap item | Ships as | Scale effect | Test assets needed |
|---|---|---|---|---|
| 7 | **Two-sample / bivariate** | **v3.10** *(was v3.6, briefly v3.5)* | additive — fixed set of test cases | A small two-group dataset (R `ToothGrowth`, 60 rows — or the in-repo `Status` split of Life Expectancy) and a **paired** dataset (R `sleep`, 20 rows). Cases: equal-variance t, Welch t, paired t, F-test of variances feeding the selector cell. |
| 8 | **Resampling & simulation** | **v3.11** *(was v3.5, briefly v3.6)* | additive | No new data. The **seeded pre-drawn `Bootstrap_Random_Draws` table** is itself the asset; Production Lots (n = 51) is the natural small-n bootstrap target (slope CI on P3b). PERT/MC cases need only parameter cells. |

**Why the numbers moved.** `Cluster` and `Time` are pulled out of the unordered bucket ahead of
WLS because a variance-estimator variant on a handful of models is cheaper to cover than a
weighted re-run of one model per dispatch family, and `Time` additionally closes the one Section-1
coverage gap that exists today. The standalone transform library leaves v3.3 for the end of the
Regression track because it is the only item that widens an axis every other model is crossed
against; the v3.3 milestone keeps its number for what already shipped (the unit-space dispatcher,
Duan back-transformation, the model formula label), and only the unshipped remainder moves.
Two-sample and Resampling move to the end of the whole ladder under key 1, keeping their relative
order — both are flat-cost, and two-sample is the more common ask (the ToolPak parity gap a user
hits first).

**Note the deliberate inversion.** Items 7 and 8 are cheaper to test than items 2–6 and still ship
last. That is key 1 doing its job, and it is the one place in this document where test scale does
*not* decide the order — recorded here so a future reader does not "fix" it back.

Unscheduled long-tail items (multi-group means/ANOVA, Fourier, decision analysis) need no assets
beyond the above: ANOVA-as-regression is `warpbreaks` + the existing categorical machinery. They
stay in the unordered v3.12+ bucket, since nothing about their test cost sequences them.

---

## Section 3 — Supplemental datasets (kept minimal)

**Principle.** A new dataset must buy a corner the three wired datasets cannot; prefer CSVs already
in `sample_data/`; keep each addition ≤ ~250 rows. Everything recommended below totals **< ~600
rows** — negligible workbook weight. The real size driver is *sheet count per model*, which the
covering-array philosophy already bounds.

**Wiring cost (future work, per dataset):** one `CsvDatasetConfig` + one `SpecDatasetProfile` +
a registry entry (`lambda_catalog/write_sheet_csv_dataset.py`, `SPEC_DATASET_PROFILES` in
`write_sheet_model_construction.py`).

### Recommended additions

| Dataset | Rows | Buys | Needed by |
|---|---|---|---|
| R `warpbreaks` | 54 | balanced 3×2 **Cat×Cat factorial** (tension × wool) — the textbook ANOVA-as-regression fixture, cleaner than M9's unbalanced cross | Section 1 (optional M9 companion); future multi-group means |
| Calendar-dated monthly series (AirPassengers-shaped, real date column) | ~144 | **the only uncovered Section-1 axis** (calendar-signature Sequence verdict); later the Time role and time-series functions | §1.5 gap; **v3.6** `Time` role |
| R `Grunfeld` | 200 | balanced 10×20 two-factor panel; proper cluster count | **v3.8** two-way FE; back-fills **v3.5** `Cluster` |
| R/MASS `Insurance` | 64 | natural weight column (`Holders` exposure) | **v3.7** WLS |
| R `sleep` | 20 | paired two-sample data | **v3.10** two-sample |
| R `ToothGrowth` | 60 | small two-group + dose (also a tidy 2×3 factorial) | **v3.10** two-sample |
| R `mtcars` | 32 | optional: a pocket-sized log-log / quadratic sandbox for fast manual iteration — nice-to-have, not required for coverage | convenience only |

### Already in `sample_data/`, currently unwired

- `cars2010.csv` (1107 rows, rich categoricals — `Transmission`, `DriveDesc`, `CarlineClassDesc`):
  a good second categorical-heavy source if wired, but **subsample to ≤ ~250 rows** if adopted;
  nothing in Section 1 requires it.
- `winequality-white.csv` (4898 rows): **recommend skipping** — large, and its only distinctive
  feature (ordinal response) isn't a modeled concept in the tool.

### Timing

Only the calendar-dated series affects Section-1 coverage; every other addition can wait until the
milestone in its `Needed by` column starts. Nothing needs to be imported in this pass.

The two two-sample datasets (`sleep`, `ToothGrowth`) are now the **last** to be needed, since
v3.10 sits at the end of the ladder. That is a scheduling fact, not a demotion — they are still
the right datasets when the milestone comes up.

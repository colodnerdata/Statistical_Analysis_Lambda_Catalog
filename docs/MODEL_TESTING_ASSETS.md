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

### 1.1 Mileage / Auto MPG (406 rows, 392 complete) — baseline, categoricals, interactions

| ID | Model | Configuration | Covers | Status |
|---|---|---|---|---|
| M1 | `MPG ~ Horsepower + Weight + C(Model Year) + C(Origin)` | Car Name = Identifier; **no Sequence axis** (see the note below); k = 16 | shipped T0 baseline; two categoricals + two continuous; (None, None) pair; unit-space reduction invariant (fit-space ≡ unit-space when no transform) | existing — `default_t0_intercept` |
| M2 | M1, intercept OFF | carries the deliberate red CF (intercept-off + included categorical) | no-intercept with categoricals | existing — `default_t0_no_intercept` |
| M3 | `MPG ~` all 5 continuous (Cylinders, Displacement, Horsepower, Weight, Acceleration), ± intercept | categoricals excluded | all-continuous fit, intercept on/off pair | existing — `v1_full_continuous_intercept` / `_no_intercept` |
| M4 | `MPG ~` curated continuous subset, ± intercept | | `Include = FALSE` candidate rows | existing — `continuous_subset_intercept` / `_no_intercept` |
| M5 | `Ln(MPG) ~ Ln(Weight) + Ln(Horsepower)` | Duan default | **(Log, Log) on a dataset other than Production Lots**; NA-propagation masking (8 missing MPG, 6 missing Horsepower) | existing — `mileage_log_log_na_masking` |
| M6 | `MPG ~ Horsepower + Horsepower × Horsepower` | self-product | quadratic (x²) declaration | existing — `interaction_quadratic_self_product` |
| M7 | continuous × continuous product | one extra design column | Cont×Cont `Product` | existing — `interaction_continuous_product` |
| M8 | `Weight × C(Origin)` | L−1 extra columns | Cont×Cat broadcast | existing — `interaction_categorical_broadcast` |
| M9 | `MPG ~ C(Model Year) + C(Origin) + C(Model Year) × C(Origin)` | Model Year as Categorical (numeric-valued, 13 levels) | **Cat×Cat full-product width** (12 × 2 = 24 interaction columns); numeric-valued categorical | existing — `interaction_categorical_cross` |
| M10 | `MPG ~ Displacement + Weight + Displacement − Acceleration` | Acceleration is an excluded operand | **first `Difference` case**; antisymmetric op; `" − "` (U+2212) header; also covers G11's amber marginality path | existing — `interaction_difference` |
| M11 | `MPG ~ Weight + Horsepower + Weight ÷ Horsepower + Horsepower ÷ Weight` | both reciprocal Ratio rows declared | **first `Ratio` case**; zero-denominator `#N/A` path; **legal reciprocal pair** (Ratio is asymmetric, so no singular-Gram flag) | existing — `interaction_ratio_reciprocal` |
| M12 | M1 with `Origin` reference = `Europe` | typed reference | explicit reference override | existing — `origin_explicit_reference` |
| M13 | M1 with `Origin` reference blank | first-in-sort-order default | default reference | existing — `origin_default_reference` |
| M14 | `MPG ~ Displacement + Horsepower + Weight + C(Model Year) + C(Origin)` | two categoricals **plus** three continuous | multi-level categorical alongside continuous predictors | existing — `model_year_origin_categorical` |
| M14b | `MPG ~ C(Model Year) + C(Origin)` | no continuous predictors at all | **categorical-only design**; the mask reduces to "response is numeric", so n grows 392 → 398; M9's interaction-free base | existing — `categorical_only_design` |
| M15 | M1 + `Is_USA` Filter | `ExtraSpecColumn` fixture | filter-induced **degenerate categorical** (Origin collapses to one level → 0 columns, red K cell) | existing — `usa_filter_degenerate_origin` |
| M16 | M1 + Model Year `Sequence = TRUE` and a typed `Sequence Period` = 2 | candidate Δ = 1 | **period override**; the verdict re-evaluates against the typed Δ (escalating to off-grid, since odd gaps are not multiples of 2) | existing — **guard state** `guard_sequence_period_override` |

**Auto MPG carries no Sequence axis, and no case here may add one.** The
dataset is cross-sectional: each row is a distinct car model observed once,
with no unit repeated across periods. `Model Year` used to be flagged
`Sequence = TRUE` in the shipped T0 spec, and every case built on that spec
inherited it — which asserted panel structure the data does not have, and
bought a Base Period Δ candidate nobody can interpret plus a Durbin-Watson
computed over an arbitrary row order. (The shipped Identifier, `Car Name`, is
very nearly unique, so `Sequence_Deltas` finds no within-group consecutive
pairs and the spacing verdict is blank regardless.)
`_DEFAULT_SEQUENCE_VARIABLES` is now empty and the flag is gone from every
Auto MPG case except **G3** and **M16**, which test the flag's *mechanics* —
the H2 cardinality rule counts flags, and the typed-override path reads the
flagged row positionally; both are dataset-independent and need a flag
present to be reachable at all. The Sequence layer's substantive coverage
lives on the two datasets that are real panels: Production Lots
(`Fiscal_Year`, §1.3) and Life Expectancy (`Year`, §1.2).
`test_sequence_is_flagged_only_on_datasets_that_have_an_ordering_axis` and
`test_only_the_two_mechanics_cases_flag_sequence_on_auto_mpg` pin both halves.

### 1.2 Life Expectancy (2938 rows) — transform dispatch, scale, missingness

Every row below was new before this pass; all nine are now implemented.

| ID | Model | Configuration | Covers | Status |
|---|---|---|---|---|
| L1 | `Life expectancy ~ Ln(Population) + Ln(GDP) + Alcohol + C(Status)` | **partial linear-log** (`y ~ ln(x)`, logs on the predictors — see the note below); Country = Identifier | **(None, Mixed)** pair; binary categorical; heavy missingness masking (Population 652, GDP 448, Alcohol 194 blanks → n = 2117) | existing — `life_partial_linear_log` |
| L2 | `Ln(Life expectancy) ~ Adult Mortality + Schooling + C(Status)` | **exponential model** (log-level, `ln(y) ~ x`); Back-Transform = Duan (default) | **(Log, None)** pair; smearing factor; unit-space R²/RMSE; Original-Units prediction + AZ/BA residual columns | existing — `life_log_response_duan` |
| L3 | L2 with Back-Transform = **Naive** | flips `AH4` | naive point estimate `EXP(ŷ)`; confirms CI/PI bounds are EXP-only under both settings. **Required an oracle change** — the Python side computed both branches and discarded the Naive one, so `AH4` had never been verified against anything | existing — `life_log_response_naive` |
| L4 | `Ln(Life expectancy) ~ Ln(GDP) + Ln(Population)` | elasticity form | (Log, Log) with large-sample masking | existing — `life_elasticity_log_log` |
| L5 | `Life expectancy ~` all 18 continuous + `C(Status)` | shipped `life_expectancy` profile; Year = Sequence, Country = Identifier | k-stress kitchen sink (k = 19); the shipped default finally gets an oracle. **Heavy** — at k = 19 with n = 2117 the statsmodels OLS reference and Excel's OLS implementation diverge in the 5th–6th decimal place on most coefficients and residuals (both go through a QR-with-column-pivoting path on an ill-conditioned Gram matrix and produce near-tied numerics), so this case lives behind `--include-heavy` as a deliberate showcase for the floor, not as a defect to paper over. | existing — `life_full_profile` (**heavy**) |
| L6 | `Life expectancy ~ Adult Mortality + Ln(Schooling)` | Schooling contains 28 true zeros | **`Ln_Positive` zero guard** — see the correction below: the rows do NOT drop out of the mask | existing — **guard state** `guard_ln_zero_propagation` |
| L7 | `Life expectancy ~ C(Country) + C(Year) +` 8 continuous | 183 countries → 182 dummies, + 15 Year dummies + 8; k = 205 | **width-guard soft warning** (k = 200 threshold, `M2` status), and the engine degrading visibly rather than returning a plausible wrong number | existing — **guard state** `guard_width_guard_warning` |
| L8 | `Life expectancy ~ Schooling + Adult Mortality \| Country` | Year = Sequence | **high-cardinality Fixed Effects** (173 surviving groups, 172 absorbed df); panel spacing verdicts at scale | existing — `life_country_fixed_effects` (**heavy**) |
| L9 | L1 with `Status` reference = `Developing` | retained dummy = `Developed` | explicit reference on a **binary** categorical | existing — `life_status_explicit_reference` |

**L1 is linear-log, not log-linear.** The two names describe opposite
specifications and L1 is unambiguously the first: the logs sit on the
*predictors* and the response is untransformed, so a coefficient reads as a
semi-elasticity (years of life expectancy per 100% change in GDP).
"Log-linear" (log-lin) is the mirror image, `ln(y) ~ x` — which in this table
is **L2**, already carrying that model's other standard name, the exponential
model. Calling L1 "log-linear" put the same label on both halves of the very
dispatch pair the two cases exist to tell apart, so the case, its spec
builder and its worksheet were renamed to `life_partial_linear_log` /
`L01 Partial Linear-Log`. "Partial" is the `Mixed` half: Alcohol stays raw
while Population and GDP are logged. The other three model names in this
document — L2's exponential model, P1's and P3/P3b's power law — are the standard
terms for what those cases fit and are unchanged.

**Three of these did not survive contact with the data as written, and the
deviations are recorded rather than papered over.**

* **L6 contradicts the shipped mask.** The row above used to read "`NA()` on zero
  → row drops out of the mask, not a silent 0". That is not what happens.
  `Sample_Include` tests `ISNUMBER(col)` on the Response and the included
  Continuous Predictors, and `ISNUMBER(0)` is TRUE — there is no Log-positivity
  term anywhere in it. So the 28 zero-Schooling rows stay in the sample,
  `Ln_Positive` returns `#N/A` for each, and the `#N/A` propagates through
  `Predictor_Columns` into every downstream statistic. L6 is therefore a **guard
  state** asserting that propagation, not a fittable model. Adding a positivity
  term to `Sample_Include` would make the original description true and is
  arguably better behaviour — it is an **open production question**, deliberately
  not decided while writing an oracle. See `_LN_ZERO_GUARD_NOTE` in
  `lambda_catalog/analyze_regression_guard_states.py`.
* **L7 cannot reach k = 200 the way the plan assumed.** The arithmetic (193
  countries → 192 dummies + 8 predictors = 200) ignores missingness: the response
  itself is blank on rows covering 10 countries, so at most 183 countries ever
  survive the mask and the dummy block caps at 182. Adding sparser predictors
  drops further countries roughly as fast as it adds columns. Declaring `Year`
  (16 levels) as a second Categorical Predictor adds 15 columns at no row cost,
  putting k at 205 — over the threshold with margin. `test_width_guard_case_
  actually_crosses_the_two_hundred_column_threshold` pins that, so a future data
  or spec change cannot silently drop the case back under 200 and leave the guard
  untested.
* **L8 has 173 groups, not 193.** Schooling is blank for every row of 20
  countries, so those panels leave the sample entirely.

### 1.3 Production Lots (51 rows) — learning curves, fixed effects, sequence

| ID | Model | Configuration | Covers | Status |
|---|---|---|---|---|
| P1 | `log Unit Cost ~ log Cum Units \| Facility` | **learning-curve power law with FE** (log-log ⟺ `cost = A·units^b`); pre-derived ln columns; `Full_Data` = Filter; Fiscal_Year = Sequence with a **typed `Sequence Period` = 1**; prediction group `Site B` | one-way FE; Filter role; group prediction; **the BFN panel Durbin-Watson** (the only cases that make that cell live) | existing — `production_lots_fixed_effects` |
| P2 | `Ln(Unit_Cost_BY) ~ Ln(Cumulative_Units) \| Facility` | raw columns with `Transform = Log`; typed `Sequence Period` = 1, matching P1 | FE + (Log, Log) via the transform axis (vs. P1's pre-derived columns); BFN, equal to P1's | existing — `production_lots_log_transform` |
| P3 | `log Unit Cost ~ log Cum Units` | **power law without FE**, from the pre-derived ln columns | the pre-derived half of the no-FE pair; P3b's twin | existing — `production_lots_derived_log_no_fe` |
| P3b | `Ln(Unit_Cost_BY) ~ Ln(Cumulative_Units)` | same model, raw columns with `Transform = Log` | (Log, Log), no level shift; **transform axis isolated from FE** (vs. P3's pre-derived columns) | existing — `production_lots_log_no_fe` |
| P4 | mixed logged/unlogged predictors | | **(Log, Mixed)** pair | existing — `production_lots_log_mixed_predictors` |
| P5 | `Unit_Cost_BY ~ Ln(Cumulative_Units)` | | **(None, Log)** pair; must reproduce ordinary fit-space stats exactly | existing — `production_lots_log_predictor_only` |
| P6 | P2 with `Facility` as **Categorical Predictor** instead of Fixed Effects | intercept ON, default reference | **LSDV ↔ within-estimator equivalence** — identical slope and residual vector as P2 (agreement to ~1e-15), by a completely different estimator path; the strongest cheap cross-oracle in the suite | existing — `production_lots_lsdv_equivalence` |
| P7 | P1 with `Facility` as the **Identifier** | per-facility Fiscal_Year gaps (A 1998–2023, B 2001–2020, C 2005–2024) | **irregular-spacing Regularity verdict** (yellow); typed Δ = 1 override on a gapped panel | existing — **guard state** `guard_irregular_panel_spacing` |

**P7's Identifier change is not cosmetic.** `Sequence_Deltas` groups by the
**Identifier** columns — the Identifier is what the spacing layer treats as the
panel unit, not the Fixed Effects column. In P1/P2's shipped spec the Identifier
is `Lot_ID`, which is unique per row, so every group is a singleton, there are no
within-group consecutive pairs, and the verdict cell is unconditionally blank no
matter how gapped the fiscal years are. Declaring `Facility` as the Identifier is
what makes the three sites the groups and the verdict reachable at all.

**Two pre-derived/transform-axis pairs, each pair adjacent.** P1/P2 and P3/P3b
fit the same model twice by different routes: one spec reads the shipped
`log Cum Units` / `log Unit Cost` columns and declares no transform, the other
points at the raw `Cumulative_Units` / `Unit_Cost_BY` columns and declares
`Transform = Log`. The shipped log columns are exact logs of the raw ones, so
each pair must agree bit-for-bit on the design matrix and response vector and
to floating point on every downstream statistic — the cheapest strong oracle
available, with neither side reading the workbook.

The pairing exists **twice on purpose**. The two mechanisms reach the design
matrix by different code paths (one reads a column, the other computes one),
and composing either with Fixed Effects demeaning is a third path again. With
only the FE pair, a transform-axis regression could hide behind the demeaning
or vice versa; P3/P3b has no FE, so the transform axis is the only thing
between the CSV and the design matrix and a disagreement can only be the
transform wiring. `tests/test_transform_threading.py` asserts both pairs.
Each pair is registered adjacently so the two land on adjacent worksheets, and
the sheet names state the route — `P03 Power Law Derived Cols` against
`P03b Power Law Transform Axis` — because the route is the only thing that
differs between the tabs and the whole reason both exist.

**The BFN cell needs a typed Sequence Period, and only P1/P2 give it one.**
`Base_Period_Delta()` is the **override** accessor: it reads the typed value in
spec column I and returns `#N/A` when the cell is blank — never a silent 1, by
design (see DECISIONS § *Sequence Period / Period In Use split*). The BFN panel
Durbin-Watson cell passes that as its Δ, so a Fixed Effects sheet with no typed
period leaves `AE12` at `#N/A` and its panel diagnostic is unverifiable. P1 and
P2 declare `Sequence Period = 1`, which is a true statement about the data —
Production Lots is an annual panel — and makes them the only cases where the
statistic itself is compared. **L8 deliberately does not**: it is the case for
high-cardinality FE degrees of freedom, and leaving its period untyped keeps one
registered case covering the honest `#N/A` state.

### 1.4 Guard-rail / error-state configurations

Not fittable models — these verify status lines, conditional formatting, and graceful degradation.
Any dataset works unless noted; Auto MPG is the default fixture.

**They have their own oracle shape.** A guard case is a `GuardStateCase` in
`lambda_catalog/analyze_regression_guard_states.py`, not a `RegressionSpecCase`:
`calculate_regression_spec_case` raises on most of these by design, and what they
assert is status text, the per-row Design Columns audit, the Model Formula, and
which CF rules fire — not fit statistics. Flags are recorded as **predicates**
recomputed from the same condition the CF expression encodes, never read back as
`DisplayFormat.Interior.Color`: reading the colour would only re-report what Excel
already decided from the rule, whereas recomputing the predicate is what makes a
silent change to a CF expression fail. Names are pinned in
`_EXPECTED_GUARD_NAMES` (`tests/test_regression_guard_states.py`), the same
regime as `_EXPECTED_CASE_NAMES`.

| ID | Configuration | Expected behavior | Status |
|---|---|---|---|
| G1 | zero `Response (y)` rows | response readout degrades to `"(none)"`; the mask loses its response term so the sample GROWS (392 → 400) | existing — `guard_no_response` |
| G1b | no included `Predictor (x)` rows | `Predictor_Columns` errors and every consumer's `IFERROR` degrades to `"(empty model)"`, never `#CALC!` | existing — `guard_empty_model` |
| G2 | two `Response (y)` rows | audit-strip `responses` count red; `XMATCH` still resolves the first | existing — `guard_two_responses` |
| G3 | two `Sequence = TRUE` flags | `E1`/`H2` "multiple Sequence flags" error + red H cells. The spacing layer keeps computing from the FIRST flagged row (that is what `XMATCH` resolves), so the oracle must not go quiet here | existing — `guard_two_sequence_flags` |
| G4 | two `Fixed Effects` rows | `B1` cardinality error + red CF | existing — `guard_two_fixed_effects` |
| G5 | Fixed Effects row + Intercept `TRUE` | red CF on the intercept toggle (double-counted demeaned design) | existing — `guard_fixed_effects_with_intercept` |
| G6 | Intercept `FALSE` + included Categorical | red CF on the toggle; the formula renders `~ 0 + …` (M2 fits anyway — flag is advisory) | existing — `guard_intercept_off_with_categorical` |
| G7 | `Transform = Log` on a Categorical predictor | red G cell; transform **never silently applied** (no `Ln(` in the model formula) | existing — `guard_log_on_categorical` |
| G8 | typed reference absent from sample (`Origin` = `99`) | red E cell via `ISNA(Dummy_Levels(...))`; row contributes 0 columns | existing — *fittable* `origin_invalid_reference` |
| G9 | Interaction Term naming a non-Predictor row (or no column) | red M cell; the row contributes its main effect only | existing — `guard_interaction_bad_operand` |
| G10 | reciprocal `Product` declaration (A×B and B×A) | red N on **both** rows — symmetric op ⇒ singular Gram. M11 is the legal counterpart under `Ratio` | existing — `guard_reciprocal_product` |
| G11 | interaction operand with `Include = FALSE` | **amber** — marginality violation, allowed. Also exercised by M10 | existing — `guard_excluded_operand` |
| G12 | unrecognized Interaction Operation pasted past the dropdown | `" ? "` header + `NA()` design column. No CF rule fires: the refusal is visible in the header itself | existing — `guard_unknown_interaction_operation` |
| G13 | width **hard** error (k > 16384 − design-matrix origin) | documented as conceptual only — not buildable at reasonable size; the soft warning is L7's job | conceptual |
| G14 | spec block built for a **narrower** dataset than `Source_Table` points at | every part of the block resizes to `COLUMNS(Source_Data)` — bands, the four computed columns, the input fill — and the model fits correctly on the wider table | existing — `guard_spec_block_retarget_widens` (L10) |

**G14 is the one case that exercises the retarget itself.** Every other case
builds its sheet for exactly the dataset it reads, so the two always agree and
the retarget path is never taken. `GuardStateCase.shell_profile_key` is what
creates the disagreement deliberately: L10 builds its block with the Auto MPG
profile (12 columns) and then points `Source_Table` at `LifeExpectancyData`
(23), which is what a user does by hand from the Name Manager — the one-name
edit the Instructions sheet promises.

Before the spec block was made table-free this state produced `#REF!` through
the entire engine. The `Spec_*` bands were structured references into a
`SpecTable` ListObject sized at build time; a 12-row band under a 23-column
table meant `TAKE` returned 12 rows (it does not pad) and `INDEX(rl, 23)` ran
off the end. Excel cannot resize a ListObject from a formula and the workbook is
macro-free, so the table was removed and the block now sizes itself.

The case earns its sheet by where its evidence sits, not by the model it fits:
`Schooling` contributes design columns from spec index 21 — sheet row 25, ten
rows past the old table's bottom edge at row 15.
`test_retarget_case_puts_its_evidence_past_the_narrow_shells_last_row` pins
that, so a future edit that moves the predictors up into the first 12 rows
fails rather than silently testing nothing.

Four rows from other sections live here, because everything they test is
spec-block state rather than a fit — and in one case because the fit does not
exist:

* **L6** (`guard_ln_zero_propagation`) — the mask has no Log-positivity term, so
  the zero rows stay in and `#N/A` propagates.
* **M16** (`guard_sequence_period_override`) and **P7**
  (`guard_irregular_panel_spacing`) fit exactly the models M1 and P2 already fit,
  so registering them as fittable cases would have added two duplicate fits and
  covered nothing new.
* **L7** (`guard_width_guard_warning`) — the workbook cannot invert a 205-column
  Gram matrix, so there are no numbers to compare. What it asserts instead is the
  guard doing its job.

The last one is worth stating as a rule, since it will come up again: **a numeric
oracle for a model the sheet cannot compute is comparing against nothing.** When
a case's whole point is a limit, the limit is the assertion.

**One reachability note worth recording.** The Sequence verdict's four branches
are priority-ordered (off-grid → regularity → no-natural-period → calendar), and
the ordering makes the last two very narrow. `no natural base period` requires Δ
to divide and equal every spacing while no spacing repeats — reachable only with
a single spacing. The `calendar` signature requires spacings to be *perfectly*
uniform at a calendar-like value, so a realistic mixed 30/31-day monthly series
reports regularity instead. Neither is a bug, but neither is as reachable as the
message text implies.

### 1.5 Coverage matrix

| Feature axis / corner | Covered by |
|---|---|
| Role = Response (exactly one) | every model; G1/G2 for the violations |
| Role = Predictor, Continuous | every model |
| Role = Identifier | M1 (Car Name), L5 (Country), P1 (Lot_ID), P7 (Facility — the panel unit) |
| Role = Filter | P1 (`Full_Data`), M15 (`Is_USA`) |
| Role = Fixed Effects | P1/P2, L8 (173 groups); G4/G5 violations |
| Role = Omit / blank ≡ Omit | M1 (Make, Model?) |
| Include = FALSE candidates | M4; M10/G11 (as interaction operand); G1b (all of them) |
| Type = Categorical, multi-level | M1, M9, M14, M14b, L7 |
| Numeric-valued categorical | M1/M14/M14b/M9 (Model Year), L7 (Year) |
| Binary categorical | L1, L9 |
| Reference: default / explicit / invalid | M13 / M12, L9 / G8 |
| Degenerate categorical (post-filter) | M15 |
| Dispatch (None, None) | M1 |
| Dispatch (None, Log) | P5 |
| Dispatch (None, Mixed) | L1 |
| Dispatch (Log, None) | L2 |
| Dispatch (Log, Log) | M5, L4, P3b, P6 (and P2 under FE) |
| Dispatch (Log, Mixed) | P4 |
| Back-Transform = Duan / Naive | L2 / L3 (the toggle's first oracle) |
| Unit-space reduction invariant (no transforms) | M1 |
| `Ln_Positive` zero/negative guard | L6 (as #N/A propagation — see § 1.2) |
| Missing-data NA propagation | M5, L1, L4, L7, L8 |
| Intercept OFF | M2, M3/M4 variants |
| FE + Log | P2 |
| FE + intercept flag | G5 |
| Serial correlation: plain DW / BFN panel form | every no-FE case / P1, P2 (the two mutually-gated cells; the oracle NaNs whichever one the sheet shows as text) |
| Sequence: candidate Δ / typed override / irregular / calendar | P1 / M16 / P7 / **uncovered — needs a dated dataset (§2 `Time` role, §3)** |
| Interaction: Product / self-product / Cont×Cat / Cat×Cat / Difference / Ratio | M7 / M6 / M8 / M9 / M10 / M11 |
| Reciprocal declaration: legal (Ratio) / illegal (Product) | M11 / G10 |
| Width guard: soft / hard | L7 (k = 205) / G13 (conceptual) |
| `Source_Table` retarget onto a wider dataset | G14 / L10 (12-column shell → 23-column table) |
| Group prediction under FE | P1 |
| LSDV ↔ FE equivalence | P6 vs P2 |
| Categorical-only design (mask without continuous predictors) | M14b |
| Empty model / no response degradation | G1b / G1 |
| Model Formula cell (AB2) text | every guard case; corrected to mirror the cell exactly |

The one axis Section 1 cannot cover with the wired data is the **calendar-signature Sequence
verdict** (~28–31 / ~90–92 / ~365–366-day spacing clusters): no wired dataset carries real dates.
See the `Time` role entry in §2 and §3.

---

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

**Sheet names state the concept, not the variables.** `M05 Log-Log NA Masking`,
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

**Two cases are opt-in.** L05 (Kitchen Sink Profile, k = 19, n = 2117) and L08
(173 Fixed Effects groups) carry `heavy=True` and are skipped unless
`--include-heavy` or an explicit `--cases L05` / `--cases L08` is given. L08
is gated on sheet-build cost; L05 is gated on the statsmodels-vs-Excel
floating-point floor at fdd = 5/6 that both implementations agree on (it
lives here as a deliberate showcase for the floor, not as a defect). Their
Python oracles always run in the unit suite — only the sheet build is gated.
L07 was the third candidate until the live run showed the workbook cannot
fit a 205-column design at all — it is a guard state now, and guard sheets
are cheap.

```
python scripts/build_test_models.py                        # 47 sheets (31 models + 16 guards)
python scripts/build_test_models.py --include-heavy        # 49, adding L05 and L08
python scripts/build_test_models.py --cases M09,G10        # just those two
python scripts/build_test_models.py --verify --no-launch   # build, check, exit 1 on drift
make verify-test-models                            # the same, verbose
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
exists for M15, which declares it as a Filter. The legacy verifier adds it and
deletes it again around that one case; this workbook keeps it, which makes it
part of *every* Auto MPG case's source table. Every constructor opens
`n_c = COLUMNS(Source_Data)` and then indexes the `Spec_*` bands at `1..n_c`, so
a 12-row spec block against a 13-column table makes `INDEX(rl, n_c)` run off the
end — the row mask errors and every engine cell downstream reads as an error.
The build succeeds, the sheet looks right, and it surfaces only as a verify run
where an entire case reads `None`. M15 was the one Auto MPG sheet that worked,
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

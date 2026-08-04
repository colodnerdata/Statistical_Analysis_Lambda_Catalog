# Model Testing Assets

A planning document for the regression test-model suite. It lists the **model configurations** the
QC harness should cover — nothing here adds code, sheets, or data. Future work turns each listed
model into a `RegressionSpecCase` (see `lambda_catalog/analyze_regression_spec.py` —
`SpecVariable` rows built with `_spec_var(...)`, per-case `source_table_ref` retargeting of
`Source_Table`, statsmodels/NumPy oracles via `calculate_regression_spec_case`) and, separately,
into an inserted sheet per model type.

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
| M1 | `MPG ~ Horsepower + Weight + C(Model Year) + C(Origin)` | Model Year = Sequence; Car Name = Identifier; k = 16 | shipped T0 baseline; two categoricals + two continuous; (None, None) pair; unit-space reduction invariant (fit-space ≡ unit-space when no transform) | existing — `default_t0_intercept` |
| M2 | M1, intercept OFF | carries the deliberate red CF (intercept-off + included categorical) | no-intercept with categoricals | existing — `default_t0_no_intercept` |
| M3 | `MPG ~` all 5 continuous (Cylinders, Displacement, Horsepower, Weight, Acceleration), ± intercept | categoricals excluded | all-continuous fit, intercept on/off pair | existing — `v1_full_continuous_intercept` / `_no_intercept` |
| M4 | `MPG ~` curated continuous subset, ± intercept | | `Include = FALSE` candidate rows | existing — `continuous_subset_intercept` / `_no_intercept` |
| M5 | `Ln(MPG) ~ Ln(Weight) + Ln(Horsepower)` | Duan default | **(Log, Log) on a dataset other than Production Lots**; NA-propagation masking (8 missing MPG, 6 missing Horsepower) | new |
| M6 | `MPG ~ Horsepower + Horsepower × Horsepower` | self-product | quadratic (x²) declaration | existing — `interaction_quadratic_self_product` |
| M7 | continuous × continuous product | one extra design column | Cont×Cont `Product` | existing — `interaction_continuous_product` |
| M8 | `Weight × C(Origin)` | L−1 extra columns | Cont×Cat broadcast | existing — `interaction_categorical_broadcast` |
| M9 | `MPG ~ C(Cylinders) + C(Origin) + C(Cylinders) × C(Origin)` | Cylinders as Categorical (numeric-valued, 5 levels) | **Cat×Cat full-product width** (4 × 2 retained levels); numeric-valued categorical | new |
| M10 | `MPG ~ Displacement + Horsepower + Displacement − Horsepower` | | **first `Difference` case**; antisymmetric op; `" − "` (U+2212) header | new |
| M11 | `MPG ~ Weight + Horsepower + Weight ÷ Horsepower + Horsepower ÷ Weight` | both reciprocal Ratio rows declared | **first `Ratio` case**; zero-denominator `#N/A` path; **legal reciprocal pair** (Ratio is asymmetric, so no singular-Gram flag) | new |
| M12 | M1 with `Origin` reference = `Europe` | typed reference | explicit reference override | existing — `origin_explicit_reference` |
| M13 | M1 with `Origin` reference blank | first-in-sort-order default | default reference | existing — `origin_default_reference` |
| M14 | `MPG ~ C(Model Year) + C(Origin)` | two categoricals, no continuous | categorical-only design | existing — `model_year_origin_categorical` |
| M15 | M1 + `Is_USA` Filter | `ExtraSpecColumn` fixture | filter-induced **degenerate categorical** (Origin collapses to one level → 0 columns, red K cell) | existing — `usa_filter_degenerate_origin` |
| M16 | M1 with typed `Sequence Period` = 2 on Model Year | candidate Δ = 1 | **period override** (yellow J cell), Regularity verdict against the overridden Δ | new |

### 1.2 Life Expectancy (2938 rows) — transform dispatch, scale, missingness

Currently **zero** regression QC cases use this dataset; every row below is new.

| ID | Model | Configuration | Covers |
|---|---|---|---|
| L1 | `Life expectancy ~ Ln(Population) + Ln(GDP) + Alcohol + C(Status)` | user-named **partial log-linear**; Country = Identifier | **(None, Mixed)** pair; binary categorical; heavy missingness masking (Population 652, GDP 448, Alcohol 194 blanks) |
| L2 | `Ln(Life expectancy) ~ Adult Mortality + Schooling + C(Status)` | user-named **exponential model**; Back-Transform = Duan (default) | **(Log, None)** pair; smearing factor; unit-space R²/RMSE; Original-Units prediction + AZ/BA residual columns |
| L3 | L2 with Back-Transform = **Naive** | flips `AH4` | naive point estimate `EXP(ŷ)`; confirms CI/PI bounds are Naive under both settings |
| L4 | `Ln(Life expectancy) ~ Ln(GDP) + Ln(Population)` | elasticity form | (Log, Log) with large-sample masking |
| L5 | `Life expectancy ~` all 18 continuous + `C(Status)` | shipped `life_expectancy` profile; Year = Sequence, Country = Identifier | k-stress kitchen sink; the shipped default finally gets an oracle |
| L6 | L2 + `Ln(Schooling)` | Schooling contains true zeros | **`Ln_Positive` zero guard**: `NA()` on zero → row drops out of the mask, not a silent 0 |
| L7 | `Life expectancy ~ C(Country) +` ~8 continuous | 193 levels → 192 dummies; k ≈ 201 | **width-guard soft warning** (k = 200 threshold, `M2` status) |
| L8 | `Life expectancy ~ Schooling + Adult Mortality \| Country` | Year = Sequence | **high-cardinality Fixed Effects** (193 groups); panel spacing verdicts at scale |
| L9 | L1 with `Status` reference = `Developing` | retained dummy = `Developed` | explicit reference on a **binary** categorical |

### 1.3 Production Lots (51 rows) — learning curves, fixed effects, sequence

| ID | Model | Configuration | Covers | Status |
|---|---|---|---|---|
| P1 | `log Unit Cost ~ log Cum Units \| Facility` | user-named **learning-curve power law with FE**; pre-derived ln columns; `Full_Data` = Filter; Fiscal_Year = Sequence; prediction group `Site B` | one-way FE; Filter role; group prediction | existing — `production_lots_fixed_effects` |
| P2 | `Ln(Unit_Cost_BY) ~ Ln(Cumulative_Units) \| Facility` | raw columns with `Transform = Log` | FE + (Log, Log) via the transform axis (vs. P1's pre-derived columns) | existing — `production_lots_log_transform` |
| P3 | `Ln(Unit_Cost_BY) ~ Ln(Cumulative_Units)` | user-named **power law without FE** | (Log, Log), no level shift | existing — `production_lots_log_no_fe` |
| P4 | mixed logged/unlogged predictors | | **(Log, Mixed)** pair | existing — `production_lots_log_mixed_predictors` |
| P5 | `Unit_Cost_BY ~ Ln(Cumulative_Units)` | | **(None, Log)** pair; must reproduce ordinary fit-space stats exactly | existing — `production_lots_log_predictor_only` |
| P6 | P2 with `Facility` as **Categorical Predictor** instead of Fixed Effects | intercept ON, default reference | **LSDV ↔ within-estimator equivalence** — identical slopes/fit as P2; the strongest cheap cross-oracle in the suite | new |
| P7 | P2 with attention on the Sequence block | per-facility Fiscal_Year gaps (A 1998–2023, B 2001–2020, C 2005–2024) | **irregular-spacing Regularity verdict** (yellow); typed Δ = 1 override on a gapped panel | new |

### 1.4 Guard-rail / error-state configurations

Not fittable models — these verify status lines, conditional formatting, and graceful degradation.
Any dataset works unless noted; Auto MPG is the default fixture.

| ID | Configuration | Expected behavior |
|---|---|---|
| G1 | zero `Response (y)` rows | model formula degrades to `"(empty model)"`, never `#CALC!` |
| G2 | two `Response (y)` rows | audit-strip `responses` count red |
| G3 | two `Sequence = TRUE` flags | `E1` "multiple Sequence flags" error + red H cells |
| G4 | two `Fixed Effects` rows | `B1` cardinality error + red CF |
| G5 | Fixed Effects row + Intercept `TRUE` | red CF on the intercept toggle (double-counted demeaned design) |
| G6 | Intercept `FALSE` + included Categorical | red CF on the toggle (M2 fits anyway — flag is advisory) |
| G7 | `Transform = Log` on a Categorical predictor | red G cell; transform **never silently applied** |
| G8 | typed reference absent from sample (`Origin` = `99`) | red E cell via `ISNA(Dummy_Levels(...))`; row contributes 0 columns | *(existing fittable case — `origin_invalid_reference`)* |
| G9 | Interaction Term naming a non-Predictor row (or no column) | red M cell |
| G10 | reciprocal `Product` declaration (A×B and B×A) | red — symmetric op ⇒ singular Gram |
| G11 | interaction operand with `Include = FALSE` | amber — marginality violation, allowed |
| G12 | unrecognized Interaction Operation pasted past the dropdown | `" ? "` header + `NA()` design column |
| G13 | width **hard** error (k > 16384 − design-matrix origin) | documented as conceptual only — not buildable at reasonable size; the soft warning is L7's job |

### 1.5 Coverage matrix

| Feature axis / corner | Covered by |
|---|---|
| Role = Response (exactly one) | every model; G1/G2 for the violations |
| Role = Predictor, Continuous | every model |
| Role = Identifier | M1 (Car Name), L5 (Country), P1 (Lot_ID) |
| Role = Filter | P1 (`Full_Data`), M15 (`Is_USA`) |
| Role = Fixed Effects | P1/P2, L8; G4/G5 violations |
| Role = Omit / blank ≡ Omit | M1 (Make, Model?) |
| Include = FALSE candidates | M4; G11 (as interaction operand) |
| Type = Categorical, multi-level | M1, M9, M14 |
| Numeric-valued categorical | M1/M14 (Model Year), M9 (Cylinders) |
| Binary categorical | L1, L9 |
| Reference: default / explicit / invalid | M13 / M12, L9 / G8 |
| Degenerate categorical (post-filter) | M15 |
| Dispatch (None, None) | M1 |
| Dispatch (None, Log) | P5 |
| Dispatch (None, Mixed) | L1 |
| Dispatch (Log, None) | L2 |
| Dispatch (Log, Log) | M5, L4, P3 (and P2 under FE) |
| Dispatch (Log, Mixed) | P4 |
| Back-Transform = Duan / Naive | L2 / L3 |
| Unit-space reduction invariant (no transforms) | M1 |
| `Ln_Positive` zero/negative guard | L6 |
| Missing-data NA propagation | M5, L1, L4 |
| Intercept OFF | M2, M3/M4 variants |
| FE + Log | P2 |
| FE + intercept flag | G5 |
| Sequence: candidate Δ / typed override / irregular / calendar | M1 / M16, P7 / P7 / **uncovered — needs a dated dataset (§2.5, §3)** |
| Interaction: Product / self-product / Cont×Cat / Cat×Cat / Difference / Ratio | M7 / M6 / M8 / M9 / M10 / M11 |
| Reciprocal declaration: legal (Ratio) / illegal (Product) | M11 / G10 |
| Width guard: soft / hard | L7 / G13 (conceptual) |
| Group prediction under FE | P1 |
| LSDV ↔ FE equivalence | P6 vs P2 |

The one axis Section 1 cannot cover with the wired data is the **calendar-signature Sequence
verdict** (~28–31 / ~90–92 / ~365–366-day spacing clusters): no wired dataset carries real dates.
See §2.5 and §3.

---

## Section 2 — Assets for roadmap features, ordered by test-scale multiplier

**Ordering principle.** Additive features first (each adds a fixed number of cases), per-model
multipliers later, axis-widening multipliers last — so the suite grows linearly for as long as
possible and the biggest lifts land when the harness is most mature.

Two framing notes:

- One multiplicative feature — the Log transform and its response × predictor dispatch pairs — is
  *already integrated*; its cost is absorbed into Section 1's covering array. That is a fact about
  the current suite, **not a constraint on the ordering logic**: the ranking below scores each
  future feature purely by the marginal growth it forces, independent of what happens to be in
  already.
- The tool is single-user and pre-release; nothing here is frozen. Existing QC cases, dataset
  profiles, and this ordering itself can be backtracked and reshaped as iteration reveals better
  structure. The order is a rework-minimizing default, not a commitment.

| # | Roadmap item | Scale effect | Test assets needed |
|---|---|---|---|
| 1 | **v3.4 Model Comparison sheet** | additive (~1×) — reads existing models | ≥3 registered models with shared prediction inputs — Section 1 already supplies them (e.g. M1, L2, P2). Add **one mismatched-predictor-set pair** (e.g. M1 vs M14) to exercise the `XLOOKUP [if_not_found]` open question. |
| 2 | **v3.6 Two-sample / bivariate** | additive — fixed set of test cases | A small two-group dataset (R `ToothGrowth`, 60 rows — or the in-repo `Status` split of Life Expectancy) and a **paired** dataset (R `sleep`, 20 rows). Cases: equal-variance t, Welch t, paired t, F-test of variances feeding the selector cell. |
| 3 | **v3.5 Resampling & simulation** | additive | No new data. The **seeded pre-drawn `Bootstrap_Random_Draws` table** is itself the asset; Production Lots (n = 51) is the natural small-n bootstrap target (slope CI on P3). PERT/MC cases need only parameter cells. |
| 4 | **Cluster role** (v3.8+ candidate) | near-additive — a variance-estimator variant on a few models | Within-group correlated data: Production Lots facilities suffice initially (3 clusters — deliberately few, to test the small-cluster warning path); `Grunfeld` (item 7) later provides 10–11 proper clusters. |
| 5 | **Time role / time series** (v3.8+ candidate) | near-additive — **and unlocks a today-gap** | A real **calendar-dated monthly series** (~144 rows, AirPassengers-shaped, with an actual date column). No wired dataset has dates; this asset also enables the Sequence **calendar-signature verdict** test in Section 1 immediately, before the Time role ships. Also serves `Moving_Average` / `Exponential_Smoothing` cases. |
| 6 | **v3.7 WLS Weight role** | ~2× over a representative subset | Grouped/heteroskedastic data with a natural weight column: R/MASS `Insurance` (64 rows, claims with exposure `Holders`) or a grouped-mean aggregation of an existing dataset. Plan **weighted variants of ~6 representative Section-1 models** (one per dispatch-pair family), not the whole suite. Include the recorded trap as an oracle assertion: `DEVSQ(√w ⊙ y)` ≠ weighted SST. |
| 7 | **Two-way Fixed Effects** (v3.8+ candidate) | ~2× over the FE family | A balanced two-factor panel: R `Grunfeld` (200 rows, 10 firms × 20 years) plus an **unbalanced variant** (rows deleted) to exercise `Is_Balanced_Panel` and the convergence check. Re-run the FE family (P1/P2/L8 analogues) two-way. |
| 8 | **v3.3 standalone transform library** | the **~10× axis-widener — deliberately last** | Each new Transform value (`Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Zscore_By`, `Decompose_By`) widens the predictor-transform axis that currently holds {None, Log}, and every widening multiplies the response × predictor dispatch table (six recognized pairs today). No new data needed — existing datasets cover all of them. Sequencing *within* the item: (a) the additive helpers first (`Numeric_Complete_Cases`, `Dummy_Column`, `Interact`, `Model_Matrix` — standalone LAMBDAs, fixed test count); (b) predictor-side location/scale transforms next (each adds pairs but not back-transform semantics); (c) **any response-side extension last** — a response transform also multiplies the back-transformation / unit-space semantics (what is the smearing analogue for Zscore⁻¹?), which is the single most expensive kind of growth this project has. |

Unscheduled long-tail items (multi-group means/ANOVA, Fourier, decision analysis) need no assets
beyond the above: ANOVA-as-regression is `warpbreaks` + the existing categorical machinery.

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
| Calendar-dated monthly series (AirPassengers-shaped, real date column) | ~144 | **the only uncovered Section-1 axis** (calendar-signature Sequence verdict); later the Time role and time-series functions | §1.5 gap; roadmap item 5 |
| R `Grunfeld` | 200 | balanced 10×20 two-factor panel; proper cluster count | roadmap items 4, 7 |
| R/MASS `Insurance` | 64 | natural weight column (`Holders` exposure) | roadmap item 6 (WLS) |
| R `sleep` | 20 | paired two-sample data | roadmap item 2 |
| R `ToothGrowth` | 60 | small two-group + dose (also a tidy 2×3 factorial) | roadmap item 2 |
| R `mtcars` | 32 | optional: a pocket-sized log-log / quadratic sandbox for fast manual iteration — nice-to-have, not required for coverage | convenience only |

### Already in `sample_data/`, currently unwired

- `cars2010.csv` (1107 rows, rich categoricals — `Transmission`, `DriveDesc`, `CarlineClassDesc`):
  a good second categorical-heavy source if wired, but **subsample to ≤ ~250 rows** if adopted;
  nothing in Section 1 requires it.
- `winequality-white.csv` (4898 rows): **recommend skipping** — large, and its only distinctive
  feature (ordinal response) isn't a modeled concept in the tool.

### Timing

Only the calendar-dated series affects Section-1 coverage; every other addition can wait until its
roadmap item starts. Nothing needs to be imported in this pass.

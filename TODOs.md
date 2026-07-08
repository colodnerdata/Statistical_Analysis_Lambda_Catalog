# TODOs



## Alias layer

Design/planning item — see ROADMAP.md for the architectural rationale. Aliases are thin ALL-CAPS wrappers whose entire body is a single call to the canonical function; the canonical function remains the single source of truth. Implement only after the canonical library is stable.

Suggested alias names:

**Regression — scalar outputs**

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

**Regression — coefficient vectors**

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

**Regression — observation vectors**

| Alias | Canonical |
|---|---|
| `PRED` | `Predictions` |
| `RESID` | `Residuals` |
| `STDR` | `Studentized_Residuals` |
| `LEV` | `Hat_Diagonal` |
| `COOK_D` | `Cooks_Distance` |
| `LOOCV` | `LOOCV_Prediction` |
| `PI` | `Prediction_Interval` |

**Regression — utilities**

| Alias | Canonical |
|---|---|
| `COMPLETE` | `Complete_Cases_Filter` |
| `CORMAT` | `Correlation_Matrix` |
| `DESIGN` | `Design_Matrix` |

**Univariate — descriptive**

| Alias | Canonical |
|---|---|
| `DSTAT` | `Descriptive_Statistics` |
| `NMISS` | `Missing_Count` |

**Univariate — histogram binning**

| Alias | Canonical |
|---|---|
| `NBINS` | `Number_Of_Histogram_Bins` |
| `EDGES` | `Bin_Edges` |
| `UEDGES` | `Upper_Bin_Edges` |
| `LEDGES` | `Bin_Lower_Edges` |
| `BIN_MIDS` | `Bin_Midpoints` |
| `BIN_FREQS` | `Bin_Counts` |

**Univariate — goodness-of-fit**

| Alias | Canonical |
|---|---|
| `GOF_AD` | `GoF_Anderson_Darling` |
| `GOF_KS` | `GoF_Kolmogorov_Smirnov` |

**Grid-search helpers**

| Alias | Canonical |
|---|---|
| `GS_MIN` | `Grid_Argument_Minimum` |
| `GS_OPT` | `Grid_Search_Optimum` |



## v1.x — Regression sheet

- TODO: Add reference lines to the Cook's Distance, PRESS Residuals, and Leverage vs. Studentized charts. Format thematically similar to the conditional formatting in the table (yellow = mild, red = strong). Use minimalist helper columns (2 anchor points using max/min of the relevant threshold; place them beneath the chart area).

---

## v1.1 — Univariate (shipped; leftovers)

### LAMBDA functions

. See ROADMAP.md § v1.1 Distribution
  fitting for the full rationale.

### Sheet writer (`write_sheet_univariate.py`)

**Additional distributions (long-term)**
- TODO: Add support for more distribution families: Bernoulli, Binomial, Geometric, Negative Binomial, Hypergeometric, Poisson, Uniform, Chi-Square, Student-t.

---

## v2.0 — Specification-Driven Regression (shipped; leftovers)

Human test plan fully executed and signed off PASS 2026-07-05 (T0–T16). One open
decision remains from it:

- TODO: Resolve the blank-categorical caveat — `Sample_Include()`'s role-aware
  completeness layer requires numeric Response and numeric included Continuous
  Predictors, but Categorical Predictors impose no non-blank condition; a blank
  category value encodes as all-zero dummies (indistinguishable from the reference
  level). Run the caveat verification step in `HUMAN_TEST_PLAN_v3_model_construction.md`
  and record the decision: accept as documented behavior, or extend `Sample_Include()`
  with a non-blank condition for included Categorical Predictors. Interim workaround:
  a completeness column declared as a Filter.

---

## v2.1 — Sequence, gap-aware longitudinal, serial-correlation diagnostics, fixed effects (in progress)

Two-way FE is deliberately deferred until this framework is finished — see the
v2.5+ section. Items below are in the locked ship order: the Sequence fix is
the prerequisite for the 2.1.0 entry; the FE Role dropdown, the CI+PI
prediction layout, and the engine are gated to ship as a single release so
users never see "FE is in the dropdown but the engine is forthcoming."

### Pending (in ship order; #1 prerequisite, #2/#3/#9 gated to #5)

- TODO: **Sequence axis auto-detection and override** (renames column I to **`Sequence Period`**, adds column J **`Period In Use`** following the Reference Level / Reference In Use pattern). The current override mechanic has a spill-collision risk for source tables wider than the shipped WHO sample: the spec block reads its own H/I cells, and a longer table could let the override spill overrun an input band. Fix: relocate the override spill and bound every read of the H/I/J band by `COLUMNS(Source_Data)` (the spill-placement principle from `CLAUDE.md`). Yellow CF on `Period In Use` when overridden; red CF on `Period In Use` when off-grid; per-row CF on the override cell. Update the Sequence Spacing block (rows 28–34), the spec layout constants, and the QC analyzers. **Significant testing. Resolve before writing the 2.1.0 Version History entry.**

- TODO: **FE Role dropdown + status-block validation** (gated to ship with the engine). `Fixed Effects` in the Role axis; status-block cells for "active FE variable," "group count" (`n/a — engine forthcoming` until the engine lands), "absorbed df" (`n/a — engine forthcoming` until the engine lands); visible error at 2+ FE variables; intercept × FE red flag; CF bands update. Sheet-only; no engine change.

- TODO: **Surface BOTH intervals in adjacent cells of the prediction outputs section** (gated to ship with the engine). Three lines: point · CI low/high · PI low/high. The PI half-width is `√(σ²·(1+1/T) + q)` and the CI half-width is `√(σ²/T + q)`; same center, same x_new, same t-critical. Sheet layout only — the math lives in the engine. The FE engine drops in group-keyed inputs at the activation step without restructuring this layout.

- TODO: **`Demean_By(x, group, [include])` and `Group_Mean(x, group, [include])`** (constructor internals, also user-callable transforms).

- TODO: **`Is_Balanced_Panel(group, time, [include])`** — one-way/panel diagnostic; ships with `Demean_By` (shares the "valid group set" primitive).

- TODO: **`Absorbed_Degrees_Of_Freedom(spec)`** — Σ(Gᵢ − 1) from the spec.

- TODO: **`y_s()`** — demeaned-Response constructor (new function, not a replacement wired into existing no-FE call sites).

- TODO: **`[DF_Absorbed]` argument (default 0) threaded through df / MS-residual / t-critical.** **Significant testing** — assert bit-equality of every existing engine test with and without the argument, plus FE-active cases for the full inferential chain (SE, t, p, CIs, AIC/BIC).

- TODO: **FE group selection dropdown + ȳᵢ / x̄ᵢ / Tᵢ cells** (gated to ship with the engine). AVERAGEIFS/COUNTIFS respecting the Include/Filter mask; the prediction outputs section from the CI+PI layout activates the group-mean form (ȳᵢ + (x_new − x̄ᵢ)′β̂) with group-keyed inputs; BFN cell (X12/Y12) flips from `n/a — no fixed effects` to active when FE is set. **#2, #3, and this item ship as one release with the engine.**

- TODO: **BFN critical values** (follow-on, ships with 2.1.0 if there's room). N,T-dependent bounds per Bhargava et al. 1982 tables; do NOT present standard DW bounds next to the BFN cell.

- TODO: **Categorical × FE prediction encoding** (follow-on, ships with 2.1.0 if there's room). x_new and x̄ᵢ formed in constructed design-matrix space; UI wire to encode through `Dummy_Code` before reaching the FE formula. Largely subsumed by v2.0 categorical prediction; recorded so the encoding step is not forgotten.

- TODO: **Relabel within-model residual outputs + Diagnostic Guide paragraph on residuals under FE** (follow-on, ships with 2.1.0 if there's room). Documentation-only.

---

## v2.2 — Transforms & the standalone transform library
v2.2 TODO — Formula Strings, Log Transforms, and Unit-Space Prediction

Goal

Make transformed regression models auditable in three layers:

1. Spec-level formula — what the user specified.
2. Fit-space formula — the linear model actually estimated.
3. Unit-space prediction formula — how predictions are returned to the response’s original units.

The implementation must be spec-aware. Display text is not an identifier. The authoritative audit trail is the coefficient metadata, not parsed formula strings.

---

1. Add formula display zones to the Regression sheet

TODO 1.1 — Add bold spec-level formula in row 2

Add a bold formula display above the regression outputs, positioned between the Coefficients zone header and the Prediction Output zone header.

Example:

ln(Cost) ~ ln(Weight) + Country + Status

This formula operates at the source-variable/spec level:

- Continuous predictors show transforms when applicable.
- Categorical predictors appear as source variables, not expanded levels.
- Omitted/reference levels are not shown here.
- This is the preferred formula string for Model Comparison registry rows.

Gate criteria

- Formula updates when Response changes.
- Formula updates when Include toggles change.
- Formula updates when a Continuous predictor transform changes.
- Categorical predictors are not expanded.
- Filter, Identifier, and Omit variables never appear.

---

2. Add coefficient metadata columns

TODO 2.1 — Add coefficient metadata columns to the left of the existing coefficient statistics

Add the following helper/display columns before the existing coefficient estimate/statistic columns:

Column| Purpose
Beta| "β₀", "β₁", … in coefficient-vector order
Term Kind| "Intercept", "Continuous", or "Categorical Level"
Display Term| Human-readable term such as "ln(GDP)" or "I(Country = "Brazil")"
Source Variable| Original source-table column name
Predictor Type| "Continuous" or "Categorical"; blank for intercept
Transform| "None", "Ln", future transform values; blank for intercept and categorical rows
Level| Categorical level represented by this coefficient; blank otherwise
Reference Level| Omitted level for that categorical source variable; blank otherwise
Design Column Index| "0" for intercept, "1..k" for constructed design columns

Do not include Source Role. Coefficient rows are already limited to intercept plus included predictors, so Role is redundant.

Gate criteria

- Intercept row appears as "β₀" only when "Allow_Intercept = TRUE".
- Non-intercept rows align exactly with "x_s()" constructed column order.
- Number of non-intercept coefficient rows equals "COLUMNS(x_s())".
- Categorical rows carry Source Variable, Level, and Reference Level separately.
- Continuous rows carry Source Variable and Transform.
- No formula-string logic parses Display Term to infer meaning.

---

3. Implement coefficient metadata as column-wise spill formulas

TODO 3.1 — Do not build one large "HSTACK" metadata formula

Each metadata field should be its own spill formula or sheet-scoped named helper.

Preferred helper names:

Coefficient_Beta_Symbols
Coefficient_Term_Kinds
Coefficient_Display_Terms
Coefficient_Source_Variables
Coefficient_Predictor_Types
Coefficient_Transforms
Coefficient_Levels
Coefficient_Reference_Levels
Coefficient_Design_Column_Indexes

The coefficient metadata zone may visually form a table, but implementation should remain column-by-column for auditability.

Gate criteria

- Each metadata column can be inspected independently.
- A failure in one metadata field does not obscure all metadata fields.
- The workbook does not rely on one giant "HSTACK" formula for the coefficient metadata table.
- Metadata spills remain aligned row-for-row.

---

4. Define display-term rules

TODO 4.1 — Continuous display terms

For Continuous predictors:

Transform| Display Term
None / blank| "VariableName"
Ln| "ln(VariableName)"

Future transforms should extend this mapping rather than special-case formula strings elsewhere.

TODO 4.2 — Categorical display terms

For Categorical predictors, use:

I(SourceVariable = "Level")

Example:

I(Country = "Brazil")

This is a display convention only. It is not an identifier.

TODO 4.3 — Preserve collision safety through metadata

A user may have a source column literally named:

I(Country = "Brazil")

Do not attempt to solve this by inventing a special display convention. Instead, rely on the structured metadata columns.

Example collision case:

Beta| Term Kind| Display Term| Source Variable| Predictor Type| Level
β₂| Categorical Level| "I(Country = "Brazil")"| "Country"| Categorical| Brazil
β₃| Continuous| "I(Country = "Brazil")"| "I(Country = "Brazil")"| Continuous| —

The display terms may collide; the metadata must not.

Gate criteria

- Formula generation never parses Display Term.
- Source Variable and Level remain separate fields.
- A source column named like an indicator expression remains distinguishable in the metadata.
- Tests include a deliberate collision case.

---

5. Build fit-space formula string

TODO 5.1 — Add fit-space formula display near the prediction section

Add a formula string showing the linear predictor actually estimated.

Example:

ln(Cost) = η = β₀ + β₁ ln(Weight) + β₂ I(Country = "Brazil")

This formula is built from coefficient metadata, not from constructed column names alone.

Gate criteria

- Uses Beta symbols from "Coefficient_Beta_Symbols".
- Uses Display Terms from "Coefficient_Display_Terms".
- Omits intercept term when "Allow_Intercept = FALSE".
- Includes categorical level terms in coefficient-vector order.
- Updates when Include toggles, transforms, references, or categorical levels change.
- Does not parse display text for semantic meaning.

---

6. Build unit-space prediction formula string

TODO 6.1 — Add unit-space formula display near Prediction Output

The unit-space formula keys off the Response transform and Back-transform Method.

If Response transform is blank / None:

ŷ = η

If Response transform is "Ln" and Back-transform Method is "Naive":

ŷ = exp(η)

If Response transform is "Ln" and Back-transform Method is "Duan":

ŷ = S · exp(η), where S = mean(exp(residuals))

The displayed unit-space formula must match the actual prediction calculation.

Gate criteria

- Formula changes when Response Transform changes.
- Formula changes when Back-transform Method changes.
- Non-log response models do not show Duan/Naive as active.
- Log-response models show the selected back-transform method.
- Unit-space formula is visibly tied to "η", not presented as if the model was fit directly in unit space.

---

7. Add Back-transform Method input

TODO 7.1 — Add prediction-section toggle

Add a Prediction Output input cell:

Back-transform Method: Duan / Naive

Default:

Duan

This toggle affects prediction output and the displayed unit-space formula. It does not affect fitted coefficients.

For untransformed response models, the cell should gray out or display an inactive status.

TODO 7.2. - Add Duan_Smear Lambda and surface it in the prediction output area.

TODO 7.3 — Add visible caveat text

Add a short caveat below the prediction output:

Duan = Duan (1983) smearing; Naive = textbook EXP(η), biased for log-response means.

Gate criteria

- Default workbook uses Duan for log-response predictions.
- Naive option remains available.
- Toggle does not affect fit-space coefficients.
- Unit-space prediction output and unit-space formula string update together.
- Caveat is visible on the sheet.

---

8. Add formula-string LAMBDAs / helpers

TODO 8.1 — Add spec-level formula helper

Candidate function:

Model_Spec_Formula_String(
    response_name,
    response_transform,
    spec_names,
    spec_roles,
    spec_includes,
    spec_types,
    spec_transforms
)

Output example:

ln(Cost) ~ ln(Weight) + Country + Status

Rules:

- Use source-variable names.
- Apply transforms only to Continuous variables.
- Do not expand categoricals.
- Exclude non-included predictors.
- Exclude Identifier, Filter, and Omit variables.

TODO 8.2 — Add fit-space formula helper

Candidate function:

Fit_Space_Formula_String(
    response_name,
    response_transform,
    beta_symbols,
    display_terms
)

Output example:

ln(Cost) = η = β₀ + β₁ ln(Weight) + β₂ I(Country = "Brazil")

Inputs should come from coefficient metadata spills, not parsed design-column labels.

TODO 8.3 — Add unit-space formula helper

Candidate function:

Unit_Space_Formula_String(
    response_name,
    response_transform,
    back_transform_method
)

Output examples:

Cost_hat = η
Cost_hat = exp(η)
Cost_hat = S · exp(η), S = mean(exp(residuals))

Gate criteria

- Helpers return readable text, not numeric results.
- Helpers use "NA()" for invalid transform values.
- Helpers handle no-intercept models.
- Helpers handle intercept-only models if the sheet allows them.
- Helpers handle categorical-only models.
- Helpers do not require parsing Display Term.

---

9. Wire Model_Formula_String to the spec-level formula

TODO 9.1 — Define public wrapper for Model Comparison

"Model_Formula_String(anchor_cell)" should return the spec-level formula string, not the expanded fit-space formula.

Reason: Model Comparison needs a compact registry label.

Example:

ln(Cost) ~ ln(Weight) + Country + Status

The expanded coefficient-level formula belongs on the Regression sheet, not in the registry row.

Gate criteria

- "Model_Formula_String(anchor_cell)" reads the target sheet via anchor-cell pattern, not "INDIRECT".
- Returns "NA()" if the anchor does not point to a valid Regression sheet.
- Output matches the row-2 spec formula on the source Regression sheet.
- Function does not depend on raw constructed design-column names.

---

10. Testing requirements

TODO 10.1 — Basic transform tests

Test models:

1. No transforms.
2. Log Response only.
3. Log Predictor only.
4. Log Response and Log Predictor.
5. Mixed transformed and untransformed predictors.

Verify:

- Spec formula.
- Fit-space formula.
- Unit-space formula.
- Prediction back-transform behavior.

TODO 10.2 — Categorical tests

Test models:

1. Single categorical predictor with intercept.
2. Categorical plus continuous predictor.
3. Categorical plus log-transformed continuous predictor.
4. Categorical with explicit reference level.
5. Categorical with default reference level.

Verify:

- Display terms use "I(Source = "Level")".
- Reference level appears in metadata.
- Omitted level has no beta row.
- Intercept note/baseline interpretation remains correct.

TODO 10.3 — Collision tests

Create a source table containing both:

- A categorical variable "Country" with level "Brazil".
- A separate source column literally named "I(Country = "Brazil")".

Verify:

- Both can appear in the coefficient metadata.
- Display terms may match, but Source Variable / Predictor Type / Level distinguish them.
- Formula-string generation does not collapse or deduplicate them incorrectly.
- Coefficient order still matches "x_s()".

TODO 10.4 — Duan / Naive tests

For log-response models, verify:

- Duan prediction equals "EXP(η) * AVERAGE(EXP(residuals))".
- Naive prediction equals "EXP(η)".
- Formula display changes with the toggle.
- Non-log response models ignore or gray out the toggle.

---

11. Documentation updates

TODO 11.1 — Regression Instructions

Add explanation of:

- Spec-level formula.
- Fit-space formula.
- Unit-space prediction formula.
- Why log-response predictions need back-transformation.
- Difference between Duan and Naive.
- Why categorical coefficients are relative to the reference level.

TODO 11.2 — Diagnostic Guide

Add log-model interpretation notes:

- In "ln(y) ~ x", coefficients are semi-elasticity-style effects.
- In "ln(y) ~ ln(x)", coefficients are elasticities.
- In log-response models, "exp(β)" is multiplicative in unit space.
- For categorical predictors in log-response models, "exp(β_level)" is the multiplicative effect relative to the reference level.
- Duan smearing adjusts the predicted mean after exponentiating log-space predictions.

TODO 11.3 — Roadmap cleanup

Move the detailed v2.2 formula-string design discussion out of "ROADMAP.md" and into this TODO section. Leave the roadmap with only the high-level milestone summary and a pointer to this section.

---

12. Final acceptance criteria for v2.2 formula-string work

v2.2 formula-string work is complete when:

- The Regression sheet displays a bold spec-level formula above outputs.
- The coefficient zone includes structured metadata columns.
- The coefficient metadata is produced column-by-column, not by one giant "HSTACK".
- Fit-space and unit-space formulas are visible and update with the model spec.
- Log-response prediction supports Duan and Naive back-transform methods.
- Formula strings are generated from spec/metadata fields, not parsed display labels.
- Collision tests prove that source-column names resembling indicator expressions do not silently corrupt formula interpretation.
- Model Comparison can call "Model_Formula_String(anchor_cell)" and receive the compact spec-level formula.
### Transform wiring (spec column G)

- TODO: `Transform` dropdown gains `Log`; wire `X_s()` / `Constructed_Column_Names()` /
  prediction to read column G.
- TODO: Unit-space fit statistics (R², Adjusted R², RMSE at minimum) — resolve the
  one-LAMBDA-per-combination vs. `Unit_Space_*` dispatcher decision BEFORE implementing
  (sets the pattern for every future transform).
- TODO: Unit-space section on the Regression sheet — SWITCH on column G, one headline
  comparable statistic (the cell v2.3 Model Comparison will reference).
- TODO: Prediction back-transformation — decide naive `EXP()` with documented caveat
  vs. Duan smearing estimator (a statistical decision, not an implementation detail).

### Standalone Data Transformation functions (specs in ROADMAP.md)

- TODO: Location & Scale — `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Ln_Positive`.
- TODO: Group & Panel — `Zscore_By`, `Decompose_By` (`Demean_By`/`Group_Mean` arrive at
  v2.1; two-way functions follow the two-way FE milestone).
- ~~Longitudinal — `Lag_By`, `Difference_By`~~ — **DONE (shipped early, base-period
  release)** with the gap-aware t−Δ semantics: exact-match lookup of
  (group, seq−Δ) pairs, `NA()` at first periods and gaps, `[delta]` defaulting
  to the spec's Base Period Δ cell via `Base_Period_Delta()` (never a silent 1).
  The same release wired spec column I (candidate + override) and the Sequence
  Spacing block (delta spectrum, Regularity/Off-grid flags, calendar-signature
  guidance). Verification: `tests/test_difference_by_verification.py`; human
  test plan T17–T19.
- TODO: Sample construction — `Numeric_Complete_Cases`.
- TODO: Categorical & model construction — `Dummy_Column`, `Interact`, `Model_Matrix`.

---

## v2.3 — Model Comparison Sheet

- TODO: Resolve the spec-string function name (`Regression_Model_Spec_String` vs.
  `Regression_Spec_Label` vs. `Model_Formula_String`) and the argument type (lean:
  anchor-cell reference, not sheet-name text — avoids volatile `INDIRECT`).
- TODO: Implement the spec-string LAMBDA with header-signature validation (`NA()` on
  non-Regression targets).
- TODO: Sheet layout — model registry (hyperlinks), GoF table referencing the v2.2
  unit-space headline cells, shared prediction inputs (Comparison sheet is the source;
  Regression sheets pull via XLOOKUP), prediction results table.
- TODO: Formalize `Comparison_Anchor` sheet-scoped named ranges (interface contract —
  becomes part of the public interface, a versioning commitment).
- TODO: Decide the mismatched-predictor-set fallback (XLOOKUP `[if_not_found]`).

---

## v2.4 — Resampling & Simulation

- TODO: Implement `Bootstrap_CI(data, stat_lambda, n_resamples, alpha, [include])` — bootstrap confidence interval for an arbitrary statistic passed as a LAMBDA. Evaluate whether `RANDARRAY`-based resampling is viable or whether a pre-drawn random table is needed.
- TODO: Implement `MC_Percentile(dist_params, n_samples, percentile)` — Monte Carlo draw from a fitted distribution; complements v2.0 fitting.
- TODO: Implement `PERT_Sample(min, mode, max, n_samples)` — BetaPERT sampling for cost/schedule risk analysis.
- TODO: Design sheet layout (bootstrap section + Monte Carlo section; may share one sheet). Implement `write_sheet_simulation.py`.

---

## v2.5+ — Future (sequence TBD)

### Two-way Fixed Effects (first candidate after v2.1)

- TODO: Implement `Absorb_Two_Way_Fixed_Effects(x, group1, group2, [include], [passes])`
  (alternating-projection demeaning for unbalanced panels).
- TODO: Implement `Demean_Two_Way_Balanced(x, group1, group2, [include])` and the
  two-way `Is_Balanced_Panel` check.
- TODO: Implement `Fixed_Effects_Convergence_Check(x, group1, group2, [include])`;
  surface in the status block whenever two FE variables are active.
- TODO: Lift the v2.1 one-FE-variable status-block error; resolve the two-way
  prediction question (group intercepts are not recoverable as simple group means).

### Weighted regression — superseded by the `Weight` Role

The standalone WLS milestone and its `[weights]`-argument-vs-parallel-function-set
debate are superseded by a **`Weight` value on the Role axis** (see ROADMAP *Future
roles*). Three-stage scope carried forward: user-supplied weights →
variance-driver-derived weights → FGLS.

- TODO: Implement the `Weight` Role (at most one; status-block validation) and thread
  weights through the engine per the Role-axis design.
- TODO: Update the Diagnostic Guide to describe which diagnostics change interpretation
  under WLS. (WLS closes the loop opened by v1's Scale-Location diagnostic.)

### Bivariate / Two-sample

- TODO: Implement `T_Test_OneSample(data, mu0, alpha, [include])` → test statistic, p-value, CI.
- TODO: Implement `T_Test_TwoSample(data1, data2, alpha, equal_var, [include1], [include2])` — equal-variance, Welch, and paired variants via `equal_var` flag.
- TODO: Implement `F_Test_Variance(data1, data2, alpha, [include1], [include2])` — test for equality of variances; output feeds a recommendation cell that selects the appropriate t-test variant.
- TODO: Implement `Covariance_Matrix(data, [include])` — complement to the existing `Correlation_Matrix`.
- TODO: Design two-sample sheet layout: inputs, test selector, F-test assumption check, output (test statistic, df, p-value, CI, effect size). Implement `write_sheet_two_sample.py`.

### Multi-group means (ANOVA)

- TODO: Implement one-way ANOVA as regression on group dummies, reusing the existing SS/MS/F machinery. Frame explicitly as "ANOVA is regression" — group means, SS decomposition, and F-test should match the MLR output exactly.
- TODO: Add post-hoc comparisons (Tukey HSD or Bonferroni) as an optional output section.

### Time series

- TODO: Implement `Moving_Average(data, window, [include])`.
- TODO: Implement `Exponential_Smoothing(data, alpha_smooth, [include])` — note: use `alpha_smooth` to distinguish from the significance-level `alpha`.
- TODO: Implement `write_sheet_time_series.py` with forecast output, error metrics (MAE, RMSE, MAPE), and an actual vs. smoothed series chart.

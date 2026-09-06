<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# Modeling Concepts

The content of the workbook's **Modeling Concepts** sheet, organized
as sections for reading on the web. Generated
from the authored lists in
`lambda_catalog/write_sheet_modeling_concepts.py`.

## Shipped features

### Intercept Control (C2 toggle)

**What Problem It Solves**

The intercept is the model's baseline — the expected response at zero on every predictor. Keep the intercept unless the model requires a zero response when all predictors are zero; the baseline need not itself be a meaningful observed case.

**Statistical Method**

A model-level on/off toggle (cell C2). With reference-coded categoricals in the model, turning it OFF is flagged red: the workbook still drops the reference category, so removing the intercept also constrains the reference baseline to zero. Under Fixed Effects it is flagged because the within transformation demeans the data, and the intercept of demeaned data is an uninterpretable artifact.

**Use Case**

Theory says cost is zero at zero production, or the instrument reads zero with no input? Turn the intercept off for a regression-through-origin fit. Otherwise leave it on and let it absorb the reference levels' baseline.

### Reference Levels (Categorical predictors)

**What Problem It Solves**

A categorical column cannot enter a regression as text — it has to become numbers. But one 0/1 column per level PLUS an intercept is redundant (perfect multicollinearity), so one level must step aside as the baseline everything else is measured against.

**Statistical Method**

Treatment (dummy) coding: a 0/1 column per retained level, with the reference level dropped. Each categorical coefficient reads as a contrast against the reference; the intercept absorbs the reference level's baseline. The default reference is the first-in-sort-order level; type another in column E to override. A typed level not present in the sample is flagged red, never silently fitted.

**Use Case**

Origin (US / Europe / Japan) predicting MPG: with US as the reference, the Europe coefficient is the Europe-vs-US difference in MPG. Pick the level that makes the contrasts meaningful — an experiment's control group, or the market standard your question is 'compared to what?'

### Log Transforms (Transform = Log / Log (drop ≤ 0))

**What Problem It Solves**

Log transforms can represent multiplicative relationships. They change the scale on which the model is fitted and its coefficients are interpreted; they do not guarantee an adequate fit or symmetric residuals.

**Statistical Method**

Only variables marked Log are transformed. Logging a predictor alone leaves the response and residuals in response units. When the response is logged, Naive exponentiates its fitted log value; Duan also multiplies by the average of exp(residuals). This common smearing factor estimates an original-unit conditional mean only when the error retransformation factor is constant across predictor values. It does not generally correct heteroscedasticity. Naive represents a conditional median when log errors have conditional median zero. Original-unit interval bounds always use Naive; they are not confidence bounds for the Duan-adjusted mean. Zeros and negatives have no logarithm: plain Log keeps them in the sample and the fit returns #N/A and flags the Transform cell red; Log (drop ≤ 0) excludes them and reports how many.

**Use Case**

In a log-log model without interactions, a coefficient gives the proportional change in the fitted geometric mean of the response for a proportional change in that predictor, holding other predictors fixed. For small changes, a 1% predictor increase corresponds to approximately that coefficient percent change in the fitted response. Inspect residual plots after transformation.

### Sample Filtering & Completeness (Role = Filter)

**What Problem It Solves**

The fit should answer a question about a specific population, and rows with missing values cannot be part of it. The workbook derives the analysis sample from the spec, so a change of population is a specification change — not hand-deleting rows on the data sheet.

**Statistical Method**

A per-row mask ANDed together: every Filter column must be TRUE, the Response must be numeric, and every included Continuous predictor must be numeric (listwise deletion). Categorical predictors impose no completeness condition — a blank category is just not a level. Log (drop ≤ 0) adds its own exclusion layer, and the excluded-row counts surface in the status cells above the spec (B2 / G2) rather than staying silent.

**Use Case**

Restrict a nationwide demand model to one market segment: add a derived column that is TRUE only for that segment, set its Role to Filter, and every statistic is refit to that population. Retarget the filter later and the sample follows — the data sheet is never touched.

### Interactions (Interaction Term / Operation, M/N)

**What Problem It Solves**

One predictor's effect can depend on another's value — a discount that lands differently by market, a drug that scales with dosage. A purely additive model cannot express that dependency; an interaction column can.

**Statistical Method**

Pairwise constructed columns in three operations: Product (symmetric), Difference (antisymmetric), Ratio (asymmetric). Width follows the operands: Continuous × Continuous adds 1 column, Continuous × Categorical adds L−1, Categorical × Categorical adds (L₁−1)(L₂−1) — all counted in the Design Columns audit. Interacting a variable with itself under Product is the documented quadratic. An interaction whose main effect is switched off is flagged amber (marginality); declaring both A×B and B×A under Product or Difference is flagged red (duplicate column, singular matrix).

**Use Case**

Does advertising work equally in every region? Interact Ad Spend (Continuous) with Region (Categorical): the design gains one ad-sensitivity slope per region vs. the reference, and each coefficient answers for its own region. For diminishing returns, interact x with itself under Product to add x².

### Fixed Effects (Role = Fixed Effects)

**What Problem It Solves**

Groups differ in ways you never measured — plant conditions, firm culture, country policy — and those stable differences can drive both the response and the predictors, biasing coefficients. Fixed Effects absorbs a separate intercept for each group so the group differences stop competing with the predictors you actually care about.

**Statistical Method**

One-way within transformation: every variable is restated as a deviation from its group's mean, which removes the group intercepts from the fit entirely. Algebraically equal to LSDV (a dummy column per group) without the columns, and the absorbed group degrees of freedom are credited back into inference.

**Use Case**

Modelling salaries across 12 plants: plant-level pay practices drive both salary and who works where. Set Plant to Role = Fixed Effects and the coefficients answer the within-plant question — holding the plant fixed, what does this predictor change? — with no plant dummy columns cluttering the coefficient table.

### Sequence Effects (Sequence flag, column H)

**What Problem It Solves**

Panel and repeated-measures data can violate the independence assumption — rows within one unit follow each other in time. Lags, differences, and serial-correlation diagnostics all need to know which variable defines that ordering, which is what the Sequence flag declares.

**Statistical Method**

Within-group exact-time matching: Lag_By and Difference_By find each group's prior period by matching its time value, never by row position, so a gap in the panel yields #N/A instead of a silently wrong neighbor. The Base Period Δ (your typed override, or the computed candidate — the most common within-group spacing) sets the period step, and the BFN Panel Durbin-Watson tests for serial correlation within groups.

**Use Case**

Yearly observations per country: set Year to Sequence = TRUE and a year-over-year difference column gives each country's change between its own consecutive years. The Sequence Spacing verdict above the spec tells you whether the panel is regular enough for those features to mean what you think they mean.

## Planned features

Declared in the sheet's structure but not built in this release — see
the workbook sheet for the same rows.

### Weight — WLS (Role = Weight, v3.7) PLANNED, future release

**What Problem It Solves**

Some observations are known to be noisier than others — each row averages a different number of underlying readings, or a variance you can quantify at row level. OLS gives the noisy rows the same vote as the precise ones.

**Statistical Method**

Weighted Least Squares: each row enters the fit scaled by its known variance, so precise observations dominate. The Weight Role will name the column carrying that variance.

**Use Case**

Today's workaround is the Diagnostic Guide's heteroscedasticity guidance: a Log transform on the response, or an interaction / quadratic term when the variance tracks a predictor's mean.

### Cluster (Role = Cluster, v3.5) PLANNED, future release

**What Problem It Solves**

Rows collected in groups — plants, firms, countries — share unmodeled shocks, so nominally-independent standard errors are too optimistic (they overstate significance).

**Statistical Method**

Clustered-robust variance estimator: the same coefficients, with standard errors that allow arbitrary correlation within a group. Partially forward-wired — the Role slot and a dormant branch in Serial_Correlation_Group() await the estimator.

**Use Case**

Survey or panel data where rows within one cluster move together: name the cluster column and the P-values and confidence intervals stop assuming every row is independent.

### Time (Role = Time, v3.6) PLANNED, future release

**What Problem It Solves**

A panel needs an explicit time index distinct from 'which group a row belongs to' — lag and difference features must follow time within a unit, not sheet order.

**Statistical Method**

Time-index designation: the Role names the column holding each row's time value; Lag_By / Difference_By and the future Time Series sheet are its consumers. The Role ships first, the sheet last.

**Use Case**

Until then, the Sequence flag (column H) plus its Sequence Period override already carries the ordering axis for the shipped lag/difference features.

### Two-way Fixed Effects (v3.8) PLANNED, future release

**What Problem It Solves**

Sometimes the stable nuisance is two-dimensional — both the plant AND the year leave a fixed mark on every observation, and absorbing only one of them leaves the other confounding the coefficients.

**Statistical Method**

Absorb_Two_Way_Fixed_Effects: the within transformation applied along two group axes at once. Today at most ONE Fixed Effects row is legal — two or more is a red spec error, by design, until this lands.

**Use Case**

Plant-and-year panels: for now, keep one Fixed Effects row and use a Filter to narrow the sample, or dummy-code the second axis as a Categorical predictor.

### Order (column F) PLANNED, future release

**What Problem It Solves**

Term order in the coefficient table currently follows the source table's column order — fine for every model so far, but a hand-authored spec may want terms presented in a chosen sequence.

**Statistical Method**

A reserved, hidden column (width 0): it will control each spec row's position in the constructed design matrix and the coefficient table. No formula reads it in this release.

**Use Case**

Nothing to do today — the column is hidden and ignored; leave it blank if you reveal it.

## Further reading

[Stata: interpreting log-transformed outcomes](https://www.stata.com/stata-news/news34-2/spotlight/) explains why exponentiating a log prediction does not in general estimate an arithmetic mean.

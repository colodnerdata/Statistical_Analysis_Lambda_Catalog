<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# Diagnostic Guide

The same tiers and thresholds as the workbook's **Diagnostic Guide**
sheet, rendered from the authored lists in
`lambda_catalog/write_sheet_diagnostic_guide.py`.

Use the charts and flagged cells on the Regression sheet to assess
model assumptions. Start with Tier 1, then review spread and influence
in Tier 2 even if the first plots show no obvious problem.

The thresholds below are the workbook's screening rules, not universal
tests of model validity. In particular, Q-Q correlation cutoffs are not
sample-size-adjusted normality tests. Consider the sampling design and
possible dependence as well as the plots.

## Tier 1 — review for every model

### Residuals vs. Fitted

**X-axis**

Predicted Y

**Y-axis**

Residuals

**What to look for**

Random scatter around zero. A curve or funnel shape signals nonlinearity or heteroscedasticity (non-constant error variance).

### Normal Q-Q

**X-axis**

Normal Scores (theoretical)

**Y-axis**

Studentized Residuals Ranked

**What to look for**

Points close to a straight diagonal line. Heavy tails or an S-curve indicate non-normal errors. Check QQ Correlation (Cell AE11): below 0.98 = mild concern, below 0.95 = stronger concern. The chart title itself shows the live r and appends a '— check normality' warning when it falls below 0.95, matching the table's red threshold.

### Actual vs. Predicted

**X-axis**

Predicted Y

**Y-axis**

Y

**What to look for**

Points close to a 45° line. A bow or fan shape can suggest the same problems seen in Residuals vs. Fitted but from a different angle. The chart title shows the response name and the live Adjusted R², so the fit's headline quality is on the plot itself.

## Tier 2 — review spread and influence

### Scale-Location

**X-axis**

Predicted Y

**Y-axis**

Scale-Location (√|Studentized Residuals|)

**What to look for**

A roughly horizontal spread is consistent with constant error variance. A trend suggests changing variance; inspect it alongside Residuals vs. Fitted. Yellow cells: value > √2 ≈ 1.41; red cells: value > √3 ≈ 1.73.

### Cook's Distance

**X-axis**

Observation (bar position)

**Y-axis**

Cook's Distance

**What to look for**

Spikes above F.INV(0.5, p, n-p) — the median of the reference F distribution, so the bar to beat scales with the model's own size — mark observations with outsized influence on the fitted coefficients. Bars are ordered by observation number. The title prints the live cutoff, so the bar to beat is visible on the plot itself, and bars above it carry data labels naming the row. Inspect those rows for data entry errors or genuine outliers before removing them. Remove outliers only when you have a definitive, non-statistical reason to believe the data point is invalid, comes from a population excluded by the study definition, or is otherwise outside the prespecified analysis scope. Never filter out data solely because it is an extreme statistical value, as doing so can introduce bias and erase genuine insights.

### Studentized Residuals vs. Leverage

**X-axis**

Hat Diagonal (leverage)

**Y-axis**

Studentized Residuals

**What to look for**

High leverage alone is not a problem. Combined high leverage and large residual (top-right or bottom-right of the plot) = influential outlier. Hat > 2p/n is flagged red; Hat > 3p/n is additionally bold. The title prints the live mean leverage (p/n), so 'Hat > 2p/n' reads directly as 'more than twice the printed mean.'

### PRESS Residuals

**X-axis**

Observation (bar position)

**Y-axis**

PRESS Residual (e / (1 − h))

**What to look for**

PRESS residuals inflate the ordinary residual by leverage. Bars are ordered by observation number. Large values (|PRESS| > 2 × SE yellow, > 3 × SE red) flag observations whose removal would substantially shift the fitted model. The chart title carries the live PRESS total.

## Diagnostic threshold reference

| Diagnostic | Location on sheet | Yellow threshold | Red threshold |
|---|---|---|---|
| GVIF (Generalized Variance Inflation Factor) | Col U, Predictor Summary | GVIF > 5  (possible collinearity) | GVIF > 10  (strong collinearity) |
| Tolerance | Col V, Predictor Summary | Tolerance < 0.2 | Tolerance < 0.1 |
| PRESS R² | Cell AE6, Diagnostics | — | PRESS R² < 0  (worse than predicting Y-mean) |
| QQ Correlation | Cell AE11, Diagnostics | < 0.98  (mild non-normality) | < 0.95  (clear non-normality) |
| Significance F | Cell AF16, ANOVA Table | — | P-value > alpha  (model not significant) |
| Coefficient P-values | Col AE, Coefficients | — | P-value > alpha  (term not significant) |
| Hat Diagonal (leverage) | Col AR, Residual Output | — | h > 2p/n  (high leverage) |
| Studentized Residuals | Col AS, Residual Output | \|r*\| > 2  (moderate outlier) | \|r*\| ≥ 3  (strong outlier) |
| Cook's Distance | Col AT, Residual Output | — | D > F.INV(0.5, p, n-p)  (high influence) |
| Scale-Location | Col AW, Residual Output | > √2 ≈ 1.41  (\|r*\| > 2 equivalent) | > √3 ≈ 1.73  (\|r*\| > 3 equivalent) |
| PRESS Residual | Col AX, Residual Output | \|PRESS\| > 2 × SE | \|PRESS\| > 3 × SE |

## Common patterns and next steps

### Heteroscedasticity (funnel residuals)

**Symptom**

Residuals vs. Fitted shows fan shape; Scale-Location trends upward.

**Next step**

Consider: log-transforming Y, adding polynomial terms, or fitting a Weighted Least Squares (WLS) model. WLS is planned for a future version of this library.

### Nonlinearity (curved residuals)

**Symptom**

Residuals vs. Fitted shows a curve; Actual vs. Predicted bows.

**Next step**

Add squared or interaction terms to the model. Toggle individual predictors on/off to isolate which variable is driving the curve.

### Non-normal errors (Q-Q deviation)

**Symptom**

Q-Q plot shows heavy tails or S-curve; QQ Correlation < 0.98.

**Next step**

Assess the severity of the departure and the sampling design; there is no universal sample-size cutoff that makes it harmless. Robust standard errors address variance misspecification, not all non-normality. A bootstrap must respect dependence in the data. These alternatives require tools beyond this template.

### Influential observations (high Cook's D or PRESS)

**Symptom**

Cook's D > F.INV(0.5, p, n-p) or |PRESS| > 2 × SE for one or more rows.

**Next step**

Inspect those rows. Verify data entry. Refit without the observation(s) and compare coefficients — if they shift substantially, see which ones and research the reasons why those data points are different. If there's a significant difference, choose one model, but report the prediction results of the other regression as a sensitivity analysis.

### Multicollinearity (high GVIF)

**Symptom**

GVIF > 10 for one or more predictors. A categorical predictor's dummy columns all show the same shared GVIF value — that's the whole variable's collinearity, not a per-level artifact.

**Next step**

Remove or combine correlated predictors. Use the Correlation_Matrix function to identify the pairs. Partial R² shows each predictor's unique contribution after controlling for the others.

## Further reading

[NIST: checking regression assumptions](https://itl.nist.gov/div898/handbook/pri/section2/pri245.htm) discusses residual shape, variance, and independence.

[NIST: regression diagnostics](https://www.itl.nist.gov/div898/software/dataplot/refman1/auxillar/regrdiag.htm) describes leverage and influence measures.

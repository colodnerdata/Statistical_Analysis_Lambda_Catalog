<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# Diagnostic Guide

The same tiers and thresholds as the workbook's **Diagnostic Guide**
sheet, rendered from the authored lists in
`lambda_catalog/write_sheet_diagnostic_guide.py`.

Use the charts and flagged cells on the Regression sheet to assess
model assumptions. Work through Tier 1 first; investigate Tier 2
only when a Tier 1 plot raises a concern.

## Tier 1 — review for every model

| Plot | X-axis | Y-axis | What to look for |
|---|---|---|---|
| Residuals vs. Fitted | Predicted Y | Residuals | Random scatter around zero. A curve or funnel shape signals nonlinearity or heteroscedasticity (non-constant error variance). |
| Normal Q-Q | Normal Scores (theoretical) | Studentized Residuals Ranked | Points close to a straight diagonal line. Heavy tails or an S-curve indicate non-normal errors. Check QQ Correlation (Cell P10): below 0.98 = mild concern, below 0.95 = stronger concern. The chart title itself shows the live r and appends a '— check normality' warning when it falls below 0.95, matching the table's red threshold. |
| Actual vs. Predicted | Predicted Y | Y | Points close to a 45° line. A bow or fan shape confirms the same problems seen in Residuals vs. Fitted but from a different angle. The chart title shows the response name and the live Adjusted R², so the fit's headline quality is on the plot itself. |

## Tier 2 — investigate when Tier 1 raises a concern

| Plot | X-axis | Y-axis | What to look for |
|---|---|---|---|
| Scale-Location | Predicted Y | Scale-Location (√\|Studentized Residuals\|) | Flat horizontal spread of points = homoscedasticity. An upward trend confirms heteroscedasticity flagged in Tier 1. Yellow cells: value > √2 ≈ 1.41; red cells: value > √3 ≈ 1.73. |
| Cook's Distance | Observation (bar position) | Cook's Distance | Spikes above F.INV(0.5, p, n-p) — the median of the reference F distribution, so the bar to beat scales with the model's own size — mark observations with outsized influence on the fitted coefficients. Bars are ordered by observation number. The title prints the live cutoff, so the bar to beat is visible on the plot itself, and bars above it carry data labels naming the row. Inspect those rows for data entry errors or genuine outliers before removing them. Remove outliers only when you have a definitive, non-statistical reason to believe the data point is invalid, comes from a different population, or disproportionately distorts the analysis. Never filter out data solely because it is an extreme statistical value, as doing so can introduce bias and erase genuine insights. |
| Studentized Residuals vs. Leverage | Hat Diagonal (leverage) | Studentized Residuals | High leverage alone is not a problem. Combined high leverage and large residual (top-right or bottom-right of the plot) = influential outlier. Hat > 2p/n is flagged red; Hat > 3p/n is additionally bold. The title prints the live mean leverage (p/n), so 'Hat > 2p/n' reads directly as 'more than twice the printed mean.' |
| PRESS Residuals | Observation (bar position) | PRESS Residual (e / (1 − h)) | PRESS residuals inflate the ordinary residual by leverage. Bars are ordered by observation number. Large values (\|PRESS\| > 2 × SE yellow, > 3 × SE red) flag observations whose removal would substantially shift the fitted model. The chart title carries the live PRESS total. |

## Diagnostic threshold reference

| Diagnostic | Location on sheet | Yellow threshold | Red threshold |
|---|---|---|---|
| GVIF (Generalized Variance Inflation Factor) | Col U, Predictor Summary | GVIF > 5  (possible collinearity) | GVIF > 10  (strong collinearity) |
| Tolerance | Col V, Predictor Summary | Tolerance < 0.2 | Tolerance < 0.1 |
| PRESS R² | Cell P5, Diagnostics | — | PRESS R² < 0  (worse than predicting Y-mean) |
| QQ Correlation | Cell P10, Diagnostics | < 0.98  (mild non-normality) | < 0.95  (clear non-normality) |
| Significance F | Cell Q15, ANOVA Table | — | P-value > alpha  (model not significant) |
| Coefficient P-values | Col AE, Coefficients | — | P-value > alpha  (term not significant) |
| Hat Diagonal (leverage) | Col AR, Residual Output | — | h > 2p/n  (high leverage) |
| Studentized Residuals | Col AS, Residual Output | \|r*\| > 2  (moderate outlier) | \|r*\| ≥ 3  (strong outlier) |
| Cook's Distance | Col AT, Residual Output | — | D > F.INV(0.5, p, n-p)  (high influence) |
| Scale-Location | Col AW, Residual Output | > √2 ≈ 1.41  (\|r*\| > 2 equivalent) | > √3 ≈ 1.73  (\|r*\| > 3 equivalent) |
| PRESS Residual | Col AX, Residual Output | \|PRESS\| > 2 × SE | \|PRESS\| > 3 × SE |

## Common patterns and next steps

| Pattern | Symptom | Next step |
|---|---|---|
| Heteroscedasticity (funnel residuals) | Residuals vs. Fitted shows fan shape; Scale-Location trends upward. | Consider: log-transforming Y, adding polynomial terms, or fitting a Weighted Least Squares (WLS) model. WLS is planned for a future version of this library. |
| Nonlinearity (curved residuals) | Residuals vs. Fitted shows a curve; Actual vs. Predicted bows. | Add squared or interaction terms to the model. Toggle individual predictors on/off to isolate which variable is driving the curve. |
| Non-normal errors (Q-Q deviation) | Q-Q plot shows heavy tails or S-curve; QQ Correlation < 0.98. | With n > 100 moderate departures rarely invalidate inference. For small n, consider robust standard errors or a bootstrap Confidence Interval approach. |
| Influential observations (high Cook's D or PRESS) | Cook's D > F.INV(0.5, p, n-p) or \|PRESS\| > 2 × SE for one or more rows. | Inspect those rows. Verify data entry. Refit without the observation(s) and compare coefficients — if they shift substantially, see which ones and research the reasons why those data points are different. If there's a significant difference, choose one model, but report the prediction results of the other regression as a sensitivity analysis. |
| Multicollinearity (high GVIF) | GVIF > 10 for one or more predictors. A categorical predictor's dummy columns all show the same shared GVIF value — that's the whole variable's collinearity, not a per-level artifact. | Remove or combine correlated predictors. Use the Correlation_Matrix function to identify the pairs. Partial R² shows each predictor's unique contribution after controlling for the others. |

# Statistical Analysis Lambda Catalog

Excel 365 LAMBDA functions that replicate and extend Analysis ToolPak regression statistics — no VBA, no add-ins, no installation. Download `Lambda_Library.xlsx`, open it in Excel, and all functions are immediately available in formulas.

## Getting started

1. Download `Lambda_Library.xlsx` from this repository.
2. Open it in Excel 365 (Windows or Mac).
3. Enter your data in columns on any sheet, then call any function by name.

All functions are defined as workbook-scoped names, so they work in any cell formula within the workbook. The **Regression** sheet provides a ready-to-use analysis interface: select your predictor and outcome columns, and it computes the full regression output automatically. The **LAMBDA_functions** sheet is a browsable catalog of every function with its full description and argument documentation.

To use these functions in a different workbook, open both files in Excel at the same time. You can reference functions as `='[Lambda_Library.xlsx]'!FunctionName(args)`, or use Name Manager (Formulas → Name Manager → New) to copy individual definitions into your own workbook.

## Common function signature

Most regression functions share this signature:

```
FunctionName(X_s, Y, [Allow_Intercept], [Filter])
```

| Argument | Description |
|---|---|
| `X_s` | Predictor column range (one or more columns) |
| `Y` | Outcome column range (single column) |
| `[Allow_Intercept]` | TRUE to fit an intercept (default), FALSE to force through the origin |
| `[Filter]` | Boolean column — TRUE includes the row, FALSE excludes it |

Square brackets indicate optional arguments. A few functions have different signatures, as noted in their entries below.

## Function reference

### Model fit — scalar outputs

These functions return a single number summarizing the regression.

| Function | Returns |
|---|---|
| `Observations(Y, [Filter])` | Number of observations (n) |
| `DF_Regression(X_s)` | Degrees of freedom for the model (k predictors) |
| `DF_Residual(X_s, Y, [Allow_Intercept], [Filter])` | Degrees of freedom for residuals (n − k − 1) |
| `DF_Total(Y, [Allow_Intercept], [Filter])` | Total degrees of freedom (n − 1) |
| `R_squared(X_s, Y, [Allow_Intercept], [Filter])` | R² — proportion of variance explained, 0 to 1 |
| `Multiple_R(X_s, Y, [Allow_Intercept], [Filter])` | √R² — multiple correlation coefficient |
| `Adjusted_R2(X_s, Y, [Allow_Intercept], [Filter])` | R² penalized for number of predictors |
| `SE_Regression(X_s, Y, [Allow_Intercept], [Filter])` | Standard error of the regression |
| `SS_Regression(X_s, Y, [Allow_Intercept], [Filter])` | Model sum of squares |
| `SS_Residual(X_s, Y, [Allow_Intercept], [Filter])` | Residual (error) sum of squares |
| `SS_Total(Y, [Allow_Intercept], [Filter])` | Total sum of squares |
| `PRESS(X_s, Y, [Allow_Intercept], [Filter])` | Leave-one-out cross-validation error sum |
| `Durbin_Watson(X_s, Y, [Allow_Intercept], [Filter])` | Serial autocorrelation test for residuals (2 = no autocorrelation) |
| `AIC(X_s, Y, [Allow_Intercept], [Filter])` | Akaike Information Criterion |
| `AICc(X_s, Y, [Allow_Intercept], [Filter])` | Corrected AIC (for small samples) |
| `BIC(X_s, Y, [Allow_Intercept], [Filter])` | Bayesian Information Criterion |
| `QQ_Correlation(X_s, Y, [Allow_Intercept], [Filter])` | Filliben Q-Q normality statistic for residuals |

### Coefficient-level outputs — (k+1)×1 vector

These functions return one value per coefficient (intercept first when included, then predictors in input order). Enter them as array formulas and let them spill.

| Function | Returns |
|---|---|
| `Coefficients(X_s, Y, [Allow_Intercept], [Filter])` | OLS coefficient estimates |
| `SE_Coefficients(X_s, Y, [Allow_Intercept], [Filter])` | Standard errors of the coefficients |
| `T_Stats(X_s, Y, [Allow_Intercept], [Filter])` | t-statistics (coefficient / standard error) |
| `P_Values(X_s, Y, [Allow_Intercept], [Filter])` | Two-tailed p-values |
| `CI_Lower(X_s, Y, [Allow_Intercept], [Filter], [Alpha])` | Lower confidence interval bounds (default 95%) |
| `CI_Upper(X_s, Y, [Allow_Intercept], [Filter], [Alpha])` | Upper confidence interval bounds (default 95%) |
| `Partial_R2(X_s, Y, [Allow_Intercept], [Filter])` | Partial R² per coefficient |
| `Partial_Correlation(X_s, Y, [Allow_Intercept], [Filter])` | Partial correlation per coefficient |

### Multicollinearity — k×1 vector

These functions return one value per predictor (no intercept row). `VIF > 10` or `Tolerance < 0.1` indicates severe collinearity.

| Function | Returns |
|---|---|
| `VIF(X_s, [Allow_Intercept], [Filter])` | Variance inflation factor per predictor |
| `Tolerance(X_s, [Allow_Intercept], [Filter])` | Tolerance (= 1 / VIF) per predictor |

### Observation-level outputs — n×1 vector

These functions return one value per observation in the filtered dataset, spilled in original data order.

| Function | Returns |
|---|---|
| `Predictions(X_s, Y, [Allow_Intercept], [Filter])` | Fitted (predicted) values |
| `Residuals(X_s, Y, [Allow_Intercept], [Filter])` | Raw residuals (Y − Ŷ) |
| `Scaled_Residuals(X_s, Y, [Allow_Intercept], [Filter])` | Residuals divided by SE_Regression |
| `Scaled_Residuals_Ranked(X_s, Y, [Allow_Intercept], [Filter])` | Scaled residuals sorted ascending |
| `Studentized_Residuals(X_s, Y, [Allow_Intercept], [Filter])` | Internally studentized residuals |
| `Studentized_Residuals_Ranked(X_s, Y, [Allow_Intercept], [Filter])` | Studentized residuals sorted ascending |
| `Hat_diagonal(X_s, [Allow_Intercept], [Filter])` | Leverage values (diagonal of the hat matrix) |
| `Cooks_Distance(X_s, Y, [Allow_Intercept], [Filter])` | Cook's distance per observation |
| `LOOCV_prediction(X_s, Y, [Allow_Intercept], [Filter])` | Leave-one-out predicted value per observation |
| `Observation_Num(Y, [Filter])` | Sequential row indices 1 through n |
| `Rank_Fraction(Y, [Filter])` | Empirical CDF fractions in original data order |
| `Y_Ranked(Y, [Filter])` | Sorted filtered outcome values |
| `Normal_Scores(Y, [Filter])` | Theoretical standard-normal quantiles |

### Prediction interval

`Prediction_Interval(X_s, Y, X_new, [Allow_Intercept], [Filter], [alpha])` returns a 6-element vertical array: point estimate, SE of prediction, critical t value, lower bound, upper bound, and confidence level. `X_new` is a single-row range of predictor values for the new observation.

### Model significance — scalar outputs

These functions return the overall F-test results for the regression model.

| Function | Returns |
|---|---|
| `F_Stat(X_s, Y, [Allow_Intercept], [Filter])` | F-statistic for overall model significance |
| `P_Value_F(X_s, Y, [Allow_Intercept], [Filter])` | p-value for the overall F-test |

### Descriptive and correlation functions

These functions support exploratory analysis before model fitting. `Pearson_R`, `Spearman_R`, `Skewness`, and `Kurtosis` accept a multi-column `X_s` and return a k×1 array — one value per predictor. Comparing `Pearson_R` and `Spearman_R` for the same predictor diagnoses nonlinearity: a large gap between the two suggests curvature that a linear model will not capture.

| Function | Signature | Returns |
|---|---|---|
| `Pearson_R` | `(X_s, Y, [Filter])` | k×1 vector of Pearson correlations between each predictor and Y |
| `Spearman_R` | `(X_s, Y, [Filter])` | k×1 vector of Spearman rank correlations between each predictor and Y |
| `Skewness` | `(X_s, [Filter])` | k×1 vector of skewness for each predictor column |
| `Kurtosis` | `(X_s, [Filter])` | k×1 vector of excess kurtosis (Fisher convention) for each predictor column |
| `Correlation_Matrix` | `(X_s, [Filter])` | k×k symmetric matrix of pairwise Pearson correlations among predictors |

### Utility functions

| Function | Returns |
|---|---|
| `Design_Matrix(X_s, [Allow_Intercept], [Filter])` | Filtered numeric design matrix as a spilled array |
| `Complete_Cases_Filter(X_s, [Y])` | Boolean column — TRUE for rows with no missing values |
| `Col_Select(table, col_nums)` | Selected columns from an array in the specified order |
| `LOO_prediction(X_s, Y, n, [Allow_Intercept], [Filter])` | Leave-one-out prediction for a single observation n |
| `This_row(array)` | 1-to-n relative row indices |
| `Exclude_row_n(array, n)` | Array with row n removed |

## Sample data and Regression sheet

`Lambda_Library.xlsx` includes the WHO Life Expectancy dataset (2,938 rows across 193 countries, 2000–2015) as a structured table on the **Life Expectancy Data** sheet. The **Regression** sheet uses this dataset to demonstrate a full multiple regression analysis: select any subset of the 18 health and economic predictors, and the sheet recomputes all statistics instantly using the LAMBDA functions.

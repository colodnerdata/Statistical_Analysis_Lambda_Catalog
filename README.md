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
FunctionName(X_s, Y, [Allow_Intercept], [Include])
```

| Argument | Description |
|---|---|
| `X_s` | Predictor column range (one or more columns) |
| `Y` | Outcome column range (single column) |
| `[Allow_Intercept]` | TRUE to fit an intercept (default), FALSE to force through the origin |
| `[Include]` | Boolean column — TRUE includes the row, FALSE excludes it |

Square brackets indicate optional arguments. A few functions have different signatures, as noted in their entries below.

## Function reference

### Model fit — scalar outputs

These functions return a single number summarizing the regression.

| Function | Returns |
|---|---|
| `Observations(Y, [Include])` | Number of observations (n) |
| `DF_Regression(X_s)` | Degrees of freedom for the model (k predictors) |
| `DF_Residual(X_s, Y, [Allow_Intercept], [Include])` | Degrees of freedom for residuals (n − k − 1) |
| `DF_Total(Y, [Allow_Intercept], [Include])` | Total degrees of freedom (n − 1) |
| `R_squared(X_s, Y, [Allow_Intercept], [Include])` | R² — proportion of variance explained, 0 to 1 |
| `Multiple_R(X_s, Y, [Allow_Intercept], [Include])` | √R² — multiple correlation coefficient |
| `Adjusted_R2(X_s, Y, [Allow_Intercept], [Include])` | R² penalized for number of predictors |
| `SE_Regression(X_s, Y, [Allow_Intercept], [Include])` | Standard error of the regression |
| `SS_Regression(X_s, Y, [Allow_Intercept], [Include])` | Model sum of squares |
| `SS_Residual(X_s, Y, [Allow_Intercept], [Include])` | Residual (error) sum of squares |
| `SS_Total(Y, [Allow_Intercept], [Include])` | Total sum of squares |
| `PRESS(X_s, Y, [Allow_Intercept], [Include])` | Leave-one-out cross-validation error sum |
| `Durbin_Watson(X_s, Y, [Allow_Intercept], [Include])` | Serial autocorrelation test for residuals (2 = no autocorrelation) |
| `AIC(X_s, Y, [Allow_Intercept], [Include])` | Akaike Information Criterion |
| `AICc(X_s, Y, [Allow_Intercept], [Include])` | Corrected AIC (for small samples) |
| `BIC(X_s, Y, [Allow_Intercept], [Include])` | Bayesian Information Criterion |
| `QQ_Correlation(X_s, Y, [Allow_Intercept], [Include])` | Filliben Q-Q normality statistic for residuals |

### Coefficient-level outputs — (k+1)×1 vector

These functions return one value per coefficient (intercept first when included, then predictors in input order). Enter them as array formulas and let them spill.

| Function | Returns |
|---|---|
| `Coefficients(X_s, Y, [Allow_Intercept], [Include])` | OLS coefficient estimates |
| `SE_Coefficients(X_s, Y, [Allow_Intercept], [Include])` | Standard errors of the coefficients |
| `T_Stats(X_s, Y, [Allow_Intercept], [Include])` | t-statistics (coefficient / standard error) |
| `P_Values(X_s, Y, [Allow_Intercept], [Include])` | Two-tailed p-values |
| `CI_Lower(X_s, Y, [Allow_Intercept], [Include], [Alpha])` | Lower confidence interval bounds (default 95%) |
| `CI_Upper(X_s, Y, [Allow_Intercept], [Include], [Alpha])` | Upper confidence interval bounds (default 95%) |
| `Partial_R2(X_s, Y, [Allow_Intercept], [Include])` | Partial R² per coefficient |
| `Partial_Correlation(X_s, Y, [Allow_Intercept], [Include])` | Partial correlation per coefficient |

### Multicollinearity — k×1 vector

These functions return one value per predictor (no intercept row). `VIF > 10` or `Tolerance < 0.1` indicates severe collinearity.

| Function | Returns |
|---|---|
| `VIF(X_s, [Allow_Intercept], [Include])` | Variance inflation factor per predictor |
| `Tolerance(X_s, [Allow_Intercept], [Include])` | Tolerance (= 1 / VIF) per predictor |

### Observation-level outputs — n×1 vector

These functions return one value per observation in the filtered dataset, spilled in original data order.

| Function | Returns |
|---|---|
| `Predictions(X_s, Y, [Allow_Intercept], [Include])` | Fitted (predicted) values |
| `Residuals(X_s, Y, [Allow_Intercept], [Include])` | Raw residuals (Y − Ŷ) |
| `Scaled_Residuals(X_s, Y, [Allow_Intercept], [Include])` | Residuals divided by SE_Regression |
| `Scaled_Residuals_Ranked(X_s, Y, [Allow_Intercept], [Include])` | Scaled residuals sorted ascending |
| `Studentized_Residuals(X_s, Y, [Allow_Intercept], [Include])` | Internally studentized residuals |
| `Studentized_Residuals_Ranked(X_s, Y, [Allow_Intercept], [Include])` | Studentized residuals sorted ascending |
| `Hat_diagonal(X_s, [Allow_Intercept], [Include])` | Leverage values (diagonal of the hat matrix) |
| `Cooks_Distance(X_s, Y, [Allow_Intercept], [Include])` | Cook's distance per observation |
| `LOOCV_prediction(X_s, Y, [Allow_Intercept], [Include])` | Leave-one-out predicted value per observation |
| `Observation_Num(Y, [Include])` | Sequential row indices 1 through n |
| `Rank_Fraction(Y, [Include])` | Empirical CDF fractions in original data order |
| `Y_Ranked(Y, [Include])` | Sorted filtered outcome values |
| `Normal_Scores(Y, [Include])` | Theoretical standard-normal quantiles |

### Prediction interval

`Prediction_Interval(X_s, Y, X_new, [Allow_Intercept], [Include], [alpha])` returns a 6-element vertical array: point estimate, SE of prediction, critical t value, lower bound, upper bound, and confidence level. `X_new` is a single-row range of predictor values for the new observation.

### Model significance — scalar outputs

These functions return the overall F-test results for the regression model.

| Function | Returns |
|---|---|
| `F_Stat(X_s, Y, [Allow_Intercept], [Include])` | F-statistic for overall model significance |
| `P_Value_F(X_s, Y, [Allow_Intercept], [Include])` | p-value for the overall F-test |

### Descriptive and correlation functions

These functions support exploratory analysis before model fitting. `Pearson_R`, `Spearman_R`, `Skewness`, and `Kurtosis` accept a multi-column `X_s` and return a k×1 array — one value per predictor. Comparing `Pearson_R` and `Spearman_R` for the same predictor diagnoses nonlinearity: a large gap between the two suggests curvature that a linear model will not capture.

| Function | Signature | Returns |
|---|---|---|
| `Pearson_R` | `(X_s, Y, [Include])` | k×1 vector of Pearson correlations between each predictor and Y |
| `Spearman_R` | `(X_s, Y, [Include])` | k×1 vector of Spearman rank correlations between each predictor and Y |
| `Skewness` | `(X_s, [Include])` | k×1 vector of skewness for each predictor column |
| `Kurtosis` | `(X_s, [Include])` | k×1 vector of excess kurtosis (Fisher convention) for each predictor column |
| `Correlation_Matrix` | `(X_s, [Include])` | k×k symmetric matrix of pairwise Pearson correlations among predictors |

### Utility functions

| Function | Returns |
|---|---|
| `Design_Matrix(X_s, [Allow_Intercept], [Include])` | Filtered numeric design matrix as a spilled array |
| `Gram_Inverse(X)` | Inverse of the Gram matrix (X'X)⁻¹ — used internally by `Hat_diagonal`, `PRESS`, `LOOCV_prediction`, and `Prediction_Interval` |
| `Complete_Cases_Filter(X_s, [Y])` | Boolean column — TRUE for rows with no missing values |
| `Col_Select(table, col_nums)` | Selected columns from an array in the specified order |
| `LOO_prediction(X_s, Y, n, [Allow_Intercept], [Include])` | Leave-one-out prediction for a single observation n |
| `This_row(array)` | 1-to-n relative row indices |
| `Exclude_row_n(array, n)` | Array with row n removed |

### Grid-search helpers and Univariate Weibull fitting

`Grid_Argmin(grid)` returns a horizontal three-value array:

```text
minimum value | 1-based row location | 1-based column location
```

Tied minima resolve to the first occurrence in row-major order. `Grid_Search_Optimum(grid)` uses those locations to return a vertical two-value array:

```text
best column-parameter value
best row-parameter value
```

For `Grid_Search_Optimum`, `grid` must be the complete rectangular Data Table body. Column-parameter values must sit immediately above it, and row-parameter values immediately to its left.

The **Univariate** sheet applies these helpers to a two-stage Weibull search. Shape is the column parameter and Scale is the row parameter. Each stage evaluates 20 values per axis—a 20×20 grid and 400 likelihood evaluations. The visible **Rows/Columns** value documents this generated physical size; manually changing it does not resize the Data Table.

Each axis includes both endpoints. Its spacing is:

```text
Step Size = (Max - Min) / (Rows/Columns - 1)
```

Stage 2 centers its narrower range on the Stage 1 optimum and extends one Stage 1 step in each direction. A red Best cell warns when the optimum lands on the corresponding grid boundary.

## Sample data and Regression sheet

`Lambda_Library.xlsx` includes the WHO Life Expectancy dataset (2,938 rows across 193 countries, 2000–2015) as a structured table on the **Life Expectancy Data** sheet. The **Regression** sheet uses this dataset to demonstrate a full multiple regression analysis.

The sheet is organized in four zones:

- **Cols A–J — Main analysis.** Prediction inputs, regression statistics, ANOVA (including `F_Stat` and `P_Value_F`), prediction interval for a new observation, and coefficient table.
- **Predictor summary (below the coefficient table).** A per-predictor panel showing `Pearson_R`, `Spearman_R`, `Skewness`, `Kurtosis`, `VIF`, and `Tolerance` for every predictor — all filtered and recomputed whenever `Full_Data` changes.
- **Cols L–X — Residual output.** One row per filtered observation: predicted values, residuals, LOOCV predictions, leverage, studentized residuals, Cook's distance, and ranked/scored variants.
- **Diagnostic charts (col AC+).** Six XY scatter charts — Actual vs Predicted, Residuals vs Fitted, Normal Q-Q, Predicted vs LOOCV, Cook's Distance, and Residuals vs Top Predictor — that update automatically.

`Correlation_Matrix(X_s, [Include])` returns a k×k pairwise correlation matrix and can be entered in any blank cell on the sheet.

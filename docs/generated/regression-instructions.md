<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# Regression Instructions

The same content as the workbook's **Regression Instructions** sheet,
rendered from the authored rows (`_ROWS` in
`lambda_catalog/write_sheet_regression_instructions.py`).

## How to use the Regression sheet for Multiple Linear Regression (MLR)

To analyze your own dataset, copy the Regression sheet into your workbook (right-click the tab → Move or Copy). This carries over all workbook-scoped LAMBDA functions as well as the sheet-scoped names the Regression sheet uses.

Ensure your data is organized as a structured Excel table. Select the range including column headers and press Ctrl+T, or go to Home → Format as Table.

## Connect the sheet to your data:

Open the Name Manager (Formulas → Name Manager) and update Source_Table (the "Refers To" field) to your table, including its header row — for example =MyTable[#All]. Everything else derives from this one name: the header row, the data body, and the variable list in the MODEL SPECIFICATION block all update automatically. Source_Table points at LifeExpectancyData by default (a curated four-driver Life Expectancy model — Adult Mortality, Alcohol, percentage expenditure, and Status); a second sample table, MileageData (on the Mileage Data sheet), ships with the workbook for a multi-level categorical-encoding demo (MPG ~ Horsepower + Weight + C(Model Year) + C(Origin)). You can use a supplied table or your own Excel Table. After changing the source, review every Role, Include setting, and Type; existing inputs are not inferred from the new headers.

## Define your model in the MODEL SPECIFICATION block (columns A–O):

The Variable column fills itself from your table's headers — one specification row per column. For each row, choose a Role:

- Response (y) — the dependent variable; declare exactly one.
- Predictor (x) — a candidate model input; the Include toggle turns it on or off for the current model run.
- Identifier (Row Label) — labeling columns such as names, IDs, or dates; they appear as row labels in the RESIDUAL OUTPUT section (multiple Identifier columns are joined with "|").
- Filter — a TRUE/FALSE or 1/0 column; only rows where every Filter column is truthy enter the regression. Declare several Filter columns to stratify — they are ANDed together.
- Omit — never used for anything; helper columns and notes.
- Fixed Effects — declare exactly one panel-grouping variable; the model absorbs a separate intercept per group (one-way within transformation) instead of pooled OLS, crediting the absorbed degrees of freedom back into inference. Type and Reference Level don't apply to this row — leave them as-is.

### Predictor types and reference levels

For each included Predictor, set its Type. Continuous predictors enter the design matrix as-is. Categorical predictors are dummy-coded automatically: each level except the reference level becomes a 0/1 column with a level-qualified name (e.g. "Status: Developing"). The reference level defaults to the first level in sort order; type a level into Reference Level to override it (the cell turns red if that level does not exist in the analysis sample). The Levels and Reference In Use columns display what the model will do: the distinct level count in the analysis sample, and the reference actually in effect. A categorical with one level (or an invalid reference) is flagged red and contributes no columns — the rest of the model still computes.

### Review the specification after changing data

Review Role for every source column, and Include and Type for each Predictor; fill these in for any additional specification rows — the dropdowns are available all the way down. A row with a blank Role is ignored. The Order column is a placeholder for a future version and is not read by any formula; it is hidden on the sheet.

### Transforms and original-unit outputs

The Transform column offers Log and Log (drop ≤ 0) on the Response row and on Continuous Predictor rows: the model fits in natural-log space for that variable, and affected outputs are labeled "(Log)" — but the Unit-Space Fit block at AG4:AH10 reports the back-transformed R² / Adj R² / RMSE in original units (Duan smearing by default, Naive on the AH5 toggle), and the Prediction Outputs block's Original Units column (AL) carries the back-transformed point estimate and the four CI/PI bounds. The two new Residual Output columns (AZ, BA) carry Predicted Y and Residual in original units. With a Log response, interval bounds always use Naive back-transformation, regardless of the toggle. They are not confidence bounds for the Duan-adjusted mean. Log is not valid on a Categorical Predictor; setting it there flags the cell red and the fit still runs with that row's Transform ignored (dummy-coded as usual). Plain Log leaves nonpositive values in the sample and returns #N/A; Log (drop ≤ 0) excludes them and reports the count.

### Sequence and period

The Sequence column marks at most one variable as the ordering axis used by lag/difference functions and serial-correlation diagnostics — it is independent of Role and Type (a Predictor can also be the sequence axis). Leave it blank for non-panel data; marking two or more rows shows a red error in the status cell above the column. The Sequence Period column (I) is the typed override input: type a number on the flagged row to declare a Δ that differs from the candidate. The Period In Use column (J) is the live companion — it shows the typed override if column I is non-blank, otherwise the computed candidate (the most common gap between consecutive periods within a group). Lag_By and Difference_By fall back to Base_Period_Delta() when their [delta] argument is omitted.

### Interactions and quadratic terms

The Interaction Term (M) and Interaction Operation (N) columns declare an interaction between this row and another Predictor — pick the other variable and one of Product, Difference, or Ratio. The interaction column is built into the design matrix automatically: Continuous × Continuous adds 1 column, Continuous × Categorical adds one per retained level, and Categorical × Categorical adds the full product of the two. Naming this row's own variable with Operation = Product is the documented way to declare a quadratic (x²) term. An interaction whose mate is switched off (Include = FALSE) still builds and is flagged amber as a marginality warning; a reciprocal declaration — two rows naming each other under Product or Difference — is flagged red, because it would duplicate or negate a column and make the matrix singular. Interaction rows also appear in PREDICTION INPUTS; leave them at the Training Mean defaults or type scenario values.

### Check the model width

The Design Columns column (O) shows how many design-matrix columns each row contributes — 1 for a Continuous predictor, one less than the Levels count for a Categorical one. The total above it, with the intercept added, is the full width of the constructed design matrix; it turns amber on a model wide enough to be slow and red on one too wide for the sheet.

## Intercept:

The Intercept toggle sits at the top of the Include column (C2). Leave it TRUE for models with Categorical predictors: reference-level dummy coding relies on the intercept to carry the baseline, and the toggle flags red if it is FALSE while a Categorical predictor is included.

## Which rows are used:

The analysis sample is derived automatically — a row is included when the Response is numeric, every included Continuous predictor is numeric, and every Filter column is truthy. Include settings and Log (drop ≤ 0) also affect the sample. Categorical predictors impose no completeness requirement: rows with a blank category still enter the sample (all of that variable's dummies are blank there) unless a Filter excludes them.

## Optional — point prediction:

Enter a value for each design-matrix column in the orange cells under PREDICTION INPUTS (column AK) — one row per constructed column, including one per dummy (use 1 for the scenario's level, 0 for its siblings; no Intercept row — the model's own baseline is handled automatically). The Training Mean column beside the inputs shows each column's mean over the analysis sample, which is also the prefilled default — so the untouched prediction is at the data's center. Results appear in the PREDICTION OUTPUTS box above as both a mean-response confidence interval (CI) and a wider new-observation prediction interval (PI); the confidence level is controlled by the Alpha cell (AB13) in REGRESSION OUTPUTS. If a Fixed Effects variable is declared, the FE Group cell above the inputs selects which group's own average the prediction is anchored to.

## Optional – faster charts on a fixed-size dataset:

The seven diagnostic charts read each series from a worksheet-scoped named range (the RegChart* names) built with OFFSET and sized to the live observation count at $AB$9. OFFSET is volatile, so those ranges re-evaluate on every recalculation pass. If your dataset size is stable, you can remove that overhead by pointing each chart series directly at the absolute range the name resolves to — from row 4 (one below the column header) down to row 3+N, where N is the count shown at $AB$9 (for 200 observations, for example ='Regression'!$AP$4:$AP$203). Select the chart, click a series, and replace the ='Regression'!RegChart... reference in the formula bar with the absolute range. The trade-off is that the chart no longer resizes if you add or drop rows — you must re-point it when the row count changes. The named ranges can stay defined; there is no need to delete them.

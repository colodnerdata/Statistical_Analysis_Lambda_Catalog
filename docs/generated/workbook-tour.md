<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# The workbook, tab by tab

The shipped `dist/Lambda_Library.xlsx` presents twelve tabs in this
order (extracted from `scripts/build_production.py` at generation
time):

1. Regression
2. Regression Instructions
3. Modeling Concepts
4. Diagnostic Guide
5. Model Comparison
6. Model Comparison Guide
7. Univariate
8. LAMBDA_functions
9. Version History
10. Production Lots
11. Life Expectancy Data
12. Mileage Data

- **Regression** — the working sheet: MODEL SPECIFICATION (A–O),
  Regression Outputs, Prediction Outputs, Residual Output, and the
  seven diagnostic charts.
- **Model Comparison** — one row per fitted model, read from the other
  model sheets: a goodness-of-fit table and a prediction comparison,
  gated per row so two models are only compared on statistics that
  actually bear comparing.
- **Regression Instructions / Modeling Concepts / Diagnostic Guide /
  Model Comparison Guide** — the built-in manual (each has a generated
  page in this site).
- **Univariate** — descriptive statistics, histograms, distribution
  fitting, and Q-Q plots for one column of data.
- **LAMBDA_functions** — the catalog itself, one row per function.
- **Version History** — what shipped when.
- **Production Lots / Life Expectancy Data / Mileage Data** — three
  Excel Tables you can practice retargeting `Source_Table` against.

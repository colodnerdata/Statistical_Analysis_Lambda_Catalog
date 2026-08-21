# Statistical Analysis Lambda Catalog

Excel 365 LAMBDA functions that replicate and extend Analysis ToolPak regression statistics — no VBA, no add-ins, no installation. Download a workbook, open it in Excel, and all functions are immediately available in formulas.

## Getting started

1. Download [`Lambda_Library.xlsx`](dist/Lambda_Library.xlsx) from the repository's `dist/` directory.
2. Open it in Excel 365 (Windows or Mac).
3. Save a working copy before replacing sample data or changing the templates.

The file in `dist/` is the ready-to-use workbook; you do not need Python or the build scripts unless you want to contribute to the project.

### Choose a starting point

| If you want to… | Start here |
|---|---|
| Explore a working regression model | Open **Regression**. It is preconfigured against one of the included sample datasets, so its outputs and diagnostic charts are populated immediately. |
| Run a regression on your own table | Read **Regression Instructions**, then follow its table-setup and `Source_Table` retargeting steps. Merely pasting data onto an arbitrary sheet does not retarget the Regression template. |
| Describe or fit a distribution | Open **Univariate** and use its input controls for descriptive statistics, histograms, fitted distributions, and Q-Q plots. |
| Call a function directly | Open **LAMBDA_functions** to find the function signature, then enter that function by name in a cell in this workbook. |
| Understand a regression warning or chart | Open **Diagnostic Guide** for thresholds, plot interpretations, and suggested next steps. |

The workbook has two layers: the **library** — all 155 LAMBDA definitions (131 workbook-scoped and shared by every sheet, plus 24 sheet-scoped template constructors) — and the **templates**, the pre-built analysis sheets that demonstrate how to drive the library. The **Regression** template installs a small set of sheet-scoped constructor names for the spec block and provides a ready-to-use analysis interface: declare each table column's Role and Type on the spec block, and the sheet derives the row mask, the constructed design matrix, and the full regression output automatically. The **Univariate** template demonstrates descriptive statistics, histogram binning, and distribution fitting via search-based MLE. The **LAMBDA_functions** sheet is the canonical, always-in-sync function catalog — see [Documentation map](#documentation-map) below. Data sheets, instructions, and the diagnostic guide are reference sheets that support the templates.

Most functions are defined as workbook-scoped names, so they work in any cell formula within the workbook. Around 30 catalog functions are called by no pre-built template. That is deliberate: they are the **standalone user-callable layer** — `Correlation_Matrix`, `Lag_By`, `Descriptive_Statistics`, `Design_Matrix` and others you call in your own cells on your own data. The templates demonstrate the library; they are not the whole of it.

### Use the library in another workbook

To use these functions in a different workbook, you have two options. **The easy one:** if you are using one of the pre-built templates (the Regression workbench or the Univariate Analysis sheet), copy the sheet into your own workbook. The sheet's named-range dependencies come with it — the **workbook-scoped LAMBDA definitions** (the 131 portable functions) and the **sheet-scoped definitions** (the Regression sheet's 24 Regression-scoped definitions — constructors like `Predictor_Columns`, `Sample_Include`, `Design_Columns`, plus the row-2 status readouts like `Role_Status`; the Univariate sheet's `UV_Data`, `UV_Include`, `GoF_AIC`) travel inside the sheet-copy and are renamed automatically. Open your workbook, the sheet calculates, and every function is ready in formulas. **The other one:** if you only want a single function, or you would rather not pull in the pre-built template at all, open both files in Excel at the same time and reference functions as `='[Lambda_Library.xlsx]'!FunctionName(args)`, or use Name Manager (Formulas → Name Manager → New) to copy individual definitions into your own workbook.

## Versions

One number — `MAJOR.MINOR.PATCH` — because there is one workbook:

- **Function library version** — the shared catalog of 155 LAMBDA definitions. Moves when a function is added, renamed, or changes what it returns.
- **Workbook version** — the workbook's sheets, input cells, and control blocks. Moves when its input surface changes.

The workbook's **Version History** sheet shows both:

```
Lambda Library 3.3.0   ·   Function Library 3.3.0
```

The **`Breaking?` flag belongs to the workbook version**, since it answers a question about your saved inputs. A library-version bump that adds a function breaks nothing.

## What ships in the workbook

### `Lambda_Library.xlsx`

One workbook, built by one script (`build_production.py`). It carries the full function library — all 155 LAMBDA definitions (131 workbook-scoped plus 24 sheet-scoped template constructors) — plus the pre-built templates and the reference and data sheets. Nine sheets, in tab order: Regression, Regression Instructions, Diagnostic Guide, Univariate, LAMBDA_functions, Version History, Production Lots, Life Expectancy Data, Mileage Data.

Includes the WHO Life Expectancy dataset (2,938 rows across 193 countries, 2000–2015) as a structured table on the **Life Expectancy Data** sheet.

The **library** and the **templates**:

- **LAMBDA_functions** — the **canonical function catalog**: every function's name, scope, full LAMBDA definition, arguments, yields, plain-language summary, and long description. This is the source of truth for "what functions exist and what they do" — generated from `lambda_functions.json` by `lambda_catalog/write_sheet_lambda_functions.py`, so the sheet and the source are never out of sync. This is the library.
- **Regression** — the spec-driven regression workbench, and a pre-built template. Each table column gets a Role (`Response (y)` / `Predictor (x)` / `Identifier (Row Label)` / `Filter` / `Omit` / `Fixed Effects`). Each predictor also gets a Type (Continuous / Categorical, with reference-level control), an optional Transform (natural log, on the Response row and/or Continuous Predictor rows), and an optional interaction with another predictor (Product, Difference, or Ratio — a row pointing at itself under Product is how you write a quadratic term). Any column can carry the structural Sequence flag for lag/difference/serial-correlation features. The spec block derives the row mask, the constructed design matrix, and level-qualified column names; the sheet then produces the full regression output, diagnostics, and — for a Fixed Effects model — both a mean-response CI and a new-observation PI in the prediction outputs.

  The template is laid out as **five content zones** — Model Specification, Predictor Summary, Regression Outputs, Prediction Outputs, and Residual Output — each in its own collapsible outline group, separated by narrow gap columns, with seven pre-built diagnostic charts to the right. Past the charts sits a band of **materialized** blocks — the Model Context, the row mask, and the Constructed Design Matrix — which exist so the engines read a computed value instead of recomputing it at every call site. Only the Model Context block collapses; the row mask and the design matrix stay expanded, because hiding the columns a spilled array occupies is what leaves it stale when the model recalculates.
- **Univariate** — a pre-built template for characterizing a single variable's distribution. Descriptive statistics, three side-by-side histogram binning methods (Sturges, Scott, Freedman-Diaconis), and two-stage distribution fitting across eight candidate distributions. Weibull and Gamma profile their scale/rate parameter out in closed form and search a 20-point profile-NLL curve per stage; Beta searches both parameters on a `Full_Factorial` dynamic-array grid (N² negative-log-likelihood evaluations per stage, N editable live in the sheet); Normal, Log-Normal, Exponential, Triangular, and BetaPERT are closed-form. The fitted-distribution Q-Q plots, the Weibull and Gamma profile-NLL line charts, and the histogram distribution overlays live alongside.
- **Regression Instructions** — step-by-step guide for adapting the Regression template to a new dataset, including Name Manager updates and table setup. A reference sheet.
- **Diagnostic Guide** — interpretation guide for regression diagnostics with Tier 1/Tier 2 plot specifications, threshold reference table, and "Common Patterns & Next Steps" guidance. A reference sheet.
- **Version History** — changelog that travels with the workbook for non-git users. Every release's "Breaking? (yes/no)" flag is here so a workbook user gets the one signal the version number is *for* — "do my existing inputs still work?" — without the number also having to convey "how big is this release."
- **Life Expectancy Data** — the WHO dataset as a structured table. It ships as a ready-made target for practicing the Regression template's `Source_Table` retarget: point `Source_Table` at `LifeExpectancyData[#All]` in Name Manager and the spec block re-populates from its columns, no data of your own required.
- **Mileage Data** — the Auto MPG dataset (406 vehicles) as a structured table. Retarget `Source_Table` at `MileageData[#All]` for a multi-level categorical-encoding example.
- **Production Lots** — a small unbalanced learning-curve panel (3 facilities, 51 lots) as a structured table. The only shipped dataset with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year) — retarget `Source_Table` at `ProductionLotsData[#All]` (or pass `--regression-dataset production_lots` to `build_production.py`) for a ready-made Fixed Effects example.

The workbook ships in full **Automatic** calculation mode — no Data Tables anywhere — so fitted parameters update live as you change the input column or the Beta grid-points cell.

## Documentation map

This README is intentionally short — it tells you what the library is and how to open it. The rest of the documentation is split by purpose:

| File | What's in it | When to read it |
|---|---|---|
| [CONTRIBUTING.md](CONTRIBUTING.md) | Dev guide: setup, tests, build, file structure, adding a new LAMBDA function, Regression sheet chart conventions | You're contributing to the codebase |
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | Foundational patterns: naming convention, function categories, the Role / Type / Sequence taxonomy, the Model Spec block (A–O), the data-transformation taxonomy, the reserved-spec-column pattern | You're adding a feature that has to honor the library's design |
| [ROADMAP.md](docs/ROADMAP.md) | Version plan: the public-interface definition, the version ladder, what's shipped, what's next, ToolPak parity reference | You want to know what's planned or whether a change breaks the public interface |
| [DECISIONS.md](docs/DECISIONS.md) | Resolved design decisions with their rationale, indexed by version, plus the supersession log | You want to know *why* a design choice was made |
| [MODEL_TESTING_ASSETS.md](docs/MODEL_TESTING_ASSETS.md) | The regression test-model suite: which model configurations the QC harness covers, which corner each one is there for, the datasets future milestones need, and the ordering the version ladder follows from v3.4 on (Regression work first, then test-scale growth) | You're adding or changing a QC model case, wiring a new dataset, or wondering why the roadmap is sequenced the way it is |
| [TODOs.md](docs/TODOs.md) | Open work only — every item tagged status · size · whether it needs Excel, with a "Pick something to work on" index at the top and cross-links to DECISIONS for the "why" | You want to pick up the next piece of work |

For the canonical function reference (every function's signature and description), open the **LAMBDA_functions** sheet in `Lambda_Library.xlsx` — the library catalog, generated from the same source as the codebase and the only place the reference is always in sync with the workbook.

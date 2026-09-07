# The Lambda Catalog — a guided tour for Excel users

The **Statistical Analysis Lambda Catalog** is an Excel workbook for
regression and univariate analysis. You can connect the supplied template
sheets to your chosen data sources, specify a model, inspect its results,
and make predictions. You can also use the named LAMBDA functions in your
own formulas.

The workbook, `dist/Lambda_Library.xlsx`, uses Excel formulas and requires
no macros or add-ins.

It was written for someone comfortable with spreadsheets but new to both
statistics workbenches and this repository. No programming knowledge is
assumed; where a formula appears, it comes with a plain-English
explanation.

The reference pages are generated from the same authored text used to
build the workbook's reference sheets, together with the catalog and
selected formulas from the sheet writers.

## Get the workbook

[Download Lambda_Library.xlsx](https://github.com/colodnerdata/Statistical_Analysis_Lambda_Catalog/raw/refs/heads/main/dist/Lambda_Library.xlsx)
and save a working copy before changing its data or specifications.

Use an up-to-date desktop Excel for Microsoft 365. The workbook depends on
LAMBDA and dynamic-array functions including TAKE, MAP, and VSTACK.
Microsoft lists Microsoft 365 and Excel 2024 for Windows and Mac as
supporting [LAMBDA](https://support.microsoft.com/en-us/excel/functions/lambda-function)
and [TAKE](https://support.microsoft.com/en-us/excel/functions/take-function).
That function availability does not establish that this entire workbook
has been tested on every edition; the project's automated workbook
verification runs in desktop Excel on Windows.

Leave **Formulas → Calculation Options → Automatic** selected so results
update after you edit the inputs.

## Using the templates

You can connect the template sheets to your chosen data sources using
{doc}`walkthrough`. For an example with values you can check, follow
{doc}`worked-example`.

```{toctree}
:maxdepth: 1
:caption: Using the workbook

worked-example
walkthrough
generated/workbook-tour
```

```{toctree}
:maxdepth: 1
:caption: Model choices and reference

generated/modeling-concepts
generated/diagnostic-guide
generated/spec-block
generated/regression-instructions
```

```{toctree}
:maxdepth: 1
:caption: Technical reference

technology
generated/formula-review
generated/lambda-reference
```

```{toctree}
:maxdepth: 1
:caption: For contributors

improvements
```

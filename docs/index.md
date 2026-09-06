# The Lambda Catalog — a guided tour for Excel users

Welcome. This site is a tutorial for the **Statistical Analysis Lambda
Catalog** — a project that ships one file, `dist/Lambda_Library.xlsx`, which
turns a copy of Excel you already own into a working statistics workbench:
multiple regression with categorical predictors, interactions, fixed effects
and log transforms; univariate analysis with distribution fitting; and a
full diagnostic suite — with **no macros, no add-ins, and nothing to
install**.

It was written for someone comfortable with spreadsheets but new to both
statistics workbenches and this repository. No programming knowledge is
assumed; where a formula appears, it comes with a plain-English
explanation.

The pages under *The built-in manual* and *Reference* are **generated from
the workbook's own source lists** — the same Python that writes the sheets
writes these pages — so what you read here is what the sheets say, and the
two cannot drift apart.

## How to read this site

Start at the beginning and read in order; every term is explained before
it is used.

```{toctree}
:maxdepth: 2
:numbered: 1

technology
generated/workbook-tour
walkthrough
generated/spec-block
generated/regression-instructions
generated/modeling-concepts
generated/diagnostic-guide
generated/formula-review
generated/lambda-reference
improvements
```
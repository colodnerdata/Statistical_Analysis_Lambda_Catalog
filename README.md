# Statistical_Analysis_Lambda_Catalog

Reproduce Analysis ToolPak-style statistics with workbook-scoped Excel LAMBDA functions.

## Build the workbook

Run `build_lambda_library.py` to open or create `Lambda_Library.xlsx` and sync the workbook Name Manager entries from `lambda_functions.json`.

The script:
- creates `Lambda_Library.xlsx` if it does not exist
- opens the existing workbook if it already exists
- ensures a starter worksheet named `MLR` is present
- converts each JSON `formula_display` value into workbook XML syntax before writing workbook-scoped names
- overwrites any existing workbook-scoped names that match the JSON definitions

This script requires desktop Excel because it uses `xlwings`.

```powershell
python build_lambda_library.py
```

Use `--validate-reopen` to confirm Excel can reopen the saved workbook after the XML name patch.

```powershell
python build_lambda_library.py --validate-reopen
```

The current build path uses `formula_display` as the source of truth for workbook name syncing. The parser strips formatting whitespace outside string literals so the multi-line display form produces the same workbook XML as the older compact form.

## Write the catalog sheet

Run `write_lambda_catalog.py` to write the human-readable catalog into a worksheet named `LAMBDA_functions` in any target workbook.

The script:
- opens the target workbook if it already exists, or creates a new workbook if it does not
- clears and rewrites the `LAMBDA_functions` worksheet starting at cell `A1`
- writes a structured Excel table named `LAMBDAFunctionsCatalog`
- loads the displayed formulas, arguments, yields, and descriptions from `lambda_functions.json`

```powershell
python write_lambda_catalog.py Lambda_Library.xlsx
```

## Write the life expectancy sheet

Run `write_life_expectancy_data.py` to import `Life Expectancy Data.csv` into a worksheet named `Life Expectancy Data` in any target workbook.

The script:
- opens the target workbook if it already exists, or creates a new workbook if it does not
- clears and rewrites the `Life Expectancy Data` worksheet starting at cell `A1`
- writes a structured Excel table named `LifeExpectancyData`
- adds a calculated column named `Full_Data` with the formula `=COUNT(LifeExpectancyData[@[Life expectancy]:[Schooling]])=19`

```powershell
python write_life_expectancy_data.py Lambda_Library.xlsx
```

After the workbook exists, `format_workbook.py` can be used as a separate formatting step.

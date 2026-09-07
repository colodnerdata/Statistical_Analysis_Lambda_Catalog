# Proposed improvements

These proposals concern documentation, data selection, and presentation.
They are suggestions; `docs/ROADMAP.md` in the repository remains the
plan of record.

## 1. Add Univariate instructions

The Regression sheet has three reference sheets: **Regression
Instructions**, **Modeling Concepts**, and **Diagnostic Guide**.
Univariate has hover Notes on its controls but no corresponding
instructions sheet.

Add instructions for selecting a data column, interpreting the bin-width
rules, and using the two-stage grid searches and Grid Points controls.
The content could follow the existing static-sheet pattern and supply a
generated page on this site.

## 2. Add a source name for Univariate

Regression uses `Source_Table` to select its data. Univariate currently
references the source column directly in two cell formulas:

- `=IF(LifeExpectancyData[Life expectancy]="","",…)`
- `=ISNUMBER(LifeExpectancyData[Life expectancy])`

Introduce a sheet-scoped source name that both formulas read. Users could
then select a different column by editing that name in the Name Manager.

## 3. Add a Beta fit chart

The Beta fit zone reserves `BY13:CD30` for a chart. Weibull and Gamma
already have charts showing their two-stage profile-NLL searches.

A heatmap or surface plot could show Beta NLL over the alpha/beta grid.
This needs a chart design suited to two searched parameters; the existing
profile charts show one searched parameter.

## 4. Check cell addresses in reference text

The static reference sheets cite cell addresses in prose. Those citations
can become outdated when the Regression layout changes, even when the
underlying formulas use layout constants.

Derive address citations from `regression_layout.py` where possible, and
check the remaining citations against their intended controls or outputs.
Use labels such as “the Alpha input” where an address is unnecessary.
Regenerate the static sheets and documentation after changing the text.

## 5. Extend formatting when grids grow

Increasing **Grid Points** grows a fit's grid and NLL column. Formatting
is applied to a fixed window, so additional rows can appear unshaded.

Evaluate extending conditional formatting and number formats over a
larger range, with a check on workbook size and recalculation cost.
Alternatively, add a note to the Grid Points control explaining that
rows beyond the formatted window retain their calculated values but may
appear unshaded.

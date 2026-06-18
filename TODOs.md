Chart `SERIES` formulas require explicit range references; the `#` spill operator is not reliably supported in chart series formulas. However, referencing all 1,048,576 rows can significantly degrade performance or crash Excel when the populated dataset is much smaller.

Instead, define dynamically sized named ranges using the row count in `$M$8`. For example:

```excel
QQPlotX = OFFSET($AJ$2,1,0,$M$8,1)
QQPlotY = OFFSET($AK$2,1,0,$M$8,1)
```

These formulas define ranges beginning at `AJ3` and `AK3`, respectively, and extending for exactly the number of rows specified in `M8`.

Create `QQPlotX` and `QQPlotY` as **worksheet-scoped names on the Regression sheet**, not as workbook-scoped names. When assigning them to the scatter chart series, qualify each name with the worksheet name:

```excel
Series X values: ='Regression'!QQPlotX
Series Y values: ='Regression'!QQPlotY
```

Do not reference the names without the sheet prefix. Charts live "above" the worksheet. Because these are worksheet-scoped names, the chart series formulas must explicitly identify the worksheet that owns them.

~~TODO: Fix the references for every chart on the Regression sheet to follow this pattern.~~ ✅ Done
  - ~~Normal Q-Q: `QQPlotX` (col AE) and `QQPlotY` (col AF)~~ ✅ Done
  - ~~Residuals vs. Fitted: `FittedY` (col Y, shared), `ResidData` (col Z)~~ ✅ Done
  - ~~Actual vs. Predicted: `FittedY` (col Y, shared), `ActualY` (col X)~~ ✅ Done
  - ~~Scale-Location: `FittedY` (col Y, shared), `ScaleLocData` (col AG)~~ ✅ Done
  - ~~Cook's Distance: `CooksDistData` (col AD)~~ ✅ Done
  - ~~Leverage vs. Studentized: `LeverageData` (col AB), `StudResidData` (col AC)~~ ✅ Done
  - ~~PRESS Residuals: `PRESSResidData` (col AH)~~ ✅ Done
~~TODO: Decrease the marker size for the scatter charts to 4.~~ ✅ Done
~~TODO: Change chart titles to Header format and font size 14.~~ ✅ Done
~~TODO: Add a Y=X Reference line to the QQ and Actual v Predicted Plots - do not add helper columns.~~ ✅ Done
~~TODO: Change axis number formatting to 0 DP for all charts.~~ ✅ Done
~~TODO: Change PRESS Residual and Cook's Distance to column charts rather than using Observation as an X variable - update this in the Diagnostic guide as well.~~ ✅ Done
TODO: Add reference lines to the Cook's Distance, PRESS Residuals, Leverage vs. Studentized. Format the reference lines thematically similar to the conditional formatting in the table. Use minimalist helper columns (2 points using max/min; they can be underneath the charts.)

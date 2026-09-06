# Part 3 — Five improvements, from a self-review

A review of this project should end by turning the same scrutiny on
itself. These five suggestions come from reading the workbook's own
source; each is grounded in something the code and docs already say, and
each is scoped to be one focused change. They are *suggestions* — the
owner's roadmap (`docs/ROADMAP.md` in the repository) remains the plan of
record.

## 1. Give the Univariate sheet an instructions sheet

The Regression sheet comes with a built-in manual: **Regression
Instructions** (the how), **Modeling Concepts** (the why), and the
**Diagnostic Guide** (the what-to-look-for). The **Univariate** sheet has
nothing comparable — its control fields carry hover Notes, and the
{doc}`generated/regression-instructions`-style orientation ("point the
sheet at your data with one edit…") simply does not exist for it. A new
user meeting the two-stage grid search, the live Grid Points cells, or
the three bin-width rules for the first time has to reverse-engineer
them from the sheet itself.

**Suggested change:** a fourth static reference sheet
("Univariate Instructions") following the established pattern — authored
in a `write_sheet_univariate_instructions.py` module, baked into
`templates/static_sheets.xlsx` by `scripts/rebuild_static_sheets.py`,
and linked from this docs site by the same generator.

## 2. Give Univariate the one-edit retarget the Regression sheet promises

The Regression sheet's central promise is one edit: retarget
`Source_Table` in the Name Manager and everything resizes. The
Univariate sheet **cannot make that promise** — its data references are
hard-coded into cell formulas:

- the sample column is `=IF(LifeExpectancyData[Life expectancy]="","",…)`
- the numeric test beside it is `=ISNUMBER(LifeExpectancyData[Life expectancy])`

Retargeting the Univariate sheet means editing those two formulas (and
knowing to), not one name. The fix follows the Regression sheet's own
pattern: a `UV_Source_Table` name that both formulas read, so pointing
the sheet at another column becomes a single Name Manager edit.

## 3. Build the reserved Beta chart

The Beta fit zone on the Univariate sheet reserves chart space —
`BY13:CD30`, between the control block and the grid body — but the zone
is blank: the Weibull and Gamma fits each get a two-stage profile-NLL
chart there, while Beta, the most parameter-rich fit, gets none. The
reservation is deliberate (the zone is documented as "reserved for a
future Beta chart"), but a user comparing three distributions sees two
charts and an empty band.

**Suggested change:** a surface plot or profile-heatmap of the Beta
NLL over the (α, β) grid — the grid is already materialized as a
spill, and the two profile fits' charts (`regression_charts.py`
patterns, OFFSET-based named ranges, masked overlay series) are the
template. The main design question is honest: a 2-D surface of N² cells
is a different chart type from the 1-D profile curves, so this is the
one suggestion that needs design, not just wiring.

## 4. Add a systemic check: static sheets must not cite stale addresses

The static reference sheets cite cell addresses in prose ("Alpha is at
AB12", "Cook's Distance in Col AT"), and those addresses drift when the
Regression layout shifts — this review found one instance shipped
(the Instructions sheet cited AB12 where the live readout sits at AB13),
and the Diagnostic Guide's own comments record a whole-zone drift found
and fixed late for the same reason. The root cause is structural: the
sheets are baked into `templates/static_sheets.xlsx` and no build
re-derives them, so a layout change cannot invalidate their text.

**Suggested change:** a drift check that extracts every
address-shaped citation from the three static-sheet modules and
validates it against `regression_layout.py` — the module whose
constants define the real layout. It could run as a unit test (pure
text/AST, no Excel), turning "a static sheet cites a moved cell" from
a find-it-late bug into a red CI check. A milder variant: make the
prose cite *labels* ("the Alpha readout", "the Cook's Distance
column") instead of addresses, so there is nothing to drift.

## 5. Make the live-grown grid rows inherit their formatting

Raising a fit's **Grid Points** (N) grows its grid — and the rows the
increase adds beyond the statically formatted window come in
**unshaded**: number formats, the color scale and the border box are
painted over the default-size window only, so a live increase produces
a body that is half-formatted. The workbook's own documentation calls
this cosmetic and deliberate, and it is — nothing computes wrong — but
it is also the one place where the workbook's "everything derives from
one edit" story shows a seam: the *data* follows N, the *look* does
not.

**Suggested change:** the same conditional-formatting trick the input
band already uses — extend the body's CF rules to the full
16000-row ceiling (they are `Formula2`-sized already) so rows added by
a live N render like their neighbors. If the CF-rule count makes that
heavy, a cheaper fix is a one-line note in the fit's own hover text
saying the added rows are expected to appear unshaded, so the user is
never surprised.

---

None of these are blockers — the workbook ships correct, documented,
and quietly ingenious. They are the difference between a tool whose
two halves (Regression and Univariate) feel like one coherent workbook,
and one
where the second half visibly grew up later.
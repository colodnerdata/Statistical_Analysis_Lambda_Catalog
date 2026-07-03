# Model Construction Sheet — Human Test Plan (WHO dataset)

Execute in order — each test is a small delta on the previous spec state, simplest
to most complex. Format: the spec values you would enter, then the expected
observations. Rows not listed are unchanged from the previous test. Reference
numbers below were computed directly from `Life_Expectancy_Data.csv` and the
`Full_Data` completeness definition, so mismatches indicate implementation bugs,
not stale expectations.

**Baseline facts:** 2,938 data rows; `Full_Data` TRUE on 1,649; within those 1,649:
Status has 2 levels (`Developed`, `Developing` — 242 / 1,407 rows), Year has 16
levels (2000–2015, first-sort reference `2000`), Country has 133 levels.

---

## T0 — Initial state sanity (build defaults)

No edits. The build ships this spec:

| Variable | Role | Include | Type | Reference |
|---|---|---|---|---|
| Country | Identifier | — | — | |
| Year | Predictor | TRUE | Categorical | |
| Status | Predictor | TRUE | Categorical | |
| Life expectancy | Response | — | — | |
| Adult Mortality | Predictor | TRUE | Continuous | |
| GDP | Predictor | TRUE | Continuous | |
| Schooling | Predictor | TRUE | Continuous | |
| (all other numerics) | Predictor | FALSE | Continuous | |
| Full_Data | Filter | — | — | |

**Expect:** k = 19 · rows = 2938 · response = `Life expectancy` · responses = 1 ·
included rows = 1649. Header strip reads, in spec order: `Year: 2001` … `Year: 2015`
(15 columns), `Status: Developing`, `Adult Mortality`, `GDP`, `Schooling`. Levels
column: Year = 16, Status = 2, blank elsewhere. C–F gray on every non-Predictor row.
First filtered label = `Afghanistan|2015`.

## T1 — Reproduce v1: full continuous set, y, filter (simplest behavior)

Set Year → Role `Identifier`; Status → Role `Omit`; then set Include = TRUE on all
18 numeric predictors (Adult Mortality … Schooling).

**Expect:** k = 18 · included rows = 1649 · header strip = the 18 table headers in
table order — this is exactly v1's `x_s`/`y`/`Regression_Sample_Include`. Spot-check
row 1 of the filtered matrix against the Regression sheet's first included
observation (Afghanistan 2015); values must match cell-for-cell, and filtered y
column = Life expectancy values.

## T2 — Continuous subset

Set Include = FALSE on all but Adult Mortality, GDP, Schooling.

**Expect:** k = 3; matrix columns in spec order; included rows still 1649
(completeness is spec-driven: only *included* continuous predictors + response
demand numeric — but every Full_Data row is complete anyway, so the count holds).

## T3 — First categorical (binary)

Status → Role `Predictor`, Include TRUE, Type `Categorical`, Reference blank.

**Expect:** k = 4 · Levels(Status) = 2 · new column named `Status: Developing`
(reference defaulted to `Developed`, first in sort — surfaced by omission in the
name) · dummy column contains only 0/1 · full-height rows still 2938.

## T4 — Reference override

Type `Developing` into Status's Reference cell.

**Expect:** the dummy column flips to `Status: Developed`; k stays 4; the 0/1
pattern inverts. Clear the cell → reverts to `Status: Developing`.

## T5 — Invalid reference (visible failure)

Type `Developped` (typo) into Status's Reference cell.

**Expect:** E cell turns red (invalid-reference CF). Constructor behavior per spec:
the typo'd reference matches no level, so *no* level is dropped — full one-hot
plus intercept = the dummy trap the red flag exists to prevent. Verify the red
appears; then fix the cell.

## T6 — Numeric-valued categorical

Year → Role `Predictor`, Include TRUE, Type `Categorical` (this is T0's state for
Year).

**Expect:** k = 19 · Levels(Year) = 16 · fifteen columns `Year: 2001` … `Year: 2015`
(reference `2000` — numeric sort, not text sort; if you see `Year: 2013` missing and
weird ordering, numeric/text coercion is broken) · each Year dummy is 0/1.

## T7 — Identifier labeling variants

(a) Country = Identifier, Year = Predictor (current state): labels are Country only
→ `Afghanistan`. (b) Set Year → `Identifier` too: labels become `Afghanistan|2015`.
(c) Set Country AND Year → `Omit`: labels fall back to `Obs. 1`, `Obs. 2`, … in
full-height H, and the *filtered* label column starts at the first included row's
observation number (not `Obs. 1` unless row 1 passes the filter — Afghanistan|2015
does, so `Obs. 1` should appear).
Restore: Country = Identifier, Year = Identifier.

## T8 — Filter composition (declarative stratification)

Add a helper column to the LifeExpectancyData table:
`Is_Developing` = `=--([@Status]="Developing")`. Its spec row appears automatically
(A spills from headers). Set its Role → `Filter`. Status: Role `Predictor`,
Include TRUE, Categorical, Reference blank. Year: Identifier.

**Expect:** included rows = **1407** (Full_Data AND Is_Developing) · Levels(Status)
collapses from 2 → **1** — computed over the *masked* sample — and the F cell turns
**red** (included, categorical, L ≤ 1) · Status contributes **zero** columns (k
drops by 1 relative to the same spec unstratified) while every other output still
computes: visible degradation, not a hard error, not silent omission. Filtered label
count = 1407; Country levels within the stratum = 114 if you inspect.

## T9 — Filter semantics edge

Set Is_Developing → `Omit`, leaving only Full_Data as Filter: included rows returns
to 1649. Then set Full_Data → `Omit` too (zero Filter columns).

**Expect:** with zero Filters, the mask is completeness-only: included rows =
count of rows where the response and every included continuous predictor are
numeric — with the T8 predictor set this is ≥ 1649 (Full_Data demanded *all 19*
numerics; the spec-driven mask demands only the ones in the model). The exact
number depends on included predictors; verify it is ≥ 1649 and equals 2938 minus
rows with a blank in y/AdultMortality/GDP/Schooling. Restore Full_Data → `Filter`.

## T10 — Response swap (derived y)

Life expectancy → `Predictor`, Include TRUE, Continuous. GDP → `Response`.

**Expect:** response audit cell reads `GDP`; responses = 1; filtered y column now
holds GDP values; Life expectancy appears in the header strip as a design-matrix
column; included-rows count recomputes (GDP-as-response must be numeric — identical
outcome here since Full_Data already demands it). Restore afterward.

## T11 — Response-count validation

(a) Set GDP → `Predictor` while Life expectancy is still `Predictor` (zero
Responses): responses = 0, **red**; response cell reads `(none)`; filtered-y zone
shows the empty-model/error fallback rather than fabricating a column.
(b) Set BOTH GDP and Life expectancy → `Response`: responses = 2, **red**;
documented behavior is first-match wins (GDP is earlier in table order — verify the
y column is GDP, and that the red count is the only alarm).
Restore: Life expectancy = Response, GDP = Predictor/TRUE/Continuous.

## T12 — Levels display is live and mask-aware (no model change)

Country → `Predictor`, Categorical, Include **FALSE**.

**Expect:** Levels(Country) = **133** — not 193 — because levels are computed over
the 1,649 mask-included rows, and no red flag (not included). k unchanged. Then
Include TRUE briefly: k jumps by 132 and the header strip floods with
`Country: …` names — the visible-count-as-warning behavior. Restore Country →
`Identifier`.

## T13 — Extreme stratification degeneracy

Add helper column `Is_2015` = `=--([@Year]=2015)`, Role → `Filter` (with Full_Data
still Filter).

**Expect:** included rows = **2** (only two 2015 rows survive Full_Data — GDP and
Population are missing for most 2015 records; this is a real dataset property, not
a bug). Year (if Predictor/Categorical/TRUE) shows Levels = 1 → red, contributes 0
columns. Any continuous fit on 2 rows is degenerate downstream but the construction
zone itself must not error. Restore Is_2015 → `Omit`.

## T14 — Empty model

Set every Predictor row's Include → FALSE.

**Expect:** every output zone shows `(empty model)` (from the `IFERROR` wrappers) —
no `#CALC!` leaks to the sheet. Audit k reads `(empty model)`. Restore T0 state.

## T15 — Full-height vs filtered contract

With T0 state: column H (labels) and I (filter booleans) run the full 2,938 rows;
K/L and N/O… run exactly 1,649. `ROWS(X_s())` audit = 2938 always — the constructor
never row-filters (the row-mask contract); only the display zones do.

## T16 — Twin alignment tripwire

In any state from the above: the header strip width must equal the k audit cell,
always, including mid-edit states (degenerate categorical, empty model). If they
ever disagree, the `X_s` / `Constructed_Column_Names` twins have drifted — a
structural bug, not a data issue.

---

## Known caveat to verify and accept (or escalate)

**Blank categorical values pass the filter.** The mask demands numeric y and
numeric included *continuous* predictors only; a Categorical Predictor cell that is
blank imposes no condition, and a blank encodes 0 against every retained level —
i.e., it is silently classified as the *reference level*. The WHO Status/Country/
Year columns have no blanks, so no current test exposes it; reproduce by blanking
one Status cell in a Full_Data-complete row and observing the row survive the
filter with reference-level encoding. If this is unacceptable, the fix is extending
`Sample_Include()` with a non-blank condition on included Categorical Predictors —
one more branch in the REDUCE — and it should be raised as a design decision rather
than patched silently (ROADMAP open item 5, auto-completeness, is the home for it).

# Model Construction Sheet — Human Test Plan (WHO dataset)

Execute in order — each test is a small delta on the previous spec state, simplest
to most complex. Every test gives an **Inputs** block (the exact cells to edit on
the `Model Construction` sheet, everything else unchanged from the previous test)
and an **Expected** block (the exact cells to read and the values they must show).
Reference numbers were computed directly from
`sample_data/Life Expectancy Data.csv` by the same spec-driven mask the sheet
applies (`calculate_model_construction_expectations`), so mismatches indicate
implementation bugs, not stale expectations.

**Baseline change (Full_Data omitted).** The shipped spec no longer uses the
`Full_Data` completeness column as a Filter — it ships as **Omit**. `Full_Data`
flags rows complete on *all* 18 numeric features, which is (a) redundant with the
mask's built-in completeness (`Sample_Include()` already requires the response
and every included **continuous** predictor to be numeric) and (b) an
over-filter: it drops rows missing a sparse predictor the model does not even
use. With no Filter declared, the mask is completeness-only on the model's own
columns, so the default model now includes **2,482** rows, not 1,649. The Filter
role is exercised in T8 via a purpose-built `Is_Stan` filter column, not the
completeness flag.

**Layout note (Sequence structural-axis release):** the spec block has since
gained column H (`Sequence` flag), column I (`Sequence Period` — typed
override input), and column J (`Period In Use` — live since the base-period
release: candidate formula with typed override via column I, see T17–T19),
so the `Levels` and `Reference In Use` displays moved to **K** and **L**,
and every derived zone right of the spec block shifted one column right
(audit strip starts at N1; filtered zones at Q/V). Cell addresses below
reflect the layout at the time this plan was written — read `Levels` at
K instead of H,
and shift any address in column K or rightward by two columns. The base-period
addendum (T17–T19) uses current addresses (headers row 3, spec rows 4–26,
Sequence Spacing block rows 28–34).

**Baseline facts:** 2,938 data rows. Completeness-only on the shipped default's
three continuous predictors (`Adult Mortality`, `GDP`, `Schooling`) plus the
response is TRUE on **2,482** rows; requiring *all* 18 features (the old
`Full_Data`) is TRUE on 1,649. Within the 2,482: Status has 2 levels
(`Developed`, `Developing`), Year has 16 levels (2000–2015, first-sort reference
`2000`), Country has 157 levels. First data row: Afghanistan, 2015, Life
expectancy 65.0.

## Cell map

Spec block (columns B Role · C Include · D Type · E Reference · H Sequence ·
J Levels), one row per table column:

| Row | Variable | Row | Variable |
|---|---|---|---|
| 3 | Country | 15 | Polio |
| 4 | Year | 16 | Total expenditure |
| 5 | Status | 17 | Diphtheria |
| 6 | Life expectancy | 18 | HIV/AIDS |
| 7 | Adult Mortality | 19 | GDP |
| 8 | infant deaths | 20 | Population |
| 9 | Alcohol | 21 | thinness 1-19 years |
| 10 | percentage expenditure | 22 | thinness 5-9 years |
| 11 | Hepatitis B | 23 | Income composition of resources |
| 12 | Measles | 24 | Schooling |
| 13 | BMI | 25 | Full_Data |
| 14 | under-five deaths | 26+ | user-added table columns |

Output cells:

| Cell | Contents |
|---|---|
| K1 | audit **k** = `COLUMNS(X_s())` |
| N1 | audit **rows** = `ROWS(X_s())` — must read 2938 always |
| Q1 | audit **response** (derived name; `(none)` when no Response role) |
| S1 | audit **responses** (count of Role=Response; **red CF when ≠ 1**) |
| U1 | audit **included rows** = `SUMPRODUCT(N(Sample_Include()))` |
| J3↓ | full-height row labels (2,938 rows always) |
| K3↓ | full-height include mask (2,938 booleans always) |
| M3↓ / P3↓ | filtered row labels |
| N3↓ | filtered response values (header N2 = `y: <name>`) |
| Q2→ | header strip = `Constructed_Column_Names()` |
| Q3→ | filtered design matrix, spec order |

---

## T0 — Initial state sanity (build defaults)

**Inputs:** none. The build ships this spec:

| Variable | Role | Include | Type | Reference | Order | Transform | Sequence |
|---|---|---|---|---|---|---|---|
| Country | Identifier | FALSE | Continuous | | | None | |
| Year | Predictor | TRUE | Categorical | | | None | TRUE |
| Status | Predictor | TRUE | Categorical | | | None | |
| Life expectancy | Response | FALSE | Continuous | | | None | |
| Adult Mortality | Predictor | TRUE | Continuous | | | None | |
| Population | Omit | FALSE | Continuous | | | None | |
| GDP | Predictor | TRUE | Continuous | | | None | |
| Schooling | Predictor | TRUE | Continuous | | | None | |
| (all other numerics) | Predictor | FALSE | Continuous | | | None | |
| Full_Data | Omit | FALSE | Continuous | | | None | |

The shipped spec demonstrates the Role axis (Identifier, Response, Predictor,
Omit — Population and Full_Data both ship as Omit) and flags Year as the
Sequence axis, so the Period In Use candidate (Δ = 1), the Sequence Spacing
block, and the gated Durbin-Watson diagnostic on the Regression sheet are all
live at T0. **No Filter is declared**, so the mask is completeness-only.

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| k | K1 | 19 |
| rows | N1 | 2938 |
| response | Q1 | `Life expectancy` |
| responses | S1 | 1, not red |
| included rows | U1 | **2482** (completeness on the response + `Adult Mortality`, `GDP`, `Schooling`; no Full_Data over-filter) |
| sequence status | H2 | blank — one flag (Year) is valid; the line errors only at two-plus |
| header strip | Q2→ | `Year: 2001` … `Year: 2015` (15 cols), `Status: Developing`, `Adult Mortality`, `GDP`, `Schooling` |
| Levels | K4 / K5 | 16 / 2, blank elsewhere |
| Period In Use | Year row, col J | 1 (Year's within-country spacing; I is the override input) |
| Durbin-Watson | Regression diagnostics (Y11) | a number computed along Year, not the `n/a — requires Sequence` token |
| BFN Panel Durbin-Watson | Regression diagnostics (Y12) | `n/a — no fixed effects` (Sequence declared, no Role="Fixed Effects" row — the role ships with the v2.1 FE engine). Clearing Year's H flag flips BOTH cells to `n/a — requires Sequence` |
| gray cascade | C–H | gray on every non-Predictor row |
| y header | N2 | `y: Life expectancy` |
| first filtered label | M3 | `Afghanistan` |

- [ ] Pass

## T1 — Reproduce the all-features model (spec-driven == Full_Data here)

**Inputs:**

| Cell | Enter |
|---|---|
| B4 (Year Role) | `Identifier` |
| B5 (Status Role) | `Omit` |
| C7:C24 (Include, all 18 numerics) | `TRUE` |

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| k | K1 | 18 |
| included rows | U1 | **1649** |
| header strip | Q2→ | the 18 numeric headers, `Adult Mortality` … `Schooling`, table order |
| labels | M3 | `Afghanistan\|2015` (Country and Year both Identifier) |
| filtered y | N3 | 65 (Life expectancy, Afghanistan 2015) |
| cross-check | Q3→ row | matches the Regression sheet's first included observation cell-for-cell |

Teaching point: with **every** numeric feature in the model, spec-driven
completeness requires all 18 columns — exactly the old `Full_Data` rule — so the
count coincides at 1,649. The two diverge only when the model omits a predictor
(T2).

- [ ] Pass

## T2 — Continuous subset (the over-filter fix, made visible)

**Inputs:**

| Cell | Enter |
|---|---|
| C8:C18 (Include) | `FALSE` |
| C20:C23 (Include) | `FALSE` |

(C7 Adult Mortality, C19 GDP, C24 Schooling stay TRUE.)

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| k | K1 | 3 |
| header strip | Q2→ | `Adult Mortality`, `GDP`, `Schooling` (spec order) |
| included rows | U1 | **2482** — completeness now demands only the three predictors in the model, so 833 rows that were missing a *dropped* sparse predictor rejoin the sample (2482 > 1649) |

This 1,649 → 2,482 jump is the whole reason `Full_Data` was demoted to Omit: the
model must not lose rows to a predictor it does not use.

- [ ] Pass

## T3 — First categorical (binary)

**Inputs:**

| Cell | Enter |
|---|---|
| B5 (Status Role) | `Predictor` |
| C5 (Status Include) | `TRUE` |
| D5 (Status Type) | `Categorical` |
| E5 (Status Reference) | leave blank |

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| k | K1 | 4 |
| Levels | J5 | 2 |
| header strip | Q2→ | `Status: Developing`, `Adult Mortality`, `GDP`, `Schooling` |
| dummy values | Q3↓ (Status column) | only 0/1 |
| included rows | U1 | 2482 (a Categorical Predictor imposes no numeric condition) |
| rows | N1 | 2938 |

Reference defaulted to `Developed` (first in sort order — surfaced by omission in
the column name).

- [ ] Pass

## T4 — Reference override

**Inputs:** E5 → `Developing`. Observe. Then clear E5.

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| strip column | Q2 | flips to `Status: Developed`; 0/1 pattern inverts; k stays 4 |
| after clearing E5 | Q2 | reverts to `Status: Developing` |

- [ ] Pass

## T5 — Invalid reference (visible failure)

**Inputs:** E5 → `Developped` (typo). Observe. Then clear E5.

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| invalid-reference CF | E5 | red fill/font |
| k | K1 | 3 — Status contributes **zero** columns (the constructor's `ISNA` skip), header strip loses `Status: …` |
| no raw error | all zones | no `#N/A` / `#CALC!` leaks; every other output still computes |
| after clearing E5 | K1 | back to 4 |

Visible degradation, not a hard error, not silent full one-hot (the collinearity
trap the reference-drop exists to prevent). Conditional formatting may need a
recalculate to render.

- [ ] Pass

## T6 — Numeric-valued categorical

**Inputs:**

| Cell | Enter |
|---|---|
| B4 (Year Role) | `Predictor` |
| C4 (Year Include) | `TRUE` |
| D4 (Year Type) | `Categorical` |

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| k | K1 | 19 |
| Levels | J4 | 16 |
| header strip | Q2→ | fifteen columns `Year: 2001` … `Year: 2015`, then `Status: Developing`, `Adult Mortality`, `GDP`, `Schooling` |
| included rows | U1 | 2482 |
| dummy values | Year columns | each 0/1 |
| labels | M3 | `Afghanistan` (Country is again the sole Identifier) |

Reference `2000` — numeric sort, not text sort; if `Year: 2013` is missing or the
ordering is weird, numeric/text coercion is broken.

- [ ] Pass

## T7 — Identifier labeling variants

**(a) Inputs:** none (current state: Country Identifier, Year Predictor).
**Expected:** J3/M3 = `Afghanistan` (Country only).

**(b) Inputs:** B4 → `Identifier`.
**Expected:** labels become `Afghanistan|2015`; k drops to 4 (Year left the
model); U1 stays 2482.

**(c) Inputs:** B3 → `Omit`, B4 → `Omit`.
**Expected:** J3 = `Obs. 1`, `Obs. 2`, … full height; M3 = `Obs. 1` (row 1 —
Afghanistan|2015 — passes the mask; if the first included row were later, the
filtered column would start at that observation number, not `Obs. 1`).

**Restore:** B3 → `Identifier`, B4 → `Identifier` (labels `Afghanistan|2015`;
Status Predictor/Categorical, three continuous predictors — k = 4, U1 = 2482).

- [ ] Pass

## T8 — Filter composition (the `stan` filter)

**Inputs:** on the `Life Expectancy Data` sheet, type `Is_Stan` in the first
header cell right of the table (the table auto-expands) and in the first data
cell below it the formula `=--(RIGHT([@Country],4)="stan")` — 1 for every country
whose name ends in "stan" (Afghanistan, Kazakhstan, Kyrgyzstan, Pakistan,
Tajikistan, Turkmenistan, Uzbekistan). Back on `Model Construction`, its spec row
appears automatically at row 26 (column A spills from the headers; the B–H
dropdowns already cover it):

| Cell | Enter |
|---|---|
| B26 (Is_Stan Role) | `Filter` |

(State entering T8: Country Identifier, Year Identifier, Status
Predictor/Categorical, three continuous predictors.)

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| included rows | U1 | **96** — completeness-on-the-model AND `Is_Stan` = 1 (7 "-stan" countries, 96 rows complete on the three predictors; Kyrgyzstan contributes no complete row) |
| Levels | J5 | collapses 2 → **1**, cell turns **red** (included, categorical, L ≤ 1) — every "-stan" country in the sample is `Developing`, so Status degenerates over the masked stratum |
| k | K1 | 3 — Status contributes **zero** columns while every other output still computes |
| filtered zones | M3↓ | exactly 96 rows; first label still `Afghanistan\|2015` |

Visible degradation, not a hard error, not silent omission — the filter both
demonstrates the Filter role and drives Status into the masked-degeneracy path.
(Optional inspect: Country levels within the stratum = **6** — set B3 →
Predictor/Categorical, read J3, then restore B3 → Identifier.)

- [ ] Pass

## T9 — Filter semantics edge (and why Full_Data is Omit)

**Inputs (step 1):** B26 (Is_Stan Role) → `Omit`.
**Expected:** U1 returns to **2482**; J5 back to 2; K1 back to 4 — with no Filter
declared the mask is completeness-only.

**Inputs (step 2 — the over-filter demonstration):** B25 (Full_Data Role) →
`Filter`.
**Expected:**

| Check | Cell | Expect |
|---|---|---|
| included rows | U1 | **1649** — re-imposing the all-features completeness flag drops 833 rows whose only missing values are in predictors the model does not use |
| k | K1 | 4 (Status keeps both levels in the narrower mask) |

This is exactly the redundant over-filter the shipped default avoids by leaving
`Full_Data` as Omit.

**Restore:** B25 (Full_Data Role) → `Omit` (U1 = 2482).

- [ ] Pass

## T10 — Response swap (derived y)

**Inputs:** (state entering T10: Country Identifier, Year Identifier, Status
Predictor/Categorical, three continuous predictors — k = 4.)

| Cell | Enter |
|---|---|
| B6 (Life expectancy Role) | `Predictor` |
| C6 (Life expectancy Include) | `TRUE` |
| B19 (GDP Role) | `Response` |

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| response | Q1 | `GDP` |
| y header | N2 | `y: GDP` |
| filtered y | N3 | GDP values (first ≈ 584.26) |
| responses | S1 | 1, not red |
| header strip | Q2→ | `Status: Developing`, `Life expectancy`, `Adult Mortality`, `Schooling` |
| k | K1 | 4 |
| included rows | U1 | 2482 (the required column set — GDP + Life expectancy/Adult Mortality/Schooling — is the same four columns as before) |

Do **not** restore yet — T11 continues from this state.

- [ ] Pass

## T11 — Response-count validation

**(a) Inputs:** B19 → `Predictor` (GDP and Life expectancy both Predictor — zero
Responses).
**Expected:**

| Check | Cell | Expect |
|---|---|---|
| responses | S1 | 0, **red** |
| response | Q1 | `(none)` |
| y header | N2 | `y: (none)` |
| filtered y | N3 | `(empty model)` — no fabricated column, no raw error |

**(b) Inputs:** B6 → `Response` AND B19 → `Response` (two Responses).
**Expected:**

| Check | Cell | Expect |
|---|---|---|
| responses | S1 | 2, **red** — the only alarm |
| response | Q1 | `Life expectancy` — first match in **table order** wins (Life expectancy is column 4, GDP is column 17) |
| filtered y | N3 | Life expectancy values (65 first) |

**Restore:** B19 → `Predictor`, C19 → `TRUE`, C6 → `FALSE` (S1 = 1; K1 = 4).

- [ ] Pass

## T12 — Levels display is live and mask-aware (no model change)

**Inputs:**

| Cell | Enter |
|---|---|
| B3 (Country Role) | `Predictor` |
| D3 (Country Type) | `Categorical` |
| C3 (Country Include) | `FALSE` |

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| Levels | J3 | **157** — computed over the 2,482 mask-included rows (wider than the 133 the old 1,649 Full_Data mask showed); **no red** (not included) |
| k | K1 | unchanged (4) |
| labels | M3 | `2015` — Country left the Identifier role, so Year is now the sole Identifier; expected side effect, reverts on restore |

Then C3 → `TRUE` briefly: K1 jumps to **160** (+156 Country dummies) and the
header strip floods with `Country: …` names — the visible-count-as-warning
behavior.

**Restore:** C3 → `FALSE`, B3 → `Identifier`, D3 → `Continuous` (labels
`Afghanistan|2015`).

- [ ] Pass

## T13 — Extreme stratification degeneracy (both categoricals collapse)

**Inputs:** on `Life Expectancy Data`, add table column `Is_2015` with formula
`=--([@Year]=2015)`. Keep the `Is_Stan` column from T8. Then:

| Cell | Enter |
|---|---|
| B26 (Is_Stan Role) | `Filter` |
| B27 (Is_2015 Role) | `Filter` |
| B4 (Year Role) | `Predictor`, C4 → `TRUE`, D4 → `Categorical` |

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| included rows | U1 | **6** — "-stan" AND 2015 AND complete on the three predictors (one row per qualifying country; Kyrgyzstan's 2015 row is incomplete) |
| Levels | J4 | 1 → **red**; Year contributes 0 columns (all six rows are 2015) |
| Levels | J5 | 1 → **red** — all six are Developing, so Status degenerates too |
| k | K1 | 3 (`Adult Mortality`, `GDP`, `Schooling` only) |
| filtered zones | M3↓ | exactly 6 rows; first label `Afghanistan` (Country sole Identifier); N3 = 65 |
| no errors | all zones | construction zone must not error; any fit on 6 rows is degenerate *downstream* |

**Restore:** B26 → `Omit`, B27 → `Omit`, B4 → `Identifier`.

- [ ] Pass

## T14 — Empty model

**Inputs:** C3:C25 (every Include) → `FALSE`.

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| k | K1 | `(empty model)` |
| rows | N1 | `(empty model)` |
| header strip | Q2 | `(empty model)` |
| matrix | Q3 | `(empty model)` |
| still alive | M3 / N3 / U1 / S1 | labels, y values, **2928**, 1 — with every predictor excluded and no Filter, the mask reduces to "response is numeric" (10 rows have a blank Life expectancy); the mask and response are spec-driven, not X_s-driven |
| no leaks | all zones | no `#CALC!` anywhere |

**Restore T0** (full reset):

| Cell | Enter |
|---|---|
| Life Expectancy Data | delete the `Is_Stan` and `Is_2015` table columns (spec rows 26–27 disappear) |
| B3 / C3 / D3 | `Identifier` / `FALSE` / `Continuous` |
| B4 / C4 / D4 | `Predictor` / `TRUE` / `Categorical` |
| B5 / C5 / D5 | `Predictor` / `TRUE` / `Categorical` |
| B6 / C6 / D6 | `Response` / `FALSE` / `Continuous` |
| C7, C19, C24 | `TRUE` |
| C8:C18, C20:C23 | `FALSE` |
| B20 (Population) / B25 (Full_Data) | `Omit` / `Omit` |
| E3:E25 | blank |

Verify the T0 Expected block again before continuing.

- [ ] Pass

## T15 — Full-height vs filtered contract

**Inputs:** none (T0 state).

**Expected:**

| Check | Cell | Expect |
|---|---|---|
| full-height labels | J3↓ | exactly 2938 rows |
| full-height mask | K3↓ | exactly 2938 booleans |
| filtered labels / y | M3↓, N3↓, P3↓ | exactly 2482 rows |
| filtered matrix | Q3→↓ | 2482 × 19 |
| rows audit | N1 | 2938 — always; the constructor never row-filters (the row-mask contract); only the display zones do |

- [ ] Pass

## T16 — Twin alignment tripwire

**Inputs:** none — check in any/every state from the tests above, including
mid-edit states (degenerate categorical, invalid reference, empty model).

**Expected:** the header strip width (count of spilled cells right of Q2) equals
the K1 audit cell, always. If they ever disagree, the `X_s` /
`Constructed_Column_Names` twins have drifted — a structural bug, not a data
issue.

- [ ] Pass

---

## Known caveat to verify and accept (or escalate)

**Blank categorical values do not fail the completeness mask.** As implemented
(per the current design decision), `Sample_Include()` demands numeric y + numeric
included *continuous* predictors; included **Categorical** Predictors impose no
completeness condition. A blank category value in an otherwise-complete row stays
included and encodes as all-zero dummies — indistinguishable from the reference
level. The ROADMAP language suggests such rows should be excluded instead.

**Steps:** on `Life Expectancy Data`, note then delete the Status value of a
mask-complete row (e.g., the first row). On `Model Construction`, check
whether U1 drops by 1 (row excluded) or holds (row included, encoded as
reference). Restore the cell afterward.

**Record the outcome:** if the row remains included, either **accept** (document
the caveat as intended v3.0 behavior) or **escalate** (extend `Sample_Include()`
with a non-blank condition for included Categorical Predictors in a follow-up).

- [ ] Verified — decision: ______________

---

## Sign-off

| Field | |
|---|---|
| Executed by | |
| Date | |
| Workbook build | |
| Result | |

---

# Addendum — Base Period Δ & gap-aware Difference_By / Lag_By (T17–T19)

Execute on the **Regression** sheet's spec block (the production home of the
spec; `Base_Period_Delta()` — the omitted-`[delta]` default — reads the
Regression sheet by name). **Current addresses:** headers row 3; spec rows 4–26
(Country 4, Year 5, Status 6, Life expectancy 7, …, Full_Data 26); Sequence
Spacing block rows 28–34 (Δ candidate B29, Δ in use B30, verdict lines
A31–A34, delta spectrum spilling at J30:K…).

Semantics under test: Δx(i,t) = x(i,t) − x(i,t−Δ) with (A) each group's first
period `#N/A` — never a fabricated 0; (B) "prior period" = the literal time
value t−Δ — a gap returns `#N/A`, never the previous available row; (C) Δ
computed-and-displayed-with-override, never a silent 1. Expected numbers below
are asserted programmatically in `tests/test_difference_by_verification.py`
from the same CSV. The spectrum counts within-group Year spacings and is
**independent of the completeness mask** (`Sequence_Deltas` ignores
`Sample_Include()`), so the Full_Data change does not move these numbers.

## T17 — WHO panel: Δ candidate, spectrum, and gap-aware differencing

**Inputs:** none required — the shipped default already flags **Year as the
Sequence axis** (H5 = TRUE, Country is the Identifier, so groups = Country). If
you cleared it in an earlier test, re-set **H5 = TRUE**.

**Expected:**

- **I5** shows **1** — the candidate (MODE of within-group consecutive Year
  spacings) populates the spec cell.
- **B29 (Δ candidate) = 1**; **B30 (Δ in use) = 1**, no yellow highlight.
- Spectrum: **J30 = 1, K30 = 2745** and no second row — the delta spectrum is
  exactly {1: 2745}.
- **A31–A34 all blank** (Regularity quiet, Off-grid quiet, no
  no-natural-base-period prompt, no calendar guidance).

On the `Life Expectancy Data` sheet, in empty columns right of the table:

| Formula | Expected |
|---|---|
| `=Difference_By(LifeExpectancyData[Life expectancy],LifeExpectancyData[Country],LifeExpectancyData[Year])` (spills 2,938 rows; `[delta]` omitted → spec Δ) | first cell (Afghanistan 2015) = **5.1** (= 65.0 − 59.9) |
| `=SUMPRODUCT(--ISNA(<spill>#))` | **193** — one `#N/A` per country: 183 first periods + all 10 single-observation countries (Cook Islands' one row is `#N/A`, not 0 and not blank) |
| `=SUMPRODUCT(--ISNUMBER(<spill>#))` | **2745**; 193 + 2745 = **2938** — no row is mask-excluded |

**Override visibility (C):** type **2** into I5 → B30 = 2 with the **yellow**
override highlight; A31 Regularity fires (spacings besides Δ exist); A32
Off-grid fires red (1 is not a whole multiple of 2); the data-sheet `ISNA`
count rises to **376** (each 16-year panel now has two start years, 2000 and
2001, with no t−2 partner: 183×2, plus the 10 single rows) and the numeric
count falls to 2,562. Restore by re-entering the candidate formula in I5:
`=IF($H5<>TRUE,"",IFERROR(Base_Period_Delta_Candidate(),""))` — B30 returns
to 1, the highlight clears, and the verdicts go quiet.

- [ ] Pass

## T18 — Synthetic: punched-out year (gap on the Δ grid)

**(a) Direct function check — scratch sheet.** Copy Albania's 16
(Year, Life expectancy) pairs to a blank sheet, delete the **2007** row
(15 rows remain), add a constant group column `Albania`, then:

`=Difference_By(<x 15-rows>, <group 15-rows>, <year 15-rows>, 1)`

**Expected:** `#N/A` at **2000** (first period, A) and at **2008** (2007
absent → gap, B — the 2008 value must NOT become x(2008) − x(2006));
**13** numeric differences; 15 rows total. Delta spectrum of this fragment
(verified in the Python script): **{1: 13, 2: 1}**, candidate = 1.

**(b) Block check — on the live table.** Delete the Albania 2007 row from
`LifeExpectancyData` (T17 state still applied: H5 = TRUE, I5 = candidate).

**Expected:** spectrum shows **{1: 2743, 2: 1}** (whole table: Albania trades
15 one-spacings for 13 ones + one two); **I5 / B29 / B30 still 1**;
**A31 Regularity fires** (yellow, "spacings besides Δ"); **A32 Off-grid stays
quiet** (2 is a whole multiple of 1); no calendar guidance. The T17
data-sheet `ISNA` count reads **194** (the 193 first periods + Albania 2008).
**Undo (Ctrl+Z)** and confirm the spectrum returns to {1: 2745}.

- [ ] Pass

## T19 — Synthetic: calendar dates (spacing that must NOT be quantized)

**Inputs:**

1. Add a table column `Month_Start`; in its first 24 data rows enter the
   first-of-month dates Jan 2020 – Dec 2021 (`=DATE(2020,1,1)`, then
   `=EDATE(prev,1)` filled down; paste as values), leave the rest blank.
   2020 is a leap year. A new spec row (27) appears for it.
2. Set **B4 (Country) = Omit** — no Identifier columns remain, so the whole
   sample is one group and only the 24 dated rows carry the axis.
3. Set **H27 = TRUE** (clear H5 first — the H2 status line must stay blank).

**Expected:**

- **B29 (Δ candidate) = 31** — MODE of the day-count spacings, which would be
  a *wrong* Δ for eleven months of the year; the tool surfaces rather than
  quantizes.
- Spectrum (ascending, 23 spacings): **J/K rows = (28, 1), (29, 1), (30, 8),
  (31, 13)**.
- **A34 calendar-signature guidance fires (red)**: recommends building an
  integer period index upstream (`YEAR`, or `YEAR*12+MONTH`) and flagging
  that as the Sequence axis instead of a day-count Δ.
- **A33 (no natural base period) stays quiet** — MODE is defined here.
- Row 27 has no pre-filled candidate formula (build pre-fills rows 4–26
  only), so B30 reads "(not set)" and the Δ-keyed verdicts hold fire until a
  Δ is typed: enter **31** into **I27** → B30 = 31 (no override highlight —
  it matches the candidate), **A31 Regularity fires** (28/29/30 exist) and
  **A32 Off-grid fires red** (28/29/30 are not whole multiples of 31).
- Optional direct check (scratch): `=Difference_By(<x>, <group>, <dates>, 31)`
  over the 24 dated rows yields **13** numeric values (only months preceded
  by a 31-day month) and **11** `#N/A` — commitment (B) on off-grid data.

**Restore:** delete the `Month_Start` table column, set B4 back to
`Identifier (Row Label)`, clear H27/I27, and re-set H5 = TRUE (Year the
shipped Sequence axis).

- [ ] Pass

## Addendum sign-off

| Field | |
|---|---|
| Executed by | |
| Date | |
| Workbook build | |
| Result | |

"""Build the Model Construction worksheet — v3.0 declarative specification block.

Two-axis specification plus the Sequence structural axis (ROADMAP: v3.0 —
Specification-Driven Regression; Sequence added post-v2.0):

    A        B      C       D     E               F       G         H        I             J      K
    Variable Role   Include Type  Reference Level Order   Transform Sequence Base Period Δ Levels Reference In Use
    (spill)  (drop) (input) (drop)(input)         (rsvd.) (rsvd.)   (flag)   (cand./ovr.)  (disp.)(disp.)

Right of the spec block, after a narrow gap column L (which also visually
reserves the future Design Columns audit column):

    M             N        O     P           Q            R     S           T →
    Row Labels    Included (brk) Filt.Labels Filt.y       (brk) Filt.Labels Filtered X_s
    (=Row_Labels() spill at M4; =Sample_Include() spill at N4 — both
     full-height, never internally filtered. P/Q/S/T are the FILTERED
     display zones: the only place on the sheet where Sample_Include()
     row-filters anything. S repeats the filtered labels so the matrix
     reads side-by-side without scrolling back to P.)

Row 1, from column M rightward, holds the bold audit cells as
label/value pairs (values on the non-narrow columns N/Q/T/V/X/Z):

    k = COLUMNS(X_s()) · rows = ROWS(X_s()) · response = <derived name> ·
    responses = <count of Role="Response (y)"> (red CF when <> 1) ·
    included rows = SUMPRODUCT(N(Sample_Include())) ·
    sequence flags = <count of Sequence=TRUE> (red CF when > 1)

Row 3 above R carries the =Constructed_Column_Names() header strip
(level-qualified names, horizontal). Every spill formula in the filtered
zones wraps IFERROR(..., "(empty model)") so an empty model degrades to a
documented string, never a raw #CALC! leak.

Row 2 is a model-level control row above the spec table: A2 labels
"Intercept" and C2 is the Allow_Intercept toggle, sitting at the top of the
C/Include boolean column (mirroring the v1 Regression sheet's A2/B2
control). It has no v3.0 consumer yet — the engine will read it. Because of
this control row the spec headers move to row 3 and the 23 variable rows to
4–26; the row-1 audit strip is unaffected.

The spec spans EVERY column of the LifeExpectancyData table (23 rows:
[Country]..[Schooling] plus [Full_Data]). Two axes:

    Variable Role  — Response (y) | Predictor (x) | Identifier (Row Label) |
                     Filter | Omit
                     (what the column IS; future: Fixed Effects/Weight/Time.
                      The parenthetical glosses are part of the stored token —
                      see the _ROLE_* constants)
    Predictor Type — Continuous | Categorical
                     (how a Predictor ENTERS; meaningful only when
                      Role = Predictor; this axis never grows)
    Sequence       — TRUE | blank (structural axis, column H)
                     (which column ORDERS the data for lag/difference/
                      serial-correlation features. Deliberately NOT a Role
                      value and NOT a Predictor Type — a column can be
                      Role = Predictor, Type = Continuous AND Sequence = TRUE
                      simultaneously. At most one flag: two-plus is a
                      status-line error (H2, same pattern as the
                      exactly-one-Response audit); zero is valid (non-panel
                      data). Not to be confused with the reserved Order
                      column F, which is term-ordering.)

Columns F (Order) and G (Transform) are reserved for a future release: they
are styled as inputs and carry sheet-scoped names (Spec_Order, Spec_Transform)
so the grid shape is final, but no formula reads them yet. Column I
(Base Period Δ) is Sequence's LIVE companion since the base-period release —
computed-with-override, the reference-level pattern: every row is pre-filled
with a formula that shows Base_Period_Delta_Candidate() (MODE of the
within-group consecutive spacings, MIN fallback when no spacing repeats) on
the Sequence-flagged row and blank elsewhere; typing a number over the
flagged row's cell overrides the candidate. The workbook-scoped
Base_Period_Delta() accessor reads whatever that cell shows and is the
default [delta] of Lag_By / Difference_By — Δ is never a silent 1.
The Sequence Spacing block below the spec (rows 28–34) surfaces the
candidate, the Δ in use (yellow when overridden), the delta spectrum
(distinct within-group spacings with counts, spilling at J/K), and four
verdict lines: Regularity (any spacing ≠ Δ), Off-grid (a spacing that is
not a whole multiple of Δ), the no-natural-base-period override prompt
(MODE undefined → MIN fallback in use), and the calendar-signature guidance
(spacings clustered at ~28–31 / ~90–92 / ~365–366 → recommend an integer
period index upstream instead of quantizing day counts to a scalar Δ).
The spectrum deliberately ignores Sample_Include(): the time grid is
dataset structure, not model-iteration state. No CONSTRUCTOR reads H or I —
the base-period layer (candidate closure, spectrum, accessor) is the only
consumer, and every read is TAKE-trimmed to COLUMNS(Source_Data) so the
block's own cells sitting inside the 16000-row Spec_* bands are never
scanned as spec rows.

The role system dissolves the v1 Regression sheet's three hard-wired names
into declarations, all late-bound zero-argument LAMBDAs on this sheet:
    y                          → Response_Column(): the Role = "Response" row
    Regression_Sample_Include  → Sample_Include(): per-row AND computed as a
                                 REDUCE product of indicator vectors — every
                                 Role = Filter column truthy (TRUE/1 pass;
                                 FALSE/0/blank/text fail), the Response
                                 numeric, every included Continuous Predictor
                                 numeric. Categorical Predictors impose no
                                 completeness condition (known caveat — see
                                 the human test plan); Identifier/Omit
                                 impose nothing
    All_Xs                     → Source_Data, derived (with Header_Names) from
                                 Source_Table — the ONE name wrapping the
                                 table reference, so a dataset changeover is
                                 a one-name edit (structured refs can't be
                                 parameterized without volatile INDIRECT)

Constructor decisions (all settled — see ROADMAP):
1. Spec-order assembly — REDUCE over spec rows, HSTACK from a full-height
   sentinel seed, DROP the sentinel. Iteration predicate: Role = Predictor
   AND Include = TRUE.
2. Level-vector split — the workbook-scoped Dummy_Levels (mask-aware
   SORT/UNIQUE, reference dropped, #N/A on degenerate input; synced from
   lambda_functions.json) is the single source of truth for retained level
   sets; encoding is --(column = retained_levels). A blank Reference cell
   means "use the default" — Dummy_Levels treats a provided "" the same as
   an omitted reference (first level in sort order).
3. Degenerate skip — a Categorical Predictor whose Dummy_Levels call
   returns #N/A (degenerate or invalid reference) contributes nothing
   (acc passthrough), flagged red on the sheet by conditional formatting:
   visible degradation, not a hard error, not silent omission.

Default configuration (the human test plan's T0 state):
    Country          → Identifier            (residual labeling; no columns)
    Year             → Predictor/Categorical/TRUE  (numeric-valued)
    Status           → Predictor/Categorical/TRUE
    Life expectancy  → Response               (derived y)
    Adult Mortality, GDP, Schooling → Predictor/Continuous/TRUE
    Population        → Omit                   (deliberate exclusion demo)
    remaining numerics → Predictor/Continuous/FALSE (candidates)
    Full_Data        → Omit                    (its all-features completeness
                                               flag is redundant with the mask's
                                               built-in completeness and
                                               over-filters; no default Filter)
Full-height contract: ROWS(X_s()) = ROWS(Row_Labels()) =
ROWS(Sample_Include()) = 2938 always — the constructor reads the mask ONLY
to fix level sets; nothing here ever row-filters. With the real mask live,
the T0 mask-dependent values are real on the sheet: k = 19 (15 Year
dummies), and SUMPRODUCT(N(Sample_Include())) = 2482 (completeness-only on
the response and the three continuous predictors, no Full_Data over-filter).

Not here (deliberately, per release scoping): the QC analyzer
(analyze_model_construction.py) and the Version History / CHANGELOG bump
to v3.0 — those land in the final wiring PR.
"""
from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

import xlwings as xw

from .catalog_schema import CatalogFunction, load_catalog_document
from .lambda_formula_parser import (
    _normalize_user_formula,
    _strip_non_string_whitespace,
)
from .sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_DARK_YELLOW_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
    MUTED_TEXT_COLOR,
)
from .workbook_helpers import (
    add_expression_format,
    bold,
    bold_row,
    col_letter,
    drop_local_name,
    f,
    format_input,
    get_or_create_sheet,
    open_or_create_workbook,
    rc,
    reset_generated_sheet,
    section_heading,
    val,
)

SHEET_NAME = "Model Construction"

# The constructor closures moved to the Regression sheet with the v3.0
# changeover (scope "Regression" in lambda_functions.json) — the spec block
# now lives there and this sheet is no longer part of the production build.
# A standalone rebuild of this sheet still works: the closures are generic
# over the sheet they're registered on, so they are loaded by this scope and
# installed sheet-scoped here exactly as before.
_CLOSURE_SCOPE = "Regression"

# The catalog file backing the constructor closures. Used when a caller does
# not pass them in explicitly.
_DEFINITIONS_PATH = Path(__file__).resolve().parent.parent / "lambda_functions.json"

# Every LifeExpectancyData column, in table order (incl. the computed
# Full_Data completeness column — the spec spans the whole table).
_VARIABLES: list[str] = [
    "Country",
    "Year",
    "Status",
    "Life expectancy",
    "Adult Mortality",
    "infant deaths",
    "Alcohol",
    "percentage expenditure",
    "Hepatitis B",
    "Measles",
    "BMI",
    "under-five deaths",
    "Polio",
    "Total expenditure",
    "Diphtheria",
    "HIV/AIDS",
    "GDP",
    "Population",
    "thinness 1-19 years",
    "thinness 5-9 years",
    "Income composition of resources",
    "Schooling",
    "Full_Data",
]
_N_VARIABLES = len(_VARIABLES)  # 23

# Row 2 is the model-level Intercept control (label A2, toggle C2 — aligned
# to the C/Include boolean column). The spec table sits one row below it:
# headers on row 3, the 23 variable rows on 4–26.
_INTERCEPT_ROW = 2
_HEADER_ROW = 3
_FIRST_DATA_ROW = 4
_LAST_DATA_ROW = _FIRST_DATA_ROW + _N_VARIABLES - 1  # 26
# Sheet row _FIRST_DATA_ROW maps to Source_Data column 1, so a row-indexed
# formula recovers its column via INDEX(Source_Data,0,ROW()-_ROW_TO_COL_OFFSET).
_ROW_TO_COL_OFFSET = _FIRST_DATA_ROW - 1  # 3

# Spec-block columns (1-based). Role precedes Include: the larger
# declaration comes first (dataset semantics before iteration state).
# F/G are the reserved Order/Transform slots; H is the Sequence structural
# flag with I as its reserved Base Period Δ companion; J and K are the
# computed displays (Levels count, Reference In Use).
(
    _C_LABEL,
    _C_ROLE,
    _C_INCLUDE,
    _C_TYPE,
    _C_REFERENCE,
    _C_ORDER,
    _C_TRANSFORM,
    _C_SEQUENCE,
    _C_BASE_PERIOD,
    _C_LEVELS,
    _C_REF_IN_USE,
) = range(1, 12)

# Derived-row zone right of the spec block. L is a narrow gap (and the
# visual reservation for the future Design Columns audit column); M and N
# hold the full-height Row_Labels() / Sample_Include() spills.
_C_GAP = 12
_C_ROW_LABELS = 13
_C_INCLUDED = 14
_GAP_COLUMN_WIDTH = 2

# Filtered display zone: the ONLY place Sample_Include() row-filters
# anything (everything left of O honors the full-height contract). O and
# R are narrow visual breaks; S repeats the filtered labels so the matrix
# reads side-by-side without scrolling back to P.
_C_BREAK_LEFT = 15
_C_FILTERED_LABELS = 16
_C_FILTERED_Y = 17
_C_BREAK_MID = 18
_C_MATRIX_LABELS = 19
_C_MATRIX_START = 20

# Row-1 audit strip: label/value pairs marching right from column M,
# values placed on the non-narrow columns (N, Q, T, V, X, Z) so no number
# lands on a width-2 break column.
_AUDIT_ROW = 1
_AUDIT_PAIRS: tuple[tuple[int, int], ...] = (
    (_C_ROW_LABELS, _C_INCLUDED),          # k
    (_C_FILTERED_LABELS, _C_FILTERED_Y),   # rows
    (_C_MATRIX_LABELS, _C_MATRIX_START),   # response
    (_C_MATRIX_START + 1, _C_MATRIX_START + 2),  # responses (red CF <> 1)
    (_C_MATRIX_START + 3, _C_MATRIX_START + 4),  # included rows
    (_C_MATRIX_START + 5, _C_MATRIX_START + 6),  # sequence flags (red CF > 1)
)

_EMPTY_MODEL_FALLBACK = '"(empty model)"'

# Role tokens — the exact strings stored in column B and compared by every
# role-driven formula, here and in the catalog closures (Sample_Include,
# Response_Column, Row_Labels, X_s, Constructed_Column_Names in
# lambda_functions.json). The parenthetical glosses are part of the stored
# value, not display-only: renaming a token means updating the JSON
# closures in lockstep.
_ROLE_RESPONSE = "Response (y)"
_ROLE_PREDICTOR = "Predictor (x)"
_ROLE_IDENTIFIER = "Identifier (Row Label)"
_ROLE_FILTER = "Filter"
_ROLE_OMIT = "Omit"
# The v2.1 panel role (ROADMAP Role-axis values). Forward wiring only: the
# token is read by the Fixed_Effects_Column() accessor and the FE-count guard
# on the Regression diagnostics (BFN panel Durbin-Watson trigger matrix), but
# it is deliberately NOT in the Role dropdown yet — the design-matrix engine
# does not absorb fixed effects until the v2.1 release, and offering the role
# before then would let a spec claim FE while silently fitting pooled OLS.
_ROLE_FIXED_EFFECTS = "Fixed Effects"

# The derived response name, shared by the audit strip and the filtered-y
# header: the header of the first Role=Response spec row, "(none)" when
# no row carries the role. XMATCH position over the TAKE-trimmed roles is
# the same lookup Response_Column() uses for its data column.
_RESPONSE_NAME_FORMULA = (
    'IFERROR(INDEX(TOROW(Header_Names),'
    f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data)))),"(none)")'
)

# Dropdown validations cover the repo's standard 16000-row input band so a
# retargeted dataset with more columns inherits them without a rebuild.
_VALIDATION_LAST_ROW = 16000

# Default spec: variable -> (role, include, type). Reference (E) starts
# blank everywhere so the first-in-sort-order default is what gets
# exercised; type an explicit level into E to exercise the override path.
# The shipped T0 spec demonstrates the Role axis (Response, Predictor,
# Identifier, Omit), both Types (Continuous, Categorical), and — via
# _DEFAULT_SEQUENCE_VARIABLES below — the structural Sequence axis with its
# Base Period Δ companion. Population is shipped as an explicit Omit (WHO
# population figures are notoriously incomplete/inconsistent) so the Omit role
# and its graying are demonstrated; Omit contributes no column and imposes no
# mask condition, leaving the fitted model identical to a plain excluded row.
#
# Full_Data ships as Omit, NOT Filter: the Full_Data completeness column
# demands EVERY numeric feature be present, which is (a) redundant with the
# built-in completeness the mask already applies to the response and the
# model's included continuous predictors, and (b) an over-filter — it drops
# rows missing a sparse predictor the model does not even use. With no Filter
# declared, the shipped model includes every row complete on its OWN columns
# (2482, vs 1649 under Full_Data). The Filter role is exercised in the human
# test plan via a purpose-built filter column, not the completeness flag.
_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Country": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Year": (_ROLE_PREDICTOR, True, "Categorical"),
    "Status": (_ROLE_PREDICTOR, True, "Categorical"),
    "Life expectancy": (_ROLE_RESPONSE, False, "Continuous"),
    "Adult Mortality": (_ROLE_PREDICTOR, True, "Continuous"),
    "Population": (_ROLE_OMIT, False, "Continuous"),
    "GDP": (_ROLE_PREDICTOR, True, "Continuous"),
    "Schooling": (_ROLE_PREDICTOR, True, "Continuous"),
    "Full_Data": (_ROLE_OMIT, False, "Continuous"),
}
_FALLBACK_SPEC: tuple[str, bool, str] = (_ROLE_PREDICTOR, False, "Continuous")

# Variables shipped with their Sequence flag (column H) set TRUE. Year is the
# canonical ordering axis for the WHO Country/Year panel: flagging it activates
# the Base Period Δ candidate (Δ = 1), the Sequence Spacing block, and the
# gated Durbin-Watson diagnostic on the Regression sheet. Structural and
# Role-independent — Year stays a Categorical Predictor, so the fitted model is
# unchanged; the flag only drives the serial-correlation / base-period layer.
# Kept to at most one entry: the H2 status line errors at two-plus flags.
_DEFAULT_SEQUENCE_VARIABLES: frozenset[str] = frozenset({"Year"})

_DEFAULT_TRANSFORM = "None"

_ROLE_VALIDATION_LIST = ",".join(
    (_ROLE_RESPONSE, _ROLE_PREDICTOR, _ROLE_IDENTIFIER, _ROLE_FILTER, _ROLE_OMIT)
)
_INCLUDE_VALIDATION_LIST = "TRUE,FALSE"
_TYPE_VALIDATION_LIST = "Continuous,Categorical"
_TRANSFORM_VALIDATION_LIST = _DEFAULT_TRANSFORM
# Sequence flag: TRUE or blank (IgnoreBlank keeps blank legal).
_SEQUENCE_VALIDATION_LIST = "TRUE"
_XL_VALIDATE_LIST = 3
_XL_VALID_ALERT_STOP = 1
_XL_BETWEEN = 1
_RESERVED_NOTE = "Reserved for a future release — not yet used by any formula."
_SEQUENCE_NOTE = (
    "Sequence structural axis: mark AT MOST ONE variable TRUE as the "
    "ordering axis for lag/difference/serial-correlation features. "
    "Independent of Role and Predictor Type — a Predictor can also be the "
    "sequence axis. Zero flags is valid (non-panel data); two or more is a "
    "spec error shown in the status line above this column. Distinct from "
    "the reserved Order column (F), which is term-ordering."
)
_BASE_PERIOD_NOTE = (
    "Base Period Δ — computed-with-override. On the Sequence-flagged row this "
    "cell shows the computed candidate (the MODE of within-group consecutive "
    "spacings; MIN when no spacing repeats). Type a number over it to "
    "override; Lag_By and Difference_By use this cell when [delta] is "
    "omitted. Blank on rows that are not the sequence axis. See the Sequence "
    "Spacing block below the spec for the delta spectrum and verdicts."
)

# Count of Sequence flags across the live spec rows — the zero-or-one
# validation shared by the H2 status line, the audit strip, and the
# multi-flag conditional format (same TAKE-trimmed idiom as the
# responses count).
_SEQUENCE_FLAG_COUNT_FORMULA = (
    "SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))"
)

# Count of Role="Fixed Effects" spec rows — the FE-active detector behind the
# Regression sheet's serial-correlation trigger matrix (plain DW vs. the BFN
# panel DW). Always 0 until the v2.1 Fixed Effects role ships (the token is
# not in the Role dropdown), which keeps the shipped workbook in the
# DW-active / BFN-token state. Same TAKE-trimmed idiom as the responses and
# sequence-flag counts.
_FIXED_EFFECTS_COUNT_FORMULA = (
    "SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))"
    f'="{_ROLE_FIXED_EFFECTS}"))'
)

# ── Sequence Spacing block (base-period release) ────────────────────────────
# A small fixed block under the spec rows: the Base Period Δ candidate and
# the Δ in use (A/B), four verdict lines (A, blank while quiet — the H2
# status-line pattern), and the delta spectrum spilling at J/K. These rows
# sit INSIDE the 16000-row Spec_* bands (columns B–I), which is safe because
# every consumer TAKE-trims to COLUMNS(Source_Data); the writer additionally
# scrubs the spec dropdown validations off the block's rows.
_ROW_SPACING_HEADING = _LAST_DATA_ROW + 2      # 28
_ROW_SPACING_CANDIDATE = _ROW_SPACING_HEADING + 1   # 29
_ROW_SPACING_IN_USE = _ROW_SPACING_HEADING + 2      # 30
_ROW_SPACING_REGULARITY = _ROW_SPACING_HEADING + 3  # 31
_ROW_SPACING_OFF_GRID = _ROW_SPACING_HEADING + 4    # 32
_ROW_SPACING_NO_NATURAL = _ROW_SPACING_HEADING + 5  # 33
_ROW_SPACING_CALENDAR = _ROW_SPACING_HEADING + 6    # 34
_ROW_SPACING_LAST = _ROW_SPACING_CALENDAR
# Spectrum display: headers on the candidate row, the two-column
# Sequence_Delta_Spectrum() spill directly below (deltas are data-dependent,
# so the spill grows downward past the fixed verdict rows when needed).
_C_SPECTRUM_DELTA = _C_LEVELS       # J
_C_SPECTRUM_COUNT = _C_REF_IN_USE   # K

# Verdict messages. Blank cell = quiet; conditional formatting keys on
# nonblank (yellow = advisory, red = the data is off the declared grid).
_MSG_REGULARITY = (
    "Sequence is not evenly spaced: spacings besides Δ exist — "
    "Difference_By / Lag_By return #N/A at gap rows."
)
_MSG_OFF_GRID = (
    "Off-grid spacings: some gaps are not whole multiples of Δ — "
    "those observations never align with the Δ grid."
)
_MSG_NO_NATURAL = (
    "No repeated spacing — no natural base period (candidate falls back "
    "to the minimum). Review the spectrum and set Δ in spec column I."
)
_MSG_CALENDAR = (
    "Calendar-day spacing (~28–31 / 90–92 / 365–366): do not "
    "set a day-count Δ — build an integer period index upstream "
    "(YEAR, or YEAR*12+MONTH) and flag that as the Sequence axis."
)


def _set_sheet_scoped_names(
    sheet: xw.Sheet, closures: Sequence[CatalogFunction]
) -> None:
    """Register this sheet's local names in dependency order.

    Two groups, installed in order so Excel resolves each reference against
    the names already added:

    1. **Wiring** (defined here) — the ``Source_Table`` dataset-retarget
       point (the only name that references the table; ``Source_Data`` and
       ``Header_Names`` derive from it) and the ``Spec_*`` column ranges.
       These hardcode *this sheet's* cell addresses and the source table, so
       they live with the sheet layout rather than in the portable catalog.
    2. **Constructor closures** (``closures``) — the zero-arg LAMBDAs
       ``Sample_Include``/``Response_Column``/``Row_Labels``/``X_s``/
       ``Constructed_Column_Names``, sourced from ``lambda_functions.json``
       (scope ``"Regression"``) so their definitions live in one declarative
       place and appear on the LAMBDA_functions catalog sheet. Passed in
       document order, which is dependency order (``Sample_Include`` before
       ``X_s``, etc.).
    """
    sname = f"'{sheet.name}'"

    local_names: dict[str, str] = {
        # ── Source-table indirection: THE dataset-retarget point ─────────
        # Source_Table is the only name that references the table; the data
        # body and header row derive from it, so a dataset changeover is a
        # ONE-name edit. DROP/TAKE (not OFFSET) keep the derivations
        # non-volatile — a volatile Header_Names would be re-evaluated on
        # every Data Table substitution pass during workbook calculation.
        "Source_Table": "=LifeExpectancyData[#All]",
        "Source_Data": "=DROP(Source_Table,1)",
        "Header_Names": "=TAKE(Source_Table,1)",
        # ── Spec ranges (local columns B–I; TAKE-trimmed at use) ─────────
        "Spec_Role": f"={sname}!$B${_FIRST_DATA_ROW}:$B${_VALIDATION_LAST_ROW}",
        "Spec_Include": f"={sname}!$C${_FIRST_DATA_ROW}:$C${_VALIDATION_LAST_ROW}",
        "Spec_Type": f"={sname}!$D${_FIRST_DATA_ROW}:$D${_VALIDATION_LAST_ROW}",
        "Spec_Reference": f"={sname}!$E${_FIRST_DATA_ROW}:$E${_VALIDATION_LAST_ROW}",
        # Reserved axes: named now so the grid shape is final, read by
        # nothing until the Order/Transform release.
        "Spec_Order": f"={sname}!$F${_FIRST_DATA_ROW}:$F${_VALIDATION_LAST_ROW}",
        "Spec_Transform": f"={sname}!$G${_FIRST_DATA_ROW}:$G${_VALIDATION_LAST_ROW}",
        # Sequence structural axis (live: read by the zero-or-one status
        # validation and its conditional formats, not by any constructor);
        # Base Period Δ is its reserved companion, read by nothing until
        # the base-period release.
        "Spec_Sequence": f"={sname}!$H${_FIRST_DATA_ROW}:$H${_VALIDATION_LAST_ROW}",
        "Spec_Base_Period_Delta": (
            f"={sname}!$I${_FIRST_DATA_ROW}:$I${_VALIDATION_LAST_ROW}"
        ),
        # Model-level Intercept toggle (row-2 control): a single boolean cell
        # in the C/Include column. No v3.0 formula reads it yet — the engine
        # will, exactly as the v1 Regression sheet's Allow_Intercept did.
        "Allow_Intercept": (
            f"={sname}!${col_letter(_C_INCLUDE)}${_INTERCEPT_ROW}"
        ),
    }

    # Wiring first: Excel resolves each name against the ones already added,
    # and the constructor closures below reference Source_Data / Spec_*.
    for name, refers_to in local_names.items():
        drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)

    # Constructor closures (Sample_Include, Response_Column, Row_Labels, X_s,
    # Constructed_Column_Names) come from lambda_functions.json, in document
    # (= dependency) order. Their full rationale — the REDUCE-product mask, the
    # (col+0) Filter coercion that avoids N()'s implicit intersection, the
    # once-bound Dummy_Levels skip, the X_s / Constructed_Column_Names twin —
    # lives in each entry's catalog description. Compact the display formula to
    # a single line for the defined name's RefersTo (no-op on already-compact
    # strings; safe if the JSON is later pretty-printed).
    for closure in closures:
        refers_to = "=" + _strip_non_string_whitespace(
            _normalize_user_formula(closure.formula_display)
        )
        drop_local_name(sheet, closure.name)
        sheet.api.Names.Add(Name=closure.name, RefersTo=refers_to)


def _add_list_validation(sheet: xw.Sheet, col: int, formula: str) -> None:
    """Attach a dropdown list to a spec column's input band."""
    rng = sheet.range((_FIRST_DATA_ROW, col), (_VALIDATION_LAST_ROW, col)).api
    rng.Validation.Delete()
    rng.Validation.Add(
        Type=_XL_VALIDATE_LIST,
        AlertStyle=_XL_VALID_ALERT_STOP,
        Operator=_XL_BETWEEN,
        Formula1=formula,
    )
    rng.Validation.IgnoreBlank = True


def _set_note(sheet: xw.Sheet, row: int, col: int, text: str) -> None:
    """Replace the cell's note/comment text."""
    cell_api = sheet.range(rc(row, col)).api
    try:
        cell_api.ClearComments()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    cell_api.AddComment(text)
    cell_api.Comment.Visible = False


def _write_spec_block(sheet: xw.Sheet) -> None:
    """The A–K specification block: headers, defaults, dropdowns, CF."""
    bold_row(sheet, _HEADER_ROW, _C_LABEL, _C_REF_IN_USE)
    for col, header in (
        (_C_LABEL, "Variable"),
        (_C_ROLE, "Role"),
        (_C_INCLUDE, "Include"),
        (_C_TYPE, "Type"),
        (_C_REFERENCE, "Reference Level"),
        (_C_ORDER, "Order"),
        (_C_TRANSFORM, "Transform"),
        (_C_SEQUENCE, "Sequence"),
        (_C_BASE_PERIOD, "Base Period Δ"),
        (_C_LEVELS, "Levels"),
        (_C_REF_IN_USE, "Reference In Use"),
    ):
        val(sheet, _HEADER_ROW, col, header)

    # A: variable names spill straight from the table's header row via the
    # Header_Names indirection (dataset-agnostic; reads no other sheet).
    f(sheet, _FIRST_DATA_ROW, _C_LABEL, "=TRANSPOSE(Header_Names)")

    for offset, variable in enumerate(_VARIABLES):
        row = _FIRST_DATA_ROW + offset
        role, include, ptype = _DEFAULT_SPEC.get(variable, _FALLBACK_SPEC)
        val(sheet, row, _C_ROLE, role)
        format_input(sheet, row, _C_ROLE)
        val(sheet, row, _C_INCLUDE, include)
        format_input(sheet, row, _C_INCLUDE)
        val(sheet, row, _C_TYPE, ptype)
        format_input(sheet, row, _C_TYPE)
        # E starts blank → first-in-sort-order default reference.
        format_input(sheet, row, _C_REFERENCE)
        # F (Order) starts blank; G (Transform) defaults to "None".
        format_input(sheet, row, _C_ORDER)
        val(sheet, row, _C_TRANSFORM, _DEFAULT_TRANSFORM)
        format_input(sheet, row, _C_TRANSFORM)
        # H (Sequence): TRUE on the shipped ordering axis (Year), blank
        # elsewhere — zero-or-one flags is the legal range, and blank stays a
        # valid non-panel spec. I (Base Period Δ) is computed-with-override:
        # pre-filled with the candidate formula — the MODE of within-group
        # spacings on the flagged row, blank elsewhere — and styled as an
        # input so a typed number overrides the candidate (the reference-
        # level pattern). Base_Period_Delta() reads whatever this cell
        # shows; the scalar $H test keeps the candidate lazy (it computes
        # only on the flagged row).
        if variable in _DEFAULT_SEQUENCE_VARIABLES:
            val(sheet, row, _C_SEQUENCE, True)
        format_input(sheet, row, _C_SEQUENCE)
        f(
            sheet,
            row,
            _C_BASE_PERIOD,
            (
                f'=IF($H{row}<>TRUE,"",'
                'IFERROR(Base_Period_Delta_Candidate(),""))'
            ),
        )
        format_input(sheet, row, _C_BASE_PERIOD)

        # J: Levels display — Categorical Predictors only; the raw distinct
        # level count L over the mask-included rows, with Dummy_Levels'
        # blank normalization mirrored inline. Deliberately NOT a
        # Dummy_Levels call: the display must show L (including 1 for a
        # degenerate column, feeding the red CF below), while Dummy_Levels
        # returns the L−1 retained levels and #N/A when degenerate.
        # ROW()-2 maps sheet row 3 → Source_Data column 1. IFERROR → 0
        # covers the empty-masked-sample edge.
        f(
            sheet,
            row,
            _C_LEVELS,
            (
                f'=IF(OR($B{row}<>"{_ROLE_PREDICTOR}",$D{row}<>"Categorical"),"",'
                f"LET(col,INDEX(Source_Data,0,ROW()-{_ROW_TO_COL_OFFSET}),"
                f'x,IF(col="","",col),'
                f'IFERROR(ROWS(UNIQUE(FILTER(x,(x<>"")*Sample_Include()))),0)))'
            ),
        )

        # K: Reference In Use display — the level the constructor will
        # actually drop, surfaced even when defaulted. A nonblank E is
        # echoed verbatim (E's invalid-reference CF carries the error
        # signal); a blank E shows Dummy_Levels' own default, the first
        # sorted level over the mask-included sample, with the same blank
        # normalization mirrored inline. Deliberately NOT a Dummy_Levels
        # call: the function returns the RETAINED levels, which is exactly
        # the set the reference has been dropped from. IFERROR → "" covers
        # the empty-masked-sample edge (J shows 0 and flags red there).
        f(
            sheet,
            row,
            _C_REF_IN_USE,
            (
                f'=IF(OR($B{row}<>"{_ROLE_PREDICTOR}",$D{row}<>"Categorical"),"",'
                f'IF($E{row}<>"",$E{row},'
                f"LET(col,INDEX(Source_Data,0,ROW()-{_ROW_TO_COL_OFFSET}),"
                f'x,IF(col="","",col),'
                f'IFERROR(INDEX(SORT(UNIQUE(FILTER(x,(x<>"")*Sample_Include()))),1,1),""))))'
            ),
        )

    _add_list_validation(sheet, _C_ROLE, _ROLE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_INCLUDE, _INCLUDE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_TYPE, _TYPE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_TRANSFORM, _TRANSFORM_VALIDATION_LIST)
    _add_list_validation(sheet, _C_SEQUENCE, _SEQUENCE_VALIDATION_LIST)

    # Cascading relevance, Role-keyed: the per-Predictor inputs (C–G) and
    # the Categorical displays (J–K) gray out whenever Role ≠ Predictor —
    # the Reference-only-for-Categorical pattern applied one level up.
    # H–I are deliberately excluded: Sequence is a structural axis, not a
    # Role property (an Identifier like Year is a typical sequence axis).
    for role_keyed_band in (
        f"$C${_FIRST_DATA_ROW}:$G${_LAST_DATA_ROW}",
        f"$J${_FIRST_DATA_ROW}:$K${_LAST_DATA_ROW}",
    ):
        add_expression_format(
            sheet,
            role_keyed_band,
            f'=$B{_FIRST_DATA_ROW}<>"{_ROLE_PREDICTOR}"',
            font_color=MUTED_TEXT_COLOR,
        )

    # Cascading relevance, Sequence-keyed: H–I gray out on every row that
    # is not the sequence axis — Base Period Δ is meaningful only for the
    # flagged row, and the flag itself keys on its own value, not on Role.
    add_expression_format(
        sheet,
        f"$H${_FIRST_DATA_ROW}:$I${_LAST_DATA_ROW}",
        f"=$H{_FIRST_DATA_ROW}<>TRUE",
        font_color=MUTED_TEXT_COLOR,
    )

    # Multi-flag error: red on every flagged Sequence cell when two-plus
    # rows are marked — points at the offending rows while the H2 status
    # line states the error.
    add_expression_format(
        sheet,
        f"$H${_FIRST_DATA_ROW}:$H${_LAST_DATA_ROW}",
        (
            f"=AND($H{_FIRST_DATA_ROW}=TRUE,"
            f"{_SEQUENCE_FLAG_COUNT_FORMULA}>1)"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Degeneracy flag: red J when an INCLUDED Categorical Predictor has
    # L <= 1 — the constructor contributes zero columns for it (visible
    # degradation, not silent omission). N() coerces "" to 0.
    add_expression_format(
        sheet,
        f"$J${_FIRST_DATA_ROW}:$J${_LAST_DATA_ROW}",
        (
            f'=AND($B{_FIRST_DATA_ROW}="{_ROLE_PREDICTOR}",'
            f"$C{_FIRST_DATA_ROW}=TRUE,"
            f'$D{_FIRST_DATA_ROW}="Categorical",'
            f"N($J{_FIRST_DATA_ROW})<=1)"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Invalid-reference flag: red E when a nonblank reference makes
    # Dummy_Levels fail — the constructor's exact skip condition, tested
    # directly rather than via level-set membership (a membership test
    # against the returned set would false-positive on the default
    # reference itself, which Dummy_Levels excludes from its output).
    add_expression_format(
        sheet,
        f"$E${_FIRST_DATA_ROW}:$E${_LAST_DATA_ROW}",
        (
            f'=AND($E{_FIRST_DATA_ROW}<>"",'
            f"ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-{_ROW_TO_COL_OFFSET}),"
            f"$E{_FIRST_DATA_ROW},Sample_Include())))"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    _write_sequence_status(sheet)


def _write_sequence_status(sheet: xw.Sheet) -> None:
    """Row-2 status line above the Sequence column (the H2 cell).

    Zero-or-one Sequence flags is the legal range: zero is a valid spec
    (non-panel data), one designates the ordering axis, two-plus is a spec
    error. Same pattern as the exactly-one-Response audit: a visible
    status cell that renders blank while the spec is legal and a red error
    line when it is not (the message overflows rightward across the empty
    row-2 display-column cells). Per-cell red CF on the flagged H cells
    (added in _write_spec_block) points at the offending rows.
    """
    status_cell = f"${col_letter(_C_SEQUENCE)}${_INTERCEPT_ROW}"  # $H$2
    f(
        sheet,
        _INTERCEPT_ROW,
        _C_SEQUENCE,
        (
            f"=IF({_SEQUENCE_FLAG_COUNT_FORMULA}>1,"
            '"ERROR: multiple Sequence flags (mark at most one variable)","")'
        ),
    )
    bold(sheet, _INTERCEPT_ROW, _C_SEQUENCE)
    add_expression_format(
        sheet,
        status_cell,
        f'={status_cell}<>""',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_sequence_spacing_block(sheet: xw.Sheet) -> None:
    """The Sequence Spacing block under the spec rows (base-period release).

    Fixed rows 28–34, columns A–K: Δ candidate and Δ in use on the left,
    the delta spectrum (distinct within-group spacings with counts)
    spilling at J/K, and four verdict lines in A that render blank while
    quiet — the H2-status-line pattern. Verdicts key on the Δ-in-use cell
    (candidate or typed override), so an override re-runs every check
    against the user's Δ.

    The block's rows sit inside the 16000-row Spec_* bands; every
    base-period consumer TAKE-trims to COLUMNS(Source_Data), and the spec
    dropdown validations are scrubbed off these rows so no spec dropdown
    haunts a display cell.
    """
    section_heading(
        sheet, _ROW_SPACING_HEADING, _C_LABEL, "Sequence Spacing (Base Period Δ)"
    )

    in_use = f"$B${_ROW_SPACING_IN_USE}"
    candidate = f"$B${_ROW_SPACING_CANDIDATE}"

    val(sheet, _ROW_SPACING_CANDIDATE, _C_LABEL, "Δ candidate")
    bold(sheet, _ROW_SPACING_CANDIDATE, _C_LABEL)
    f(
        sheet,
        _ROW_SPACING_CANDIDATE,
        _C_ROLE,
        '=IFERROR(Base_Period_Delta_Candidate(),"")',
    )

    # Δ in use: the flagged row's I cell — candidate formula or typed
    # override. Local Spec_* reads (not the workbook Base_Period_Delta()
    # accessor) so a standalone Model Construction rebuild reads its own
    # spec. Blank when no Sequence axis is flagged; "(not set)" when an
    # axis is flagged but its I cell holds no number.
    val(sheet, _ROW_SPACING_IN_USE, _C_LABEL, "Δ in use")
    bold(sheet, _ROW_SPACING_IN_USE, _C_LABEL)
    f(
        sheet,
        _ROW_SPACING_IN_USE,
        _C_ROLE,
        (
            "=LET(n_c,COLUMNS(Source_Data),"
            "p,IFERROR(XMATCH(TRUE,TAKE(Spec_Sequence,n_c),0),0),"
            'IF(p=0,"",'
            "LET(v,INDEX(TAKE(Spec_Base_Period_Delta,n_c),p),"
            'IF(ISNUMBER(v),v,"(not set)"))))'
        ),
    )
    # Visible override: yellow when the Δ in use is a number that differs
    # from the computed candidate (the reference-level pattern's "surfaced
    # even when defaulted", inverted: surfaced ESPECIALLY when not default).
    add_expression_format(
        sheet,
        f"{in_use}",
        f"=AND(ISNUMBER({in_use}),ISNUMBER({candidate}),{in_use}<>{candidate})",
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # Delta spectrum: headers + the two-column Sequence_Delta_Spectrum()
    # spill. IFERROR degrades the no-axis / no-spacings #N/A to a quiet
    # blank.
    val(sheet, _ROW_SPACING_CANDIDATE, _C_SPECTRUM_DELTA, "Delta")
    val(sheet, _ROW_SPACING_CANDIDATE, _C_SPECTRUM_COUNT, "Count")
    bold_row(
        sheet, _ROW_SPACING_CANDIDATE, _C_SPECTRUM_DELTA, _C_SPECTRUM_COUNT
    )
    f(
        sheet,
        _ROW_SPACING_IN_USE,
        _C_SPECTRUM_DELTA,
        '=IFERROR(Sequence_Delta_Spectrum(),"")',
    )

    # Verdict lines: blank while quiet, message text when firing. Long
    # messages overflow rightward across the empty B–I display cells.
    verdicts: tuple[tuple[int, str, tuple[int, int, int], tuple[int, int, int]], ...] = (
        (
            _ROW_SPACING_REGULARITY,
            (
                f"=LET(d,Sequence_Deltas(),v,{in_use},"
                'IF(OR(COUNT(d)=0,NOT(ISNUMBER(v))),"",'
                f'IF(SUM(--(d<>v))>0,"{_MSG_REGULARITY}","")))'
            ),
            CF_YELLOW_FILL,
            CF_DARK_YELLOW_TEXT,
        ),
        (
            _ROW_SPACING_OFF_GRID,
            (
                f"=LET(d,Sequence_Deltas(),v,{in_use},"
                'IF(OR(COUNT(d)=0,NOT(ISNUMBER(v)),N(v)<=0),"",'
                f'IF(SUM(--(MOD(d,v)<>0))>0,"{_MSG_OFF_GRID}","")))'
            ),
            CF_LIGHT_RED_FILL,
            CF_DARK_RED_TEXT,
        ),
        (
            _ROW_SPACING_NO_NATURAL,
            (
                "=LET(d,Sequence_Deltas(),"
                'IF(COUNT(d)=0,"",'
                f'IF(ISNA(MODE.SNGL(d)),"{_MSG_NO_NATURAL}","")))'
            ),
            CF_YELLOW_FILL,
            CF_DARK_YELLOW_TEXT,
        ),
        (
            _ROW_SPACING_CALENDAR,
            (
                "=LET(d,Sequence_Deltas(),n,COUNT(d),"
                "cal,SUM(--((((d>=28)*(d<=31))+((d>=90)*(d<=92))"
                "+((d>=365)*(d<=366)))>0)),"
                'IF(n=0,"",'
                f'IF(cal*2>n,"{_MSG_CALENDAR}","")))'
            ),
            CF_LIGHT_RED_FILL,
            CF_DARK_RED_TEXT,
        ),
    )
    for row, formula, fill, font_color in verdicts:
        f(sheet, row, _C_LABEL, formula)
        cell = f"$A${row}"
        add_expression_format(
            sheet,
            cell,
            f'={cell}<>""',
            fill=fill,
            font_color=font_color,
        )

    # The block's rows sit inside the spec columns' 16000-row dropdown
    # validations — scrub them off the display band so no spec dropdown
    # haunts these cells.
    sheet.range(
        (_ROW_SPACING_HEADING, _C_ROLE), (_ROW_SPACING_LAST, _C_BASE_PERIOD)
    ).api.Validation.Delete()


def _write_intercept_control(sheet: xw.Sheet) -> None:
    """Row-2 model-level Intercept toggle (the ``Allow_Intercept`` cell).

    Mirrors the v1 Regression sheet's A2 label / boolean-column toggle, here
    aligned to column C so the toggle sits at the top of the Include column,
    one row above the per-variable Include toggles. No v3.0 formula consumes
    it yet — it restores the visible control and declares the intercept with
    the rest of the spec for the future engine to read.

    Conditional formatting encodes the reference-coding coupling
    (ROADMAP: "Intercept coupling — flag, don't switch"): treatment coding
    drops one level and relies on the intercept to carry the baseline, so an
    included Categorical predictor makes the intercept effectively required.

    * **Gray** whenever an included Categorical predictor is present — the
      toggle is required-here and reads as locked-on even while (correctly)
      TRUE.
    * **Red** when the toggle is nonetheless set FALSE in that state — the
      invalid combination, flagged not forced. Added first with StopIfTrue so
      it outranks the gray rule on the same cell.
    """
    val(sheet, _INTERCEPT_ROW, _C_LABEL, "Intercept")
    bold(sheet, _INTERCEPT_ROW, _C_LABEL)
    val(sheet, _INTERCEPT_ROW, _C_INCLUDE, True)
    format_input(sheet, _INTERCEPT_ROW, _C_INCLUDE)

    # TRUE/FALSE dropdown on the single toggle cell (the spec block's Include
    # validation covers only rows _FIRST_DATA_ROW onward, not this row).
    cell = sheet.range(
        rc(_INTERCEPT_ROW, _C_INCLUDE), rc(_INTERCEPT_ROW, _C_INCLUDE)
    ).api
    cell.Validation.Delete()
    cell.Validation.Add(
        Type=_XL_VALIDATE_LIST,
        AlertStyle=_XL_VALID_ALERT_STOP,
        Operator=_XL_BETWEEN,
        Formula1=_INCLUDE_VALIDATION_LIST,
    )
    cell.Validation.IgnoreBlank = True

    toggle = f"${col_letter(_C_INCLUDE)}${_INTERCEPT_ROW}"  # $C$2
    # At least one included Categorical predictor anywhere in the spec.
    cat_included = (
        "SUMPRODUCT("
        f'N(TAKE(Spec_Role,COLUMNS(Source_Data))="{_ROLE_PREDICTOR}"),'
        "N(TAKE(Spec_Include,COLUMNS(Source_Data))=TRUE),"
        'N(TAKE(Spec_Type,COLUMNS(Source_Data))="Categorical"))>0'
    )
    # Red first (StopIfTrue → outranks gray on this cell): FALSE while a
    # Categorical needs the intercept.
    add_expression_format(
        sheet,
        toggle,
        f"=AND({toggle}=FALSE,{cat_included})",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
        stop_if_true=True,
    )
    # Gray: required-here signal, applies even while the toggle is still TRUE.
    add_expression_format(
        sheet,
        toggle,
        f"={cat_included}",
        font_color=MUTED_TEXT_COLOR,
    )


def _write_row_zones(sheet: xw.Sheet) -> None:
    """The M/N derived-row zone: full-height label and mask spills.

    Row 1 of M/N is not written here — _write_audit_row owns the audit
    strip that occupies it.
    """
    sheet.range(rc(1, _C_GAP)).column_width = _GAP_COLUMN_WIDTH

    bold_row(sheet, _HEADER_ROW, _C_ROW_LABELS, _C_INCLUDED)
    val(sheet, _HEADER_ROW, _C_ROW_LABELS, "Row Labels")
    val(sheet, _HEADER_ROW, _C_INCLUDED, "Included")

    f(sheet, _FIRST_DATA_ROW, _C_ROW_LABELS, "=Row_Labels()")
    f(sheet, _FIRST_DATA_ROW, _C_INCLUDED, "=Sample_Include()")


def _write_audit_row(sheet: xw.Sheet) -> None:
    """Row-1 audit strip: bold label/value pairs from column K rightward.

    Values live in their own cells (not concatenated into the labels) so
    the QC analyzer can assert the numbers directly. The X_s()-derived
    cells wrap IFERROR — an empty model makes DROP(built,,1) error, and
    the audit strip must degrade to the documented string, never leak a
    raw #CALC!. The two SUMPRODUCT counts are total functions over
    full-height inputs and cannot error, so they stay unwrapped.
    """
    audit_cells: tuple[tuple[str, str], ...] = (
        ("k", f"=IFERROR(COLUMNS(X_s()),{_EMPTY_MODEL_FALLBACK})"),
        ("rows", f"=IFERROR(ROWS(X_s()),{_EMPTY_MODEL_FALLBACK})"),
        ("response", f"={_RESPONSE_NAME_FORMULA}"),
        (
            "responses",
            "=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))"
            f'="{_ROLE_RESPONSE}"))',
        ),
        ("included rows", "=SUMPRODUCT(N(Sample_Include()))"),
        ("sequence flags", f"={_SEQUENCE_FLAG_COUNT_FORMULA}"),
    )
    for (label_col, value_col), (label, formula) in zip(
        _AUDIT_PAIRS, audit_cells
    ):
        val(sheet, _AUDIT_ROW, label_col, label)
        f(sheet, _AUDIT_ROW, value_col, formula)
        bold_row(sheet, _AUDIT_ROW, label_col, value_col)

    # The model must declare exactly one Response — flag the count red
    # otherwise (zero and multiple are both spec errors the filtered
    # zones can only partially absorb).
    responses_col = col_letter(_AUDIT_PAIRS[3][1])
    add_expression_format(
        sheet,
        f"${responses_col}${_AUDIT_ROW}",
        f"=N(${responses_col}${_AUDIT_ROW})<>1",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Zero-or-one Sequence flags — red only at two-plus (zero is a valid
    # non-panel spec, so <>1 would be the wrong test here).
    sequence_col = col_letter(_AUDIT_PAIRS[5][1])
    add_expression_format(
        sheet,
        f"${sequence_col}${_AUDIT_ROW}",
        f"=N(${sequence_col}${_AUDIT_ROW})>1",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_filtered_zones(sheet: xw.Sheet) -> None:
    """The P/Q and S/T→ filtered display zones.

    The only row-filtering on the sheet: FILTER(<full-height name>(),
    Sample_Include()). Every spill wraps IFERROR(..., "(empty model)") —
    an empty model (no included predictors, or a mask that excludes
    everything) degrades to the documented string.
    """
    for break_col in (_C_BREAK_LEFT, _C_BREAK_MID):
        sheet.range(rc(1, break_col)).column_width = _GAP_COLUMN_WIDTH

    bold_row(sheet, _HEADER_ROW, _C_FILTERED_LABELS, _C_MATRIX_START)
    val(sheet, _HEADER_ROW, _C_FILTERED_LABELS, "Row Labels")
    # Q header carries the derived response name ("y: Life expectancy")
    # so the filtered-y column is self-describing under response swaps.
    f(
        sheet,
        _HEADER_ROW,
        _C_FILTERED_Y,
        f'="y: "&{_RESPONSE_NAME_FORMULA}',
    )
    val(sheet, _HEADER_ROW, _C_MATRIX_LABELS, "Row Labels")
    # Header strip above the matrix: the structural twin guarantees this
    # spills exactly COLUMNS(X_s()) level-qualified names.
    f(
        sheet,
        _HEADER_ROW,
        _C_MATRIX_START,
        f"=IFERROR(Constructed_Column_Names(),{_EMPTY_MODEL_FALLBACK})",
    )

    filtered_spills: tuple[tuple[int, str], ...] = (
        (_C_FILTERED_LABELS, "Row_Labels()"),
        (_C_FILTERED_Y, "Response_Column()"),
        (_C_MATRIX_LABELS, "Row_Labels()"),
        (_C_MATRIX_START, "X_s()"),
    )
    for col, source in filtered_spills:
        f(
            sheet,
            _FIRST_DATA_ROW,
            col,
            (
                f"=IFERROR(FILTER({source},Sample_Include()),"
                f"{_EMPTY_MODEL_FALLBACK})"
            ),
        )


def write_model_construction_sheet(
    workbook: xw.Book,
    closures: Sequence[CatalogFunction] | None = None,
) -> xw.Sheet:
    """Create or rebuild the Model Construction sheet.

    Parameters
    ----------
    workbook : xw.Book
        The open workbook to write into.
    closures : Sequence[CatalogFunction] or None, optional
        The sheet-scoped constructor functions (scope ``"Model Construction"``),
        in dependency order. When None, they are loaded from
        ``lambda_functions.json`` — so a standalone rebuild works without the
        caller threading the catalog through.
    """
    if closures is None:
        closures = load_catalog_document(_DEFINITIONS_PATH).functions_for_sheet(
            _CLOSURE_SCOPE
        )

    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    section_heading(sheet, 1, _C_LABEL, "Model Construction")

    _set_sheet_scoped_names(sheet, closures)
    _write_spec_block(sheet)
    _write_sequence_spacing_block(sheet)
    _write_intercept_control(sheet)
    _write_row_zones(sheet)
    _write_audit_row(sheet)
    _write_filtered_zones(sheet)

    # Reserved-column and Sequence notes are COM comment calls; keep them
    # out of the RecordingSheet-testable spec block.
    _set_note(sheet, _FIRST_DATA_ROW, _C_ORDER, _RESERVED_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_TRANSFORM, _RESERVED_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_SEQUENCE, _SEQUENCE_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_BASE_PERIOD, _BASE_PERIOD_NOTE)

    sheet.range(
        (_HEADER_ROW, _C_LABEL), (_HEADER_ROW, _C_REF_IN_USE)
    ).columns.autofit()
    return sheet


def main(workbook_path: str | Path = "Lambda_Library.xlsx") -> None:
    """Standalone runner: add/rebuild the sheet in the given workbook."""
    path = Path(workbook_path).resolve()
    app = xw.App(visible=True, add_book=False)
    try:
        workbook, existed = open_or_create_workbook(app, path)
        write_model_construction_sheet(workbook)
        workbook.save(str(path))
        state = "existing" if existed else "new"
        print(f"Sheet written: {SHEET_NAME} ({state} workbook: {path})")
    finally:
        app.quit()


if __name__ == "__main__":
    import sys

    main(sys.argv[1] if len(sys.argv) > 1 else "Lambda_Library.xlsx")

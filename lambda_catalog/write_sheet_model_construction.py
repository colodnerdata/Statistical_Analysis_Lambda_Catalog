"""Build the Model Construction worksheet — v2.0 declarative specification block.

Two-axis specification plus the Sequence structural axis (ROADMAP: v2.0 —
Specification-Driven Regression; Sequence added post-v2.0):

Note on numbering: this changeover was planned as "v3.0" and renumbered to
v2.0 before release. Other modules and tests still carry the old label in
comments; v3.0 now means the engine-interface release (ROADMAP). See
TODOs.md for the cleanup item.

    A        B      C       D     E               F       G         H        I              J              K      L
    Variable Role   Include Type  Reference Level Order   Transform Sequence Sequence Period Period In Use Levels Reference In Use
    (spill)  (drop) (input) (drop)(input)         (rsvd.) (rsvd.)   (flag)   (cand./ovr.)   (disp.)        (disp.)(disp.)

Right of the spec block, after a narrow gap column O (which also visually
reserves the future Design Columns audit column):

    O (gap)   P             Q        R     S           T            U     V           W →
              Row Labels    Included (brk) Filt.Labels Filt.y       (brk) Filt.Labels Filtered X_s
    (=Row_Labels() spill at P4; =Sample_Include() spill at Q4 — both
     full-height, never internally filtered. S/T/V/W are the FILTERED
     display zones: the only place on the sheet where Sample_Include()
     row-filters anything. V repeats the filtered labels so the matrix
     reads side-by-side without scrolling back to S.)

Row 1, from column O rightward, holds the bold audit cells as
label/value pairs (values on the non-narrow columns P/S/V/W/Y/AA):

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
this control row the spec headers move to row 3 and the variable rows to
4–(4+N−1), where N is the source table's column count; the row-1 audit
strip is unaffected.

The spec spans EVERY column of the Source_Table-targeted table (currently
MileageData, 12 rows: [MPG]..[Model?] plus [Full_Data]). Two axes:

    Variable Role  — Response (y) | Predictor (x) | Identifier (Row Label) |
                     Filter | Omit | Fixed Effects
                     (what the column IS; future: Weight/Time.
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

Column F (Order) is reserved for a future release: it is styled as an
input and carries a sheet-scoped name (Spec_Order) so the grid shape is
final, but no formula reads it yet. Column G (Transform) went live at
v2.2 — its dropdown offers None (default; unchanged fit) or Log (natural
log). Log is read by Response_Column(), X_s(), Constructed_Column_Names(),
and Constructed_Column_Transforms() on the Response row and on Continuous
Predictor rows; it is disallowed (flagged red, not silently ignored) on
Categorical Predictors. See ARCHITECTURE.md §4/§5 for the full contract.
Column I
(Sequence Period) is the typed override cell on the Sequence-flagged row
(what the user sees typed; default empty); column J (Period In Use) is
Sequence's LIVE companion since the base-period release — computed-with-
override, the reference-level pattern: every row is pre-filled with a
formula that shows Base_Period_Delta_Candidate() (MODE of the within-group
consecutive spacings, MIN fallback when no spacing repeats) on the
Sequence-flagged row and blank elsewhere. Typing a number into I on the
flagged row overrides the candidate. The workbook-scoped Base_Period_Delta()
accessor reads column J and is the default [delta] of Lag_By /
Difference_By — Δ is never a silent 1. Override flagging lives ONLY on
the Sequence Spacing block (verdict lines on rows 31–34), never on the
spec block itself: the J cell stays plain so the spec block reads
top-to-bottom as a clean declaration, with the override verdict surfaced
where the user actually inspects the time grid.
The Sequence Spacing block below the spec (rows 28–34) surfaces the
candidate, the Δ in use (yellow when overridden), the delta spectrum
(distinct within-group spacings with counts, spilling at K/L), and four
verdict lines: Regularity (any spacing ≠ Δ), Off-grid (a spacing that is
not a whole multiple of Δ), the no-natural-base-period override prompt
(MODE undefined → MIN fallback in use), and the calendar-signature guidance
(spacings clustered at ~28–31 / ~90–92 / ~365–366 → recommend an integer
period index upstream instead of quantizing day counts to a scalar Δ).
The spectrum deliberately ignores Sample_Include(): the time grid is
dataset structure, not model-iteration state. No CONSTRUCTOR reads H, I, or
J — the base-period layer (candidate closure, spectrum, accessor) is the only
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

Default configuration (the human test plan's T0 state, retargeted to the
Mileage/Auto MPG dataset since Source_Table now defaults to MileageData):
    Car Name         → Identifier            (residual labeling; no columns)
    Model Year       → Predictor/Categorical/TRUE  (numeric-valued; Sequence axis)
    Origin           → Predictor/Categorical/TRUE  (numeric-valued: 1/2/3)
    MPG              → Response               (derived y)
    Horsepower, Weight → Predictor/Continuous/TRUE
    Cylinders, Displacement, Acceleration → Predictor/Continuous/FALSE (candidates)
    Make, Model?     → Omit                   (text columns parsed out of Car Name)
    Full_Data        → Omit                    (its all-features completeness
                                               flag is redundant with the mask's
                                               built-in completeness and
                                               over-filters; no default Filter)
Full-height contract: ROWS(X_s()) = ROWS(Row_Labels()) =
ROWS(Sample_Include()) = 406 always — the constructor reads the mask ONLY
to fix level sets; nothing here ever row-filters. With the real mask live,
the T0 mask-dependent values are real on the sheet: k = 16 (2 continuous +
2 Origin dummies + 12 Model Year dummies), and
SUMPRODUCT(N(Sample_Include())) = 392 (completeness-only on the response
and the two continuous predictors, no Full_Data over-filter).

Not here (deliberately, per release scoping): the QC analyzer
(analyze_model_construction.py) and the Version History / CHANGELOG bump
to v3.0 — those land in the final wiring PR.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

import xlwings as xw

from .catalog_schema import CatalogFunction, load_catalog_document
from .lambda_formula_parser import (
    _normalize_user_formula,
    _strip_non_string_whitespace,
)
from .regression_shared import FEATURE_COLUMNS as _LIFE_EXPECTANCY_FEATURE_COLUMNS
from .sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_DARK_YELLOW_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
    INPUT_COLOR,
    SUBHDR_COLOR,
)
from .workbook_helpers import (
    XL_SRC_RANGE,
    XL_YES,
    add_expression_format,
    bold,
    bold_row,
    col_letter,
    drop_local_name,
    f,
    f_structured,
    format_input,
    get_or_create_sheet,
    open_or_create_workbook,
    rc,
    reset_generated_sheet,
    section_heading,
    set_column_widths,
    val,
)
from .write_sheet_csv_dataset import LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS

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

# Every MileageData column, in table order (incl. the computed Full_Data
# completeness column — the spec spans the whole table).
_VARIABLES: list[str] = [
    "MPG",
    "Cylinders",
    "Displacement",
    "Horsepower",
    "Weight",
    "Acceleration",
    "Model Year",
    "Origin",
    "Car Name",
    "Make",
    "Model?",
    "Full_Data",
]
_N_VARIABLES = len(_VARIABLES)  # 12

# Row 2 is the model-level Intercept control (label A2, toggle C2 — aligned
# to the C/Include boolean column). The spec table sits one row below it:
# headers on row 3, the N variable rows from _FIRST_DATA_ROW to
# _LAST_DATA_ROW (N = len(_VARIABLES); currently 12, rows 4-15).
_INTERCEPT_ROW = 2
_HEADER_ROW = 3
_FIRST_DATA_ROW = 4
_LAST_DATA_ROW = _FIRST_DATA_ROW + _N_VARIABLES - 1  # 15
# Sheet row _FIRST_DATA_ROW maps to Source_Data column 1, so a row-indexed
# formula recovers its column via INDEX(Source_Data,0,ROW()-_ROW_TO_COL_OFFSET).
_ROW_TO_COL_OFFSET = _FIRST_DATA_ROW - 1  # 3

# Spec-block columns (1-based). Role precedes Include: the larger
# declaration comes first (dataset semantics before iteration state).
# F/G are the reserved Order/Transform slots; H is the Sequence structural
# flag with I/J as its Sequence Period / Period In Use pair (the
# reference-level pattern: I is the candidate-with-override input,
# J is the in-use display); K and L are the computed Categorical
# displays (Levels count, Reference In Use).
(
    _C_LABEL,
    _C_ROLE,
    _C_INCLUDE,
    _C_TYPE,
    _C_REFERENCE,
    _C_ORDER,
    _C_TRANSFORM,
    _C_SEQUENCE,
    _C_SEQUENCE_PERIOD,
    _C_PERIOD_IN_USE,
    _C_LEVELS,
    _C_REF_IN_USE,
) = range(1, 13)

# Widths for the shared A-L spec block — owned here (not by
# write_sheet_regression.py, which imports and calls _set_spec_block_column_widths)
# so the standalone build and the shared-block build can never drift.
# F (Order) is reserved-but-unwired, hence width 0 — visually collapsed
# until a future release wires it up. G (Transform) went live at v2.2
# (Log wiring); it now gets a real width, matching _C_TYPE's (both hold
# comparably-sized dropdown tokens).
_SPEC_COLUMN_WIDTHS: dict[int, float] = {
    _C_LABEL: 28,
    _C_ROLE: 20,
    _C_INCLUDE: 9,
    _C_TYPE: 11,
    _C_REFERENCE: 15,
    _C_ORDER: 0,
    _C_TRANSFORM: 11,
    _C_SEQUENCE: 10,
    _C_SEQUENCE_PERIOD: 14,
    _C_PERIOD_IN_USE: 14,
    _C_LEVELS: 7,
    _C_REF_IN_USE: 16,
}


def _set_spec_block_column_widths(sheet: xw.Sheet) -> None:
    set_column_widths(sheet, _SPEC_COLUMN_WIDTHS.items())


# Spec feedback zone (M, N, I — the verdict overlay): the delta spectrum
# (Sequence_Delta_Spectrum() spill at M2:N?) sits in M and N; the combined
# verdict switch lives at I2 (the Sequence_Period column's row-1/row-2
# cells are unused by the spec block, so the verdict overlays them
# without disturbing anything below row 3). Headers on row 1, content on
# row 2 — both sit INSIDE the spec block's zone (which extends from A:N,
# see the Regression sheet's _ZONES), so a single click on the spec
# outline collapses the spec and its feedback together.
_C_FEEDBACK_DELTA = 13     # M — Δ header / spectrum column 1
_C_FEEDBACK_COUNT = 14     # N — Count header / spectrum column 2

# Gap before the derived-row zone. One ungrouped column (width 2) so the
# spec outline and the derived-row outline collapse independently.
_C_GAP = 15

# Derived-row zone right of the spec block. P and Q hold the full-height
# Row_Labels() / Sample_Include() spills — they honor the full-height
# contract (never row-filtered); the FILTERED display zone is further right.
_C_ROW_LABELS = 16
_C_INCLUDED = 17
_GAP_COLUMN_WIDTH = 2

# Filtered display zone: the ONLY place Sample_Include() row-filters
# anything (everything left of Q honors the full-height contract). R and
# U are narrow visual breaks; V repeats the filtered labels so the matrix
# reads side-by-side without scrolling back to S.
_C_BREAK_LEFT = 18
_C_FILTERED_LABELS = 19
_C_FILTERED_Y = 20
_C_BREAK_MID = 21
_C_MATRIX_LABELS = 22
_C_MATRIX_START = 23

# Row-1 audit strip: label/value pairs marching right from column P,
# values placed on the non-narrow columns (P, S, V, W, Y, AA) so no number
# lands on a width-2 break column.
_AUDIT_ROW = 1
_AUDIT_PAIRS: tuple[tuple[int, int], ...] = (
    (_C_ROW_LABELS, _C_INCLUDED),          # k
    (_C_FILTERED_LABELS, _C_FILTERED_Y),   # rows
    (_C_MATRIX_LABELS, _C_MATRIX_START),   # response
    (_C_MATRIX_START + 1, _C_MATRIX_START + 2),  # responses (red CF <> 1)
    (_C_MATRIX_START + 3, _C_MATRIX_START + 4),  # included rows
    (_C_MATRIX_START + 5, _C_MATRIX_START + 6),  # sequence flags (red CF > 1)
    (_C_MATRIX_START + 7, _C_MATRIX_START + 8),  # fixed effects (red CF > 1)
    (_C_MATRIX_START + 9, _C_MATRIX_START + 10),  # FE absorbed df
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
# Unlike the other role tokens, "Omit" is never tested for by name — no
# constructor closure in lambda_functions.json and no check in the QC
# oracle (analyze_model_construction.py) compares Role against _ROLE_OMIT;
# they only test for the active roles (Response/Predictor/Filter). Omit is
# purely the implicit "none of the above" bucket, and a blank Role cell is
# indistinguishable from it in every formula and every QC check — both
# contribute no column and impose no mask condition. This is what makes it
# safe for a freshly-added spec row (e.g. one that just joined SpecTable
# via its native auto-extend when typed past the table's bottom edge) to
# sit with an unset Role: it is inert by construction, not merely
# "untested," so the user can classify new rows at their own pace instead
# of being forced to pick a Role the instant the row exists.
# The v2.1 panel role. Read by the Fixed_Effects_Column() accessor, the
# FE-count guard on the Regression diagnostics (BFN panel Durbin-Watson
# trigger matrix), and — since the phase 1-3 engine work — by
# Absorbed_Degrees_Of_Freedom() and the fit-time y_s()/X_s_Within() pair that
# the whole inference chain reads. Now in the Role dropdown: a spec claiming
# Fixed Effects actually gets the one-way within transformation and the
# absorbed-df correction, not silent pooled OLS.
_ROLE_FIXED_EFFECTS = "Fixed Effects"

# The derived response name, shared by the audit strip and the filtered-y
# header: the header of the first Role=Response spec row, "(none)" when
# no row carries the role. XMATCH position over the TAKE-trimmed roles is
# the same lookup Response_Column() uses for its data column. Wrapped in
# Ln(...) when that row's Transform is Log, so the audit strip never
# claims the model fits the raw column when it actually fits its log —
# same XMATCH position, so this reads Spec_Transform at the identical row
# Response_Column() transforms.
_RESPONSE_NAME_FORMULA = (
    "LET(n_c,COLUMNS(Source_Data),"
    f'p,XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,n_c)),'
    "h,INDEX(TOROW(Header_Names),p),"
    'IFERROR(IF(INDEX(TAKE(Spec_Transform,n_c),p)="Log","Ln("&h&")",h),"(none)"))'
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
# Base Period Δ companion. Make/Model? ship as explicit Omit (text columns
# parsed out of Car Name, not usable as numeric predictors) so the Omit role
# and its graying are demonstrated; Omit contributes no column and imposes no
# mask condition, leaving the fitted model identical to a plain excluded row.
#
# Full_Data ships as Omit, NOT Filter: the Full_Data completeness column
# demands EVERY continuous-measurement column be present, which is (a)
# redundant with the built-in completeness the mask already applies to the
# response and the model's included continuous predictors, and (b) an
# over-filter — it drops rows missing a sparse predictor the model does not
# even use. With no Filter declared, the shipped model includes every row
# complete on its OWN columns (392, vs fewer under Full_Data). The Filter
# role is exercised in the human test plan via a purpose-built filter
# column, not the completeness flag.
_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "MPG": (_ROLE_RESPONSE, False, "Continuous"),
    "Horsepower": (_ROLE_PREDICTOR, True, "Continuous"),
    "Weight": (_ROLE_PREDICTOR, True, "Continuous"),
    "Model Year": (_ROLE_PREDICTOR, True, "Categorical"),
    "Origin": (_ROLE_PREDICTOR, True, "Categorical"),
    "Car Name": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Make": (_ROLE_OMIT, False, "Continuous"),
    "Model?": (_ROLE_OMIT, False, "Continuous"),
    "Full_Data": (_ROLE_OMIT, False, "Continuous"),
}
_FALLBACK_SPEC: tuple[str, bool, str] = (_ROLE_PREDICTOR, False, "Continuous")

# Variables shipped with their Sequence flag (column H) set TRUE. Model Year
# is the canonical ordering axis for the Auto MPG panel: flagging it
# activates the Base Period Δ candidate (Δ = 1), the Sequence Spacing block,
# and the gated Durbin-Watson diagnostic on the Regression sheet. Structural
# and Role-independent — Model Year stays a Categorical Predictor, so the
# fitted model is unchanged; the flag only drives the serial-correlation /
# base-period layer. Kept to at most one entry: the H2 status line errors at
# two-plus flags.
_DEFAULT_SEQUENCE_VARIABLES: frozenset[str] = frozenset({"Model Year"})


# ── Per-dataset spec profiles ──────────────────────────────────────────────
# _VARIABLES/_DEFAULT_SPEC/_DEFAULT_SEQUENCE_VARIABLES above are the shipped
# Auto MPG defaults; SpecDatasetProfile wraps a dataset's variable list,
# default Role/Include/Type spec, and Sequence-flagged columns as one unit
# so retargeting Source_Table (the --regression-dataset CLI choice) can
# also retarget the spec block's defaults, instead of leaving every column
# of a newly-targeted dataset to _FALLBACK_SPEC's un-flagged Predictor.
# _write_spec_block sizes SpecTable to len(profile.variables) — a dataset
# with more or fewer columns than Auto MPG gets a table sized to match, so
# every column has a Spec_Role/Spec_Include/etc. entry from the first
# build, rather than depending on the user manually typing values past the
# table's edge until Excel's native AutoExpand catches up.
@dataclass(frozen=True)
class SpecDatasetProfile:
    """One dataset's Source_Table target and shipped spec-block defaults."""

    source_table_ref: str
    variables: tuple[str, ...]
    default_spec: dict[str, tuple[str, bool, str]]
    sequence_variables: frozenset[str] = frozenset()


_AUTO_MPG_PROFILE = SpecDatasetProfile(
    source_table_ref=f"={MILEAGE.table_name}[#All]",
    variables=tuple(_VARIABLES),
    default_spec=_DEFAULT_SPEC,
    sequence_variables=_DEFAULT_SEQUENCE_VARIABLES,
)

# Column order matches the Life Expectancy CSV's normalized header order
# plus the appended Full_Data column. Response/Predictor set mirrors
# regression_shared.FEATURE_COLUMNS, the same 18-predictor model
# analyze_life_expectancy.py validates against — Country is a text
# identifier (row labeling only), Year is the natural panel ordering axis
# (Sequence-flagged, Role Omit so it never enters the design matrix itself),
# Status is the one categorical predictor, and Full_Data ships Omit for the
# same reason Auto MPG's does: the built-in per-predictor completeness mask
# already covers it, so a Filter role would only over-filter.
_LIFE_EXPECTANCY_VARIABLES: tuple[str, ...] = (
    "Country",
    "Year",
    "Status",
    "Life expectancy",
    *_LIFE_EXPECTANCY_FEATURE_COLUMNS,
    "Full_Data",
)
_LIFE_EXPECTANCY_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Country": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Year": (_ROLE_OMIT, False, "Continuous"),
    "Status": (_ROLE_PREDICTOR, True, "Categorical"),
    "Life expectancy": (_ROLE_RESPONSE, False, "Continuous"),
    **{
        column: (_ROLE_PREDICTOR, True, "Continuous")
        for column in _LIFE_EXPECTANCY_FEATURE_COLUMNS
    },
    "Full_Data": (_ROLE_OMIT, False, "Continuous"),
}
_LIFE_EXPECTANCY_SEQUENCE_VARIABLES: frozenset[str] = frozenset({"Year"})
_LIFE_EXPECTANCY_PROFILE = SpecDatasetProfile(
    source_table_ref=f"={LIFE_EXPECTANCY.table_name}[#All]",
    variables=_LIFE_EXPECTANCY_VARIABLES,
    default_spec=_LIFE_EXPECTANCY_DEFAULT_SPEC,
    sequence_variables=_LIFE_EXPECTANCY_SEQUENCE_VARIABLES,
)

# Column order matches the Production Lots CSV's header order plus the
# appended Full_Data column — the same shape as
# analyze_regression_spec.py's _production_lots_fixed_effects_spec(), the
# QC-validated Crawford/Wright learning-curve model (ln(unit cost) = a +
# b*ln(cumulative units)), reused here verbatim so the shipped default
# matches a spec the test suite already proves fits correctly: Facility is
# the Fixed Effects panel-grouping column, Fiscal_Year is the Sequence
# axis, log Cum Units is the sole predictor, log Unit Cost is the
# response, and Full_Data is a Filter (unlike the other two datasets — this
# is the one shipped case that exercises Role=Filter by default).
_PRODUCTION_LOTS_VARIABLES: tuple[str, ...] = (
    "Lot_ID",
    "Facility",
    "Fiscal_Year",
    "Lot_Quantity",
    "Cumulative_Units",
    "Experience_Stock",
    "Unit_Cost_BY",
    "log Cum Units",
    "log experience",
    "log Unit Cost",
    "Full_Data",
)
_PRODUCTION_LOTS_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Lot_ID": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Facility": (_ROLE_FIXED_EFFECTS, False, "Continuous"),
    "Fiscal_Year": (_ROLE_OMIT, False, "Continuous"),
    "Lot_Quantity": (_ROLE_OMIT, False, "Continuous"),
    "Cumulative_Units": (_ROLE_OMIT, False, "Continuous"),
    "Experience_Stock": (_ROLE_OMIT, False, "Continuous"),
    "Unit_Cost_BY": (_ROLE_OMIT, False, "Continuous"),
    "log Cum Units": (_ROLE_PREDICTOR, True, "Continuous"),
    "log experience": (_ROLE_OMIT, False, "Continuous"),
    "log Unit Cost": (_ROLE_RESPONSE, False, "Continuous"),
    "Full_Data": (_ROLE_FILTER, False, "Continuous"),
}
_PRODUCTION_LOTS_SEQUENCE_VARIABLES: frozenset[str] = frozenset({"Fiscal_Year"})
_PRODUCTION_LOTS_PROFILE = SpecDatasetProfile(
    source_table_ref=f"={PRODUCTION_LOTS.table_name}[#All]",
    variables=_PRODUCTION_LOTS_VARIABLES,
    default_spec=_PRODUCTION_LOTS_DEFAULT_SPEC,
    sequence_variables=_PRODUCTION_LOTS_SEQUENCE_VARIABLES,
)

# The --regression-dataset CLI choice (build_production.py) indexes this
# registry for both the Source_Table retarget and the spec-block defaults —
# adding a new dataset means adding one SpecDatasetProfile and one entry
# here, nothing else.
SPEC_DATASET_PROFILES: dict[str, SpecDatasetProfile] = {
    "auto_mpg": _AUTO_MPG_PROFILE,
    "life_expectancy": _LIFE_EXPECTANCY_PROFILE,
    "production_lots": _PRODUCTION_LOTS_PROFILE,
}

_DEFAULT_TRANSFORM = "None"

_ROLE_VALIDATION_LIST = ",".join(
    (
        _ROLE_RESPONSE,
        _ROLE_PREDICTOR,
        _ROLE_IDENTIFIER,
        _ROLE_FILTER,
        _ROLE_OMIT,
        _ROLE_FIXED_EFFECTS,
    )
)
_INCLUDE_VALIDATION_LIST = "TRUE,FALSE"
_TYPE_VALIDATION_LIST = "Continuous,Categorical"
_TRANSFORM_VALIDATION_LIST = ",".join((_DEFAULT_TRANSFORM, "Log"))
# Sequence flag: TRUE or blank (IgnoreBlank keeps blank legal).
_SEQUENCE_VALIDATION_LIST = "TRUE"
_XL_VALIDATE_LIST = 3
_XL_VALID_ALERT_STOP = 1
_XL_BETWEEN = 1
_RESERVED_NOTE = "Reserved for a future release — not yet used by any formula."
_TRANSFORM_NOTE = (
    "Transform applied to this variable before fitting. “None” "
    "(default) fits the raw column. “Log” fits the natural log — "
    "available on the Response row and on Continuous Predictor rows; not "
    "allowed on Categorical Predictors (flagged red if set — dummy-coded "
    "columns are never logged). The constructed column is relabelled "
    "“Ln(name)” everywhere it appears (Predictor Summary, "
    "coefficient table, Prediction Inputs). Every model statistic "
    "(coefficients, R², residuals, prediction interval) is then in "
    "log space; predictions are NOT back-transformed to the original "
    "units. A zero, negative, or non-numeric value on an included row "
    "makes the model return #N/A rather than fit silently."
)
_SEQUENCE_NOTE = (
    "Sequence structural axis: mark AT MOST ONE variable TRUE as the "
    "ordering axis for lag/difference/serial-correlation features. "
    "Independent of Role and Predictor Type — a Predictor can also be the "
    "sequence axis. Zero flags is valid (non-panel data); two or more is a "
    "spec error shown in the status line above this column. Distinct from "
    "the reserved Order column (F), which is term-ordering."
)
_SEQUENCE_PERIOD_NOTE = (
    "Sequence Period — typed override input for Base Period Δ "
    "(renamed in v2.1.0). On the Sequence-flagged row this cell is blank "
    "by default; type a number here to override the computed candidate "
    "delta. Lag_By and Difference_By read the Period In Use column when "
    "[delta] is omitted. Period In Use shows what the engine will "
    "actually use (the typed override if non-blank, else the computed "
    "candidate from within-group consecutive spacings). "
    "Blank on rows that are not the sequence axis. See the Sequence "
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
# panel DW) and the audit strip's fixed-effects cardinality check. Zero-or-one
# is the legal range (same pattern as the Sequence flag count); two-plus is a
# visible audit-strip error, since two-way absorption is a post-v2.1
# milestone. Same TAKE-trimmed idiom as the responses and sequence-flag
# counts.
_FIXED_EFFECTS_COUNT_FORMULA = (
    "SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))"
    f'="{_ROLE_FIXED_EFFECTS}"))'
)

# The active Fixed Effects variable's header name — the same XMATCH-on-Role
# lookup that fills the spec feedback block's "FE Variable" cell (J2 below),
# factored out so the Residual Output headers (write_sheet_regression.py) can
# reuse it verbatim to build "(Within <name>)" suffixes instead of the bare
# "(Within)" token. Resolves the FIRST FE row even in the 2-plus-rows error
# state, exactly like Fixed_Effects_Column() itself. IFERROR -> "FE" is a
# quiet fallback for the (gated-out) case where no Fixed Effects row exists;
# every call site only evaluates this once {_FIXED_EFFECTS_COUNT_FORMULA} > 0.
_FIXED_EFFECTS_NAME_FORMULA = (
    "IFERROR(INDEX(TOROW(Header_Names),"
    f'XMATCH("{_ROLE_FIXED_EFFECTS}",TAKE(Spec_Role,COLUMNS(Source_Data)))),"FE")'
)

# TRUE when the Response row's Transform (spec column G) is Log — the same
# XMATCH position Response_Column() uses for its own data column, so this
# can never disagree with what Response_Column() is actually returning.
# Feeds the residual-output headers' "(Log)" suffix (write_sheet_regression.py):
# response-scale columns (Y, Predicted Y, Residuals, PRESS Residual) need to
# say so, the same way "(Within)" already flags Fixed Effects demeaning —
# dimensionless diagnostics (Hat Diagonal, Studentized Residuals, Cook's
# Distance, Normal Scores Ranked, Studentized Residuals Ranked,
# Scale-Location) are not in response units and do not get the suffix.
# IFERROR-wrapped to FALSE: during a transient invalid spec state (zero or
# multiple Response rows — already flagged elsewhere by the audit strip's
# responses count), XMATCH itself returns #N/A, which would otherwise
# propagate through INDEX and this comparison into every consumer's AND/IF
# and show #N/A in the residual headers instead of degrading to the plain
# (non-"(Log)") label.
_RESPONSE_LOG_FORMULA = (
    'IFERROR(INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data)),'
    f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data))))="Log",FALSE)'
)

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
    "to the minimum). Review the spectrum and set Δ in spec column I "
    "(Sequence Period)."
)
_MSG_CALENDAR = (
    "Calendar-day spacing (~28–31 / 90–92 / 365–366): do not "
    "set a day-count Δ — build an integer period index upstream "
    "(YEAR, or YEAR*12+MONTH) and flag that as the Sequence axis."
)


def _set_sheet_scoped_names(
    sheet: xw.Sheet,
    closures: Sequence[CatalogFunction],
    source_table_ref: str = "=MileageData[#All]",
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
        "Source_Table": source_table_ref,
        "Source_Data": "=DROP(Source_Table,1)",
        "Header_Names": "=TAKE(Source_Table,1)",
        # ── Spec ranges (table-column structured references) ─────────────
        # The spec data area is a structured table (SpecTable) at
        # B_HEADER_ROW:L_LAST_DATA_ROW; these band names bind to its
        # columns via SpecTable[[#Data],[Column]] structured references.
        # Each column header carries the actual human-readable name (with
        # spaces — Excel requires the exact header text, not a sanitized
        # underscore form, in structured references). The [#Data]
        # qualifier restricts the range to the data body (the spec rows),
        # which is what every TAKE-trimmed consumer expects: the spec
        # rows, not the headers.
        "Spec_Role": f"={sname}!SpecTable[[#Data],[Role]]",
        "Spec_Include": f"={sname}!SpecTable[[#Data],[Include]]",
        "Spec_Type": f"={sname}!SpecTable[[#Data],[Type]]",
        "Spec_Reference": f"={sname}!SpecTable[[#Data],[Reference Level]]",
        # Reserved axes: named now so the grid shape is final, read by
        # nothing until the Order/Transform release.
        "Spec_Order": f"={sname}!SpecTable[[#Data],[Order]]",
        "Spec_Transform": f"={sname}!SpecTable[[#Data],[Transform]]",
        # Sequence structural axis (live: read by the zero-or-one status
        # validation and its conditional formats, not by any constructor);
        # Base Period Δ is its reserved companion, read by nothing until
        # the base-period release. The Period In Use band name is new —
        # it parallels the other Spec_* names so every spec column has a
        # single binding, even though no formula reads it yet.
        "Spec_Sequence": f"={sname}!SpecTable[[#Data],[Sequence]]",
        "Spec_Sequence_Period": (
            f"={sname}!SpecTable[[#Data],[Sequence Period]]"
        ),
        "Spec_Period_In_Use": (
            f"={sname}!SpecTable[[#Data],[Period In Use]]"
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


def _write_spec_block(
    sheet: xw.Sheet, profile: SpecDatasetProfile | None = None
) -> None:
    """The A–L specification block: headers, defaults, dropdowns, CF.

    ``profile`` supplies the variable list and default Role/Include/Type/
    Sequence values — defaults to the shipped Auto MPG profile
    (``SPEC_DATASET_PROFILES["auto_mpg"]``) when omitted, matching this
    function's original hardcoded-to-Auto-MPG behavior.
    """
    profile = profile or _AUTO_MPG_PROFILE
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
        (_C_SEQUENCE_PERIOD, "Sequence Period"),
        (_C_PERIOD_IN_USE, "Period In Use"),
        (_C_LEVELS, "Levels"),
        (_C_REF_IN_USE, "Reference In Use"),
    ):
        val(sheet, _HEADER_ROW, col, header)

    # Create the table before writing any [@Column] formulas below. Excel
    # rejects row-scoped structured references until the target cell belongs
    # to a ListObject with the referenced headers.
    _create_spec_table(sheet, profile)

    # TableStyle overrides the header row's fill; re-pin it to SUBHDR_COLOR
    # (the shared column-sub-header convention) so the header reads
    # consistently with every other sheet regardless of the table style
    # underneath. Covers column A's header too (outside the ListObject, so
    # untouched by TableStyle) for a uniform row.
    sheet.range(
        (_HEADER_ROW, _C_LABEL), (_HEADER_ROW, _C_REF_IN_USE)
    ).color = SUBHDR_COLOR

    # A: variable names spill straight from the table's header row via the
    # Header_Names indirection (dataset-agnostic; reads no other sheet).
    f(sheet, _FIRST_DATA_ROW, _C_LABEL, "=TRANSPOSE(Header_Names)")

    for offset, variable in enumerate(profile.variables):
        row = _FIRST_DATA_ROW + offset
        role, include, ptype = profile.default_spec.get(variable, _FALLBACK_SPEC)
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
        # valid non-panel spec.
        # I (Sequence Period) is the typed override input: the user types a
        # number here on the Sequence-flagged row to declare a Δ that
        # differs from the computed candidate. The cell is styled as an
        # input; left blank, the spec falls back to the candidate.
        # J (Period In Use) is the candidate-with-override display, the
        # reference-level pattern: it shows the typed I value when I is
        # non-blank, otherwise the candidate closure's value. The pattern
        # keeps I as the load-bearing override cell (the candidate closure
        # never overwrites user input), and J is a pure display that the
        # Sequence Spacing block reads.
        if variable in profile.sequence_variables:
            val(sheet, row, _C_SEQUENCE, True)
        format_input(sheet, row, _C_SEQUENCE)
        # I is a pure input: no candidate formula here. The pre-filled
        # candidate is in J; the user types a number into I to override,
        # and the J formula picks the override via the I reference.
        # The J/K/L formulas use structured references ([@Column]) because
        # the spec data area is a structured table (SpecTable); Formula2
        # rejects structured refs, so they go through f_structured.
        format_input(sheet, row, _C_SEQUENCE_PERIOD)
        f_structured(
            sheet,
            row,
            _C_PERIOD_IN_USE,
            (
                '=IF([@Sequence]<>TRUE,"",'
                'IF(N([@[Sequence Period]])<>0,[@[Sequence Period]],'
                'IFERROR(Base_Period_Delta_Candidate(),"")))'
            ),
        )

        # K: Levels display — Categorical Predictors only; the raw distinct
        # level count L over the mask-included rows, with Dummy_Levels'
        # blank normalization mirrored inline. Deliberately NOT a
        # Dummy_Levels call: the display must show L (including 1 for a
        # degenerate column, feeding the red CF below), while Dummy_Levels
        # returns the L−1 retained levels and #N/A when degenerate.
        # ROW()−_ROW_TO_COL_OFFSET maps the sheet row to a Source_Data
        # column index. IFERROR → 0 covers the empty-masked-sample edge.
        f_structured(
            sheet,
            row,
            _C_LEVELS,
            (
                f'=IF(OR([@Role]<>"{_ROLE_PREDICTOR}",'
                f'[@Type]<>"Categorical"),"",'
                f"LET(col,INDEX(Source_Data,0,ROW()-{_ROW_TO_COL_OFFSET}),"
                f'x,IF(col="","",col),'
                f'IFERROR(ROWS(UNIQUE(FILTER(x,(x<>"")*Sample_Include()))),0)))'
            ),
        )

        # L: Reference In Use display — the level the constructor will
        # actually drop, surfaced even when defaulted. A nonblank
        # [@Reference Level] is echoed verbatim (its invalid-reference CF
        # carries the error signal); a blank E shows Dummy_Levels' own
        # default, the first sorted level over the mask-included sample,
        # with the same blank normalization mirrored inline. Deliberately
        # NOT a Dummy_Levels call: the function returns the RETAINED
        # levels, which is exactly the set the reference has been dropped
        # from. IFERROR → "" covers the empty-masked-sample edge (K shows
        # 0 and flags red there).
        f_structured(
            sheet,
            row,
            _C_REF_IN_USE,
            (
                f'=IF(OR([@Role]<>"{_ROLE_PREDICTOR}",'
                f'[@Type]<>"Categorical"),"",'
                f'IF([@[Reference Level]]<>"",[@[Reference Level]],'
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

    # Cascading relevance, Role-keyed: the per-Predictor inputs (C–F) and
    # the Categorical displays (K–L) hide in place whenever Role ≠
    # Predictor — the Reference-only-for-Categorical pattern applied one
    # level up. H–J are deliberately excluded: Sequence is a structural
    # axis, not a Role property (an Identifier like Year is a typical
    # sequence axis). G (Transform) is handled separately, immediately
    # below: unlike C–F, it is also meaningful on the Response row (a
    # Log-transformed response is a first-class case, not a Predictor-only
    # concept), so it cannot share this rule's Role ≠ Predictor test.
    #
    # "Hide in place" means the font color matches each band's own static
    # fill (INPUT_COLOR for format_input-colored cells, white for unfilled
    # computed-display cells) rather than a single muted gray — the same
    # font-matches-fill idiom used for the boundary guard on the Univariate
    # sheet's grid-search tables (write_sheet_univariate.py, cf.Font.Color =
    # 0xFFFFFF). Every range here runs out to _VALIDATION_LAST_ROW, not just
    # _LAST_DATA_ROW: SpecTable is a ListObject, so typing a row directly
    # below its current bottom edge auto-extends the table (structured
    # names and the J/K/L calculated-column formulas follow automatically);
    # pre-applying these rules out to the same 16000-row ceiling the B/C/D/
    # G/H dropdown Validation already uses means a freshly-added row is
    # fully formatted the instant it joins the table, with no rebuild.
    add_expression_format(
        sheet,
        f"$C${_FIRST_DATA_ROW}:$F${_VALIDATION_LAST_ROW}",
        f'=$B{_FIRST_DATA_ROW}<>"{_ROLE_PREDICTOR}"',
        font_color=INPUT_COLOR,
    )
    # G (Transform) hides only when Role is neither Predictor nor Response
    # — visible on both, since Log is declarable on either.
    add_expression_format(
        sheet,
        f"$G${_FIRST_DATA_ROW}:$G${_VALIDATION_LAST_ROW}",
        (
            f'=AND($B{_FIRST_DATA_ROW}<>"{_ROLE_PREDICTOR}",'
            f'$B{_FIRST_DATA_ROW}<>"{_ROLE_RESPONSE}")'
        ),
        font_color=INPUT_COLOR,
    )
    add_expression_format(
        sheet,
        f"$K${_FIRST_DATA_ROW}:$L${_VALIDATION_LAST_ROW}",
        f'=$B{_FIRST_DATA_ROW}<>"{_ROLE_PREDICTOR}"',
        font_color=(255, 255, 255),
    )

    # Cascading relevance, Sequence-keyed: H–J hide in place on every row
    # that is not the sequence axis — Sequence Period and Period In Use are
    # meaningful only for the flagged row, and the flag itself keys on its
    # own value, not on Role. H–I are format_input-colored inputs; J is an
    # unfilled computed display.
    add_expression_format(
        sheet,
        f"$H${_FIRST_DATA_ROW}:$I${_VALIDATION_LAST_ROW}",
        f"=$H{_FIRST_DATA_ROW}<>TRUE",
        font_color=INPUT_COLOR,
    )
    add_expression_format(
        sheet,
        f"$J${_FIRST_DATA_ROW}:$J${_VALIDATION_LAST_ROW}",
        f"=$H{_FIRST_DATA_ROW}<>TRUE",
        font_color=(255, 255, 255),
    )

    # Multi-flag error: red on every flagged Sequence cell when two-plus
    # rows are marked — points at the offending rows while the H2 status
    # line states the error.
    add_expression_format(
        sheet,
        f"$H${_FIRST_DATA_ROW}:$H${_VALIDATION_LAST_ROW}",
        (
            f"=AND($H{_FIRST_DATA_ROW}=TRUE,"
            f"{_SEQUENCE_FLAG_COUNT_FORMULA}>1)"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Period In Use (J) override flagging is intentionally NOT applied on
    # the spec block. The J cells stay plain; the override verdict lives
    # on the Sequence Spacing block (verdict lines on rows 31–34) where
    # the user actually inspects the time grid. The spec block is a
    # declaration — it should read top-to-bottom as a clean grid.

    # Degeneracy flag: red K when an INCLUDED Categorical Predictor has
    # L <= 1 — the constructor contributes zero columns for it (visible
    # degradation, not silent omission). N() coerces "" to 0.
    add_expression_format(
        sheet,
        f"$K${_FIRST_DATA_ROW}:$K${_VALIDATION_LAST_ROW}",
        (
            f'=AND($B{_FIRST_DATA_ROW}="{_ROLE_PREDICTOR}",'
            f"$C{_FIRST_DATA_ROW}=TRUE,"
            f'$D{_FIRST_DATA_ROW}="Categorical",'
            f"N($K{_FIRST_DATA_ROW})<=1)"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Categorical x Log flag: red G when an INCLUDED Categorical Predictor
    # has Transform = Log — disallowed, flagged rather than silently
    # ignored (the "flag red and instruct, never silently switch"
    # precedent already used for the Intercept x Categorical case). The
    # constructor still fits the column as ordinary dummy-coding
    # (Constructed_Column_Transforms() forces every dummy column's flag to
    # "None" regardless of this cell's value) — the red flag is the
    # correction signal, not a computation abort.
    add_expression_format(
        sheet,
        f"$G${_FIRST_DATA_ROW}:$G${_VALIDATION_LAST_ROW}",
        (
            f'=AND($B{_FIRST_DATA_ROW}="{_ROLE_PREDICTOR}",'
            f'$D{_FIRST_DATA_ROW}="Categorical",'
            f'$G{_FIRST_DATA_ROW}="Log")'
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
        f"$E${_FIRST_DATA_ROW}:$E${_VALIDATION_LAST_ROW}",
        (
            f'=AND($E{_FIRST_DATA_ROW}<>"",'
            f"ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-{_ROW_TO_COL_OFFSET}),"
            f"$E{_FIRST_DATA_ROW},Sample_Include())))"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

def _create_spec_table(
    sheet: xw.Sheet, profile: SpecDatasetProfile | None = None
) -> None:
    """Convert the spec data area at B3:L(last data row) into a structured ListObject.

    The table is named ``SpecTable`` (Excel strips special characters and
    prefixes automatically; the name field is the user-visible label). A
    column is outside the table by design — the variable-names spill at
    A4:A(last data row) must not be absorbed by the table's spill scope,
    since the spill lives outside the structured-reference world.

    The table is sized to ``len(profile.variables)`` data rows (the Auto
    MPG profile when ``profile`` is omitted) so every column of the
    targeted dataset gets a Spec_Role/Spec_Include/etc. entry — those are
    ``SpecTable[[#Data],[Column]]`` structured references, so a table sized
    too short for the dataset silently drops the extra columns from every
    constructor closure instead of erroring.

    Headers on row 3 are the existing column labels written by
    _write_spec_block; XlListObjectHasHeaders=xlYes tells Excel to
    promote the first row to headers. The table must exist before the
    Spec_* band names are registered in _set_sheet_scoped_names (Excel
    validates each name's RefersTo at registration time).

    TableStyleLight9 (dark-teal accent) with banding off gives the table a
    neutral body so the conditional formatting (hide-in-place cascading
    relevance, red/yellow error flags) reads clearly on top of it, instead
    of competing with Excel's un-pinned default banding.
    """
    profile = profile or _AUTO_MPG_PROFILE
    last_data_row = _FIRST_DATA_ROW + len(profile.variables) - 1
    table_range = sheet.range(
        (_HEADER_ROW, _C_ROLE), (last_data_row, _C_REF_IN_USE)
    )
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = "SpecTable"
    table.TableStyle = "TableStyleLight9"
    table.ShowTableStyleRowStripes = False
    table.ShowTableStyleColumnStripes = False


def _write_spec_feedback(sheet: xw.Sheet) -> None:
    """The M/N spectrum and the I1/I2 verdict overlay.

    Layout (cells land on row 1 for headers, row 2 for content; row 1 is
    shared with the row-1 audit strip on the right side of the sheet, and
    row 2 holds the row-2 Intercept control — both unaffected by the
    feedback cells to the right of E2):

        M1 = "Δ"          (bold header)
        N1 = "Count"      (bold header)
        M2 = IFERROR(Sequence_Delta_Spectrum(), "")
            — spills downward into empty territory (M and N are spec
              feedback columns with no other content, so the spill never
              collides with the spec block below row 3)
        I1 = "Verdict"    (bold header — overlays the Sequence_Period
                          column's row-1 cell, which is unused)
        I2 = combined switch formula (priority-ordered, single message)

    The E1 Sequence error status is written here too (moved from H2 when
    the spec data area became a structured table (SpecTable) — H2 is now
    the table's "Sequence" header cell, and a status cell on top of a
    table header reads as a visual collision). E1 keeps the same pattern
    as the old H2: blank while the spec is legal, a red error line when
    it is not.

    B1 carries the parallel Fixed Effects cardinality error (same pattern,
    Role's own column instead of Sequence's), and J1/K1/L1 (headers) with
    J2/K2/L2 (values) surface the active FE variable, its group count, and
    the degrees of freedom it absorbs — TODOs #2's status-block cells, live
    now that the phase 1-3 engine backs them instead of showing a
    forthcoming-engine placeholder. J/K/L are the Period In Use / Levels /
    Reference In Use spec columns; rows 1-2 sit above the table's own
    per-row content (row 3 header, row 4+ data), the same free space I1/I2
    already uses for the Sequence Verdict.
    """
    feedback_header_row = 1
    feedback_content_row = 2

    # E1: the Sequence multi-flag error status (moved from H2 — see above).
    # E1 sits inside the A1–L1 row-1 zone above the spec data area; the E
    # column is otherwise Reference Level (an input, blank on row 1 by
    # default), so the status cell shares the column with no other content.
    status_cell = f"${col_letter(_C_REFERENCE)}${feedback_header_row}"  # $E$1
    f(
        sheet,
        feedback_header_row,
        _C_REFERENCE,
        (
            f"=IF({_SEQUENCE_FLAG_COUNT_FORMULA}>1,"
            '"ERROR: multiple Sequence flags (mark at most one variable)","")'
        ),
    )
    bold(sheet, feedback_header_row, _C_REFERENCE)
    add_expression_format(
        sheet,
        status_cell,
        f'={status_cell}<>""',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # B1: the parallel Fixed Effects cardinality error. B is otherwise Role
    # (a dropdown input, blank on row 1 by default), so the status cell
    # shares the column with no other content — same placement logic as E1
    # sharing Reference Level's row-1 cell.
    fe_status_cell = f"${col_letter(_C_ROLE)}${feedback_header_row}"  # $B$1
    f(
        sheet,
        feedback_header_row,
        _C_ROLE,
        (
            f"=IF({_FIXED_EFFECTS_COUNT_FORMULA}>1,"
            '"ERROR: multiple Fixed Effects rows (mark at most one variable)","")'
        ),
    )
    bold(sheet, feedback_header_row, _C_ROLE)
    add_expression_format(
        sheet,
        fe_status_cell,
        f'={fe_status_cell}<>""',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # J1/K1/L1 + J2/K2/L2: the Fixed Effects status block. All three values
    # key off the same FE-count gate, self-guarding like the BFN/DW trigger
    # matrix's "n/a" tokens — "n/a" when no Fixed Effects row is declared,
    # live values once one is (still resolves the FIRST FE row even in the
    # 2-plus-rows error state, exactly like Fixed_Effects_Column() itself;
    # the B1 error above is what flags that state, not these display cells).
    val(sheet, feedback_header_row, _C_PERIOD_IN_USE, "FE Variable")
    val(sheet, feedback_header_row, _C_LEVELS, "FE Groups")
    val(sheet, feedback_header_row, _C_REF_IN_USE, "FE df absorbed")
    bold(sheet, feedback_header_row, _C_PERIOD_IN_USE)
    bold(sheet, feedback_header_row, _C_LEVELS)
    bold(sheet, feedback_header_row, _C_REF_IN_USE)
    f(
        sheet,
        feedback_content_row,
        _C_PERIOD_IN_USE,
        (
            f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"n/a",'
            f"{_FIXED_EFFECTS_NAME_FORMULA})"
        ),
    )
    f(
        sheet,
        feedback_content_row,
        _C_LEVELS,
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"n/a",Absorbed_Degrees_Of_Freedom()+1)',
    )
    f(
        sheet,
        feedback_content_row,
        _C_REF_IN_USE,
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"n/a",Absorbed_Degrees_Of_Freedom())',
    )

    # M1/N1: bold headers (no fill, default font size). The Verdict header
    # (I1) is bolded separately below; it lives in column I (the spec
    # block's Sequence_Period column) and so is NOT in the M:N range.
    val(sheet, feedback_header_row, _C_FEEDBACK_DELTA, "Δ")
    val(sheet, feedback_header_row, _C_FEEDBACK_COUNT, "Count")
    bold_row(
        sheet,
        feedback_header_row,
        _C_FEEDBACK_DELTA,
        _C_FEEDBACK_COUNT,
    )

    # M2: Sequence_Delta_Spectrum() — an N×2 array of (delta, count) pairs,
    # spilling downward into empty territory. IFERROR degrades the
    # no-axis / no-spacings #N/A to a quiet blank.
    f(
        sheet,
        feedback_content_row,
        _C_FEEDBACK_DELTA,
        '=IFERROR(Sequence_Delta_Spectrum(),"")',
    )

    # I1: the Verdict header (bold). I is the Sequence_Period spec column,
    # but only the spec data rows (_FIRST_DATA_ROW.._LAST_DATA_ROW) use it
    # for the per-variable override; the row-1 and row-2 cells are above
    # the spec table and free to carry feedback.
    val(sheet, feedback_header_row, _C_SEQUENCE_PERIOD, "Verdict")
    bold(sheet, feedback_header_row, _C_SEQUENCE_PERIOD)

    # I2: the combined verdict switch — one cell, one message, with
    # conditional formatting carrying the priority (red outranks yellow
    # via StopIfTrue). Picks the most severe applicable verdict:
    #
    #   off-grid       (red)  — data is off the declared grid
    #   regularity     (yel)  — data has spacings besides Δ
    #   no-natural     (yel)  — MODE undefined, candidate falls back to MIN
    #   calendar       (red)  — day-count signature; recommend integer
    #                           period index upstream
    #
    # The switch looks up the Sequence-flagged row's Period In Use value
    # via the registered Spec_Period_In_Use band, paired with XMATCH over
    # Spec_Sequence to find the flagged row. Using the names here avoids
    # Formula2's rejection of direct structured references.
    # outer guard returns blank when no Sequence axis is flagged OR the
    # flagged row's I/J is non-numeric.
    o2_formula = (
        "=LET("
        "d,Sequence_Deltas(),"
        "v,INDEX(Spec_Period_In_Use,"
        "XMATCH(TRUE,Spec_Sequence,0)),"
        "IF(OR(COUNT(d)=0,NOT(ISNUMBER(v))),"
        '"",'
        'IF(SUM(--(MOD(d,v)<>0))>0,"' + _MSG_OFF_GRID + '",'
        'IF(SUM(--(d<>v))>0,"' + _MSG_REGULARITY + '",'
        'IF(ISNA(MODE.SNGL(d)),"' + _MSG_NO_NATURAL + '",'
        'IF(SUM(--((((d>=28)*(d<=31))+((d>=90)*(d<=92))'
        '+((d>=365)*(d<=366)))>0))*2>COUNT(d),'
        '"' + _MSG_CALENDAR + '",'
        '""'
        "))))))"
    )
    f(sheet, feedback_content_row, _C_SEQUENCE_PERIOD, o2_formula)

    # I2 CF: red for off-grid or calendar (StopIfTrue outranks yellow);
    # yellow for regularity or no-natural. Each rule keys on a SEARCH
    # of the cell's rendered text for the message keyword — the same
    # four message constants _MSG_* used to build the formula.
    verdict_cell = f"${col_letter(_C_SEQUENCE_PERIOD)}${feedback_content_row}"  # $I$2
    add_expression_format(
        sheet,
        verdict_cell,
        f'=OR(ISNUMBER(SEARCH("off-grid",{verdict_cell})),'
        f'ISNUMBER(SEARCH("calendar",{verdict_cell})))',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
        stop_if_true=True,
    )
    add_expression_format(
        sheet,
        verdict_cell,
        f'=OR(ISNUMBER(SEARCH("evenly spaced",{verdict_cell})),'
        f'ISNUMBER(SEARCH("no natural",{verdict_cell})))',
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )


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

    * **Hidden (font matches the cell's own INPUT_COLOR fill)** whenever an
      included Categorical predictor is present — the toggle is
      required-here and reads as locked-on even while (correctly) TRUE.
    * **Red** when the toggle is nonetheless set FALSE in that state — the
      invalid combination, flagged not forced. Added first with StopIfTrue so
      it outranks the hide rule on the same cell.
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
    # Required-here signal, applies even while the toggle is still TRUE.
    # C2 is format_input-colored, so hide-in-place uses INPUT_COLOR (same
    # idiom as the spec block's cascading-relevance rules).
    add_expression_format(
        sheet,
        toggle,
        f"={cat_included}",
        font_color=INPUT_COLOR,
    )
    # Intercept x Fixed Effects: fitting an explicit intercept on top of the
    # already-demeaned X_s_Within()/y_s() pair is not a numerical error (the
    # within transform plus Absorbed_Degrees_Of_Freedom() correctly accounts
    # for it — see the DF_Absorbed threading tests) and LINEST just estimates
    # it near zero, but the resulting "Intercept" coefficient row is not the
    # original model's intercept in any interpretable sense — it is an
    # incidental unbalanced-panel artifact. Flagged so a user does not read
    # significance into it.
    add_expression_format(
        sheet,
        toggle,
        f"=AND({toggle}=TRUE,{_FIXED_EFFECTS_COUNT_FORMULA}>0)",
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
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
        ("fixed effects", f"={_FIXED_EFFECTS_COUNT_FORMULA}"),
        ("FE absorbed df", "=Absorbed_Degrees_Of_Freedom()"),
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

    # Zero-or-one Fixed Effects rows — same pattern as Sequence: zero is a
    # valid non-panel spec, two-plus (two-way absorption) is out of scope
    # until its own milestone and a visible spec error until then.
    fixed_effects_col = col_letter(_AUDIT_PAIRS[6][1])
    add_expression_format(
        sheet,
        f"${fixed_effects_col}${_AUDIT_ROW}",
        f"=N(${fixed_effects_col}${_AUDIT_ROW})>1",
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

    # The spec block must run before the names are registered: it creates
    # the structured table (SpecTable), which the Spec_* band names bind
    # to via SpecTable[[#Data],[Column]] references — Excel
    # validates the RefersTo at registration time.
    _write_spec_block(sheet)
    _set_sheet_scoped_names(sheet, closures)
    _write_spec_feedback(sheet)
    _write_intercept_control(sheet)
    _write_row_zones(sheet)
    _write_audit_row(sheet)
    _write_filtered_zones(sheet)

    # Reserved-column and Sequence notes are COM comment calls; keep them
    # out of the RecordingSheet-testable spec block.
    _set_note(sheet, _FIRST_DATA_ROW, _C_ORDER, _RESERVED_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_TRANSFORM, _TRANSFORM_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_SEQUENCE, _SEQUENCE_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_SEQUENCE_PERIOD, _SEQUENCE_PERIOD_NOTE)

    _set_spec_block_column_widths(sheet)
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

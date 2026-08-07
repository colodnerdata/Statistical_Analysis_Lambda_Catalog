"""Spec-block component library for the Regression sheet.

This module is the canonical home of the column/row constants, role tokens,
validation lists, ``SpecDatasetProfile`` records, and the ~15 helper writers
(``_write_spec_block``, ``_write_spec_feedback``, ``_set_sheet_scoped_names``,
``_set_spec_block_column_widths``, ``_set_spec_block_optional_outline_group``,
``_write_intercept_control``, ...) consumed by ``write_sheet_regression.py``.

The module deliberately contains no runner: the spec block lives on the
Regression sheet in the shipped artifact, and the standalone ``Model
Construction`` sheet and ``__main__`` CLI were dropped at the v2.0 release.

Two-axis specification plus the Sequence structural axis (ROADMAP: v2.0 —
Specification-Driven Regression; Sequence added post-v2.0):

Note on numbering: the spec-block changeover shipped as v2.0. In current
comments, v3.0 means the engine-interface release (ROADMAP).

    A        B      C       D     E               F       G         H        I              J              K      L
    Variable Role   Include Type  Reference Level Order   Transform Sequence Sequence Period Period In Use Levels Reference In Use
    (spill)  (drop) (input) (drop)(input)         (rsvd.) (rsvd.)   (flag)   (cand./ovr.)   (disp.)        (disp.)(disp.)

    M                N                     O
    Interaction Term Interaction Operation  Design Columns
    (rsvd. input)    (rsvd. drop)           (disp.)

M and N are the layout-break MAJOR's interaction pair — declared,
validated, and flagged now, read by no constructor until the interaction
wiring release (the same reserved-column pattern F and G went live under).
O is the per-row Design Columns audit: how many columns this spec row
contributes to the constructed design matrix. It is a computed display,
bound by "display derives, never feeds" exactly like J/K/L, and it is
what supplies the PRE-FLIGHT width number the §4b guard reads — the check
has to be answerable from the spec, because constructing a 16,000-column
array in order to discover it does not fit is the failure being prevented.

Right of the spec block sit the Δ/Count spectrum feedback columns (P/Q)
and then a narrow gap column R:

    P    Q      R (gap) S             T        U     V           W            X     Y           Z →
    Δ    Count          Row Labels    Included (brk) Filt.Labels Filt.y       (brk) Filt.Labels Filtered Predictor_Columns
    (=Row_Labels() spill at S4; =Sample_Include() spill at T4 — both
     full-height, never internally filtered. V/W/Y/Z are the FILTERED
     display zones: the only place on the sheet where Sample_Include()
     row-filters anything. Y repeats the filtered labels so the matrix
     reads side-by-side without scrolling back to V.)

Row 1, from column S rightward, holds the bold audit cells as
label/value pairs — labels on S/V/Y/AA/AC/AE/AG/AI, values on the
non-narrow columns T/W/Z/AB/AD/AF/AH/AJ (never on a width-2 break):

    k = COLUMNS(Predictor_Columns()) · rows = ROWS(Predictor_Columns()) · response = <derived name> ·
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
control). It has no engine consumer yet — the engine will read it. Because of
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
log). Log is read by Response_Column(), Predictor_Columns(), Constructed_Column_Names(),
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
    Model Year       → Predictor/Categorical/TRUE  (numeric-valued; NOT a
                                               Sequence axis — Auto MPG is
                                               cross-sectional, so nothing
                                               here is flagged; see
                                               _DEFAULT_SEQUENCE_VARIABLES)
    Origin           → Predictor/Categorical/TRUE  (numeric-valued: 1/2/3)
    MPG              → Response               (derived y)
    Horsepower, Weight → Predictor/Continuous/TRUE
    Cylinders, Displacement, Acceleration → Predictor/Continuous/FALSE (candidates)
    Make, Model?     → Omit                   (text columns parsed out of Car Name)
    Full_Data        → Omit                    (its all-features completeness
                                               flag is redundant with the mask's
                                               built-in completeness and
                                               over-filters; no default Filter)
Full-height contract: ROWS(Predictor_Columns()) = ROWS(Row_Labels()) =
ROWS(Sample_Include()) = 406 always — the constructor reads the mask ONLY
to fix level sets; nothing here ever row-filters. With the real mask live,
the T0 mask-dependent values are real on the sheet: k = 16 (2 continuous +
2 Origin dummies + 12 Model Year dummies), and
SUMPRODUCT(N(Sample_Include())) = 392 (completeness-only on the response
and the two continuous predictors, no Full_Data over-filter).

Not here (deliberately, per release scoping): the QC analyzer
(analyze_model_construction.py) and the Version History / CHANGELOG bump
to v2.0 — those landed in the final wiring PR of that release.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

import xlwings as xw

from .catalog_schema import CatalogFunction
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
    HEADER_COLOR,
    INPUT_COLOR,
)
from .workbook_helpers import (
    add_expression_format,
    anchor_comment_right_of_cell,
    bold,
    bold_row,
    col_letter,
    drop_local_name,
    excel_color,
    f,
    format_input,
    group_and_hide_columns,
    note_dimensions,
    open_or_create_workbook,
    rc,
    section_heading,
    set_column_widths,
    val,
)
from .write_sheet_csv_dataset import LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS

# The constructor closures live on the Regression sheet (scope "Regression"
# in lambda_functions.json) — the spec block now lives there and the
# standalone ``Model Construction`` sheet was dropped at the v2.0 release.
# The helpers in this module remain the canonical owners of the spec-block
# column/row constants, role tokens, validation lists, and sheet-scoped
# named ranges consumed by ``write_sheet_regression.py``.
_CLOSURE_SCOPE = "Regression"

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
# to the C/Include boolean column). The spec block sits one row below it:
# headers on row 3, spec rows from _FIRST_DATA_ROW down.
_INTERCEPT_ROW = 2
_HEADER_ROW = 3
_FIRST_DATA_ROW = 4
# The last spec row of the SHIPPED Auto MPG default — NOT the block's height.
# The block has no fixed height: every part of it sizes itself from
# COLUMNS(Source_Data), so a Source_Table retarget resizes it. This constant
# survives only as the floor apply_spec_case clears to before writing a case,
# so a shorter spec cannot leave a longer one's rows behind.
_LAST_DATA_ROW = _FIRST_DATA_ROW + _N_VARIABLES - 1  # 15
# Sheet row _FIRST_DATA_ROW maps to Source_Data column 1. The CF rules still
# recover a row's column index with ROW()-_ROW_TO_COL_OFFSET; the four
# computed columns no longer need to, since a spill's MAP index IS the
# column number (see the _*_SPILL_FORMULA definitions).
_ROW_TO_COL_OFFSET = _FIRST_DATA_ROW - 1  # 3

# Spec-block columns (1-based). Role precedes Include: the larger
# declaration comes first (dataset semantics before iteration state).
# F/G are the reserved Order/Transform slots; H is the Sequence structural
# flag with I/J as its Sequence Period / Period In Use pair (the
# reference-level pattern: I is the candidate-with-override input,
# J is the in-use display); K and L are the computed Categorical
# displays (Levels count, Reference In Use).
#
# M/N are the interaction pair added by the layout-break MAJOR: M names the
# OTHER operand (dropdown over the variable names), N the operation
# (Product | Difference | Ratio, a closed axis). They are APPENDED rather
# than inserted so every cell a saved spec already filled in keeps both its
# address and its meaning — the cost being that two inputs now sit right of
# the J/K/L computed displays, which reads slightly against the block's
# inputs-then-displays order. That was the cheaper of the two; the
# alternative shifts eight columns to preserve a reading convention.
# Both are RESERVED-and-unwired at this release: validated and flagged on
# the sheet, read by no constructor until the interaction wiring release.
#
# O is the per-row Design Columns audit — a computed display bound by
# "display derives, never feeds" like J/K/L, and the pre-flight width
# number the ARCHITECTURE §4b guard reads.
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
    _C_INTERACTION_TERM,
    _C_INTERACTION_OPERATION,
    _C_DESIGN_COLUMNS,
) = range(1, 16)

# The rightmost spec-block column — the single place the block's extent is
# stated, so header bolding, the SUBHDR fill strip and the outline group all
# move together when a column is appended.
_C_SPEC_LAST = _C_DESIGN_COLUMNS

# Widths for the shared A-O spec block — owned here (not by
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
    _C_INTERACTION_TERM: 18,
    _C_INTERACTION_OPERATION: 18,
    _C_DESIGN_COLUMNS: 14,
}


def _set_spec_block_column_widths(sheet: xw.Sheet) -> None:
    set_column_widths(sheet, _SPEC_COLUMN_WIDTHS.items())


# Spec-block optional columns. The first four columns (Variable, Role,
# Include, Type) are what a regular MLR user actually edits; the rest
# (Reference Level, Order, Transform, Sequence, Sequence Period, Period
# In Use, Levels, Reference In Use, plus the Interaction Term /
# Interaction Operation / Design Columns audit columns) only matter for
# Categorical predictors, the Transform feature, the Sequence axis, panel
# data, and the interaction-pair / audit workflow. The Regression sheet's
# zone-level outline group already wraps A:O as one collapsible block;
# this sub-group nests underneath it so the user can collapse the optional
# columns down to the MLR essentials on demand. The sub-group is collapsed
# by default (columns hidden) so the shipped artifact shows only what a
# regular MLR user needs; click the "+" to expand when Reference Level,
# Sequence, the interaction pair, or the Design Columns audit is in play.
_SPEC_OPTIONAL_FIRST_COL = 5    # E — Reference Level
_SPEC_OPTIONAL_LAST_COL = _C_DESIGN_COLUMNS  # O — Design Columns audit


def _set_spec_block_optional_outline_group(sheet: xw.Sheet) -> None:
    """Group the optional spec columns (E:O) into a sub-outline and collapse.

    Layered under the Regression sheet's zone-level A:O outline group, so
    the spec block has two outline levels: the outer one collapses the
    whole zone, this inner one collapses only the optional part. Collapsed
    by default — the first MLR experience is four visible columns
    (Variable, Role, Include, Type) plus the zone title and intercept
    toggle. The optional columns (Reference Level, Order, Transform,
    Sequence, Sequence Period, Period In Use, Levels, Reference In Use,
    Interaction Term, Interaction Operation, Design Columns) and the
    P/Q spec feedback are all hidden behind the same outline button, so
    the Sequence / interaction / audit workflows are one click away.

    F (Order) is width 0 already; including it in the group is harmless.
    Hiding a spec column does not break anything that reads it: the Spec_*
    bands are per-column range references and the four computed columns are
    spills, neither of which cares whether the column is visible.
    """
    group_and_hide_columns(
        sheet, _SPEC_OPTIONAL_FIRST_COL, _SPEC_OPTIONAL_LAST_COL
    )


# Spec feedback zone (P, Q, I — the verdict overlay): the delta spectrum
# (Sequence_Delta_Spectrum() spill at P2:Q?) sits in P and Q; the combined
# verdict switch lives at I2 (the Sequence_Period column's row-1/row-2
# cells are unused by the spec block, so the verdict overlays them
# without disturbing anything below row 3). Headers on row 1, content on
# row 2 — both sit INSIDE the spec block's zone (which extends from A:Q,
# see the Regression sheet's _ZONES), so a single click on the spec
# outline collapses the spec and its feedback together.
#
# The spectrum used to sit at M/N; the layout-break MAJOR took those two
# columns for the interaction pair (and O for the Design Columns audit), so
# it moved three columns right along with everything after it.
_C_FEEDBACK_DELTA = 16     # P — Δ header / spectrum column 1
_C_FEEDBACK_COUNT = 17     # Q — Count header / spectrum column 2

# Gap before the derived-row zone. One ungrouped column (width 2) so the
# spec outline and the derived-row outline collapse independently.
_C_GAP = 18

# Derived-row zone right of the spec block. S and T hold the full-height
# Row_Labels() / Sample_Include() spills — they honor the full-height
# contract (never row-filtered); the FILTERED display zone is further right.
_C_ROW_LABELS = 19
_C_INCLUDED = 20
_GAP_COLUMN_WIDTH = 2

# Filtered display zone: the ONLY place Sample_Include() row-filters
# anything (everything left of T honors the full-height contract). U and
# X are narrow visual breaks; Y repeats the filtered labels so the matrix
# reads side-by-side without scrolling back to V.
_C_BREAK_LEFT = 21
_C_FILTERED_LABELS = 22
_C_FILTERED_Y = 23
_C_BREAK_MID = 24
_C_MATRIX_LABELS = 25
_C_MATRIX_START = 26

# Row-1 audit strip: label/value pairs marching right from column S,
# values placed on the non-narrow columns (T, W, Z, AB, AD, AF, AH, AJ) so
# no number lands on a width-2 break column.
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
# Response_Column, Row_Labels, Predictor_Columns, Constructed_Column_Names in
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
# safe for a spec row a Source_Table retarget has just brought into the
# block's range, which arrives with no defaults at all, to sit with an
# unset Role: it is inert by construction, not merely
# "untested," so the user can classify new rows at their own pace instead
# of being forced to pick a Role the instant the row exists.
# The v2.1 panel role. Read by the Fixed_Effects_Column() accessor, the
# FE-count guard on the Regression diagnostics (BFN panel Durbin-Watson
# trigger matrix), and — since the phase 1-3 engine work — by
# Absorbed_Degrees_Of_Freedom() and the fit-time Design_Response()/Design_Columns() pair that
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

# The Spec_* bands span the same range for the same reason, so a retarget
# never has to reach outside a band to find the rows it just brought into
# play. ONE ceiling for validation, conditional formatting and the bands
# means there is no way for them to disagree about how far the spec block
# can grow. Each band is TAKE-trimmed to COLUMNS(Source_Data) (see
# _spec_band), so the blank cells below the live spec are never read.
_SPEC_BAND_LAST_ROW = _VALIDATION_LAST_ROW

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

# Variables shipped with their Sequence flag (column H) set TRUE — EMPTY for
# Auto MPG, deliberately.
#
# Model Year used to ship flagged, on the reading that it is "the ordering
# axis for the Auto MPG panel". Auto MPG is not a panel. Each row is a
# distinct car model observed once; there is no unit repeated across periods,
# so there is no time axis to order along. What the flag actually bought was a
# Base Period Δ candidate nobody can interpret and a Durbin-Watson statistic
# computed over an arbitrary row order — the shipped Identifier (Car Name) is
# very nearly unique, so Sequence_Deltas finds no within-group consecutive
# pairs and the spacing verdict comes back blank regardless. A default that
# asserts panel structure the data does not have is worse than no default.
#
# The Sequence axis is still demonstrated by default, on the two datasets that
# genuinely have one: _LIFE_EXPECTANCY_SEQUENCE_VARIABLES (Country x Year) and
# _PRODUCTION_LOTS_SEQUENCE_VARIABLES (Facility x Fiscal_Year), both reachable
# through --regression-dataset. On Auto MPG the layer self-reports
# "n/a — requires Sequence" until a user types TRUE in column H, which is the
# honest state.
#
# Whatever a profile puts here is kept to at most one entry: the H2 status line
# errors at two-plus flags.
_DEFAULT_SEQUENCE_VARIABLES: frozenset[str] = frozenset()


# ── Per-dataset spec profiles ──────────────────────────────────────────────
# _VARIABLES/_DEFAULT_SPEC/_DEFAULT_SEQUENCE_VARIABLES above are the shipped
# Auto MPG defaults; SpecDatasetProfile wraps a dataset's variable list,
# default Role/Include/Type spec, and Sequence-flagged columns as one unit
# so retargeting Source_Table (the --regression-dataset CLI choice) can
# also retarget the spec block's defaults, instead of leaving every column
# of a newly-targeted dataset to _FALLBACK_SPEC's un-flagged Predictor.
# The profile decides which rows arrive PRE-FILLED, not how many spec rows
# exist — the block sizes itself from COLUMNS(Source_Data). So every column
# of the targeted dataset carries a real Role/Include/Type from the first
# build instead of falling back to _FALLBACK_SPEC, and a dataset the build
# did not target still gets a working (if unfilled) block on retarget.
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
# plus the appended Full_Data column. Country is a text identifier (row
# labeling only), Year is the natural panel ordering axis (Sequence-flagged,
# Role Omit so it never enters the design matrix itself), Status is the one
# categorical predictor, and Full_Data ships Omit for the same reason Auto
# MPG's does: the built-in per-predictor completeness mask already covers it,
# so a Filter role would only over-filter.
#
# The SHIPPED DEFAULT is the curated four-driver model both presentation decks
# headline (the slide-19 coefficient table): Life expectancy ~ Adult Mortality
# + Alcohol + percentage expenditure + C(Status). Those four predictors ship
# Include=True; every other FEATURE_COLUMNS entry ships Include=False — present
# in the block (so the block still sizes to all 23 source columns) but off,
# ready to toggle on for the EDA / VIF-trim / kitchen-sink beats. This is the
# cold open: k ≈ 5, fast on a low-compute machine, matching the deck. The
# 18-predictor kitchen sink is still a registered case (L05,
# ``_life_full_profile_spec``), just no longer what the workbook opens with —
# and the curated default's own oracle is the registered case ``life_talk_demo``
# (L11). ``_LIFE_EXPECTANCY_VARIABLES`` is unchanged: the block sizes itself
# from COLUMNS(Source_Data), so the 19 dormant rows cost nothing and are one
# toggle away.
_LIFE_EXPECTANCY_VARIABLES: tuple[str, ...] = (
    "Country",
    "Year",
    "Status",
    "Life expectancy",
    *_LIFE_EXPECTANCY_FEATURE_COLUMNS,
    "Full_Data",
)
# The three curated drivers that ship active.
_LIFE_TALK_DEMO_PREDICTORS: frozenset[str] = frozenset(
    {"Adult Mortality", "Alcohol", "percentage expenditure"}
)
_LIFE_EXPECTANCY_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Country": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Year": (_ROLE_OMIT, False, "Continuous"),
    "Status": (_ROLE_PREDICTOR, True, "Categorical"),
    "Life expectancy": (_ROLE_RESPONSE, False, "Continuous"),
    **{
        column: (
            _ROLE_PREDICTOR,
            column in _LIFE_TALK_DEMO_PREDICTORS,
            "Continuous",
        )
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

# Interaction Operation (N): a CLOSED axis in the same sense as Predictor
# Type — these three exhaust the operations an interaction column can be
# built from, and the list never grows. Each carries a symmetry attribute
# the reciprocal-declaration flag below keys on: Product is symmetric and
# Difference antisymmetric (a reciprocal declaration duplicates or negates
# a column, giving a singular Gram matrix), while Ratio is asymmetric, so
# its reciprocal is a legitimately different column.
_INTERACTION_PRODUCT = "Product"
_INTERACTION_DIFFERENCE = "Difference"
_INTERACTION_RATIO = "Ratio"
_INTERACTION_OPERATION_VALIDATION_LIST = ",".join(
    (_INTERACTION_PRODUCT, _INTERACTION_DIFFERENCE, _INTERACTION_RATIO)
)
# The operations whose reciprocal declaration is degenerate. Ratio is
# deliberately absent — B/A is not A/B, so declaring both is legal.
_INTERACTION_SYMMETRIC_OPERATIONS = (
    _INTERACTION_PRODUCT,
    _INTERACTION_DIFFERENCE,
)

# The operator each operation contributes to an interaction column's HEADER
# (v3.1). One symbol per operation, because a single separator cannot say
# which of the three built the column — and the colon this replaced was
# doubly ambiguous, since a level-qualified categorical name already contains
# ": " ("Weight:Status: Developing" reads as one name with two colons).
#
# U+2212 MINUS SIGN, not a hyphen: a hyphen is a legal character in a source
# column name, so "Unit-Cost - Weight" would be unreadable with one. The
# symbols are spaced so they stay legible beside names that contain spaces.
#
# `Constructed_Column_Names()` renders these via a SWITCH over the same
# operation strings; `test_interaction_header_symbols_match_the_catalog_formula`
# pins the two together so this table cannot drift from the formula.
_INTERACTION_HEADER_SYMBOLS = {
    _INTERACTION_PRODUCT: " × ",
    _INTERACTION_DIFFERENCE: " − ",
    _INTERACTION_RATIO: " ÷ ",
}
# Rendered when the operation is none of the three — reachable only by a
# paste past the dropdown, and paired with the NA() column Predictor_Columns()
# emits for the same input. The header still exists so the strip stays exactly
# as wide as the design matrix.
_INTERACTION_HEADER_UNKNOWN = " ? "

# Interaction Term (M): the dropdown source is the variable-name spill at
# A{_FIRST_DATA_ROW}, referenced with the spill operator so the list is
# exactly the dataset's columns and resizes with a retarget — no fixed
# range to keep in sync, and no volatile OFFSET. The spill itself is
# =TRANSPOSE(Header_Names), so the offered names are always the live
# table headers.
_INTERACTION_TERM_VALIDATION_FORMULA = f"=$A${_FIRST_DATA_ROW}#"
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
_INTERACTION_TERM_NOTE = (
    "Interaction Term — names the OTHER operand of an interaction "
    "involving this row. Blank (the default) means no interaction. Only a "
    "Predictor may be an operand: any other Role on the named row is an "
    "error (flagged red), as is a name that is not a variable in this "
    "table. Naming a Predictor whose Include is FALSE is ALLOWED and "
    "flagged amber — an interaction without its main effect is a "
    "marginality violation, usually a mistake but occasionally "
    "deliberate, and blocking it would be the library deciding a "
    "modeling question. Pointing at this row's OWN variable with "
    "Operation = Product is the documented way to declare a quadratic "
    "(x squared) term. The columns this adds are counted in the Design "
    "Columns audit: one for Continuous x Continuous, one per retained "
    "level for Continuous x Categorical, and the full product for "
    "Categorical x Categorical."
)
_INTERACTION_OPERATION_NOTE = (
    "Interaction Operation — how the two operands combine: Product "
    "(symmetric), Difference (antisymmetric), or Ratio (asymmetric). A "
    "closed axis, like Predictor Type: these three never grow. Declaring "
    "B on A as well as A on B is flagged red for Product and Difference "
    "— the reciprocal produces a duplicate or exact-negative column and a "
    "singular Gram matrix, and it is flagged, never silently "
    "deduplicated. Ratio is asymmetric, so its reciprocal is a different "
    "column and is allowed. Ratio returns #N/A where the denominator is "
    "zero, rather than a divide-by-zero error."
)
_DESIGN_COLUMNS_NOTE = (
    "Design Columns — how many columns THIS spec row contributes to the "
    "constructed design matrix. Blank when the row is not a Predictor; 0 "
    "when it is excluded or degenerate; 1 for a Continuous Predictor; "
    "L-1 for a Categorical one, where L is the Levels count beside it — "
    "plus this row's interaction columns, which are its own count times "
    "the operand's. This is the column where one dropdown change becomes "
    "visible: switching a high-cardinality variable to Categorical, or "
    "interacting two of them, can add hundreds of columns. A computed display — no constructor reads it — "
    "and the pre-flight number behind the design-matrix width guard "
    "above, which is why the check is answerable from the spec instead "
    "of by building a matrix that turns out not to fit."
)

# Plain-language tooltips for the spec-block column headers. Eight of the
# twelve spec headers get their own note; the other four (Order, Transform,
# Sequence, Sequence Period) use the longer notes defined above. Tone
# matches the existing notes: one short paragraph, no formula jargon.
_LABEL_NOTE = (
    "The header names from your Source_Table, in source order. Edit the "
    "table headers in Name Manager or on the data sheet, not this column "
    "— these cells spill from Header_Names and are read-only."
)
_ROLE_NOTE = (
    "What this column is used as: Response (y), Predictor (x), Identifier "
    "(row label), Filter (sample mask), Fixed Effects (panel group), or "
    "Omit (ignored). Exactly one Response is allowed; at most one Fixed "
    "Effects; zero or many of the others."
)
_INCLUDE_NOTE = (
    "TRUE/FALSE on/off switch. When FALSE, this variable is excluded from "
    "the model without losing its column on the data sheet. The C2 cell "
    "above the table is the model-level Intercept toggle."
)
_TYPE_NOTE = (
    "Continuous, Categorical, or Identifier. Continuous goes in raw; "
    "Categorical is dummy-coded with the level in column E as the "
    "reference. Identifier rows contribute no model terms and are not "
    "counted in the degrees of freedom."
)
_REFERENCE_NOTE = (
    "For Categorical predictors, the level whose value the intercept "
    "absorbs. Defaults to the first-in-sort-order level; type any level to "
    "override. The cell turns red if the typed level is not present in the "
    "analysis sample."
)
_PERIOD_IN_USE_NOTE = (
    "The Base Period Δ actually in effect on the Sequence-flagged row: "
    "your typed override from column I, or the computed candidate from "
    "Sequence_Delta_Spectrum() (the most common gap within the FE group) "
    "when I is blank. Lag_By and Difference_By read this cell."
)
_LEVELS_NOTE = (
    "Count of distinct non-blank values of this Categorical predictor in "
    "the analysis sample. A value of 1 on a Categorical row means a "
    "single-level predictor, which contributes no columns — the cell is "
    "flagged red so you can drop the row or widen the filter."
)
_REF_IN_USE_NOTE = (
    "The level that is actually serving as the reference for dummy "
    "coding: the typed value from column E if you supplied one, else the "
    "first-in-sort-order default. Read this cell when an unexpected level "
    "shows up as the comparison baseline."
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

# ── The four computed spec columns (J, K, L, O) ────────────────────────────
#
# Each is a SINGLE dynamic-array formula written once at _FIRST_DATA_ROW,
# spilling one value per source-table column. This is the mechanism that
# makes a Source_Table retarget resize the spec block.
#
# They used to be per-row formulas using [@Column] structured references
# inside the SpecTable ListObject, which pinned the block to the row count
# baked in at build time — a wider dataset left the computed columns with no
# formulas at all in the rows it added. A spill cannot live inside a
# ListObject, and J/K/L sit between the input columns I and M while O sits
# after N, so making these self-sizing is exactly why the table was dropped.
#
# Consequences worth knowing when editing them:
#   * `[@Column]` becomes INDEX(<band>,i) and ROW()-_ROW_TO_COL_OFFSET
#     becomes `i`, so a formula no longer depends on which row it occupies.
#   * They are written with `f` (Formula2), never `f_structured`
#     (Formula) — Formula enters an array formula as a legacy CSE range,
#     which would NOT resize on retarget and so would silently reintroduce
#     the bug this change removes.
#   * Nothing may be written into the cells below them in columns J/K/L/O;
#     a spill blocked by stray content is a #SPILL! error, not a truncation.
#
# J — Period In Use. The candidate-with-override display: the typed I value
# when non-blank, else the computed candidate, and blank on every row that
# is not the sequence axis. Base_Period_Delta_Candidate() is hoisted out of
# the MAP because it does not vary by row — the per-row version called it
# once per row, and it walks the whole sequence column each time.
_PERIOD_IN_USE_SPILL_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "sq,TAKE(Spec_Sequence,nc),"
    "sp,TAKE(Spec_Sequence_Period,nc),"
    'cand,IFERROR(Base_Period_Delta_Candidate(),""),'
    "MAP(SEQUENCE(nc),LAMBDA(i,"
    'IF(INDEX(sq,i)<>TRUE,"",'
    "IF(N(INDEX(sp,i))<>0,INDEX(sp,i),cand)))))"
)

# K — Levels display. Categorical Predictors only; the raw distinct level
# count L over the mask-included rows, with Dummy_Levels' blank
# normalization mirrored inline. Deliberately NOT a Dummy_Levels call: the
# display must show L (including 1 for a degenerate column, feeding the red
# CF), while Dummy_Levels returns the L−1 retained levels and #N/A when
# degenerate. IFERROR -> 0 covers the empty-masked-sample edge.
_LEVELS_SPILL_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "rl,TAKE(Spec_Role,nc),"
    "typ,TAKE(Spec_Type,nc),"
    "si,Sample_Include(),"
    "MAP(SEQUENCE(nc),LAMBDA(i,"
    f'IF(OR(INDEX(rl,i)<>"{_ROLE_PREDICTOR}",'
    'INDEX(typ,i)<>"Categorical"),"",'
    "LET(col,INDEX(Source_Data,0,i),"
    'x,IF(col="","",col),'
    'IFERROR(ROWS(UNIQUE(FILTER(x,(x<>"")*si))),0))))))'
)

# L — Reference In Use display. The level the constructor will actually
# drop, surfaced even when defaulted. A nonblank Reference Level is echoed
# verbatim (its invalid-reference CF carries the error signal); a blank one
# shows Dummy_Levels' own default, the first sorted level over the
# mask-included sample. Deliberately NOT a Dummy_Levels call: that function
# returns the RETAINED levels, which is the set the reference has been
# dropped from. IFERROR -> "" covers the empty-masked-sample edge (K shows 0
# and flags red there).
_REF_IN_USE_SPILL_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "rl,TAKE(Spec_Role,nc),"
    "typ,TAKE(Spec_Type,nc),"
    "refs,TAKE(Spec_Reference,nc),"
    "si,Sample_Include(),"
    "MAP(SEQUENCE(nc),LAMBDA(i,"
    f'IF(OR(INDEX(rl,i)<>"{_ROLE_PREDICTOR}",'
    'INDEX(typ,i)<>"Categorical"),"",'
    'IF(INDEX(refs,i)<>"",INDEX(refs,i),'
    "LET(col,INDEX(Source_Data,0,i),"
    'x,IF(col="","",col),'
    'IFERROR(INDEX(SORT(UNIQUE(FILTER(x,(x<>"")*si))),1,1),"")))))))'
)

# O — the per-row Design Columns audit. Mirrors
# Predictor_Columns()'s own iteration predicate and its degenerate skip
# EXACTLY, rather than re-deriving the count from the K (Levels) display:
#
#   Role <> Predictor            -> ""  (the column is not a candidate)
#   Include <> TRUE              -> 0   (candidate, currently out)
#   Type <> Categorical          -> 1   (one column, transform or not)
#   otherwise                    -> COLUMNS(Dummy_Levels(...)), i.e. L-1,
#                                   and 0 when Dummy_Levels signals #N/A
#                                   (degenerate column or invalid
#                                   reference — the constructor's acc
#                                   passthrough, which contributes nothing)
#
# The reference normalization is the constructor's own
# (IF(LEN(d&"")=0,"",d)), so a blank Reference Level resolves to the same
# default level here as it does inside Predictor_Columns(). Reading K
# instead would make one display depend on another; reading the same
# closure the constructor reads makes them provably consistent, which is
# the "one source of truth is the FUNCTION" rule from ARCHITECTURE §4.
#
# From v3.1 the row's INTERACTION columns are counted too, because the
# constructor now builds them. The count is k(row) * k(operand), the width
# of the pairwise combination Predictor_Columns() emits, and it reuses the
# SAME per-row width helper (``kk``) for both operands — so the audit cannot
# disagree with the constructor about how wide a categorical operand is.
# The gating mirrors the constructor's mate() exactly: blank M or blank N,
# a name matching no column, or an operand whose Role is not Predictor all
# contribute 0, leaving the row's main-effect count alone. An operand that
# is a Predictor with Include = FALSE still counts — that is the
# flagged-amber marginality case, which builds columns.
#
# A degenerate row needs no special case in either direction: kk returns 0,
# and 0 * anything is 0, which is exactly what the constructor's skip does.
#
# ONE spill for the whole column, not one formula per row. Every computed
# spec column is written this way (see _PERIOD_IN_USE_SPILL_FORMULA for the
# full rationale): the row count follows COLUMNS(Source_Data), so a
# Source_Table retarget resizes the audit instead of leaving it short.
#
# MAP(SEQUENCE(nc),...) rather than BYROW: the body needs the column INDEX
# to reach INDEX(Source_Data,0,i), and BYROW passes a row's values, not its
# position. That index also replaces the old ROW()-_ROW_TO_COL_OFFSET
# arithmetic, so the formula no longer depends on where it is written.
#
# Everything invariant across rows is hoisted into the outer LET —
# Sample_Include(), TOROW(Header_Names) and the kk helper itself. The
# per-row version re-evaluated all three once per row; kk still calls
# Sample_Include() once per categorical operand, but the closure is now
# built once for the column rather than once per row.
_DESIGN_COLUMNS_SPILL_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "typ,TAKE(Spec_Type,nc),"
    "refs,TAKE(Spec_Reference,nc),"
    "rl,TAKE(Spec_Role,nc),"
    "inc,TAKE(Spec_Include,nc),"
    "it,TAKE(Spec_Interaction_Term,nc),"
    "io,TAKE(Spec_Interaction_Operation,nc),"
    "si,Sample_Include(),"
    "hdr,TOROW(Header_Names),"
    'kk,LAMBDA(x,IF(INDEX(typ,x)<>"Categorical",1,'
    "IFERROR(COLUMNS(Dummy_Levels(INDEX(Source_Data,0,x),"
    'IF(LEN(INDEX(refs,x)&"")=0,"",INDEX(refs,x)),'
    "si)),0))),"
    "MAP(SEQUENCE(nc),LAMBDA(i,"
    f'IF(INDEX(rl,i)<>"{_ROLE_PREDICTOR}","",'
    "IF(INDEX(inc,i)<>TRUE,0,"
    "LET(k,kk(i),"
    "t,INDEX(it,i),"
    "o,INDEX(io,i),"
    "q,IFERROR(XMATCH(t,hdr),0),"
    'ki,IF(OR(LEN(t&"")=0,LEN(o&"")=0,q=0),0,'
    f'IF(INDEX(rl,q)<>"{_ROLE_PREDICTOR}",0,k*kk(q))),'
    "k+ki))))))"
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


def _spec_band(sname: str, col: int) -> str:
    """One spec column, TAKE-trimmed to the live source-table width.

    **This is what makes a ``Source_Table`` retarget a genuine one-name
    edit.** These bands used to be structured references into the
    ``SpecTable`` ListObject (``SpecTable[[#Data],[Role]]``), which fixed
    them at the row count baked in at build time. Retargeting to a wider
    table then left every band short: ``TAKE`` does not pad, so
    ``INDEX(rl, n_c)`` ran off the end and every engine cell downstream
    read as an error. Excel offers no formula-driven way to resize a
    ListObject and this workbook is macro-free, so the table was removed
    and the bands size themselves instead.

    ``TAKE`` (not ``OFFSET``) for the same reason ``Source_Data`` and
    ``Header_Names`` use it: it is non-volatile, so the band is not
    re-evaluated on every recalculation pass.

    ``MAX(1,...)`` keeps the name resolvable while ``Source_Table`` is
    momentarily broken — mid-retarget, a zero-row TAKE would be an error
    that every dependent name would inherit.

    The band runs to ``_SPEC_BAND_LAST_ROW`` rather than to the spec rows
    in use, so rows a retarget brings into play are already inside it. The
    cells below the live spec are ordinary blanks; the trim is what keeps
    them out, and it is why the Regression sheet may never place content
    below the spec block in columns B–O.
    """
    first = f"${col_letter(col)}${_FIRST_DATA_ROW}"
    last = f"${col_letter(col)}${_SPEC_BAND_LAST_ROW}"
    return f"=TAKE({sname}!{first}:{last},MAX(1,COLUMNS(Source_Data)))"


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
       ``Sample_Include``/``Response_Column``/``Row_Labels``/``Predictor_Columns``/
       ``Constructed_Column_Names``, sourced from ``lambda_functions.json``
       (scope ``"Regression"``) so their definitions live in one declarative
       place and appear on the LAMBDA_functions catalog sheet. Passed in
       document order, which is dependency order (``Sample_Include`` before
       ``Predictor_Columns``, etc.).

    This runs BEFORE the spec block is written, because the four computed
    spec columns are spill formulas that reference these band names. It
    used to run after, when the bands bound to a ListObject that
    ``_write_spec_block`` had to create first.
    """
    sname = f"'{sheet.name}'"

    local_names: dict[str, str] = {
        # ── Source-table indirection: THE dataset-retarget point ─────────
        # Source_Table is the only name that references the table; the data
        # body and header row derive from it, so a dataset changeover is a
        # ONE-name edit. DROP/TAKE (not OFFSET) keep the derivations
        # non-volatile — a volatile Header_Names would be re-evaluated on
        # every recalculation pass during workbook calculation.
        "Source_Table": source_table_ref,
        "Source_Data": "=DROP(Source_Table,1)",
        "Header_Names": "=TAKE(Source_Table,1)",
        # ── Spec ranges (dataset-sized dynamic bands) ────────────────────
        # Each band is the column's full input range TAKE-trimmed to the
        # live column count, so retargeting Source_Table resizes every
        # band with it. See _spec_band for why this is not a structured
        # reference into a ListObject any more.
        "Spec_Role": _spec_band(sname, _C_ROLE),
        "Spec_Include": _spec_band(sname, _C_INCLUDE),
        "Spec_Type": _spec_band(sname, _C_TYPE),
        "Spec_Reference": _spec_band(sname, _C_REFERENCE),
        # Reserved axes: named now so the grid shape is final, read by
        # nothing until the Order/Transform release.
        "Spec_Order": _spec_band(sname, _C_ORDER),
        "Spec_Transform": _spec_band(sname, _C_TRANSFORM),
        # Sequence structural axis (live: read by the zero-or-one status
        # validation and its conditional formats, not by any constructor);
        # Base Period Δ is its reserved companion, read by nothing until
        # the base-period release. The Period In Use band name is new —
        # it parallels the other Spec_* names so every spec column has a
        # single binding, even though no formula reads it yet.
        "Spec_Sequence": _spec_band(sname, _C_SEQUENCE),
        "Spec_Sequence_Period": _spec_band(sname, _C_SEQUENCE_PERIOD),
        "Spec_Period_In_Use": _spec_band(sname, _C_PERIOD_IN_USE),
        # The interaction pair (M/N) and the Design Columns audit (O),
        # added by the layout-break MAJOR. The first two are RESERVED —
        # bound so the grid shape is final and so the conditional-format
        # rules on the pair have a band to read, but consumed by no
        # constructor until the interaction wiring release. The third is
        # a computed display, bound by "display derives, never feeds":
        # only the width guard reads it, and the guard is a display too.
        "Spec_Interaction_Term": _spec_band(sname, _C_INTERACTION_TERM),
        "Spec_Interaction_Operation": _spec_band(
            sname, _C_INTERACTION_OPERATION
        ),
        "Spec_Design_Columns": _spec_band(sname, _C_DESIGN_COLUMNS),
        # Model-level Intercept toggle (row-2 control): a single boolean cell
        # in the C/Include column. No engine formula reads it yet — the engine
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

    # Constructor closures (Sample_Include, Response_Column, Row_Labels, Predictor_Columns,
    # Constructed_Column_Names) come from lambda_functions.json, in document
    # (= dependency) order. Their full rationale — the REDUCE-product mask, the
    # (col+0) Filter coercion that avoids N()'s implicit intersection, the
    # once-bound Dummy_Levels skip, the Predictor_Columns / Constructed_Column_Names twin —
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


def _set_note(
    sheet: xw.Sheet, row: int, col: int, text: str, *, label: str | None = None
) -> None:
    """Replace the cell's note/comment text.

    Sized and anchored to the right of the cell via the shared
    `note_dimensions` + `anchor_comment_right_of_cell` helpers (same
    heuristic as the regression sheet's notes). The ``label`` keyword
    mirrors the regression sheet's signature so a notes override map
    can be threaded through if any of the spec-block notes clip.
    """
    cell = sheet.range(rc(row, col))
    cell_api = cell.api
    try:
        cell_api.ClearComments()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    cell_api.AddComment(text)
    try:
        cell_api.Comment.Visible = False
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    width, height = note_dimensions(label if label is not None else text, text)
    anchor_comment_right_of_cell(sheet, row, col, width, height)


def _interaction_error_formats(sheet: xw.Sheet) -> None:
    """The three conditional-formatting flags on the interaction pair M/N.

    All three resolve the named OTHER operand the same way: XMATCH the
    typed name against the live table headers to get its source-column
    index ``j``, then read the target row's own spec cells at that index.
    ``j`` is 0 when the name matches nothing, and every INDEX is taken at
    ``MAX(j,1)`` so a miss produces a clean FALSE instead of an #N/A that
    would propagate out of the enclosing AND (and silently disable the
    rule — an error result means "do not format", which is exactly the
    wrong answer for a flag whose job is to catch bad input).

    1. **Red on M** — the named operand is not a variable in this table,
       or its Role is not Predictor. Only a Predictor may be an operand.
    2. **Amber on M** — the operand IS a Predictor but is excluded. An
       interaction without its main effect is a marginality violation:
       usually a mistake, occasionally deliberate, so it is flagged and
       allowed. Blocking it would be the library deciding a modeling
       question.
    3. **Red on N** — a reciprocal declaration under a symmetric or
       antisymmetric operation: this row names the other, the other names
       this one back, and both carry the same Product or Difference.
       That is a duplicate or exact-negative column and a singular Gram
       matrix. Flagged, never silently deduplicated. Ratio is excluded
       (it is asymmetric — B/A is a different column from A/B), and a row
       naming ITSELF is excluded (self × self under Product is the
       documented way to declare a quadratic term, not a reciprocal).
    """
    r = _FIRST_DATA_ROW
    last = _VALIDATION_LAST_ROW
    m = f"$M{r}"
    n = f"$N{r}"
    # Shared prologue: nc = spec width, hdr = the live header row, j = the
    # named operand's source-column index (0 = no match), p = j clamped
    # into range so every INDEX below is well-formed.
    lookup = (
        "nc,COLUMNS(Source_Data),"
        "hdr,TOROW(Header_Names),"
        f"j,IFERROR(XMATCH({m},hdr),0),"
        "p,MAX(j,1),"
    )

    add_expression_format(
        sheet,
        f"$M${r}:$M${last}",
        (
            f"=LET({lookup}"
            f'AND({m}<>"",OR(j=0,'
            f'INDEX(TAKE(Spec_Role,nc),p)<>"{_ROLE_PREDICTOR}")))'
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
        stop_if_true=True,
    )
    add_expression_format(
        sheet,
        f"$M${r}:$M${last}",
        (
            f"=LET({lookup}"
            f'AND({m}<>"",j>0,'
            f'INDEX(TAKE(Spec_Role,nc),p)="{_ROLE_PREDICTOR}",'
            "INDEX(TAKE(Spec_Include,nc),p)<>TRUE))"
        ),
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
        stop_if_true=True,
    )

    # This row's own source-column index, clamped the same way: the CF
    # band runs out to _VALIDATION_LAST_ROW, well past the spec rows, and
    # an unclamped INDEX(hdr,1,i) there would error. Clamping is safe
    # because the rule also requires a non-blank Operation, which no row
    # below the table has.
    reciprocal_ops = ",".join(
        f'{n}="{op}"' for op in _INTERACTION_SYMMETRIC_OPERATIONS
    )
    add_expression_format(
        sheet,
        f"$N${r}:$N${last}",
        (
            f"=LET({lookup}"
            f"i,ROW()-{_ROW_TO_COL_OFFSET},"
            "q,MIN(MAX(i,1),nc),"
            f"AND(OR({reciprocal_ops}),j>0,j<>i,"
            "INDEX(TAKE(Spec_Interaction_Term,nc),p)=INDEX(hdr,1,q),"
            f"INDEX(TAKE(Spec_Interaction_Operation,nc),p)={n}))"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
        stop_if_true=True,
    )


def _write_spec_block(
    sheet: xw.Sheet,
    profile: SpecDatasetProfile | None = None,
) -> None:
    """The A–O specification block: headers, defaults, dropdowns, CF.

    ``profile`` supplies the variable list and default Role/Include/Type/
    Sequence values — defaults to the shipped Auto MPG profile
    (``SPEC_DATASET_PROFILES["auto_mpg"]``) when omitted, matching this
    function's original hardcoded-to-Auto-MPG behavior.

    **The block has no fixed height.** The profile decides which rows get
    shipped *defaults*, not how many rows exist: the variable-name column,
    the four computed columns and the input band's fill all size themselves
    from ``COLUMNS(Source_Data)``, so retargeting ``Source_Table`` resizes
    the block. This block used to build a ``SpecTable`` ListObject sized to
    ``len(profile.variables)``, which pinned it to the build-time dataset.

    Must run AFTER ``_set_sheet_scoped_names``: the computed columns are
    spills that reference the ``Spec_*`` bands, ``Source_Data`` and the
    constructor closures. The dependency used to run the other way.
    """
    profile = profile or _AUTO_MPG_PROFILE
    bold_row(sheet, _HEADER_ROW, _C_LABEL, _C_SPEC_LAST)
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
        (_C_INTERACTION_TERM, "Interaction Term"),
        (_C_INTERACTION_OPERATION, "Interaction Operation"),
        (_C_DESIGN_COLUMNS, "Design Columns"),
    ):
        val(sheet, _HEADER_ROW, col, header)

    # Create the table before writing any [@Column] formulas below. Excel
    # rejects row-scoped structured references until the target cell belongs
    # to a ListObject with the referenced headers.
    # The specification header row. Nothing overrides it any more — this
    # used to have to be re-pinned AFTER _create_spec_table, because
    # applying a TableStyle silently replaced the header fill and font.
    # test_spec_block_prefills_the_t0_default_configuration still asserts
    # all three properties.
    header_range = sheet.range((_HEADER_ROW, _C_LABEL), (_HEADER_ROW, _C_SPEC_LAST))
    header_range.color = HEADER_COLOR
    header_range.api.Font.Bold = True
    header_range.api.Font.Color = excel_color((0, 0, 0))

    # A: variable names spill straight from the source table's header row
    # via the Header_Names indirection (dataset-agnostic; reads no other
    # sheet). This column has always resized with a retarget; the four
    # computed columns below now do the same, and the input columns follow
    # via the CF band rule.
    f(sheet, _FIRST_DATA_ROW, _C_LABEL, "=TRANSPOSE(Header_Names)")

    # J/K/L/O: one spill each, written once. See the _*_SPILL_FORMULA
    # definitions above for why these are single dynamic arrays rather than
    # per-row formulas, and why they must go through `f` (Formula2).
    f(sheet, _FIRST_DATA_ROW, _C_PERIOD_IN_USE, _PERIOD_IN_USE_SPILL_FORMULA)
    f(sheet, _FIRST_DATA_ROW, _C_LEVELS, _LEVELS_SPILL_FORMULA)
    f(sheet, _FIRST_DATA_ROW, _C_REF_IN_USE, _REF_IN_USE_SPILL_FORMULA)
    f(sheet, _FIRST_DATA_ROW, _C_DESIGN_COLUMNS, _DESIGN_COLUMNS_SPILL_FORMULA)

    # The typed input columns keep one written cell per row: they are what
    # the user edits, so they cannot be spills. The profile supplies the
    # shipped defaults for the columns THIS dataset has; a retarget to a
    # wider table leaves the extra rows blank, which is a legal spec (Role
    # blank contributes nothing) and is exactly what the user then fills in.
    #
    # I (Sequence Period) is the typed override input: the user types a
    # number on the Sequence-flagged row to declare a Δ that differs from
    # the computed candidate. Left blank, the spec falls back to the
    # candidate. It is a PURE input — no formula here; the candidate lives
    # in J, and J's spill picks the override up by reading this band.
    for offset, variable in enumerate(profile.variables):
        row = _FIRST_DATA_ROW + offset
        role, include, ptype = profile.default_spec.get(variable, _FALLBACK_SPEC)
        val(sheet, row, _C_ROLE, role)
        val(sheet, row, _C_INCLUDE, include)
        val(sheet, row, _C_TYPE, ptype)
        # E (Reference) starts blank → first-in-sort-order default.
        # F (Order) starts blank; G (Transform) defaults to "None".
        val(sheet, row, _C_TRANSFORM, _DEFAULT_TRANSFORM)
        # H (Sequence): TRUE on the shipped ordering axis (Year), blank
        # elsewhere — zero-or-one flags is the legal range, and blank stays
        # a valid non-panel spec.
        if variable in profile.sequence_variables:
            val(sheet, row, _C_SEQUENCE, True)
        # M/N (the interaction pair) start blank: no interaction.

    _add_list_validation(sheet, _C_ROLE, _ROLE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_INCLUDE, _INCLUDE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_TYPE, _TYPE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_TRANSFORM, _TRANSFORM_VALIDATION_LIST)
    _add_list_validation(sheet, _C_SEQUENCE, _SEQUENCE_VALIDATION_LIST)
    _add_list_validation(
        sheet, _C_INTERACTION_TERM, _INTERACTION_TERM_VALIDATION_FORMULA
    )
    _add_list_validation(
        sheet,
        _C_INTERACTION_OPERATION,
        _INTERACTION_OPERATION_VALIDATION_LIST,
    )

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
    # "Hide in place" means the font color matches each band's own fill
    # (INPUT_COLOR for the input columns, white for the unfilled
    # computed-display cells) rather than a single muted gray — the same
    # font-matches-fill idiom used for the boundary guard on the Univariate
    # sheet's grid-search tables (write_sheet_univariate.py, cf.Font.Color =
    # 0xFFFFFF). Every range here runs out to _VALIDATION_LAST_ROW, not just
    # the rows currently in use, for the same reason the Spec_* bands and
    # the dropdown Validation do: a Source_Table retarget must not have to
    # reach outside a pre-applied range to find the rows it brought into
    # play. ONE ceiling across bands, validation and CF means a wider
    # dataset is fully formatted the instant it is retargeted, no rebuild.
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

    # ── Interaction pair (M/N) flags ─────────────────────────────────────
    # Rule ORDER is the priority order: FormatConditions.Add appends, and
    # the earlier rule wins, so the two error flags go in FIRST with
    # StopIfTrue and the hide-in-place rule last. Without that, a
    # non-blank M typed onto a non-Predictor row would be silently grayed
    # out by the hide rule instead of showing its error.
    _interaction_error_formats(sheet)

    # Cascading relevance, Role-keyed (continued): M–N hide in place
    # whenever Role ≠ Predictor, exactly like C–F. Both are
    # format_input-colored inputs, so the font matches INPUT_COLOR.
    add_expression_format(
        sheet,
        f"$M${_FIRST_DATA_ROW}:$N${_VALIDATION_LAST_ROW}",
        f'=$B{_FIRST_DATA_ROW}<>"{_ROLE_PREDICTOR}"',
        font_color=INPUT_COLOR,
    )
    # O (Design Columns) is an unfilled computed display, so it hides the
    # same way K and L do — white font on the white cell.
    add_expression_format(
        sheet,
        f"$O${_FIRST_DATA_ROW}:$O${_VALIDATION_LAST_ROW}",
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

    # ── The input band, sized to the dataset ─────────────────────────────
    # LAST, and deliberately so. Every rule above is added earlier and so
    # outranks this one, which means a red or yellow flag still wins on the
    # cells it applies to and the hide-in-place font rules still compose on
    # top (they set a font color, this sets a fill).
    #
    # This replaces the per-row format_input() calls the writer used to
    # make. Those painted exactly the profile's variables, so the input
    # surface was pinned to the build-time dataset in the same way the old
    # ListObject pinned the bands: retarget to a wider table and the rows it
    # brought into play were functional but unpainted. Keying the fill on
    # ROW() vs. COLUMNS(Source_Data) makes the visible input band track the
    # source table exactly, the same predicate the Spec_* bands use.
    for first_col, last_col in (
        (_C_ROLE, _C_SEQUENCE_PERIOD),
        (_C_INTERACTION_TERM, _C_INTERACTION_OPERATION),
    ):
        add_expression_format(
            sheet,
            f"${col_letter(first_col)}${_FIRST_DATA_ROW}:"
            f"${col_letter(last_col)}${_VALIDATION_LAST_ROW}",
            f"=ROW()-{_ROW_TO_COL_OFFSET}<=COLUMNS(Source_Data)",
            fill=INPUT_COLOR,
        )


def _write_spec_feedback(sheet: xw.Sheet) -> None:
    """The P/Q spectrum and the I1/I2 verdict overlay.

    Layout (cells land on row 1 for headers, row 2 for content; row 1 is
    shared with the row-1 audit strip on the right side of the sheet, and
    row 2 holds the row-2 Intercept control — both unaffected by the
    feedback cells to the right of E2):

        P1 = "Δ"          (bold header)
        Q1 = "Count"      (bold header)
        P2 = IFERROR(Sequence_Delta_Spectrum(), "")
            — spills downward into empty territory (P and Q are spec
              feedback columns with no other content, so the spill never
              collides with the spec block below row 3)
        I1 = "Verdict"    (bold header — overlays the Sequence_Period
                          column's row-1 cell, which is unused)
        I2 = combined switch formula (priority-ordered, single message)

    The E1 Sequence error status is written here too. H2 is the "Sequence"
    column header, and a status cell sitting on top of a header reads as a
    visual collision — so E1 it is. (It moved off H2 when the spec area
    became a structured table; that table is gone, the placement stands.)
    E1 keeps the same pattern as the old H2: blank while the spec is legal,
    a red error line when it is not.

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

    # P1/Q1: bold headers (no fill, default font size). The Verdict header
    # (I1) is bolded separately below; it lives in column I (the spec
    # block's Sequence_Period column) and so is NOT in the P:Q range.
    val(sheet, feedback_header_row, _C_FEEDBACK_DELTA, "Δ")
    val(sheet, feedback_header_row, _C_FEEDBACK_COUNT, "Count")
    bold_row(
        sheet,
        feedback_header_row,
        _C_FEEDBACK_DELTA,
        _C_FEEDBACK_COUNT,
    )

    # P2: Sequence_Delta_Spectrum() — an N×2 array of (delta, count) pairs,
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
    one row above the per-variable Include toggles. No engine formula consumes
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
    # already-demeaned Design_Columns()/Design_Response() pair is not a numerical error (the
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
    """The S/T derived-row zone: full-height label and mask spills.

    Row 1 of S/T is not written here — _write_audit_row owns the audit
    strip that occupies it.
    """
    sheet.range(rc(1, _C_GAP)).column_width = _GAP_COLUMN_WIDTH

    bold_row(sheet, _HEADER_ROW, _C_ROW_LABELS, _C_INCLUDED)
    val(sheet, _HEADER_ROW, _C_ROW_LABELS, "Row Labels")
    val(sheet, _HEADER_ROW, _C_INCLUDED, "Included")

    f(sheet, _FIRST_DATA_ROW, _C_ROW_LABELS, "=Row_Labels()")
    f(sheet, _FIRST_DATA_ROW, _C_INCLUDED, "=Sample_Include()")


def _write_audit_row(sheet: xw.Sheet) -> None:
    """Row-1 audit strip: bold label/value pairs from column S rightward.

    Values live in their own cells (not concatenated into the labels) so
    the QC analyzer can assert the numbers directly. The Predictor_Columns()-derived
    cells wrap IFERROR — an empty model makes DROP(built,,1) error, and
    the audit strip must degrade to the documented string, never leak a
    raw #CALC!. The two SUMPRODUCT counts are total functions over
    full-height inputs and cannot error, so they stay unwrapped.
    """
    audit_cells: tuple[tuple[str, str], ...] = (
        ("k", f"=IFERROR(COLUMNS(Predictor_Columns()),{_EMPTY_MODEL_FALLBACK})"),
        ("rows", f"=IFERROR(ROWS(Predictor_Columns()),{_EMPTY_MODEL_FALLBACK})"),
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
    """The V/W and Y/Z→ filtered display zones.

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
    # spills exactly COLUMNS(Predictor_Columns()) level-qualified names.
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
        (_C_MATRIX_START, "Predictor_Columns()"),
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



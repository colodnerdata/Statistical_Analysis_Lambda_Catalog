"""
spec_layout.py

Layout constants for the spec block -- column/row constants, role tokens,
transform tokens, validation lists, default specs, note text, status formulas,
and the shared helpers (_is_log, _set_spec_block_column_widths,
_set_spec_block_optional_outline_group).

These constants are the shared contract between the spec-block writers
(write_spec_block.py), the Regression sheet writer (write_sheet_regression.py),
the spec-block QC analyzers, the test-model sheet I/O, and the build scripts.

Extracted from write_spec_block.py to make the spec layout contract an
explicit, discoverable module rather than something buried at the top of a
2,455-line writer file.
"""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import (
    HEADER_COLOR,
    INPUT_COLOR,
)
from .workbook_helpers import (
    group_and_hide_columns,
    set_column_widths,
)

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

# ── The two rows above the spec block ─────────────────────────────────────────
# One grammar, applied without exception:
#
#   ROW 1 — LABELS for the readouts on row 2, and nothing else.
#   ROW 2 — the CONTROL, VALUE or STATUS itself, in the spec column it is about.
#
# "In the spec column it is about" is the rule that decides every placement
# here. The Role cardinality errors sit above Role, the Log-domain message above
# Transform, the Sequence cardinality error above Sequence, the design-matrix
# width guard above Design Columns. Before this, each status had been parked in
# whichever cell happened to be free — the Sequence error above Reference Level,
# the Fixed Effects error above Role, the width guard above Interaction Term —
# so nothing about a message's position told you what it was about, and there
# was no principled place to put the next one.
#
# Status cells get NO row-1 label. They are blank whenever the spec is legal, so
# a permanent label would be a caption for nothing most of the time; the message
# names its own subject when it appears. Only readouts — the Intercept toggle,
# the Fixed Effects trio, Σ Design Columns, the Δ spectrum — carry labels.
#
# The cost of putting each status in its own column is that row 2 has almost no
# horizontal runway: a long message has nowhere to overflow. Every status cell is
# therefore WrapText with a short, imperative message and a hover Note carrying
# the full guidance. Row 2 is left on automatic height (Rows(2).AutoFit() in
# write_sheet_regression), so it is one line tall while the spec is legal and
# grows the moment a message fires — which makes an error more prominent, not
# less.
_FEEDBACK_LABEL_ROW = 1
_FEEDBACK_STATUS_ROW = _INTERCEPT_ROW  # 2 — the controls and the toggle share it
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
    # 14, not 11: wide enough for the "Log (drop ≤ 0)" token to render in the
    # cell rather than only in the dropdown.
    _C_TRANSFORM: 14,
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

_DEFAULT_TRANSFORM = "None"

# The two Log tokens. They build the IDENTICAL constructed column — Ln(x), with
# Constructed_Column_Transforms() reporting "Log" for both — and differ in
# exactly one thing: what happens to a row whose value is zero or negative.
#
#   Log             the row stays in the sample and Ln_Positive returns #N/A,
#                   which propagates through the whole fit. Loud, and correct:
#                   the user asked to log a value that has no log.
#   Log (drop ≤ 0)  the row leaves the sample (Sample_Include grows a
#                   positivity term for this token only) and the count is
#                   reported at G2.
#
# Two tokens rather than one token that silently filters: dropping rows changes
# the sample the user is fitting, and that must be something they DECLARED, not
# something the workbook did on their behalf. Same "flag red and instruct, never
# silently switch" precedent as Intercept x Categorical and Categorical x Log.
#
# Because both report as "Log" to Constructed_Column_Transforms(), the
# (response_transform, predictor_transform) unit-space dispatcher gains no new
# combination and the Duan / back-transformation family is untouched — which is
# what keeps this from being the ~10x axis-widener MODEL_TESTING_ASSETS section 2
# warns every new Transform value about.
#
# These strings ALSO appear literally inside the catalog bodies in
# lambda_functions.json, which no import can reach; test_transform_tokens_match
# _the_catalog_bodies pins the two spellings together so a rename cannot
# half-land.
_TRANSFORM_LOG = "Log"
_TRANSFORM_LOG_DROP = "Log (drop ≤ 0)"


def _is_log(expr: str) -> str:
    """``OR(<expr>="Log",<expr>="Log (drop ≤ 0)")`` — "this row logs its column".

    Every Excel-side test of "is this a Log row?" goes through here rather than
    comparing against one token, so adding the second token could not leave a
    call site testing only the first. ``expr`` is any formula fragment that
    evaluates to a Transform cell's value.
    """
    return f'OR({expr}="{_TRANSFORM_LOG}",{expr}="{_TRANSFORM_LOG_DROP}")'

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
    "IFERROR(IF("
    + _is_log("INDEX(TAKE(Spec_Transform,n_c),p)")
    + ',"Ln("&h&")",h),"(none)"))'
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
_TRANSFORM_VALIDATION_LIST = ",".join(
    (_DEFAULT_TRANSFORM, _TRANSFORM_LOG, _TRANSFORM_LOG_DROP)
)
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
    "units.\n\n"
    "Zeros and negatives have no logarithm, and the two Log options differ "
    "only in what happens to those rows. “Log” keeps them in the "
    "sample, so the model returns #N/A rather than fit silently — the cell "
    "turns red and the message above this column names the variable and the "
    "row count. “Log (drop ≤ 0)” excludes them from the sample "
    "instead and reports how many it dropped. Both build the same "
    "“Ln(name)” column; excluding rows changes the sample you are "
    "fitting, so it is a choice you declare here rather than one the "
    "workbook makes for you."
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

# Count of Role="Response (y)" spec rows. Exactly one is the legal range —
# unlike Sequence and Fixed Effects, where zero is also legal — so the Role
# status line below flags both zero and two-plus. Same TAKE-trimmed idiom.
_RESPONSE_COUNT_FORMULA = (
    "SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))"
    f'="{_ROLE_RESPONSE}"))'
)

# "A Sequence axis is declared AND produced spacings" — the gate the Verdict
# label, the Δ spectrum headers and their hide-in-place rules all share, so a
# workbook with no ordering axis shows none of that machinery rather than
# showing it empty. Sequence_Deltas() returns #N/A with no axis; COUNT of an
# error is 0, so no IFERROR is needed.
_SEQUENCE_ACTIVE_FORMULA = "COUNT(Sequence_Deltas())>0"

# ── The row-2 status messages ─────────────────────────────────────────────────
# Each is one cell holding one message, picked in severity order by a nested IF
# — the idiom the Sequence verdict (I2) and the design-matrix width guard (O2)
# already use. Keeping the priority inside the formula rather than spreading it
# over several cells is what lets each column own exactly one status cell.
#
# Messages are short and imperative because they wrap inside a single column;
# the hover Note beside each carries the long form.

# B2 — Role cardinality. Exactly one Response is required (zero and two-plus are
# both errors, and they need different instructions), while Sequence and Fixed
# Effects allow zero, so those only flag the two-plus case. Fixed Effects is
# checked here rather than in its own cell because Role is the column all three
# conditions are declared in.
_ROLE_STATUS_FORMULA = (
    f"=IF({_RESPONSE_COUNT_FORMULA}=0,"
    '"ERROR: no Response (y) row — mark the variable being modeled.",'
    f"IF({_RESPONSE_COUNT_FORMULA}>1,"
    '"ERROR: multiple Response (y) rows — mark exactly one.",'
    f"IF({_FIXED_EFFECTS_COUNT_FORMULA}>1,"
    '"ERROR: multiple Fixed Effects rows — mark at most one.",'
    '"")))'
)
_ROLE_STATUS_NOTE = (
    "Role cardinality, in severity order.\n\n"
    "A model needs exactly one Response (y) row: with none there is nothing to "
    "fit, and with two the constructor silently takes the first, so both states "
    "are errors rather than warnings.\n\n"
    "Sequence and Fixed Effects each allow zero — plenty of models have no "
    "ordering axis and no panel structure — so only a second row of either is "
    "flagged. Two Fixed Effects rows would be two-way absorption, which this "
    "workbook does not implement; the engine absorbs the first row and ignores "
    "the second, which is why it is called out here rather than left to be "
    "discovered in the coefficients."
)

# H2 — Sequence cardinality, above the Sequence column it is declared in.
_SEQUENCE_STATUS_FORMULA = (
    f"=IF({_SEQUENCE_FLAG_COUNT_FORMULA}>1,"
    '"ERROR: multiple Sequence rows — mark at most one.","")'
)
_SEQUENCE_STATUS_NOTE = (
    "Zero or one Sequence flag is the legal range. Zero is an ordinary "
    "non-panel spec; one designates the ordering axis for the lag, difference "
    "and serial-correlation features; two or more is a spec error, because "
    "there is no defined answer to which axis a lag is taken along.\n\n"
    "The flagged cells in the Sequence column turn red at the same time, so "
    "this line tells you what is wrong and the column tells you where."
)

# G2 — the Log domain, above the Transform column. Two states, most severe
# first:
#
#   red    a row declares the strict "Log" token on a column that holds a zero
#          or a negative among the rows the model would fit. Those rows stay in
#          the sample, so Ln_Positive returns #N/A for each and the #N/A reaches
#          every statistic. The message names the variable, the count and the
#          fix, because the fix is a different dropdown value and the user has
#          no way to guess that from a sheet full of #N/A.
#   amber  "Log (drop ≤ 0)" is in use and actually excluded something. Not a
#          problem — it is the declared behaviour — but the sample is now
#          smaller than the data, and that must never be invisible.
#
# `bad` counts, per spec column, the non-positive rows a strict Log would poison;
# MAP over SEQUENCE(nc) gives one count per column, and XMATCH recovers which
# column the worst one came from. Sample_Include(FALSE) is the mask BEFORE the
# positivity layer — the rows the fit would otherwise have used — which is what
# makes both halves of this formula count the same population.
_LOG_DOMAIN_STATUS_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "rl,TAKE(Spec_Role,nc),"
    "inc,TAKE(Spec_Include,nc),"
    "typ,TAKE(Spec_Type,nc),"
    "trn,TAKE(Spec_Transform,nc),"
    "hdr,TOROW(Header_Names),"
    "base,Sample_Include(FALSE),"
    f'elig,((rl="{_ROLE_RESPONSE}")+((rl="{_ROLE_PREDICTOR}")*(inc=TRUE)'
    '*(typ="Continuous")))>0,'
    "bad,MAP(SEQUENCE(nc),LAMBDA(j,"
    f'IF(AND(INDEX(elig,j),INDEX(trn,j)="{_TRANSFORM_LOG}"),'
    "SUMPRODUCT(--base,--IFERROR((INDEX(Source_Data,0,j)+0)<=0,FALSE)),0))),"
    "worst,MAX(bad),"
    "IF(worst>0,"
    '"ERROR: "&INDEX(hdr,XMATCH(worst,bad))&" has "&worst&'
    f'" values ≤ 0 under Log — the fit is #N/A. Use {_TRANSFORM_LOG_DROP}.",'
    "LET(d,SUMPRODUCT(N(Sample_Include(FALSE)))-SUMPRODUCT(N(Sample_Include())),"
    'IF(d=0,"",d&" rows excluded: Log of ≤ 0"))))'
)
_LOG_DOMAIN_STATUS_NOTE = (
    "What the Log transforms are doing to the sample.\n\n"
    "RED — a variable declaring “Log” contains zeros or negatives on rows "
    "the model would fit. A non-positive number has no logarithm, so "
    "Ln_Positive returns #N/A for those rows, and because they stay in the "
    "sample the #N/A propagates into every coefficient, statistic and residual "
    "on the sheet. Nothing is salvageable until it is resolved. Either exclude "
    "those rows by switching that variable to “Log (drop ≤ 0)”, "
    "declare a Filter column that removes them, or drop the transform.\n\n"
    "AMBER — “Log (drop ≤ 0)” is declared somewhere and has excluded "
    "rows. That is what it is for; the count is here so a shrinking sample is "
    "never something you have to notice on your own. Compare it against "
    "Observations in the Regression Statistics block.\n\n"
    "Blank means neither applies: either no Log is declared, or every logged "
    "column is strictly positive throughout."
)

# I2 — how the declared Sequence axis is actually spaced.
_SPACING_VERDICT_NOTE = (
    "How the Sequence axis is spaced, compared against the period in use.\n\n"
    "RED “off-grid” — some spacings are not whole multiples of the "
    "period, so the data does not sit on the grid the spec declares. Lags and "
    "differences computed along it will pair the wrong observations.\n\n"
    "RED “calendar” — the spacings look like month, quarter or year "
    "day-counts (28-31, 90-92, 365-366). Calendar arithmetic is not a constant "
    "period; add an integer period index upstream and sequence on that "
    "instead.\n\n"
    "AMBER “not evenly spaced” — the axis has spacings besides the "
    "period. Everything still computes, but a gap is silently treated as one "
    "step by any lag taken along it.\n\n"
    "AMBER “no natural period” — no spacing occurs more often than "
    "the others, so there is no modal value to adopt and the candidate falls "
    "back to the minimum. Type a Sequence Period on the flagged row if that "
    "is not what you want.\n\n"
    "Blank means either no axis is declared or its spacing is uniform."
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
    "IFERROR("
    + _is_log(
        "INDEX(TAKE(Spec_Transform,COLUMNS(Source_Data)),"
        f'XMATCH("{_ROLE_RESPONSE}",TAKE(Spec_Role,COLUMNS(Source_Data))))'
    )
    + ",FALSE)"
)

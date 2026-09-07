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
validated, and flagged on the sheet, and built into the design matrix by
the constructor's interaction wiring (the mate() closure reading the pair
through Predictor_Columns / Constructed_Column_Names).
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
    included rows = SUMPRODUCT(--Fit_Sample_Include()) ·
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
MileageData, 11 rows: [MPG]..[Model?]). Two axes:

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
Full-height contract: ROWS(Predictor_Columns()) = ROWS(Row_Labels()) =
ROWS(Sample_Include()) = 406 always — the constructor reads the mask ONLY
to fix level sets; nothing here ever row-filters. With the real mask live,
the T0 mask-dependent values are real on the sheet: k = 16 (2 continuous +
2 Origin dummies + 12 Model Year dummies), and
SUMPRODUCT(--Fit_Sample_Include()) = 392 in the shipped no-Log-drop default (completeness-only on the response
and the two continuous predictors; Auto MPG ships no Filter-by-default
column, so the active Filter role is exercised only by the Is_USA fixture).

Not here (deliberately, per release scoping): the QC analyzer
(analyze_model_construction.py) and the Version History / CHANGELOG bump
to v2.0 — those landed in the final wiring PR of that release.
"""
from __future__ import annotations

from collections.abc import Sequence

import xlwings as xw

from .catalog_schema import CatalogFunction
from .lambda_formula_parser import (
    _normalize_user_formula,
    _strip_non_string_whitespace,
)
from .regression_materialization import qualify_spill_reader_references
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

# ── Spec-block layout constants ────────────────────────────────────────────────
# Column/row constants, role tokens, transform tokens, validation lists, note
# text, status formulas, and the shared helpers (_is_log,
# _set_spec_block_column_widths, _set_spec_block_optional_outline_group) live
# in spec_layout.py and are re-exported here so existing importers keep
# working. See that module for the full documentation.
from .spec_layout import (  # noqa: F401  — re-exported for importers
    _AUDIT_PAIRS,
    _AUDIT_ROW,
    _CLOSURE_SCOPE,
    _C_BREAK_LEFT,
    _C_BREAK_MID,
    _C_DESIGN_COLUMNS,
    _C_FEEDBACK_COUNT,
    _C_FEEDBACK_DELTA,
    _C_FILTERED_LABELS,
    _C_FILTERED_Y,
    _C_GAP,
    _C_INCLUDE,
    _C_INCLUDED,
    _C_INTERACTION_OPERATION,
    _C_INTERACTION_TERM,
    _C_LABEL,
    _C_LEVELS,
    _C_MATRIX_LABELS,
    _C_MATRIX_START,
    _C_ORDER,
    _C_PERIOD_IN_USE,
    _C_REFERENCE,
    _C_REF_IN_USE,
    _C_ROLE,
    _C_ROW_LABELS,
    _C_SEQUENCE,
    _C_SEQUENCE_PERIOD,
    _C_SPEC_LAST,
    _C_TRANSFORM,
    _C_TYPE,
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _DEFAULT_TRANSFORM,
    _DESIGN_COLUMNS_NOTE,
    _EMPTY_MODEL_FALLBACK,
    _FALLBACK_SPEC,
    _FEEDBACK_LABEL_ROW,
    _FEEDBACK_STATUS_ROW,
    _FIRST_DATA_ROW,
    _FIXED_EFFECTS_COUNT_FORMULA,
    _FIXED_EFFECTS_NAME_FORMULA,
    _GAP_COLUMN_WIDTH,
    _HEADER_ROW,
    _INCLUDE_NOTE,
    _INCLUDE_VALIDATION_LIST,
    _INTERACTION_DIFFERENCE,
    _INTERACTION_HEADER_SYMBOLS,
    _INTERACTION_HEADER_UNKNOWN,
    _INTERACTION_OPERATION_NOTE,
    _INTERACTION_OPERATION_VALIDATION_LIST,
    _INTERACTION_PRODUCT,
    _INTERACTION_RATIO,
    _INTERACTION_SYMMETRIC_OPERATIONS,
    _INTERACTION_TERM_NOTE,
    _INTERACTION_TERM_VALIDATION_FORMULA,
    _INTERCEPT_ROW,
    _LABEL_NOTE,
    _LAST_DATA_ROW,
    _LEVELS_NOTE,
    _LOG_DOMAIN_STATUS_NOTE,
    _N_VARIABLES,
    _PERIOD_IN_USE_NOTE,
    _REFERENCE_NOTE,
    _REF_IN_USE_NOTE,
    _RESERVED_NOTE,
    _RESPONSE_COUNT_FORMULA,
    _RESPONSE_LOG_FORMULA,
    _RESPONSE_NAME_FORMULA,
    _ROLE_FILTER,
    _ROLE_FIXED_EFFECTS,
    _ROLE_IDENTIFIER,
    _ROLE_NOTE,
    _ROLE_OMIT,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    _ROLE_STATUS_NOTE,
    _ROLE_VALIDATION_LIST,
    _ROW_TO_COL_OFFSET,
    _SEQUENCE_ACTIVE_FORMULA,
    _SEQUENCE_FLAG_COUNT_FORMULA,
    _SEQUENCE_NOTE,
    _SEQUENCE_PERIOD_NOTE,
    _SEQUENCE_STATUS_NOTE,
    _SEQUENCE_VALIDATION_LIST,
    _SPACING_VERDICT_NOTE,
    _SPEC_BAND_LAST_ROW,
    _SPEC_COLUMN_WIDTHS,
    _SPEC_OPTIONAL_FIRST_COL,
    _SPEC_OPTIONAL_LAST_COL,
    _TRANSFORM_LOG,
    _TRANSFORM_LOG_DROP,
    _TRANSFORM_NOTE,
    _TRANSFORM_VALIDATION_LIST,
    _TYPE_NOTE,
    _TYPE_VALIDATION_LIST,
    _VALIDATION_LAST_ROW,
    _VARIABLES,
    _XL_BETWEEN,
    _XL_VALIDATE_LIST,
    _XL_VALID_ALERT_STOP,
    _is_log,
    _set_spec_block_column_widths,
    _set_spec_block_optional_outline_group,
)

# ── Per-dataset spec profiles ──────────────────────────────────────────────
# SpecDatasetProfile and the three dataset profiles (Auto MPG, Life Expectancy,
# Production Lots) live in spec_dataset_profiles.py and are re-exported here
# so existing importers keep working. See that module for the full documentation.
from .spec_dataset_profiles import (  # noqa: F401  — re-exported for importers
    LIFE_EXPECTANCY,
    MILEAGE,
    PRODUCTION_LOTS,
    _AUTO_MPG_PROFILE,
    _LIFE_EXPECTANCY_DEFAULT_SPEC,
    _LIFE_EXPECTANCY_PROFILE,
    _LIFE_EXPECTANCY_SEQUENCE_VARIABLES,
    _LIFE_EXPECTANCY_VARIABLES,
    _LIFE_TALK_DEMO_PREDICTORS,
    _PRODUCTION_LOTS_DEFAULT_SPEC,
    _PRODUCTION_LOTS_PROFILE,
    _PRODUCTION_LOTS_SEQUENCE_VARIABLES,
    _PRODUCTION_LOTS_VARIABLES,
    SPEC_DATASET_PROFILES,
    SpecDatasetProfile,
)

# ── The four computed spec columns (J, K, L, O) ────────────────────────────
#
# Each is a SINGLE dynamic-array formula written once at _FIRST_DATA_ROW,
# spilling one value per source-table column. This is the mechanism that
# makes a Source_Table retarget resize the spec block.
#
# They are single spills rather than per-row formulas, so a Source_Table
# retarget resizes them automatically — a fixed row-by-row layout would pin
# the block to the row count baked in at build time and leave the computed
# columns with no formulas in any rows a wider dataset added. A spill cannot
# live inside a ListObject, and J/K/L sit between the input columns I and M
# while O sits after N, so the only way to keep them self-sizing is to leave
# them as spills with no surrounding ListObject.
#
# Consequences worth knowing when editing them:
#   * Each value is INDEX(<band>,i) with `i` from the MAP, so a formula does
#     not depend on which row it occupies.
#   * They are written with `f` (Formula2), never `f_structured`
#     (Formula) — Formula enters an array formula as a legacy CSE range,
#     which does NOT resize on retarget and would silently pin the block
#     to the build-time row count.
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
# count L over the fitted rows, with Dummy_Levels' blank
# normalization mirrored inline. Deliberately NOT a Dummy_Levels call: the
# display must show L (including 1 for a degenerate column, feeding the red
# CF), while Dummy_Levels returns the L−1 retained levels and #N/A when
# degenerate. IFERROR -> 0 covers the empty-masked-sample edge.
_LEVELS_SPILL_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "rl,TAKE(Spec_Role,nc),"
    "typ,TAKE(Spec_Type,nc),"
    "si,Fit_Sample_Include(),"
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
# fitted sample. Deliberately NOT a Dummy_Levels call: that function
# returns the RETAINED levels, which is the set the reference has been
# dropped from. IFERROR -> "" covers the empty-masked-sample edge (K shows 0
# and flags red there).
_REF_IN_USE_SPILL_FORMULA = (
    "=LET(nc,COLUMNS(Source_Data),"
    "rl,TAKE(Spec_Role,nc),"
    "typ,TAKE(Spec_Type,nc),"
    "refs,TAKE(Spec_Reference,nc),"
    "si,Fit_Sample_Include(),"
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
# Fit_Sample_Include(), TOROW(Header_Names) and the kk helper itself. The
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
    "si,Fit_Sample_Include(),"
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
    edit.** The bands are ``TAKE``-trimmed to the live source-table width
    rather than fixed-width references, which would lock them at the row
    count baked in at build time. Retargeting to a wider table would then
    leave every band short: ``TAKE`` does not pad, so ``INDEX(rl, n_c)``
    would run off the end and every engine cell downstream would read as an
    error. Excel offers no formula-driven way to resize a fixed table and
    this workbook is macro-free, so the bands size themselves instead.

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
    spec columns are spill formulas that reference these band names.
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
        # added by the layout-break MAJOR. The first two are live inputs —
        # bound so the grid shape is final, read by the constructor's
        # interaction wiring to build Product/Difference/Ratio columns
        # (and by the pair's conditional-format rules). The third is
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

    # The Name Manager is the surface a user reads to learn what a name is
    # for — or to retarget Source_Table at another dataset — so every wiring
    # name carries a brief purpose comment there, mirroring the chart-range
    # convention in write_sheet_regression's _setup_local_names.
    local_name_comments: dict[str, str] = {
        "Source_Table": (
            "The sheet's data source: the source Excel Table (header + all "
            "rows). Edit this one name to retarget the sheet at a different "
            "dataset."
        ),
        "Source_Data": "Data body of Source_Table (header row dropped).",
        "Header_Names": "Header row of Source_Table.",
        "Spec_Role": (
            "Spec block input band: the Role column (Response / Predictor / "
            "Identifier / Filter / Omit / Fixed Effects)."
        ),
        "Spec_Include": "Spec block input band: the Include toggle column.",
        "Spec_Type": (
            "Spec block input band: the Type column (Continuous / "
            "Categorical)."
        ),
        "Spec_Reference": (
            "Spec block input band: the categorical Reference-level column."
        ),
        "Spec_Order": (
            "Spec block band: the Order column — reserved, bound but not "
            "read by any formula yet."
        ),
        "Spec_Transform": (
            "Spec block input band: the Transform column (None / Log / "
            "Log (drop ≤ 0))."
        ),
        "Spec_Sequence": (
            "Spec block band: the Sequence structural flag column (H)."
        ),
        "Spec_Sequence_Period": (
            "Spec block input band: the Sequence Period typed override (I)."
        ),
        "Spec_Period_In_Use": (
            "Spec block band: the Period In Use display (J) — the candidate "
            "Δ with the typed override applied."
        ),
        "Spec_Interaction_Term": (
            "Spec block input band: the Interaction Term column (M)."
        ),
        "Spec_Interaction_Operation": (
            "Spec block input band: the Interaction Operation column (N)."
        ),
        "Spec_Design_Columns": (
            "Spec block band: the Design Columns audit display (O); only "
            "the width guard reads it."
        ),
        "Allow_Intercept": (
            "Model-level Intercept toggle — the row-2 control cell in the "
            "Include column (C2)."
        ),
    }
    assert set(local_names) == set(local_name_comments), (
        "every spec-wiring name needs a Name Manager comment"
    )

    # Wiring first: Excel resolves each name against the ones already added,
    # and the constructor closures below reference Source_Data / Spec_*.
    for name, refers_to in local_names.items():
        drop_local_name(sheet, name)
        _nm = sheet.api.Names.Add(Name=name, RefersTo=refers_to)
        _nm.Comment = local_name_comments[name]

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
        # Qualify the spill-reader references (Fit_Sample_Include, ...) per
        # sheet: those names do not exist yet at this point (the
        # materialization zone creates them later in the sheet's build), so an
        # unqualified reference would resolve — at calculation, not here —
        # against the workbook's whole name collection and pin to a foreign
        # sheet's copy in any multi-Regression-sheet workbook. See
        # SPILL_READER_NAMES in regression_materialization.py.
        refers_to = qualify_spill_reader_references(refers_to, sheet.name)
        drop_local_name(sheet, closure.name)
        _nm = sheet.api.Names.Add(Name=closure.name, RefersTo=refers_to)
        # The catalog entry's notes ARE the Name Manager comment: one
        # description of the closure's purpose, shared with the
        # LAMBDA_functions sheet's tooltip column.
        _nm.Comment = closure.notes


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
    (``SPEC_DATASET_PROFILES["auto_mpg"]``) when omitted.

    **The block has no fixed height.** The profile decides which rows get
    shipped *defaults*, not how many rows exist: the variable-name column,
    the four computed columns and the input band's fill all size themselves
    from ``COLUMNS(Source_Data)``, so retargeting ``Source_Table`` resizes
    the block. A fixed-height layout sized to ``len(profile.variables)``
    would pin it to the build-time dataset, which is exactly what the
    self-sizing bands avoid.

    Must run AFTER ``_set_sheet_scoped_names``: the computed columns are
    spills that reference the ``Spec_*`` bands, ``Source_Data`` and the
    constructor closures, which must exist first.
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

    # The specification header row. Nothing overrides it after the defaults
    # are written — the fill and font are set once here and stay. Applying a
    # TableStyle would silently replace the header fill and font, which is
    # why the block carries no ListObject.
    # test_spec_block_prefills_the_t0_default_configuration still asserts
    # all three properties (fill, color, bold).
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
            f"{_is_log(f'$G{_FIRST_DATA_ROW}')})"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Log-domain flag: red G when a row declares the STRICT "Log" token on a
    # column that actually contains a zero or a negative among the rows the
    # model would fit. Those rows stay in the sample by design, so Ln_Positive
    # returns #N/A for each one and the #N/A propagates through
    # Predictor_Columns() into every statistic on the sheet — a fit that is
    # dead, not degraded. The cell is where the user can act on it, and the G2
    # message beside it names the variable, the count, and the fix.
    #
    # "Log (drop ≤ 0)" never fires this rule. It is a correct declaration whose
    # consequence — a narrower sample — is reported at G2 in amber instead.
    #
    # The row test mirrors the transform eligibility branch exactly:
    # only the Response and included Continuous Predictors reach Ln_Positive,
    # so a Log left on an Identifier or an excluded row is inert and unflagged.
    # Sample_Include() is the mask BEFORE the positivity layer, which is
    # what makes this count the rows the fit would otherwise have used.
    # Calling a closure with INDEX(Source_Data,0,ROW()-offset) from inside a CF
    # expression is the same idiom the invalid-reference rule below uses.
    add_expression_format(
        sheet,
        f"$G${_FIRST_DATA_ROW}:$G${_VALIDATION_LAST_ROW}",
        (
            f'=AND($G{_FIRST_DATA_ROW}="{_TRANSFORM_LOG}",'
            f'OR($B{_FIRST_DATA_ROW}="{_ROLE_RESPONSE}",'
            f'AND($B{_FIRST_DATA_ROW}="{_ROLE_PREDICTOR}",'
            f"$C{_FIRST_DATA_ROW}=TRUE,"
            f'$D{_FIRST_DATA_ROW}="Continuous")),'
            "SUMPRODUCT(--Fit_Sample_Include(),"
            "--IFERROR((INDEX(Source_Data,0,"
            f"ROW()-{_ROW_TO_COL_OFFSET})+0)<=0,FALSE))>0)"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # Invalid-reference flag: red E when a nonblank reference makes
    # Dummy_Levels fail on the fitted sample — the constructor's exact skip condition, tested
    # directly rather than via level-set membership (a membership test
    # against the returned set would false-positive on the default
    # reference itself, which Dummy_Levels excludes from its output).
    add_expression_format(
        sheet,
        f"$E${_FIRST_DATA_ROW}:$E${_VALIDATION_LAST_ROW}",
        (
            f'=AND($E{_FIRST_DATA_ROW}<>"",'
            f"ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-{_ROW_TO_COL_OFFSET}),"
            f"$E{_FIRST_DATA_ROW},Fit_Sample_Include())))"
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
    # The input-band fill is a single lowest-priority CF rule keyed on
    # ROW() vs. COLUMNS(Source_Data), so the visible input band tracks the
    # source table exactly — the same predicate the Spec_* bands use — and
    # retargeting to a wider table paints the rows it brings into play
    # instead of leaving them functional but unpainted. A fixed per-row
    # paint would pin the input surface to the build-time dataset the same
    # way a fixed-width band would.
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


def _status_cell(
    sheet: xw.Sheet,
    col: int,
    formula: str,
    note: str,
    *,
    label: str,
) -> str:
    """Write one row-2 status cell and return its ``$G$2``-style address.

    Every status cell in this band is built the same way: bold, WrapText (row 2
    is on automatic height, so a message that fires makes the row grow rather
    than truncating against its neighbour), and a hover Note carrying the long
    form the cell itself has no width for. The caller adds the conditional
    formatting, because severity differs per status.
    """
    f(sheet, _FEEDBACK_STATUS_ROW, col, formula)
    bold(sheet, _FEEDBACK_STATUS_ROW, col)
    cell = sheet.range(rc(_FEEDBACK_STATUS_ROW, col))
    try:
        cell.api.WrapText = True
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    _set_note(sheet, _FEEDBACK_STATUS_ROW, col, note, label=label)
    return f"${col_letter(col)}${_FEEDBACK_STATUS_ROW}"


def _hide_when(sheet: xw.Sheet, address: str, condition: str) -> None:
    """White-on-white a label or readout whose feature is not in play.

    The font-matches-fill idiom the spec rows already use for cascading
    relevance, applied to the band above them: rather than showing "n/a" or an
    empty labelled cell, an inactive readout and its label both disappear. The
    cells here have no fill, so white font on white cell is the match.
    """
    add_expression_format(
        sheet,
        address,
        f"={condition}",
        font_color=(255, 255, 255),
    )


def _write_spec_feedback(sheet: xw.Sheet) -> None:
    """Rows 1-2 above the spec block: labels on row 1, status/readouts on row 2.

    See the ``_FEEDBACK_LABEL_ROW`` comment block for the grammar. What this
    function writes, left to right:

        B2  Role cardinality status      (red)      — above Role
        G2  Log domain status            (red/amber)— above Transform
        I1  "Spacing Verdict" label                 — hidden with no axis
        I2  Sequence spacing verdict     (red/amber)— above Sequence Period
        J1  "FE Variable"   / J2 value              — hidden with no FE row
        K1  "FE Groups"     / K2 value              — hidden with no FE row
        L1  "FE df Absorbed"/ L2 value              — hidden with no FE row
        P3  "Δ" / Q3 "Count" labels                 — hidden with no axis
        P4  IFERROR(Sequence_Delta_Spectrum(), "")  — spills down P:Q

    The Sequence cardinality error (H2) is ``_write_sequence_status``, the
    Intercept label and toggle (C1/C2) are ``_write_intercept_control``, and the
    Σ Design Columns total (N1/O1) with its width guard (O2) belong to the
    Regression sheet writer, which owns the design-matrix layout the guard's
    thresholds are derived from.

    The Δ spectrum sits on rows 3+ rather than rows 1-2 so its header aligns
    with the spec block's own header row and its body with the spec data rows —
    it reads as a second table beside the first instead of a third row-2 thing.
    Moving it down also gives O2's width-guard message P2:Q2 to overflow into,
    which is the only runway anything on row 2 has.
    """
    # ── B2: Role cardinality ────────────────────────────────────────────────
    role_status = _status_cell(
        sheet,
        _C_ROLE,
        "=Role_Status()",
        _ROLE_STATUS_NOTE,
        label="Role status",
    )
    add_expression_format(
        sheet,
        role_status,
        f'={role_status}<>""',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )

    # ── G2: the Log domain ──────────────────────────────────────────────────
    # Red first with StopIfTrue so a poisoned strict-Log column outranks the
    # amber excluded-row count, which is otherwise true at the same time (a
    # spec can declare both tokens on different variables). Keyed on the
    # message's own leading token, the same way the width guard is, so the
    # formula above stays the single source of which state is which.
    log_status = _status_cell(
        sheet,
        _C_TRANSFORM,
        "=Log_Domain_Status()",
        _LOG_DOMAIN_STATUS_NOTE,
        label="Log domain",
    )
    add_expression_format(
        sheet,
        log_status,
        f'=ISNUMBER(SEARCH("ERROR",{log_status}))',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
        stop_if_true=True,
    )
    add_expression_format(
        sheet,
        log_status,
        f'={log_status}<>""',
        fill=CF_YELLOW_FILL,
        font_color=CF_DARK_YELLOW_TEXT,
    )

    # ── J1:L2 — the Fixed Effects readouts ──────────────────────────────────
    # All three key off the same FE-count gate and return "" so that with
    # the labels hidden alongside them, an inactive block leaves no trace
    # instead of three cells of filler. They
    # still resolve the FIRST FE row in the 2-plus-rows error state, exactly
    # like Fixed_Effects_Column() itself — B2 above is what flags that state,
    # not these display cells.
    for col, label, value in (
        (
            _C_PERIOD_IN_USE,
            "FE Variable",
            f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"",{_FIXED_EFFECTS_NAME_FORMULA})',
        ),
        (
            _C_LEVELS,
            "FE Groups",
            f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"",'
            "Absorbed_Degrees_Of_Freedom()+1)",
        ),
        (
            _C_REF_IN_USE,
            "FE df Absorbed",
            f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"",'
            "Absorbed_Degrees_Of_Freedom())",
        ),
    ):
        val(sheet, _FEEDBACK_LABEL_ROW, col, label)
        bold(sheet, _FEEDBACK_LABEL_ROW, col)
        f(sheet, _FEEDBACK_STATUS_ROW, col, value)
        _hide_when(
            sheet,
            f"${col_letter(col)}${_FEEDBACK_LABEL_ROW}:"
            f"${col_letter(col)}${_FEEDBACK_STATUS_ROW}",
            f"{_FIXED_EFFECTS_COUNT_FORMULA}=0",
        )

    # ── P3/Q3 + P4: the Δ spectrum ──────────────────────────────────────────
    # Headers on the spec block's own header row, body from its first data row.
    # P and Q carry no other content, so the N×2 spill has the whole column
    # below it and never collides with anything.
    val(sheet, _HEADER_ROW, _C_FEEDBACK_DELTA, "Δ")
    val(sheet, _HEADER_ROW, _C_FEEDBACK_COUNT, "Count")
    bold_row(sheet, _HEADER_ROW, _C_FEEDBACK_DELTA, _C_FEEDBACK_COUNT)
    _hide_when(
        sheet,
        f"${col_letter(_C_FEEDBACK_DELTA)}${_HEADER_ROW}:"
        f"${col_letter(_C_FEEDBACK_COUNT)}${_HEADER_ROW}",
        f"NOT({_SEQUENCE_ACTIVE_FORMULA})",
    )
    # IFERROR degrades the no-axis / no-spacings #N/A to a quiet blank.
    f(
        sheet,
        _FIRST_DATA_ROW,
        _C_FEEDBACK_DELTA,
        '=IFERROR(Sequence_Delta_Spectrum(),"")',
    )

    # ── I1/I2: the spacing verdict ──────────────────────────────────────────
    # "Spacing Verdict", not "Verdict": the sheet has several verdicts now, and
    # this one is specifically about how the Sequence axis is spaced. Hidden
    # with the rest of the sequence machinery when no axis is declared.
    val(sheet, _FEEDBACK_LABEL_ROW, _C_SEQUENCE_PERIOD, "Spacing Verdict")
    bold(sheet, _FEEDBACK_LABEL_ROW, _C_SEQUENCE_PERIOD)
    _hide_when(
        sheet,
        f"${col_letter(_C_SEQUENCE_PERIOD)}${_FEEDBACK_LABEL_ROW}",
        f"NOT({_SEQUENCE_ACTIVE_FORMULA})",
    )

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
    verdict_formula = (
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
    verdict_cell = _status_cell(
        sheet,
        _C_SEQUENCE_PERIOD,
        verdict_formula,
        _SPACING_VERDICT_NOTE,
        label="Spacing Verdict",
    )

    # I2 CF: red for off-grid or calendar (StopIfTrue outranks yellow);
    # yellow for regularity or no-natural. Each rule keys on a SEARCH
    # of the cell's rendered text for the message keyword — the same
    # four message constants _MSG_* used to build the formula.
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
    """H2 — the Sequence cardinality status, above the Sequence column.

    Zero-or-one Sequence flags is the legal range: zero is a valid spec
    (non-panel data), one designates the ordering axis, two-plus is a spec
    error. The per-cell red CF on the flagged H cells (added in
    ``_write_spec_block``) points at the offending rows while this line says
    what is wrong.

    This cell is the single Sequence status message and lives in the
    Sequence column itself, the column it is about — keep it as the only
    copy rather than duplicating it in an unrelated column (E1 carries none).
    """
    status_cell = _status_cell(
        sheet,
        _C_SEQUENCE,
        "=Sequence_Status()",
        _SEQUENCE_STATUS_NOTE,
        label="Sequence status",
    )
    add_expression_format(
        sheet,
        status_cell,
        f'={status_cell}<>""',
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_intercept_control(sheet: xw.Sheet) -> None:
    """C1/C2 — the model-level Intercept label and ``Allow_Intercept`` toggle.

    The toggle sits at the top of the Include column, one row above the
    per-variable Include toggles, because that is what it is: the intercept's
    Include cell. Its label is at C1, directly above the thing it names, like
    every other labelled readout in this band.

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
    val(sheet, _FEEDBACK_LABEL_ROW, _C_INCLUDE, "Intercept")
    bold(sheet, _FEEDBACK_LABEL_ROW, _C_INCLUDE)
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
    val(sheet, _HEADER_ROW, _C_INCLUDED, "Eligible")

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
        ("included rows", "=SUMPRODUCT(--Fit_Sample_Include())"),
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
                f"=IFERROR(FILTER({source},Fit_Sample_Include()),"
                f"{_EMPTY_MODEL_FALLBACK})"
            ),
        )



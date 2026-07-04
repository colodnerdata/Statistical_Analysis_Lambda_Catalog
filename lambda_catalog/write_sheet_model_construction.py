"""Build the Model Construction worksheet — v3.0 declarative specification block.

Two-axis specification (ROADMAP: v3.0 — Specification-Driven Regression):

    A         B      C        D     E                F        G          H
    Variable  Role   Include  Type  Reference Level  Order    Transform  Levels
    (spill)   (drop) (input)  (drop)(input)          (rsvd.)  (rsvd.)    (disp.)

Right of the spec block, after a narrow gap column I (which also visually
reserves the future Design Columns audit column):

    J             K        L     M           N            O     P           Q →
    Row Labels    Included (brk) Filt.Labels Filt.y       (brk) Filt.Labels Filtered X_s
    (=Row_Labels() spill at J3; =Sample_Include() spill at K3 — both
     full-height, never internally filtered. M/N/P/Q are the FILTERED
     display zones: the only place on the sheet where Sample_Include()
     row-filters anything. P repeats the filtered labels so the matrix
     reads side-by-side without scrolling back to M.)

Row 1, from column J rightward, holds the bold audit cells as
label/value pairs (values on the non-narrow columns K/N/Q/S/U):

    k = COLUMNS(X_s()) · rows = ROWS(X_s()) · response = <derived name> ·
    responses = <count of Role="Response"> (red CF when <> 1) ·
    included rows = SUMPRODUCT(N(Sample_Include()))

Row 2 above Q carries the =Constructed_Column_Names() header strip
(level-qualified names, horizontal). Every spill formula in the filtered
zones wraps IFERROR(..., "(empty model)") so an empty model degrades to a
documented string, never a raw #CALC! leak.

The spec spans EVERY column of the LifeExpectancyData table (23 rows:
[Country]..[Schooling] plus [Full_Data]). Two axes:

    Variable Role  — Response | Predictor | Identifier | Filter | Omit
                     (what the column IS; future: Fixed Effects/Weight/Time)
    Predictor Type — Continuous | Categorical
                     (how a Predictor ENTERS; meaningful only when
                      Role = Predictor; this axis never grows)

Columns F (Order) and G (Transform) are reserved for a future release: they
are styled as inputs and carry sheet-scoped names (Spec_Order, Spec_Transform)
so the grid shape is final, but no formula reads them yet.

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
    All_Xs                     → Source_Data: one name wrapping the table
                                 reference (the dataset-retarget point;
                                 structured refs can't be parameterized
                                 without volatile INDIRECT)

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
    remaining numerics → Predictor/Continuous/FALSE (candidates)
    Full_Data        → Filter                 (ANDed into Sample_Include()
                                               with role-aware completeness)
Full-height contract: ROWS(X_s()) = ROWS(Row_Labels()) =
ROWS(Sample_Include()) = 2938 always — the constructor reads the mask ONLY
to fix level sets; nothing here ever row-filters. With the real mask live,
the T0 mask-dependent values are real on the sheet: k = 19 (15 Year
dummies), SUMPRODUCT(N(Sample_Include())) = 1649.

Not here (deliberately, per release scoping): the QC analyzer
(analyze_model_construction.py) and the Version History / CHANGELOG bump
to v3.0 — those land in the final wiring PR.
"""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

from .sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_LIGHT_RED_FILL,
    MUTED_TEXT_COLOR,
)
from .workbook_helpers import (
    add_expression_format,
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

_HEADER_ROW = 2
_FIRST_DATA_ROW = 3
_LAST_DATA_ROW = _FIRST_DATA_ROW + _N_VARIABLES - 1  # 25

# Spec-block columns (1-based). Role precedes Include: the larger
# declaration comes first (dataset semantics before iteration state).
# F/G are the reserved Order/Transform slots; H is the computed display.
(
    _C_LABEL,
    _C_ROLE,
    _C_INCLUDE,
    _C_TYPE,
    _C_REFERENCE,
    _C_ORDER,
    _C_TRANSFORM,
    _C_LEVELS,
) = range(1, 9)

# Derived-row zone right of the spec block. I is a narrow gap (and the
# visual reservation for the future Design Columns audit column); J and K
# hold the full-height Row_Labels() / Sample_Include() spills.
_C_GAP = 9
_C_ROW_LABELS = 10
_C_INCLUDED = 11
_GAP_COLUMN_WIDTH = 2

# Filtered display zone: the ONLY place Sample_Include() row-filters
# anything (everything left of L honors the full-height contract). L and
# O are narrow visual breaks; P repeats the filtered labels so the matrix
# reads side-by-side without scrolling back to M.
_C_BREAK_LEFT = 12
_C_FILTERED_LABELS = 13
_C_FILTERED_Y = 14
_C_BREAK_MID = 15
_C_MATRIX_LABELS = 16
_C_MATRIX_START = 17

# Row-1 audit strip: label/value pairs marching right from column J,
# values placed on the non-narrow columns (K, N, Q, S, U) so no number
# lands on a width-2 break column.
_AUDIT_ROW = 1
_AUDIT_PAIRS: tuple[tuple[int, int], ...] = (
    (_C_ROW_LABELS, _C_INCLUDED),          # k
    (_C_FILTERED_LABELS, _C_FILTERED_Y),   # rows
    (_C_MATRIX_LABELS, _C_MATRIX_START),   # response
    (_C_MATRIX_START + 1, _C_MATRIX_START + 2),  # responses (red CF <> 1)
    (_C_MATRIX_START + 3, _C_MATRIX_START + 4),  # included rows
)

_EMPTY_MODEL_FALLBACK = '"(empty model)"'

# The derived response name, shared by the audit strip and the filtered-y
# header: the header of the first Role="Response" spec row, "(none)" when
# no row carries the role. XMATCH position over the TAKE-trimmed roles is
# the same lookup Response_Column() uses for its data column.
_RESPONSE_NAME_FORMULA = (
    'IFERROR(INDEX(TOROW(Header_Names),'
    'XMATCH("Response",TAKE(Spec_Role,COLUMNS(Source_Data)))),"(none)")'
)

# Dropdown validations cover the repo's standard 16000-row input band so a
# retargeted dataset with more columns inherits them without a rebuild.
_VALIDATION_LAST_ROW = 16000

# Default spec: variable -> (role, include, type). Reference (E) starts
# blank everywhere so the first-in-sort-order default is what gets
# exercised; type an explicit level into E to exercise the override path.
_ROLE_RESPONSE = "Response"
_ROLE_PREDICTOR = "Predictor"
_ROLE_IDENTIFIER = "Identifier"
_ROLE_FILTER = "Filter"
_ROLE_OMIT = "Omit"

_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Country": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Year": (_ROLE_PREDICTOR, True, "Categorical"),
    "Status": (_ROLE_PREDICTOR, True, "Categorical"),
    "Life expectancy": (_ROLE_RESPONSE, False, "Continuous"),
    "Adult Mortality": (_ROLE_PREDICTOR, True, "Continuous"),
    "GDP": (_ROLE_PREDICTOR, True, "Continuous"),
    "Schooling": (_ROLE_PREDICTOR, True, "Continuous"),
    "Full_Data": (_ROLE_FILTER, False, "Continuous"),
}
_FALLBACK_SPEC: tuple[str, bool, str] = (_ROLE_PREDICTOR, False, "Continuous")

_DEFAULT_TRANSFORM = "None"

_ROLE_VALIDATION_LIST = "Response,Predictor,Identifier,Filter,Omit"
_INCLUDE_VALIDATION_LIST = "TRUE,FALSE"
_TYPE_VALIDATION_LIST = "Continuous,Categorical"
_TRANSFORM_VALIDATION_LIST = _DEFAULT_TRANSFORM
_XL_VALIDATE_LIST = 3
_XL_VALID_ALERT_STOP = 1
_XL_BETWEEN = 1
_RESERVED_NOTE = "Reserved for a future release — not yet used by any formula."


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    """Register local names in dependency order (Excel resolves at Add time)."""
    sname = f"'{sheet.name}'"

    local_names: dict[str, str] = {
        # ── Source-table indirection: THE dataset-retarget point ─────────
        # Everything below references Source_Data / Header_Names, never the
        # table name directly, so a dataset changeover is a two-name edit.
        "Source_Data": "=LifeExpectancyData[#Data]",
        "Header_Names": "=LifeExpectancyData[#Headers]",
        # ── Spec ranges (local columns B–G; TAKE-trimmed at use) ─────────
        "Spec_Role": f"={sname}!$B$3:$B${_VALIDATION_LAST_ROW}",
        "Spec_Include": f"={sname}!$C$3:$C${_VALIDATION_LAST_ROW}",
        "Spec_Type": f"={sname}!$D$3:$D${_VALIDATION_LAST_ROW}",
        "Spec_Reference": f"={sname}!$E$3:$E${_VALIDATION_LAST_ROW}",
        # Reserved axes: named now so the grid shape is final, read by
        # nothing until the Order/Transform release.
        "Spec_Order": f"={sname}!$F$3:$F${_VALIDATION_LAST_ROW}",
        "Spec_Transform": f"={sname}!$G$3:$G${_VALIDATION_LAST_ROW}",
    }

    # ── Sample_Include(): the derived row mask ───────────────────────────
    # Per-row AND as a REDUCE product of indicator vectors over spec rows:
    #   Filter columns        — truthy: --(IFERROR((col+0)=1,FALSE)) passes
    #                           TRUE and 1 only (FALSE/0/blank/text/errors
    #                           multiply in a 0).
    #   Response / included   — completeness: N(ISNUMBER(col)).
    #   Continuous Predictors
    #   Everything else       — acc passthrough (Categorical Predictors
    #                           deliberately impose no completeness — the
    #                           known caveat in the human test plan;
    #                           Identifier/Omit impose nothing).
    # Multiplication over {0,1} IS logical AND; the ones seed and the final
    # prod=1 keep it a full-height boolean column with no per-row BYROW.
    #
    # Load-bearing detail — the Filter test coerces with (col+0), NOT N(col).
    # col is INDEX(Source_Data,0,j), a bare RANGE REFERENCE, and N() of a bare
    # reference implicit-intersects: it crushes the whole column to a single
    # value (row 1). Because the shipped Full_Data row 1 (Afghanistan 2015) is
    # TRUE, N(col)=1 collapsed to the scalar 1, so acc*1 was a no-op and the
    # Filter silently dropped out — the mask fell through to completeness-only
    # (2482 instead of 1649 on the T0 default spec). Arithmetic (col+0)
    # broadcasts element-wise over a reference, so the truthiness vector is
    # full-height. (The completeness branch was never affected: ISNUMBER(col)
    # builds a computed array first, and N() of a computed array broadcasts.)
    local_names["Sample_Include"] = (
        "=LAMBDA("
        "LET("
        "n_c,COLUMNS(Source_Data),"
        "rl,TAKE(Spec_Role,n_c),"
        "inc,TAKE(Spec_Include,n_c),"
        "typ,TAKE(Spec_Type,n_c),"
        "seed,SEQUENCE(ROWS(Source_Data),1,1,0),"
        "prod,REDUCE(seed,SEQUENCE(n_c),LAMBDA(acc,j,"
        "LET(col,INDEX(Source_Data,0,j),"
        "IF(INDEX(rl,j)=\"Filter\",acc*--(IFERROR((col+0)=1,FALSE)),"
        "IF(OR(INDEX(rl,j)=\"Response\","
        "AND(INDEX(rl,j)=\"Predictor\",INDEX(inc,j)=TRUE,"
        "INDEX(typ,j)=\"Continuous\")),"
        "acc*N(ISNUMBER(col)),"
        "acc))"
        ")"
        ")),"
        "prod=1"
        ")"
        ")"
    )

    # ── Response_Column(): the derived y ─────────────────────────────────
    # First Role = "Response" match; #N/A when none (consumers IFERROR).
    # Exactly-one validation belongs to the future audit row, not here.
    local_names["Response_Column"] = (
        "=LAMBDA("
        "LET("
        "n_c,COLUMNS(Source_Data),"
        "rl,TAKE(Spec_Role,n_c),"
        "INDEX(Source_Data,0,XMATCH(\"Response\",rl))"
        ")"
        ")"
    )

    # ── Row_Labels(): the derived observation labels ─────────────────────
    # Type dispatch on whether any Identifier columns exist:
    #   none — positional labels "Obs. 1", "Obs. 2", ...
    #   some — per-row TEXTJOIN of ALL Identifier columns in table order,
    #          "|"-separated, ignore_empty=FALSE so field positions stay
    #          aligned when an identifier cell is blank.
    # Dispatch is structural, not data-dependent: no Identifier role means
    # positional labels. ids LET-binds a FILTER wrapped by IFERROR(...,NA())
    # so the all-FALSE case is still safe. Full-height always (the row-mask
    # contract).
    local_names["Row_Labels"] = (
        "=LAMBDA("
        "LET("
        "n_c,COLUMNS(Source_Data),"
        "rl,TAKE(Spec_Role,n_c),"
        "ids,IFERROR(TRANSPOSE(FILTER(TRANSPOSE(Source_Data),"
        "rl=\"Identifier\")),NA()),"
        "IF(SUM(--(rl=\"Identifier\"))=0,"
        "\"Obs. \"&SEQUENCE(ROWS(Source_Data)),"
        "BYROW(ids,LAMBDA(r,TEXTJOIN(\"|\",FALSE,r)))"
        ")"
        ")"
        ")"
    )

    # ── X_s(): the spec-order REDUCE constructor ─────────────────────────
    # Iteration predicate: Role = "Predictor" AND Include = TRUE (the
    # two-axis change — everything else is v1-identical mechanics).
    #   seed        — full-height zeros sentinel; DROPped at the end.
    #   Continuous  — full-height raw passthrough.
    #   Categorical — Dummy_Levels is bound ONCE and its #N/A is the single
    #                 skip signal: degenerate level set and invalid
    #                 reference both surface as ISNA(lv) → acc passthrough.
    #                 d normalizes an empty E-cell (INDEX reads it as 0)
    #                 to "", which Dummy_Levels treats as "use the default"
    #                 (first level in sort order). lv already excludes the
    #                 reference, so --(col = lv) broadcasts the n×1 column
    #                 against the 1×(L−1) row into the 0/1 block directly —
    #                 no FILTER, so no eager-empty-FILTER hazard in LET.
    #   Mask        — read ONLY to fix level sets; output is always
    #                 full-height (the X_s row-mask contract).
    local_names["X_s"] = (
        "=LAMBDA("
        "LET("
        "n_c,COLUMNS(Source_Data),"
        "rl,TAKE(Spec_Role,n_c),"
        "inc,TAKE(Spec_Include,n_c),"
        "typ,TAKE(Spec_Type,n_c),"
        "refs,TAKE(Spec_Reference,n_c),"
        "seed,SEQUENCE(ROWS(Source_Data),1,0,0),"
        "built,REDUCE(seed,SEQUENCE(n_c),LAMBDA(acc,j,"
        "IF(OR(INDEX(rl,j)<>\"Predictor\",INDEX(inc,j)<>TRUE),acc,"
        "LET(col,INDEX(Source_Data,0,j),"
        "IF(INDEX(typ,j)<>\"Categorical\","
        "HSTACK(acc,col),"
        "LET("
        "d,INDEX(refs,j),"
        "r,IF(LEN(d&\"\")=0,\"\",d),"
        "lv,Dummy_Levels(col,r,Sample_Include()),"
        "IF(ISNA(lv),acc,HSTACK(acc,--(col=lv)))"
        ")"
        ")"
        ")"
        ")"
        ")),"
        "DROP(built,,1)"
        ")"
        ")"
    )

    # ── Constructed_Column_Names(): structural twin of X_s() ─────────────
    # Same iteration, same predicate, same skip conditions — twinning is
    # what guarantees name/column alignment (widths must always agree).
    # Emits a ROW for the header strip. Seed is a 1×1 sentinel, dropped.
    local_names["Constructed_Column_Names"] = (
        "=LAMBDA("
        "LET("
        "n_c,COLUMNS(Source_Data),"
        "rl,TAKE(Spec_Role,n_c),"
        "inc,TAKE(Spec_Include,n_c),"
        "typ,TAKE(Spec_Type,n_c),"
        "refs,TAKE(Spec_Reference,n_c),"
        "hdrs,TOROW(Header_Names),"
        "built,REDUCE(\"\",SEQUENCE(n_c),LAMBDA(acc,j,"
        "IF(OR(INDEX(rl,j)<>\"Predictor\",INDEX(inc,j)<>TRUE),acc,"
        "LET(h,INDEX(hdrs,1,j),"
        "IF(INDEX(typ,j)<>\"Categorical\","
        "HSTACK(acc,h),"
        "LET("
        "col,INDEX(Source_Data,0,j),"
        "d,INDEX(refs,j),"
        "r,IF(LEN(d&\"\")=0,\"\",d),"
        "lv,Dummy_Levels(col,r,Sample_Include()),"
        "IF(ISNA(lv),acc,HSTACK(acc,h&\": \"&lv))"
        ")"
        ")"
        ")"
        ")"
        ")),"
        "DROP(built,,1)"
        ")"
        ")"
    )

    for name, refers_to in local_names.items():
        drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)


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
    """The A–H specification block: headers, defaults, dropdowns, CF."""
    bold_row(sheet, _HEADER_ROW, _C_LABEL, _C_LEVELS)
    for col, header in (
        (_C_LABEL, "Variable"),
        (_C_ROLE, "Role"),
        (_C_INCLUDE, "Include"),
        (_C_TYPE, "Type"),
        (_C_REFERENCE, "Reference Level"),
        (_C_ORDER, "Order"),
        (_C_TRANSFORM, "Transform"),
        (_C_LEVELS, "Levels"),
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

        # H: Levels display — Categorical Predictors only; the raw distinct
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
                f'=IF(OR($B{row}<>"Predictor",$D{row}<>"Categorical"),"",'
                f"LET(col,INDEX(Source_Data,0,ROW()-2),"
                f'x,IF(col="","",col),'
                f'IFERROR(ROWS(UNIQUE(FILTER(x,(x<>"")*Sample_Include()))),0)))'
            ),
        )

    _add_list_validation(sheet, _C_ROLE, _ROLE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_INCLUDE, _INCLUDE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_TYPE, _TYPE_VALIDATION_LIST)
    _add_list_validation(sheet, _C_TRANSFORM, _TRANSFORM_VALIDATION_LIST)

    # Cascading relevance: C–H gray out whenever Role ≠ Predictor — the
    # Reference-only-for-Categorical pattern applied one level up.
    add_expression_format(
        sheet,
        f"$C${_FIRST_DATA_ROW}:$H${_LAST_DATA_ROW}",
        f'=$B{_FIRST_DATA_ROW}<>"Predictor"',
        font_color=MUTED_TEXT_COLOR,
    )

    # Degeneracy flag: red H when an INCLUDED Categorical Predictor has
    # L <= 1 — the constructor contributes zero columns for it (visible
    # degradation, not silent omission). N() coerces "" to 0.
    add_expression_format(
        sheet,
        f"$H${_FIRST_DATA_ROW}:$H${_LAST_DATA_ROW}",
        (
            f'=AND($B{_FIRST_DATA_ROW}="Predictor",'
            f"$C{_FIRST_DATA_ROW}=TRUE,"
            f'$D{_FIRST_DATA_ROW}="Categorical",'
            f"N($H{_FIRST_DATA_ROW})<=1)"
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
            f"ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-2),"
            f"$E{_FIRST_DATA_ROW},Sample_Include())))"
        ),
        fill=CF_LIGHT_RED_FILL,
        font_color=CF_DARK_RED_TEXT,
    )


def _write_row_zones(sheet: xw.Sheet) -> None:
    """The J/K derived-row zone: full-height label and mask spills.

    Row 1 of J/K is not written here — _write_audit_row owns the audit
    strip that occupies it.
    """
    sheet.range(rc(1, _C_GAP)).column_width = _GAP_COLUMN_WIDTH

    bold_row(sheet, _HEADER_ROW, _C_ROW_LABELS, _C_INCLUDED)
    val(sheet, _HEADER_ROW, _C_ROW_LABELS, "Row Labels")
    val(sheet, _HEADER_ROW, _C_INCLUDED, "Included")

    f(sheet, _FIRST_DATA_ROW, _C_ROW_LABELS, "=Row_Labels()")
    f(sheet, _FIRST_DATA_ROW, _C_INCLUDED, "=Sample_Include()")


def _write_audit_row(sheet: xw.Sheet) -> None:
    """Row-1 audit strip: bold label/value pairs from column J rightward.

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
            '=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response"))',
        ),
        ("included rows", "=SUMPRODUCT(N(Sample_Include()))"),
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


def _write_filtered_zones(sheet: xw.Sheet) -> None:
    """The M/N and P/Q→ filtered display zones.

    The only row-filtering on the sheet: FILTER(<full-height name>(),
    Sample_Include()). Every spill wraps IFERROR(..., "(empty model)") —
    an empty model (no included predictors, or a mask that excludes
    everything) degrades to the documented string.
    """
    for break_col in (_C_BREAK_LEFT, _C_BREAK_MID):
        sheet.range(rc(1, break_col)).column_width = _GAP_COLUMN_WIDTH

    bold_row(sheet, _HEADER_ROW, _C_FILTERED_LABELS, _C_MATRIX_START)
    val(sheet, _HEADER_ROW, _C_FILTERED_LABELS, "Row Labels")
    # N header carries the derived response name ("y: Life expectancy")
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


def write_model_construction_sheet(workbook: xw.Book) -> xw.Sheet:
    """Create or rebuild the Model Construction sheet."""
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    section_heading(sheet, 1, _C_LABEL, "Model Construction")

    _set_sheet_scoped_names(sheet)
    _write_spec_block(sheet)
    _write_row_zones(sheet)
    _write_audit_row(sheet)
    _write_filtered_zones(sheet)

    # Reserved-column notes are COM comment calls; keep them out of the
    # RecordingSheet-testable spec block.
    _set_note(sheet, _FIRST_DATA_ROW, _C_ORDER, _RESERVED_NOTE)
    _set_note(sheet, _FIRST_DATA_ROW, _C_TRANSFORM, _RESERVED_NOTE)

    sheet.range(
        (_HEADER_ROW, _C_LABEL), (_HEADER_ROW, _C_LEVELS)
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

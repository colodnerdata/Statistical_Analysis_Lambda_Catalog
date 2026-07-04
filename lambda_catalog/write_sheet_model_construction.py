"""Build the Model Construction worksheet — v3.0 declarative specification block.

Two-axis specification (ROADMAP: v3.0 — Specification-Driven Regression):

    A         B      C        D     E                F        G          H
    Variable  Role   Include  Type  Reference Level  Order    Transform  Levels
    (spill)   (drop) (input)  (drop)(input)          (rsvd.)  (rsvd.)    (disp.)

Right of the spec block, after a narrow gap column I (which also visually
reserves the future Design Columns audit column):

    J             K
    Row Labels    Included
    (=Row_Labels() spill at J3; =Sample_Include() spill at K3 — both
     full-height, never internally filtered)

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

Not here (deliberately, per release scoping): the filtered display/audit
zones (filtered matrix, header strip, row-1 audit cells) — those land in
the next PR, which is also why row 1 of columns J/K stays empty.
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
    #   Filter columns        — truthy: N(IFERROR(N(col)=1,FALSE)) passes
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
        "IF(INDEX(rl,j)=\"Filter\",acc*N(IFERROR(N(col)=1,FALSE)),"
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

    Row 1 of J/K stays empty — the next PR's audit cells land there.
    """
    sheet.range(rc(1, _C_GAP)).column_width = _GAP_COLUMN_WIDTH

    bold_row(sheet, _HEADER_ROW, _C_ROW_LABELS, _C_INCLUDED)
    val(sheet, _HEADER_ROW, _C_ROW_LABELS, "Row Labels")
    val(sheet, _HEADER_ROW, _C_INCLUDED, "Included")

    f(sheet, _FIRST_DATA_ROW, _C_ROW_LABELS, "=Row_Labels()")
    f(sheet, _FIRST_DATA_ROW, _C_INCLUDED, "=Sample_Include()")


def write_model_construction_sheet(workbook: xw.Book) -> xw.Sheet:
    """Create or rebuild the Model Construction sheet."""
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    section_heading(sheet, 1, _C_LABEL, "Model Construction")

    _set_sheet_scoped_names(sheet)
    _write_spec_block(sheet)
    _write_row_zones(sheet)

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

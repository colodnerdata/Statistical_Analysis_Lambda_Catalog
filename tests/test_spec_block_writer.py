"""RecordingSheet tests for the spec-block component library (Regression sheet).

Excel-side behavior (spill evaluation, Dummy_Levels calls, conditional
formatting rendering) is exercised by the Excel verifier; these tests pin
everything checkable without Excel — the sheet-scoped name definitions and
their order, the T0 default-spec prefill, the dropdown and conditional-
formatting registrations, and the structural invariants of the Predictor_Columns /
Constructed_Column_Names twins.
"""
# pylint: disable=missing-function-docstring,protected-access
from pathlib import Path
from typing import cast

import xlwings as xw

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_DARK_YELLOW_TEXT,
    CF_LIGHT_RED_FILL,
    CF_YELLOW_FILL,
    HEADER_COLOR,
    INPUT_COLOR,
)
from lambda_catalog.workbook_helpers import col_letter, excel_color
from lambda_catalog.write_spec_block import (
    _AUDIT_PAIRS,
    _AUDIT_ROW,
    _C_BREAK_LEFT,
    _C_BREAK_MID,
    _C_DESIGN_COLUMNS,
    _C_FEEDBACK_COUNT,
    _C_FEEDBACK_DELTA,
    _C_FILTERED_LABELS,
    _C_FILTERED_Y,
    _C_INCLUDE,
    _C_LABEL,
    _C_LEVELS,
    _C_MATRIX_LABELS,
    _C_MATRIX_START,
    _C_ORDER,
    _C_PERIOD_IN_USE,
    _C_REF_IN_USE,
    _C_REFERENCE,
    _C_ROLE,
    _C_SEQUENCE,
    _C_SEQUENCE_PERIOD,
    _C_SPEC_LAST,
    _C_TRANSFORM,
    _C_TYPE,
    _CLOSURE_SCOPE,
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _FALLBACK_SPEC,
    _FEEDBACK_LABEL_ROW,
    _FEEDBACK_STATUS_ROW,
    _FIRST_DATA_ROW,
    _FIXED_EFFECTS_COUNT_FORMULA,
    _HEADER_ROW,
    _INTERACTION_HEADER_SYMBOLS,
    _INTERACTION_HEADER_UNKNOWN,
    _INTERCEPT_ROW,
    _MSG_CALENDAR,
    _MSG_NO_NATURAL,
    _MSG_OFF_GRID,
    _MSG_REGULARITY,
    _N_VARIABLES,
    _RESPONSE_COUNT_FORMULA,
    _VALIDATION_LAST_ROW,
    _VARIABLES,
    SPEC_DATASET_PROFILES,
    _set_sheet_scoped_names,
    _set_spec_block_column_widths,
    _write_audit_row,
    _write_filtered_zones,
    _write_intercept_control,
    _write_row_zones,
    _write_spec_block,
)
from tests.recording_sheet import RecordingSheet

ROOT_DIR = Path(__file__).resolve().parents[1]

# Local label for the ``RecordingSheet`` instances the tests build mocks
# with. The spec block lives on the Regression sheet; this constant is only
# a sheet name the mocks need, not a reference to a shipped sheet.
SHEET_NAME = "Model Construction"

_EXPECTED_NAME_ORDER = [
    "Source_Table",
    "Source_Data",
    "Header_Names",
    "Spec_Role",
    "Spec_Include",
    "Spec_Type",
    "Spec_Reference",
    "Spec_Order",
    "Spec_Transform",
    "Spec_Sequence",
    "Spec_Sequence_Period",
    "Spec_Period_In_Use",
    "Spec_Interaction_Term",
    "Spec_Interaction_Operation",
    "Spec_Design_Columns",
    "Allow_Intercept",
    # Closures follow, in catalog document order (= dependency order).
    # Base_Period_Delta leads: it reads only the wiring names above, never
    # another closure.
    "Base_Period_Delta",
    # v3.2 name-promotion: the REDUCE bodies live in the _Calc computational
    # leaves, which the spill-source cells call and the public reader names
    # delegate to. _Calc precedes its public reader (dependency order).
    "Sample_Include_Calc",
    "Sample_Include",
    "Response_Column",
    "Row_Labels",
    "Predictor_Columns",
    "Constructed_Column_Names",
    "Constructed_Column_Transforms",
    "Sequence_Column",
    "Fixed_Effects_Column",
    "Absorbed_Degrees_Of_Freedom",
    "Prediction_Group_Column",
    "Design_Response",
    "Design_Columns_Calc",
    "Design_Columns",
    "Serial_Correlation_Group",
    "Sequence_Deltas",
    "Base_Period_Delta_Candidate",
    "Sequence_Delta_Spectrum",
    "Model_Formula",
    # The row-2 status readouts. Three read only the wiring names, but
    # Log_Domain_Status calls Sample_Include(FALSE) (which delegates to the
    # _Calc leaf), so the whole group installs after the constructors rather
    # than being interleaved among them.
    "Role_Status",
    "Sequence_Status",
    "Log_Domain_Status",
    "Design_Width_Status",
]


def _as_xw_sheet(sheet: RecordingSheet) -> xw.Sheet:
    return cast(xw.Sheet, sheet)


def _model_construction_closures():
    """The sheet-scoped constructor functions as a standalone rebuild installs them.

    The closures are sheet-scoped to "Regression"; this module installs the
    same set when its sheet is rebuilt standalone.
    """
    document = load_catalog_document(ROOT_DIR / "lambda_functions.json")
    return document.functions_for_sheet(_CLOSURE_SCOPE)


def _named_sheet() -> RecordingSheet:
    sheet = RecordingSheet(name=SHEET_NAME)
    _set_sheet_scoped_names(_as_xw_sheet(sheet), _model_construction_closures())
    return sheet


def _refers_to(sheet: RecordingSheet, name: str) -> str:
    return sheet.api.Names.by_short_name(name).RefersTo


_LAMBDA_OPEN = "=LAMBDA(\n    "
_LAMBDA_CLOSE = "\n)"


def _catalog_body(name: str) -> str:
    """One status LAMBDA's body, as the cell formula now reads it.

    The row-2 status logic lives in lambda_functions.json, so these
    assertions read the shipped catalog body. No import can reach a JSON
    string literal, which is exactly why these assertions have to stay:
    they are the only thing standing between a message edit and a silently
    changed verdict on the sheet.

    The zero-argument wrapper is stripped rather than the whole display
    string being compared, so what comes back is character-for-character the
    catalog body — which keeps these assertions readable AND exact.
    Compacting whitespace instead would have been the easy way and the wrong
    one: these bodies carry message strings with meaningful spaces.
    """
    document = load_catalog_document(ROOT_DIR / "lambda_functions.json")
    entry = next((fn for fn in document.functions if fn.name == name), None)
    # A missing name is the likeliest way this helper gets called wrong — a
    # rename that half-landed, or a typo in an assertion. Say which name, not
    # StopIteration from somewhere inside the generator.
    assert entry is not None, f"{name!r} is not in lambda_functions.json"
    assert entry.scope == _CLOSURE_SCOPE, (name, entry.scope)
    display = entry.formula_display
    assert display.startswith(_LAMBDA_OPEN) and display.endswith(_LAMBDA_CLOSE), name
    return "=" + display[len(_LAMBDA_OPEN) : -len(_LAMBDA_CLOSE)]


def _all_written_formulas(sheet: RecordingSheet) -> list[str]:
    return [
        rng.state.formula2
        for rng in sheet.ranges.values()
        if rng.state.formula2 is not None
    ]


def _write_all_zones(sheet: RecordingSheet) -> None:
    _write_spec_block(_as_xw_sheet(sheet))
    _write_row_zones(_as_xw_sheet(sheet))
    _write_audit_row(_as_xw_sheet(sheet))
    _write_filtered_zones(_as_xw_sheet(sheet))


def test_names_are_created_in_dependency_order() -> None:
    sheet = _named_sheet()

    local_name_order = [item.Name.split("!", 1)[-1] for item in sheet.api.Names.items]
    assert local_name_order == _EXPECTED_NAME_ORDER


def test_both_writers_register_names_before_the_spec_block() -> None:
    """The block's four computed columns are spills that read the Spec_*
    bands, Source_Data, Header_Names and the constructor closures, so every
    one of those names has to exist first.

    The bands are TAKE-trimmed spills, not structured references into a
    ListObject, so Excel validates nothing at Names.Add time and the names
    must be registered first. Getting it backwards does not raise — the
    spills just parse against names that do not exist yet and sit at #NAME?
    until something re-registers them — which is exactly why it is pinned
    here rather than left to the build to reveal.

    Scoped to each WRITER's own source, not its module's. Searching the whole
    module matches the callee's `def` line — which sits above both call sites
    and never moves — so the assertion would hold whatever order the writer
    actually used.
    """
    import inspect

    from lambda_catalog.write_sheet_regression import write_regression_output_sheet

    # The Regression sheet's write_regression_output_sheet registers
    # sheet-scoped names via _setup_local_names *before* calling
    # _write_spec_block — so a name like RegChartFitY is created before
    # any formula that references it lands on the sheet.
    source = inspect.getsource(write_regression_output_sheet)
    # Each appears exactly once, so "first occurrence" is the call site
    # and cannot drift to some other mention.
    assert source.count("_setup_local_names(") == 1
    assert source.count("_write_spec_block(") == 1
    assert source.index("_setup_local_names(") < source.index("_write_spec_block(")


def test_only_the_retarget_name_references_the_table_directly() -> None:
    sheet = _named_sheet()

    # Source_Table is THE dataset-retarget point — a changeover is a one-name
    # edit. The body and header row derive from it via non-volatile DROP/TAKE
    # (OFFSET would be re-evaluated on every Data Table substitution pass).
    assert _refers_to(sheet, "Source_Table") == "=MileageData[#All]"
    assert _refers_to(sheet, "Source_Data") == "=DROP(Source_Table,1)"
    assert _refers_to(sheet, "Header_Names") == "=TAKE(Source_Table,1)"
    for name in _EXPECTED_NAME_ORDER[1:]:
        assert "LifeExpectancyData" not in _refers_to(sheet, name), name

    _write_all_zones(sheet)
    for formula in _all_written_formulas(sheet):
        assert "LifeExpectancyData" not in formula, formula


def test_source_table_wiring_can_be_overridden() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _set_sheet_scoped_names(
        _as_xw_sheet(sheet),
        _model_construction_closures(),
        source_table_ref="=LifeExpectancyData[#All]",
    )
    assert _refers_to(sheet, "Source_Table") == "=LifeExpectancyData[#All]"
    assert _refers_to(sheet, "Source_Data") == "=DROP(Source_Table,1)"
    assert _refers_to(sheet, "Header_Names") == "=TAKE(Source_Table,1)"


def test_spec_ranges_cover_the_standard_input_band() -> None:
    sheet = _named_sheet()

    # Each Spec_* band is its column's full input range TAKE-trimmed to the
    # live source-table width, so a Source_Table retarget resizes every band
    # with it — a fixed-width reference would pin them to the row count baked
    # in at build time. The band's ceiling is the same _VALIDATION_LAST_ROW
    # the dropdowns and the conditional formatting already use — one ceiling,
    # so they cannot disagree about how far the block can grow.
    for name, column in (
        ("Spec_Role", "B"),
        ("Spec_Include", "C"),
        ("Spec_Type", "D"),
        ("Spec_Reference", "E"),
        ("Spec_Order", "F"),
        ("Spec_Transform", "G"),
        ("Spec_Sequence", "H"),
        ("Spec_Sequence_Period", "I"),
        ("Spec_Period_In_Use", "J"),
    ):
        assert _refers_to(sheet, name) == (
            f"=TAKE('{SHEET_NAME}'!${column}$4:${column}$16000,"
            "MAX(1,COLUMNS(Source_Data)))"
        )


def test_sample_include_is_the_reduce_product_mask() -> None:
    sheet = _named_sheet()
    # v3.2 name-promotion: the REDUCE body lives in the _Calc computational
    # leaf; the public Sample_Include name is a reader that delegates to it
    # (see test_sample_include_is_a_reader_over_its_spill). The leaf carries
    # the optional apply_log_domain argument verbatim.
    mask = _refers_to(sheet, "Sample_Include_Calc")

    # One optional argument: Sample_Include_Calc(FALSE) is the mask WITHOUT the
    # Log positivity layer, which is what the G2 status cell (via the public
    # Sample_Include(FALSE) reader) differences against the default to report
    # the excluded-row count. Every existing call site omits it and is unaffected.
    assert mask.startswith("=LAMBDA([apply_log_domain],LET(")
    assert (
        "use_log,IF(ISOMITTED(apply_log_domain),TRUE,apply_log_domain)"
    ) in mask
    # Filter columns: truthy — TRUE and 1 pass, FALSE/0/blank/text fail.
    # Coercion is (col+0), NOT N(col): col is a bare range reference, and N()
    # of a bare reference implicit-intersects it to a scalar, silently voiding
    # the Filter (the 2482-vs-1649 mask bug). Arithmetic broadcasts instead.
    assert 'IF(INDEX(rl,j)="Filter",acc*--(IFERROR((col+0)=1,FALSE))' in mask
    assert "N(IFERROR(N(col)" not in mask
    # Completeness: the Response and every included Continuous Predictor —
    # and, for those declaring Log (drop ≤ 0) only, strict positivity. Same
    # (col+0) coercion as the Filter branch, for the same reason.
    assert (
        'IF(OR(INDEX(rl,j)="Response (y)",'
        'AND(INDEX(rl,j)="Predictor (x)",INDEX(inc,j)=TRUE,'
        'INDEX(typ,j)="Continuous")),acc*N(ISNUMBER(col))'
        '*IF(AND(use_log,INDEX(trn,j)="Log (drop ≤ 0)"),'
        "--(IFERROR((col+0)>0,FALSE)),1),acc)"
    ) in mask
    # Plain Log must NOT filter — the #N/A is the signal, and the token test
    # is an equality against the filtering token alone.
    assert 'INDEX(trn,j)="Log"' not in mask
    # Full-height ones seed; product over {0,1} is the AND, no per-row loop.
    assert "seed,SEQUENCE(ROWS(Source_Data),1,1,0)" in mask
    assert "BYROW(" not in mask
    assert mask.endswith("prod=1))")
    # Reads the model axes only — never the reserved Order column or the
    # Sequence structural axis (which no constructor may consume).
    # Spec_Transform IS read now, and only for the positivity layer above.
    for non_model_axis in (
        "Spec_Order",
        "Spec_Sequence",
        "Spec_Sequence_Period",
    ):
        assert non_model_axis not in mask


def test_row_labels_dispatches_on_identifier_presence() -> None:
    sheet = _named_sheet()
    labels = _refers_to(sheet, "Row_Labels")

    assert labels.startswith("=LAMBDA(LET(")
    # The LET-bound FILTER is wrapped in IFERROR so the all-FALSE case is
    # still safe at binding time.
    assert (
        'ids,IFERROR(TRANSPOSE(FILTER(TRANSPOSE(Source_Data),'
        'rl="Identifier (Row Label)")),NA())'
    ) in labels
    assert 'IF(SUM(--(rl="Identifier (Row Label)"))=0,' in labels
    # No Identifier columns: positional fallback, full height.
    assert '"Obs. "&SEQUENCE(ROWS(Source_Data))' in labels
    # ignore_empty=FALSE keeps field positions aligned across rows.
    assert 'BYROW(ids,LAMBDA(r,TEXTJOIN("|",FALSE,r)))' in labels


def test_row_zones_spill_full_height_next_to_the_spec_block() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_row_zones(_as_xw_sheet(sheet))

    # R: narrow gap, splitting the spec block and the derived-row zone.
    assert sheet.range((1, 18)).column_width == 2

    # S/T headers on the spec-header row, bold like the A–O headers.
    assert sheet.cell(_HEADER_ROW, 19).value == "Row Labels"
    assert sheet.cell(_HEADER_ROW, 20).value == "Included"
    assert sheet.range((_HEADER_ROW, 19), (_HEADER_ROW, 20)).api.Font.Bold is True

    # Full-height spills at the first data row; row 1 belongs to
    # _write_audit_row, so this writer must leave it untouched.
    assert sheet.cell(_FIRST_DATA_ROW, 19).api.Formula2 == "=Row_Labels()"
    assert sheet.cell(_FIRST_DATA_ROW, 20).api.Formula2 == "=Sample_Include()"
    for col in (19, 20):
        assert sheet.cell(1, col).value is None
        assert sheet.cell(1, col).api.Formula2 is None


def test_spec_block_column_widths_hide_the_reserved_order_column() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _set_spec_block_column_widths(_as_xw_sheet(sheet))

    assert sheet.range((1, _C_LABEL), (1, _C_LABEL)).column_width == 28
    assert sheet.range((1, _C_ORDER), (1, _C_ORDER)).column_width == 0
    # G (Transform) went live at v2.2. Widened from 11 to 14 when the second
    # Log token arrived — "Log (drop ≤ 0)" has to render in the cell, not only
    # in the dropdown, or the two tokens are indistinguishable at a glance.
    assert sheet.range((1, _C_TRANSFORM), (1, _C_TRANSFORM)).column_width == 14
    assert sheet.range((1, _C_REF_IN_USE), (1, _C_REF_IN_USE)).column_width == 16


def test_x_s_binds_dummy_levels_once_and_skips_on_isna() -> None:
    sheet = _named_sheet()
    x_s = _refers_to(sheet, "Predictor_Columns")

    assert x_s.startswith("=LAMBDA(LET(")
    # Still bound exactly once, even though v3.1 evaluates blk() for two
    # rows per iteration (the declaring row and its interaction operand) —
    # one textual site is what keeps the two encodings identical.
    assert x_s.count("Dummy_Levels(") == 1
    assert 'lv,Dummy_Levels(col,r,Sample_Include_Calc())' in x_s
    # Scalar skip guard: ISNA(INDEX(...,1,1)), NOT ISNA(lv). lv is a 1x(L-1)
    # row; an array condition in front of a wider HSTACK branch broadcasts to
    # #N/A (the T6 header-strip bug). INDEX(...,1,1) makes the test scalar,
    # and keep() applies it only on the Categorical branch — a Continuous
    # column whose first cell is an error must not be mistaken for the
    # degenerate-categorical sentinel.
    assert "IF(ISNA(INDEX(lv,1,1)),NA(),--(col=lv))" in x_s
    assert 'keep,LAMBDA(x,arr,IF(INDEX(typ,x)<>"Categorical",TRUE,' in x_s
    assert "NOT(ISNA(INDEX(arr,1,1)))" in x_s
    assert "IF(ISNA(lv)," not in x_s
    # Empty E-cell normalization: INDEX reads a blank cell as 0; "" is the
    # "use the default" sentinel Dummy_Levels expects.
    assert 'r,IF(LEN(d&"")=0,"",d)' in x_s
    # No FILTER anywhere — level filtering is Dummy_Levels' job, and a
    # LET-bound FILTER would evaluate eagerly on empty results.
    assert "FILTER(" not in x_s
    # Full-height sentinel seed, dropped at the end (the row-mask contract).
    assert "seed,SEQUENCE(ROWS(Source_Data),1,0,0)" in x_s
    assert x_s.endswith("DROP(built,,1)))")
    # v2.2 Log wiring: a Continuous column is Ln_Positive-transformed when
    # its row's Transform is Log; the Categorical branch never reads trn.
    assert "trn,TAKE(Spec_Transform,n_c)" in x_s
    assert (
        'IF(OR(INDEX(trn,x)="Log",INDEX(trn,x)="Log (drop ≤ 0)"),'
        "Ln_Positive(col,Fit_Sample_Include()),col)"
    ) in x_s


def test_constructed_column_names_is_a_structural_twin_of_x_s() -> None:
    sheet = _named_sheet()
    x_s = _refers_to(sheet, "Predictor_Columns")
    names = _refers_to(sheet, "Constructed_Column_Names")
    transforms = _refers_to(sheet, "Constructed_Column_Transforms")

    # Identical iteration predicate and skip conditions — twinning is what
    # guarantees the header strip and transform-flag widths always match
    # COLUMNS(Predictor_Columns()). Three-way twin since v2.2's Log wiring added
    # Constructed_Column_Transforms alongside
    # Predictor_Columns/Constructed_Column_Names.
    predicate = 'IF(OR(INDEX(rl,j)<>"Predictor (x)",INDEX(inc,j)<>TRUE),acc,'
    assert predicate in x_s
    assert predicate in names
    assert predicate in transforms
    assert 'lv,Dummy_Levels(col,r,Sample_Include_Calc())' in names
    assert names.count("Dummy_Levels(") == 1
    assert 'lv,Dummy_Levels(col,r,Sample_Include_Calc())' in transforms
    assert transforms.count("Dummy_Levels(") == 1
    # Same scalar skip guard as Predictor_Columns (the twin must match).
    for formula in (x_s, names, transforms):
        assert "IF(ISNA(INDEX(lv,1,1)),NA()," in formula
        assert "IF(ISNA(lv)," not in formula
        assert "IF(NOT(keep(j,a)),acc," in formula
        assert "IF(NOT(keep(q,b)),m," in formula
    # Level-qualified headers: "Status: Developing", "Year: 2001", ... —
    # relabelled "Ln(header)" for a Log-transformed Continuous predictor.
    assert 'INDEX(hdrs,1,x)&": "&lv' in names
    assert 'trn,TAKE(Spec_Transform,n_c)' in names
    assert (
        'IF(OR(INDEX(trn,x)="Log",INDEX(trn,x)="Log (drop ≤ 0)"),'
        '"Ln("&INDEX(hdrs,1,x)&")",INDEX(hdrs,1,x))'
    ) in names
    assert names.endswith("DROP(built,,1)))")
    # Constructed_Column_Transforms: "Log"/"None" per Continuous column;
    # every dummy column from a Categorical Predictor reads "None"
    # unconditionally, regardless of its spec row's own Transform cell.
    # Both tokens report "Log" here: the unit-space dispatcher keys on the
    # SPACE the column is in, and they produce the same one.
    assert (
        'IF(OR(INDEX(trn,x)="Log",INDEX(trn,x)="Log (drop ≤ 0)"),"Log","None")'
    ) in transforms
    assert 'EXPAND("None",1,COLUMNS(lv),"None")' in transforms


def test_the_three_twins_gate_interactions_identically() -> None:
    # The twin property is what guarantees the header strip and the
    # transform-flag strip stay exactly as wide as the design matrix. v3.1
    # adds a second emission point per iteration, so the gating in front of
    # it has to be identical in all three — a mismatch would silently
    # desynchronise names from columns.
    sheet = _named_sheet()
    formulas = [
        _refers_to(sheet, name)
        for name in (
            "Predictor_Columns",
            "Constructed_Column_Names",
            "Constructed_Column_Transforms",
        )
    ]

    # mate(): the operand resolver, byte-identical across the three.
    mate = (
        "mate,LAMBDA(j,LET(t,INDEX(itm,j),o,INDEX(iop,j),"
        "q,IFERROR(XMATCH(t,hdrs),0),"
        'IF(OR(LEN(t&"")=0,LEN(o&"")=0,q=0),0,'
        'IF(INDEX(rl,q)<>"Predictor (x)",0,q))))'
    )
    for formula in formulas:
        assert mate in formula
        # Both operands' blocks come from the SAME blk(), so an operand
        # encodes exactly as it would as a main effect.
        assert "a,blk(j)" in formula
        assert "b,blk(q)" in formula
        assert "q,mate(j)" in formula
        # No interaction declared → the row's own block only.
        assert "IF(q=0,m," in formula

    columns, names, transforms = formulas
    # Predictor_Columns and Constructed_Column_Names walk the pair with the
    # same nested REDUCE, so the k-th name always describes the k-th column.
    pairwise = (
        "REDUCE(m,SEQUENCE(COLUMNS(a)),LAMBDA(p,ai,"
        "REDUCE(p,SEQUENCE(COLUMNS(b)),LAMBDA(pp,bi,"
    )
    assert pairwise in columns
    assert pairwise in names
    # The closed operation vocabulary, on the columns side only.
    assert 'SWITCH(o,"Product",INDEX(a,0,ai)*INDEX(b,0,bi)' in columns
    assert '"Difference",INDEX(a,0,ai)-INDEX(b,0,bi)' in columns
    # Ratio is an EXPLICIT case, not the trailing fallthrough. SWITCH's last
    # argument is its DEFAULT, so leaving Ratio implicit would silently treat
    # any unrecognized operation as a ratio — and data validation does not
    # block a paste into N. The default is NA(): refuse, never guess.
    assert '"Ratio",IFERROR(INDEX(a,0,ai)/INDEX(b,0,bi),NA()),NA())' in columns
    # The operands' own names joined by the OPERATION's own symbol — a single
    # separator could not say which of the three built the column, and a colon
    # additionally collided with the ": " inside a level-qualified name.
    assert "INDEX(a,0,ai)&SWITCH(o," in names
    assert "&INDEX(b,0,bi)" in names
    # Transforms needs no pairwise walk — COLUMNS(a)*COLUMNS(b) is exactly
    # the width that walk emits, and every interaction column reads "None"
    # (the transform lives on each operand's column, applied before they
    # are combined).
    assert 'EXPAND("None",1,COLUMNS(a)*COLUMNS(b),"None")' in transforms
    assert "SWITCH(" not in transforms
    assert transforms.endswith("DROP(built,,1)))")


def test_reserved_spec_order_is_defined_but_read_by_nothing() -> None:
    # F (Order) remains reserved. G (Transform) went live at v2.2 — see
    # test_spec_transform_is_read_only_by_the_transform_aware_constructors
    # for its (positive) counterpart. I (Base Period Δ) went live with the
    # base-period release and has its own reader tests below.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    reserved = "Spec_Order"
    readers = [
        item.Name
        for item in sheet.api.Names.items
        if reserved in item.RefersTo
        and item.Name.split("!", 1)[-1] != reserved
    ]
    assert readers == [], reserved
    for formula in _all_written_formulas(sheet):
        assert reserved not in formula, (reserved, formula)


def test_spec_transform_is_read_only_by_the_transform_aware_constructors() -> None:
    # Confirm-by-construction property: Spec_Transform is read by exactly the
    # four constructors the Log wiring touches, plus Model_Formula (the
    # DISPLAY that renders the response's Log wrapping into the formula
    # caption), plus Sample_Include — and by nothing else; in particular NOT
    # by Row_Labels, which never transforms anything.
    #
    # Sample_Include_Calc joined this list with the second Log token (the REDUCE
    # body lives in the _Calc leaf; the public Sample_Include name is a reader
    # over the materialized spill and does not read Spec_Transform). It is the
    # one reader that does not transform: it reads Spec_Transform ONLY to
    # decide whether a column's non-positive rows leave the sample, which is
    # the sole difference between "Log" and "Log (drop ≤ 0)". The narrower
    # assertion below is what keeps that from widening into the mask making
    # transform decisions of its own.
    #
    # Log_Domain_Status joined it with Part 6.2, and is the second non-
    # transforming reader: it REPORTS on the Log declarations (which column a
    # strict Log poisoned, how many rows the filtering token dropped) without
    # building a single column. A status readout is allowed to read the spec;
    # what this test still forbids is a new CONSTRUCTOR appearing here.
    sheet = _named_sheet()
    readers = sorted(
        item.Name.split("!", 1)[-1]
        for item in sheet.api.Names.items
        if "Spec_Transform" in item.RefersTo
        and item.Name.split("!", 1)[-1] != "Spec_Transform"
    )
    assert readers == [
        "Constructed_Column_Names",
        "Constructed_Column_Transforms",
        "Log_Domain_Status",
        "Model_Formula",
        "Predictor_Columns",
        "Response_Column",
        "Sample_Include_Calc",
    ]
    assert "Spec_Transform" not in _refers_to(sheet, "Row_Labels")
    # The public Sample_Include reader delegates to the spill / _Calc leaf and
    # does not read Spec_Transform itself — that read lives in the _Calc body.
    assert "Spec_Transform" not in _refers_to(sheet, "Sample_Include")
    # Sample_Include_Calc tests the filtering token and nothing else: no Ln,
    # no renaming, no branch on plain "Log".
    mask = _refers_to(sheet, "Sample_Include_Calc")
    assert 'INDEX(trn,j)="Log (drop ≤ 0)"' in mask
    assert "Ln_Positive" not in mask
    assert 'INDEX(trn,j)="Log"' not in mask


def test_sequence_name_is_read_only_by_validation_and_axis_layers() -> None:
    # Spec_Sequence is live for the zero-or-one validation (H2 status line,
    # audit count), the sequence-spacing layer (Sequence_Deltas), and — since
    # the DW-gate release — the serial-correlation accessor (Sequence_Column,
    # which feeds Durbin_Watson_By). No design-matrix CONSTRUCTOR closure may
    # consume it: Sequence orders the data, it does not enter the model matrix.
    sheet = _named_sheet()
    readers = sorted(
        item.Name.split("!", 1)[-1]
        for item in sheet.api.Names.items
        if "Spec_Sequence" in item.RefersTo
        and item.Name.split("!", 1)[-1] != "Spec_Sequence"
    )
    # Base_Period_Delta joined this list when it became sheet-scoped: it reads
    # the flagged row's Period In Use cell, which is the sequence layer's job.
    # It is an ACCESSOR, not a constructor — the assertion below is the one
    # that matters, and it is unchanged.
    #
    # Sequence_Status joined it in Part 6.2. It is the H2 zero-or-one
    # validation named in this test's own first line; it simply moved out of
    # the cell and into the catalog, so the reader set describes the same
    # layers it always did.
    assert readers == [
        "Base_Period_Delta",
        "Sequence_Column",
        "Sequence_Deltas",
        "Sequence_Status",
    ]
    for constructor in (
        "Sample_Include",
        "Response_Column",
        "Row_Labels",
        "Predictor_Columns",
        "Constructed_Column_Names",
    ):
        assert "Spec_Sequence" not in _refers_to(sheet, constructor), constructor


def test_sequence_period_name_is_read_only_by_the_base_period_layer() -> None:
    # Spec_Sequence_Period is live, but only for the Δ-in-use display in
    # the Sequence Spacing block (on-sheet) and the Base_Period_Delta()
    # accessor (in lambda_functions.json). No CONSTRUCTOR reads it: the
    # candidate is computed from the data, never from its own display cell.
    #
    # Base_Period_Delta became sheet-scoped so a workbook with more than one
    # Regression-shaped sheet gives each its own Δ, so it now appears in this
    # sheet's Name Manager rather than the workbook's. It is the one
    # legitimate name reader, and naming it here is what keeps the check
    # meaningful — a second reader appearing would still fail.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    name_readers = [
        item.Name.split("!", 1)[-1]
        for item in sheet.api.Names.items
        if "Spec_Sequence_Period" in item.RefersTo
        and item.Name.split("!", 1)[-1] != "Spec_Sequence_Period"
    ]
    assert name_readers == ["Base_Period_Delta"]

    formula_readers = [
        formula
        for formula in _all_written_formulas(sheet)
        if "Spec_Sequence_Period" in formula
    ]
    # Exactly one, and it is the Period In Use spill. That display reads the
    # typed override to decide whether to show it in place of the computed
    # candidate; it reaches the same cell through the band name (no
    # structured reference — the block is table-free). It is a DISPLAY, not
    # a constructor — a second reader, or one that does not consult the
    # candidate, is the regression this test exists to catch.
    assert len(formula_readers) == 1, formula_readers
    assert "Base_Period_Delta_Candidate()" in formula_readers[0]


def test_reserved_spec_order_is_not_referenced_repo_wide() -> None:
    # Spec_Transform is deliberately excluded here now that it's live —
    # it legitimately appears in lambda_functions.json (the four
    # transform-aware constructors) and in write_sheet_regression.py's
    # note-swap import.
    own_module = "write_spec_block.py"
    sources = [
        path
        for path in (ROOT_DIR / "lambda_catalog").glob("*.py")
        if path.name != own_module
    ]
    sources.append(ROOT_DIR / "lambda_functions.json")
    for path in sources:
        text = path.read_text(encoding="utf-8")
        assert "Spec_Order" not in text, path.name


def test_every_name_and_formula_string_is_balanced() -> None:
    sheet = _named_sheet()
    _write_all_zones(sheet)

    strings = [item.RefersTo for item in sheet.api.Names.items]
    strings.extend(_all_written_formulas(sheet))
    for text in strings:
        assert text.count("(") == text.count(")"), text
        assert text.count('"') % 2 == 0, text


def test_spec_block_prefills_the_t0_default_configuration() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    assert _N_VARIABLES == 11
    assert sheet.cell(_FIRST_DATA_ROW, 1).api.Formula2 == "=TRANSPOSE(Header_Names)"
    header_row = sheet.range((_HEADER_ROW, _C_LABEL), (_HEADER_ROW, _C_SPEC_LAST))
    assert header_row.color == HEADER_COLOR
    assert header_row.api.Font.Bold is True
    assert header_row.api.Font.Color == excel_color((0, 0, 0))
    for offset, variable in enumerate(_VARIABLES):
        row = _FIRST_DATA_ROW + offset
        role, include, ptype = _DEFAULT_SPEC.get(variable, _FALLBACK_SPEC)
        assert sheet.cell(row, 2).value == role, variable
        assert sheet.cell(row, 3).value is include, variable
        assert sheet.cell(row, 4).value == ptype, variable
        assert sheet.cell(row, 5).value is None, variable  # E blank → default
        assert sheet.cell(row, 6).value is None, variable  # F reserved, blank
        assert sheet.cell(row, 7).value == "None", variable  # G reserved
        # H (Sequence): TRUE on the shipped ordering axis (Year), blank
        # elsewhere. Zero-or-one flags is the legal range.
        expected_sequence = True if variable in _DEFAULT_SEQUENCE_VARIABLES else None
        assert sheet.cell(row, 8).value is expected_sequence, variable
        # I (Sequence Period) is the typed override input — blank by
        # default; the user types a number here to override the candidate.
        # The candidate display lives in J.
        assert sheet.cell(row, 9).value is None, variable
        # J/K/L/O carry no per-row formula any more: each is ONE spill
        # written at _FIRST_DATA_ROW, asserted separately below.
        if row != _FIRST_DATA_ROW:
            assert sheet.cell(row, 10).api.Formula2 is None, variable

    # Spot-check the named T0 roles — the shipped spec demonstrates Identifier,
    # Response, Predictor, and Omit. Auto MPG ships no Filter-by-default column:
    # the old Full_Data completeness column was redundant with the per-predictor
    # mask (it checked the same measurement columns the built-in mask already
    # filters for blanks) and is gone, so the active Filter role is exercised
    # only by the Is_USA fixture in the test-model suite (M15).
    by_variable = {v: _FIRST_DATA_ROW + i for i, v in enumerate(_VARIABLES)}
    assert "Full_Data" not in by_variable
    assert sheet.cell(by_variable["Car Name"], 2).value == "Identifier (Row Label)"
    assert sheet.cell(by_variable["MPG"], 2).value == "Response (y)"
    assert sheet.cell(by_variable["Make"], 2).value == "Omit"
    assert sheet.cell(by_variable["Model?"], 2).value == "Omit"
    for categorical in ("Model Year", "Origin"):
        row = by_variable[categorical]
        assert sheet.cell(row, 2).value == "Predictor (x)"
        assert sheet.cell(row, 3).value is True
        assert sheet.cell(row, 4).value == "Categorical"
    # No column is Sequence-flagged. Auto MPG is cross-sectional, so the
    # shipped profile declares no ordering axis — column H is blank on
    # every row, including Model Year.
    for variable in by_variable:
        assert sheet.cell(by_variable[variable], _C_SEQUENCE).value is None


def test_the_four_computed_columns_are_one_spill_each() -> None:
    """J/K/L/O are single dynamic arrays anchored at _FIRST_DATA_ROW, each
    sized by COLUMNS(Source_Data). That is the mechanism that resizes the
    block on a Source_Table retarget.

    They must go through Formula2. Writing a dynamic array through .Formula
    enters it as a legacy CSE range, which does NOT resize — it would look
    correct on the shipped dataset and silently reintroduce the truncation
    the spills exist to remove.
    """
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    for col in (_C_PERIOD_IN_USE, _C_LEVELS, _C_REF_IN_USE, _C_DESIGN_COLUMNS):
        formula = sheet.cell(_FIRST_DATA_ROW, col).api.Formula2
        assert formula is not None, col
        assert formula.startswith("=LET(nc,COLUMNS(Source_Data),"), col
        assert "MAP(SEQUENCE(nc),LAMBDA(i," in formula, col
        # No structured reference may survive — the block is table-free, so
        # one would never resolve.
        assert "[@" not in formula, col
        # Nor may the old row-arithmetic: the spill's position must not
        # determine which source column a row maps to.
        assert "ROW()-" not in formula, col


def test_spec_block_defaults_to_the_given_profile() -> None:
    """--regression-dataset life_expectancy must pre-fill its own defaults.

    Life Expectancy has 23 columns vs. Auto MPG's 11. The profile no longer
    sizes anything — the block's height follows COLUMNS(Source_Data) — but
    it still decides which rows arrive with shipped defaults. The shipped
    default is the curated four-driver model (Adult Mortality, Alcohol,
    percentage expenditure, Status); every other predictor is present in
    the block with Include off, ready to toggle on.
    """
    profile = SPEC_DATASET_PROFILES["life_expectancy"]
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet), profile)

    by_variable = {v: _FIRST_DATA_ROW + i for i, v in enumerate(profile.variables)}
    assert sheet.cell(by_variable["Life expectancy"], _C_ROLE).value == "Response (y)"
    assert sheet.cell(by_variable["Country"], _C_ROLE).value == "Identifier (Row Label)"
    assert sheet.cell(by_variable["Developed Country after 2013"], _C_ROLE).value == "Omit"
    assert sheet.cell(by_variable["Year"], _C_SEQUENCE).value is True

    # The curated, shipped-on predictors.
    adult_mortality_row = by_variable["Adult Mortality"]
    assert sheet.cell(adult_mortality_row, _C_ROLE).value == "Predictor (x)"
    assert sheet.cell(adult_mortality_row, _C_INCLUDE).value is True
    assert sheet.cell(adult_mortality_row, _C_TYPE).value == "Continuous"
    status_row = by_variable["Status"]
    assert sheet.cell(status_row, _C_ROLE).value == "Predictor (x)"
    assert sheet.cell(status_row, _C_INCLUDE).value is True
    assert sheet.cell(status_row, _C_TYPE).value == "Categorical"

    # A predictor the curated default leaves off — present in the block,
    # ready to toggle on for the screen/trim beats, but Include=False.
    schooling_row = by_variable["Schooling"]
    assert sheet.cell(schooling_row, _C_ROLE).value == "Predictor (x)"
    assert sheet.cell(schooling_row, _C_INCLUDE).value is False
    assert sheet.cell(schooling_row, _C_TYPE).value == "Continuous"


def test_spec_block_defaults_to_the_auto_mpg_profile_when_omitted() -> None:
    """Calling _write_spec_block with no profile keeps the shipped Auto MPG behavior."""
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    for offset, variable in enumerate(_VARIABLES):
        role, _, _ = _DEFAULT_SPEC.get(variable, _FALLBACK_SPEC)
        assert sheet.cell(_FIRST_DATA_ROW + offset, _C_ROLE).value == role, variable
    # One row past the profile is left blank rather than defaulted. A blank
    # Role contributes nothing, so this is a legal spec — and it is what a
    # retarget to a wider table leaves for the user to fill in.
    assert sheet.cell(_FIRST_DATA_ROW + _N_VARIABLES, _C_ROLE).value is None


def test_life_expectancy_and_production_lots_profiles_need_no_fallback() -> None:
    """Unlike Auto MPG (which intentionally leaves 3 candidate predictors on
    _FALLBACK_SPEC — see _DEFAULT_SPEC's docstring), every column of the two
    profiles this feature adds must have its own explicit default_spec
    entry, so retargeting to either dataset never silently depends on
    _write_spec_block's fallback tuple."""
    for name in ("life_expectancy", "production_lots"):
        profile = SPEC_DATASET_PROFILES[name]
        missing = [v for v in profile.variables if v not in profile.default_spec]
        assert missing == [], (name, missing)


def test_spec_dataset_profiles_cover_every_regression_dataset_choice() -> None:
    """Every profile's effective spec — default_spec entries falling back to
    _FALLBACK_SPEC exactly as _write_spec_block does — has exactly one
    Response row and Sequence flags restricted to its own variables."""
    for name, profile in SPEC_DATASET_PROFILES.items():
        responses = [
            variable
            for variable in profile.variables
            if profile.default_spec.get(variable, _FALLBACK_SPEC)[0]
            == "Response (y)"
        ]
        assert len(responses) == 1, (name, responses)
        assert profile.sequence_variables <= set(profile.variables), name


def test_levels_column_counts_raw_levels_without_dummy_levels() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    formula = cast(str, sheet.cell(_FIRST_DATA_ROW, _C_LEVELS).api.Formula2)
    # Must display L itself (1 for degenerate columns, which Dummy_Levels
    # signals as #N/A instead), so it counts UNIQUE directly.
    assert (
        'IF(OR(INDEX(rl,i)<>"Predictor (x)",INDEX(typ,i)<>"Categorical"),"",'
        in formula
    )
    assert "ROWS(UNIQUE(FILTER(" in formula
    assert "Dummy_Levels" not in formula
    # Sample_Include() is hoisted out of the MAP: it does not vary by row,
    # and the per-row version re-evaluated the whole mask once per row.
    assert "si,Sample_Include()," in formula
    # The map index selects the source column — no row arithmetic, so the
    # spill does not care which row it sits on.
    assert "col,INDEX(Source_Data,0,i)" in formula
    assert 'x,IF(col="","",col)' in formula  # blank normalization mirrored


def test_design_columns_audit_mirrors_the_constructors_own_skip_rules() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    formula = cast(
        str, sheet.cell(_FIRST_DATA_ROW, _C_DESIGN_COLUMNS).api.Formula2
    )
    # The audit has to agree with Predictor_Columns() by construction, not
    # by coincidence: same iteration predicate (Role/Include), same
    # Continuous-vs-Categorical split, same reference normalization, and
    # the same degenerate skip (Dummy_Levels' #N/A means the constructor
    # contributes nothing, so the audit must read 0 and not error).
    assert 'IF(INDEX(rl,i)<>"Predictor (x)","",' in formula
    assert "IF(INDEX(inc,i)<>TRUE,0," in formula
    assert 'kk,LAMBDA(x,IF(INDEX(typ,x)<>"Categorical",1,' in formula
    assert "COLUMNS(Dummy_Levels(" in formula
    assert 'IF(LEN(INDEX(refs,x)&"")=0,"",INDEX(refs,x))' in formula
    assert "si,Sample_Include()," in formula
    # The map index selects the source column, same mapping K and L use.
    assert "k,kk(i)," in formula
    # Reading the Levels display instead would make one display depend on
    # another; the audit reads the same closure the constructor reads.
    assert "Spec_Levels" not in formula

    # v3.1: the interaction term. ONE width helper serves both operands, so
    # the audit cannot disagree with the constructor about how wide a
    # categorical operand is, and the count is the pairwise product.
    assert formula.count("kk,LAMBDA(") == 1
    assert "k*kk(q)" in formula
    assert "k+ki))))))" in formula
    # Gating mirrors the constructor's mate() exactly: blank M, blank N, a
    # name matching no column, or a non-Predictor operand all contribute 0.
    # TOROW(Header_Names) is hoisted out of the MAP as `hdr`.
    assert "hdr,TOROW(Header_Names)," in formula
    assert "q,IFERROR(XMATCH(t,hdr),0)" in formula
    assert 'IF(OR(LEN(t&"")=0,LEN(o&"")=0,q=0),0,' in formula
    assert 'IF(INDEX(rl,q)<>"Predictor (x)",0,' in formula
    # Include is deliberately NOT tested on the operand — an excluded
    # Predictor operand is the flagged-amber marginality case, which builds.
    assert "INDEX(inc,q)" not in formula


def test_design_columns_audit_is_read_only_by_the_width_guard() -> None:
    # Spec_Design_Columns is a computed display: "display derives, never
    # feeds". No constructor closure may reference it — only the
    # pre-flight width guard, which is itself a display.
    #
    # Since Part 6.2 the guard is a defined name rather than an inline cell
    # formula, so it shows up in the reader list. That is the same single
    # reader this test always allowed, now spelled out: what still fails is
    # any CONSTRUCTOR appearing beside it, which would mean a display had
    # started feeding the model.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    band = "Spec_Design_Columns"
    readers = [
        item.Name.split("!", 1)[-1]
        for item in sheet.api.Names.items
        if band in item.RefersTo and item.Name.split("!", 1)[-1] != band
    ]
    assert readers == ["Design_Width_Status"]
    for formula in _all_written_formulas(sheet):
        assert band not in formula, formula


def test_interaction_bands_are_read_by_the_three_constructor_twins() -> None:
    # The v3.1 wiring release ends the RESERVED state established at v3.0
    # stage 3: M/N are now consumed by the constructor and its two twins.
    # Exactly those three — no other closure may read them, because an
    # interaction is a property of the design matrix, not of the row mask
    # or the row labels.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    expected = {
        "Predictor_Columns",
        "Constructed_Column_Names",
        "Constructed_Column_Transforms",
    }
    for band in ("Spec_Interaction_Term", "Spec_Interaction_Operation"):
        readers = {
            item.Name.split("!", 1)[-1]
            for item in sheet.api.Names.items
            if band in item.RefersTo and item.Name.split("!", 1)[-1] != band
        }
        assert readers == expected, (band, sorted(readers))

    # Sample_Include and Row_Labels must stay clear of them: the row mask is
    # what every spilled array is aligned to, and it cannot depend on a
    # declaration that only changes the matrix's width.
    for name in ("Sample_Include", "Row_Labels", "Response_Column"):
        refers_to = _refers_to(sheet, name)
        assert "Spec_Interaction" not in refers_to, name


def test_interaction_flags_key_on_the_named_operands_own_spec_row() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    r = _FIRST_DATA_ROW
    term = sheet.range(
        f"$M${r}:$M${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    # Error flags come FIRST with StopIfTrue, so a bad interaction typed
    # onto a non-Predictor row shows its error instead of being grayed out
    # by the hide-in-place rule that follows.
    assert len(term) == 2
    invalid, marginality = term

    # Red: the name matches no variable (j=0), or the row it names is not
    # a Predictor. MAX(j,1) keeps the INDEX well-formed on a miss, so the
    # AND yields FALSE rather than an #N/A that would disable the rule.
    assert "j,IFERROR(XMATCH($M4,hdr),0)" in invalid.Formula1
    assert "p,MAX(j,1)" in invalid.Formula1
    assert 'OR(j=0,INDEX(TAKE(Spec_Role,nc),p)<>"Predictor (x)")' in invalid.Formula1
    assert invalid.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert invalid.StopIfTrue is True

    # Amber: a marginality violation — the named operand IS a Predictor
    # but is excluded. Allowed and flagged, never blocked.
    assert 'INDEX(TAKE(Spec_Role,nc),p)="Predictor (x)"' in marginality.Formula1
    assert "INDEX(TAKE(Spec_Include,nc),p)<>TRUE" in marginality.Formula1
    assert marginality.Interior.Color == excel_color(CF_YELLOW_FILL)
    assert marginality.StopIfTrue is True

    # The hide-in-place rule covers M and N together, so it lands on the
    # M:N band rather than M's own — and it is registered AFTER both error
    # rules, which is what lets their StopIfTrue outrank it.
    hide = sheet.range(
        f"$M${r}:$N${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    # Two rules on this band, in priority order: hide-in-place first, then
    # the input-band fill LAST so every rule above it still wins. The fill
    # rule replaced the per-row format_input() calls, which painted only the
    # build-time profile's rows and so left a retarget's new rows unpainted.
    assert [c.Formula1 for c in hide] == [
        f'=$B{r}<>"Predictor (x)"',
        f"=ROW()-{_FIRST_DATA_ROW - 1}<=COLUMNS(Source_Data)",
    ]
    assert hide[0].Font.Color == excel_color(INPUT_COLOR)
    assert hide[1].Interior.Color == excel_color(INPUT_COLOR)

    # Red on N: a reciprocal declaration under a SYMMETRIC operation. Ratio
    # is excluded (B/A is a different column from A/B), and a row naming
    # itself is excluded (self x self under Product is a quadratic term).
    operation = sheet.range(
        f"$N${r}:$N${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert len(operation) == 1
    reciprocal = operation[0].Formula1
    assert f'OR($N{r}="Product",$N{r}="Difference")' in reciprocal
    assert "Ratio" not in reciprocal
    assert "j>0,j<>i" in reciprocal
    assert "INDEX(TAKE(Spec_Interaction_Term,nc),p)=INDEX(hdr,1,q)" in reciprocal
    assert f"INDEX(TAKE(Spec_Interaction_Operation,nc),p)=$N{r}" in reciprocal
    # Both indices are clamped: the CF band runs past the spec rows, and an
    # unclamped INDEX there would error the rule into silence.
    assert f"i,ROW()-{_FIRST_DATA_ROW - 1}" in reciprocal
    assert "q,MIN(MAX(i,1),nc)" in reciprocal
    assert operation[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)


def test_reference_in_use_echoes_e_or_shows_the_sorted_default() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    r = _FIRST_DATA_ROW
    formula = cast(str, sheet.cell(r, _C_REF_IN_USE).api.Formula2)
    # Same relevance guard as the Levels display: Categorical Predictors only.
    assert (
        'IF(OR(INDEX(rl,i)<>"Predictor (x)",INDEX(typ,i)<>"Categorical"),"",'
        in formula
    )
    # An explicit Reference Level is echoed verbatim (its invalid-reference
    # CF carries the error signal); blank Reference Level falls through to
    # the default.
    assert 'IF(INDEX(refs,i)<>"",INDEX(refs,i),' in formula
    # The default mirrors Dummy_Levels: first sorted level over the
    # mask-included sample, with the same blank normalization. NOT a
    # Dummy_Levels call — that returns the retained levels, i.e. everything
    # EXCEPT the reference.
    assert "INDEX(SORT(UNIQUE(FILTER(" in formula
    assert "Dummy_Levels" not in formula
    assert "si,Sample_Include()," in formula
    assert "col,INDEX(Source_Data,0,i)" in formula
    assert 'x,IF(col="","",col)' in formula
    # Empty masked sample degrades to blank (H shows 0 and flags red there).
    assert formula.endswith(',"")))))))')


def test_dropdowns_cover_exactly_the_list_columns() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    validated = {
        key: rng.api.Validation
        for key, rng in sheet.ranges.items()
        if rng.api.Validation.rules
    }
    r = _FIRST_DATA_ROW
    assert set(validated) == {
        ((r, 2), (16000, 2)),  # B Role
        ((r, 3), (16000, 3)),  # C Include
        ((r, 4), (16000, 4)),  # D Type
        ((r, 7), (16000, 7)),  # G Transform (None or either Log token)
        ((r, 8), (16000, 8)),  # H Sequence (TRUE or blank)
        ((r, 13), (16000, 13)),  # M Interaction Term (the variable-name spill)
        ((r, 14), (16000, 14)),  # N Interaction Operation (closed axis)
    }
    formulas = {
        key[0][1]: validation.rules[0]["Formula1"]
        for key, validation in validated.items()
    }
    assert formulas[2] == "Response (y),Predictor (x),Identifier (Row Label),Filter,Omit,Fixed Effects"
    assert formulas[3] == "TRUE,FALSE"
    assert formulas[4] == "Continuous,Categorical"
    assert formulas[7] == "None,Log,Log (drop ≤ 0)"
    assert formulas[8] == "TRUE"
    # M sources its list from the variable-name spill at A4, so the offered
    # names resize with the dataset instead of a fixed range going stale.
    assert formulas[13] == f"=$A${r}#"
    assert formulas[14] == "Product,Difference,Ratio"
    for validation in validated.values():
        assert validation.delete_count == 1
        assert validation.rules[0]["Type"] == 3  # xlValidateList
        assert validation.IgnoreBlank is True


def test_conditional_formats_cover_cascading_relevance_degeneracy_and_reference() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    r = _FIRST_DATA_ROW
    off = _FIRST_DATA_ROW - 1
    # Role-keyed relevance: the per-Predictor inputs (C–F) hide behind their
    # own INPUT_COLOR fill, the Categorical displays (K–L) hide behind
    # white (unfilled computed cells); H–J are deliberately NOT in these
    # bands. Every range runs out to _VALIDATION_LAST_ROW so a row a
    # Source_Table retarget brings into the block is already covered.
    role_keyed_input = sheet.range(
        f"$C${r}:$F${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in role_keyed_input] == [f'=$B{r}<>"Predictor (x)"']
    assert role_keyed_input[0].Font.Color == excel_color(INPUT_COLOR)

    # G (Transform) is relevant on Predictor OR Response rows (unlike C–F,
    # Predictor-only) — its own cascading-relevance rule, plus a red flag
    # for the disallowed Categorical x Log combination.
    transform_col = sheet.range(
        f"$G${r}:$G${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in transform_col] == [
        f'=AND($B{r}<>"Predictor (x)",$B{r}<>"Response (y)")',
        # Categorical x Log — either token, since dummy columns are never
        # logged under either.
        f'=AND($B{r}="Predictor (x)",$D{r}="Categorical",'
        f'OR($G{r}="Log",$G{r}="Log (drop ≤ 0)"))',
        # Strict Log on a column that holds a zero or a negative among the
        # rows the model would fit. Equality against "Log" alone, NOT _is_log:
        # the whole point of the filtering token is that it does not fire this.
        # Sample_Include(FALSE) is the mask before the positivity layer, so
        # the count is of rows the fit would otherwise have used.
        f'=AND($G{r}="Log",'
        f'OR($B{r}="Response (y)",'
        f'AND($B{r}="Predictor (x)",$C{r}=TRUE,$D{r}="Continuous")),'
        "SUMPRODUCT(--Sample_Include(FALSE),"
        f"--IFERROR((INDEX(Source_Data,0,ROW()-{off})+0)<=0,FALSE))>0)",
    ]
    assert transform_col[0].Font.Color == excel_color(INPUT_COLOR)
    for rule in transform_col[1:]:
        assert rule.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
        assert rule.Font.Color == excel_color(CF_DARK_RED_TEXT)

    role_keyed_computed = sheet.range(
        f"$K${r}:$L${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in role_keyed_computed] == [f'=$B{r}<>"Predictor (x)"']
    assert role_keyed_computed[0].Font.Color == excel_color((255, 255, 255))

    # Sequence-keyed relevance: H–I hide behind INPUT_COLOR, J (computed)
    # behind white, on rows that are not the sequence axis — keyed on the
    # flag itself, not on Role.
    seq_input = sheet.range(
        f"$H${r}:$I${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in seq_input] == [f"=$H{r}<>TRUE"]
    assert seq_input[0].Font.Color == excel_color(INPUT_COLOR)

    seq_computed = sheet.range(
        f"$J${r}:$J${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in seq_computed] == [f"=$H{r}<>TRUE"]
    assert seq_computed[0].Font.Color == excel_color((255, 255, 255))

    # Multi-flag error: red on every flagged H cell at two-plus flags.
    multi = sheet.range(f"$H${r}:$H${_VALIDATION_LAST_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in multi] == [
        f"=AND($H{r}=TRUE,"
        "SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))>1)"
    ]
    assert multi[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert multi[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Per-row override flagging is intentionally NOT applied on column J
    # of the spec block. The Period In Use cells stay plain; the override
    # verdict lives on the Sequence Spacing block (rows 31–34) where the
    # user actually reads it, not on the spec block.

    # Degeneracy flag: red K when an INCLUDED Categorical Predictor has
    # L <= 1 — the constructor contributes zero columns for it (visible
    # degradation, not silent omission). N() coerces "" to 0.
    degenerate = sheet.range(
        f"$K${r}:$K${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in degenerate] == [
        f'=AND($B{r}="Predictor (x)",$C{r}=TRUE,$D{r}="Categorical",N($K{r})<=1)'
    ]
    assert degenerate[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert degenerate[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Invalid reference: the constructor's exact skip condition, tested
    # directly — a membership test against Dummy_Levels' output would
    # false-positive on the default reference itself.
    invalid = sheet.range(f"$E${r}:$E${_VALIDATION_LAST_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in invalid] == [
        f'=AND($E{r}<>"",ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-{off}),'
        f"$E{r},Sample_Include())))"
    ]
    assert invalid[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert invalid[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def _feedback_sheet() -> RecordingSheet:
    """A sheet carrying the whole rows 1-2 band, as the Regression writer builds it."""
    from lambda_catalog.write_spec_block import (
        _write_sequence_status,
        _write_spec_feedback,
    )

    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_feedback(_as_xw_sheet(sheet))
    _write_sequence_status(_as_xw_sheet(sheet))
    return sheet


def test_every_status_line_sits_in_the_spec_column_it_is_about() -> None:
    """The rule the whole band is arranged by, asserted as one statement.

    Role cardinality above Role, the Log domain above Transform, Sequence
    cardinality above Sequence, the spacing verdict above Sequence Period.
    Before this they were scattered across whichever cells happened to be
    free — Sequence's error above Reference Level, Fixed Effects' above Role,
    the width guard above Interaction Term — so a message's position said
    nothing about its subject.
    """
    sheet = _feedback_sheet()

    # Since Part 6.2 three of the four cells hold a CALL rather than the logic,
    # so the check runs in two hops: the cell names its function, and that
    # function's catalog body reads the spec band the column is about. Both
    # halves are needed. The cell alone would pass a wall of nested IFs that
    # happened to mention the right token, which is the state this test was
    # written against; the function name alone would pass a cell wired to a
    # plausibly-named wrong function, which is the state moving the logic out
    # newly makes possible. I2's spacing verdict is still inline, so it has no
    # second hop — its marker is matched in the cell.
    for col, function_name, marker in (
        (_C_ROLE, "Role_Status", "Response (y)"),
        (_C_TRANSFORM, "Log_Domain_Status", "Spec_Transform"),
        (_C_SEQUENCE, "Sequence_Status", "Spec_Sequence"),
        (_C_SEQUENCE_PERIOD, None, "Spec_Period_In_Use"),
    ):
        formula = cast(str, sheet.cell(_FEEDBACK_STATUS_ROW, col).api.Formula2)
        assert formula is not None, col
        if function_name is None:
            assert marker in formula, (col, marker)
        else:
            assert formula == f"={function_name}()", (col, function_name)
            assert marker in _catalog_body(function_name), (col, marker)
        # Every status cell wraps: row 2 has no runway between columns, so a
        # message grows the row rather than truncating against its neighbour.
        assert sheet.cell(_FEEDBACK_STATUS_ROW, col).api.WrapText is True, col
        assert sheet.cell(_FEEDBACK_STATUS_ROW, col).api.Font.Bold is True, col
        # ...and carries the long form as a hover note.
        assert sheet.cell(_FEEDBACK_STATUS_ROW, col).api.Comment.Text, col

    # The status cells that do not own a column are empty. E1 in particular
    # carries no Sequence status — H2 is the single home for it.
    for row, col in ((1, _C_ROLE), (1, _C_REFERENCE), (2, _C_LABEL)):
        assert sheet.cell(row, col).value is None, (row, col)
        assert sheet.cell(row, col).api.Formula2 is None, (row, col)


def test_role_status_ranks_response_cardinality_above_fixed_effects() -> None:
    """B2 — one cell, three conditions, most severe first.

    Exactly one Response is required, so zero and two-plus are both errors and
    they need different instructions. Sequence and Fixed Effects allow zero, so
    only a second row of either is flagged; Fixed Effects is checked here
    because Role is the column it is declared in.
    """
    sheet = _feedback_sheet()

    status = sheet.cell(_FEEDBACK_STATUS_ROW, _C_ROLE)
    assert cast(str, status.api.Formula2) == "=Role_Status()"
    # The logic moved into the catalog (Part 6.2), so the three conditions and
    # their order are asserted where they now live. Same statement, one
    # indirection further out: the count sub-formulas are still built in
    # Python, so this also pins the two halves together across the JSON gap.
    assert _catalog_body("Role_Status") == (
        f"=IF({_RESPONSE_COUNT_FORMULA}=0,"
        '"ERROR: no Response (y) row — mark the variable being modeled.",'
        f"IF({_RESPONSE_COUNT_FORMULA}>1,"
        '"ERROR: multiple Response (y) rows — mark exactly one.",'
        f"IF({_FIXED_EFFECTS_COUNT_FORMULA}>1,"
        '"ERROR: multiple Fixed Effects rows — mark at most one.",'
        '"")))'
    )
    # Severity order is inside the formula, so the cell needs only one rule.
    conditions = sheet.range("$B$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ['=$B$2<>""']
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_log_domain_status_reports_the_poisoned_column_then_the_dropped_count() -> None:
    """G2 — the two Log states, red outranking amber.

    RED names the variable, its count of non-positive rows IN THE SAMPLE, and
    the token that would exclude them: with strict Log the fit is #N/A
    everywhere, and the fix is a different dropdown value the user has no way
    to guess from a sheet full of errors.

    AMBER is Log (drop ≤ 0) doing its job — not a problem, but the sample is
    now smaller than the data and that must never be invisible.

    StopIfTrue on red matters because a spec can declare both tokens on
    different variables, making both states true at once.
    """
    sheet = _feedback_sheet()

    assert (
        cast(str, sheet.cell(_FEEDBACK_STATUS_ROW, _C_TRANSFORM).api.Formula2)
        == "=Log_Domain_Status()"
    )
    formula = _catalog_body("Log_Domain_Status")
    # Only the STRICT token is counted as poisoned — the whole point of the
    # filtering token is that these rows leaving is intended.
    assert 'INDEX(trn,j)="Log"' in formula
    assert "Use Log (drop ≤ 0)." in formula
    # The eligibility test mirrors Sample_Include's own branch, so a Log left
    # on an Identifier or an excluded row is inert and uncounted.
    assert (
        'elig,((rl="Response (y)")+((rl="Predictor (x)")*(inc=TRUE)'
        '*(typ="Continuous")))>0'
    ) in formula
    # Sample_Include(FALSE) — the mask BEFORE the positivity layer — is what
    # makes both halves count the same population. It is bound ONCE, to `base`,
    # and the excluded-row count reuses that binding rather than calling again:
    # Sample_Include is a REDUCE over the spec, so a second call is a second
    # full evaluation for a value already in hand. (Copilot review on #222;
    # deferred from that PR to keep the body a byte-identical move of the
    # formula it replaced, and landed here where it takes the same Excel pass
    # as the rest of the v3.2 spike.)
    assert "base,Sample_Include(FALSE)" in formula
    assert formula.count("Sample_Include(FALSE)") == 1
    # The DEFAULT-mask half reads the materialized spill via Fit_Sample_Include()
    # — the same reader the ~30 engine call sites use — but sums it with `--`
    # rather than N(). N() of a range-returning thunk (the reader, whether
    # spelled `#`/ANCHORARRAY or OFFSET) collapses to the top-left cell, so
    # SUMPRODUCT(N(Fit_Sample_Include())) returns 1 and the amber fires
    # "n-1 rows excluded" on EVERY sheet regardless of transforms. `--` coerces
    # a range OR an array to a summable 1/0 array, so it is robust to the
    # reader's form (and to Sample_Include()'s promotion to the reader on main,
    # which would make N(Sample_Include()) collapse the same way).
    assert "d,SUMPRODUCT(N(base))-SUMPRODUCT(--(Fit_Sample_Include()))" in formula
    assert '" rows excluded: Log of ≤ 0"' in formula
    # Guard the regression: the amber must never sum the reader with N().
    assert "N(Fit_Sample_Include())" not in formula

    conditions = sheet.range("$G$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == [
        '=ISNUMBER(SEARCH("ERROR",$G$2))',
        '=$G$2<>""',
    ]
    red, amber = conditions
    assert red.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert red.Font.Color == excel_color(CF_DARK_RED_TEXT)
    assert red.StopIfTrue is True
    assert amber.Interior.Color == excel_color(CF_YELLOW_FILL)
    assert amber.Font.Color == excel_color(CF_DARK_YELLOW_TEXT)
    assert amber.StopIfTrue is False


def test_sequence_status_line_validates_zero_or_one_flags() -> None:
    """H2 — back in the Sequence column, and the only copy.

    This cell was written by a function that stopped being called, so the
    message lived at E1 (above Reference Level) instead. Both existed in the
    source; only E1 reached the sheet.
    """
    sheet = _feedback_sheet()

    status = sheet.cell(_FEEDBACK_STATUS_ROW, _C_SEQUENCE)
    assert status.api.Formula2 == "=Sequence_Status()"
    assert _catalog_body("Sequence_Status") == (
        "=IF(SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))>1,"
        '"ERROR: multiple Sequence rows — mark at most one.","")'
    )
    assert status.api.Font.Bold is True

    conditions = sheet.range("$H$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ['=$H$2<>""']
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_fixed_effects_status_block_shows_variable_groups_and_absorbed_df() -> None:
    """J1:L2 — labels over values, and both disappear when there is no FE row.

    The values return "" (not a literal "n/a") and the labels white out
    with them, so an inactive block leaves no trace instead of three cells
    of filler.
    """
    sheet = _feedback_sheet()

    for col, label in (
        (_C_PERIOD_IN_USE, "FE Variable"),
        (_C_LEVELS, "FE Groups"),
        (_C_REF_IN_USE, "FE df Absorbed"),
    ):
        cell = sheet.cell(1, col)
        assert cell.value == label, (col, label)
        assert cell.api.Font.Bold is True

        # Label and value share one hide rule over the two-row range.
        letter = col_letter(col)
        hide = sheet.range(f"${letter}$1:${letter}$2").api.FormatConditions.items
        assert [c.Formula1 for c in hide] == [
            f"={_FIXED_EFFECTS_COUNT_FORMULA}=0"
        ], col
        assert hide[0].Font.Color == excel_color((255, 255, 255)), col

    variable = cast(str, sheet.cell(2, _C_PERIOD_IN_USE).api.Formula2)
    assert variable.startswith(f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"",')
    assert 'XMATCH("Fixed Effects",TAKE(Spec_Role,COLUMNS(Source_Data)))' in variable

    groups = sheet.cell(2, _C_LEVELS).api.Formula2
    assert groups == (
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"",Absorbed_Degrees_Of_Freedom()+1)'
    )

    absorbed = sheet.cell(2, _C_REF_IN_USE).api.Formula2
    assert absorbed == (
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"",Absorbed_Degrees_Of_Freedom())'
    )


def test_spec_feedback_writes_delta_count_verdict_with_priority_cf() -> None:
    """The P/Q spectrum and the I1/I2 spacing verdict.

    The spectrum moved off rows 1-2 onto the spec block's own rows — headers
    at P3/Q3 beside the spec headers, body from P4 beside the spec data — so
    it reads as a second table rather than a third row-2 thing, and so O2's
    width-guard message gains P2:Q2 as the only overflow runway anything on
    row 2 has.

    I1/I2: the combined switch — one cell, one message, with red CF outranking
    yellow via StopIfTrue.
    """
    sheet = _feedback_sheet()

    for col, label in (
        (_C_FEEDBACK_DELTA, "Δ"),
        (_C_FEEDBACK_COUNT, "Count"),
    ):
        cell = sheet.cell(_HEADER_ROW, col)
        assert cell.value == label, (col, label)
    assert (
        sheet.range(
            (_HEADER_ROW, _C_FEEDBACK_DELTA), (_HEADER_ROW, _C_FEEDBACK_COUNT)
        ).api.Font.Bold
        is True
    )
    # Nothing left on rows 1-2 of P:Q — that space is O2's runway now.
    for row in (1, 2):
        for col in (_C_FEEDBACK_DELTA, _C_FEEDBACK_COUNT):
            assert sheet.cell(row, col).value is None, (row, col)
            assert sheet.cell(row, col).api.Formula2 is None, (row, col)

    # P4 spectrum spill, aligned with the spec block's first data row.
    assert sheet.cell(_FIRST_DATA_ROW, _C_FEEDBACK_DELTA).api.Formula2 == (
        '=IFERROR(Sequence_Delta_Spectrum(),"")'
    )

    # "Spacing Verdict", not "Verdict" — the sheet has several verdicts now,
    # and this one is specifically about how the Sequence axis is spaced.
    assert sheet.cell(1, _C_SEQUENCE_PERIOD).value == "Spacing Verdict"
    assert sheet.cell(1, _C_SEQUENCE_PERIOD).api.Font.Bold is True

    # The label and the spectrum headers white out together when no axis is
    # declared, on the same gate: COUNT of Sequence_Deltas()'s #N/A is 0.
    for address in ("$I$1", f"$P${_HEADER_ROW}:$Q${_HEADER_ROW}"):
        hide = sheet.range(address).api.FormatConditions.items
        assert [c.Formula1 for c in hide] == [
            "=NOT(COUNT(Sequence_Deltas())>0)"
        ], address
        assert hide[0].Font.Color == excel_color((255, 255, 255)), address

    # I2 combined switch formula — the priority-ordered switch replacing
    # the four A31:A34 verdict cells. Reads Spec_Period_In_Use via the
    # structured INDEX/XMATCH pair, branches in priority order
    # (off-grid > regularity > no-natural > calendar), and emits the
    # declared message constant for the first hit.
    i2 = cast(str, sheet.cell(2, _C_SEQUENCE_PERIOD).api.Formula2)
    assert "INDEX(Spec_Period_In_Use" in i2
    assert "XMATCH(TRUE,Spec_Sequence,0)" in i2
    assert f'"{_MSG_OFF_GRID}"' in i2
    assert f'"{_MSG_REGULARITY}"' in i2
    assert f'"{_MSG_NO_NATURAL}"' in i2
    assert f'"{_MSG_CALENDAR}"' in i2

    # Priority CF: red outranks yellow on the same cell via StopIfTrue.
    # The red rule is added first with stop_if_true=True; yellow follows.
    conditions = sheet.range("$I$2").api.FormatConditions.items
    assert len(conditions) == 2
    red, yellow = conditions
    assert red.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert red.Font.Color == excel_color(CF_DARK_RED_TEXT)
    assert red.StopIfTrue is True
    assert yellow.Interior.Color == excel_color(CF_YELLOW_FILL)
    assert yellow.Font.Color == excel_color(CF_DARK_YELLOW_TEXT)
    assert yellow.StopIfTrue is False


_CAT_INCLUDED = (
    "SUMPRODUCT("
    'N(TAKE(Spec_Role,COLUMNS(Source_Data))="Predictor (x)"),'
    "N(TAKE(Spec_Include,COLUMNS(Source_Data))=TRUE),"
    'N(TAKE(Spec_Type,COLUMNS(Source_Data))="Categorical"))>0'
)


def test_allow_intercept_names_the_row2_toggle_cell() -> None:
    sheet = _named_sheet()
    assert _refers_to(sheet, "Allow_Intercept") == f"='{SHEET_NAME}'!$C$2"


def test_intercept_control_is_a_toggle_with_coupling_cf() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_intercept_control(_as_xw_sheet(sheet))

    # C1 label (bold), C2 toggle prefilled TRUE with input styling. The label
    # sits directly above the cell it names.
    assert sheet.cell(_FEEDBACK_LABEL_ROW, _C_INCLUDE).value == "Intercept"
    assert sheet.cell(_FEEDBACK_LABEL_ROW, _C_INCLUDE).api.Font.Bold is True
    assert sheet.cell(_INTERCEPT_ROW, _C_LABEL).value is None
    toggle_cell = sheet.cell(_INTERCEPT_ROW, _C_INCLUDE)
    assert toggle_cell.value is True
    assert toggle_cell.color == INPUT_COLOR

    # TRUE/FALSE dropdown on the single toggle cell.
    validation = sheet.range(
        (_INTERCEPT_ROW, _C_INCLUDE), (_INTERCEPT_ROW, _C_INCLUDE)
    ).api.Validation
    assert validation.delete_count == 1
    assert validation.rules[0]["Type"] == 3  # xlValidateList
    assert validation.rules[0]["Formula1"] == "TRUE,FALSE"
    assert validation.IgnoreBlank is True

    # Coupling CF, on C2: red (toggle FALSE while an included Categorical
    # needs the intercept) added first with StopIfTrue so it outranks the
    # hide-in-place "required-here" rule; that rule applies whenever a
    # Categorical is in, hiding C2 behind its own INPUT_COLOR fill. A third
    # rule flags the opposite state: toggle TRUE while Fixed Effects is
    # active (the resulting "Intercept" row is a nuisance artifact, not a
    # numerical error — see the DF_Absorbed threading, which accounts for it
    # correctly either way).
    conditions = sheet.range("$C$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == [
        f"=AND($C$2=FALSE,{_CAT_INCLUDED})",
        f"={_CAT_INCLUDED}",
        f"=AND($C$2=TRUE,{_FIXED_EFFECTS_COUNT_FORMULA}>0)",
    ]
    red, hidden, fe_red = conditions
    assert red.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert red.Font.Color == excel_color(CF_DARK_RED_TEXT)
    assert red.StopIfTrue is True
    assert hidden.Font.Color == excel_color(INPUT_COLOR)
    assert hidden.StopIfTrue is False
    assert fe_red.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert fe_red.Font.Color == excel_color(CF_DARK_RED_TEXT)


_RESPONSE_NAME = (
    "LET(n_c,COLUMNS(Source_Data),"
    'p,XMATCH("Response (y)",TAKE(Spec_Role,n_c)),'
    "h,INDEX(TOROW(Header_Names),p),"
    'IFERROR(IF(OR(INDEX(TAKE(Spec_Transform,n_c),p)="Log",INDEX(TAKE(Spec_Transform,n_c),p)="Log (drop ≤ 0)"),"Ln("&h&")",h),"(none)"))'
)


def test_audit_row_is_bold_label_value_pairs_with_response_count_cf() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_audit_row(_as_xw_sheet(sheet))

    expected = [
        (19, 20, "k", '=IFERROR(COLUMNS(Predictor_Columns()),"(empty model)")'),
        (22, 23, "rows", '=IFERROR(ROWS(Predictor_Columns()),"(empty model)")'),
        (25, 26, "response", f"={_RESPONSE_NAME}"),
        (
            27,
            28,
            "responses",
            '=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))',
        ),
        (29, 30, "included rows", "=SUMPRODUCT(N(Sample_Include()))"),
        (
            31,
            32,
            "sequence flags",
            "=SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))",
        ),
        (
            33,
            34,
            "fixed effects",
            '=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))',
        ),
        (35, 36, "FE absorbed df", "=Absorbed_Degrees_Of_Freedom()"),
    ]
    assert list(_AUDIT_PAIRS) == [(lc, vc) for lc, vc, _, _ in expected]
    assert _AUDIT_ROW == 1
    for label_col, value_col, label, formula in expected:
        # No audit cell may land on a width-2 break column (U=21, X=24).
        assert label_col not in (_C_BREAK_LEFT, _C_BREAK_MID)
        assert value_col not in (_C_BREAK_LEFT, _C_BREAK_MID)
        assert sheet.cell(1, label_col).value == label
        assert sheet.cell(1, value_col).api.Formula2 == formula
        assert (
            sheet.range((1, label_col), (1, value_col)).api.Font.Bold is True
        )

    # Exactly-one-Response validation: red CF on the responses count cell (AB=28).
    conditions = sheet.range("$AB$1").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ["=N($AB$1)<>1"]
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Zero-or-one-Sequence validation: red CF only at two-plus flags (AF=32).
    seq_conditions = sheet.range("$AF$1").api.FormatConditions.items
    assert [c.Formula1 for c in seq_conditions] == ["=N($AF$1)>1"]
    assert seq_conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert seq_conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Zero-or-one-Fixed-Effects validation: red CF only at two-plus (AH=34).
    fe_conditions = sheet.range("$AH$1").api.FormatConditions.items
    assert [c.Formula1 for c in fe_conditions] == ["=N($AH$1)>1"]
    assert fe_conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert fe_conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_filtered_zones_filter_by_the_mask_and_degrade_gracefully() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_filtered_zones(_as_xw_sheet(sheet))

    # O and R: narrow visual breaks, same width as the L gap.
    assert sheet.range((1, _C_BREAK_LEFT)).column_width == 2
    assert sheet.range((1, _C_BREAK_MID)).column_width == 2

    # Header-row labels: static labels over the two Row Labels columns, the
    # derived response name over filtered y, the twin strip over the
    # matrix. All bold like the spec headers.
    assert sheet.cell(_HEADER_ROW, _C_FILTERED_LABELS).value == "Row Labels"
    assert sheet.cell(_HEADER_ROW, _C_MATRIX_LABELS).value == "Row Labels"
    assert sheet.cell(_HEADER_ROW, _C_FILTERED_Y).api.Formula2 == (
        f'="y: "&{_RESPONSE_NAME}'
    )
    assert sheet.cell(_HEADER_ROW, _C_MATRIX_START).api.Formula2 == (
        '=IFERROR(Constructed_Column_Names(),"(empty model)")'
    )
    assert (
        sheet.range(
            (_HEADER_ROW, _C_FILTERED_LABELS), (_HEADER_ROW, _C_MATRIX_START)
        ).api.Font.Bold
        is True
    )

    # First-data-row spills: the ONLY row-filtering on the sheet, every one
    # wrapped so an empty model degrades to the documented string
    # instead of leaking a raw #CALC!.
    expected_spills = {
        _C_FILTERED_LABELS: "Row_Labels()",
        _C_FILTERED_Y: "Response_Column()",
        _C_MATRIX_LABELS: "Row_Labels()",
        _C_MATRIX_START: "Predictor_Columns()",
    }
    for col, source in expected_spills.items():
        assert sheet.cell(_FIRST_DATA_ROW, col).api.Formula2 == (
            f'=IFERROR(FILTER({source},Sample_Include()),"(empty model)")'
        ), source


def test_interaction_header_symbols_match_the_catalog_formula() -> None:
    # _INTERACTION_HEADER_SYMBOLS is what the Python QC mirror renders;
    # Constructed_Column_Names() is what the sheet renders. They are two
    # spellings of one rule, so drift between them would make the oracle
    # disagree with the sheet about a column NAME while agreeing about its
    # values — a mismatch that reads like a formula bug and is not one.
    sheet = _named_sheet()
    names = _refers_to(sheet, "Constructed_Column_Names")

    for operation, symbol in _INTERACTION_HEADER_SYMBOLS.items():
        assert f'"{operation}","{symbol}"' in names, operation
    # The fall-through: an operation that is none of the three still yields a
    # header, so the strip stays exactly as wide as the design matrix, and it
    # is visibly not one of the three.
    assert f'"{_INTERACTION_HEADER_UNKNOWN}")' in names


def test_interaction_header_symbols_are_distinct_and_operation_specific() -> None:
    # The point of the change: one symbol per operation, none of them a
    # substring of another, and none of them a character that can appear in a
    # source column name (a hyphen can, which is why Difference uses U+2212
    # MINUS SIGN and not "-").
    symbols = list(_INTERACTION_HEADER_SYMBOLS.values())
    assert len(set(symbols)) == len(symbols)
    assert set(_INTERACTION_HEADER_SYMBOLS) == {"Product", "Difference", "Ratio"}
    assert _INTERACTION_HEADER_UNKNOWN not in symbols
    for symbol in symbols:
        assert "-" not in symbol, symbol
        assert ":" not in symbol, symbol
        for other in symbols:
            if other != symbol:
                assert symbol.strip() not in other, (symbol, other)


def test_status_lambda_messages_are_the_ones_the_guard_oracle_expects() -> None:
    """The sheet says what the oracle predicts, pinned across the JSON gap.

    The guard-state QC compares each row-2 cell's text against a Python mirror
    in analyze_regression_guard_states. Before Part 6.2 both halves were
    Python, so an edit to one was an edit to a shared constant or an obvious
    two-place change. Now the sheet's half is a JSON string literal no import
    can reach, and the two could drift silently — the message would change,
    the oracle would keep predicting the old text, and the mismatch would only
    surface on a machine with Excel running Layer 2.

    Pinning them here means a message edit fails in the unit suite, in CI,
    against the exact string the QC harness will look for.
    """
    from lambda_catalog.analyze_regression_guard_states import (
        _FIXED_EFFECTS_MULTI_ROW_ERROR,
        _MULTIPLE_RESPONSES_ERROR,
        _NO_RESPONSE_ERROR,
        _SEQUENCE_MULTI_FLAG_ERROR,
    )

    role = _catalog_body("Role_Status")
    for expected in (
        _NO_RESPONSE_ERROR,
        _MULTIPLE_RESPONSES_ERROR,
        _FIXED_EFFECTS_MULTI_ROW_ERROR,
    ):
        assert f'"{expected}"' in role, expected

    assert f'"{_SEQUENCE_MULTI_FLAG_ERROR}"' in _catalog_body("Sequence_Status")

    # The Log messages are built around live counts, so the oracle formats
    # them rather than holding them whole; the fixed fragments either side of
    # the interpolated values are what both sides have to agree on.
    log = _catalog_body("Log_Domain_Status")
    assert '" values ≤ 0 under Log — the fit is #N/A. Use Log (drop ≤ 0)."' in log
    assert '" rows excluded: Log of ≤ 0"' in log

    # The width guard's oracle deliberately mirrors only the leading token —
    # the full message embeds counts and prose — so that is what is pinned.
    width = _catalog_body("Design_Width_Status")
    assert '"ERROR: ' in width
    assert '"WARNING: ' in width


def test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform() -> None:
    """The two spellings live in Python AND in lambda_functions.json.

    No import can bridge that gap: the catalog bodies are JSON string data, so
    the token appears there as a literal. Five closures read Spec_Transform,
    and a rename or a half-applied edit that updated only some of them would
    leave the sheet silently treating one token as unrecognized — the column
    would fit raw, unlogged, with no error anywhere. This asserts the two
    tokens travel together through every body that mentions either.

    Sample_Include is the deliberate exception in the other direction: it must
    mention ONLY the filtering token, since plain Log not filtering is the
    whole distinction between them.
    """
    import json
    from pathlib import Path

    from lambda_catalog.write_spec_block import _TRANSFORM_LOG, _TRANSFORM_LOG_DROP

    document = json.loads(
        (Path(__file__).resolve().parents[1] / "lambda_functions.json").read_text(
            encoding="utf-8"
        )
    )
    bodies = {
        entry["name"]: entry["formula_display"]
        for entry in document["functions"]
        if "Spec_Transform" in entry.get("formula_display", "")
    }

    assert set(bodies) == {
        "Constructed_Column_Names",
        "Constructed_Column_Transforms",
        "Log_Domain_Status",
        "Model_Formula",
        "Predictor_Columns",
        "Response_Column",
        # The v3.2 name-promotion moved the REDUCE body (which reads
        # Spec_Transform) from Sample_Include to the Sample_Include_Calc leaf;
        # the public Sample_Include name is now a reader over the spill and no
        # longer mentions Spec_Transform.
        "Sample_Include_Calc",
    }

    for name, body in bodies.items():
        if name == "Sample_Include_Calc":
            assert _TRANSFORM_LOG_DROP in body
            assert f'="{_TRANSFORM_LOG}"' not in body, name
            continue
        if name == "Log_Domain_Status":
            # The mirror of Sample_Include's exception, and the other half of
            # the same distinction. Sample_Include tests ONLY the filtering
            # token because plain Log not filtering is the whole difference
            # between them; this body tests ONLY the strict token because the
            # rows it counts are the ones a strict Log leaves in the sample to
            # poison the fit. Rows the filtering token removed are not an
            # error, so there is nothing here to equality-test them against —
            # the drop token appears as the REMEDY the message names, which is
            # prose, not a comparison. Pairing the counts here would mean
            # inventing a test that must never fire.
            assert f'="{_TRANSFORM_LOG}"' in body.replace(" ", "").replace("\n", "")
            assert f"Use {_TRANSFORM_LOG_DROP}." in body
            assert f'="{_TRANSFORM_LOG_DROP}"' not in body, name
            continue
        # Whitespace around "=" varies between bodies, so normalize it out
        # rather than pinning each body's own formatting.
        compact = body.replace(" ", "").replace("\n", "")
        assert f'="{_TRANSFORM_LOG}"' in compact, name
        # Paired, not merely both present: every equality test against the
        # strict token is matched by one against the other.
        assert compact.count(f'="{_TRANSFORM_LOG}"') == compact.count(
            f'="{_TRANSFORM_LOG_DROP.replace(" ", "")}"'
        ), name

"""RecordingSheet tests for the Model Construction sheet writer.

Excel-side behavior (spill evaluation, Dummy_Levels calls, conditional
formatting rendering) is exercised by the QC build; these tests pin
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
from lambda_catalog.workbook_helpers import excel_color
from lambda_catalog.write_sheet_model_construction import (
    _AUDIT_PAIRS,
    _AUDIT_ROW,
    _C_BREAK_LEFT,
    _C_DESIGN_COLUMNS,
    _C_BREAK_MID,
    _C_FEEDBACK_COUNT,
    _C_FEEDBACK_DELTA,
    _C_FILTERED_LABELS,
    _C_FILTERED_Y,
    _C_MATRIX_LABELS,
    _C_MATRIX_START,
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _FIXED_EFFECTS_COUNT_FORMULA,
    _C_INCLUDE,
    _C_ROLE,
    _C_LABEL,
    _C_LEVELS,
    _C_ORDER,
    _C_REFERENCE,
    _C_REF_IN_USE,
    _C_SPEC_LAST,
    _C_SEQUENCE,
    _C_SEQUENCE_PERIOD,
    _C_PERIOD_IN_USE,
    _C_TRANSFORM,
    _C_TYPE,
    _CLOSURE_SCOPE,
    _FALLBACK_SPEC,
    _FIRST_DATA_ROW,
    _INTERACTION_HEADER_SYMBOLS,
    _INTERACTION_HEADER_UNKNOWN,
    _HEADER_ROW,
    _INTERCEPT_ROW,
    _VALIDATION_LAST_ROW,
    _N_VARIABLES,
    _VARIABLES,
    _MSG_CALENDAR,
    _MSG_NO_NATURAL,
    _MSG_OFF_GRID,
    _MSG_REGULARITY,
    _set_sheet_scoped_names,
    _set_spec_block_column_widths,
    _write_audit_row,
    _write_filtered_zones,
    _write_intercept_control,
    _write_row_zones,
    _write_spec_block,
    SHEET_NAME,
    SPEC_DATASET_PROFILES,
)
from tests.recording_sheet import RecordingSheet

ROOT_DIR = Path(__file__).resolve().parents[1]

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
    "Design_Columns",
    "Serial_Correlation_Group",
    "Sequence_Deltas",
    "Base_Period_Delta_Candidate",
    "Sequence_Delta_Spectrum",
]


def _as_xw_sheet(sheet: RecordingSheet) -> xw.Sheet:
    return cast(xw.Sheet, sheet)


def _model_construction_closures():
    """The sheet-scoped constructor functions as a standalone rebuild installs them.

    The closures moved to scope "Regression" with the v3.0 changeover; this
    module keeps installing the same set when its sheet is rebuilt standalone.
    """
    document = load_catalog_document(ROOT_DIR / "lambda_functions.json")
    return document.functions_for_sheet(_CLOSURE_SCOPE)


def _named_sheet() -> RecordingSheet:
    sheet = RecordingSheet(name=SHEET_NAME)
    _set_sheet_scoped_names(_as_xw_sheet(sheet), _model_construction_closures())
    return sheet


def _refers_to(sheet: RecordingSheet, name: str) -> str:
    return sheet.api.Names.by_short_name(name).RefersTo


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
    # with it. These used to be structured references into the SpecTable
    # ListObject, which pinned them to the row count baked in at build time.
    # The band's ceiling is the same _VALIDATION_LAST_ROW the dropdowns and
    # the conditional formatting already use — one ceiling, so they cannot
    # disagree about how far the block can grow.
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
    mask = _refers_to(sheet, "Sample_Include")

    assert mask.startswith("=LAMBDA(LET(")
    # Filter columns: truthy — TRUE and 1 pass, FALSE/0/blank/text fail.
    # Coercion is (col+0), NOT N(col): col is a bare range reference, and N()
    # of a bare reference implicit-intersects it to a scalar, silently voiding
    # the Filter (the 2482-vs-1649 mask bug). Arithmetic broadcasts instead.
    assert 'IF(INDEX(rl,j)="Filter",acc*--(IFERROR((col+0)=1,FALSE))' in mask
    assert "N(IFERROR(N(col)" not in mask
    # Completeness: the Response and every included Continuous Predictor.
    assert (
        'IF(OR(INDEX(rl,j)="Response (y)",'
        'AND(INDEX(rl,j)="Predictor (x)",INDEX(inc,j)=TRUE,'
        'INDEX(typ,j)="Continuous")),acc*N(ISNUMBER(col)),acc)'
    ) in mask
    # Full-height ones seed; product over {0,1} is the AND, no per-row loop.
    assert "seed,SEQUENCE(ROWS(Source_Data),1,1,0)" in mask
    assert "BYROW(" not in mask
    assert mask.endswith("prod=1))")
    # Reads the model axes only — never the reserved columns or the
    # Sequence structural axis (which no constructor may consume).
    for non_model_axis in (
        "Spec_Order",
        "Spec_Transform",
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
    # G (Transform) went live at v2.2 — it gets a real width now, matching
    # _C_TYPE's (both hold comparably-sized dropdown tokens).
    assert sheet.range((1, _C_TRANSFORM), (1, _C_TRANSFORM)).column_width == 11
    assert sheet.range((1, _C_REF_IN_USE), (1, _C_REF_IN_USE)).column_width == 16


def test_x_s_binds_dummy_levels_once_and_skips_on_isna() -> None:
    sheet = _named_sheet()
    x_s = _refers_to(sheet, "Predictor_Columns")

    assert x_s.startswith("=LAMBDA(LET(")
    # Still bound exactly once, even though v3.1 evaluates blk() for two
    # rows per iteration (the declaring row and its interaction operand) —
    # one textual site is what keeps the two encodings identical.
    assert x_s.count("Dummy_Levels(") == 1
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in x_s
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
    assert 'IF(INDEX(trn,x)="Log",Ln_Positive(col,Sample_Include()),col)' in x_s


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
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in names
    assert names.count("Dummy_Levels(") == 1
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in transforms
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
    assert 'IF(INDEX(trn,x)="Log","Ln("&INDEX(hdrs,1,x)&")",INDEX(hdrs,1,x))' in names
    assert names.endswith("DROP(built,,1)))")
    # Constructed_Column_Transforms: "Log"/"None" per Continuous column;
    # every dummy column from a Categorical Predictor reads "None"
    # unconditionally, regardless of its spec row's own Transform cell.
    assert 'IF(INDEX(trn,x)="Log","Log","None")' in transforms
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
    # Confirm-by-construction property, preserved rather than merely
    # relaxed when G went live: Spec_Transform is read by exactly the four
    # constructors the Log wiring touches, and by nothing else — in
    # particular NOT by Sample_Include or Row_Labels, which never
    # transform anything.
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
        "Predictor_Columns",
        "Response_Column",
    ]
    for non_reader in ("Sample_Include", "Row_Labels"):
        assert "Spec_Transform" not in _refers_to(sheet, non_reader), non_reader


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
    assert readers == ["Base_Period_Delta", "Sequence_Column", "Sequence_Deltas"]
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
    # candidate; it used to reach the same cell through a
    # [@[Sequence Period]] structured reference, so the band name did not
    # appear in any formula. With the block table-free it goes through the
    # band, which is why it shows up here. It is a DISPLAY, not a
    # constructor — a second reader, or one that does not consult the
    # candidate, is the regression this test exists to catch.
    assert len(formula_readers) == 1, formula_readers
    assert "Base_Period_Delta_Candidate()" in formula_readers[0]


def test_reserved_spec_order_is_not_referenced_repo_wide() -> None:
    # Spec_Transform is deliberately excluded here now that it's live —
    # it legitimately appears in lambda_functions.json (the four
    # transform-aware constructors) and in write_sheet_regression.py's
    # note-swap import.
    own_module = "write_sheet_model_construction.py"
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

    assert _N_VARIABLES == 12
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
    # Response, Predictor, and Omit. Full_Data ships as Omit (not Filter): its
    # all-features completeness flag is redundant with the built-in mask and
    # over-filters, so it is not the default Filter.
    by_variable = {v: _FIRST_DATA_ROW + i for i, v in enumerate(_VARIABLES)}
    assert sheet.cell(by_variable["Car Name"], 2).value == "Identifier (Row Label)"
    assert sheet.cell(by_variable["MPG"], 2).value == "Response (y)"
    assert sheet.cell(by_variable["Full_Data"], 2).value == "Omit"
    assert sheet.cell(by_variable["Make"], 2).value == "Omit"
    assert sheet.cell(by_variable["Model?"], 2).value == "Omit"
    for categorical in ("Model Year", "Origin"):
        row = by_variable[categorical]
        assert sheet.cell(row, 2).value == "Predictor (x)"
        assert sheet.cell(row, 3).value is True
        assert sheet.cell(row, 4).value == "Categorical"
    # No column is Sequence-flagged. Auto MPG is cross-sectional, so the
    # shipped profile declares no ordering axis — column H is blank on every
    # row, including Model Year, which used to carry it.
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
        # No structured reference may survive — they only resolve inside a
        # ListObject, which the block no longer creates.
        assert "[@" not in formula, col
        # Nor may the old row-arithmetic: the spill's position must not
        # determine which source column a row maps to.
        assert "ROW()-" not in formula, col


def test_spec_block_defaults_to_the_given_profile() -> None:
    """--regression-dataset life_expectancy must pre-fill its own defaults.

    Life Expectancy has 23 columns vs. Auto MPG's 12. The profile no longer
    sizes anything — the block's height follows COLUMNS(Source_Data) — but
    it still decides which rows arrive with shipped defaults.
    """
    profile = SPEC_DATASET_PROFILES["life_expectancy"]
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet), profile)

    by_variable = {v: _FIRST_DATA_ROW + i for i, v in enumerate(profile.variables)}
    assert sheet.cell(by_variable["Life expectancy"], _C_ROLE).value == "Response (y)"
    assert sheet.cell(by_variable["Country"], _C_ROLE).value == "Identifier (Row Label)"
    assert sheet.cell(by_variable["Full_Data"], _C_ROLE).value == "Omit"
    schooling_row = by_variable["Schooling"]
    assert sheet.cell(schooling_row, _C_ROLE).value == "Predictor (x)"
    assert sheet.cell(schooling_row, _C_INCLUDE).value is True
    assert sheet.cell(schooling_row, _C_TYPE).value == "Continuous"
    assert sheet.cell(by_variable["Year"], _C_SEQUENCE).value is True


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
    sheet = _named_sheet()
    _write_all_zones(sheet)

    band = "Spec_Design_Columns"
    readers = [
        item.Name.split("!", 1)[-1]
        for item in sheet.api.Names.items
        if band in item.RefersTo and item.Name.split("!", 1)[-1] != band
    ]
    assert readers == []
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
        ((r, 7), (16000, 7)),  # G Transform (None or Log, live at v2.2)
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
    assert formulas[7] == "None,Log"
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
    # bands. Every range runs out to _VALIDATION_LAST_ROW so a row added by
    # typing past SpecTable's current bottom edge (auto-extending the
    # ListObject) is already covered.
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
        f'=AND($B{r}="Predictor (x)",$D{r}="Categorical",$G{r}="Log")',
    ]
    assert transform_col[0].Font.Color == excel_color(INPUT_COLOR)
    assert transform_col[1].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert transform_col[1].Font.Color == excel_color(CF_DARK_RED_TEXT)

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


def test_sequence_status_line_validates_zero_or_one_flags() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))
    # The status cell lives in _write_spec_feedback (E1) once the spec
    # data area becomes a structured table (SpecTable) — H2 is now the
    # table's "Sequence" header cell, and a status cell on top of a
    # table header reads as a visual collision.
    from lambda_catalog.write_sheet_model_construction import _write_spec_feedback
    _write_spec_feedback(_as_xw_sheet(sheet))

    # E1: blank while the spec carries zero-or-one flags, a red error line
    # at two-plus — the exactly-one-Response pattern with a >1 threshold
    # (zero flags is a valid non-panel spec).
    status = sheet.cell(1, _C_REFERENCE)
    assert status.api.Formula2 == (
        "=IF(SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))>1,"
        '"ERROR: multiple Sequence flags (mark at most one variable)","")'
    )
    assert status.api.Font.Bold is True

    conditions = sheet.range("$E$1").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ['=$E$1<>""']
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_fixed_effects_status_line_validates_zero_or_one_rows() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))
    from lambda_catalog.write_sheet_model_construction import _write_spec_feedback
    _write_spec_feedback(_as_xw_sheet(sheet))

    # B1: the Fixed Effects cardinality error — same pattern as E1's
    # Sequence check, on Role's own row-1 cell instead of Reference Level's.
    status = sheet.cell(1, _C_ROLE)
    assert status.api.Formula2 == (
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}>1,'
        '"ERROR: multiple Fixed Effects rows (mark at most one variable)","")'
    )
    assert status.api.Font.Bold is True

    conditions = sheet.range("$B$1").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ['=$B$1<>""']
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_fixed_effects_status_block_shows_variable_groups_and_absorbed_df() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    from lambda_catalog.write_sheet_model_construction import _write_spec_feedback
    _write_spec_feedback(_as_xw_sheet(sheet))

    for col, label in (
        (_C_PERIOD_IN_USE, "FE Variable"),
        (_C_LEVELS, "FE Groups"),
        (_C_REF_IN_USE, "FE df absorbed"),
    ):
        cell = sheet.cell(1, col)
        assert cell.value == label, (col, label)
        assert cell.api.Font.Bold is True

    variable = cast(str, sheet.cell(2, _C_PERIOD_IN_USE).api.Formula2)
    assert variable.startswith(f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"n/a",')
    assert 'XMATCH("Fixed Effects",TAKE(Spec_Role,COLUMNS(Source_Data)))' in variable

    groups = sheet.cell(2, _C_LEVELS).api.Formula2
    assert groups == (
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"n/a",Absorbed_Degrees_Of_Freedom()+1)'
    )

    absorbed = sheet.cell(2, _C_REF_IN_USE).api.Formula2
    assert absorbed == (
        f'=IF({_FIXED_EFFECTS_COUNT_FORMULA}=0,"n/a",Absorbed_Degrees_Of_Freedom())'
    )


def test_spec_feedback_writes_delta_count_verdict_with_priority_cf() -> None:
    """The M/N spectrum and the I1/I2 verdict overlay (Verdict overlays the
    Sequence_Period column's row-1/row-2 cells, which are unused by the spec
    table, SpecTable).

    M1/N1: bold headers (Δ, Count). M2: the Sequence_Delta_Spectrum() spill,
    wrapped in IFERROR so a no-axis / no-spacings case degrades to blank.
    I1/I2: the combined switch — one cell, one message, with red CF outranking
    yellow via StopIfTrue.
    """
    sheet = RecordingSheet(name=SHEET_NAME)
    from lambda_catalog.write_sheet_model_construction import _write_spec_feedback
    _write_spec_feedback(_as_xw_sheet(sheet))

    # M1/N1 headers, bold via bold_row (range-level bold, not per-cell).
    for col, label in (
        (_C_FEEDBACK_DELTA, "Δ"),
        (_C_FEEDBACK_COUNT, "Count"),
        (_C_SEQUENCE_PERIOD, "Verdict"),
    ):
        cell = sheet.cell(1, col)
        assert cell.value == label, (col, label)
    # Bold-row applied across M:N — the Verdict header (I1) is bolded
    # independently (lives in column I, outside the M:N range).
    assert (
        sheet.range((1, _C_FEEDBACK_DELTA), (1, _C_FEEDBACK_COUNT)).api.Font.Bold
        is True
    )
    assert sheet.cell(1, _C_SEQUENCE_PERIOD).api.Font.Bold is True

    # M2 spectrum spill.
    assert sheet.cell(2, _C_FEEDBACK_DELTA).api.Formula2 == (
        '=IFERROR(Sequence_Delta_Spectrum(),"")'
    )

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

    # A2 label (bold), C2 toggle prefilled TRUE with input styling.
    assert sheet.cell(_INTERCEPT_ROW, _C_LABEL).value == "Intercept"
    assert sheet.cell(_INTERCEPT_ROW, _C_LABEL).api.Font.Bold is True
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
    'IFERROR(IF(INDEX(TAKE(Spec_Transform,n_c),p)="Log","Ln("&h&")",h),"(none)"))'
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

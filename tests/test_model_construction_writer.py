"""RecordingSheet tests for the Model Construction sheet writer.

Excel-side behavior (spill evaluation, Dummy_Levels calls, conditional
formatting rendering) is exercised by the QC build; these tests pin
everything checkable without Excel — the sheet-scoped name definitions and
their order, the T0 default-spec prefill, the dropdown and conditional-
formatting registrations, and the structural invariants of the X_s /
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
    INPUT_COLOR,
)
from lambda_catalog.workbook_helpers import excel_color
from lambda_catalog.write_sheet_model_construction import (
    _AUDIT_PAIRS,
    _AUDIT_ROW,
    _C_BREAK_LEFT,
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
    _C_SEQUENCE,
    _C_SEQUENCE_PERIOD,
    _C_PERIOD_IN_USE,
    _C_TRANSFORM,
    _CLOSURE_SCOPE,
    _FALLBACK_SPEC,
    _FIRST_DATA_ROW,
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
    "Allow_Intercept",
    "Sample_Include",
    "Response_Column",
    "Row_Labels",
    "X_s",
    "Constructed_Column_Names",
    "Sequence_Column",
    "Fixed_Effects_Column",
    "Absorbed_Degrees_Of_Freedom",
    "Prediction_Group_Column",
    "y_s",
    "X_s_Within",
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

    # Spec_* band names bind to the SpecTable via structured references:
    # SpecTable[[#Data],[Column]]. The [#Data] qualifier restricts the
    # range to the data body (the spec rows), which is what every
    # TAKE-trimmed consumer expects: the spec rows, not the headers.
    # The column header text (with spaces) is the actual Excel header —
    # structured references are case- and whitespace-sensitive, so the
    # "Reference Level" / "Sequence Period" forms are correct.
    for name, header in (
        ("Spec_Role", "Role"),
        ("Spec_Include", "Include"),
        ("Spec_Type", "Type"),
        ("Spec_Reference", "Reference Level"),
        ("Spec_Order", "Order"),
        ("Spec_Transform", "Transform"),
        ("Spec_Sequence", "Sequence"),
        ("Spec_Sequence_Period", "Sequence Period"),
        ("Spec_Period_In_Use", "Period In Use"),
    ):
        assert _refers_to(sheet, name) == (
            f"='{SHEET_NAME}'!SpecTable[[#Data],[{header}]]"
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

    # O: narrow gap, splitting the spec block and the derived-row zone.
    assert sheet.range((1, 15)).column_width == 2

    # P/Q headers on the spec-header row, bold like the A–L headers.
    assert sheet.cell(_HEADER_ROW, 16).value == "Row Labels"
    assert sheet.cell(_HEADER_ROW, 17).value == "Included"
    assert sheet.range((_HEADER_ROW, 16), (_HEADER_ROW, 17)).api.Font.Bold is True

    # Full-height spills at the first data row; row 1 belongs to
    # _write_audit_row, so this writer must leave it untouched.
    assert sheet.cell(_FIRST_DATA_ROW, 16).api.Formula2 == "=Row_Labels()"
    assert sheet.cell(_FIRST_DATA_ROW, 17).api.Formula2 == "=Sample_Include()"
    for col in (16, 17):
        assert sheet.cell(1, col).value is None
        assert sheet.cell(1, col).api.Formula2 is None


def test_spec_block_column_widths_hide_reserved_columns() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _set_spec_block_column_widths(_as_xw_sheet(sheet))

    assert sheet.range((1, _C_LABEL), (1, _C_LABEL)).column_width == 28
    assert sheet.range((1, _C_ORDER), (1, _C_ORDER)).column_width == 0
    assert sheet.range((1, _C_TRANSFORM), (1, _C_TRANSFORM)).column_width == 0
    assert sheet.range((1, _C_REF_IN_USE), (1, _C_REF_IN_USE)).column_width == 16


def test_x_s_binds_dummy_levels_once_and_skips_on_isna() -> None:
    sheet = _named_sheet()
    x_s = _refers_to(sheet, "X_s")

    assert x_s.startswith("=LAMBDA(LET(")
    assert x_s.count("Dummy_Levels(") == 1
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in x_s
    # Scalar skip guard: ISNA(INDEX(lv,1,1)), NOT ISNA(lv). lv is a 1x(L-1)
    # row; an array condition in front of a wider HSTACK branch broadcasts to
    # #N/A (the T6 header-strip bug). INDEX(lv,1,1) makes the test scalar.
    assert "IF(ISNA(INDEX(lv,1,1)),acc,HSTACK(acc,--(col=lv)))" in x_s
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


def test_constructed_column_names_is_a_structural_twin_of_x_s() -> None:
    sheet = _named_sheet()
    x_s = _refers_to(sheet, "X_s")
    names = _refers_to(sheet, "Constructed_Column_Names")

    # Identical iteration predicate and skip conditions — twinning is what
    # guarantees the header strip width always matches COLUMNS(X_s()).
    predicate = 'IF(OR(INDEX(rl,j)<>"Predictor (x)",INDEX(inc,j)<>TRUE),acc,'
    assert predicate in x_s
    assert predicate in names
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in names
    assert names.count("Dummy_Levels(") == 1
    # Same scalar skip guard as X_s (the twin must match).
    assert "IF(ISNA(INDEX(lv,1,1)),acc," in names
    assert "IF(ISNA(lv)," not in names
    # Level-qualified headers: "Status: Developing", "Year: 2001", ...
    assert 'HSTACK(acc,h&": "&lv)' in names
    assert names.endswith("DROP(built,,1)))")


def test_reserved_spec_names_are_defined_but_read_by_nothing() -> None:
    # F (Order) and G (Transform) remain reserved. I (Base Period Δ) went
    # live with the base-period release and has its own reader tests below.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    for reserved in ("Spec_Order", "Spec_Transform"):
        readers = [
            item.Name
            for item in sheet.api.Names.items
            if reserved in item.RefersTo
            and item.Name.split("!", 1)[-1] != reserved
        ]
        assert readers == [], reserved
        for formula in _all_written_formulas(sheet):
            assert reserved not in formula, (reserved, formula)


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
    assert readers == ["Sequence_Column", "Sequence_Deltas"]
    for constructor in (
        "Sample_Include",
        "Response_Column",
        "Row_Labels",
        "X_s",
        "Constructed_Column_Names",
    ):
        assert "Spec_Sequence" not in _refers_to(sheet, constructor), constructor


def test_sequence_period_name_is_read_only_by_the_base_period_layer() -> None:
    # Spec_Sequence_Period is live, but only for the Δ-in-use display in
    # the Sequence Spacing block (on-sheet) and the workbook-scoped
    # Base_Period_Delta() accessor (in lambda_functions.json). No sheet
    # closure — constructor or otherwise — reads it: the candidate is
    # computed from the data, never from its own display cell.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    name_readers = [
        item.Name
        for item in sheet.api.Names.items
        if "Spec_Sequence_Period" in item.RefersTo
        and item.Name.split("!", 1)[-1] != "Spec_Sequence_Period"
    ]
    assert name_readers == []

    formula_readers = [
        formula
        for formula in _all_written_formulas(sheet)
        if "Spec_Sequence_Period" in formula
    ]
    assert formula_readers == []


def test_reserved_spec_names_are_not_referenced_repo_wide() -> None:
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
        assert "Spec_Transform" not in text, path.name


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
        # I (Sequence Period) is the typed override input — input-styled,
        # blank by default; the user types a number here to override the
        # candidate. The candidate formula lives in J.
        assert sheet.cell(row, 9).value is None, variable
        # J (Period In Use) is the candidate-with-override display:
        # the typed override if I is non-blank, otherwise the candidate
        # closure's value. The scalar [@Sequence] test keeps the
        # candidate lazy (it computes only on the flagged row). The
        # formula is written through .Formula (not .Formula2) because
        # structured references are rejected by Formula2.
        assert sheet.cell(row, 10).api.Formula == (
            '=IF([@Sequence]<>TRUE,"",'
            'IF(N([@[Sequence Period]])<>0,[@[Sequence Period]],'
            'IFERROR(Base_Period_Delta_Candidate(),"")))'
        ), variable
        for col in range(2, 10):
            assert sheet.cell(row, col).color == INPUT_COLOR, (variable, col)

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
    # Model Year is additionally flagged as the Sequence (ordering) axis.
    assert sheet.cell(by_variable["Model Year"], _C_SEQUENCE).value is True


def test_levels_column_counts_raw_levels_without_dummy_levels() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    formula = cast(str, sheet.cell(_FIRST_DATA_ROW, _C_LEVELS).api.Formula)
    # Must display L itself (1 for degenerate columns, which Dummy_Levels
    # signals as #N/A instead), so it counts UNIQUE directly.
    assert formula.startswith(
        '=IF(OR([@Role]<>"Predictor (x)",[@Type]<>"Categorical"),"",'
    )
    assert "ROWS(UNIQUE(FILTER(" in formula
    assert "Dummy_Levels" not in formula
    assert "Sample_Include()" in formula
    # First data row → Source_Data column 1.
    assert f"ROW()-{_FIRST_DATA_ROW - 1}" in formula
    assert 'x,IF(col="","",col)' in formula  # blank normalization mirrored


def test_reference_in_use_echoes_e_or_shows_the_sorted_default() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    r = _FIRST_DATA_ROW
    formula = cast(str, sheet.cell(r, _C_REF_IN_USE).api.Formula)
    # Same relevance guard as the Levels display: Categorical Predictors only.
    assert formula.startswith(
        '=IF(OR([@Role]<>"Predictor (x)",[@Type]<>"Categorical"),"",'
    )
    # An explicit Reference Level is echoed verbatim (its invalid-reference
    # CF carries the error signal); blank Reference Level falls through to
    # the default.
    assert 'IF([@[Reference Level]]<>"",[@[Reference Level]],' in formula
    # The default mirrors Dummy_Levels: first sorted level over the
    # mask-included sample, with the same blank normalization. NOT a
    # Dummy_Levels call — that returns the retained levels, i.e. everything
    # EXCEPT the reference.
    assert "INDEX(SORT(UNIQUE(FILTER(" in formula
    assert "Dummy_Levels" not in formula
    assert "Sample_Include()" in formula
    assert f"ROW()-{_FIRST_DATA_ROW - 1}" in formula
    assert 'x,IF(col="","",col)' in formula
    # Empty masked sample degrades to blank (H shows 0 and flags red there).
    assert formula.endswith(',""))))')


def test_dropdowns_cover_exactly_the_four_list_columns() -> None:
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
        ((r, 7), (16000, 7)),  # G Transform (reserved; only "None" valid)
        ((r, 8), (16000, 8)),  # H Sequence (TRUE or blank)
    }
    formulas = {
        key[0][1]: validation.rules[0]["Formula1"]
        for key, validation in validated.items()
    }
    assert formulas[2] == "Response (y),Predictor (x),Identifier (Row Label),Filter,Omit,Fixed Effects"
    assert formulas[3] == "TRUE,FALSE"
    assert formulas[4] == "Continuous,Categorical"
    assert formulas[7] == "None"
    assert formulas[8] == "TRUE"
    for validation in validated.values():
        assert validation.delete_count == 1
        assert validation.rules[0]["Type"] == 3  # xlValidateList
        assert validation.IgnoreBlank is True


def test_conditional_formats_cover_cascading_relevance_degeneracy_and_reference() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    r = _FIRST_DATA_ROW
    off = _FIRST_DATA_ROW - 1
    # Role-keyed relevance: the per-Predictor inputs (C–G) hide behind their
    # own INPUT_COLOR fill, the Categorical displays (K–L) hide behind
    # white (unfilled computed cells); H–J are deliberately NOT in these
    # bands. Every range runs out to _VALIDATION_LAST_ROW so a row added by
    # typing past SpecTable's current bottom edge (auto-extending the
    # ListObject) is already covered.
    role_keyed_input = sheet.range(
        f"$C${r}:$G${_VALIDATION_LAST_ROW}"
    ).api.FormatConditions.items
    assert [c.Formula1 for c in role_keyed_input] == [f'=$B{r}<>"Predictor (x)"']
    assert role_keyed_input[0].Font.Color == excel_color(INPUT_COLOR)

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
    conditions = sheet.range(f"$I$2").api.FormatConditions.items
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
    'IFERROR(INDEX(TOROW(Header_Names),'
    'XMATCH("Response (y)",TAKE(Spec_Role,COLUMNS(Source_Data)))),"(none)")'
)


def test_audit_row_is_bold_label_value_pairs_with_response_count_cf() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_audit_row(_as_xw_sheet(sheet))

    expected = [
        (16, 17, "k", '=IFERROR(COLUMNS(X_s()),"(empty model)")'),
        (19, 20, "rows", '=IFERROR(ROWS(X_s()),"(empty model)")'),
        (22, 23, "response", f"={_RESPONSE_NAME}"),
        (
            24,
            25,
            "responses",
            '=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))',
        ),
        (26, 27, "included rows", "=SUMPRODUCT(N(Sample_Include()))"),
        (
            28,
            29,
            "sequence flags",
            "=SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))",
        ),
        (
            30,
            31,
            "fixed effects",
            '=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))',
        ),
        (32, 33, "FE absorbed df", "=Absorbed_Degrees_Of_Freedom()"),
    ]
    assert list(_AUDIT_PAIRS) == [(lc, vc) for lc, vc, _, _ in expected]
    assert _AUDIT_ROW == 1
    for label_col, value_col, label, formula in expected:
        # No audit cell may land on a width-2 break column (R=18, U=21).
        assert label_col not in (_C_BREAK_LEFT, _C_BREAK_MID)
        assert value_col not in (_C_BREAK_LEFT, _C_BREAK_MID)
        assert sheet.cell(1, label_col).value == label
        assert sheet.cell(1, value_col).api.Formula2 == formula
        assert (
            sheet.range((1, label_col), (1, value_col)).api.Font.Bold is True
        )

    # Exactly-one-Response validation: red CF on the responses count cell (Y=25).
    conditions = sheet.range("$Y$1").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ["=N($Y$1)<>1"]
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Zero-or-one-Sequence validation: red CF only at two-plus flags (AC=29).
    seq_conditions = sheet.range("$AC$1").api.FormatConditions.items
    assert [c.Formula1 for c in seq_conditions] == ["=N($AC$1)>1"]
    assert seq_conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert seq_conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Zero-or-one-Fixed-Effects validation: red CF only at two-plus (AE=31).
    fe_conditions = sheet.range("$AE$1").api.FormatConditions.items
    assert [c.Formula1 for c in fe_conditions] == ["=N($AE$1)>1"]
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
        _C_MATRIX_START: "X_s()",
    }
    for col, source in expected_spills.items():
        assert sheet.cell(_FIRST_DATA_ROW, col).api.Formula2 == (
            f'=IFERROR(FILTER({source},Sample_Include()),"(empty model)")'
        ), source

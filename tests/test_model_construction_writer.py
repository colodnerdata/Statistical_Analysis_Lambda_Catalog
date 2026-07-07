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
    MUTED_TEXT_COLOR,
)
from lambda_catalog.workbook_helpers import excel_color
from lambda_catalog.write_sheet_model_construction import (
    _AUDIT_PAIRS,
    _AUDIT_ROW,
    _C_BREAK_LEFT,
    _C_BREAK_MID,
    _C_FILTERED_LABELS,
    _C_FILTERED_Y,
    _C_MATRIX_LABELS,
    _C_MATRIX_START,
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _C_BASE_PERIOD,
    _C_INCLUDE,
    _C_LABEL,
    _C_LEVELS,
    _C_REF_IN_USE,
    _C_SEQUENCE,
    _CLOSURE_SCOPE,
    _FALLBACK_SPEC,
    _FIRST_DATA_ROW,
    _HEADER_ROW,
    _INTERCEPT_ROW,
    _LAST_DATA_ROW,
    _N_VARIABLES,
    _VARIABLES,
    _MSG_CALENDAR,
    _MSG_NO_NATURAL,
    _MSG_OFF_GRID,
    _MSG_REGULARITY,
    _ROW_SPACING_CALENDAR,
    _ROW_SPACING_CANDIDATE,
    _ROW_SPACING_HEADING,
    _ROW_SPACING_IN_USE,
    _ROW_SPACING_NO_NATURAL,
    _ROW_SPACING_OFF_GRID,
    _ROW_SPACING_REGULARITY,
    _set_sheet_scoped_names,
    _write_audit_row,
    _write_filtered_zones,
    _write_intercept_control,
    _write_row_zones,
    _write_sequence_spacing_block,
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
    "Spec_Base_Period_Delta",
    "Allow_Intercept",
    "Sample_Include",
    "Response_Column",
    "Row_Labels",
    "X_s",
    "Constructed_Column_Names",
    "Sequence_Column",
    "Fixed_Effects_Column",
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
    _write_sequence_spacing_block(_as_xw_sheet(sheet))
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
    assert _refers_to(sheet, "Source_Table") == "=LifeExpectancyData[#All]"
    assert _refers_to(sheet, "Source_Data") == "=DROP(Source_Table,1)"
    assert _refers_to(sheet, "Header_Names") == "=TAKE(Source_Table,1)"
    for name in _EXPECTED_NAME_ORDER[1:]:
        assert "LifeExpectancyData" not in _refers_to(sheet, name), name

    _write_all_zones(sheet)
    for formula in _all_written_formulas(sheet):
        assert "LifeExpectancyData" not in formula, formula


def test_spec_ranges_cover_the_standard_input_band() -> None:
    sheet = _named_sheet()

    for name, column in (
        ("Spec_Role", "B"),
        ("Spec_Include", "C"),
        ("Spec_Type", "D"),
        ("Spec_Reference", "E"),
        ("Spec_Order", "F"),
        ("Spec_Transform", "G"),
        ("Spec_Sequence", "H"),
        ("Spec_Base_Period_Delta", "I"),
    ):
        assert _refers_to(sheet, name) == (
            f"='{SHEET_NAME}'!${column}${_FIRST_DATA_ROW}:${column}$16000"
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
        "Spec_Base_Period_Delta",
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

    # L: narrow gap, visually reserving the future Design Columns column.
    assert sheet.range((1, 12)).column_width == 2

    # M/N headers on the spec-header row, bold like the A–K headers.
    assert sheet.cell(_HEADER_ROW, 13).value == "Row Labels"
    assert sheet.cell(_HEADER_ROW, 14).value == "Included"
    assert sheet.range((_HEADER_ROW, 13), (_HEADER_ROW, 14)).api.Font.Bold is True

    # Full-height spills at the first data row; row 1 belongs to
    # _write_audit_row, so this writer must leave it untouched.
    assert sheet.cell(_FIRST_DATA_ROW, 13).api.Formula2 == "=Row_Labels()"
    assert sheet.cell(_FIRST_DATA_ROW, 14).api.Formula2 == "=Sample_Include()"
    for col in (13, 14):
        assert sheet.cell(1, col).value is None
        assert sheet.cell(1, col).api.Formula2 is None


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


def test_base_period_name_is_read_only_by_the_base_period_layer() -> None:
    # Spec_Base_Period_Delta is live, but only for the Δ-in-use display in
    # the Sequence Spacing block (on-sheet) and the workbook-scoped
    # Base_Period_Delta() accessor (in lambda_functions.json). No sheet
    # closure — constructor or otherwise — reads it: the candidate is
    # computed from the data, never from its own display cell.
    sheet = _named_sheet()
    _write_all_zones(sheet)

    name_readers = [
        item.Name
        for item in sheet.api.Names.items
        if "Spec_Base_Period_Delta" in item.RefersTo
        and item.Name.split("!", 1)[-1] != "Spec_Base_Period_Delta"
    ]
    assert name_readers == []

    formula_readers = [
        formula
        for formula in _all_written_formulas(sheet)
        if "Spec_Base_Period_Delta" in formula
    ]
    in_use = cast(
        str, sheet.cell(_ROW_SPACING_IN_USE, 2).api.Formula2
    )
    assert formula_readers == [in_use]
    # Every read of the live H/I bands must TAKE-trim to the table width —
    # the Sequence Spacing block itself sits inside the 16000-row bands.
    assert "TAKE(Spec_Base_Period_Delta,n_c)" in in_use
    assert "TAKE(Spec_Sequence,n_c)" in in_use


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

    assert _N_VARIABLES == 23
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
        # I (Base Period Δ) is computed-with-override: pre-filled with the
        # candidate formula (blank off the flagged row via the scalar $H
        # test, so the candidate stays lazy), input-styled so a typed
        # number overrides it.
        assert sheet.cell(row, 9).api.Formula2 == (
            f'=IF($H{row}<>TRUE,"",'
            'IFERROR(Base_Period_Delta_Candidate(),""))'
        ), variable
        for col in range(2, 10):
            assert sheet.cell(row, col).color == INPUT_COLOR, (variable, col)

    # Spot-check the named T0 roles — the shipped spec demonstrates Identifier,
    # Response, Predictor, and Omit. Full_Data ships as Omit (not Filter): its
    # all-features completeness flag is redundant with the built-in mask and
    # over-filters, so it is not the default Filter.
    by_variable = {v: _FIRST_DATA_ROW + i for i, v in enumerate(_VARIABLES)}
    assert sheet.cell(by_variable["Country"], 2).value == "Identifier (Row Label)"
    assert sheet.cell(by_variable["Life expectancy"], 2).value == "Response (y)"
    assert sheet.cell(by_variable["Full_Data"], 2).value == "Omit"
    assert sheet.cell(by_variable["Population"], 2).value == "Omit"
    for categorical in ("Year", "Status"):
        row = by_variable[categorical]
        assert sheet.cell(row, 2).value == "Predictor (x)"
        assert sheet.cell(row, 3).value is True
        assert sheet.cell(row, 4).value == "Categorical"
    # Year is additionally flagged as the Sequence (ordering) axis.
    assert sheet.cell(by_variable["Year"], _C_SEQUENCE).value is True


def test_levels_column_counts_raw_levels_without_dummy_levels() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    formula = cast(str, sheet.cell(_FIRST_DATA_ROW, _C_LEVELS).api.Formula2)
    # Must display L itself (1 for degenerate columns, which Dummy_Levels
    # signals as #N/A instead), so it counts UNIQUE directly.
    assert formula.startswith(
        f'=IF(OR($B{_FIRST_DATA_ROW}<>"Predictor (x)",$D{_FIRST_DATA_ROW}<>"Categorical"),"",'
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
    formula = cast(str, sheet.cell(r, _C_REF_IN_USE).api.Formula2)
    # Same relevance guard as the Levels display: Categorical Predictors only.
    assert formula.startswith(
        f'=IF(OR($B{r}<>"Predictor (x)",$D{r}<>"Categorical"),"",'
    )
    # An explicit E is echoed verbatim (E's invalid-reference CF carries the
    # error signal); blank E falls through to the default.
    assert f'IF($E{r}<>"",$E{r},' in formula
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
    assert formulas[2] == "Response (y),Predictor (x),Identifier (Row Label),Filter,Omit"
    assert formulas[3] == "TRUE,FALSE"
    assert formulas[4] == "Continuous,Categorical"
    assert formulas[7] == "None"
    assert formulas[8] == "TRUE"
    for validation in validated.values():
        assert validation.delete_count == 1
        assert validation.rules[0]["Type"] == 3  # xlValidateList
        assert validation.IgnoreBlank is True


def test_conditional_formats_cover_gray_cascade_degeneracy_and_reference() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    r = _FIRST_DATA_ROW
    off = _FIRST_DATA_ROW - 1
    # Role-keyed relevance: the per-Predictor inputs (C–G) and the
    # Categorical displays (J–K); H–I are deliberately NOT in these bands.
    for band in (
        f"$C${r}:$G${_LAST_DATA_ROW}",
        f"$J${r}:$K${_LAST_DATA_ROW}",
    ):
        gray = sheet.range(band).api.FormatConditions.items
        assert [c.Formula1 for c in gray] == [f'=$B{r}<>"Predictor (x)"'], band
        assert gray[0].Font.Color == excel_color(MUTED_TEXT_COLOR)

    # Sequence-keyed relevance: H–I gray on rows that are not the sequence
    # axis — keyed on the flag itself, not on Role.
    seq_gray = sheet.range(f"$H${r}:$I${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in seq_gray] == [f"=$H{r}<>TRUE"]
    assert seq_gray[0].Font.Color == excel_color(MUTED_TEXT_COLOR)

    # Multi-flag error: red on every flagged H cell at two-plus flags.
    multi = sheet.range(f"$H${r}:$H${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in multi] == [
        f"=AND($H{r}=TRUE,"
        "SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))>1)"
    ]
    assert multi[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert multi[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    degenerate = sheet.range(f"$J${r}:$J${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in degenerate] == [
        f'=AND($B{r}="Predictor (x)",$C{r}=TRUE,$D{r}="Categorical",N($J{r})<=1)'
    ]
    assert degenerate[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert degenerate[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Invalid reference: the constructor's exact skip condition, tested
    # directly — a membership test against Dummy_Levels' output would
    # false-positive on the default reference itself.
    invalid = sheet.range(f"$E${r}:$E${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in invalid] == [
        f'=AND($E{r}<>"",ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-{off}),'
        f"$E{r},Sample_Include())))"
    ]
    assert invalid[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert invalid[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_sequence_status_line_validates_zero_or_one_flags() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    # H2: blank while the spec carries zero-or-one flags, a red error line
    # at two-plus — the exactly-one-Response pattern with a >1 threshold
    # (zero flags is a valid non-panel spec).
    status = sheet.cell(_INTERCEPT_ROW, _C_SEQUENCE)
    assert status.api.Formula2 == (
        "=IF(SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))>1,"
        '"ERROR: multiple Sequence flags (mark at most one variable)","")'
    )
    assert status.api.Font.Bold is True

    conditions = sheet.range("$H$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ['=$H$2<>""']
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


def test_sequence_spacing_block_layout_and_verdicts() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_sequence_spacing_block(_as_xw_sheet(sheet))

    # Fixed rows below the spec: heading 28, Δ candidate/in-use 29–30,
    # verdict lines 31–34; spectrum headers on 29 with the spill on 30.
    assert (_ROW_SPACING_HEADING, _ROW_SPACING_CANDIDATE, _ROW_SPACING_IN_USE) == (
        28, 29, 30,
    )
    assert (
        _ROW_SPACING_REGULARITY,
        _ROW_SPACING_OFF_GRID,
        _ROW_SPACING_NO_NATURAL,
        _ROW_SPACING_CALENDAR,
    ) == (31, 32, 33, 34)

    heading = sheet.cell(_ROW_SPACING_HEADING, 1)
    assert heading.value == "Sequence Spacing (Base Period Δ)"
    assert heading.api.Font.Bold is True

    # Δ candidate: the computed MODE-with-MIN-fallback closure, degrading
    # to blank when there are no within-group spacings.
    assert sheet.cell(_ROW_SPACING_CANDIDATE, 1).value == "Δ candidate"
    assert sheet.cell(_ROW_SPACING_CANDIDATE, 2).api.Formula2 == (
        '=IFERROR(Base_Period_Delta_Candidate(),"")'
    )

    # Δ in use: the flagged row's I cell via LOCAL Spec_* names (a
    # standalone Model Construction rebuild must read its own spec), both
    # bands TAKE-trimmed so the block's own rows are never scanned.
    assert sheet.cell(_ROW_SPACING_IN_USE, 1).value == "Δ in use"
    in_use = cast(str, sheet.cell(_ROW_SPACING_IN_USE, 2).api.Formula2)
    assert "TAKE(Spec_Sequence,n_c)" in in_use
    assert "TAKE(Spec_Base_Period_Delta,n_c)" in in_use
    assert '"(not set)"' in in_use

    # Visible override: yellow when the Δ in use differs from the candidate.
    override = sheet.range("$B$30").api.FormatConditions.items
    assert [c.Formula1 for c in override] == [
        "=AND(ISNUMBER($B$30),ISNUMBER($B$29),$B$30<>$B$29)"
    ]
    assert override[0].Interior.Color == excel_color(CF_YELLOW_FILL)
    assert override[0].Font.Color == excel_color(CF_DARK_YELLOW_TEXT)

    # Delta spectrum: two-column spill at J/K under bold headers.
    assert sheet.cell(_ROW_SPACING_CANDIDATE, 10).value == "Delta"
    assert sheet.cell(_ROW_SPACING_CANDIDATE, 11).value == "Count"
    assert sheet.cell(_ROW_SPACING_IN_USE, 10).api.Formula2 == (
        '=IFERROR(Sequence_Delta_Spectrum(),"")'
    )

    # Verdict lines: blank while quiet; each carries its message and its
    # nonblank-keyed CF (yellow = advisory, red = off the declared grid).
    expected_verdicts = (
        (_ROW_SPACING_REGULARITY, _MSG_REGULARITY, CF_YELLOW_FILL, CF_DARK_YELLOW_TEXT),
        (_ROW_SPACING_OFF_GRID, _MSG_OFF_GRID, CF_LIGHT_RED_FILL, CF_DARK_RED_TEXT),
        (_ROW_SPACING_NO_NATURAL, _MSG_NO_NATURAL, CF_YELLOW_FILL, CF_DARK_YELLOW_TEXT),
        (_ROW_SPACING_CALENDAR, _MSG_CALENDAR, CF_LIGHT_RED_FILL, CF_DARK_RED_TEXT),
    )
    for row, message, fill, font_color in expected_verdicts:
        formula = cast(str, sheet.cell(row, 1).api.Formula2)
        assert formula.startswith("=LET(d,Sequence_Deltas(),"), row
        assert f'"{message}"' in formula, row
        conditions = sheet.range(f"$A${row}").api.FormatConditions.items
        assert [c.Formula1 for c in conditions] == [f'=$A${row}<>""'], row
        assert conditions[0].Interior.Color == excel_color(fill), row
        assert conditions[0].Font.Color == excel_color(font_color), row

    # Verdicts key on the Δ-in-use cell, so a typed override re-runs every
    # check against the user's Δ; the calendar check is Δ-independent
    # (surfacing, never quantizing) and the MODE prompt reads the data only.
    regularity = cast(str, sheet.cell(_ROW_SPACING_REGULARITY, 1).api.Formula2)
    off_grid = cast(str, sheet.cell(_ROW_SPACING_OFF_GRID, 1).api.Formula2)
    calendar = cast(str, sheet.cell(_ROW_SPACING_CALENDAR, 1).api.Formula2)
    assert "$B$30" in regularity and "SUM(--(d<>v))>0" in regularity
    assert "$B$30" in off_grid and "MOD(d,v)<>0" in off_grid
    assert "$B$30" not in calendar
    no_natural = cast(str, sheet.cell(_ROW_SPACING_NO_NATURAL, 1).api.Formula2)
    assert "ISNA(MODE.SNGL(d))" in no_natural

    # The spec dropdown validations are scrubbed off the block's rows.
    scrub = sheet.range((28, 2), (34, 9)).api.Validation
    assert scrub.delete_count == 1


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
    # gray "required-here" rule; gray applies whenever a Categorical is in.
    conditions = sheet.range("$C$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == [
        f"=AND($C$2=FALSE,{_CAT_INCLUDED})",
        f"={_CAT_INCLUDED}",
    ]
    red, gray = conditions
    assert red.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert red.Font.Color == excel_color(CF_DARK_RED_TEXT)
    assert red.StopIfTrue is True
    assert gray.Font.Color == excel_color(MUTED_TEXT_COLOR)
    assert gray.StopIfTrue is False


_RESPONSE_NAME = (
    'IFERROR(INDEX(TOROW(Header_Names),'
    'XMATCH("Response (y)",TAKE(Spec_Role,COLUMNS(Source_Data)))),"(none)")'
)


def test_audit_row_is_bold_label_value_pairs_with_response_count_cf() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_audit_row(_as_xw_sheet(sheet))

    expected = [
        (13, 14, "k", '=IFERROR(COLUMNS(X_s()),"(empty model)")'),
        (16, 17, "rows", '=IFERROR(ROWS(X_s()),"(empty model)")'),
        (19, 20, "response", f"={_RESPONSE_NAME}"),
        (
            21,
            22,
            "responses",
            '=SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))',
        ),
        (23, 24, "included rows", "=SUMPRODUCT(N(Sample_Include()))"),
        (
            25,
            26,
            "sequence flags",
            "=SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))",
        ),
    ]
    assert list(_AUDIT_PAIRS) == [(lc, vc) for lc, vc, _, _ in expected]
    assert _AUDIT_ROW == 1
    for label_col, value_col, label, formula in expected:
        # No audit cell may land on a width-2 break column (O=15, R=18).
        assert label_col not in (_C_BREAK_LEFT, _C_BREAK_MID)
        assert value_col not in (_C_BREAK_LEFT, _C_BREAK_MID)
        assert sheet.cell(1, label_col).value == label
        assert sheet.cell(1, value_col).api.Formula2 == formula
        assert (
            sheet.range((1, label_col), (1, value_col)).api.Font.Bold is True
        )

    # Exactly-one-Response validation: red CF on the responses count cell.
    conditions = sheet.range("$V$1").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == ["=N($V$1)<>1"]
    assert conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Zero-or-one-Sequence validation: red CF only at two-plus flags.
    seq_conditions = sheet.range("$Z$1").api.FormatConditions.items
    assert [c.Formula1 for c in seq_conditions] == ["=N($Z$1)>1"]
    assert seq_conditions[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert seq_conditions[0].Font.Color == excel_color(CF_DARK_RED_TEXT)


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

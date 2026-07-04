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

from lambda_catalog.sheet_styles import (
    CF_DARK_RED_TEXT,
    CF_LIGHT_RED_FILL,
    INPUT_COLOR,
    MUTED_TEXT_COLOR,
)
from lambda_catalog.workbook_helpers import excel_color
from lambda_catalog.write_sheet_model_construction import (
    _DEFAULT_SPEC,
    _FALLBACK_SPEC,
    _FIRST_DATA_ROW,
    _LAST_DATA_ROW,
    _N_VARIABLES,
    _VARIABLES,
    _set_sheet_scoped_names,
    _write_row_zones,
    _write_spec_block,
    SHEET_NAME,
)
from tests.recording_sheet import RecordingSheet

ROOT_DIR = Path(__file__).resolve().parents[1]

_EXPECTED_NAME_ORDER = [
    "Source_Data",
    "Header_Names",
    "Spec_Role",
    "Spec_Include",
    "Spec_Type",
    "Spec_Reference",
    "Spec_Order",
    "Spec_Transform",
    "Sample_Include",
    "Response_Column",
    "Row_Labels",
    "X_s",
    "Constructed_Column_Names",
]


def _as_xw_sheet(sheet: RecordingSheet) -> xw.Sheet:
    return cast(xw.Sheet, sheet)


def _named_sheet() -> RecordingSheet:
    sheet = RecordingSheet(name=SHEET_NAME)
    _set_sheet_scoped_names(_as_xw_sheet(sheet))
    return sheet


def _refers_to(sheet: RecordingSheet, name: str) -> str:
    return sheet.api.Names.by_short_name(name).RefersTo


def _all_written_formulas(sheet: RecordingSheet) -> list[str]:
    return [
        rng.state.formula2
        for rng in sheet.ranges.values()
        if rng.state.formula2 is not None
    ]


def test_names_are_created_in_dependency_order() -> None:
    sheet = _named_sheet()

    local_name_order = [item.Name.split("!", 1)[-1] for item in sheet.api.Names.items]
    assert local_name_order == _EXPECTED_NAME_ORDER


def test_only_the_retarget_names_reference_the_table_directly() -> None:
    sheet = _named_sheet()

    assert _refers_to(sheet, "Source_Data") == "=LifeExpectancyData[#Data]"
    assert _refers_to(sheet, "Header_Names") == "=LifeExpectancyData[#Headers]"
    for name in _EXPECTED_NAME_ORDER[2:]:
        assert "LifeExpectancyData" not in _refers_to(sheet, name), name

    _write_spec_block(_as_xw_sheet(sheet))
    _write_row_zones(_as_xw_sheet(sheet))
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
    ):
        assert _refers_to(sheet, name) == (
            f"='{SHEET_NAME}'!${column}$3:${column}$16000"
        )


def test_sample_include_is_the_reduce_product_mask() -> None:
    sheet = _named_sheet()
    mask = _refers_to(sheet, "Sample_Include")

    assert mask.startswith("=LAMBDA(LET(")
    # Filter columns: truthy — TRUE and 1 pass, FALSE/0/blank/text fail.
    assert 'IF(INDEX(rl,j)="Filter",acc*N(IFERROR(N(col)=1,FALSE))' in mask
    # Completeness: the Response and every included Continuous Predictor.
    assert (
        'IF(OR(INDEX(rl,j)="Response",'
        'AND(INDEX(rl,j)="Predictor",INDEX(inc,j)=TRUE,'
        'INDEX(typ,j)="Continuous")),acc*N(ISNUMBER(col)),acc)'
    ) in mask
    # Full-height ones seed; product over {0,1} is the AND, no per-row loop.
    assert "seed,SEQUENCE(ROWS(Source_Data),1,1,0)" in mask
    assert "BYROW(" not in mask
    assert mask.endswith("prod=1))")
    # Reads the model axes only — never the reserved columns.
    for reserved in ("Spec_Order", "Spec_Transform"):
        assert reserved not in mask


def test_row_labels_dispatches_on_identifier_presence() -> None:
    sheet = _named_sheet()
    labels = _refers_to(sheet, "Row_Labels")

    assert labels.startswith("=LAMBDA(LET(")
    # The LET-bound FILTER is wrapped in IFERROR so the all-FALSE case is
    # still safe at binding time.
    assert (
        'ids,IFERROR(TRANSPOSE(FILTER(TRANSPOSE(Source_Data),'
        'rl="Identifier")),NA())'
    ) in labels
    assert 'IF(SUM(--(rl="Identifier"))=0,' in labels
    # No Identifier columns: positional fallback, full height.
    assert '"Obs. "&SEQUENCE(ROWS(Source_Data))' in labels
    # ignore_empty=FALSE keeps field positions aligned across rows.
    assert 'BYROW(ids,LAMBDA(r,TEXTJOIN("|",FALSE,r)))' in labels


def test_row_zones_spill_full_height_next_to_the_spec_block() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_row_zones(_as_xw_sheet(sheet))

    # I: narrow gap, visually reserving the future Design Columns column.
    assert sheet.range((1, 9)).column_width == 2

    # J/K headers on the spec-header row, bold like the A–H headers.
    assert sheet.cell(2, 10).value == "Row Labels"
    assert sheet.cell(2, 11).value == "Included"
    assert sheet.range((2, 10), (2, 11)).api.Font.Bold is True

    # Full-height spills at row 3; row 1 stays empty for the next PR's
    # audit cells.
    assert sheet.cell(3, 10).api.Formula2 == "=Row_Labels()"
    assert sheet.cell(3, 11).api.Formula2 == "=Sample_Include()"
    for col in (10, 11):
        assert sheet.cell(1, col).value is None
        assert sheet.cell(1, col).api.Formula2 is None


def test_x_s_binds_dummy_levels_once_and_skips_on_isna() -> None:
    sheet = _named_sheet()
    x_s = _refers_to(sheet, "X_s")

    assert x_s.startswith("=LAMBDA(LET(")
    assert x_s.count("Dummy_Levels(") == 1
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in x_s
    assert "IF(ISNA(lv),acc,HSTACK(acc,--(col=lv)))" in x_s
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
    predicate = 'IF(OR(INDEX(rl,j)<>"Predictor",INDEX(inc,j)<>TRUE),acc,'
    assert predicate in x_s
    assert predicate in names
    assert 'lv,Dummy_Levels(col,r,Sample_Include())' in names
    assert names.count("Dummy_Levels(") == 1
    assert "IF(ISNA(lv),acc," in names
    # Level-qualified headers: "Status: Developing", "Year: 2001", ...
    assert 'HSTACK(acc,h&": "&lv)' in names
    assert names.endswith("DROP(built,,1)))")


def test_reserved_spec_names_are_defined_but_read_by_nothing() -> None:
    sheet = _named_sheet()
    _write_spec_block(_as_xw_sheet(sheet))
    _write_row_zones(_as_xw_sheet(sheet))

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
    _write_spec_block(_as_xw_sheet(sheet))
    _write_row_zones(_as_xw_sheet(sheet))

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
        for col in range(2, 8):
            assert sheet.cell(row, col).color == INPUT_COLOR, (variable, col)

    # Spot-check the named T0 roles.
    by_variable = {v: _FIRST_DATA_ROW + i for i, v in enumerate(_VARIABLES)}
    assert sheet.cell(by_variable["Country"], 2).value == "Identifier"
    assert sheet.cell(by_variable["Life expectancy"], 2).value == "Response"
    assert sheet.cell(by_variable["Full_Data"], 2).value == "Filter"
    for categorical in ("Year", "Status"):
        row = by_variable[categorical]
        assert sheet.cell(row, 2).value == "Predictor"
        assert sheet.cell(row, 3).value is True
        assert sheet.cell(row, 4).value == "Categorical"


def test_levels_column_counts_raw_levels_without_dummy_levels() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    formula = cast(str, sheet.cell(_FIRST_DATA_ROW, 8).api.Formula2)
    # Must display L itself (1 for degenerate columns, which Dummy_Levels
    # signals as #N/A instead), so it counts UNIQUE directly.
    assert formula.startswith(
        f'=IF(OR($B{_FIRST_DATA_ROW}<>"Predictor",$D{_FIRST_DATA_ROW}<>"Categorical"),"",'
    )
    assert "ROWS(UNIQUE(FILTER(" in formula
    assert "Dummy_Levels" not in formula
    assert "Sample_Include()" in formula
    assert "ROW()-2" in formula  # sheet row 3 → Source_Data column 1
    assert 'x,IF(col="","",col)' in formula  # blank normalization mirrored


def test_dropdowns_cover_exactly_the_four_list_columns() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    validated = {
        key: rng.api.Validation
        for key, rng in sheet.ranges.items()
        if rng.api.Validation.rules
    }
    assert set(validated) == {
        ((3, 2), (16000, 2)),  # B Role
        ((3, 3), (16000, 3)),  # C Include
        ((3, 4), (16000, 4)),  # D Type
        ((3, 7), (16000, 7)),  # G Transform (reserved; only "None" valid)
    }
    formulas = {
        key[0][1]: validation.rules[0]["Formula1"]
        for key, validation in validated.items()
    }
    assert formulas[2] == "Response,Predictor,Identifier,Filter,Omit"
    assert formulas[3] == "TRUE,FALSE"
    assert formulas[4] == "Continuous,Categorical"
    assert formulas[7] == "None"
    for validation in validated.values():
        assert validation.delete_count == 1
        assert validation.rules[0]["Type"] == 3  # xlValidateList
        assert validation.IgnoreBlank is True


def test_conditional_formats_cover_gray_cascade_degeneracy_and_reference() -> None:
    sheet = RecordingSheet(name=SHEET_NAME)
    _write_spec_block(_as_xw_sheet(sheet))

    gray = sheet.range(f"$C$3:$H${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in gray] == ['=$B3<>"Predictor"']
    assert gray[0].Font.Color == excel_color(MUTED_TEXT_COLOR)

    degenerate = sheet.range(f"$H$3:$H${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in degenerate] == [
        '=AND($B3="Predictor",$C3=TRUE,$D3="Categorical",N($H3)<=1)'
    ]
    assert degenerate[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert degenerate[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

    # Invalid reference: the constructor's exact skip condition, tested
    # directly — a membership test against Dummy_Levels' output would
    # false-positive on the default reference itself.
    invalid = sheet.range(f"$E$3:$E${_LAST_DATA_ROW}").api.FormatConditions.items
    assert [c.Formula1 for c in invalid] == [
        '=AND($E3<>"",ISNA(Dummy_Levels(INDEX(Source_Data,0,ROW()-2),'
        "$E3,Sample_Include())))"
    ]
    assert invalid[0].Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert invalid[0].Font.Color == excel_color(CF_DARK_RED_TEXT)

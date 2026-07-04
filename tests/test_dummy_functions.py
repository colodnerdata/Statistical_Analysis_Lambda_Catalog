"""Static and reference-mirror tests for the rebuilt Dummy_Levels / Dummy_Code.

The v3.0 rebuild changed the failure contract: degenerate input returns a
genuine Excel error (NA()) instead of a descriptive string. Excel itself is
exercised by the Dummy_Test QC sheet (see write_sheet_dummy_test.py and
build_qc.py); these tests cover everything checkable without Excel — the
formula strings, the catalog metadata, the parser translation, and the
pure-Python mirrors that generate the QC sheet's expected values.
"""
from pathlib import Path

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.write_sheet_dummy_test import (
    CATEGORY_VALUES,
    CHECK_CASES,
    EXPECTED_DEFAULT_LEVELS,
    EXPECTED_EXPLICIT_LEVELS,
    INCLUDE_VALUES,
    SINGLE_LEVEL_VALUES,
    dummy_code_reference,
    dummy_levels_reference,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
_DOCUMENT = load_catalog_document(ROOT_DIR / "lambda_functions.json")
_BY_NAME = {fn.name: fn for fn in _DOCUMENT.functions}
_DUMMY_LEVELS = _BY_NAME["Dummy_Levels"]
_DUMMY_CODE = _BY_NAME["Dummy_Code"]

_LEGACY_ERROR_STRINGS = (
    "ERROR: only one category level",
    "ERROR: reference level not found",
    "ERROR: no included rows",
)


def test_formula_strings_are_balanced() -> None:
    for fn in (_DUMMY_LEVELS, _DUMMY_CODE):
        formula = fn.formula_display
        assert formula.count("(") == formula.count(")"), fn.name
        assert formula.count('"') % 2 == 0, fn.name


def test_failure_is_signaled_by_na_not_text() -> None:
    for fn in (_DUMMY_LEVELS, _DUMMY_CODE):
        assert "NA()" in fn.formula_display, fn.name
        assert "ISNA" in fn.formula_display, fn.name
        assert "ERROR:" not in fn.formula_display, fn.name


def test_no_legacy_error_strings_anywhere_in_catalog() -> None:
    raw = (ROOT_DIR / "lambda_functions.json").read_text(encoding="utf-8")
    for legacy in _LEGACY_ERROR_STRINGS:
        assert legacy not in raw, legacy


def test_signatures_unchanged() -> None:
    for fn in (_DUMMY_LEVELS, _DUMMY_CODE):
        assert fn.argument_names == ("category", "reference", "include"), fn.name
        required = [a.name for a in fn.arguments if not a.optional]
        assert required == ["category"], fn.name


def test_empty_string_reference_means_default() -> None:
    # A provided "" behaves exactly like an omitted reference (use the first
    # sorted level) — the Model Construction sheet passes blank Reference
    # cells straight through, and "" must mean "default", not "invalid".
    assert 'IF(ISOMITTED(reference), "", reference)' in _DUMMY_LEVELS.formula_display
    assert (
        'IF(ref_raw = "", IFERROR(INDEX(All_Levels, 1, 1), NA()), ref_raw)'
        in _DUMMY_LEVELS.formula_display
    )


def test_dummy_code_delegates_level_determination() -> None:
    # One source of truth: Dummy_Code must call Dummy_Levels with pass-through
    # arguments rather than re-deriving the level set.
    assert "Dummy_Levels(category, reference, include)" in _DUMMY_CODE.formula_display


def test_blank_cells_are_normalized_before_masking() -> None:
    # TOCOL coerces truly empty cells to numeric 0; without normalization a
    # blank category cell would become a spurious "0" level instead of being
    # excluded by the (x<>"") mask.
    for fn in (_DUMMY_LEVELS, _DUMMY_CODE):
        assert 'TOCOL(IF(category = "", "", category), 0)' in fn.formula_display, fn.name


def test_formulas_translate_to_workbook_xml() -> None:
    levels_xml = _DUMMY_LEVELS.workbook_xml_formula_from_display
    code_xml = _DUMMY_CODE.workbook_xml_formula_from_display
    assert "_xlfn.LAMBDA" in levels_xml
    assert "_xlfn.TOROW" in levels_xml
    assert "_xlfn._xlws.FILTER" in levels_xml
    assert "_xlfn.LAMBDA" in code_xml
    assert "_xlfn.MAKEARRAY" in code_xml


def test_test_table_tag_set() -> None:
    assert _DUMMY_LEVELS.test_table == "Dummy_Test"
    assert _DUMMY_CODE.test_table == "Dummy_Test"


# ---------------------------------------------------------------------------
# Pure-Python mirrors: pinned to independently hand-derived expectations, so
# the QC sheet's expected values are not self-certifying.
# ---------------------------------------------------------------------------


def test_fixture_shape() -> None:
    assert len(CATEGORY_VALUES) == len(INCLUDE_VALUES) == len(SINGLE_LEVEL_VALUES) == 8
    assert CATEGORY_VALUES[5] == ""  # blank category cell (row 6)
    assert INCLUDE_VALUES[6] is False  # excluded row (row 7)


def test_mirror_levels_clean_binary_default_reference() -> None:
    # Levels {Developed, Developing}; default reference = first sorted level.
    assert dummy_levels_reference(CATEGORY_VALUES) == ["Developing"]
    assert EXPECTED_DEFAULT_LEVELS == ["Developing"]


def test_mirror_levels_empty_string_reference_defaults() -> None:
    # "" and None request the same default reference.
    assert dummy_levels_reference(CATEGORY_VALUES, "") == ["Developing"]
    assert dummy_code_reference(CATEGORY_VALUES, "") == dummy_code_reference(
        CATEGORY_VALUES
    )


def test_mirror_levels_explicit_valid_reference() -> None:
    assert dummy_levels_reference(CATEGORY_VALUES, "Developing") == ["Developed"]
    assert EXPECTED_EXPLICIT_LEVELS == ["Developed"]


def test_mirror_levels_failure_conditions_return_none() -> None:
    # Invalid reference — must fail, never silently one-hot.
    assert dummy_levels_reference(CATEGORY_VALUES, "Developped") is None
    # Single level — dropping the reference leaves nothing.
    assert dummy_levels_reference(SINGLE_LEVEL_VALUES) is None
    # Everything excluded by the mask.
    assert dummy_levels_reference(CATEGORY_VALUES, "Developed", False) is None
    # No non-blank values at all.
    assert dummy_levels_reference([""] * 8) is None


def test_mirror_code_default_matrix() -> None:
    # include omitted: active = non-blank rows; retained level = Developing.
    expected = [[0], [1], [0], [1], [1], [""], [0], [1]]
    assert dummy_code_reference(CATEGORY_VALUES) == expected


def test_mirror_code_masked_matrix() -> None:
    # Row 6 blank and row 7 include=FALSE both emit "" (full height preserved).
    expected = [[0], [1], [0], [1], [1], [""], [""], [1]]
    assert dummy_code_reference(CATEGORY_VALUES, "Developed", INCLUDE_VALUES) == expected


def test_mirror_code_failure_conditions_return_none() -> None:
    assert dummy_code_reference(CATEGORY_VALUES, "Developped", INCLUDE_VALUES) is None
    assert dummy_code_reference(SINGLE_LEVEL_VALUES) is None


def test_check_case_formulas_are_wellformed_and_unique() -> None:
    labels = [label for label, _ in CHECK_CASES]
    assert len(labels) == len(set(labels))
    for label, formula in CHECK_CASES:
        assert formula.startswith("="), label
        assert formula.count("(") == formula.count(")"), label
        assert formula.count('"') % 2 == 0, label

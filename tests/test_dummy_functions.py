"""Static tests for the rebuilt Dummy_Levels / Dummy_Code catalog formulas.

The v3.0 rebuild changed the failure contract: degenerate input returns a
genuine Excel error (NA()) instead of a descriptive string. These tests cover
what is checkable without Excel — the formula strings, the signatures, and the
parser translation to workbook XML.

The functions' Excel-side behaviour is exercised through the Regression
engine's own spec-driven verifier; ``analyze_model_construction._masked_levels``
carries the pure-Python mirror of ``Dummy_Levels`` used by that oracle.
"""
from pathlib import Path

from lambda_catalog.catalog_schema import load_catalog_document

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


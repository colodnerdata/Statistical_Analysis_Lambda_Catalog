"""Tests for lambda_formula_parser — LAMBDA formula → workbook.xml translation."""
from __future__ import annotations

import unittest

from lambda_catalog.lambda_formula_parser import (
    _consume_string_literal,
    _normalize_user_formula,
    _split_lambda_signature,
    _split_top_level_arguments,
    _strip_non_string_whitespace,
    to_workbook_xml_formula,
    to_workbook_xml_formula_from_display,
)


class NormalizeUserFormulaTests(unittest.TestCase):
    def test_strips_leading_equals(self) -> None:
        self.assertEqual(_normalize_user_formula("=LAMBDA(x,x+1)"), "LAMBDA(x,x+1)")

    def test_no_change_when_no_equals(self) -> None:
        self.assertEqual(_normalize_user_formula("LAMBDA(x,x+1)"), "LAMBDA(x,x+1)")

    def test_strips_surrounding_whitespace(self) -> None:
        self.assertEqual(_normalize_user_formula("  =LAMBDA(x,x+1)  "), "LAMBDA(x,x+1)")

    def test_empty_string_unchanged(self) -> None:
        self.assertEqual(_normalize_user_formula(""), "")


class StripNonStringWhitespaceTests(unittest.TestCase):
    def test_removes_spaces_outside_strings(self) -> None:
        self.assertEqual(_strip_non_string_whitespace("A + B"), "A+B")

    def test_preserves_spaces_inside_double_quoted_string(self) -> None:
        self.assertEqual(_strip_non_string_whitespace('"Hello World"'), '"Hello World"')

    def test_mixed_content(self) -> None:
        self.assertEqual(
            _strip_non_string_whitespace('IF(x, "yes value", 0)'),
            'IF(x,"yes value",0)',
        )

    def test_newlines_and_tabs_removed_outside_strings(self) -> None:
        result = _strip_non_string_whitespace("LAMBDA(\n  x,\n\tx+1\n)")
        self.assertEqual(result, "LAMBDA(x,x+1)")

    def test_escaped_quote_inside_string_preserved(self) -> None:
        result = _strip_non_string_whitespace('"say ""hi"""')
        self.assertEqual(result, '"say ""hi"""')


class ConsumeStringLiteralTests(unittest.TestCase):
    def test_simple_string(self) -> None:
        text = '"hello"rest'
        end = _consume_string_literal(text, 0)
        self.assertEqual(end, 7)
        self.assertEqual(text[:end], '"hello"')

    def test_escaped_quote(self) -> None:
        text = '"say ""hi"""rest'
        end = _consume_string_literal(text, 0)
        self.assertEqual(text[:end], '"say ""hi"""')

    def test_raises_on_unterminated_string(self) -> None:
        with self.assertRaises(ValueError):
            _consume_string_literal('"unterminated', 0)


class SplitLambdaSignatureTests(unittest.TestCase):
    def test_single_parameter(self) -> None:
        params, body = _split_lambda_signature("LAMBDA(x,x+1)")
        self.assertEqual(params, ["x"])
        self.assertEqual(body, "x+1")

    def test_two_parameters(self) -> None:
        params, body = _split_lambda_signature("LAMBDA(a,b,a+b)")
        self.assertEqual(params, ["a", "b"])
        self.assertEqual(body, "a+b")

    def test_nested_function_in_body_not_split(self) -> None:
        params, body = _split_lambda_signature("LAMBDA(x,SUM(x,1))")
        self.assertEqual(params, ["x"])
        self.assertEqual(body, "SUM(x,1)")

    def test_raises_when_not_lambda(self) -> None:
        with self.assertRaises(ValueError):
            _split_lambda_signature("SUM(A1:A10)")

    def test_single_argument_is_a_zero_parameter_lambda(self) -> None:
        # Excel semantics: LAMBDA(body) is a valid zero-parameter lambda,
        # called as Name(). The catalog's zero-argument workbook accessors
        # (Base_Period_Delta) rely on this surviving the workbook.xml sync.
        params, body = _split_lambda_signature("LAMBDA(x)")
        self.assertEqual(params, [])
        self.assertEqual(body, "x")

    def test_raises_when_body_is_empty(self) -> None:
        with self.assertRaises(ValueError):
            _split_lambda_signature("LAMBDA(a,)")


class SplitTopLevelArgumentsTests(unittest.TestCase):
    def test_two_simple_args(self) -> None:
        result = _split_top_level_arguments("a,b")
        self.assertEqual(result, ["a", "b"])

    def test_nested_call_not_split(self) -> None:
        result = _split_top_level_arguments("SUM(a,b),c")
        self.assertEqual(result, ["SUM(a,b)", "c"])

    def test_string_literal_comma_not_split(self) -> None:
        result = _split_top_level_arguments('"a,b",c')
        self.assertEqual(result, ['"a,b"', "c"])

    def test_single_arg(self) -> None:
        self.assertEqual(_split_top_level_arguments("x"), ["x"])


class ToWorkbookXmlFormulaTests(unittest.TestCase):
    def test_identity_lambda(self) -> None:
        self.assertEqual(
            to_workbook_xml_formula("LAMBDA(x,x)"),
            "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)",
        )

    def test_two_parameter_lambda(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(a,b,a+b)")
        self.assertIn("_xlpm.a", result)
        self.assertIn("_xlpm.b", result)
        self.assertTrue(result.startswith("_xlfn.LAMBDA("))

    def test_excel_365_function_gets_prefix(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(x,SEQUENCE(x))")
        self.assertIn("_xlfn.SEQUENCE", result)

    def test_classic_function_not_prefixed(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(x,SUM(x))")
        self.assertIn("SUM(", result)
        self.assertNotIn("_xlfn.SUM(", result)
        self.assertNotIn("_xludf.SUM(", result)

    def test_filter_function_gets_xlws_prefix(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(x,FILTER(x,x>0))")
        self.assertIn("_xlfn._xlws.FILTER", result)

    def test_leading_equals_stripped(self) -> None:
        result = to_workbook_xml_formula("=LAMBDA(x,x*2)")
        self.assertEqual(result, "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)")

    def test_raises_when_not_a_lambda(self) -> None:
        with self.assertRaises(ValueError):
            to_workbook_xml_formula("SUM(A1:A10)")

    def test_lambda_prefix_applied(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(x,x+1)")
        self.assertTrue(result.startswith("_xlfn.LAMBDA("))

    def test_optional_parameter_uses_xlop_prefix(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(x,[y],x)")
        self.assertIn("_xlop.y", result)

    def test_zero_parameter_lambda_translates(self) -> None:
        result = to_workbook_xml_formula("LAMBDA(LET(n,COUNT(A:A),n))")
        self.assertEqual(result, "_xlfn.LAMBDA(_xlfn.LET(_xlpm.n,COUNT(A:A),_xlpm.n))")

    def test_sheet_qualified_name_reference_passes_through(self) -> None:
        # Workbook-scoped accessors may reference sheet-scoped names with an
        # explicit sheet prefix; the quoted prefix must survive untouched.
        result = to_workbook_xml_formula(
            "LAMBDA(XMATCH(TRUE,'Regression'!Spec_Sequence,0))"
        )
        self.assertEqual(
            result,
            "_xlfn.LAMBDA(_xlfn.XMATCH(TRUE,'Regression'!Spec_Sequence,0))",
        )


class ToWorkbookXmlFormulaFromDisplayTests(unittest.TestCase):
    def test_multiline_lambda_compacted(self) -> None:
        multiline = "=LAMBDA(\n  x,\n  x + 1\n)"
        result = to_workbook_xml_formula_from_display(multiline)
        self.assertEqual(result, "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)")

    def test_indented_lambda_with_let(self) -> None:
        formula = "=LAMBDA(data, LET(n, COUNT(data), n))"
        result = to_workbook_xml_formula_from_display(formula)
        self.assertIn("_xlfn.LET", result)
        self.assertIn("_xlpm.n", result)

    def test_preserves_string_literal_spaces(self) -> None:
        formula = 'LAMBDA(x, IF(x > 0, "yes value", "no value"))'
        result = to_workbook_xml_formula_from_display(formula)
        self.assertIn('"yes value"', result)
        self.assertIn('"no value"', result)


if __name__ == "__main__":
    unittest.main()

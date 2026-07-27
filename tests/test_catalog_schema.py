"""Tests for catalog_schema — CatalogDocument loading, validation, and projections."""
from __future__ import annotations

import unittest
from pathlib import Path

from lambda_catalog.catalog_schema import (
    CatalogArgument,
    CatalogDocument,
    CatalogFunction,
    load_catalog_document,
)


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


def _minimal_function(**overrides) -> dict:
    base = {
        "name": "My_Func",
        "formula_display": "=LAMBDA(x, x)",
        "arguments": [],
        "yields": "scalar",
        "description": "A test function.",
        "plain_language_summary": "Does a thing.",
    }
    base.update(overrides)
    return base


def _payload(*functions, notes=None) -> dict:
    result: dict = {"functions": list(functions)}
    if notes is not None:
        result["regression_sheet_notes"] = notes
    return result


class LoadCatalogDocumentValidDocumentTests(unittest.TestCase):
    def test_minimal_valid_document_parses(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload(_minimal_function()))
        self.assertIsInstance(doc, CatalogDocument)
        self.assertEqual(len(doc.functions), 1)
        self.assertIsInstance(doc.functions[0], CatalogFunction)

    def test_function_fields_are_mapped_correctly(self) -> None:
        fn = _minimal_function(
            name="Test_Fn",
            formula_display="=LAMBDA(x, x + 1)",
            yields="number",
            description="Returns x plus one.",
            plain_language_summary="Adds one.",
            test_table="MLR_Scalar_Test",
            number_format="0.000",
        )
        doc = load_catalog_document(Path(), payload=_payload(fn))
        cf = doc.functions[0]
        self.assertEqual(cf.name, "Test_Fn")
        self.assertEqual(cf.formula_display, "=LAMBDA(x, x + 1)")
        self.assertEqual(cf.yields, "number")
        self.assertEqual(cf.description, "Returns x plus one.")
        self.assertEqual(cf.plain_language_summary, "Adds one.")
        self.assertEqual(cf.test_table, "MLR_Scalar_Test")
        self.assertEqual(cf.number_format, "0.000")

    def test_number_format_defaults_to_general(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload(_minimal_function()))
        self.assertEqual(doc.functions[0].number_format, "General")

    def test_scope_defaults_to_workbook(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload(_minimal_function()))
        self.assertEqual(doc.functions[0].scope, "workbook")

    def test_scope_parsed_when_present(self) -> None:
        fn = _minimal_function(scope="Model Construction")
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertEqual(doc.functions[0].scope, "Model Construction")

    def test_blank_scope_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(Path(), payload=_payload(_minimal_function(scope="  ")))

    def test_non_string_scope_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(Path(), payload=_payload(_minimal_function(scope=3)))

    def test_workbook_functions_and_functions_for_sheet_partition(self) -> None:
        payload = _payload(
            _minimal_function(name="Portable"),
            _minimal_function(name="Local_A", scope="Model Construction"),
            _minimal_function(name="Local_B", scope="Model Construction"),
        )
        doc = load_catalog_document(Path(), payload=payload)
        self.assertEqual([f.name for f in doc.workbook_functions], ["Portable"])
        self.assertEqual(
            [f.name for f in doc.functions_for_sheet("Model Construction")],
            ["Local_A", "Local_B"],
        )

    def test_test_table_defaults_to_none(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload(_minimal_function()))
        self.assertIsNone(doc.functions[0].test_table)

    def test_function_order_is_preserved(self) -> None:
        names = ["Alpha", "Beta", "Gamma"]
        fns = [_minimal_function(name=n) for n in names]
        doc = load_catalog_document(Path(), payload=_payload(*fns))
        self.assertEqual([f.name for f in doc.functions], names)

    def test_empty_functions_array_is_valid(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload())
        self.assertEqual(doc.functions, ())

    def test_regression_sheet_notes_parsed(self) -> None:
        notes = {"note_a": "Text A", "note_b": "Text B"}
        doc = load_catalog_document(Path(), payload=_payload(notes=notes))
        self.assertEqual(doc.regression_sheet_notes, notes)

    def test_regression_sheet_notes_absent_gives_empty_dict(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload())
        self.assertEqual(doc.regression_sheet_notes, {})

    def test_leading_trailing_whitespace_stripped_from_name(self) -> None:
        fn = _minimal_function(name="  Trimmed  ")
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertEqual(doc.functions[0].name, "Trimmed")

    def test_functions_field_is_tuple(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload(_minimal_function()))
        self.assertIsInstance(doc.functions, tuple)

    def test_document_is_frozen(self) -> None:
        doc = load_catalog_document(Path(), payload=_payload())
        with self.assertRaises((AttributeError, TypeError)):
            doc.functions = ()  # type: ignore[misc]


class LoadCatalogDocumentDuplicateNameTests(unittest.TestCase):
    def test_duplicate_function_names_rejected(self) -> None:
        fn1 = _minimal_function(name="Dup")
        fn2 = _minimal_function(name="Dup")
        with self.assertRaises(ValueError) as ctx:
            load_catalog_document(Path(), payload=_payload(fn1, fn2))
        self.assertIn("Dup", str(ctx.exception))

    def test_unique_names_accepted(self) -> None:
        fn1 = _minimal_function(name="Alpha")
        fn2 = _minimal_function(name="Beta")
        doc = load_catalog_document(Path(), payload=_payload(fn1, fn2))
        self.assertEqual(len(doc.functions), 2)


class LoadCatalogDocumentMissingRequiredFieldTests(unittest.TestCase):
    def _assert_raises_for(self, fn_dict: dict) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(Path(), payload=_payload(fn_dict))

    def test_missing_name_rejected(self) -> None:
        fn = _minimal_function()
        del fn["name"]
        self._assert_raises_for(fn)

    def test_blank_name_rejected(self) -> None:
        self._assert_raises_for(_minimal_function(name="   "))

    def test_missing_formula_display_rejected(self) -> None:
        fn = _minimal_function()
        del fn["formula_display"]
        self._assert_raises_for(fn)

    def test_blank_formula_display_rejected(self) -> None:
        self._assert_raises_for(_minimal_function(formula_display=""))

    def test_missing_yields_rejected(self) -> None:
        fn = _minimal_function()
        del fn["yields"]
        self._assert_raises_for(fn)

    def test_blank_yields_rejected(self) -> None:
        self._assert_raises_for(_minimal_function(yields="  "))

    def test_missing_description_rejected(self) -> None:
        fn = _minimal_function()
        del fn["description"]
        self._assert_raises_for(fn)

    def test_blank_description_rejected(self) -> None:
        self._assert_raises_for(_minimal_function(description=""))

    def test_missing_plain_language_summary_rejected(self) -> None:
        fn = _minimal_function()
        del fn["plain_language_summary"]
        self._assert_raises_for(fn)

    def test_blank_plain_language_summary_rejected(self) -> None:
        self._assert_raises_for(_minimal_function(plain_language_summary=""))

    def test_whitespace_only_plain_language_summary_rejected(self) -> None:
        self._assert_raises_for(_minimal_function(plain_language_summary="   "))

    def test_no_functions_array_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(Path(), payload={"not_functions": []})

    def test_functions_not_array_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(Path(), payload={"functions": "wrong"})

    def test_non_object_entry_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(Path(), payload={"functions": ["not_a_dict"]})


class LoadCatalogDocumentTestTableValidationTests(unittest.TestCase):
    def test_valid_test_table_accepted(self) -> None:
        fn = _minimal_function(test_table="MLR_Scalar_Test")
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertEqual(doc.functions[0].test_table, "MLR_Scalar_Test")

    def test_invalid_test_table_too_long_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(test_table="A" * 32))
            )

    def test_invalid_test_table_contains_bracket_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(test_table="Bad[Name]"))
            )

    def test_invalid_test_table_starts_with_digit_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(test_table="1BadName"))
            )

    def test_invalid_test_table_looks_like_cell_ref_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(test_table="A1"))
            )

    def test_blank_test_table_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(test_table=""))
            )

    def test_none_test_table_is_allowed(self) -> None:
        fn = _minimal_function()
        fn["test_table"] = None
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertIsNone(doc.functions[0].test_table)


class LoadCatalogDocumentRegressionSheetNotesTests(unittest.TestCase):
    def test_non_dict_notes_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload={"functions": [], "regression_sheet_notes": "bad"}
            )

    def test_non_string_value_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(),
                payload={"functions": [], "regression_sheet_notes": {"key": 42}},
            )


class LoadCatalogDocumentArgumentTests(unittest.TestCase):
    def test_argument_ordering_preserved(self) -> None:
        fn = _minimal_function(arguments=[
            {"name": "first", "description": "First arg"},
            {"name": "second", "description": "Second arg"},
            {"name": "third", "description": "Third arg"},
        ])
        doc = load_catalog_document(Path(), payload=_payload(fn))
        cf = doc.functions[0]
        self.assertEqual(cf.argument_names, ("first", "second", "third"))

    def test_optional_flag_preserved(self) -> None:
        fn = _minimal_function(arguments=[
            {"name": "required_arg", "description": "Required."},
            {"name": "optional_arg", "description": "Optional.", "optional": True},
        ])
        doc = load_catalog_document(Path(), payload=_payload(fn))
        args = doc.functions[0].arguments
        self.assertFalse(args[0].optional)
        self.assertTrue(args[1].optional)

    def test_arguments_tuple_is_frozen(self) -> None:
        fn = _minimal_function(arguments=[{"name": "x", "description": "The x."}])
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertIsInstance(doc.functions[0].arguments, tuple)

    def test_empty_arguments_list_is_valid(self) -> None:
        fn = _minimal_function(arguments=[])
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertEqual(doc.functions[0].arguments, ())

    def test_absent_arguments_field_treated_as_empty(self) -> None:
        fn = _minimal_function()
        del fn["arguments"]
        doc = load_catalog_document(Path(), payload=_payload(fn))
        self.assertEqual(doc.functions[0].arguments, ())

    def test_non_list_arguments_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(arguments="wrong"))
            )

    def test_non_dict_argument_entry_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(), payload=_payload(_minimal_function(arguments=["not_a_dict"]))
            )

    def test_missing_argument_name_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(),
                payload=_payload(_minimal_function(arguments=[{"description": "No name."}])),
            )

    def test_blank_argument_name_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(),
                payload=_payload(
                    _minimal_function(arguments=[{"name": "  ", "description": "Blank name."}])
                ),
            )

    def test_missing_argument_description_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(),
                payload=_payload(_minimal_function(arguments=[{"name": "x"}])),
            )

    def test_blank_argument_description_rejected(self) -> None:
        with self.assertRaises(ValueError):
            load_catalog_document(
                Path(),
                payload=_payload(
                    _minimal_function(arguments=[{"name": "x", "description": ""}])
                ),
            )


class CatalogArgumentProjectionTests(unittest.TestCase):
    def test_display_name_required_arg(self) -> None:
        arg = CatalogArgument(name="x", description="Some x.", optional=False)
        self.assertEqual(arg.display_name(), "x")

    def test_display_name_optional_arg(self) -> None:
        arg = CatalogArgument(name="x", description="Some x.", optional=True)
        self.assertEqual(arg.display_name(), "[x]")


class CatalogFunctionProjectionTests(unittest.TestCase):
    def _make_function(self, **overrides) -> CatalogFunction:
        fn = _minimal_function(**overrides)
        doc = load_catalog_document(Path(), payload=_payload(fn))
        return doc.functions[0]

    def test_argument_names_returns_tuple_of_names(self) -> None:
        cf = self._make_function(arguments=[
            {"name": "x_s", "description": "Predictors."},
            {"name": "y", "description": "Target."},
        ])
        self.assertEqual(cf.argument_names, ("x_s", "y"))

    def test_argument_names_empty_when_no_arguments(self) -> None:
        cf = self._make_function(arguments=[])
        self.assertEqual(cf.argument_names, ())

    def test_name_manager_comment_required_args(self) -> None:
        cf = self._make_function(arguments=[
            {"name": "x_s", "description": "Predictors."},
            {"name": "y", "description": "Target."},
        ])
        self.assertEqual(cf.name_manager_comment, "x_s: Predictors.\n\ny: Target.")

    def test_name_manager_comment_optional_arg_uses_brackets(self) -> None:
        cf = self._make_function(arguments=[
            {"name": "x_s", "description": "Predictors."},
            {"name": "filter", "description": "Row mask.", "optional": True},
        ])
        self.assertIn("[filter]: Row mask.", cf.name_manager_comment)

    def test_name_manager_comment_empty_when_no_args(self) -> None:
        cf = self._make_function(arguments=[])
        self.assertEqual(cf.name_manager_comment, "")

    def test_arguments_cell_text_uses_single_newline_separator(self) -> None:
        cf = self._make_function(arguments=[
            {"name": "x_s", "description": "Predictors."},
            {"name": "y", "description": "Target."},
        ])
        self.assertEqual(cf.arguments_cell_text(), "x_s: Predictors.\ny: Target.")

    def test_arguments_cell_text_empty_when_no_args(self) -> None:
        cf = self._make_function(arguments=[])
        self.assertEqual(cf.arguments_cell_text(), "")

    def test_arguments_cell_text_optional_arg_uses_brackets(self) -> None:
        cf = self._make_function(arguments=[
            {"name": "filter", "description": "Row mask.", "optional": True},
        ])
        self.assertEqual(cf.arguments_cell_text(), "[filter]: Row mask.")

    def test_workbook_xml_formula_delegates_to_parser(self) -> None:
        cf = self._make_function(formula_display="=LAMBDA(x, x + 1)")
        xml_formula = cf.workbook_xml_formula_from_display
        self.assertTrue(xml_formula.startswith("_xlfn.LAMBDA("))
        self.assertIn("_xlpm.x", xml_formula)

    def test_name_manager_comment_separator_is_double_newline(self) -> None:
        cf = self._make_function(arguments=[
            {"name": "a", "description": "First."},
            {"name": "b", "description": "Second."},
        ])
        self.assertIn("\n\n", cf.name_manager_comment)
        self.assertNotIn("\n\n", cf.arguments_cell_text())

    def test_notes_field_round_trips_through_loader(self) -> None:
        cf = self._make_function(notes="A 255-character-or-shorter tooltip.")
        self.assertEqual(cf.notes, "A 255-character-or-shorter tooltip.")

    def test_notes_missing_defaults_to_empty_string(self) -> None:
        cf = self._make_function()
        self.assertEqual(cf.notes, "")

    def test_notes_over_255_chars_raises_value_error(self) -> None:
        with self.assertRaises(ValueError):
            self._make_function(notes="x" * 256)


class RealCatalogIntegrationTests(unittest.TestCase):
    """Smoke tests that load the actual lambda_functions.json on disk."""

    def setUp(self) -> None:
        self.document = load_catalog_document(DEFINITIONS_PATH)

    def test_document_has_functions(self) -> None:
        self.assertTrue(self.document.functions)

    def test_all_functions_have_plain_language_summary(self) -> None:
        for fn in self.document.functions:
            self.assertTrue(
                fn.plain_language_summary,
                f"Function {fn.name!r} has a blank plain_language_summary",
            )

    def test_all_function_names_are_unique(self) -> None:
        names = [fn.name for fn in self.document.functions]
        self.assertEqual(len(names), len(set(names)))

    def test_functions_with_test_table_have_valid_tags(self) -> None:
        tagged = [fn for fn in self.document.functions if fn.test_table is not None]
        self.assertTrue(tagged, "Expected at least some functions with test_table set")

    def test_regression_sheet_notes_is_string_mapping(self) -> None:
        for key, value in self.document.regression_sheet_notes.items():
            self.assertIsInstance(key, str)
            self.assertIsInstance(value, str)

    def test_load_twice_with_payload_gives_same_result(self) -> None:
        import json
        with DEFINITIONS_PATH.open("r", encoding="utf-8") as fh:
            payload = json.load(fh)
        doc_from_file = load_catalog_document(DEFINITIONS_PATH)
        doc_from_payload = load_catalog_document(DEFINITIONS_PATH, payload=payload)
        self.assertEqual(
            [f.name for f in doc_from_file.functions],
            [f.name for f in doc_from_payload.functions],
        )

    def test_model_construction_closures_are_sheet_scoped(self) -> None:
        # v3.0 changeover: the constructor closures are registered on the
        # Regression sheet (the spec block moved there).
        closures = self.document.functions_for_sheet("Regression")
        self.assertEqual(
            [f.name for f in closures],
            [
                "Sample_Include",
                "Response_Column",
                "Row_Labels",
                "X_s",
                "Constructed_Column_Names",
                "Constructed_Column_Transforms",
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
            ],
        )
        # Sheet-scoped closures must never be synced as workbook names — a
        # workbook-scoped X_s would carry Model-Construction-relative refs into
        # the global namespace.
        workbook_names = {f.name for f in self.document.workbook_functions}
        for closure in closures:
            self.assertNotIn(closure.name, workbook_names)

    def test_grid_search_helpers_load_with_expected_contracts(self) -> None:
        functions = {fn.name: fn for fn in self.document.functions}
        self.assertIn("Grid_Argument_Minimum", functions)
        self.assertIn("Grid_Search_Optimum", functions)
        self.assertEqual(functions["Grid_Argument_Minimum"].argument_names, ("grid",))
        self.assertEqual(functions["Grid_Search_Optimum"].argument_names, ("grid",))

    def test_grid_search_optimum_uses_scalar_index_and_native_offset(self) -> None:
        functions = {fn.name: fn for fn in self.document.functions}
        display = functions["Grid_Search_Optimum"].formula_display.replace(" ", "")
        self.assertIn("INDEX(argmin,1,2)", display)
        self.assertIn("INDEX(argmin,1,3)", display)
        self.assertNotIn("CHOOSECOLS", display)
        self.assertIn("OFFSET(grid", display)

        xml = functions["Grid_Search_Optimum"].workbook_xml_formula_from_display
        self.assertIn("INDEX(_xlpm.argmin,1,2)", xml)
        self.assertIn("INDEX(_xlpm.argmin,1,3)", xml)
        self.assertIn("OFFSET(_xlpm.grid", xml)
        self.assertNotIn("_xlfn.OFFSET", xml)
        self.assertIn("_xlfn.VSTACK", xml)

    def test_grid_argmin_dynamic_functions_receive_parser_prefixes(self) -> None:
        functions = {fn.name: fn for fn in self.document.functions}
        xml = functions["Grid_Argument_Minimum"].workbook_xml_formula_from_display
        for function_name in ("LAMBDA", "LET", "HSTACK", "XMATCH", "TOCOL"):
            self.assertIn(f"_xlfn.{function_name}", xml)

    def test_loocv_residual_is_cataloged_and_used_by_press(self) -> None:
        functions = {fn.name: fn for fn in self.document.functions}
        self.assertIn("LOOCV_Residual", functions)
        self.assertEqual(
            functions["LOOCV_Residual"].argument_names,
            ("X_s", "Y", "Allow_Intercept", "Include"),
        )
        self.assertIn(
            "Residuals(X_s, Y, allow_arg, filt_arg)",
            functions["LOOCV_Residual"].formula_display,
        )
        self.assertIn(
            "Hat_Diagonal(X_s, allow_arg, filt_arg)",
            functions["LOOCV_Residual"].formula_display,
        )
        self.assertIn(
            "SUMSQ(LOOCV_Residual(X_s, Y, allow_arg, filt_arg))",
            functions["PRESS"].formula_display,
        )


if __name__ == "__main__":
    unittest.main()

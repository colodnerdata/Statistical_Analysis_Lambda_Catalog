"""Canonical document model and loader for lambda_functions.json."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

from lambda_catalog.lambda_formula_parser import to_workbook_xml_formula_from_display


_MAX_TEST_SHEET_NAME_LEN = 31
_INVALID_WORKSHEET_NAME_CHARS = set("[]:*?/\\")
_VALID_TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_MAX_EXCEL_ROW = 1_048_576


@dataclass(frozen=True)
class CatalogArgument:
    """A single LAMBDA function argument with its display metadata.

    Attributes
    ----------
    name : str
        The argument identifier as used in the LAMBDA formula.
    description : str
        Human-readable description of what the argument represents.
    optional : bool
        If True, the argument is optional and is displayed in brackets.
    """

    name: str
    description: str
    optional: bool = False

    def display_name(self) -> str:
        """Return the argument name formatted for display.

        Returns
        -------
        str
            The name wrapped in square brackets when optional, otherwise the bare name.
        """
        return f"[{self.name}]" if self.optional else self.name


@dataclass(frozen=True)
class CatalogFunction:
    """One LAMBDA function entry from the catalog.

    This is the single runtime object used by all consumers: workbook name sync,
    the catalog worksheet writer, and QC/test-sheet generation.

    Attributes
    ----------
    name : str
        The Excel defined name for this LAMBDA.
    formula_display : str
        Human-readable LAMBDA formula (may include whitespace and newlines).
    arguments : tuple[CatalogArgument, ...]
        Ordered arguments accepted by the function.
    yields : str
        Short description of the value the function returns.
    description : str
        Full description shown in the catalog.
    plain_language_summary : str
        Short plain-English explanation shown in the catalog column.
    test_table : str or None
        Tag identifying which test sheet this function belongs to.
    number_format : str
        Excel number format string for test result cells.
    """

    name: str
    formula_display: str
    arguments: tuple[CatalogArgument, ...]
    yields: str
    description: str
    plain_language_summary: str
    test_table: str | None
    number_format: str

    @property
    def argument_names(self) -> tuple[str, ...]:
        """Ordered tuple of bare argument names, without optional brackets."""
        return tuple(arg.name for arg in self.arguments)

    @property
    def name_manager_comment(self) -> str:
        """Argument documentation formatted for the Name Manager comment field.

        Each argument appears as ``display_name: description`` on its own block,
        separated by double newlines.

        Returns
        -------
        str
            The formatted comment string, or an empty string when there are no arguments.
        """
        lines = [
            f"{arg.display_name()}: {arg.description}"
            for arg in self.arguments
        ]
        return "\n\n".join(lines)

    def arguments_cell_text(self) -> str:
        """Format all arguments as multi-line text for the catalog Arguments cell.

        Each argument appears as ``display_name: description`` on its own line,
        separated by single newlines.

        Returns
        -------
        str
            The formatted text, or an empty string when the function takes no arguments.
        """
        if not self.arguments:
            return ""
        lines = [
            f"{arg.display_name()}: {arg.description}"
            for arg in self.arguments
        ]
        return "\n".join(lines)

    @property
    def workbook_xml_formula_from_display(self) -> str:
        """Display formula translated into workbook.xml token syntax.

        Returns
        -------
        str
            The formula string with platform-specific XML prefixes applied.
        """
        return to_workbook_xml_formula_from_display(self.formula_display)


@dataclass(frozen=True)
class CatalogDocument:
    """The full parsed contents of lambda_functions.json.

    Attributes
    ----------
    functions : tuple[CatalogFunction, ...]
        Ordered sequence of all function definitions.
    regression_sheet_notes : dict[str, str]
        Mapping of sheet label to plain-language note text.
    """

    functions: tuple[CatalogFunction, ...]
    regression_sheet_notes: dict[str, str]


def _looks_like_a1_reference(value: str) -> bool:
    """Return True when a candidate name resembles an Excel A1-style cell reference."""
    match = re.fullmatch(r"([A-Za-z]{1,3})([1-9][0-9]{0,6})", value)
    if match is None:
        return False
    column_label = match.group(1).upper()
    row_number = int(match.group(2))
    column_number = 0
    for char in column_label:
        column_number = (column_number * 26) + (ord(char) - ord("A") + 1)
    return column_number <= 16384 and row_number <= _MAX_EXCEL_ROW


def _validate_test_table_tag(tag: str, entry_index: int) -> None:
    """Validate a catalog test-table tag for worksheet and table-name usage.

    Parameters
    ----------
    tag : str
        The test_table string to validate.
    entry_index : int
        1-based index of the catalog entry, used in error messages.

    Raises
    ------
    ValueError
        If the tag exceeds the maximum worksheet name length, contains invalid
        characters, starts or ends with an apostrophe, fails the table-name
        regex, or resembles a cell reference.
    """
    if len(tag) > _MAX_TEST_SHEET_NAME_LEN:
        raise ValueError(
            f"'test_table' value {tag!r} in entry {entry_index} is "
            f"{len(tag)} characters; worksheet names may be at most "
            f"{_MAX_TEST_SHEET_NAME_LEN} characters."
        )
    if any(char in _INVALID_WORKSHEET_NAME_CHARS for char in tag):
        invalid_chars = "".join(
            sorted({c for c in tag if c in _INVALID_WORKSHEET_NAME_CHARS})
        )
        raise ValueError(
            f"'test_table' value {tag!r} in entry {entry_index} contains invalid "
            f"worksheet characters: {invalid_chars!r}."
        )
    if tag[0] == "'" or tag[-1] == "'":
        raise ValueError(
            f"'test_table' value {tag!r} in entry {entry_index} cannot begin or end "
            "with an apostrophe in Excel worksheet names."
        )
    if not _VALID_TABLE_NAME_RE.fullmatch(tag):
        raise ValueError(
            f"'test_table' value {tag!r} in entry {entry_index} must be a valid Excel "
            "table name: start with a letter or underscore and contain only letters, "
            "numbers, and underscores."
        )
    if _looks_like_a1_reference(tag):
        raise ValueError(
            f"'test_table' value {tag!r} in entry {entry_index} cannot look like an "
            "Excel cell reference when used as a table name."
        )


def load_catalog_document(
    path: Path, *, payload: dict | None = None
) -> CatalogDocument:
    """Load, validate, and parse lambda_functions.json into a CatalogDocument.

    Parameters
    ----------
    path : Path
        Path to the JSON file. Only read when ``payload`` is None.
    payload : dict or None, optional
        Pre-parsed JSON payload. When supplied, ``path`` is not re-read,
        avoiding redundant I/O when the caller already has the data in memory.

    Returns
    -------
    CatalogDocument
        The fully validated catalog document.

    Raises
    ------
    ValueError
        If the JSON structure is invalid, required fields are missing or blank,
        function names are duplicated, any test_table tag fails validation, or
        regression_sheet_notes is not a string-to-string mapping.
    """
    if payload is None:
        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)

    if not isinstance(payload, dict):
        raise ValueError("lambda_functions.json must be a JSON object at the top level.")

    functions_raw = payload.get("functions")
    if not isinstance(functions_raw, list):
        raise ValueError("lambda_functions.json must contain a top-level 'functions' array.")

    functions: list[CatalogFunction] = []
    seen_names: set[str] = set()

    for index, item in enumerate(functions_raw, start=1):
        if not isinstance(item, dict):
            raise ValueError(f"Entry {index} in lambda_functions.json must be an object.")

        raw_name = item.get("name")
        raw_formula_display = item.get("formula_display")

        if not isinstance(raw_name, str) or not raw_name.strip():
            raise ValueError(f"Entry {index} is missing a non-empty 'name'.")
        if not isinstance(raw_formula_display, str) or not raw_formula_display.strip():
            raise ValueError(f"Entry {index} is missing a non-empty 'formula_display'.")

        name = raw_name.strip()
        formula_display = raw_formula_display.strip()

        if name in seen_names:
            raise ValueError(f"Duplicate function name in lambda_functions.json: {name!r}.")

        raw_yields = item.get("yields")
        raw_description = item.get("description")
        if not isinstance(raw_yields, str) or not raw_yields.strip():
            raise ValueError(f"Entry {index} ({name!r}) is missing a non-empty 'yields'.")
        if not isinstance(raw_description, str) or not raw_description.strip():
            raise ValueError(f"Entry {index} ({name!r}) is missing a non-empty 'description'.")

        raw_pls = item.get("plain_language_summary")
        plain_language_summary = str(raw_pls).strip() if raw_pls is not None else ""
        if not plain_language_summary:
            raise ValueError(
                f"Function {name!r} is missing a non-empty "
                "'plain_language_summary' in lambda_functions.json."
            )

        raw_arguments = item.get("arguments", [])
        if not isinstance(raw_arguments, list):
            raise ValueError(f"Entry {index} ({name!r}) 'arguments' must be an array.")

        arguments: list[CatalogArgument] = []
        for arg_index, arg in enumerate(raw_arguments, start=1):
            if not isinstance(arg, dict):
                raise ValueError(
                    f"Entry {index} ({name!r}) argument {arg_index} must be an object."
                )
            raw_arg_name = arg.get("name")
            raw_arg_desc = arg.get("description")
            if not isinstance(raw_arg_name, str) or not raw_arg_name.strip():
                raise ValueError(
                    f"Entry {index} ({name!r}) argument {arg_index} is missing a "
                    "non-empty 'name'."
                )
            if not isinstance(raw_arg_desc, str) or not raw_arg_desc.strip():
                raise ValueError(
                    f"Entry {index} ({name!r}) argument {arg_index} "
                    f"({raw_arg_name.strip()!r}) is missing a non-empty 'description'."
                )
            arguments.append(CatalogArgument(
                name=raw_arg_name.strip(),
                description=raw_arg_desc.strip(),
                optional=bool(arg.get("optional", False)),
            ))

        raw_test_table = item.get("test_table")
        test_table: str | None = None
        if raw_test_table is not None:
            if not isinstance(raw_test_table, str) or not raw_test_table.strip():
                raise ValueError(f"Entry {index} 'test_table' must be a non-empty string.")
            test_table = raw_test_table.strip()
            _validate_test_table_tag(test_table, index)

        raw_number_format = item.get("number_format", "General")
        number_format = str(raw_number_format).strip() if raw_number_format else "General"

        seen_names.add(name)
        functions.append(CatalogFunction(
            name=name,
            formula_display=formula_display,
            arguments=tuple(arguments),
            yields=raw_yields.strip(),
            description=raw_description.strip(),
            plain_language_summary=plain_language_summary,
            test_table=test_table,
            number_format=number_format,
        ))

    notes_raw = payload.get("regression_sheet_notes", {})
    if not isinstance(notes_raw, dict):
        raise ValueError(
            "regression_sheet_notes in lambda_functions.json must be an object."
        )
    for key, value in notes_raw.items():
        if not isinstance(key, str) or not isinstance(value, str):
            raise ValueError(
                f"regression_sheet_notes entry {key!r} must have string key and string value; "
                f"got {type(value).__name__}."
            )

    return CatalogDocument(
        functions=tuple(functions),
        regression_sheet_notes=dict(notes_raw),
    )

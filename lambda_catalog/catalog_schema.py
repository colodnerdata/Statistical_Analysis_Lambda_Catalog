"""Canonical document model and loader for lambda_functions.json."""
from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from lambda_catalog.lambda_formula_parser import to_workbook_xml_formula_from_display


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
    notes : str
        Tooltip text (max 255 characters) written to the workbook's defined-name
        ``comment`` attribute, shown by Excel in Name Manager and formula-bar
        autocomplete.
    scope : str
        Where the name lives. ``"workbook"`` (the default) means a portable,
        workbook-scoped defined name synced into workbook.xml. Any other value
        is a worksheet name: the function is a sheet-scoped closure installed on
        that sheet (e.g. the Model Construction constructor names), excluded from
        the workbook-scope sync and labelled by its owning sheet in the catalog.
    """

    name: str
    formula_display: str
    arguments: tuple[CatalogArgument, ...]
    yields: str
    description: str
    plain_language_summary: str
    notes: str
    scope: str = "workbook"

    @property
    def argument_names(self) -> tuple[str, ...]:
        """Ordered tuple of bare argument names, without optional brackets."""
        return tuple(arg.name for arg in self.arguments)

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
    univariate_sheet_notes : dict[str, str]
        Mapping of Univariate sheet label to plain-language note text.
    """

    functions: tuple[CatalogFunction, ...]
    regression_sheet_notes: dict[str, str]
    univariate_sheet_notes: dict[str, str]

    @property
    def workbook_functions(self) -> tuple[CatalogFunction, ...]:
        """Functions synced as portable workbook-scoped defined names."""
        return tuple(fn for fn in self.functions if fn.scope == "workbook")

    def functions_for_sheet(self, sheet_name: str) -> tuple[CatalogFunction, ...]:
        """Sheet-scoped closures owned by ``sheet_name``, in document order.

        Document order is dependency order for the installer (Excel resolves
        each name against the ones already added), so callers install these
        exactly as returned.
        """
        return tuple(fn for fn in self.functions if fn.scope == sheet_name)


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
        function names are duplicated, or regression_sheet_notes /
        univariate_sheet_notes are not string-to-string mappings.
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

        raw_notes = item.get("notes", "")
        notes = str(raw_notes).strip() if raw_notes else ""
        if notes and len(notes) > 255:
            raise ValueError(
                f"Entry {index} ({name!r}) 'notes' is {len(notes)} characters; "
                "limit is 255."
            )

        raw_scope = item.get("scope", "workbook")
        if not isinstance(raw_scope, str) or not raw_scope.strip():
            raise ValueError(
                f"Entry {index} ({name!r}) 'scope' must be a non-empty string "
                "('workbook' or a worksheet name)."
            )
        scope = raw_scope.strip()

        seen_names.add(name)
        functions.append(CatalogFunction(
            name=name,
            formula_display=formula_display,
            arguments=tuple(arguments),
            yields=raw_yields.strip(),
            description=raw_description.strip(),
            plain_language_summary=plain_language_summary,
            notes=notes,
            scope=scope,
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

    uv_notes_raw = payload.get("univariate_sheet_notes", {})
    if not isinstance(uv_notes_raw, dict):
        raise ValueError(
            "univariate_sheet_notes in lambda_functions.json must be an object."
        )
    for key, value in uv_notes_raw.items():
        if not isinstance(key, str) or not isinstance(value, str):
            raise ValueError(
                f"univariate_sheet_notes entry {key!r} must have string key and string value; "
                f"got {type(value).__name__}."
            )

    return CatalogDocument(
        functions=tuple(functions),
        regression_sheet_notes=dict(notes_raw),
        univariate_sheet_notes=dict(uv_notes_raw),
    )


DEFAULT_CATALOG_PATH = Path(__file__).resolve().parent.parent / "lambda_functions.json"


@lru_cache(maxsize=1)
def _default_document() -> CatalogDocument:
    return load_catalog_document(DEFAULT_CATALOG_PATH)


def catalog_argument_names(function_name: str) -> tuple[str, ...]:
    """Return the declared argument names of ``function_name`` from the catalog.

    The QC test-sheet writers render their formulas from these names rather than
    from a hard-coded positional list, so a signature change in the catalog can
    never leave a test sheet calling the old shape. The parsed document is
    cached, so repeated lookups during a build cost one file read.

    Parameters
    ----------
    function_name : str
        The catalog function to look up.

    Returns
    -------
    tuple[str, ...]
        Argument names in declaration order.

    Raises
    ------
    KeyError
        If no catalog function has that name.
    """
    for function in _default_document().functions:
        if function.name == function_name:
            return function.argument_names
    raise KeyError(f"{function_name!r} is not in the catalog")

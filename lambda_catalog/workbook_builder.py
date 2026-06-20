"""Shared LAMBDA definitions, name-syncing utilities, and workbook helpers."""
from __future__ import annotations

import json
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

import lxml.etree as etree  # type: ignore[import-untyped]  # pyright: ignore[reportMissingTypeStubs]
import xlwings as xw

from lambda_catalog.lambda_formula_parser import to_workbook_xml_formula_from_display
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, excel_error_message


WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CALC_CHAIN_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
_MAX_TEST_SHEET_NAME_LEN = 31
_INVALID_WORKSHEET_NAME_CHARS = set("[]:*?/\\")
_VALID_TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_MAX_EXCEL_ROW = 1_048_576
XL_CALCULATION_AUTOMATIC     = -4105  # Excel XlCalculation.xlCalculationAutomatic
XL_CALCULATION_MANUAL        = -4135  # Excel XlCalculation.xlCalculationManual
XL_CALCULATION_SEMIAUTOMATIC = -4133  # Excel XlCalculation.xlCalculationSemiAutomatic (auto except tables)


@dataclass(frozen=True)
class LambdaDefinition:
    """One workbook-scoped LAMBDA definition loaded from the JSON catalog.

    Attributes
    ----------
    name : str
        The Excel defined name for this LAMBDA.
    formula_display : str
        Human-readable LAMBDA formula (may include whitespace and newlines).
    comment : str
        Argument documentation formatted for the Name Manager comment field.
    argument_names : tuple[str, ...]
        Ordered argument names extracted from the argument list.
    test_table : str or None
        Tag identifying which test sheet this function belongs to.
    number_format : str
        Excel number format string for test result cells.
    """

    name: str
    formula_display: str
    comment: str = ""
    argument_names: tuple[str, ...] = ()
    test_table: str | None = None
    number_format: str = "General"

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
class NameSyncResult:
    """Summary of how many workbook names were created versus replaced.

    Attributes
    ----------
    created : int
        Number of new workbook names added.
    updated : int
        Number of existing workbook names replaced.
    """

    created: int
    updated: int


def _looks_like_a1_reference(value: str) -> bool:
    """Return True when a candidate name resembles an Excel A1-style reference.

    Parameters
    ----------
    value : str
        The string to test.

    Returns
    -------
    bool
        True if the string matches an Excel A1-style cell reference pattern
        within the valid column (≤ XFD) and row (≤ 1,048,576) ranges.
    """
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
        If the tag exceeds the maximum worksheet name length, contains
        invalid characters, starts or ends with an apostrophe, fails the
        table-name regex, or resembles a cell reference.
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


def load_lambda_definitions(
    path: Path, *, payload: dict | None = None
) -> list[LambdaDefinition]:
    """Load, validate, and normalize the JSON catalog into LambdaDefinition objects.

    Parameters
    ----------
    path : Path
        Path to the JSON file containing the lambda definitions.
    payload : dict or None, optional
        Pre-parsed JSON payload. When supplied the file at ``path`` is not
        re-read, avoiding redundant I/O when the caller already has the data.

    Returns
    -------
    list[LambdaDefinition]
        Ordered list of validated definitions.

    Raises
    ------
    ValueError
        If the JSON structure is invalid, required fields are missing,
        names are duplicated, or any test_table tag fails validation.
    """
    if payload is None:
        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)

    functions = payload.get("functions")
    if not isinstance(functions, list):
        raise ValueError("lambda_functions.json must contain a top-level 'functions' array.")

    definitions: list[LambdaDefinition] = []
    seen_names: set[str] = set()

    for index, item in enumerate(functions, start=1):
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
            raise ValueError(f"Duplicate function name in lambda_functions.json: {name}")

        raw_arguments = item.get("arguments", [])
        comment_lines = []
        arg_names: list[str] = []
        for arg in raw_arguments:
            arg_name = str(arg.get("name", "")).strip()
            arg_desc = str(arg.get("description", "")).strip()
            display = f"[{arg_name}]" if arg.get("optional") else arg_name
            comment_lines.append(f"{display}: {arg_desc}")
            arg_names.append(arg_name)
        comment = "\n\n".join(comment_lines)

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
        definitions.append(
            LambdaDefinition(
                name=name,
                formula_display=formula_display,
                comment=comment,
                argument_names=tuple(arg_names),
                test_table=test_table,
                number_format=number_format,
            )
        )

    return definitions


def _defined_names_insert_index(workbook_root) -> int:
    """Return the schema-safe insertion point for a definedNames element.

    Parameters
    ----------
    workbook_root : lxml.etree._Element
        The root element of workbook.xml.

    Returns
    -------
    int
        Index at which to insert the definedNames child element.
    """
    ordered_followers = {
        "calcPr",
        "oleSize",
        "customWorkbookViews",
        "pivotCaches",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    }

    children = list(workbook_root)
    for index, child in enumerate(children):
        tag_name = child.tag.rsplit("}", 1)[-1]
        if tag_name in ordered_followers:
            return index
    return len(children)


def sync_workbook_names(
    workbook_path: Path, definitions: list[LambdaDefinition]
) -> NameSyncResult:
    """Replace workbook-scoped names in workbook.xml with the supplied definitions.

    Parameters
    ----------
    workbook_path : Path
        Path to the .xlsx file to patch in place.
    definitions : list[LambdaDefinition]
        Definitions whose names will be inserted or replaced.

    Returns
    -------
    NameSyncResult
        Counts of created versus updated workbook names.
    """
    target_names = {definition.name for definition in definitions}
    existing_target_names: set[str] = set()
    temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=workbook_path.parent)
    os.close(temp_fd)
    Path(temp_name).unlink(missing_ok=True)

    try:
        with zipfile.ZipFile(workbook_path, "r") as input_zip, zipfile.ZipFile(
            temp_name, "w", compression=zipfile.ZIP_DEFLATED
        ) as output_zip:
            for item in input_zip.infolist():
                data = input_zip.read(item.filename)

                if item.filename == "xl/workbook.xml":
                    workbook_root = etree.fromstring(data)
                    defined_names_elements = workbook_root.findall(
                        f"{{{WORKBOOK_NS}}}definedNames"
                    )

                    if defined_names_elements:
                        defined_names = defined_names_elements[0]
                        for extra in defined_names_elements[1:]:
                            workbook_root.remove(extra)
                    else:
                        defined_names = etree.Element(f"{{{WORKBOOK_NS}}}definedNames")
                        workbook_root.insert(
                            _defined_names_insert_index(workbook_root), defined_names
                        )

                    for name_element in list(defined_names):
                        if name_element.tag != f"{{{WORKBOOK_NS}}}definedName":
                            continue

                        target_name = name_element.get("name")
                        is_workbook_scope = name_element.get("localSheetId") is None
                        if is_workbook_scope and target_name in target_names:
                            existing_target_names.add(target_name)
                            defined_names.remove(name_element)

                    for definition in definitions:
                        name_element = etree.SubElement(
                            defined_names, f"{{{WORKBOOK_NS}}}definedName"
                        )
                        name_element.set("name", definition.name)
                        name_element.text = definition.workbook_xml_formula_from_display

                    xml_body = str(etree.tostring(workbook_root, encoding="unicode"))
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        + xml_body
                    ).encode("UTF-8")

                elif item.filename == "xl/calcChain.xml":
                    continue  # omit stale calc chain; Excel rebuilds on open

                elif item.filename == "xl/_rels/workbook.xml.rels":
                    rels_root = etree.fromstring(data)
                    for rel in rels_root.findall(f"{{{_RELS_NS}}}Relationship"):
                        if rel.get("Type") == _CALC_CHAIN_REL_TYPE:
                            rels_root.remove(rel)
                    xml_body = str(etree.tostring(rels_root, encoding="unicode"))
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        + xml_body
                    ).encode("UTF-8")

                elif item.filename == "[Content_Types].xml":
                    ct_root = etree.fromstring(data)
                    for override in ct_root.findall(f"{{{_CT_NS}}}Override"):
                        part_name = override.get("PartName") or ""
                        if part_name.lower() == "/xl/calcchain.xml":
                            ct_root.remove(override)
                    xml_body = str(etree.tostring(ct_root, encoding="unicode"))
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        + xml_body
                    ).encode("UTF-8")

                output_zip.writestr(item, data)

        Path(temp_name).replace(workbook_path)
    finally:
        Path(temp_name).unlink(missing_ok=True)

    return NameSyncResult(
        created=len(definitions) - len(existing_target_names),
        updated=len(existing_target_names),
    )


def _validate_workbook_reopen(workbook_path: Path) -> None:
    """Reopen the saved workbook in Excel to confirm the patched package is valid.

    Parameters
    ----------
    workbook_path : Path
        Path to the .xlsx file to reopen and immediately close.

    Raises
    ------
    RuntimeError
        If Excel rejects the workbook during open.
    """
    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook = app.books.open(str(workbook_path))
            workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise RuntimeError(
            f"Workbook reopen validation failed for {workbook_path.name!r}: "
            f"{excel_error_message(exc)}"
        ) from exc


def _delete_sheet_if_present(workbook: xw.Book, sheet_name: str) -> None:
    """Delete a worksheet by name if it exists in the workbook.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook.
    sheet_name : str
        Name of the sheet to delete.
    """
    for sheet in list(workbook.sheets):
        if sheet.name == sheet_name:
            sheet.delete()
            return

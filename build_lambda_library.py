from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from pathlib import Path
import tempfile
import zipfile

from lxml import etree
from lambda_formula_parser import to_workbook_xml_formula, to_workbook_xml_formula_from_display
import pywintypes  # type: ignore[import-untyped]
import xlwings as xw


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
STARTER_SHEET_NAME = "MLR"
WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)


@dataclass(frozen=True)
class LambdaDefinition:
    """Represents one workbook-scoped LAMBDA definition from the JSON catalog."""

    name: str
    formula_compact: str
    formula_display: str

    @property
    def workbook_xml_formula(self) -> str:
        """Return the formula translated into workbook.xml token syntax."""

        return to_workbook_xml_formula(self.formula_compact)

    @property
    def workbook_xml_formula_from_display(self) -> str:
        """Return the display formula translated into workbook.xml token syntax."""

        return to_workbook_xml_formula_from_display(self.formula_display)


@dataclass(frozen=True)
class NameSyncResult:
    """Summarizes how many workbook names were created versus replaced."""

    created: int
    updated: int


def load_lambda_definitions(path: Path) -> list[LambdaDefinition]:
    """Load, validate, and normalize the JSON catalog into LambdaDefinition objects."""

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
        raw_formula = item.get("formula_compact")
        raw_formula_display = item.get("formula_display")

        if not isinstance(raw_name, str) or not raw_name.strip():
            raise ValueError(f"Entry {index} is missing a non-empty 'name'.")
        if not isinstance(raw_formula, str) or not raw_formula.strip():
            raise ValueError(f"Entry {index} is missing a non-empty 'formula_compact'.")
        if not isinstance(raw_formula_display, str) or not raw_formula_display.strip():
            raise ValueError(f"Entry {index} is missing a non-empty 'formula_display'.")

        name = raw_name.strip()
        formula_compact = raw_formula.strip()
        formula_display = raw_formula_display.strip()

        if name in seen_names:
            raise ValueError(f"Duplicate function name in lambda_functions.json: {name}")

        seen_names.add(name)
        definitions.append(
            LambdaDefinition(
                name=name,
                formula_compact=formula_compact,
                formula_display=formula_display,
            )
        )

    return definitions


def ensure_starter_sheet(workbook: xw.Book, definition_count: int) -> None:
    """Ensure the workbook has the expected starter sheet and summary text."""

    sheet_names = {sheet.name for sheet in workbook.sheets}
    if STARTER_SHEET_NAME in sheet_names:
        sheet = workbook.sheets[STARTER_SHEET_NAME]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = STARTER_SHEET_NAME
    else:
        sheet = workbook.sheets.add(name=STARTER_SHEET_NAME, after=workbook.sheets[-1])

    sheet["A1"].value = "Lambda Library"
    sheet["A2"].value = (
        "Workbook-scoped LAMBDA functions are synced from lambda_functions.json. "
        "Open Formulas > Name Manager in Excel to inspect or use them."
    )
    sheet["A3"].value = f"Registered definitions: {definition_count}"
    sheet.range("A1:A3").columns.autofit()


def sync_workbook_names(workbook_path: Path, definitions: list[LambdaDefinition]) -> NameSyncResult:
    """Replace target workbook-scoped names in workbook.xml with the supplied definitions."""

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
                    defined_names_elements = workbook_root.findall(f"{{{WORKBOOK_NS}}}definedNames")

                    if defined_names_elements:
                        defined_names = defined_names_elements[0]
                        for extra in defined_names_elements[1:]:
                            workbook_root.remove(extra)
                    else:
                        defined_names = etree.Element(f"{{{WORKBOOK_NS}}}definedNames")
                        workbook_root.insert(_defined_names_insert_index(workbook_root), defined_names)

                    for name_element in list(defined_names):
                        if name_element.tag != f"{{{WORKBOOK_NS}}}definedName":
                            continue

                        target_name = name_element.get("name")
                        is_workbook_scope = name_element.get("localSheetId") is None
                        if is_workbook_scope and target_name in target_names:
                            existing_target_names.add(target_name)
                            defined_names.remove(name_element)

                    for definition in definitions:
                        name_element = etree.SubElement(defined_names, f"{{{WORKBOOK_NS}}}definedName")
                        name_element.set("name", definition.name)
                        name_element.text = definition.workbook_xml_formula_from_display

                    data = etree.tostring(
                        workbook_root,
                        encoding="UTF-8",
                        xml_declaration=True,
                        standalone=True,
                    )

                output_zip.writestr(item, data)

        Path(temp_name).replace(workbook_path)
    finally:
        Path(temp_name).unlink(missing_ok=True)

    return NameSyncResult(created=len(definitions) - len(existing_target_names), updated=len(existing_target_names))


def _defined_names_insert_index(workbook_root) -> int:
    """Return the schema-safe insertion point for a definedNames element."""

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


def _excel_error_message(exc: BaseException) -> str:
    """Extract a readable message from an Excel automation exception."""

    return str(exc).strip() or exc.__class__.__name__


def _raise_excel_access_error(workbook_path: Path, action: str, exc: BaseException) -> None:
    """Raise a clearer workbook access error when Excel has the file locked or read-only."""

    message = _excel_error_message(exc)
    normalized = message.lower()
    likely_locked = any(
        phrase in normalized
        for phrase in (
            "cannot access read-only document",
            "read-only",
            "currently in use",
            "sharing violation",
            "permission denied",
        )
    )

    if likely_locked:
        raise RuntimeError(
            f"Excel could not {action} {workbook_path.name!r}. "
            "The workbook is likely open in Excel or locked by another process. "
            f"Close Excel and retry. Original error: {message}"
        ) from exc

    raise RuntimeError(f"Excel could not {action} {workbook_path.name!r}: {message}") from exc


def _validate_workbook_reopen(workbook_path: Path) -> None:
    """Reopen the saved workbook in Excel to confirm the patched package is still valid."""

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook = app.books.open(str(workbook_path))
            workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise RuntimeError(
            f"Workbook reopen validation failed for {workbook_path.name!r}: {_excel_error_message(exc)}"
        ) from exc


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for workbook generation and optional validation."""

    parser = argparse.ArgumentParser(description="Build Lambda_Library.xlsx from lambda_functions.json.")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK_PATH,
        help="Path to the workbook to create or update.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda definitions.",
    )
    parser.add_argument(
        "--validate-reopen",
        action="store_true",
        help="Reopen the workbook after syncing names to verify Excel accepts the result.",
    )
    return parser.parse_args()


def build_lambda_library(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    validate_reopen: bool = False,
) -> NameSyncResult:
    """Open or create the workbook, sync all JSON-backed names, and optionally validate reopen."""

    definitions = load_lambda_definitions(definitions_path)
    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    try:
        with xw.App(visible=False, add_book=False) as app:
            if workbook_exists:
                workbook = app.books.open(str(workbook_path))
            else:
                workbook = app.books.add()

            try:
                ensure_starter_sheet(workbook, len(definitions))
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        _raise_excel_access_error(workbook_path, "open or save", exc)

    result = sync_workbook_names(workbook_path, definitions)
    if validate_reopen:
        _validate_workbook_reopen(workbook_path)

    return result


def main() -> None:
    """Build the workbook and print a short sync summary for interactive use."""

    args = parse_args()
    result = build_lambda_library(
        workbook_path=args.workbook,
        definitions_path=args.definitions,
        validate_reopen=args.validate_reopen,
    )
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    print("Invalid entries: 0")
    if args.validate_reopen:
        print("Reopen validation: passed")


if __name__ == "__main__":
    main()
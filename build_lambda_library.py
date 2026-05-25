"""Build Lambda_Library.xlsx from lambda_functions.json."""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from lambda_catalog.lambda_formula_parser import to_workbook_xml_formula_from_display
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, excel_error_message, raise_excel_access_error
from lambda_catalog.write_sheet_lambda_functions import load_catalog_entries, write_catalog_sheet
from lambda_catalog.write_sheet_life_expectancy_data import (
    DEFAULT_CSV_PATH,
    load_life_expectancy_rows,
    write_life_expectancy_sheet,
)
from lambda_catalog.analysis_cache import get_analysis_results
from lambda_catalog.write_sheet_mlr_scalar_test import write_mlr_scalar_test_sheet
from lambda_catalog.write_sheet_mlr_observation_test import write_mlr_observation_test_sheet
from lambda_catalog.write_sheet_mlr_vector_outputs_test import write_mlr_vector_outputs_test_sheet
from lambda_catalog.write_sheet_regression import write_regression_output_sheet
import xlwings as xw


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_MAX_TEST_SHEET_NAME_LEN = 31
_INVALID_WORKSHEET_NAME_CHARS = set("[]:*?/\\")
_VALID_TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_MAX_EXCEL_ROW = 1_048_576
XL_CALCULATION_AUTOMATIC = -4105  # Excel XlCalculation.xlCalculationAutomatic


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


def load_lambda_definitions(path: Path) -> list[LambdaDefinition]:
    """Load, validate, and normalize the JSON catalog into LambdaDefinition objects.

    Parameters
    ----------
    path : Path
        Path to the JSON file containing the lambda definitions.

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
                        if definition.comment:
                            name_element.set("comment", definition.comment)
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

    return NameSyncResult(
        created=len(definitions) - len(existing_target_names),
        updated=len(existing_target_names),
    )


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


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for workbook generation and validation.

    Returns
    -------
    argparse.Namespace
        Parsed arguments with workbook, definitions, csv, and
        validate_reopen attributes.
    """
    parser = argparse.ArgumentParser(
        description="Build Lambda_Library.xlsx from lambda_functions.json."
    )
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
        "--csv",
        type=Path,
        default=DEFAULT_CSV_PATH,
        help="Path to the Life Expectancy CSV data file.",
    )
    parser.add_argument(
        "--validate-reopen",
        action="store_true",
        help="Reopen the workbook after syncing names to verify Excel accepts the result.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print timing information for each build phase.",
    )
    return parser.parse_args()


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



def verify_test_sheets(
    workbook: xw.Book,
    scalar_row_configs: list,
    vector_row_configs: list,
    observation_row_configs: list,
) -> None:
    """Compare Excel Calc columns against Python-computed expected values in both MLR test sheets.

    Forces a full recalculation on the open workbook, then reads each sheet's
    Calc columns via inspect_test_sheets, compares them against expected values
    derived from the supplied analysis configs (not from the sheet's (Exp.) columns),
    and warns (without raising) for every value that diverges within the tolerance band.

    Parameters
    ----------
    workbook : xw.Book
        Open xlwings Book to inspect (must already be saved so formulas are
        linked to the workbook-scoped defined names).
    scalar_row_configs : list
        Per-row configs from ``build_mlr_row_configs``; provides the Python-computed
        expected scalar values used for comparison.
    vector_row_configs : list
        Per-section configs from ``build_mlr_vector_row_configs``; provides the
        Python-computed expected coefficient vectors used for comparison.
    """
    import importlib.util

    tool_path = ROOT_DIR / "tools" / "inspect_test_sheets.py"
    spec = importlib.util.spec_from_file_location("inspect_test_sheets", tool_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load inspect_test_sheets from {tool_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    workbook.app.calculate()

    scalar_df = mod.read_scalar_df(workbook, scalar_row_configs)
    vector_df = mod.read_vector_df(workbook, vector_row_configs)
    observation_df = mod.read_observation_df(workbook, observation_row_configs)
    tol = mod.TOLERANCE_DECIMALS

    for _, row in scalar_df.iterrows():
        fdd = row["first_digit_deviation"]
        if fdd is not None and fdd <= tol:
            k = row["k"]
            ai = row["allow_intercept"]
            stat = row["stat_name"]
            print(
                f"WARNING [MLR_Scalar_Test] k={k} intercept={'TRUE' if ai else 'FALSE'} "
                f"stat={stat!r}: Calc vs Python-expected mismatch — first digit of deviation at "
                f"decimal place {fdd} (tolerance={tol}). "
                f"expected={row['expected']!r}, excel_calc={row['excel_calc']!r}, "
                f"abs_diff={row['abs_diff']!r}",
                flush=True,
            )

    for _, row in vector_df.iterrows():
        fdd = row["first_digit_deviation"]
        if fdd is not None and fdd <= tol:
            k = row["k"]
            ai = row["allow_intercept"]
            stat = row["stat_name"]
            term = row["term_name"]
            print(
                f"WARNING [MLR_Vector_Outputs_Test] k={k} intercept={'TRUE' if ai else 'FALSE'} "
                f"stat={stat!r} term={term!r}: Calc vs Python-expected mismatch — first digit of deviation "
                f"at decimal place {fdd} (tolerance={tol}). "
                f"expected={row['expected']!r}, excel_calc={row['excel_calc']!r}, "
                f"abs_diff={row['abs_diff']!r}",
                flush=True,
            )
    for _, row in observation_df.iterrows():
        fdd = row["first_digit_deviation"]
        if fdd is not None and fdd <= tol:
            print(
                f"WARNING [MLR_Observation_Test] k={row['k']} intercept={'TRUE' if row['allow_intercept'] else 'FALSE'} "
                f"stat={row['stat_name']!r}: expected={row['expected']!r}, excel_calc={row['excel_calc']!r}",
                flush=True,
            )


def build_lambda_library(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = DEFAULT_CSV_PATH,
    validate_reopen: bool = False,
    verbose: bool = False,
) -> NameSyncResult:
    """Build all workbook assets, sync JSON-backed names, and optionally validate.

    Parameters
    ----------
    workbook_path : Path, optional
        Path to the workbook to create or update.
    definitions_path : Path, optional
        Path to the JSON catalog file.
    csv_path : Path, optional
        Path to the Life Expectancy CSV file.
    validate_reopen : bool, optional
        If True, reopens the workbook in Excel after patching to verify it.
    verbose : bool, optional
        If True, prints timing information for each build phase to stdout.

    Returns
    -------
    NameSyncResult
        Counts of created versus updated workbook names.
    """
    _t = time.monotonic()
    definitions = load_lambda_definitions(definitions_path)
    catalog_entries = load_catalog_entries(definitions_path)
    row_configs, vector_row_configs, observation_row_configs = get_analysis_results(csv_path)
    csv_headers, csv_rows = load_life_expectancy_rows(csv_path)
    if verbose:
        print(f"  Prep:           {time.monotonic() - _t:.1f}s", flush=True)

    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    _t = time.monotonic()
    try:
        with xw.App(visible=True, add_book=False) as app:
            if workbook_exists:
                workbook = app.books.open(str(workbook_path))
            else:
                workbook = app.books.add()

            try:
                _delete_sheet_if_present(workbook, PREDICTIONS_SHEET_NAME)
                if "Sheet1" in {sheet.name for sheet in workbook.sheets}:
                    workbook.sheets["Sheet1"].name = "LAMBDA_functions"
                write_catalog_sheet(workbook, catalog_entries)
                write_life_expectancy_sheet(workbook, csv_headers, csv_rows)
                write_regression_output_sheet(workbook)
                write_mlr_scalar_test_sheet(
                    workbook,
                    definitions,
                    row_configs,
                )
                write_mlr_vector_outputs_test_sheet(
                    workbook,
                    vector_row_configs,
                )
                write_mlr_observation_test_sheet(workbook, observation_row_configs)
                app.api.Calculation = XL_CALCULATION_AUTOMATIC
                workbook.save(str(workbook_path))
                verify_test_sheets(workbook, row_configs, vector_row_configs, observation_row_configs)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)
    if verbose:
        print(f"  Write sheets:   {time.monotonic() - _t:.1f}s", flush=True)

    _t = time.monotonic()
    try:
        result = sync_workbook_names(workbook_path, definitions)
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    except (PermissionError, OSError) as exc:
        raise_excel_access_error(workbook_path, "update", exc)
    if verbose:
        print(f"  Sync names:     {time.monotonic() - _t:.1f}s", flush=True)

    if validate_reopen:
        _validate_workbook_reopen(workbook_path)

    return result


def main() -> None:
    """Build the workbook and print a short sync summary for interactive use."""
    args = parse_args()
    while True:
        try:
            result = build_lambda_library(
                workbook_path=args.workbook,
                definitions_path=args.definitions,
                csv_path=args.csv,
                validate_reopen=args.validate_reopen,
                verbose=args.verbose,
            )
            break
        except RuntimeError as exc:
            if "likely open in Excel" in str(exc):
                if not sys.stdin.isatty():
                    raise
                prompt = (
                    f"\n{args.workbook.name} is open in Excel — close it "
                    "and press Enter to retry (or Ctrl+C to cancel): "
                )
                input(prompt)
            else:
                raise

    print(f"Workbook: {args.workbook.resolve()}")
    print("Sheet updated: MLR_Scalar_Test")
    print("Sheet updated: MLR_Vector_Outputs_Test")
    print("Sheet updated: MLR_Observation_Test")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print("Sheet updated: Regression")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    print("Invalid entries: 0")
    if args.validate_reopen:
        print("Reopen validation: passed")

    subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])


if __name__ == "__main__":
    main()

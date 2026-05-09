from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from pathlib import Path
import tempfile
from typing import Any
from typing import Iterable
import zipfile

from lxml import etree
from analyze_life_expectancy import (
    RegressionSummary,
    calculate_regression_summary,
)
from lambda_formula_parser import to_workbook_xml_formula_from_display
import pywintypes  # type: ignore[import-untyped]
from write_sheet_lambda_functions import load_catalog_entries, write_catalog_sheet
from write_sheet_life_expectancy_data import (
    DEFAULT_CSV_PATH,
    load_life_expectancy_rows,
    write_life_expectancy_sheet,
)
import xlwings as xw


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
STARTER_SHEET_NAME = "MLR_Testing"
PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_SRC_RANGE = 1
XL_YES = 1
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)

_MLR_TABLE_HEADER_ROW = 4
_MLR_TABLE_NAME = "MLRTestingResults"

_SMOKE_TEST_FORMULA = (
    "=AND("
    "[@[Observations (Exp.)]]=[@[Observations (Calc.)]],"
    "[@[DF_Regression (Exp.)]]=[@[DF_Regression (Calc.)]],"
    "[@[DF_Total (Exp.)]]=[@[DF_Total (Calc.)]],"
    "ROUND([@[R_squared (Exp.)]],10)=ROUND([@[R_squared (Calc.)]],10),"
    "[@[DF_Residual (Exp.)]]=[@[DF_Residual (Calc.)]],"
    "ROUND([@[Multiple_R (Exp.)]],10)=ROUND([@[Multiple_R (Calc.)]],10),"
    "ROUND([@[Adjusted_R2 (Exp.)]],10)=ROUND([@[Adjusted_R2 (Calc.)]],10)"
    ")"
)

# (header, expected_key_or_None, formula_or_None, number_format)
_MLR_TABLE_COLUMNS: list[tuple[str, str | None, str | None, str]] = [
    ("X_Variables",           None,            '=TEXTJOIN(", ",TRUE,OFFSET(x_s,-1,0,1,COLUMNS(x_s)))', "General"),
    ("Allow_Intercept",       None,            None,                                                     "General"),
    ("Smoke Test",            None,            _SMOKE_TEST_FORMULA,                                     "General"),
    ("Observations (Exp.)",   "Observations",  None,                                                    "0"),
    ("Observations (Calc.)",  None,            "=Observations(y, fil)",                                 "0"),
    ("DF_Regression (Exp.)",  "DF_Regression", None,                                                    "0"),
    ("DF_Regression (Calc.)", None,            "=DF_Regression(x_s)",                                  "0"),
    ("DF_Total (Exp.)",       "DF_Total",      None,                                                    "0"),
    ("DF_Total (Calc.)",      None,            "=DF_Total(y, [@[Allow_Intercept]], fil)",                 "0"),
    ("R_squared (Exp.)",      "R_squared",     None,                                                    "0.000"),
    ("R_squared (Calc.)",     None,            "=R_squared(x_s, y, [@[Allow_Intercept]], fil)",           "0.000"),
    ("DF_Residual (Exp.)",    "DF_Residual",   None,                                                    "0"),
    ("DF_Residual (Calc.)",   None,            "=DF_Residual(x_s, y, [@[Allow_Intercept]], fil)",         "0"),
    ("Multiple_R (Exp.)",     "Multiple_R",    None,                                                    "0.000"),
    ("Multiple_R (Calc.)",    None,            "=Multiple_R(x_s, y, [@[Allow_Intercept]], fil)",          "0.000"),
    ("Adjusted_R2 (Exp.)",    "Adjusted_R2",   None,                                                    "0.000"),
    ("Adjusted_R2 (Calc.)",   None,            "=Adjusted_R2(x_s, y, [@[Allow_Intercept]], fil)",         "0.000"),
]


@dataclass(frozen=True)
class LambdaDefinition:
    """Represents one workbook-scoped LAMBDA definition from the JSON catalog."""

    name: str
    formula_display: str
    comment: str = ""

    @property
    def workbook_xml_formula_from_display(self) -> str:
        """Return the display formula translated into workbook.xml token syntax."""

        return to_workbook_xml_formula_from_display(self.formula_display)


@dataclass(frozen=True)
class NameSyncResult:
    """Summarizes how many workbook names were created versus replaced."""

    created: int
    updated: int


def _delete_sheet_scoped_name_if_present(sheet: xw.Sheet, target_name: str) -> None:
    """Delete a local worksheet name if it already exists."""

    for index in range(sheet.api.Names.Count, 0, -1):
        existing_name = sheet.api.Names(index).Name
        local_name = existing_name.split("!", 1)[-1]
        if local_name.lower() == target_name.lower():
            sheet.api.Names(index).Delete()


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    """Create worksheet-scoped helper names used by the MLR formula tests."""

    local_names = {
        "y": "=LifeExpectancyData[Life expectancy]",
        "x_s": "=LifeExpectancyData[[Adult Mortality]:[Schooling]]",
        "fil": "=LifeExpectancyData[Full_Data]",
    }

    for name, refers_to in local_names.items():
        _delete_sheet_scoped_name_if_present(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)

    _delete_sheet_scoped_name_if_present(sheet, "Allow_Intercept")


def _actual_formula(function_name: str, argument_names: Iterable[str]) -> str:
    """Build an executable test formula that references local helper names."""

    reference_map = {
        "y": "y",
        "x_s": "x_s",
        "filter": "fil",
        "allow_intercept": "Allow_Intercept",
    }

    resolved_arguments = [
        reference_map.get(argument_name.strip().lower(), argument_name)
        for argument_name in argument_names
    ]
    return f"={function_name}({', '.join(resolved_arguments)})"


def _expected_values_map(summary: RegressionSummary) -> dict[str, float | int]:
    """Map workbook function names to expected values from regression analysis."""

    return {
        "Observations": summary.observations,
        "DF_Regression": summary.df_regression,
        "DF_Total": summary.df_total,
        "R_squared": summary.r_squared,
        "DF_Residual": summary.df_residual,
        "Multiple_R": summary.multiple_r,
        "Adjusted_R2": summary.adjusted_r2,
    }


def _analysis_cache_signature(csv_path: Path, definitions: list[LambdaDefinition]) -> dict[str, Any]:
    """Return the cache key fields for analysis artifacts."""

    resolved_csv_path = csv_path.resolve()
    csv_stat = resolved_csv_path.stat()
    return {
        "csv_path": str(resolved_csv_path),
        "csv_size": csv_stat.st_size,
        "csv_mtime_ns": csv_stat.st_mtime_ns,
        "definition_names": [definition.name for definition in definitions],
        "definitions_content": [
            [definition.name, definition.formula_display] for definition in definitions
        ],
        "include_intercept": True,
    }


def _load_cached_expected_values(
    cache_path: Path,
    cache_signature: dict[str, Any],
) -> dict[str, float | int] | None:
    """Load cached expected values when the dataset/functions signature still matches."""

    if not cache_path.exists():
        return None

    try:
        with cache_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return None

    if not isinstance(payload, dict):
        return None
    if payload.get("signature") != cache_signature:
        return None

    expected_values = payload.get("expected_values")
    if not isinstance(expected_values, dict):
        return None

    required_keys = set(EXPECTED_RESULT_KEYS)
    if not required_keys.issubset(expected_values):
        return None

    return expected_values


def _write_analysis_cache(
    cache_path: Path,
    cache_signature: dict[str, Any],
    expected_values: dict[str, float | int],
) -> None:
    """Persist expected-values cache for later workbook builds."""

    payload = {"signature": cache_signature, "expected_values": expected_values}
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with cache_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)
        handle.write("\n")


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
        for arg in raw_arguments:
            arg_name = str(arg.get("name", "")).strip()
            arg_desc = str(arg.get("description", "")).strip()
            display = f"[{arg_name}]" if arg.get("optional") else arg_name
            comment_lines.append(f"{display}: {arg_desc}")
        comment = "\n\n".join(comment_lines)

        seen_names.add(name)
        definitions.append(
            LambdaDefinition(
                name=name,
                formula_display=formula_display,
                comment=comment,
            )
        )

    return definitions


def ensure_formula_test_sheet(
    workbook: xw.Book,
    definitions,
    definition_count: int,
    expected_true: dict[str, float | int],
    expected_false: dict[str, float | int],
) -> None:
    """Create or refresh the MLR_Testing sheet with a metric comparison table."""

    sheet_names = {sheet.name for sheet in workbook.sheets}
    if STARTER_SHEET_NAME in sheet_names:
        sheet = workbook.sheets[STARTER_SHEET_NAME]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = STARTER_SHEET_NAME
    else:
        sheet = workbook.sheets.add(name=STARTER_SHEET_NAME, after=workbook.sheets[-1])

    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()

    # Title area
    sheet["A1"].value = "MLR Testing"
    sheet["A1"].api.Font.Bold = True
    sheet["A2"].value = f"Registered definitions: {definition_count}"

    # Set sheet-scoped names
    _set_sheet_scoped_names(sheet)

    # Build metric comparison table with two rows: Allow_Intercept TRUE and FALSE
    header_row = _MLR_TABLE_HEADER_ROW
    last_data_row = header_row + 2
    col_count = len(_MLR_TABLE_COLUMNS)

    for col_idx, (header, _, _, _) in enumerate(_MLR_TABLE_COLUMNS, start=1):
        sheet.range((header_row, col_idx)).value = header

    table_range = sheet.range((header_row, 1), (last_data_row, col_count))
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = _MLR_TABLE_NAME
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = False
    table.ShowTableStyleColumnStripes = True

    rows_data = [
        (header_row + 1, True,  expected_true),
        (header_row + 2, False, expected_false),
    ]

    for data_row, allow_intercept_val, expected_values in rows_data:
        is_first_row = data_row == header_row + 1
        for col_idx, (header, exp_key, formula, num_fmt) in enumerate(_MLR_TABLE_COLUMNS, start=1):
            cell = sheet.range((data_row, col_idx))
            cell.api.NumberFormat = num_fmt
            if header == "Allow_Intercept":
                cell.value = allow_intercept_val
            elif formula is not None:
                # Formulas auto-propagate in a table; only write on the first row
                if is_first_row:
                    cell.formula = formula
            elif exp_key is not None:
                val = expected_values.get(exp_key)
                if val is not None:
                    cell.value = val

    sheet.range((1, 1), (last_data_row + 100, 1)).api.WrapText = True
    sheet.range("A1").column_width = 100
    sheet.range((header_row, 2), (last_data_row, col_count)).columns.autofit()
    sheet.range((header_row, 1), (last_data_row, col_count)).api.EntireRow.AutoFit()
    sheet.api.Application.ActiveWindow.SplitRow = header_row - 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True


def _get_or_create_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet:
    """Return an existing worksheet by name or append a new one."""

    for sheet in workbook.sheets:
        if sheet.name == sheet_name:
            return sheet
    return workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])


def _reset_generated_sheet(sheet: xw.Sheet) -> None:
    """Clear worksheet cells and remove existing list objects."""

    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()


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
    return parser.parse_args()


def _delete_sheet_if_present(workbook: xw.Book, sheet_name: str) -> None:
    """Delete a worksheet by name if it exists in the workbook."""

    for sheet in list(workbook.sheets):
        if sheet.name == sheet_name:
            sheet.delete()
            return


def _write_name_comments(workbook_path: Path, definitions: list[LambdaDefinition]) -> None:
    """Set Name Manager comments via COM so newlines render correctly in intellisense."""

    comments = {d.name: d.comment for d in definitions if d.comment}
    if not comments:
        return

    try:
        with xw.App(visible=False, add_book=False) as app:
            workbook = app.books.open(str(workbook_path))
            try:
                for name_obj in workbook.api.Names:
                    bare = name_obj.Name.split("!")[-1]
                    if bare in comments:
                        name_obj.Comment = comments[bare]
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        _raise_excel_access_error(workbook_path, "set comments on", exc)


def build_lambda_library(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    csv_path: Path = DEFAULT_CSV_PATH,
    validate_reopen: bool = False,
) -> NameSyncResult:
    """Build all workbook assets, sync JSON-backed names, and optionally validate reopen."""

    definitions = load_lambda_definitions(definitions_path)
    catalog_entries = load_catalog_entries(definitions_path)
    summary_true = calculate_regression_summary(input_csv_path=csv_path, include_intercept=True)
    summary_false = calculate_regression_summary(input_csv_path=csv_path, include_intercept=False)
    expected_true = _expected_values_map(summary_true)
    expected_false = _expected_values_map(summary_false)
    csv_headers, csv_rows = load_life_expectancy_rows(csv_path)
    workbook_path = workbook_path.resolve()
    workbook_exists = workbook_path.exists()

    try:
        with xw.App(visible=False, add_book=False) as app:
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
                ensure_formula_test_sheet(
                    workbook,
                    catalog_entries,
                    len(definitions),
                    expected_true,
                    expected_false,
                )
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        _raise_excel_access_error(workbook_path, "open or save", exc)

    try:
        result = sync_workbook_names(workbook_path, definitions)
    except OPEN_WORKBOOK_ERRORS as exc:
        _raise_excel_access_error(workbook_path, "update", exc)
    except (PermissionError, OSError) as exc:
        _raise_excel_access_error(workbook_path, "update", exc)

    _write_name_comments(workbook_path, definitions)

    if validate_reopen:
        _validate_workbook_reopen(workbook_path)

    return result


def main() -> None:
    """Build the workbook and print a short sync summary for interactive use."""

    args = parse_args()
    result = build_lambda_library(
        workbook_path=args.workbook,
        definitions_path=args.definitions,
        csv_path=args.csv,
        validate_reopen=args.validate_reopen,
    )
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Sheet updated: {STARTER_SHEET_NAME}")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    print("Invalid entries: 0")
    if args.validate_reopen:
        print("Reopen validation: passed")


if __name__ == "__main__":
    main()

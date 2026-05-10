from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from pathlib import Path
import re
import tempfile
from typing import Any
from typing import Iterable
import zipfile

from lxml import etree
from analyze_life_expectancy import (
    FEATURE_COLUMNS,
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
PREDICTIONS_SHEET_NAME = "Life Expectancy Predictions"
WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_SRC_RANGE = 1
XL_YES = 1
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)

_MLR_TABLE_HEADER_ROW = 4
_MAX_TEST_SHEET_NAME_LEN = 31
_INVALID_WORKSHEET_NAME_CHARS = set("[]:*?/\\")
_VALID_TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_MAX_EXCEL_ROW = 1_048_576

# Number of predictor columns (k) tested — one pair of rows per value (TRUE/FALSE intercept).
_MLR_K_VALUES: list[int] = [1, 5, 10, 18]

# Replaces x_s in every (Calc.) formula; resolves dynamically from each row's ind_vars value.
_MLR_X_S_OFFSET = "OFFSET(y,0,1,ROWS(y),[@[ind_vars]])"

# (header, expected_key_or_None, formula_or_None, number_format)
_ColumnSpec = tuple[str, str | None, str | None, str]

# Per-row config: (per-row cell values, expected-values dict)
_RowConfig = tuple[dict[str, Any], dict[str, float | int]]


@dataclass(frozen=True)
class LambdaDefinition:
    """Represents one workbook-scoped LAMBDA definition from the JSON catalog."""

    name: str
    formula_display: str
    comment: str = ""
    argument_names: tuple[str, ...] = ()
    test_table: str | None = None
    number_format: str = "General"

    @property
    def workbook_xml_formula_from_display(self) -> str:
        """Return the display formula translated into workbook.xml token syntax."""

        return to_workbook_xml_formula_from_display(self.formula_display)


@dataclass(frozen=True)
class NameSyncResult:
    """Summarizes how many workbook names were created versus replaced."""

    created: int
    updated: int


def _looks_like_a1_reference(value: str) -> bool:
    """Return True when a candidate name resembles an Excel A1-style reference."""

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
    """Validate a catalog test-table tag for both worksheet and table-name usage."""

    if len(tag) > _MAX_TEST_SHEET_NAME_LEN:
        raise ValueError(
            f"'test_table' value {tag!r} in entry {entry_index} is "
            f"{len(tag)} characters; worksheet names may be at most "
            f"{_MAX_TEST_SHEET_NAME_LEN} characters."
        )
    if any(char in _INVALID_WORKSHEET_NAME_CHARS for char in tag):
        invalid_chars = "".join(sorted({char for char in tag if char in _INVALID_WORKSHEET_NAME_CHARS}))
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
        "y":   "=LifeExpectancyData[Life expectancy]",
        "fil": "=LifeExpectancyData[Full_Data]",
    }

    # Remove legacy x_s names from older workbook versions
    for legacy in ("x_s", "x_s_k1", "x_s_k5", "x_s_k10", "x_s_k18"):
        _delete_sheet_scoped_name_if_present(sheet, legacy)

    for name, refers_to in local_names.items():
        _delete_sheet_scoped_name_if_present(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)

    _delete_sheet_scoped_name_if_present(sheet, "Allow_Intercept")


def _actual_formula(
    function_name: str,
    argument_names: Iterable[str],
    x_s_name: str = "x_s",
) -> str:
    """Build an executable test formula that references local helper names.

    When x_s_name is not the literal 'x_s' and the function takes an x_s argument,
    the formula is wrapped in LET(x_s, <x_s_name>, ...) so the LAMBDA is called
    with a plain 'x_s' reference rather than an inlined expression.
    """

    args = list(argument_names)
    uses_x_s = any(a.strip().lower() == "x_s" for a in args)
    use_let = uses_x_s and x_s_name != "x_s"

    reference_map = {
        "y": "y",
        "x_s": "x_s" if use_let else x_s_name,
        "filter": "fil",
        "allow_intercept": "[@[Allow_Intercept]]",
    }

    resolved_arguments = [
        reference_map.get(argument_name.strip().lower(), argument_name)
        for argument_name in args
    ]
    call = f"{function_name}({', '.join(resolved_arguments)})"
    return f"=LET(x_s, {x_s_name}, {call})" if use_let else f"={call}"


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


def load_lambda_definitions(path: Path) -> list[LambdaDefinition]:
    """Load, validate, and normalize the JSON catalog into LambdaDefinition objects."""

    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    functions = payload.get("functions")
    if not isinstance(functions, list):
        raise ValueError("lambda_functions.json must contain a top-level 'functions' array.")

    definitions: list[LambdaDefinition] = []
    seen_names: set[str] = set()
    seen_test_tables: set[str] = set()

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
            normalized_test_table = test_table.casefold()
            if normalized_test_table in seen_test_tables:
                raise ValueError(
                    f"Duplicate 'test_table' value in lambda_functions.json: {test_table}"
                )
            seen_test_tables.add(normalized_test_table)

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


def _test_sheet_name(test_table: str) -> str:
    """Return the worksheet name for a given test-table tag."""
    return test_table


def _test_table_name(test_table: str) -> str:
    """Return the Excel ListObject name for a given test-table tag."""
    return test_table


def build_smoke_test_formula(
    test_table: str,  # reserved for per-table customization
    metric_columns: list[_ColumnSpec],
) -> str:
    """Build the =AND(...) smoke-test formula from the metric column pairs."""

    _ = test_table
    comparisons: list[str] = []
    for exp_col, calc_col in zip(metric_columns[::2], metric_columns[1::2]):
        exp_header = exp_col[0]
        calc_header = calc_col[0]
        num_fmt = exp_col[3]
        if num_fmt == "0":
            comparisons.append(f"[@[{exp_header}]]=[@[{calc_header}]]")
        else:
            comparisons.append(
                f"ROUND([@[{exp_header}]],10)=ROUND([@[{calc_header}]],10)"
            )
    return "=AND(" + ",".join(comparisons) + ")"


def build_test_columns(
    test_table: str,
    definitions: list[LambdaDefinition],
) -> list[_ColumnSpec]:
    """Build the full column spec list for a test table from the matching definitions."""

    filtered = [d for d in definitions if d.test_table == test_table]
    metric_columns: list[_ColumnSpec] = []
    for d in filtered:
        metric_columns.append((f"{d.name} (Exp.)", d.name, None, d.number_format))
        metric_columns.append(
            (
                f"{d.name} (Calc.)",
                None,
                _actual_formula(d.name, d.argument_names, _MLR_X_S_OFFSET),
                d.number_format,
            )
        )

    smoke_formula = build_smoke_test_formula(test_table, metric_columns)
    fixed_columns: list[_ColumnSpec] = [
        (
            "X_Variables",
            None,
            "=TEXTJOIN(\", \",TRUE,OFFSET(y,-1,[@[ind_vars]],1,[@[ind_vars]]))",
            "General",
        ),
        ("ind_vars",       None, None, "0"),
        ("Allow_Intercept", None, None, "General"),
        ("Smoke Test",     None, smoke_formula, "General"),
    ]
    return fixed_columns + metric_columns


def write_test_table(
    sheet: xw.Sheet,
    table_name: str,
    columns: list[_ColumnSpec],
    header_row: int,
    rows_data: list[tuple[int, dict[str, Any], dict[str, float | int]]],
) -> None:
    """Create and populate an Excel ListObject for a test table."""

    last_data_row = header_row + len(rows_data)
    col_count = len(columns)

    for col_idx, (header, _, _, _) in enumerate(columns, start=1):
        sheet.range((header_row, col_idx)).value = header

    table_range = sheet.range((header_row, 1), (last_data_row, col_count))
    table = sheet.api.ListObjects.Add(
        SourceType=XL_SRC_RANGE,
        Source=table_range.api,
        XlListObjectHasHeaders=XL_YES,
    )
    table.Name = table_name
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = False
    table.ShowTableStyleColumnStripes = True

    for data_row, row_values, expected_values in rows_data:
        is_first_row = data_row == header_row + 1
        for col_idx, (header, exp_key, formula, num_fmt) in enumerate(columns, start=1):
            cell = sheet.range((data_row, col_idx))
            cell.api.NumberFormat = num_fmt
            if header in row_values:
                cell.value = row_values[header]
            elif formula is not None:
                # Formulas auto-propagate within a table; only write on the first row
                if is_first_row:
                    cell.formula = formula
            elif exp_key is not None:
                val = expected_values.get(exp_key)
                if val is not None:
                    cell.value = val


def ensure_mlr_scalar_test_sheet(
    workbook: xw.Book,
    definitions: list[LambdaDefinition],
    definition_count: int,
    row_configs: list[_RowConfig],
) -> None:
    """Create or refresh the MLR_Scalar_Test sheet with a single metric comparison table."""

    test_table = "MLR_Scalar_Test"
    sheet_name = _test_sheet_name(test_table)

    sheet_names = {sheet.name for sheet in workbook.sheets}
    if sheet_name in sheet_names:
        sheet = workbook.sheets[sheet_name]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = sheet_name
    else:
        sheet = workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])

    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()

    sheet["A1"].value = "MLR Testing"
    sheet["A1"].api.Font.Bold = True
    sheet["A2"].value = f"Registered definitions: {definition_count}"

    _set_sheet_scoped_names(sheet)

    columns = build_test_columns(test_table, definitions)
    header_row = _MLR_TABLE_HEADER_ROW
    last_data_row = header_row + len(row_configs)
    col_count = len(columns)

    rows_data: list[tuple[int, dict[str, Any], dict[str, float | int]]] = [
        (header_row + 1 + i, row_vals, expected)
        for i, (row_vals, expected) in enumerate(row_configs)
    ]

    write_test_table(sheet, _test_table_name(test_table), columns, header_row, rows_data)

    sheet.range((header_row, 2), (last_data_row, col_count)).columns.autofit()
    sheet.range((header_row, 1), (last_data_row, col_count)).api.EntireRow.AutoFit()
    sheet.api.Application.ActiveWindow.SplitRow = header_row - 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True
    sheet.range((1, 1), (last_data_row + 100, 1)).api.WrapText = True
    sheet.range("A1").column_width = 100


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

    row_configs: list[_RowConfig] = []
    for k in _MLR_K_VALUES:
        feature_cols = FEATURE_COLUMNS[:k]
        summary_true = calculate_regression_summary(
            input_csv_path=csv_path,
            include_intercept=True,
            feature_columns=feature_cols,
        )
        summary_false = calculate_regression_summary(
            input_csv_path=csv_path,
            include_intercept=False,
            feature_columns=feature_cols,
        )
        row_configs.append(({"ind_vars": k, "Allow_Intercept": True},  _expected_values_map(summary_true)))
        row_configs.append(({"ind_vars": k, "Allow_Intercept": False}, _expected_values_map(summary_false)))

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
                ensure_mlr_scalar_test_sheet(
                    workbook,
                    definitions,
                    len(definitions),
                    row_configs,
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
    print(f"Sheet updated: {_test_sheet_name('MLR_Scalar_Test')}")
    print("Sheet updated: LAMBDA_functions")
    print("Sheet updated: Life Expectancy Data")
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    print("Invalid entries: 0")
    if args.validate_reopen:
        print("Reopen validation: passed")


if __name__ == "__main__":
    main()

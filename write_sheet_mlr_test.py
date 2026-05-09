"""
build_mlr.py
Generates MLR_Lambda_Library.xlsx from LambdaFunction dataclasses loaded from JSON.
To add a new function: append a new object to lambda_functions.json.

Dependencies: xlwings, lxml
"""
from __future__ import annotations
from dataclasses import dataclass
import json
import os
from typing import Any
import zipfile
from lxml import etree
from lambda_formula_parser import to_workbook_xml_formula_from_display
import xlwings as xw


# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class Argument:
    name: str
    description: str
    optional: bool = False

    def display_name(self) -> str:
        return f"[{self.name}]" if self.optional else self.name


@dataclass
class LambdaFunction:
    name: str
    formula_display: str        # Indented/newline form shown in the Definition cell
    arguments: list[Argument]
    yields: str
    description: str

    def arguments_cell_text(self) -> str:
        """Generate the Arguments cell value from the Argument list."""
        lines = [
            f"{a.display_name()}: {a.description}"
            for a in self.arguments
        ]
        return "\n".join(lines)


# ── Layout constants ──────────────────────────────────────────────────────────

COL_WIDTHS = {"A": 22, "B": 72, "C": 55, "D": 36, "E": 64}
TABLE_COLS  = ["Function Name", "Definition", "Arguments", "Yields", "Description"]
ROW_HEIGHT  = 130
SHEET_NAME  = "MLR"
TABLE_NAME  = "Lambda_Definitions"
FUNCTIONS_JSON_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "lambda_functions.json",
)


# ── XML patch for DefinedNames ────────────────────────────────────────────────

def _patch_defined_names(path: str, functions: list[LambdaFunction]) -> None:
    """
    Excel's COM Names.Add rejects LAMBDA formulas with optional [params] at
    write time. This patches workbook.xml directly — same technique Excel uses
    internally when reading a saved file.
    """
    ns  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    tmp = path + ".tmp.xlsx"

    with zipfile.ZipFile(path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                tree = etree.fromstring(data)
                for el in tree.findall(f"{{{ns}}}definedNames"):
                    tree.remove(el)
                dn_el = etree.SubElement(tree, f"{{{ns}}}definedNames")
                for fn in functions:
                    name_el = etree.SubElement(dn_el, f"{{{ns}}}definedName")
                    name_el.set("name", fn.name)
                    name_el.text = to_workbook_xml_formula_from_display(fn.formula_display)
                data = etree.tostring(
                    tree, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            zout.writestr(item, data)

    os.replace(tmp, path)


# ── Workbook builder ──────────────────────────────────────────────────────────

def _open_or_create_workbook(app: xw.App, path: str) -> tuple[xw.Book, bool]:
    if os.path.exists(path):
        try:
            return app.books.open(path), True
        except Exception as exc:
            raise RuntimeError(
                f"Could not open existing workbook at {path}. Close it in Excel and try again."
            ) from exc

    return app.books.add(), False


def _get_or_create_sheet(wb: xw.Book, name: str) -> xw.Sheet:
    for sheet in wb.sheets:
        if sheet.name == name:
            return sheet
    return wb.sheets.add(name=name, after=wb.sheets[-1])


def _reset_generated_sheet(ws: xw.Sheet) -> None:
    for idx in range(ws.api.ListObjects.Count, 0, -1):
        ws.api.ListObjects(idx).Delete()
    ws.api.Cells.Clear()


def build_workbook(functions: list[LambdaFunction], path: str) -> None:
    app = xw.App(visible=False)
    wb, existing_workbook = _open_or_create_workbook(app, path)
    ws = _get_or_create_sheet(wb, SHEET_NAME)
    _reset_generated_sheet(ws)
    ws.activate()

    # Info rows
    ws["A1"].value = "MLR Lambda Library"
    ws["A2"].value = (
        "Workbook-scoped LAMBDA functions for multilinear regression. "
        "Copy this sheet into any workbook to carry the named formulas with it."
    )

    # Table header row (row 4)
    for col_idx, header in enumerate(TABLE_COLS, start=1):
        ws.range((4, col_idx)).value = header

    # Data rows
    last_data_row = 4 + len(functions)
    # NumberFormat "@" forces text storage so =LAMBDA(... is not evaluated as a formula
    ws.range(f"B5:B{last_data_row}").api.NumberFormat = "@"

    for row_offset, fn in enumerate(functions):
        r = 5 + row_offset
        ws.range((r, 1)).value = fn.name
        ws.range((r, 2)).value = fn.formula_display
        ws.range((r, 3)).value = fn.arguments_cell_text()
        ws.range((r, 4)).value = fn.yields
        ws.range((r, 5)).value = fn.description
        ws.range(f"{r}:{r}").row_height = ROW_HEIGHT

    # Excel Table
    tbl_range = ws.range(f"A4:E{last_data_row}")
    lo = ws.api.ListObjects.Add(
        SourceType=1,            # xlSrcRange
        Source=tbl_range.api,
        XlListObjectHasHeaders=1 # xlYes
    )
    lo.Name = TABLE_NAME

    # Column widths
    for col_letter, width in COL_WIDTHS.items():
        ws.range(f"{col_letter}:{col_letter}").column_width = width

    # Freeze rows 1–4 (first scrollable row = 5)
    ws.api.Application.ActiveWindow.SplitRow    = 4
    ws.api.Application.ActiveWindow.SplitColumn = 0
    ws.api.Application.ActiveWindow.FreezePanes = True

    # Save and close so we can patch the file; app stays alive for the reopen
    if existing_workbook:
        wb.save()
    else:
        wb.save(path)
    wb.close()

    # COM's Names.Add rejects LAMBDA with optional [params] — patch XML directly
    _patch_defined_names(path, functions)
    print(f"Saved: {path}")

    # Reopen visibly so the user can inspect
    app.visible = True
    app.books.open(path)


# ── Function definitions ──────────────────────────────────────────────────────

def _build_argument(argument_data: dict[str, Any]) -> Argument:
    return Argument(
        name=str(argument_data["name"]),
        description=str(argument_data["description"]),
        optional=bool(argument_data.get("optional", False)),
    )


def _build_function(function_data: dict[str, Any]) -> LambdaFunction:
    return LambdaFunction(
        name=str(function_data["name"]),
        formula_display=str(function_data["formula_display"]),
        arguments=[
            _build_argument(argument_data)
            for argument_data in function_data["arguments"]
        ],
        yields=str(function_data["yields"]),
        description=str(function_data["description"]),
    )


def load_functions(path: str = FUNCTIONS_JSON_PATH) -> list[LambdaFunction]:
    with open(path, encoding="utf-8") as file:
        payload = json.load(file)

    return [_build_function(function_data) for function_data in payload["functions"]]


FUNCTIONS = load_functions()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import os
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MLR_Lambda_Library.xlsx")
    build_workbook(FUNCTIONS, out)
    print("Done. Excel is open — the existing workbook was updated in place, or created if missing.")

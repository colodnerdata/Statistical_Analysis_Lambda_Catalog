"""Sync workbook Named Ranges from lambda_functions.json without xlwings (Linux-safe)."""
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path

import lxml.etree as etree

sys.path.insert(0, str(Path(__file__).parent))
from lambda_catalog.lambda_formula_parser import to_workbook_xml_formula_from_display

ROOT_DIR = Path(__file__).parent
WORKBOOK_PATH = ROOT_DIR / "Lambda_Library.xlsx"
DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"
WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def load_definitions(path: Path) -> list[tuple[str, str]]:
    with path.open("r", encoding="utf-8") as f:
        payload = json.load(f)
    return [(item["name"], item["formula_display"]) for item in payload["functions"]]


def sync_names(workbook_path: Path, definitions: list[tuple[str, str]]) -> tuple[int, int]:
    target_names = {name for name, _ in definitions}
    existing: set[str] = set()
    temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=workbook_path.parent)
    os.close(temp_fd)
    Path(temp_name).unlink(missing_ok=True)

    try:
        with zipfile.ZipFile(workbook_path, "r") as zin, zipfile.ZipFile(
            temp_name, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == "xl/workbook.xml":
                    root = etree.fromstring(data)
                    ns = WORKBOOK_NS
                    dn_list = root.findall(f"{{{ns}}}definedNames")
                    if dn_list:
                        dn = dn_list[0]
                        for extra in dn_list[1:]:
                            root.remove(extra)
                    else:
                        dn = etree.Element(f"{{{ns}}}definedNames")
                        root.append(dn)

                    for el in list(dn):
                        if el.tag != f"{{{ns}}}definedName":
                            continue
                        if el.get("localSheetId") is None and el.get("name") in target_names:
                            existing.add(el.get("name"))
                            dn.remove(el)

                    for name, formula_display in definitions:
                        el = etree.SubElement(dn, f"{{{ns}}}definedName")
                        el.set("name", name)
                        el.text = to_workbook_xml_formula_from_display(formula_display)

                    data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

                elif item.filename == "xl/calcChain.xml":
                    continue
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    r = etree.fromstring(data)
                    _RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
                    CALC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
                    for rel in r.findall(f"{{{_RELS_NS}}}Relationship"):
                        if rel.get("Type") == CALC_REL:
                            r.remove(rel)
                    data = etree.tostring(r, encoding="UTF-8", xml_declaration=True, standalone=True)

                zout.writestr(item, data)

        Path(temp_name).replace(workbook_path)
    finally:
        Path(temp_name).unlink(missing_ok=True)

    created = len(definitions) - len(existing)
    updated = len(existing)
    return created, updated


if __name__ == "__main__":
    defs = load_definitions(DEFINITIONS_PATH)
    created, updated = sync_names(WORKBOOK_PATH, defs)
    print(f"Named ranges synced: {created} created, {updated} updated")
    print(f"Total: {len(defs)} lambdas in workbook")

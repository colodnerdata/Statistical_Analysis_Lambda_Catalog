"""LAMBDA name-syncing via workbook.xml patching and workbook lifecycle helpers."""
from __future__ import annotations

import os
import tempfile
import zipfile
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path

import lxml.etree as etree  # type: ignore[import-untyped]  # pyright: ignore[reportMissingTypeStubs]
import xlwings as xw

from lambda_catalog.catalog_schema import CatalogFunction
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, excel_error_message


WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CALC_CHAIN_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
XL_CALCULATION_AUTOMATIC     = -4105  # Excel XlCalculation.xlCalculationAutomatic
XL_CALCULATION_MANUAL        = -4135  # Excel XlCalculation.xlCalculationManual
XL_CALCULATION_SEMIAUTOMATIC = 2      # Excel XlCalculation.xlCalculationSemiautomatic
                                       # (auto recalc, but ignore changes in data tables)


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
    workbook_path: Path, definitions: Sequence[CatalogFunction]
) -> NameSyncResult:
    """Replace workbook-scoped names in workbook.xml with the supplied definitions.

    Parameters
    ----------
    workbook_path : Path
        Path to the .xlsx file to patch in place.
    definitions : Sequence[CatalogFunction]
        Functions whose names will be inserted or replaced.

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
                        if definition.notes:
                            name_element.set("comment", definition.notes)
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


def drop_workbook_names(workbook_path: Path, names: Iterable[str]) -> int:
    """Remove workbook-scoped defined names from a closed workbook package.

    Parameters
    ----------
    workbook_path : Path
        Path to the .xlsx file to patch in place.
    names : Iterable[str]
        Workbook-scoped defined names to remove.

    Returns
    -------
    int
        Number of names removed.
    """
    target_names = set(names)
    removed = 0
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
                    for defined_names in workbook_root.findall(
                        f"{{{WORKBOOK_NS}}}definedNames"
                    ):
                        for name_element in list(defined_names):
                            if name_element.tag != f"{{{WORKBOOK_NS}}}definedName":
                                continue

                            target_name = name_element.get("name")
                            is_workbook_scope = name_element.get("localSheetId") is None
                            if is_workbook_scope and target_name in target_names:
                                defined_names.remove(name_element)
                                removed += 1

                    xml_body = str(etree.tostring(workbook_root, encoding="unicode"))
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        + xml_body
                    ).encode("UTF-8")

                elif item.filename == "xl/calcChain.xml":
                    continue

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

    return removed


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

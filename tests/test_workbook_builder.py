"""Tests for workbook package patching helpers that do not require Excel."""
from __future__ import annotations

import zipfile

from lambda_catalog.workbook_builder import drop_workbook_names


def test_drop_workbook_names_removes_only_workbook_scoped_targets(tmp_path) -> None:
    workbook_path = tmp_path / "book.xlsx"
    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets/>
  <definedNames>
    <definedName name="Predictions">LAMBDA(x,x)</definedName>
    <definedName name="Predictions" localSheetId="0">'Sheet1'!$A$1</definedName>
    <definedName name="Keep_Me">LAMBDA(x,x)</definedName>
  </definedNames>
  <calcPr/>
</workbook>"""
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"""
    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"""

    with zipfile.ZipFile(workbook_path, "w") as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("xl/calcChain.xml", "<calcChain/>")
        zf.writestr("xl/worksheets/sheet1.xml", "<worksheet/>")

    removed = drop_workbook_names(workbook_path, {"Predictions"})

    assert removed == 1
    with zipfile.ZipFile(workbook_path) as zf:
        patched_workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")
        patched_rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        patched_content_types_xml = zf.read("[Content_Types].xml").decode("utf-8")
        names = set(zf.namelist())

    assert '<definedName name="Predictions">LAMBDA(x,x)</definedName>' not in patched_workbook_xml
    assert '<definedName name="Predictions" localSheetId="0">' in patched_workbook_xml
    assert '<definedName name="Keep_Me">LAMBDA(x,x)</definedName>' in patched_workbook_xml
    assert "calcChain" not in patched_rels_xml
    assert "calcChain" not in patched_content_types_xml
    assert "xl/calcChain.xml" not in names
    assert "xl/worksheets/sheet1.xml" in names

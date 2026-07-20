"""Tests for workbook package patching helpers that do not require Excel."""
from __future__ import annotations

import zipfile

from lambda_catalog.catalog_schema import CatalogFunction
from lambda_catalog.workbook_builder import drop_workbook_names, sync_workbook_names


def _catalog_function(name: str, formula: str = "LAMBDA(x,x)") -> CatalogFunction:
    """Build a minimal CatalogFunction for sync_workbook_names tests."""
    return CatalogFunction(
        name=name,
        formula_display=formula,
        arguments=(),
        yields="x",
        description="",
        plain_language_summary="",
        test_table=None,
        number_format="General",
        notes="",
    )


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


def test_sync_workbook_names_strips_broken_workbook_scoped_residue(tmp_path) -> None:
    """The committed workbook sometimes carries workbook-scoped residue from
    previous builds: entries whose body is '#REF!' or who duplicate a name
    that already has a sheet-scoped entry. sync_workbook_names must strip
    these so they never propagate to a fresh build.

    See the TestRealWorkbook::test_no_orphan_named_ranges finding in
    tests/test_workbook_invariants.py for the production-artifact signal.
    """
    workbook_path = tmp_path / "book.xlsx"
    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets/>
  <definedNames>
    <definedName name="Source_Data" localSheetId="5">'Regression'!$A$1:$A$100</definedName>
    <definedName name="Source_Data">#REF!</definedName>
    <definedName name="Spec_Role" localSheetId="5">'Regression'!$B$1:$B$100</definedName>
    <definedName name="Spec_Role">#NAME?</definedName>
    <definedName name="Sample_Include" localSheetId="5">LAMBDA(x,x)</definedName>
    <definedName name="Sample_Include">_xlfn.LAMBDA(_xlfn.LET([0]!Source_Data,1))</definedName>
    <definedName name="Keep_Me">LAMBDA(x,x)</definedName>
  </definedNames>
</workbook>"""

    with zipfile.ZipFile(workbook_path, "w") as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>""")
        zf.writestr("xl/_rels/workbook.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""")

    result = sync_workbook_names(workbook_path, [_catalog_function("Keep_Me")])

    assert result.updated == 1  # Keep_Me was replaced
    assert result.created == 0  # Keep_Me already existed, no new name

    with zipfile.ZipFile(workbook_path) as zf:
        patched = zf.read("xl/workbook.xml").decode("utf-8")

    # Sheet-scoped entries preserved
    assert '<definedName name="Source_Data" localSheetId="5">' in patched
    assert '<definedName name="Spec_Role" localSheetId="5">' in patched

    # Workbook-scoped residue stripped (regardless of whether broken-body or
    # duplicate-of-sheet-scoped).
    assert '<definedName name="Source_Data">#REF!</definedName>' not in patched
    assert '<definedName name="Spec_Role">#NAME?</definedName>' not in patched
    assert '<definedName name="Sample_Include">_xlfn.LAMBDA' not in patched

    # Keep_Me replaced (single entry, the new one).
    assert patched.count('<definedName name="Keep_Me">') == 1
    assert 'LAMBDA(x,x)' in patched


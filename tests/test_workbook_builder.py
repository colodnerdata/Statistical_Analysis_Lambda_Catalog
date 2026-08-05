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


def _write_workbook(
    workbook_path,
    workbook_xml: str,
    *,
    extra_members: dict[str, str] | None = None,
    content_types_xml: str = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>""",
    rels_xml: str = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""",
) -> None:
    """Write a minimal .xlsx package with the given workbook.xml."""
    with zipfile.ZipFile(workbook_path, "w") as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        for member_name, member_body in (extra_members or {}).items():
            zf.writestr(member_name, member_body)


def test_sync_workbook_names_strips_broken_workbook_scoped_residue(tmp_path) -> None:
    """The committed workbook sometimes carries workbook-scoped residue from
    previous builds: entries whose body is '#REF!' or who duplicate a name
    that already has a sheet-scoped entry. sync_workbook_names must strip
    these so they never propagate to a fresh build.

    See the TestRealWorkbook::test_no_orphan_named_ranges finding in
    tests/test_workbook_invariants.py for the production-artifact signal.
    """
    workbook_path = tmp_path / "book.xlsx"
    _write_workbook(
        workbook_path,
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</workbook>""",
    )

    result = sync_workbook_names(workbook_path, [_catalog_function("Keep_Me")])

    assert result.updated == 1  # Keep_Me was replaced
    assert result.created == 0  # Keep_Me already existed, no new name
    assert result.removed == 3  # the three residue entries

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


def test_sync_workbook_names_purges_other_artifacts_chart_ranges(tmp_path) -> None:
    """The v3.0 split's residue: RegChart* entries promoted to workbook scope.

    The Univariate artifact carried twelve of these, each an ``OFFSET(#REF!,…)``
    body — a wrapped error, not a bare error literal, which is why matching on
    the body text never caught them. Workbook scope belongs to the catalog, so
    they go regardless of body. The sheet-scoped RegChart* range on the
    Regression sheet — where the real one lives — must survive untouched.
    """
    workbook_path = tmp_path / "book.xlsx"
    _write_workbook(
        workbook_path,
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Univariate" sheetId="1" r:id="rId1"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
  <definedNames>
    <definedName name="RegChartFitY">OFFSET(#REF!,1,0,MAX(IFERROR(#REF!,1),1),1)</definedName>
    <definedName name="RegChartFitY" localSheetId="0">OFFSET(Univariate!$AP$2,1,0,1,1)</definedName>
    <definedName name="UV_QQ_Normal" localSheetId="0">OFFSET(Univariate!$CY$4,1,0,1,1)</definedName>
    <definedName name="Retired_Lambda">_xlfn.LAMBDA(_xlpm.x,_xlpm.x)</definedName>
    <definedName name="_xlnm.Print_Area">Univariate!$A$1:$Z$99</definedName>
    <definedName name="Keep_Me">LAMBDA(x,x)</definedName>
  </definedNames>
</workbook>""",
    )

    result = sync_workbook_names(workbook_path, [_catalog_function("Keep_Me")])

    with zipfile.ZipFile(workbook_path) as zf:
        patched = zf.read("xl/workbook.xml").decode("utf-8")

    assert '<definedName name="RegChartFitY">' not in patched
    assert '<definedName name="Retired_Lambda">' not in patched
    assert result.removed == 2

    # Sheet-scoped ranges and Excel's own reserved names are left alone.
    assert '<definedName name="RegChartFitY" localSheetId="0">' in patched
    assert '<definedName name="UV_QQ_Normal" localSheetId="0">' in patched
    assert '<definedName name="_xlnm.Print_Area">' in patched


def test_sync_workbook_names_skips_definitions_referencing_absent_sheets(tmp_path) -> None:
    """A catalog LAMBDA that names a sheet this artifact lacks is not written.

    Writing it is what produced the standalone Univariate workbook's broken
    link: Excel rebinds ``Regression!Source_Data`` to ``[1]!Source_Data``
    against a missing external workbook rather than leaving it unresolved.
    """
    workbook_path = tmp_path / "book.xlsx"
    _write_workbook(
        workbook_path,
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Univariate" sheetId="1" r:id="rId1"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
  <definedNames/>
</workbook>""",
    )

    result = sync_workbook_names(
        workbook_path,
        [
            _catalog_function(
                "Base_Period_Delta", "LAMBDA(COLUMNS('Regression'!Source_Data))"
            ),
            _catalog_function("Local_Reader", "LAMBDA(COUNT('Univariate'!$A:$A))"),
            _catalog_function("Portable", "LAMBDA(x,x)"),
        ],
    )

    assert result.skipped == ("Base_Period_Delta",)
    assert result.created == 2

    with zipfile.ZipFile(workbook_path) as zf:
        patched = zf.read("xl/workbook.xml").decode("utf-8")

    assert '<definedName name="Base_Period_Delta">' not in patched
    assert '<definedName name="Local_Reader">' in patched
    assert '<definedName name="Portable">' in patched


def test_sync_workbook_names_reads_references_past_string_literals(tmp_path) -> None:
    """Text inside a string literal is not a reference.

    A message ending in an exclamation mark reads as ``Sheet!`` to a naive
    scan, and ``[1]`` inside a label reads as an external workbook — both
    would make a perfectly portable definition look unportable.
    """
    workbook_path = tmp_path / "book.xlsx"
    _write_workbook(
        workbook_path,
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Univariate" sheetId="1" r:id="rId1"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
  <definedNames/>
</workbook>""",
    )

    result = sync_workbook_names(
        workbook_path,
        [_catalog_function("Chatty", 'LAMBDA(x,IF(x>0,"Positive!","see note [1]"))')],
    )

    assert result.skipped == ()
    with zipfile.ZipFile(workbook_path) as zf:
        assert '<definedName name="Chatty">' in zf.read("xl/workbook.xml").decode("utf-8")


_EXTERNAL_LINK_WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Univariate" sheetId="1" r:id="rId1"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
  <externalReferences xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <externalReference r:id="rId5"/>
  </externalReferences>
  <definedNames>
    <definedName name="Base_Period_Delta">_xlfn.LAMBDA(COLUMNS([1]!Source_Data))</definedName>
  </definedNames>
</workbook>"""

_EXTERNAL_LINK_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>"""

_EXTERNAL_LINK_CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>"""


def _sheet_xml(formula: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData><row r="1"><c r="A1"><f>{formula}</f></c></row></sheetData>'
        "</worksheet>"
    )


def test_sync_workbook_names_strips_orphaned_external_link(tmp_path) -> None:
    """Dropping the last external reference must drop the link parts with it.

    An externalLink part left behind still shows up as ``<externalReferences>``
    in workbook.xml, so Excel keeps prompting about broken links even after the
    name that caused it is gone.
    """
    workbook_path = tmp_path / "book.xlsx"
    _write_workbook(
        workbook_path,
        _EXTERNAL_LINK_WORKBOOK_XML,
        rels_xml=_EXTERNAL_LINK_RELS_XML,
        content_types_xml=_EXTERNAL_LINK_CONTENT_TYPES_XML,
        extra_members={
            "xl/externalLinks/externalLink1.xml": "<externalLink/>",
            "xl/externalLinks/_rels/externalLink1.xml.rels": "<Relationships/>",
            "xl/worksheets/sheet1.xml": _sheet_xml("SUM(A2:A9)"),
        },
    )

    sync_workbook_names(workbook_path, [_catalog_function("Portable")])

    with zipfile.ZipFile(workbook_path) as zf:
        patched = zf.read("xl/workbook.xml").decode("utf-8")
        patched_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        patched_content_types = zf.read("[Content_Types].xml").decode("utf-8")
        members = set(zf.namelist())

    assert "externalReferences" not in patched
    assert "externalLink" not in patched_rels
    assert "externalLink" not in patched_content_types
    assert not any(member.startswith("xl/externalLinks/") for member in members)


def test_sync_workbook_names_keeps_external_link_a_cell_still_uses(tmp_path) -> None:
    """A live external reference in a cell formula keeps its link parts."""
    workbook_path = tmp_path / "book.xlsx"
    _write_workbook(
        workbook_path,
        _EXTERNAL_LINK_WORKBOOK_XML,
        rels_xml=_EXTERNAL_LINK_RELS_XML,
        content_types_xml=_EXTERNAL_LINK_CONTENT_TYPES_XML,
        extra_members={
            "xl/externalLinks/externalLink1.xml": "<externalLink/>",
            "xl/worksheets/sheet1.xml": _sheet_xml("[1]!Source_Data"),
        },
    )

    sync_workbook_names(workbook_path, [_catalog_function("Portable")])

    with zipfile.ZipFile(workbook_path) as zf:
        patched = zf.read("xl/workbook.xml").decode("utf-8")
        members = set(zf.namelist())

    assert "externalReferences" in patched
    assert "xl/externalLinks/externalLink1.xml" in members

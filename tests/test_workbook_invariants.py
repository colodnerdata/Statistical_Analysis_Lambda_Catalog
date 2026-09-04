"""Headless structural invariants for the produced Lambda_Library.xlsx package.

These tests are pure-Python (zipfile + lxml); they do not import xlwings and run
on every Linux CI commit in <1 s. They catch packaging regressions that the
existing `RecordingSheet` unit tests cannot see:

  * sheet-name drift against the spec
  * orphan defined names (empty body, invalid characters, duplicates)
  * workbook-scope ownership: only catalog functions and Excel's `_xlnm.*`
    names may be workbook-scoped, no defined-name body may carry an error
    literal, and the workbook may not link to another workbook (the pre-reunify
    v3.0 split left one artifact's chart ranges at workbook scope in the other
    file, and a Regression-coupled catalog LAMBDA left a broken external link
    in the Univariate one)
  * localSheetId values that fall outside [0, sheet_count)
  * #REF! / #NAME? / #VALUE! / #NULL! / #DIV/0! / #N/A / #NUM! literals leaking
    into cached cell values (the writers' LAMBDA formulas legitimately wrap
    `NA()` inside `IFERROR`, so we only scan <v> and <t> elements, not <f>)
  * [Content_Types].xml / xl/_rels/workbook.xml.rels inconsistencies
  * chart-relationship Targets that don't resolve to a real zip member
  * a chart referencing a sheet other than the one it sits on — SERIES formulas
    and the `c15:datalabelsRange` "Value From Cells" extension alike, since a
    chart lives in its own package part and every reference in it names a sheet
    explicitly, so a wrong name plots another sheet's model without erroring
  * workbook_builder.py: post-`sync_workbook_names` rebuilds must leave no
    calcChain part or relationship behind
  * the shipped Univariate distribution fits are at the maximum likelihood, and
    the shipped workbook is the one the current writer produces
    (``TestShippedUnivariateFits`` / ``TestShippedUnivariateLayout``)

The synthetic 4-sheet fixture (``build_headless_fixture``) is the always-on
source of truth for these invariants, alongside ``TestRealWorkbookNameScope``
and ``TestRealWorkbook``, which run the same checks against the committed
unified workbook on every commit. Those committed-artifact checks need no
Excel — they are zipfile and lxml — so the committed workbook is screened on
every push.

The deep spec-driven check (``lambda_catalog.deep_verify.verify_test_sheets``, xlwings + Excel
required) is the source of truth for cell-level correctness — *except* for the
Univariate fitting table, which it does not read at all. It checks the twelve
descriptive statistics and the three histogram blocks and stops there, so a
passing ``build_production.py --verify`` says nothing about whether the
distribution fits are right. That is the gap the two ``TestShipped*`` classes
close, and they close it headlessly: a built .xlsx caches every formula's last
computed result, so the shipped parameters can be read out of the zip and
checked against the reference MLEs in ``analyze_univariate`` with no Excel.
Those two are the only checks here that do numeric work rather than structural.
"""
# Pytest fixtures intentionally share names with the parameters they inject.
# pylint: disable=redefined-outer-name
from __future__ import annotations

import os
import posixpath
import re
import tempfile
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

import pytest
from lxml import etree  # type: ignore[import-untyped]

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import drop_workbook_names

# ---------------------------------------------------------------------------
# Namespaces and the constant set of Excel error literals we never want to see
# leak into a cached cell value. The cell value lives in a <v> element (or in
# an inline <t>). Formulas <f> legitimately reference NA() inside IFERROR
# wrappers, so we deliberately exclude those from the scan.
# ---------------------------------------------------------------------------

WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
# The r:id attribute on <sheet> lives in the officeDocument namespace, which is
# a different one from the package namespace the .rels parts themselves use.
OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EXCEL_ERROR_LITERALS = ("#REF!", "#NAME?", "#VALUE!", "#NULL!", "#DIV/0!", "#N/A", "#NUM!")
_VALID_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_WB = f"{{{WORKBOOK_NS}}}"
_CT = f"{{{CT_NS}}}"
_RELS = f"{{{RELS_NS}}}"
_OVERRIDE_TAG = f"{_CT}Override"
_RELATIONSHIP_TAG = f"{_RELS}Relationship"
_SHEET_TAG = f"{_WB}sheet"
_DEFINED_NAMES_TAG = f"{_WB}definedNames"
_DEFINED_NAME_TAG = f"{_WB}definedName"

# ---------------------------------------------------------------------------
# Deliberately-#N/A columns.
#
# The Cook's Distance chart labels only the influential points, from a masked
# overlay column — the real value where D exceeds the cutoff, a masking token
# everywhere else (see CLAUDE.md -> "Selective data labels"). The token was
# NA() while the labels came from ShowValue/ShowCategoryName, and every
# non-flagged row of that column cached a perfectly correct #N/A, which a naive
# scan reports as hundreds of false offenders. The labels now read the column
# through Value From Cells and the token is "", so a freshly built artifact has
# no #N/A there at all — but the exemption stays: it is what keeps an artifact
# built before that change (and any future NA()-masked overlay) from failing a
# scan over cells that are behaving exactly as designed.
#
# The exempt column is NOT hard-coded. It is read per sheet from the
# sheet-scoped RegChart named range that feeds the overlay series, so a layout
# change that moves the column moves the exemption with it, and every
# Regression-templated sheet is covered — the shipped artifact has one, the
# test-model workbook has 47, each with its own copy of the name.
#
# Only #N/A is forgiven, and only in that column. A #REF! or #VALUE! there is
# still a broken cell.
# ---------------------------------------------------------------------------
_NA_MASKED_RANGE_NAMES = ("RegChartCookDistFlag",)
_NA_LITERAL = "#N/A"
# $AY$3 out of "OFFSET(Regression!$AY$3,1,0,MAX(...),1)" — the anchor column of
# an OFFSET-based RegChart range.
_OFFSET_ANCHOR_RE = re.compile(r"OFFSET\([^!,]*!\$([A-Z]{1,3})\$\d+", re.IGNORECASE)
_CELL_COLUMN_RE = re.compile(r"^([A-Z]{1,3})\d+$")

# Real Lambda_Library.xlsx (the unified production artifact) ships nine sheets
# in the tab order that build_production._reorder_and_style_sheet_tabs
# applies: the three Regression workbench sheets front-most, then the Univariate
# workbench, the LAMBDA_functions catalog, Version History, and the three data
# sheets last.
EXPECTED_REAL_SHEETS: tuple[str, ...] = (
    "Regression",
    "Regression Instructions",
    "Diagnostic Guide",
    "Univariate",
    "LAMBDA_functions",
    "Version History",
    "Production Lots",
    "Life Expectancy Data",
    "Mileage Data",
)

# Synthetic 4-sheet fixture.
_SYNTHETIC_SHEETS = ("Sheet1", "Sheet2", "Sheet3", "Sheet4")


# ---------------------------------------------------------------------------
# Synthetic 4-sheet fixture. Mirrors the hand-rolled-xlsx style of
# tests/test_workbook_builder.py:9-52. Pure zipfile + lxml, no xlwings.
# ---------------------------------------------------------------------------


def _workbook_xml(sheet_names: Iterable[str]) -> bytes:
    sheet_elements = "".join(
        f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
        for i, name in enumerate(sheet_names)
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        b"<sheets>" + sheet_elements.encode("utf-8") + b"</sheets>"
        b"<definedNames>"
        b'<definedName name="Real_Name">LAMBDA(x,x)</definedName>'
        b'<definedName name="Local_Name" localSheetId="0">\'Sheet1\'!$A$1</definedName>'
        b"</definedNames>"
        b"</workbook>"
    )


def _workbook_rels(sheet_count: int) -> bytes:
    rels = "".join(
        f'<Relationship Id="rId{i + 1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{i + 1}.xml"/>'
        for i in range(sheet_count)
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels.encode("utf-8")
        + b"</Relationships>"
    )


def _content_types(sheet_count: int) -> bytes:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(sheet_count)
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" '
        b'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/xl/workbook.xml" '
        b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + overrides.encode("utf-8")
        + b"</Types>"
    )


def _blank_sheet_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<sheetData/>"
        b"</worksheet>"
    )


def _blank_drawing_rels_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )


def build_headless_fixture(
    tmp_path: Path,
    *,
    sheet_names: tuple[str, ...] = _SYNTHETIC_SHEETS,
    include_drawing: bool = False,
) -> Path:
    """Write a minimal but spec-conformant .xlsx to ``tmp_path`` and return it."""
    workbook_path = tmp_path / "fixture.xlsx"
    with zipfile.ZipFile(workbook_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types(len(sheet_names)))
        zf.writestr("xl/workbook.xml", _workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheet_names)))
        for i in range(len(sheet_names)):
            zf.writestr(f"xl/worksheets/sheet{i + 1}.xml", _blank_sheet_xml())
        if include_drawing:
            # A drawing whose rels file is intentionally empty (so a real test
            # can verify "no chart relationships" passes; a chart-relationship
            # test against this fixture would be vacuous).
            zf.writestr("xl/drawings/drawing1.xml", b"<?xml version='1.0'?><drawing/>")
            zf.writestr("xl/drawings/_rels/drawing1.xml.rels", _blank_drawing_rels_xml())
    return workbook_path


@dataclass(frozen=True)
class WorkbookPackage:
    """A view onto a parsed .xlsx package, shared by all invariant tests."""

    path: Path

    @property
    def namelist(self) -> set[str]:
        with zipfile.ZipFile(self.path) as zf:
            return set(zf.namelist())

    @property
    def workbook_root(self):
        with zipfile.ZipFile(self.path) as zf:
            return etree.fromstring(zf.read("xl/workbook.xml"))

    @property
    def content_types_root(self):
        with zipfile.ZipFile(self.path) as zf:
            return etree.fromstring(zf.read("[Content_Types].xml"))

    @property
    def workbook_rels_root(self):
        with zipfile.ZipFile(self.path) as zf:
            return etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    def sheet_part_for(self, sheet_name: str) -> str:
        """Return the zip member holding a named worksheet, via workbook rels."""
        rels = {
            rel.get("Id"): rel.get("Target")
            for rel in self.workbook_rels_root
        }
        for sheet in self.workbook_root.findall(f".//{_WB}sheets/{_WB}sheet"):
            if sheet.get("name") == sheet_name:
                rid = sheet.get(f"{{{OFFICE_RELS_NS}}}id")
                target = rels[rid]
                return posixpath.normpath(posixpath.join("xl", target))
        raise KeyError(f"sheet {sheet_name!r} not in {self.path.name}")

    def cell_values(self, sheet_part: str) -> dict[str, str]:
        """Return {A1 address: cached <v> text} for one worksheet.

        ``worksheet_cached_values`` flattens every value in the sheet, which
        cannot answer "what is in I8". A built workbook stores each formula's
        last computed result, so this is how the shipped fit parameters are
        read without Excel.
        """
        with zipfile.ZipFile(self.path) as zf:
            root = etree.fromstring(zf.read(sheet_part))
        values: dict[str, str] = {}
        for cell in root.iter(f"{_WB}c"):
            value = cell.find(f"{_WB}v")
            if value is not None and value.text is not None:
                values[cell.get("r")] = value.text
        return values

    @property
    def sheet_parts_by_index(self) -> list[str]:
        """Worksheet zip members in workbook order.

        ``localSheetId`` on a <definedName> is an index into this ordering,
        which is how a sheet-scoped name is tied back to the part its cells
        live in.
        """
        rels = {rel.get("Id"): rel.get("Target") for rel in self.workbook_rels_root}
        parts: list[str] = []
        for sheet in self.workbook_root.findall(f".//{_WB}sheets/{_WB}sheet"):
            rid = sheet.get(f"{{{OFFICE_RELS_NS}}}id")
            parts.append(posixpath.normpath(posixpath.join("xl", rels[rid])))
        return parts

    def cached_values_by_cell(self, sheet_part: str) -> list[tuple[str, str]]:
        """Return [(A1 address, text)] for every <v>/<t> in one worksheet.

        ``worksheet_cached_values`` flattens the same text but drops the
        address, which cannot answer "which column is this in" — the question
        a column-scoped exemption has to ask.
        """
        with zipfile.ZipFile(self.path) as zf:
            root = etree.fromstring(zf.read(sheet_part))
        pairs: list[tuple[str, str]] = []
        for cell in root.iter(f"{_WB}c"):
            ref = cell.get("r") or ""
            for element in cell.iter():
                if element.tag in (f"{_WB}v", f"{_WB}t", f"{_WB}is"):
                    for text in element.itertext():
                        if text:
                            pairs.append((ref, text))
        return pairs

    def na_masked_columns(self) -> dict[str, set[str]]:
        """Return {worksheet part: {column letters allowed to cache #N/A}}.

        Derived from the sheet-scoped RegChart overlay ranges rather than from
        a literal column letter, so this tracks the writer's layout and covers
        every Regression-templated sheet in the workbook automatically.
        """
        parts = self.sheet_parts_by_index
        exempt: dict[str, set[str]] = {}
        containers = self.workbook_root.findall(_DEFINED_NAMES_TAG)
        if not containers:
            return exempt
        for name_element in containers[0].findall(_DEFINED_NAME_TAG):
            if name_element.get("name") not in _NA_MASKED_RANGE_NAMES:
                continue
            local_sheet_id = name_element.get("localSheetId")
            if local_sheet_id is None:
                continue
            index = int(local_sheet_id)
            if not 0 <= index < len(parts):
                continue
            match = _OFFSET_ANCHOR_RE.search((name_element.text or "").strip())
            if match is None:
                continue
            exempt.setdefault(parts[index], set()).add(match.group(1).upper())
        return exempt

    def chart_sheet_references(self) -> dict[str, tuple[str, set[str]]]:
        """Return {chart part: (sheet the chart sits on, sheet names it references)}.

        A chart is its own package part, not sheet content, so every reference
        inside it has to name a sheet — which is exactly why a wrong name still
        parses and silently reads another sheet's numbers. Both formula kinds
        are collected: the `<c:f>` of a SERIES and the `<c15:f>` of the
        `datalabelsRange` extension that backs "Value From Cells" labels. They
        live in different namespaces, so the scan matches on the local name `f`
        and picks up any future extension for free.

        Which sheet owns a chart is a relationship walk, worksheet ->
        `<drawing r:id>` -> drawing rels -> chart part; a chart part no drawing
        claims is not in the result at all.
        """
        with zipfile.ZipFile(self.path) as zf:

            def _rels_for(part: str) -> dict[str, str]:
                folder, base = posixpath.split(part)
                rels_part = posixpath.join(folder, "_rels", base + ".rels")
                try:
                    data = zf.read(rels_part)
                except KeyError:
                    return {}
                return {
                    rel.get("Id") or "": rel.get("Target") or ""
                    for rel in etree.fromstring(data).findall(_RELATIONSHIP_TAG)
                }

            def _resolve(base_part: str, target: str) -> str:
                if target.startswith("/"):
                    return target.lstrip("/")
                return posixpath.normpath(
                    posixpath.join(posixpath.dirname(base_part), target)
                )

            sheet_names = [
                sheet.get("name") or ""
                for sheet in self.workbook_root.findall(f".//{_WB}sheets/{_WB}sheet")
            ]
            references: dict[str, tuple[str, set[str]]] = {}
            for sheet_name, sheet_part in zip(sheet_names, self.sheet_parts_by_index):
                try:
                    sheet_root = etree.fromstring(zf.read(sheet_part))
                except KeyError:
                    continue
                sheet_rels = _rels_for(sheet_part)
                for drawing in sheet_root.iter(f"{_WB}drawing"):
                    rid = drawing.get(f"{{{OFFICE_RELS_NS}}}id")
                    if rid is None or rid not in sheet_rels:
                        continue
                    drawing_part = _resolve(sheet_part, sheet_rels[rid])
                    for target in _rels_for(drawing_part).values():
                        chart_part = _resolve(drawing_part, target)
                        if "/charts/" not in chart_part:
                            continue
                        try:
                            chart_root = etree.fromstring(zf.read(chart_part))
                        except KeyError:
                            continue
                        prefixes: set[str] = set()
                        for element in chart_root.iter():
                            if etree.QName(element).localname != "f":
                                continue
                            text = (element.text or "").strip()
                            if "!" not in text:
                                continue
                            prefixes.add(text.split("!", 1)[0].strip("=").strip("'"))
                        references[chart_part] = (sheet_name, prefixes)
        return references

    def worksheet_cached_values(self, sheet_part: str) -> list[str]:
        """Return the text of every <v> and <t> element in the given worksheet."""
        with zipfile.ZipFile(self.path) as zf:
            data = zf.read(sheet_part)
        root = etree.fromstring(data)
        text_pieces: list[str] = []
        for element in root.iter():
            tag = element.tag
            if tag in (f"{_WB}v", f"{_WB}t", f"{_WB}is"):
                # <is><t> is the inline-string container; pull every <t> descendant
                for child_text in element.itertext():
                    if child_text:
                        text_pieces.append(child_text)
        return text_pieces


# ---------------------------------------------------------------------------
# Invariant assertions, exposed as plain functions so the fault-injection
# cases can reuse the exact same assertion with a precise failure message.
# ---------------------------------------------------------------------------


def _assert_no_orphan_named_ranges(package: WorkbookPackage) -> None:
    root = package.workbook_root
    defined_names_elements = root.findall(_DEFINED_NAMES_TAG)
    assert len(defined_names_elements) == 1, (
        f"workbook.xml must have exactly one <definedNames> container, "
        f"found {len(defined_names_elements)}"
    )
    defined_names = defined_names_elements[0]
    seen_workbook: set[str] = set()
    seen_local: set[tuple[str, str]] = set()
    offenders: list[str] = []
    for name_element in defined_names.findall(_DEFINED_NAME_TAG):
        name = name_element.get("name")
        body = (name_element.text or "").strip()
        local_sheet_id = name_element.get("localSheetId")
        if not name:
            offenders.append("<definedName> missing name attribute")
            continue
        if not body:
            offenders.append(f"<definedName name={name!r}> has empty body")
        if not _VALID_NAME_RE.match(name):
            offenders.append(
                f"<definedName name={name!r}> contains invalid characters"
            )
        if local_sheet_id is None:
            if name in seen_workbook:
                offenders.append(f"duplicate workbook-scoped definedName: {name!r}")
            seen_workbook.add(name)
        else:
            key = (name, local_sheet_id)
            if key in seen_local:
                offenders.append(
                    f"duplicate sheet-scoped definedName: {name!r} localSheetId={local_sheet_id!r}"
                )
            seen_local.add(key)
    assert not offenders, "Defined-name violations:\n  " + "\n  ".join(offenders)


def _assert_charts_reference_only_their_own_sheet(package: WorkbookPackage) -> None:
    """No chart may read data from a sheet other than the one it sits on.

    Every Regression-templated sheet carries its own copy of the sheet-scoped
    `RegChart*` ranges, and each sheet's charts are written against that
    sheet's name. A chart whose reference names a *different* sheet is the
    silent-wrong-answer failure: the formula still resolves, the chart still
    draws, and it plots another model's numbers. In the pre-reunify workbook
    this is precisely what a promoted-to-workbook-scope name did.
    """
    offenders: list[str] = []
    for chart_part, (owner, prefixes) in sorted(package.chart_sheet_references().items()):
        foreign = prefixes - {owner}
        if foreign:
            offenders.append(
                f"{chart_part} sits on {owner!r} but references "
                + ", ".join(repr(name) for name in sorted(foreign))
            )
    assert not offenders, (
        "Charts referencing another sheet's data:\n  " + "\n  ".join(offenders)
    )


def _assert_no_error_literals_in_cached_values(package: WorkbookPackage) -> None:
    """No <v> or <t> element contains an Excel error literal.

    Formulas <f> legitimately reference NA() inside IFERROR wrappers, so they
    are excluded from the scan.

    One class of cached #N/A is legitimate and is exempted per column: the
    NA()-masked overlay columns that drive selective chart data labels, located
    per sheet via ``na_masked_columns``. Every other literal, in every other
    column, still fails.
    """
    namelist = package.namelist
    worksheet_parts = sorted(
        n for n in namelist if n.startswith("xl/worksheets/") and n.endswith(".xml")
    )
    exempt_columns = package.na_masked_columns()
    offenders: list[tuple[str, str, str]] = []
    for part in worksheet_parts:
        masked = exempt_columns.get(part, set())
        for ref, value in package.cached_values_by_cell(part):
            for literal in EXCEL_ERROR_LITERALS:
                if literal not in value:
                    continue
                if literal == _NA_LITERAL and masked:
                    column_match = _CELL_COLUMN_RE.match(ref)
                    if column_match and column_match.group(1) in masked:
                        continue
                offenders.append((f"{part}!{ref}" if ref else part, literal, value))
    assert not offenders, (
        "Cached-value error literals detected (these are post-recalc cached values; "
        "they indicate the build produced a broken cell, not a documented formula):\n  "
        + "\n  ".join(f"{part}: literal {lit!r} in {val!r}" for part, lit, val in offenders)
    )


def _assert_workbook_scope_belongs_to_the_catalog(package: WorkbookPackage) -> None:
    """Workbook scope holds catalog names and Excel's ``_xlnm.*`` names, nothing else.

    Every range the sheet writers create (``RegChart*``, ``UV_*``, the spec
    wiring) is sheet-scoped. A workbook-scoped entry outside the catalog is
    therefore residue — most damagingly the other artifact's ranges, which the
    v3.0 split left pointing at ``#REF!`` in both workbooks.
    ``workbook_builder.sync_workbook_names`` clears them on every build.
    """
    catalog_names = {definition.name for definition in _catalog_workbook_names()}
    offenders = [
        name_element.get("name")
        for name_element in package.workbook_root.iter(_DEFINED_NAME_TAG)
        if name_element.get("localSheetId") is None
        and not (name_element.get("name") or "").startswith("_xlnm.")
        and name_element.get("name") not in catalog_names
    ]
    assert not offenders, (
        "Workbook-scoped defined names that are not catalog functions "
        f"({len(offenders)}): {sorted(offenders)!r}. These are residue from a "
        "previous build; sync_workbook_names should have cleared them."
    )


def _assert_no_error_literals_in_defined_names(package: WorkbookPackage) -> None:
    """No defined-name body, at either scope, contains an Excel error literal.

    ``#N/A`` is excluded: the catalog's LAMBDA bodies legitimately call ``NA()``
    and several return it by design. A ``#REF!`` inside an ``OFFSET(...)`` body
    is the residue signature this check exists for.
    """
    offenders: list[str] = []
    for name_element in package.workbook_root.iter(_DEFINED_NAME_TAG):
        body = name_element.text or ""
        for literal in EXCEL_ERROR_LITERALS:
            if literal == "#N/A":
                continue
            if literal in body:
                offenders.append(f"{name_element.get('name')!r}: {literal}")
    assert not offenders, (
        "Defined names with error literals in their body:\n  " + "\n  ".join(offenders)
    )


def _assert_no_external_links(package: WorkbookPackage) -> None:
    """The artifact links to no other workbook.

    Both artifacts are self-contained. An external link here is never
    deliberate: it is Excel rebinding a reference to a sheet this workbook does
    not have (``Regression!Source_Data`` -> ``[1]!Source_Data``), and it makes
    Excel prompt about broken links on every open.
    """
    external_parts = sorted(
        name for name in package.namelist if name.startswith("xl/externalLinks/")
    )
    assert not external_parts, f"External link parts present: {external_parts!r}"

    external_references = package.workbook_root.findall(f"{_WB}externalReferences")
    assert not external_references, (
        "workbook.xml declares <externalReferences>; the artifact must not link "
        "to another workbook."
    )


def _assert_sheet_inventory(package: WorkbookPackage, expected: tuple[str, ...]) -> None:
    sheets = [s.get("name") for s in package.workbook_root.iter(_SHEET_TAG)]
    assert tuple(sheets) == expected, (
        f"Sheet inventory drifted: got {sheets!r}, expected {expected!r}. "
        f"If a sheet was added or removed, update the expected list."
    )


# ---------------------------------------------------------------------------
# Fixtures: synthetic (always) and real-workbook (for the opt-in class below)
# ---------------------------------------------------------------------------

ROOT_DIR = Path(__file__).resolve().parents[1]
# The shipped artifacts moved into dist/ in the Chunk 1 reorganization. These
# paths are what decides whether the real-workbook class runs or SKIPS, so a
# stale one does not fail — it silently drops every check in this file that
# reads a committed artifact, which is the whole of Layer 1 in CI.
DIST_DIR = ROOT_DIR / "dist"
REAL_WORKBOOK_PATH = DIST_DIR / "Lambda_Library.xlsx"
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"


def _catalog_workbook_names():
    """Return the catalog's workbook-scoped functions (pure Python, no xlwings)."""
    return load_catalog_document(CATALOG_PATH).workbook_functions


@pytest.fixture
def synthetic_workbook(tmp_path: Path) -> WorkbookPackage:
    return WorkbookPackage(build_headless_fixture(tmp_path))


@pytest.fixture
def post_sync_synthetic_workbook(tmp_path: Path) -> WorkbookPackage:
    """A synthetic fixture that has been through drop_workbook_names.

    This is the post-``workbook_builder.sync_workbook_names`` shape and is the
    place where the no-calcChain invariant must hold.
    """
    workbook_path = build_headless_fixture(tmp_path)
    # First, hand-stamp a calcChain into the fixture so the patcher has
    # something to strip. This mirrors what an unsynced build leaves behind.
    _inject_calc_chain(workbook_path)
    drop_workbook_names(workbook_path, {"Real_Name"})
    return WorkbookPackage(workbook_path)


@pytest.fixture(scope="module")
def real_workbook_package() -> WorkbookPackage | None:
    """Return a parsed view of the committed production workbook, or skip."""
    if not REAL_WORKBOOK_PATH.exists():
        pytest.skip(f"Real workbook not present at {REAL_WORKBOOK_PATH}")
    return WorkbookPackage(REAL_WORKBOOK_PATH)


def _inject_calc_chain(workbook_path: Path) -> None:
    """Add a calcChain part + Override + relationship to a workbook in place."""
    temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=workbook_path.parent)
    os.close(temp_fd)
    Path(temp_name).unlink(missing_ok=True)

    with zipfile.ZipFile(workbook_path, "r") as input_zip, zipfile.ZipFile(
        temp_name, "w", compression=zipfile.ZIP_DEFLATED
    ) as output_zip:
        for item in input_zip.infolist():
            data = input_zip.read(item.filename)
            if item.filename == "[Content_Types].xml":
                root = etree.fromstring(data)
                override = etree.SubElement(root, _OVERRIDE_TAG)
                override.set("PartName", "/xl/calcChain.xml")
                override.set(
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
                )
                data = (
                    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    + etree.tostring(root, encoding="UTF-8")
                )
            elif item.filename == "xl/_rels/workbook.xml.rels":
                root = etree.fromstring(data)
                rel = etree.SubElement(root, _RELATIONSHIP_TAG)
                rel.set("Id", "rIdCalcChain")
                rel.set(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
                )
                rel.set("Target", "calcChain.xml")
                data = (
                    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    + etree.tostring(root, encoding="UTF-8")
                )
            output_zip.writestr(item, data)
        output_zip.writestr("xl/calcChain.xml", b"<calcChain/>")

    Path(temp_name).replace(workbook_path)


# ---------------------------------------------------------------------------
# Always-on invariant tests (synthetic fixture)
# ---------------------------------------------------------------------------


def test_zipfile_is_valid(synthetic_workbook: WorkbookPackage) -> None:
    assert zipfile.ZipFile(synthetic_workbook.path).testzip() is None


def test_sheet_inventory_matches_synthetic(
    synthetic_workbook: WorkbookPackage,
) -> None:
    _assert_sheet_inventory(synthetic_workbook, _SYNTHETIC_SHEETS)


def test_no_orphan_named_ranges_synthetic(synthetic_workbook: WorkbookPackage) -> None:
    _assert_no_orphan_named_ranges(synthetic_workbook)


def test_no_error_literals_in_cached_values_synthetic(
    synthetic_workbook: WorkbookPackage,
) -> None:
    _assert_no_error_literals_in_cached_values(synthetic_workbook)


def test_post_sync_workbook_has_no_calc_chain(
    post_sync_synthetic_workbook: WorkbookPackage,
) -> None:
    """After drop_workbook_names runs, calcChain must be fully stripped.

    The patcher at workbook_builder.py:78-184 (sync_workbook_names) and
    workbook_builder.py:187-268 (drop_workbook_names) both:
      * remove the calcChain.xml part
      * drop the calcChain Override from [Content_Types].xml
      * drop the calcChain Relationship from xl/_rels/workbook.xml.rels
    """
    namelist = post_sync_synthetic_workbook.namelist
    assert "xl/calcChain.xml" not in namelist, (
        "calcChain part should be stripped by drop_workbook_names"
    )

    ct_root = post_sync_synthetic_workbook.content_types_root
    for override in ct_root.findall(_OVERRIDE_TAG):
        assert (override.get("PartName") or "").lower() != "/xl/calcchain.xml", (
            "calcChain Override in [Content_Types].xml should be stripped"
        )

    rels_root = post_sync_synthetic_workbook.workbook_rels_root
    for rel in rels_root.findall(_RELATIONSHIP_TAG):
        assert not (rel.get("Type") or "").endswith("/calcChain"), (
            f"calcChain Relationship should be stripped, found Id={rel.get('Id')!r}"
        )


# ---------------------------------------------------------------------------
# Fault-injection cases. These exercise the failure-detection path so we
# know the invariants above would actually catch a regression.
# ---------------------------------------------------------------------------


def test_fault_inject_invalid_defined_name_causes_no_orphan_named_ranges_to_fail(
    tmp_path: Path,
) -> None:
    workbook_path = build_headless_fixture(tmp_path)
    with zipfile.ZipFile(workbook_path) as zf:
        original = zf.read("xl/workbook.xml")
    root = etree.fromstring(original)
    defined_names = root.find(_DEFINED_NAMES_TAG)
    bad = etree.SubElement(defined_names, _DEFINED_NAME_TAG)
    bad.set("name", "Bad Name With Spaces")
    bad.text = "LAMBDA(x,x)"
    patched = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + etree.tostring(root, encoding="UTF-8")
    )
    _replace_member(workbook_path, "xl/workbook.xml", patched)

    package = WorkbookPackage(workbook_path)
    with pytest.raises(AssertionError, match="Defined-name violations"):
        _assert_no_orphan_named_ranges(package)


def test_fault_inject_ref_literal_in_cached_value_causes_check_to_fail(
    tmp_path: Path,
) -> None:
    workbook_path = build_headless_fixture(tmp_path)
    bad_sheet = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<sheetData>"
        b'<row r="1"><c r="A1"><v>#REF!</v></c></row>'
        b"</sheetData>"
        b"</worksheet>"
    )
    _replace_member(workbook_path, "xl/worksheets/sheet1.xml", bad_sheet)

    package = WorkbookPackage(workbook_path)
    with pytest.raises(AssertionError, match="Cached-value error literals"):
        _assert_no_error_literals_in_cached_values(package)


def test_fault_inject_non_catalog_workbook_scoped_name_is_rejected(
    synthetic_workbook: WorkbookPackage,
) -> None:
    """The fixture's workbook-scoped ``Real_Name`` stands in for build residue."""
    with pytest.raises(AssertionError, match="not catalog functions"):
        _assert_workbook_scope_belongs_to_the_catalog(synthetic_workbook)


def test_fault_inject_ref_literal_in_defined_name_body_is_rejected(
    tmp_path: Path,
) -> None:
    """An ``OFFSET(#REF!,…)`` body is the cross-artifact residue signature."""
    workbook_path = build_headless_fixture(tmp_path)
    with zipfile.ZipFile(workbook_path) as zf:
        root = etree.fromstring(zf.read("xl/workbook.xml"))
    residue = etree.SubElement(root.find(_DEFINED_NAMES_TAG), _DEFINED_NAME_TAG)
    residue.set("name", "RegChartFitY")
    residue.text = "OFFSET(#REF!,1,0,MAX(IFERROR(#REF!,1),1),1)"
    _replace_member(
        workbook_path,
        "xl/workbook.xml",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + etree.tostring(root, encoding="UTF-8"),
    )

    package = WorkbookPackage(workbook_path)
    with pytest.raises(AssertionError, match="error literals in their body"):
        _assert_no_error_literals_in_defined_names(package)


def test_fault_inject_external_link_is_rejected(tmp_path: Path) -> None:
    """An externalLink part means Excel will prompt about broken links."""
    workbook_path = build_headless_fixture(tmp_path)
    _replace_member(
        workbook_path, "xl/externalLinks/externalLink1.xml", b"<externalLink/>"
    )

    package = WorkbookPackage(workbook_path)
    with pytest.raises(AssertionError, match="External link parts present"):
        _assert_no_external_links(package)


def test_fault_inject_extra_sheet_causes_inventory_check_to_fail(
    tmp_path: Path,
) -> None:
    workbook_path = build_headless_fixture(tmp_path)
    with zipfile.ZipFile(workbook_path) as zf:
        original = zf.read("xl/workbook.xml")
    root = etree.fromstring(original)
    sheets = root.find(f"{_WB}sheets")
    extra = etree.SubElement(sheets, _SHEET_TAG)
    extra.set("name", "Unexpected_Sheet")
    extra.set("sheetId", "99")
    extra.set("{%s}id" % "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId99")
    patched = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + etree.tostring(root, encoding="UTF-8")
    )
    _replace_member(workbook_path, "xl/workbook.xml", patched)

    package = WorkbookPackage(workbook_path)
    with pytest.raises(AssertionError, match="Sheet inventory drifted"):
        _assert_sheet_inventory(package, _SYNTHETIC_SHEETS)


# ---------------------------------------------------------------------------
# The NA()-masked overlay column exemption.
#
# These build workbooks whose Regression-templated sheets each declare their
# own sheet-scoped RegChartCookDistFlag range, because that name — not a
# hard-coded column letter — is what the scan reads to decide which column may
# cache #N/A.
# ---------------------------------------------------------------------------


def _error_cells_sheet_xml(cells: dict[str, str]) -> bytes:
    """A worksheet whose given addresses hold cached error literals."""
    body = "".join(
        f'<row r="{"".join(ch for ch in ref if ch.isdigit())}">'
        f'<c r="{ref}" t="e"><v>{literal}</v></c></row>'
        for ref, literal in cells.items()
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<sheetData>" + body.encode("utf-8") + b"</sheetData>"
        b"</worksheet>"
    )


def _workbook_xml_with_flag_ranges(
    sheet_names: tuple[str, ...], flag_columns: dict[int, str]
) -> bytes:
    """Workbook XML declaring a sheet-scoped flag range per given sheet index.

    ``flag_columns`` maps sheet index to the anchor column its
    RegChartCookDistFlag range points at, mirroring what ``_setup_local_names``
    writes for each Regression-templated sheet.
    """
    sheet_elements = "".join(
        f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
        for i, name in enumerate(sheet_names)
    )
    names = "".join(
        f'<definedName name="RegChartCookDistFlag" localSheetId="{index}">'
        f"OFFSET({sheet_names[index]}!${column}$3,1,0,"
        f"MAX(IFERROR({sheet_names[index]}!$AB$9,1),1),1)"
        f"</definedName>"
        for index, column in sorted(flag_columns.items())
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        b"<sheets>" + sheet_elements.encode("utf-8") + b"</sheets>"
        b"<definedNames>" + names.encode("utf-8") + b"</definedNames>"
        b"</workbook>"
    )


def _build_masked_fixture(
    tmp_path: Path,
    *,
    sheet_names: tuple[str, ...],
    flag_columns: dict[int, str],
    cells_by_index: dict[int, dict[str, str]],
) -> WorkbookPackage:
    tmp_path.mkdir(parents=True, exist_ok=True)
    workbook_path = tmp_path / "masked.xlsx"
    with zipfile.ZipFile(workbook_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types(len(sheet_names)))
        zf.writestr(
            "xl/workbook.xml", _workbook_xml_with_flag_ranges(sheet_names, flag_columns)
        )
        zf.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheet_names)))
        for index in range(len(sheet_names)):
            cells = cells_by_index.get(index)
            zf.writestr(
                f"xl/worksheets/sheet{index + 1}.xml",
                _error_cells_sheet_xml(cells) if cells else _blank_sheet_xml(),
            )
    return WorkbookPackage(workbook_path)


def test_na_in_the_masked_overlay_column_is_not_an_offender(tmp_path: Path) -> None:
    """The non-flagged rows of the Cook's D overlay cache #N/A by design."""
    package = _build_masked_fixture(
        tmp_path,
        sheet_names=("Regression",),
        flag_columns={0: "AY"},
        cells_by_index={0: {"AY3": "#N/A", "AY4": "#N/A", "AY5": "#N/A"}},
    )
    _assert_no_error_literals_in_cached_values(package)


def test_na_outside_the_masked_column_still_fails(tmp_path: Path) -> None:
    package = _build_masked_fixture(
        tmp_path,
        sheet_names=("Regression",),
        flag_columns={0: "AY"},
        cells_by_index={0: {"AX3": "#N/A"}},
    )
    with pytest.raises(AssertionError, match=r"AX3"):
        _assert_no_error_literals_in_cached_values(package)


def test_non_na_literal_in_the_masked_column_still_fails(tmp_path: Path) -> None:
    """The exemption is for #N/A only — a broken cell there is still broken."""
    package = _build_masked_fixture(
        tmp_path,
        sheet_names=("Regression",),
        flag_columns={0: "AY"},
        cells_by_index={0: {"AY3": "#REF!"}},
    )
    with pytest.raises(AssertionError, match=r"#REF!"):
        _assert_no_error_literals_in_cached_values(package)


def test_the_exemption_follows_the_named_range_when_the_column_moves(
    tmp_path: Path,
) -> None:
    """The whole point of reading the name: insert a column, and the exemption
    moves with the overlay instead of staying behind on a stale letter."""
    package = _build_masked_fixture(
        tmp_path,
        sheet_names=("Regression",),
        flag_columns={0: "BB"},
        cells_by_index={0: {"BB3": "#N/A"}},
    )
    _assert_no_error_literals_in_cached_values(package)

    stale = _build_masked_fixture(
        tmp_path / "stale",
        sheet_names=("Regression",),
        flag_columns={0: "BB"},
        cells_by_index={0: {"AY3": "#N/A"}},
    )
    with pytest.raises(AssertionError, match=r"AY3"):
        _assert_no_error_literals_in_cached_values(stale)


def test_every_regression_templated_sheet_gets_its_own_exemption(
    tmp_path: Path,
) -> None:
    """The test-model workbook is 47 Regression-shaped sheets, each with its own
    sheet-scoped copy of the range. The exemption is per sheet, not per
    workbook, and a sheet that declares no such range gets none."""
    package = _build_masked_fixture(
        tmp_path,
        sheet_names=("M05 Log-Log NA Masking", "G10 Two Sequence Flags", "Mileage Data"),
        flag_columns={0: "AY", 1: "BB"},
        cells_by_index={
            0: {"AY3": "#N/A"},
            1: {"BB3": "#N/A"},
            2: {"C7": "#N/A"},
        },
    )
    with pytest.raises(AssertionError) as excinfo:
        _assert_no_error_literals_in_cached_values(package)

    message = str(excinfo.value)
    # Only the sheet with no flag range is reported; the other two are exempt
    # at their own, different columns.
    assert "sheet3.xml!C7" in message
    assert "sheet1.xml" not in message
    assert "sheet2.xml" not in message


def test_one_sheets_exemption_does_not_leak_to_another(tmp_path: Path) -> None:
    """Sheet 1's column being exempt must not exempt the same column on sheet 2."""
    package = _build_masked_fixture(
        tmp_path,
        sheet_names=("Regression", "Mileage Data"),
        flag_columns={0: "AY"},
        cells_by_index={1: {"AY3": "#N/A"}},
    )
    with pytest.raises(AssertionError, match=r"sheet2\.xml!AY3"):
        _assert_no_error_literals_in_cached_values(package)


# ---------------------------------------------------------------------------
# Charts must reference only the sheet they sit on.
#
# These fixtures wire the full ownership walk a real workbook uses — worksheet
# -> <drawing r:id> -> drawing rels -> chart part — because that walk is what
# decides which sheet a chart is allowed to name.
# ---------------------------------------------------------------------------

_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_CHART_EXT_NS = "http://schemas.microsoft.com/office/drawing/2012/chart"


def _sheet_xml_with_drawing() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:r="' + OFFICE_RELS_NS.encode("utf-8") + b'">'
        b"<sheetData/><drawing r:id=\"rId1\"/>"
        b"</worksheet>"
    )


def _chart_xml(series_ref: str, label_range_ref: str) -> bytes:
    """A chart part carrying one SERIES formula and one Value-From-Cells range.

    The two references sit in different namespaces on purpose: `c:f` is the
    series, `c15:f` inside the `datalabelsRange` extension is what Excel writes
    for "Value From Cells" labels. A scan that only knew the `c` namespace
    would pass this fixture while the label range pointed anywhere at all.
    """
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace xmlns:c="{_CHART_NS}" xmlns:c15="{_CHART_EXT_NS}">'
        f"<c:ser><c:val><c:numRef><c:f>{series_ref}</c:f></c:numRef></c:val>"
        f"<c:extLst><c:ext><c15:datalabelsRange>"
        f"<c15:f>{label_range_ref}</c15:f>"
        f"</c15:datalabelsRange></c:ext></c:extLst></c:ser>"
        f"</c:chartSpace>"
    ).encode("utf-8")


def _rels_xml(pairs: list[tuple[str, str, str]]) -> bytes:
    body = "".join(
        f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>'
        for rid, rel_type, target in pairs
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + body.encode("utf-8")
        + b"</Relationships>"
    )


def _build_chart_fixture(
    tmp_path: Path, *, series_ref: str, label_range_ref: str
) -> WorkbookPackage:
    """A 2-sheet workbook with one chart, anchored to the first sheet."""
    tmp_path.mkdir(parents=True, exist_ok=True)
    workbook_path = tmp_path / "charts.xlsx"
    sheet_names = ("Regression", "Mileage Data")
    drawing_type = f"{OFFICE_RELS_NS}/drawing"
    chart_type = f"{OFFICE_RELS_NS}/chart"
    with zipfile.ZipFile(workbook_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types(len(sheet_names)))
        zf.writestr("xl/workbook.xml", _workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheet_names)))
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml_with_drawing())
        zf.writestr("xl/worksheets/sheet2.xml", _blank_sheet_xml())
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            _rels_xml([("rId1", drawing_type, "../drawings/drawing1.xml")]),
        )
        zf.writestr("xl/drawings/drawing1.xml", b"<?xml version='1.0'?><drawing/>")
        zf.writestr(
            "xl/drawings/_rels/drawing1.xml.rels",
            _rels_xml([("rId1", chart_type, "../charts/chart1.xml")]),
        )
        zf.writestr("xl/charts/chart1.xml", _chart_xml(series_ref, label_range_ref))
    return WorkbookPackage(workbook_path)


def test_a_chart_referencing_its_own_sheet_passes(tmp_path: Path) -> None:
    package = _build_chart_fixture(
        tmp_path,
        series_ref="Regression!RegChartCookDist",
        label_range_ref="Regression!RegChartCookDistFlag",
    )
    assert package.chart_sheet_references() == {
        "xl/charts/chart1.xml": ("Regression", {"Regression"})
    }
    _assert_charts_reference_only_their_own_sheet(package)


def test_a_chart_series_naming_another_sheet_fails(tmp_path: Path) -> None:
    package = _build_chart_fixture(
        tmp_path,
        series_ref="'Mileage Data'!RegChartCookDist",
        label_range_ref="Regression!RegChartCookDistFlag",
    )
    with pytest.raises(AssertionError, match="Mileage Data"):
        _assert_charts_reference_only_their_own_sheet(package)


def test_a_value_from_cells_range_naming_another_sheet_fails(tmp_path: Path) -> None:
    """The data-label range is an extension element, not a SERIES formula.

    It is the reference most easily left pointing at the template sheet — it is
    written once, by name, and nothing about a wrong sheet shows up as an
    error. This is the case that must not slip through.
    """
    package = _build_chart_fixture(
        tmp_path,
        series_ref="Regression!RegChartCookDist",
        label_range_ref="'Mileage Data'!RegChartCookDistFlag",
    )
    with pytest.raises(AssertionError, match="Mileage Data"):
        _assert_charts_reference_only_their_own_sheet(package)


def test_masked_column_matches_the_writers_column_constant() -> None:
    """Cross-check the artifact against the writer.

    The scan reads the workbook's own name, so it cannot notice if the writer
    and the built artifact disagree. This is the assertion that would.
    """
    from lambda_catalog.workbook_helpers import col_letter
    from lambda_catalog.write_sheet_regression import _C_AY

    if not REAL_WORKBOOK_PATH.exists():
        pytest.skip(f"Real workbook not present at {REAL_WORKBOOK_PATH}")

    package = WorkbookPackage(REAL_WORKBOOK_PATH)
    exempt = package.na_masked_columns()
    assert exempt, "the Regression artifact declares no RegChartCookDistFlag range"
    for part, columns in exempt.items():
        assert columns == {col_letter(_C_AY)}, part


def _replace_member(workbook_path: Path, member_name: str, new_bytes: bytes) -> None:
    """Replace a single member inside an .xlsx zipfile in place."""
    temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=workbook_path.parent)
    os.close(temp_fd)
    Path(temp_name).unlink(missing_ok=True)

    with zipfile.ZipFile(workbook_path, "r") as input_zip, zipfile.ZipFile(
        temp_name, "w", compression=zipfile.ZIP_DEFLATED
    ) as output_zip:
        replaced = False
        for item in input_zip.infolist():
            data = new_bytes if item.filename == member_name else input_zip.read(item.filename)
            if item.filename == member_name:
                replaced = True
            output_zip.writestr(item, data)
        if not replaced:
            output_zip.writestr(member_name, new_bytes)

    Path(temp_name).replace(workbook_path)


# ---------------------------------------------------------------------------
# Always-on real-workbook tests. Unlike the opt-in class below, these run on
# every commit against the committed unified workbook: they are the
# regression guard for the pre-reunify v3.0 split's cross-artifact leakage
# (one artifact's chart ranges promoted to workbook scope in the other file,
# and the external link a Regression-coupled catalog LAMBDA produced in the
# Univariate workbook). Both Regression and Univariate sheets now live in the
# one artifact, so a single set of checks covers it.
# ---------------------------------------------------------------------------


class TestRealWorkbookNameScope:
    """Workbook scope and link hygiene for the committed unified workbook."""

    def test_workbook_scope_belongs_to_the_catalog(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        _assert_workbook_scope_belongs_to_the_catalog(real_workbook_package)

    def test_defined_names_have_no_error_literals(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        _assert_no_error_literals_in_defined_names(real_workbook_package)

    def test_workbook_has_no_external_links(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        _assert_no_external_links(real_workbook_package)


# ---------------------------------------------------------------------------
# Shipped-Univariate checks. Always on, like every other committed-artifact
# class in this file — these two guard the one thing no automated check covered
# before: whether the workbook we ship is the one the current writer produces,
# and whether its fits are right.
#
# The spec-driven verifier (lambda_catalog.deep_verify.verify_test_sheets, Excel required) reads
# only the descriptive statistics and the three histogram blocks — it never
# touches the fitting table. So a passing `build_production.py --verify` says
# nothing about the distribution fits. A built .xlsx caches every formula's
# last computed result, which is what lets these run headlessly in CI.
# ---------------------------------------------------------------------------

# How far above the reference MLE's NLL a shipped fit may sit. The two-stage
# search is a grid, so it lands near the optimum, not on it: measured 0.091 for
# Weibull and 0.0025 for Gamma on the shipped sample, and never worse than 0.09
# across simulated Weibull shapes 0.4-30 and Gamma shapes 0.3-200. The failure
# this guards against is coarser by two orders of magnitude — the pre-shrink 2-D
# Gamma grid shipped a fit 6.8 NLL above the optimum, worth ~13.7 of AIC.
_FIT_NLL_TOLERANCE = 0.5

# The IFERROR sentinel every NLL LAMBDA falls back to on undefined likelihood.
_NLL_SENTINEL = 1e15

UNIVARIATE_SHEET = "Univariate"
LIFE_EXPECTANCY_CSV = (
    Path(__file__).resolve().parents[1] / "sample_data" / "Life Expectancy Data.csv"
)


def _life_expectancy_column() -> list[float | None]:
    """The dataset column the Univariate sheet fits.

    Same three steps as ``tools/inspect_univariate_sheet._load_source_data``,
    off the same ``load_csv_rows`` / ``LIFE_EXPECTANCY`` pair, so this test and
    the spec verifier cannot disagree about which rows are in the sample. Kept
    inline rather than importing that module because it pulls in xlwings, which
    these headless invariants deliberately never touch.
    """
    from lambda_catalog.write_sheet_csv_dataset import LIFE_EXPECTANCY, load_csv_rows

    headers, rows = load_csv_rows(LIFE_EXPECTANCY_CSV, LIFE_EXPECTANCY)
    column = headers.index("Life expectancy")
    return [row[column] if column < len(row) else None for row in rows]


def _defined_name_bodies(package: WorkbookPackage) -> list[tuple[str, str]]:
    """Every defined name in the package as (name, body) pairs."""
    return [
        (element.get("name") or "", (element.text or "").strip())
        for element in package.workbook_root.iter(_DEFINED_NAME_TAG)
    ]


def _shipped_fit_row(distribution: str) -> int:
    """Row of a distribution in the fitting table, from the writer's own spec."""
    from lambda_catalog.write_sheet_univariate import _ROW_DIST_START, _dist_rows

    for row, name, *_ in _dist_rows(_ROW_DIST_START):
        if name == distribution:
            return row
    raise KeyError(f"{distribution!r} is not a row in the fitting table")


def _shipped_cells(
    package: WorkbookPackage, distribution: str, columns: tuple[int, ...]
) -> tuple[float, ...]:
    """Read cached values from one fitting-table row.

    Addresses come from the writer's column constants, so a column move breaks
    this loudly rather than silently reading the wrong cell.
    """
    from lambda_catalog.workbook_helpers import col_letter

    values = package.cell_values(package.sheet_part_for(UNIVARIATE_SHEET))
    row = _shipped_fit_row(distribution)
    out = []
    for col in columns:
        address = f"{col_letter(col)}{row}"
        raw = values.get(address)
        assert raw is not None, (
            f"{distribution} {address} has no cached value in the shipped workbook — "
            "it was saved without calculating, or the fitting table moved"
        )
        out.append(float(raw))
    return tuple(out)


def _shipped_two_parameter_fit(
    package: WorkbookPackage, distribution: str
) -> tuple[float, float, float]:
    """Return (theta1, theta2, NLL) for a two-parameter fit."""
    from lambda_catalog.write_sheet_univariate import _C_NLL, _C_T1_VAL, _C_T2_VAL

    theta1, theta2, nll = _shipped_cells(
        package, distribution, (_C_T1_VAL, _C_T2_VAL, _C_NLL)
    )
    return theta1, theta2, nll


def _shipped_nll(package: WorkbookPackage, distribution: str) -> float:
    """Return just the NLL, for rows whose parameter count varies.

    Exponential fills only theta1 and Triangular/BetaPERT fill three, so the
    NLL column is the one cell every row is guaranteed to have.
    """
    from lambda_catalog.write_sheet_univariate import _C_NLL

    return _shipped_cells(package, distribution, (_C_NLL,))[0]


class TestShippedUnivariateFits:
    """The fits in the committed artifact are the ones the data supports.

    Guards the search itself: a wrong profile formula, a bad bracket, or a
    boundary-pinned optimum all show up as an NLL well above the reference MLE.
    """

    def test_weibull_fit_is_at_the_maximum_likelihood(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        from lambda_catalog.analyze_univariate import mle_weibull, nll_weibull

        shape, scale, cached_nll = _shipped_two_parameter_fit(
            real_workbook_package, "Weibull"
        )
        data = _life_expectancy_column()

        assert 0.0 < shape < _NLL_SENTINEL and 0.0 < scale < _NLL_SENTINEL
        # The cached NLL cell must agree with the parameters beside it.
        assert cached_nll == pytest.approx(nll_weibull(data, shape, scale), rel=1e-9)

        best_shape, best_scale = mle_weibull(data)
        excess = nll_weibull(data, shape, scale) - nll_weibull(data, best_shape, best_scale)
        assert excess == pytest.approx(0.0, abs=_FIT_NLL_TOLERANCE), (
            f"shipped Weibull fit (k={shape:.4f}, lambda={scale:.4f}) sits {excess:.4f} "
            f"NLL above the MLE (k={best_shape:.4f}, lambda={best_scale:.4f})"
        )

    def test_gamma_fit_is_at_the_maximum_likelihood(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        from lambda_catalog.analyze_univariate import mle_gamma, nll_gamma

        shape, rate, cached_nll = _shipped_two_parameter_fit(
            real_workbook_package, "Gamma"
        )
        data = _life_expectancy_column()

        assert 0.0 < shape < _NLL_SENTINEL and 0.0 < rate < _NLL_SENTINEL
        assert cached_nll == pytest.approx(nll_gamma(data, shape, rate), rel=1e-9)

        best_shape, best_rate = mle_gamma(data)
        excess = nll_gamma(data, shape, rate) - nll_gamma(data, best_shape, best_rate)
        assert excess == pytest.approx(0.0, abs=_FIT_NLL_TOLERANCE), (
            f"shipped Gamma fit (alpha={shape:.4f}, beta={rate:.4f}) sits {excess:.4f} "
            f"NLL above the MLE (alpha={best_shape:.4f}, beta={best_rate:.4f})"
        )

    def test_beta_fit_is_finite_and_not_the_error_sentinel(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """Beta has no closed-form MLE to compare against, so this is the
        weaker invariant: the grid search returned real parameters rather than
        collapsing to the IFERROR sentinel."""
        alpha, beta, cached_nll = _shipped_two_parameter_fit(
            real_workbook_package, "Beta"
        )

        assert 0.0 < alpha < _NLL_SENTINEL and 0.0 < beta < _NLL_SENTINEL
        assert cached_nll < _NLL_SENTINEL, (
            f"shipped Beta NLL is the {_NLL_SENTINEL:g} sentinel — the grid search "
            "found no admissible parameter pair"
        )

    def test_every_fit_beats_the_worst_and_none_is_the_sentinel(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """No distribution row may ship the sentinel: it would rank first or
        last in the AIC/BIC comparison for the wrong reason."""
        for distribution in (
            "Normal", "Lognormal", "Exponential", "Weibull",
            "Gamma", "Triangular", "Beta", "BetaPERT",
        ):
            cached_nll = _shipped_nll(real_workbook_package, distribution)
            assert cached_nll < _NLL_SENTINEL, (
                f"{distribution} ships the {_NLL_SENTINEL:g} NLL sentinel"
            )


class TestShippedUnivariateLayout:
    """The committed artifact is the one the current writer produces.

    Pairs with TestShippedUnivariateFits: a stale workbook can still hold
    perfectly good fit values, so numeric correctness alone cannot catch a
    forgotten rebuild. Both sides of every assertion derive from the writer's
    zone table, so this compares artifact against writer rather than against a
    transcribed constant.
    """

    def test_fit_zone_bodies_match_the_writer(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        from lambda_catalog.workbook_helpers import col_letter
        from lambda_catalog.write_sheet_univariate import (
            _BAND_COL,
            _BETA_C_LABEL,
            _BETA_C_S1_A,
            _BETA_C_S1_B,
            _BETA_C_S2_A,
            _BETA_C_S2_B,
            _BETA_C_S2_NLL,
            _BETA_R_GRID_POINTS,
            _PR_C_LABEL,
            _PR_C_S1,
            _PR_C_S2,
            _PR_C_S2_NLL,
            _PR_R_GRID_POINTS,
            _R_BODY,
            _ROW_FIT_ZONE,
        )

        first = _ROW_FIT_ZONE + _R_BODY                       # row 33
        expected: dict[str, str] = {}

        # Weibull / Gamma: OFFSET bodies whose height tracks that stage's live
        # Grid Points cell (unsquared — a profile axis is N points, not N²).
        # Two stages side by side: S1 body/axis in cols 1/0, S2 in cols 3/2,
        # each stage's N cell on row 4 of its own value column.
        for name, zone in (("UV_WB", "weibull"), ("UV_GAMMA", "gamma")):
            c0 = _BAND_COL[zone]
            for suffix, off, n_off in (
                ("S1", _PR_C_S1, _PR_C_S1), ("S1_Axis", _PR_C_LABEL, _PR_C_S1),
                ("S2", _PR_C_S2_NLL, _PR_C_S2), ("S2_Axis", _PR_C_S2, _PR_C_S2),
            ):
                cl = col_letter(c0 + off)
                n_ref = f"${col_letter(c0 + n_off)}${_ROW_FIT_ZONE + _PR_R_GRID_POINTS}"
                expected[f"{name}_{suffix}"] = (
                    f"OFFSET(Univariate!${cl}${first},0,0,"
                    f"MAX(IFERROR(Univariate!{n_ref},1),1),1)"
                )

        # Beta: OFFSET bodies whose height tracks the live Grid Points cell
        # (^2). Three names per stage (Alpha / Beta / NLL); S1's N cell is the
        # S1 α column (row 4), S2's is the S2 α column.
        beta = _BAND_COL["beta"]
        n_row = _ROW_FIT_ZONE + _BETA_R_GRID_POINTS             # row 4
        for stage, a_off, b_off, nll_off, n_off in (
            ("S1", _BETA_C_LABEL, _BETA_C_S1_A, _BETA_C_S1_B, _BETA_C_S1_A),
            ("S2", _BETA_C_S2_A, _BETA_C_S2_B, _BETA_C_S2_NLL, _BETA_C_S2_A),
        ):
            n_ref = f"${col_letter(beta + n_off)}${n_row}"
            for suffix, body_off in (("Alpha", a_off), ("Beta", b_off), ("NLL", nll_off)):
                cl = col_letter(beta + body_off)
                expected[f"UV_BETA_{stage}_{suffix}"] = (
                    f"OFFSET(Univariate!${cl}${first},0,0,"
                    f"MAX(IFERROR(Univariate!{n_ref},1),1)^2,1)"
                )

        actual = {
            name.split("!", 1)[-1]: body
            for name, body in _defined_name_bodies(real_workbook_package)
        }
        mismatched = {
            key: (want, actual.get(key))
            for key, want in expected.items()
            if actual.get(key) != want
        }
        assert not mismatched, (
            "the committed Lambda_Library.xlsx's Univariate sheet does not match "
            "the current writer — rebuild it with `python scripts/build_production.py "
            f"--verify --no-launch` on a machine with Excel and commit it. {mismatched}"
        )


# ---------------------------------------------------------------------------
# Real-workbook tests; always on, against the committed artifact
# ---------------------------------------------------------------------------


class TestRealWorkbook:
    """Structural checks against the committed ``dist/Lambda_Library.xlsx``.

    These answer the one question the synthetic fixture cannot: is the file we
    ship the file the writers produce? They were opt-in for as long as a known
    bad artifact sat committed; that artifact is rebuilt, so they run on every
    push. Nothing here needs Excel — the checks read the .xlsx as a zip.

    A failure here means the committed workbook is stale or was saved by hand.
    The fix is to rebuild and commit it, not to relax the check:
    ``python scripts/build_production.py --verify --no-launch`` (needs Excel).
    """

    def test_sheet_inventory_matches_spec(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """Real workbook's sheet set must equal the documented production inventory."""
        _assert_sheet_inventory(real_workbook_package, EXPECTED_REAL_SHEETS)

    def test_no_orphan_named_ranges(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        _assert_no_orphan_named_ranges(real_workbook_package)

    def test_no_error_literals_in_cached_values(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        _assert_no_error_literals_in_cached_values(real_workbook_package)

    def test_content_types_consistent(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """[Content_Types].xml must reference every actual worksheet part, and vice versa."""
        namelist = real_workbook_package.namelist
        actual_worksheets = {
            f"/{name}"
            for name in namelist
            if name.startswith("xl/worksheets/") and name.endswith(".xml")
        }
        root = real_workbook_package.content_types_root
        declared_worksheets = {
            override.get("PartName")
            for override in root.findall(_OVERRIDE_TAG)
            if (override.get("ContentType") or "").endswith("worksheet+xml")
        }
        missing = actual_worksheets - declared_worksheets
        extra = declared_worksheets - actual_worksheets
        assert not missing and not extra, (
            f"Worksheet <-> Content_Types mismatch: missing={missing!r}, extra={extra!r}"
        )

    def test_workbook_rels_consistent(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """Every <Relationship> in xl/_rels/workbook.xml.rels must resolve to a zip member."""
        namelist = real_workbook_package.namelist
        rels = real_workbook_package.workbook_rels_root
        missing: list[tuple[str, str]] = []
        for rel in rels.findall(_RELATIONSHIP_TAG):
            target = rel.get("Target") or ""
            # Targets are relative to xl/; the workbook.xml.rels lives at xl/_rels/.
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = f"xl/{target}"
            if resolved not in namelist:
                missing.append((rel.get("Id", "?"), resolved))
        assert not missing, (
            "xl/_rels/workbook.xml.rels references parts not in the zip:\n  "
            + "\n  ".join(f"Id={rid!r} Target={t!r}" for rid, t in missing)
        )

    def test_charts_reference_only_their_own_sheet(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """Every chart in the shipped artifact reads its own sheet, and only it."""
        references = real_workbook_package.chart_sheet_references()
        assert references, "the artifact declares no charts to check"
        _assert_charts_reference_only_their_own_sheet(real_workbook_package)

    def test_chart_relationship_targets_resolve(
        self, real_workbook_package: WorkbookPackage
    ) -> None:
        """Every chart/drawing relationship Target must point to a real zip member."""
        namelist = real_workbook_package.namelist
        rels_parts = [
            name
            for name in namelist
            if name.startswith(("xl/drawings/_rels/", "xl/charts/_rels/"))
            and name.endswith(".rels")
        ]
        missing: list[tuple[str, str, str]] = []
        for rels_part in rels_parts:
            with zipfile.ZipFile(real_workbook_package.path) as zf:
                data = zf.read(rels_part)
            root = etree.fromstring(data)
            # The .rels file lives at e.g. xl/drawings/_rels/drawing1.xml.rels;
            # relative Targets are resolved against its parent directory.
            base_dir = rels_part.rsplit("/", 2)[0]  # strip "_rels/foo.rels"
            for rel in root.findall(_RELATIONSHIP_TAG):
                target = rel.get("Target") or ""
                resolved = (
                    target.lstrip("/")
                    if target.startswith("/")
                    else posixpath.normpath(f"{base_dir}/{target}")
                )
                if resolved not in namelist:
                    missing.append((rels_part, rel.get("Id", "?"), resolved))
        assert not missing, (
            "Chart/drawing relationship Targets don't resolve:\n  "
            + "\n  ".join(f"{part}: Id={rid!r} -> {t!r}" for part, rid, t in missing)
        )

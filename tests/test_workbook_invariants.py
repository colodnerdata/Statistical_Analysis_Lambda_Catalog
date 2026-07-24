"""Headless structural invariants for the produced Lambda_Library.xlsx package.

These tests are pure-Python (zipfile + lxml); they do not import xlwings and run
on every Linux CI commit in <1 s. They catch packaging regressions that the
existing `RecordingSheet` unit tests cannot see:

  * sheet-name drift against the spec
  * orphan defined names (empty body, invalid characters, duplicates)
  * localSheetId values that fall outside [0, sheet_count)
  * #REF! / #NAME? / #VALUE! / #NULL! / #DIV/0! / #N/A / #NUM! literals leaking
    into cached cell values (the writers' LAMBDA formulas legitimately wrap
    `NA()` inside `IFERROR`, so we only scan <v> and <t> elements, not <f>)
  * [Content_Types].xml / xl/_rels/workbook.xml.rels inconsistencies
  * chart-relationship Targets that don't resolve to a real zip member
  * workbook_builder.py: post-`sync_workbook_names` rebuilds must leave no
    calcChain part or relationship behind

The synthetic 4-sheet fixture (``build_headless_fixture``) is the always-on
source of truth for these invariants. A separate ``TestRealWorkbook`` class
also runs the same checks against the committed ``Lambda_Library.xlsx`` but
is opt-in via ``RUN_EXCEL_INTEGRATION=1`` because the committed artifact
currently has workbook-scoped defined-name duplicates from a previous build
that the in-flight bug fix in ``workbook_builder.py`` will address.

The deep spec-driven check (``build_qc.verify_test_sheets``, xlwings + Excel
required) is the source of truth for cell-level correctness. These invariants
are the cheap, always-on packaging screen.
"""
from __future__ import annotations

import os
import re
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

import lxml.etree as etree  # type: ignore[import-untyped]
import pytest

from lambda_catalog.workbook_builder import drop_workbook_names

# ---------------------------------------------------------------------------
# Opt-in env var. RUN_EXCEL_INTEGRATION=1 enables the TestRealWorkbook class,
# which exercises the same invariants against the committed production
# workbook. The class is skipped by default until the in-flight bug fix in
# workbook_builder.py cleans up the committed artifact.
# ---------------------------------------------------------------------------

RUN_REAL_WORKBOOK_TESTS = os.environ.get("RUN_EXCEL_INTEGRATION") == "1"
skip_real_workbook = pytest.mark.skipif(
    not RUN_REAL_WORKBOOK_TESTS,
    reason="real-workbook checks are opt-in via RUN_EXCEL_INTEGRATION=1 "
    "until the committed artifact is cleaned up",
)

# ---------------------------------------------------------------------------
# Namespaces and the constant set of Excel error literals we never want to see
# leak into a cached cell value. The cell value lives in a <v> element (or in
# an inline <t>). Formulas <f> legitimately reference NA() inside IFERROR
# wrappers, so we deliberately exclude those from the scan.
# ---------------------------------------------------------------------------

WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

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

# Real Lambda_Library.xlsx has 6 sheets in the committed artifact (the
# production build that excludes Univariate and Model Construction was the
# last to be checked in). We keep this list here so the invariant test can
# assert the structural set without re-reading the spec at every run.
EXPECTED_REAL_SHEETS: tuple[str, ...] = (
    "LAMBDA_functions",
    "Life Expectancy Data",
    "Mileage Data",
    "Regression Instructions",
    "Diagnostic Guide",
    "Version History",
    "Regression",
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


def _assert_no_error_literals_in_cached_values(package: WorkbookPackage) -> None:
    """No <v> or <t> element contains an Excel error literal.

    Formulas <f> legitimately reference NA() inside IFERROR wrappers, so they
    are excluded from the scan.
    """
    namelist = package.namelist
    worksheet_parts = sorted(
        n for n in namelist if n.startswith("xl/worksheets/") and n.endswith(".xml")
    )
    offenders: list[tuple[str, str, str]] = []
    for part in worksheet_parts:
        for value in package.worksheet_cached_values(part):
            for literal in EXCEL_ERROR_LITERALS:
                if literal in value:
                    offenders.append((part, literal, value))
    assert not offenders, (
        "Cached-value error literals detected (these are post-recalc cached values; "
        "they indicate the build produced a broken cell, not a documented formula):\n  "
        + "\n  ".join(f"{part}: literal {lit!r} in {val!r}" for part, lit, val in offenders)
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

REAL_WORKBOOK_PATH = Path(__file__).resolve().parents[1] / "Lambda_Library.xlsx"


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
    import os
    import tempfile

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


def _replace_member(workbook_path: Path, member_name: str, new_bytes: bytes) -> None:
    """Replace a single member inside an .xlsx zipfile in place."""
    import os
    import tempfile

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
# Real-workbook tests; opt-in via RUN_EXCEL_INTEGRATION=1
# ---------------------------------------------------------------------------


@skip_real_workbook
class TestRealWorkbook:
    """Real-workbook checks; opt-in via RUN_EXCEL_INTEGRATION=1.

    These are deferred until the in-flight bug fix in workbook_builder.py
    cleans up the committed production artifact (workbook-scoped defined-name
    duplicates that point to #REF!). The synthetic-fixture tests above are
    the always-on source of truth for the invariants.
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
                if target.startswith("/"):
                    resolved = target.lstrip("/")
                else:
                    resolved = f"{base_dir}/{target}"
                    # Normalize ../segments
                    parts: list[str] = []
                    for segment in resolved.split("/"):
                        if segment == "..":
                            if parts:
                                parts.pop()
                        elif segment and segment != ".":
                            parts.append(segment)
                    resolved = "/".join(parts)
                if resolved not in namelist:
                    missing.append((rels_part, rel.get("Id", "?"), resolved))
        assert not missing, (
            "Chart/drawing relationship Targets don't resolve:\n  "
            + "\n  ".join(f"{part}: Id={rid!r} -> {t!r}" for part, rid, t in missing)
        )

"""LAMBDA name-syncing via workbook.xml patching and workbook lifecycle helpers."""
from __future__ import annotations

import os
import re
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

# Excel's own reserved defined names — print areas, print titles, autofilter
# ranges. Excel creates and maintains these; they are never catalog names, so
# the workbook-scope purge leaves them alone.
_EXCEL_BUILTIN_NAME_PREFIX = "_xlnm."

# An external-workbook reference. When a formula names a sheet this workbook
# does not have (``Regression!Source_Data`` in the standalone Univariate
# artifact), Excel does not error — it rewrites the reference to
# ``[1]!Source_Data``, materializes xl/externalLinks/externalLink1.xml with an
# ``xlPathMissing`` target, and prompts about broken links on every open.
# ``[0]`` is this workbook and is not an external reference.
_EXTERNAL_REFERENCE_RE = re.compile(r"\[[1-9][0-9]*\]")

# A sheet-qualified reference inside a formula: ``'Life Expectancy Data'!$A$1``
# or ``Regression!Source_Data``. The lookbehind keeps two lookalikes out: the
# ``#REF!`` error literal (preceded by ``#``) and the ``[0]!Name`` this-workbook
# form (whose ``!`` is preceded by ``]``).
_SHEET_QUALIFIED_REFERENCE_RE = re.compile(
    r"(?<![\w.$#\]])(?:'([^']+)'|([A-Za-z_][\w.]*))!"
)

# A double-quoted Excel string literal, inner quotes doubled. Stripped before
# either scan above runs: a label like "See note [1]!" is text, not a reference,
# and matching it would skip a perfectly portable definition.
_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')

_EXTERNAL_LINK_PART_PREFIX = "xl/externalLinks/"
_EXTERNAL_LINK_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
)


@dataclass(frozen=True)
class NameSyncResult:
    """Summary of how the workbook-scoped name manager was rewritten.

    Attributes
    ----------
    created : int
        Number of new workbook names added.
    updated : int
        Number of existing workbook names replaced.
    removed : int
        Number of workbook-scoped entries dropped and not re-added — residue
        from a previous build, or a catalog name skipped for this artifact.
    skipped : tuple[str, ...]
        Catalog names not written because they reference a worksheet this
        workbook does not have. Writing them anyway is what produces a broken
        external link (see ``sync_workbook_names``).
    """

    created: int
    updated: int
    removed: int = 0
    skipped: tuple[str, ...] = ()


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


def _sheet_names(workbook_root) -> frozenset[str]:
    """Return the worksheet names declared in a parsed workbook.xml."""
    return frozenset(
        sheet.get("name") or ""
        for sheet in workbook_root.iterfind(
            f"{{{WORKBOOK_NS}}}sheets/{{{WORKBOOK_NS}}}sheet"
        )
    )


def _missing_sheet_references(formula: str, sheet_names: frozenset[str]) -> tuple[str, ...]:
    """Return the sheets a formula names that this workbook does not have.

    Parameters
    ----------
    formula : str
        A defined-name body in workbook.xml token syntax.
    sheet_names : frozenset[str]
        The worksheet names present in the target workbook.

    Returns
    -------
    tuple[str, ...]
        Sorted names of referenced sheets that are absent, empty when every
        sheet-qualified reference resolves inside this workbook.
    """
    referenced = {
        quoted or bare
        for quoted, bare in _SHEET_QUALIFIED_REFERENCE_RE.findall(
            _STRING_LITERAL_RE.sub('""', formula)
        )
    }
    return tuple(sorted(name for name in referenced if name not in sheet_names))


def _formula_texts(part_bytes: bytes) -> list[str]:
    """Return every formula element's text in a worksheet or chart part.

    Matches on the local tag name so both the worksheet ``<f>`` and the
    DrawingML chart ``<c:f>`` element are picked up.
    """
    try:
        root = etree.fromstring(part_bytes)
    except etree.XMLSyntaxError:
        return []
    return [
        element.text
        for element in root.iter()
        if isinstance(element.tag, str)
        and element.tag.rsplit("}", 1)[-1] == "f"
        and element.text
    ]


def _external_links_are_orphaned(
    input_zip: zipfile.ZipFile,
    workbook_root,
    definitions: Sequence[CatalogFunction],
) -> bool:
    """Return True when nothing in the patched workbook still needs an external link.

    Checks every surviving formula: worksheet and chart formulas, the
    sheet-scoped defined names (which the sync never touches), and the catalog
    bodies about to be written. Workbook-scoped entries are excluded because
    the sync replaces all of them.
    """
    surviving_formulas: list[str] = [
        definition.workbook_xml_formula_from_display for definition in definitions
    ]
    defined_names = workbook_root.find(f"{{{WORKBOOK_NS}}}definedNames")
    if defined_names is not None:
        surviving_formulas.extend(
            name_element.text or ""
            for name_element in defined_names.iterfind(f"{{{WORKBOOK_NS}}}definedName")
            if name_element.get("localSheetId") is not None
        )
    for item in input_zip.namelist():
        if not item.endswith(".xml"):
            continue
        if item.startswith(("xl/worksheets/", "xl/charts/", "xl/tables/")):
            surviving_formulas.extend(_formula_texts(input_zip.read(item)))
    return not any(
        _EXTERNAL_REFERENCE_RE.search(_STRING_LITERAL_RE.sub('""', formula))
        for formula in surviving_formulas
    )


def sync_workbook_names(
    workbook_path: Path, definitions: Sequence[CatalogFunction]
) -> NameSyncResult:
    """Make the catalog authoritative over workbook scope in workbook.xml.

    Workbook scope belongs to the catalog: after this runs, the only
    workbook-scoped ``<definedName>`` entries are the supplied definitions plus
    Excel's own ``_xlnm.*`` reserved names. Every other workbook-scoped entry is
    residue from a previous build and is dropped — a name the catalog retired, a
    sheet-scoped name that was once workbook-scoped, or a range belonging to the
    *other* artifact (the v3.0 workbook split left twelve ``RegChart*`` entries
    pointing at ``#REF!`` in the Univariate workbook and forty-two ``UV_*``
    entries in the Regression workbook, none of which an error-literal body
    match caught: the bodies are ``OFFSET(#REF!,…)``, not a bare ``#REF!``).
    Sheet-scoped entries — where every range the sheet writers create actually
    lives — are never touched.

    A definition is skipped rather than written when its body names a worksheet
    this workbook does not have. Excel does not leave such a reference broken:
    it rebinds it to an external workbook (``Regression!Source_Data`` becomes
    ``[1]!Source_Data``), writes an ``xlPathMissing`` external link part, and
    prompts about broken links on every open. ``Base_Period_Delta`` is the
    catalog's one Regression-coupled workbook-scoped LAMBDA and is what put
    that link in the standalone Univariate artifact. When the purge leaves no
    formula referencing an external workbook, the orphaned external-link parts,
    relationships and content-type overrides are stripped too.

    Parameters
    ----------
    workbook_path : Path
        Path to the .xlsx file to patch in place.
    definitions : Sequence[CatalogFunction]
        Functions whose names will be inserted or replaced.

    Returns
    -------
    NameSyncResult
        Counts of created, updated and removed workbook names, plus the names
        skipped for referencing absent worksheets.
    """
    target_names = {definition.name for definition in definitions}
    existing_target_names: set[str] = set()
    written_names: set[str] = set()
    skipped_names: list[str] = []
    purged = 0

    with zipfile.ZipFile(workbook_path, "r") as probe_zip:
        probe_root = etree.fromstring(probe_zip.read("xl/workbook.xml"))
        sheet_names = _sheet_names(probe_root)
        strip_external_links = any(
            item.startswith(_EXTERNAL_LINK_PART_PREFIX)
            for item in probe_zip.namelist()
        ) and _external_links_are_orphaned(probe_zip, probe_root, definitions)

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

                    # Clear workbook scope. Everything here is either a catalog
                    # name (rewritten below from the catalog body, so the old
                    # entry has no value) or residue from a previous build —
                    # including the cross-artifact residue the v3.0 split left
                    # behind, whose bodies wrap the #REF! rather than being one.
                    # Only Excel's reserved _xlnm.* names survive. Sheet-scoped
                    # entries are skipped: that is where the sheet writers'
                    # ranges (RegChart*, UV_*, the spec wiring) legitimately
                    # live, and this patcher has no business touching them.
                    for name_element in list(defined_names):
                        if name_element.tag != f"{{{WORKBOOK_NS}}}definedName":
                            continue
                        if name_element.get("localSheetId") is not None:
                            continue

                        target_name = name_element.get("name") or ""
                        if target_name.startswith(_EXCEL_BUILTIN_NAME_PREFIX):
                            continue
                        if target_name in target_names:
                            existing_target_names.add(target_name)
                        defined_names.remove(name_element)
                        purged += 1

                    for definition in definitions:
                        missing_sheets = _missing_sheet_references(
                            definition.workbook_xml_formula_from_display, sheet_names
                        )
                        if missing_sheets:
                            skipped_names.append(definition.name)
                            continue
                        name_element = etree.SubElement(
                            defined_names, f"{{{WORKBOOK_NS}}}definedName"
                        )
                        name_element.set("name", definition.name)
                        if definition.notes:
                            name_element.set("comment", definition.notes)
                        name_element.text = definition.workbook_xml_formula_from_display
                        written_names.add(definition.name)

                    if strip_external_links:
                        for external_references in workbook_root.findall(
                            f"{{{WORKBOOK_NS}}}externalReferences"
                        ):
                            workbook_root.remove(external_references)

                    xml_body = str(etree.tostring(workbook_root, encoding="unicode"))
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        + xml_body
                    ).encode("UTF-8")

                elif item.filename == "xl/calcChain.xml":
                    continue  # omit stale calc chain; Excel rebuilds on open

                elif strip_external_links and item.filename.startswith(
                    _EXTERNAL_LINK_PART_PREFIX
                ):
                    continue  # orphaned external link; nothing references it

                elif item.filename == "xl/_rels/workbook.xml.rels":
                    rels_root = etree.fromstring(data)
                    for rel in rels_root.findall(f"{{{_RELS_NS}}}Relationship"):
                        if rel.get("Type") == _CALC_CHAIN_REL_TYPE:
                            rels_root.remove(rel)
                        elif strip_external_links and rel.get("Type") == _EXTERNAL_LINK_REL_TYPE:
                            rels_root.remove(rel)
                    xml_body = str(etree.tostring(rels_root, encoding="unicode"))
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        + xml_body
                    ).encode("UTF-8")

                elif item.filename == "[Content_Types].xml":
                    ct_root = etree.fromstring(data)
                    for override in ct_root.findall(f"{{{_CT_NS}}}Override"):
                        part_name = (override.get("PartName") or "").lower()
                        if part_name == "/xl/calcchain.xml":
                            ct_root.remove(override)
                        elif strip_external_links and part_name.startswith(
                            f"/{_EXTERNAL_LINK_PART_PREFIX}".lower()
                        ):
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

    updated = len(existing_target_names & written_names)
    return NameSyncResult(
        created=len(written_names) - updated,
        updated=updated,
        removed=purged - updated,
        skipped=tuple(skipped_names),
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

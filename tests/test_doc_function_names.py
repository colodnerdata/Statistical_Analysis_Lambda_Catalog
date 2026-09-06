"""Every function name the docs call resolves — the name half of check 3.

This finishes the documentation-drift proposal in CONTRIBUTING.md. The count
half (``test_doc_catalog_counts.py``) asserts the *number* of catalog functions
the docs state; this one asserts the *names*. A doc that writes
``No_Such_Function(x)`` — or keeps calling a function that was renamed, the
failure the 2026-08-03 review caught by hand in the stale ``X_s`` references —
still renders as prose, and only fails for whoever follows it.

**What counts as a function reference.** A name written as a CALL: an
identifier immediately followed by ``(``, inside an inline code span in prose or
anywhere inside a fenced block. Backticked names in doc TABLE cells are prose
spans, so tables are covered by the same pass. The name must also match the
catalog's own naming convention — ``Upper_Snake_Case`` with at least one
underscore, or a short stats acronym (``AIC`` / ``AICc`` / ``BIC`` / ``GVIF`` /
``PRESS`` / ``VIF`` are the catalog's only underscore-free shapes). That second
condition is what makes the check tractable, and it is a bargain, not an
accident:

* CamelCase single words — natives like ``TAKE`` are 4-plus caps and escape the
  acronym shape; COM methods (``ChartObjects``), doc placeholders
  (``Predictor(``), and the catalog's own CamelCase entries
  (``Coefficients``, ``Interact``, …) are invisible to the detector.
* lowercase names — every Python helper the developer docs legitimately call
  (``safe_activate``, ``build_production_workbook``) is lowercase, and the
  catalog is not.
* bare names — ``Some_Function`` in backticks without a call. Policing those
  would require excluding every named range, dataset column, Role value, and
  Python constant the docs also write bare, which is the unbounded version of
  the exclusion-list problem. A stale bare name is real drift; this check does
  not see it, and ``test_every_catalog_name_is_either_a_candidate_shape_or_a_
  camel_case_word`` keeps that blind spot from ever growing silently.

A name preceded by ``.`` is a fragment of a dotted native (``NORM.INV``,
``MODE.SNGL``) and is not a candidate on its own.

**The exclusion lists, and the guards that keep them honest.** The TODO warned
that the exclusion list is the hard part, so every list is pinned here AND
guarded against rotting:

* ``_NATIVE_EXCEL`` — native functions the docs call. Guard: disjoint from the
  catalog, and every entry actually occurs as a candidate (a dead entry fails,
  so the list tracks the docs' vocabulary, not a dump of Excel).
* ``_PLANNED`` — functions ROADMAP / TODOs name as planned that do not exist
  yet by design. Guard: disjoint from the catalog (when one ships, the test
  forces it out of this list, so the docs' "planned" claim cannot quietly turn
  false while still resolving), and each entry appears in ROADMAP.md or
  TODOs.md — the docs' own planning tags are the authority, not this file.
* ``_SHEET_READER_NAMES`` — the spill readers (``Fit_Context`` & co.): defined
  names installed by the materialization zone, called like functions because
  reading them is ``Name()``.
* ``_RETIRED`` — names that no longer exist and that the shipped-changelog
  prose deliberately still names in rename history (``X_s`` →
  ``Response_Column``). Historical, like DECISIONS.md's entries — but ROADMAP
  is policed, so the names are pinned here instead of excluding the whole
  history section.
* ``_DOC_SHORTHAND`` — call-shaped tokens that are not function references at
  all: ``CDF`` as prose shorthand for the ``CDF_*`` family, and the zone column
  letters in ARCHITECTURE's Univariate band diagram, whose parenthesis is a
  width annotation.

All five are checked for dead entries and none may shadow a catalog function.

**Scope.** The policed list is pinned, and the classification guard asserts
policed ∪ excluded == every markdown file in the repo, so a new hand-authored
doc page must be consciously added here. DECISIONS.md is excluded for the same
reason the count half excludes it (a dated log's superseded names are the
record doing its job), and ``docs/generated/`` is excluded because those pages
are rendered from the authored sheet lists by ``poe docs-generate``, whose
extraction assertions fail at generation time.

Pure ``json`` + ``re`` + ``pathlib``: no Excel, so it runs in the Linux CI job.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

# The documents that describe the catalog and the workbook as they are now (or
# as they are planned to be). Pinned, like the count half's doc set — and the
# classification guard below fails when a markdown file exists that this list
# and _EXCLUDED_PREFIXES do not account for. The tutorial-site pages
# (docs/index.md, technology.md, walkthrough.md, worked-example.md,
# improvements.md) join this list when the docs-site PR lands: they do not
# exist on main yet, and "a policed entry that no longer exists" fails the
# same guard, so the docs-site merge is what adds them.
_POLICED_DOCS = (
    "README.md",
    "CONTRIBUTING.md",
    "CLAUDE.md",
    "AGENTS.md",
    "docs/ROADMAP.md",
    "docs/ARCHITECTURE.md",
    "docs/TODOs.md",
    "docs/MODEL_TESTING_ASSETS.md",
    "excel-only-runs/README.md",
)

# Discovered but not policed, each for a stated reason. Prefixes are matched
# against POSIX repo-relative paths.
_EXCLUDED_PREFIXES: dict[str, str] = {
    "docs/DECISIONS.md": (
        "dated decision log — a superseded name is the record doing its job, "
        "the same reason the count half leaves it out"
    ),
    "docs/generated/": (
        "regenerated by `poe docs-generate` from the authored sheet lists; "
        "its extraction assertions fail at generation time"
    ),
    "docs/_build/": "Sphinx HTML build output",
    ".hermes/": "dated planning documents, frozen at write time",
    ".claude/": "tool skill definitions, not project documentation",
    ".agents/": "tool skill definitions, not project documentation",
}

# Trees that are not project documentation at all: dependency and cache dirs.
_SKIP_DIRS = {".git", ".venv", "venv", "node_modules", "build", "dist", ".cache"}

# Fenced blocks (``` / ~~~) and inline spans (`...`, ``...``). Fences first so a
# stray backtick inside a block cannot pair with one outside it — the same
# ordering test_doc_links.py uses.
_FENCED_BLOCK = re.compile(r"^(?P<fence>```|~~~).*?^(?P=fence)", re.DOTALL | re.MULTILINE)
_INLINE_CODE = re.compile(r"(?P<ticks>`+)(?:(?!(?P=ticks)).)*(?P=ticks)", re.DOTALL)

# A call-shaped reference to a catalog-style name, not the tail of a dotted
# native. The underscore arm is the catalog's Snake_Case convention; the
# acronym arm (2-5 caps, optional trailing lowercase) covers AIC/AICc/BIC/
# GVIF/PRESS/VIF and the short all-caps natives (NA, IF, LN, MAX, TAKE, ...).
_CALL = re.compile(r"(?<![.\w])([A-Z][A-Za-z0-9]*(?:_[A-Za-z0-9]+)+|[A-Z]{2,5}[a-z]?)\s*\(")

# The name shapes above, without the call requirement — used to derive the
# blind spot from the catalog itself rather than pinning it by hand.
_CANDIDATE_NAME = re.compile(r"^[A-Z][A-Za-z0-9]*(?:_[A-Za-z0-9]+)+$|^[A-Z]{2,5}[a-z]?$")
# The one other shape a catalog name may have: a CamelCase word
# (Coefficients, Interact, Observations, ...), invisible to the detector on
# purpose — it is the shape the docs share with roles, classes, and datasets.
_CAMEL_CASE_WORD = re.compile(r"^[A-Z][a-z]+$")

# Native Excel functions the policed docs call. Pinned to the docs' actual
# vocabulary: every entry is exercised (the dead-entry guard prunes any that
# stops appearing), and none may be a catalog function.
_NATIVE_EXCEL = frozenset(
    {
        "AND", "BYROW", "COUNT", "DEVSQ", "EXP", "IF", "INDEX", "ISNA",
        "LET", "LN", "MAP", "MAX", "MIN", "MOD", "NA", "ROWS",
        "SUMSQ", "TAKE",
    }
)

# Functions the planning docs name that do not exist yet, by design. The guard
# is that each one appears in docs/ROADMAP.md or docs/TODOs.md — this file is
# not the authority on what is planned, the planning docs are.
_PLANNED = frozenset(
    {
        "ADF_Critical_Value", "ADF_Statistic", "Absorb_Two_Way_Fixed_Effects",
        "Bootstrap_CI", "Covariance_Matrix", "Decompose_By",
        "Demean_Two_Way_Balanced", "Exponential_Smoothing", "F_Test_Variance",
        "Fixed_Effects_Convergence_Check", "KPSS_Critical_Value",
        "KPSS_Statistic", "Ljung_Box_Q", "MC_Percentile", "Minmax_Scale",
        "Model_Formula_String", "Moving_Average", "PERT_Sample",
        "T_Test_OneSample", "T_Test_TwoSample", "Zscore_By",
    }
)

# The spill readers the v3.2 materialization zone installs: defined names on
# the Regression sheet, called like functions because reading a spill is
# `Name()`. Not catalog entries, but not drift either.
_SHEET_READER_NAMES = frozenset({"Fit_Context", "Fit_Design_Columns", "Fit_Sample_Include"})

# Names the v3.0 constructor pipeline retired, which the shipped-changelog
# prose in ROADMAP still names in its rename history — "shipped at v2.1 as
# `y_s` / `X_s_Within`; renamed by the v3.0 constructor pipeline". Historical
# reference, not a live claim.
_RETIRED = frozenset({"X_s", "X_s_Within"})

# Call-shaped tokens that are not function references at all: CDF is prose
# shorthand for the CDF_* family ("`CDF(upper edge) − CDF(lower edge)`"), and
# BS/BU/BW are the zone column letters in ARCHITECTURE's Univariate band
# diagram, where the parenthesis is a width annotation, not a call.
_DOC_SHORTHAND = frozenset({"CDF", "BS", "BU", "BW"})


def _catalog_names() -> set[str]:
    functions = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))["functions"]
    return {entry["name"] for entry in functions}


def _all_exclusions() -> set[str]:
    return _NATIVE_EXCEL | _PLANNED | _SHEET_READER_NAMES | _RETIRED | _DOC_SHORTHAND


def _candidates(path: Path) -> list[tuple[int, str]]:
    """Every call-shaped ``(line, name)`` reference in one document.

    The text is split on its fences first, and inline spans are paired within
    each prose region separately. Pairing backticks across the whole text
    instead would work only while every fenced block holds an even number of
    them — the docs' fences do carry backticks (escaped markdown-syntax
    examples), so one odd count there would mis-pair every span after it and
    silently skip prose references. Positions stay against the original text
    so the reported line numbers are true.
    """
    text = path.read_text(encoding="utf-8")
    fences = [block.span() for block in _FENCED_BLOCK.finditer(text)]

    def line_of(position: int) -> int:
        return text.count("\n", 0, position) + 1

    prose_regions: list[tuple[int, str]] = []
    cursor = 0
    for start, end in fences:
        if text[cursor:start]:
            prose_regions.append((cursor, text[cursor:start]))
        cursor = end
    if text[cursor:]:
        prose_regions.append((cursor, text[cursor:]))

    found: list[tuple[int, str]] = []
    for offset, chunk in prose_regions:
        for span in _INLINE_CODE.finditer(chunk):
            for call in _CALL.finditer(span.group(0)):
                found.append((line_of(offset + span.start() + call.start(1)), call.group(1)))
    for (start, end) in fences:
        for call in _CALL.finditer(text[start:end]):
            found.append((line_of(start + call.start(1)), call.group(1)))
    return found


def _unresolved(path: Path, catalog: set[str]) -> list[str]:
    allowed = catalog | _all_exclusions()
    return [
        f"line {line}: {name}"
        for line, name in _candidates(path)
        if name not in allowed
    ]


def _markdown_files() -> list[Path]:
    """Every markdown file in the repo, dependency and cache trees excluded."""
    return sorted(
        path
        for path in ROOT_DIR.rglob("*.md")
        if not _SKIP_DIRS & set(path.relative_to(ROOT_DIR).parts)
    )


def test_every_catalog_name_is_either_a_candidate_shape_or_a_camel_case_word() -> None:
    """The blind spot is derived from the catalog, never grown in silence.

    Every catalog name matches the detector's shapes OR is a CamelCase word
    (the documented blind spot). A catalog function named outside both shapes
    — a lowercase name, a dotted name — fails HERE, naming it, so whoever
    added it learns the naming convention and updates the detector.
    """
    outside = sorted(
        name
        for name in _catalog_names()
        if not _CANDIDATE_NAME.match(name) and not _CAMEL_CASE_WORD.match(name)
    )
    assert not outside, f"catalog names outside every known shape: {outside}"


def test_exclusions_cannot_shadow_catalog_functions() -> None:
    """A name can be in one universe only.

    The dangerous case is `_PLANNED`: a planned function that ships stays in
    the exclusion list, keeps resolving, and the check keeps passing while the
    docs' "planned" claim quietly turned false. When a name enters the
    catalog, remove it from every exclusion list in the same commit.
    """
    shadowed = sorted(_catalog_names() & _all_exclusions())
    assert not shadowed, f"names claimed by both the catalog and an exclusion list: {shadowed}"


def test_planned_names_appear_in_the_planning_docs() -> None:
    """`_PLANNED` takes its authority from ROADMAP / TODOs, not from itself.

    A name in `_PLANNED` that no planning doc names is a dump-ground entry —
    the failure mode the exclusion-list warning in TODOs.md predicted. Each
    entry must appear in docs/ROADMAP.md or docs/TODOs.md.
    """
    roadmap = (ROOT_DIR / "docs/ROADMAP.md").read_text(encoding="utf-8")
    todos = (ROOT_DIR / "docs/TODOs.md").read_text(encoding="utf-8")
    missing = sorted(name for name in _PLANNED if name not in roadmap and name not in todos)
    assert not missing, f"planned names no planning document plans: {missing}"


def test_exclusion_lists_carry_no_dead_entries() -> None:
    """Every exclusion entry still matches a candidate the docs actually write.

    An entry that stopped appearing is a name the check no longer has a reason
    to allow. Prune it when the prose moves on — the lists track the docs, not
    the other way round. (Every planned entry is call-shaped in the policed
    docs today too; if one stops being so, that is the same signal.)
    """
    seen = {
        name
        for relative_path in _POLICED_DOCS
        for _, name in _candidates(ROOT_DIR / relative_path)
    }
    dead = sorted(_all_exclusions() - seen)
    assert not dead, f"exclusion entries matching nothing in the policed docs: {dead}"


def test_every_markdown_file_is_classified() -> None:
    """Policed ∪ excluded == every markdown file in the repo.

    A new hand-authored doc page is drift surface: this fails until it is
    either added to `_POLICED_DOCS` or given a reason in `_EXCLUDED_PREFIXES`.
    A policed entry that no longer exists fails the same way.
    """
    discovered = {
        path.relative_to(ROOT_DIR).as_posix() for path in _markdown_files()
    }
    unclassified = sorted(
        relative_path
        for relative_path in discovered
        if relative_path not in _POLICED_DOCS
        and not any(
            relative_path == prefix.rstrip("/") or relative_path.startswith(prefix)
            for prefix in _EXCLUDED_PREFIXES
        )
    )
    assert not unclassified, (
        "markdown files neither policed nor excluded: "
        f"{unclassified} — classify each in tests/test_doc_function_names.py"
    )

    missing = sorted(set(_POLICED_DOCS) - discovered)
    assert not missing, f"policed docs that do not exist: {missing}"


def test_the_check_still_sees_work() -> None:
    """Guard against the detector silently matching nothing and vacuously passing."""
    found = [
        (relative_path, name)
        for relative_path in _POLICED_DOCS
        for _, name in _candidates(ROOT_DIR / relative_path)
    ]
    assert len(found) >= 300, f"only {len(found)} call-shaped references found"
    catalog = _catalog_names()
    resolving = {name for _, name in found if name in catalog}
    assert len(resolving) >= 30, (
        f"only {len(resolving)} distinct catalog names resolve — the docs' "
        "function vocabulary has moved away from this check"
    )


@pytest.mark.parametrize("relative_path", _POLICED_DOCS, ids=str)
def test_function_references_resolve(relative_path: str) -> None:
    """Each call-shaped, catalog-style name in a policed doc resolves.

    Resolves means: an entry in `lambda_functions.json`, or one of the five
    pinned exclusion categories. The failure names the line and the name so
    whoever hits it can see both sides, the same bargain the count half makes.
    """
    unresolved = _unresolved(ROOT_DIR / relative_path, _catalog_names())
    assert not unresolved, (
        f"{relative_path} calls function name(s) nothing defines — not a "
        "catalog entry, native Excel function, planned name, sheet reader, "
        "retired name, or doc shorthand: " + "; ".join(unresolved)
        + ". If it is planned, tag it in _PLANNED (it must then appear in "
        "ROADMAP/TODOs); if it no longer exists, the doc is stale."
    )


def test_a_stale_reference_is_reported_with_its_line(tmp_path: Path) -> None:
    """The failure has to say WHERE and WHAT, or whoever hits it still has to
    read the whole document looking for the call."""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "# Title\n\nuses `No_Such_Function(x)` here\n", encoding="utf-8"
    )
    assert _candidates(doc) == [(3, "No_Such_Function")]


def test_dotted_native_fragments_are_not_candidates(tmp_path: Path) -> None:
    """`NORM.INV(` is one native call, not a reference to a function INV."""
    doc = tmp_path / "doc.md"
    doc.write_text("`=NORM.INV(0.5, p, n)` and `INV(x)` bare\n", encoding="utf-8")
    assert _candidates(doc) == [(1, "INV")]


def test_the_blind_spots_are_not_candidates(tmp_path: Path) -> None:
    """CamelCase words, lowercase helpers, and bare names are invisible by
    design — pinned here so the bargain is a decision, not an accident."""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "`Predictor(x)` `Interact` bare `safe_activate(sheet)` `ln(x)` "
        "`XLOOKUP(x)` `Some_Function` bare\n",
        encoding="utf-8",
    )
    assert _candidates(doc) == []


def test_calls_inside_fenced_blocks_are_checked(tmp_path: Path) -> None:
    """Fenced formulas are the load-bearing surface — a stale name in a formula
    is the failure class the 2026-08-03 review actually found."""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "intro\n\n```excel\n=No_Such_Function(A1)\n```\n\ntail\n",
        encoding="utf-8",
    )
    assert _candidates(doc) == [(4, "No_Such_Function")]


def test_an_odd_backtick_inside_a_fence_does_not_break_prose_pairing(
    tmp_path: Path,
) -> None:
    """The docs' fences do carry backticks (escaped markdown-syntax examples);
    pairing prose spans globally would let one odd count there swallow every
    span after it. Region-splitting is what keeps this a doc quirk, not a
    silent skip."""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "```\nhow to write ` a fence\n```\n\nthen `No_Such_Function(x)` after\n",
        encoding="utf-8",
    )
    assert _candidates(doc) == [(5, "No_Such_Function")]
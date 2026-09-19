"""The committed static-sheet template carries the text its writers produce.

**The failure this closes.** Three modules author the sheets in
`templates/static_sheets.xlsx` — `write_sheet_regression_instructions.py`,
`write_sheet_modeling_concepts.py`, `write_sheet_diagnostic_guide.py`. Every
artifact build copies the *already-baked* sheet out of that template
(`workbook_helpers.copy_static_sheet`); nothing at build time executes the
writers. So editing one of those modules changes nothing any build can see until
`scripts/rebuild_static_sheets.py` is run and the regenerated template is
committed. That regeneration needs desktop Excel, so it is a manual step that
has been missed at least twice, shipping stale doc text.

A skill can only remind; this is the guard. It reads the writers' *source* and
the template's *bytes* — no import of the writer modules, no xlwings, no Excel —
so it runs in the Linux CI job beside the other committed-artifact checks.

**What it asserts.** Every string literal those writers put in a cell appears in
the corresponding sheet of the committed template. It is deliberately
one-directional: it catches "edited the Python, forgot to regenerate" — the
failure that actually happens — and does not attempt to prove the template has
nothing *extra* (a hand-edited template, or one regenerated from different
source, is invisible here), nor to check layout (rows, columns, merges, widths).
Those need the Excel-machine path.

**How the literals are collected, and why it is not "every string constant".**
The three modules write cell text three ways, and only those three paths are
read — which is what keeps docstrings, log messages and other prose out, with no
filtering rule of their own:

* a module-level content constant — `_ROWS` / `_FEATURES` / `_THRESHOLDS` / …
  (pinned per module in ``_CONTENT_CONSTANTS``, with a dead-entry guard);
* a helper call — `_heading` / `_subheading` / `_table_header_row` / `_row` —
  whose *third* argument carries the text: a single string for the two heading
  helpers, a list of strings for the two row helpers. Every helper in all three
  modules puts the text in that position, which is what lets one rule cover them;
* a direct assignment — `sheet.range(...).value = "…"`.

A helper's text argument is sometimes a local name rather than a literal
(`for vals in fe_residuals: _row(sheet, r, vals)`) — and the name the helper
receives is the *loop* variable, not the list — so local bindings are resolved
too: see ``_local_literals``, which resolves an assignment to a literal, and a
loop target over one, and documents the two deliberately different rules for a
name bound twice. The check reports what it can prove; where it cannot prove,
the gap is stated rather than guessed at.

Only **static** literals can be checked. A cell whose text is built by
interpolation — `"plain " f"interpolated {addr}" "plain"` — parses as a
``JoinedStr``, and its value at runtime is not knowable here. Those are handled
in a second, weaker tier: any literal *segment* of the interpolation at least
``_MIN_INTERPOLATED_SEGMENT`` characters long must appear as a **substring** of
some cell text in that sheet. Shorter segments are skipped as non-evidence —
they are boilerplate like ``"Cell "`` or ``", Diagnostics"``, which would match
almost anything. This keeps the paragraph text around an interpolation
(`"below 0.95 = stronger concern. The chart title itself shows…"`) under
coverage while accepting that the interpolated value itself is not.

``_ROWS`` carries one shape rule: its entries are ``(row, text, style)`` and the
writer ``continue``s when ``style is None``, so such a row is never written and
its text is *correctly* absent from the template. Applying it here is what keeps
the check from reporting a false positive on a deliberately-blank row.

Strings starting with ``=`` are formulas, which the package stores as formula
parts rather than cell text, and are skipped.
"""
from __future__ import annotations

import ast
import posixpath
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

import pytest
from lxml import etree

ROOT_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = ROOT_DIR / "templates/static_sheets.xlsx"

# Which module authors which template sheet is not pinned here: it is read from
# each module's own ``SHEET_NAME``, so a rename cannot leave this table stale.
#
# The content constants are pinned, exactly as the doc-drift checks pin their
# doc sets: a module's cell text lives in a named module-level constant, and
# discovering those by shape would be guesswork. ``_COLUMNS`` is deliberately
# absent, and the reason is worth stating because it *does* carry a header
# string: only ``ColumnSpec.width`` (``set_column_widths``) and ``.index``
# (``_LAST_COL``) are ever read, so those headers never reach a cell — the
# header rows are written from literals at the call sites, which are collected.
_CONTENT_CONSTANTS: dict[str, tuple[str, ...]] = {
    "write_sheet_regression_instructions.py": ("_ROWS",),
    "write_sheet_modeling_concepts.py": ("_FEATURES", "_PLANNED_FEATURES"),
    "write_sheet_diagnostic_guide.py": (
        "_TIER1", "_TIER2", "_THRESHOLDS", "_GUIDANCE",
    ),
}

# The four helpers every one of these modules writes cell text through. The text
# is the third positional argument in all four signatures
# (``(sheet, row, text…)``), which is what makes one rule cover them.
_TEXT_HELPERS = frozenset({"_row", "_heading", "_subheading", "_table_header_row"})

# An interpolated segment shorter than this is boilerplate, not evidence.
_MIN_INTERPOLATED_SEGMENT = 20

_TEXT_ARG_INDEX = 2


def _norm(text: str) -> str:
    """Whitespace-insensitive form, for comparing source text to stored text.

    Source literals are wrapped across lines and the package may store a
    newline as a literal character or an entity; collapsing every whitespace run
    keeps those differences from reading as content drift. Only whitespace is
    forgiven — every word still has to match.
    """
    return " ".join(text.split())


@dataclass(frozen=True)
class ModuleLiterals:
    """What one writer module says it will write."""

    sheet_name: str
    exact: tuple[str, ...]
    interpolated: tuple[str, ...]

    @property
    def checked_segments(self) -> tuple[str, ...]:
        """Interpolated segments long enough to be evidence."""
        return tuple(
            segment
            for segment in self.interpolated
            if len(_norm(segment)) >= _MIN_INTERPOLATED_SEGMENT and " " in _norm(segment)
        )


def _is_formula(text: str) -> bool:
    return text.lstrip().startswith("=")


def _keep(text: str) -> bool:
    return bool(text.strip()) and not _is_formula(text)


class _LiteralCollector(ast.NodeVisitor):
    """Collect cell-text literals under one node.

    An interpolated string is recorded as *segments* and not descended into: a
    ``FormattedValue``'s contents are runtime values, and the surrounding
    ``Constant`` parts are only checkable as substrings.
    """

    def __init__(self) -> None:
        self.exact: list[str] = []
        self.interpolated: list[str] = []

    def visit_JoinedStr(self, node: ast.JoinedStr) -> None:
        for part in node.values:
            if isinstance(part, ast.Constant) and isinstance(part.value, str):
                self.interpolated.append(part.value)
        # No generic_visit: the nested FormattedValue trees are not literals.

    def visit_Constant(self, node: ast.Constant) -> None:
        if isinstance(node.value, str):
            self.exact.append(node.value)


def _collect_from(node: ast.AST | None) -> _LiteralCollector:
    collector = _LiteralCollector()
    if node is not None:
        collector.visit(node)
    return collector


def _local_literals(func: ast.FunctionDef) -> dict[str, tuple[ast.AST, ...]]:
    """Local names in ``func`` that hold a literal list/tuple/string.

    Cell text is sometimes assembled in a local before it reaches a row helper
    (``for vals in fe_residuals: _row(sheet, r, vals)``), so a literal-argument
    rule alone would leave that block uncollected — silently, which is the
    failure this file exists to prevent.

    Two kinds of binding resolve, and they resolve *differently*:

    * an **assignment** to a literal (``block = [["a", "b"]]``) resolves only
      when it is the name's single binding. A re-bound assignment is skipped
      rather than guessed at — which write reached the helper is not statically
      obvious, and guessing would be worse than a gap that is stated.
    * a **loop target** (``for vals in fe_residuals``, and the ``for vals in
      block`` case whose iterable is itself a resolved local) resolves to the
      *union* of its iterables. A reused loop variable is the ordinary pattern —
      the Diagnostic Guide binds ``vals`` in five sibling loops, each writing its
      own rows — so skipping on a second binding would gut the rule for no gain.
      Collecting from the iterable covers every row that loop writes.
    """
    assigned: dict[str, ast.AST] = {}
    rebound: set[str] = set()
    for node in ast.walk(func):
        if (
            isinstance(node, ast.Assign)
            and len(node.targets) == 1
            and isinstance(node.targets[0], ast.Name)
            and isinstance(node.value, (ast.List, ast.Tuple, ast.Constant))
        ):
            name = node.targets[0].id
            if name in assigned:
                rebound.add(name)
            assigned[name] = node.value

    resolved: dict[str, list[ast.AST]] = {
        name: [value] for name, value in assigned.items() if name not in rebound
    }
    for node in ast.walk(func):
        if isinstance(node, ast.For) and isinstance(node.target, ast.Name):
            iterable = node.iter
            if isinstance(iterable, ast.Name):
                iterable = assigned.get(iterable.id)
            if isinstance(iterable, (ast.List, ast.Tuple, ast.Constant)):
                resolved.setdefault(node.target.id, []).append(iterable)
    return {name: tuple(nodes) for name, nodes in resolved.items()}


def _rows_collector(node: ast.AST) -> _LiteralCollector:
    """``_ROWS`` shape: ``(row, text, style)``, where ``style is None`` is skipped.

    The writer reads ``for row, text, style in _ROWS`` and ``continue``s on a
    ``None`` style, so such a row's text never reaches the sheet.

    The text slot goes through the collector rather than a bare ``isinstance``
    check: taking only a ``Constant`` there would drop an interpolated cell
    *silently*, which is the failure mode this whole file exists to avoid.
    """
    collector = _LiteralCollector()
    for element in getattr(node, "elts", ()):
        if not isinstance(element, ast.Tuple) or len(element.elts) < 3:
            continue
        style = element.elts[2]
        if isinstance(style, ast.Constant) and style.value is None:
            continue  # the writer `continue`s on this row
        collector.visit(element.elts[1])
    return collector


def _absorb_module_constants(
    tree: ast.Module, constants: tuple[str, ...], absorb: Callable[[_LiteralCollector], None]
) -> str:
    """Absorb the pinned content constants, and return the module's ``SHEET_NAME``."""
    sheet_name = ""
    for node in tree.body:
        if not isinstance(node, (ast.Assign, ast.AnnAssign)):
            continue
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        if len(targets) != 1 or not isinstance(targets[0], ast.Name):
            continue
        name, value = targets[0].id, node.value
        if name == "SHEET_NAME" and isinstance(value, ast.Constant):
            sheet_name = value.value
        elif name in constants:
            absorb(_rows_collector(value) if name == "_ROWS" else _collect_from(value))
    return sheet_name


def _absorb_template_writes(
    func: ast.FunctionDef, absorb: Callable[[_LiteralCollector], None]
) -> None:
    """Absorb the text written by a ``_write_template_sheet`` body."""
    local_literals = _local_literals(func)
    for sub in ast.walk(func):
        if isinstance(sub, ast.Assign):
            for target in sub.targets:
                if isinstance(target, ast.Attribute) and target.attr == "value":
                    absorb(_collect_from(sub.value))
        if (
            isinstance(sub, ast.Call)
            and isinstance(sub.func, ast.Name)
            and sub.func.id in _TEXT_HELPERS
            and len(sub.args) > _TEXT_ARG_INDEX
        ):
            argument = sub.args[_TEXT_ARG_INDEX]
            if isinstance(argument, ast.Name):
                for resolved in local_literals.get(argument.id, ()):
                    absorb(_collect_from(resolved))
            else:
                absorb(_collect_from(argument))


def _settle(exact: list[str], interpolated: list[str]) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """De-duplicate, order-preserving.

    One literal can be reached twice — through a pinned constant and through the
    loop that writes it — and a duplicate would be reported twice and counted
    twice by the floors. The constants themselves also repeat strings (the
    Diagnostic Guide's header cells recur across its tables), so a raw count is
    not the coverage; a unique count is.
    """
    return tuple(dict.fromkeys(exact)), tuple(dict.fromkeys(interpolated))


def _module_literals(path: Path, constants: tuple[str, ...]) -> ModuleLiterals:
    """Read one writer module's source and extract the cell text it writes."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    exact: list[str] = []
    interpolated: list[str] = []

    def absorb(collector: _LiteralCollector) -> None:
        exact.extend(t for t in collector.exact if _keep(t))
        interpolated.extend(t for t in collector.interpolated if _keep(t))

    sheet_name = _absorb_module_constants(tree, constants, absorb)
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == "_write_template_sheet":
            _absorb_template_writes(node, absorb)

    assert sheet_name, f"{path.name} has no SHEET_NAME constant"
    return ModuleLiterals(sheet_name, *_settle(exact, interpolated))


def _template_sheet_parts() -> dict[str, str]:
    """{sheet name: zip member} for the committed template.

    Relationship ``Target`` is absolute (``/xl/worksheets/sheet1.xml``) in this
    artifact — Excel's own save writes that form, where the xlwings save path in
    ``dist/`` writes relative ones. Both are normalised here; the suite's
    existing ``WorkbookPackage.sheet_part_for`` assumes the relative form only.
    """
    with zipfile.ZipFile(TEMPLATE_PATH) as archive:
        workbook = etree.fromstring(archive.read("xl/workbook.xml"))
        rels = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") or "" for rel in rels}
    parts: dict[str, str] = {}
    for sheet in workbook.findall(".//{*}sheet"):
        rid = next((v for k, v in sheet.attrib.items() if k.endswith("}id")), "")
        raw = targets.get(rid, "").lstrip("/")
        parts[sheet.get("name")] = raw if raw.startswith("xl/") else posixpath.join("xl", raw)
    return parts


def _cell_texts(sheet_part: str) -> set[str]:
    """Every string stored in one worksheet, normalised.

    The template has no ``sharedStrings.xml``: xlwings writes cell text inline
    (``t="inlineStr"`` → ``<is><t>``), so the text is in the worksheet part. Each
    ``is``/``si`` is joined across its runs — a rich-text cell splits one value
    into several ``<t>`` elements — and every individual ``<t>`` is also kept,
    so a value split by runs still matches either way.
    """
    with zipfile.ZipFile(TEMPLATE_PATH) as archive:
        root = etree.fromstring(archive.read(sheet_part))
    texts: set[str] = set()
    for element in root.iter():
        if element.tag.endswith(("}is", "}si")):
            texts.add(_norm("".join(t.text or "" for t in element.iter() if t.tag.endswith("}t"))))
        elif element.tag.endswith("}t"):
            texts.add(_norm(element.text or ""))
    return {text for text in texts if text}


def _missing(literals: ModuleLiterals, cell_texts: set[str]) -> list[str]:
    """Literals the module writes that the sheet's stored text does not carry."""
    blob = "\n".join(sorted(cell_texts))
    absent = [text for text in literals.exact if _norm(text) not in cell_texts]
    absent += [
        segment for segment in literals.checked_segments if _norm(segment) not in blob
    ]
    return absent


_CASES = tuple(_CONTENT_CONSTANTS)


@pytest.mark.parametrize("module_name", _CASES, ids=str)
def test_every_static_literal_reaches_the_committed_template(module_name: str) -> None:
    """The template on disk carries the text the writers produce today.

    A failure here means the writer module was edited without running
    ``python scripts/rebuild_static_sheets.py`` and committing the result — so
    every artifact build is copying the older sheet. Fix by regenerating (which
    needs desktop Excel) and committing the template, not by editing this test.
    """
    literals = _module_literals(ROOT_DIR / "lambda_catalog" / module_name, _CONTENT_CONSTANTS[module_name])
    parts = _template_sheet_parts()
    assert literals.sheet_name in parts, (
        f"{module_name} names sheet {literals.sheet_name!r}, which the template "
        f"does not contain: {sorted(parts)}"
    )
    absent = _missing(literals, _cell_texts(parts[literals.sheet_name]))
    assert not absent, (
        f"{module_name} writes {len(absent)} string(s) that "
        f"{literals.sheet_name!r} in templates/static_sheets.xlsx does not carry — "
        "the template is stale. Regenerate it with "
        "`python scripts/rebuild_static_sheets.py` and commit it:\n  "
        + "\n  ".join(_norm(text)[:120] for text in absent)
    )


def test_the_template_carries_all_three_static_sheets() -> None:
    """The template is the only source of these sheets, so all three must be in it."""
    parts = _template_sheet_parts()
    expected = {
        _module_literals(ROOT_DIR / "lambda_catalog" / name, constants).sheet_name
        for name, constants in _CONTENT_CONSTANTS.items()
    }
    assert expected <= set(parts), f"template is missing {sorted(expected - set(parts))}"


def test_every_pinned_content_constant_still_exists_and_holds_text() -> None:
    """Dead-entry guard on ``_CONTENT_CONSTANTS``.

    A pinned constant that stopped existing, or that yields nothing this check
    can see, silently narrows the check to nothing — the vacuous-pass failure
    the floors below exist for, one level up. Prune or repoint the entry.
    """
    for module_name, constants in _CONTENT_CONSTANTS.items():
        literals = _module_literals(ROOT_DIR / "lambda_catalog" / module_name, constants)
        assert literals.exact, f"{module_name} yielded no static literals from {constants}"
        for constant in constants:
            source = (ROOT_DIR / "lambda_catalog" / module_name).read_text(encoding="utf-8")
            assert f"\n{constant}" in source, f"{module_name} no longer defines {constant}"


def test_the_check_still_sees_work() -> None:
    """Floors, so a detector that stopped matching cannot pass quietly.

    The first draft of this check collected *nothing* from `_ROWS` and reported
    success: the content constants are annotated assignments (``ast.AnnAssign``),
    and the extractor only looked at plain ``Assign`` nodes. Floors are how that
    is caught rather than trusted.
    """
    counts = {
        module_name: len(
            _module_literals(ROOT_DIR / "lambda_catalog" / module_name, constants).exact
        )
        for module_name, constants in _CONTENT_CONSTANTS.items()
    }
    assert sum(counts.values()) >= 150, f"only {sum(counts.values())} static literals found: {counts}"
    thin = {name: count for name, count in counts.items() if count < 5}
    assert not thin, f"modules yielding almost nothing (extractor regression?): {thin}"


def test_interpolated_segments_are_still_found() -> None:
    """The substring tier is load-bearing for the Diagnostic Guide's prose.

    That sheet's longest guidance cell is built around an interpolation, so the
    exact tier cannot see it. If the segments ever stop being collected, the
    guide's coverage silently drops to the exact tier alone.
    """
    guide = _module_literals(
        ROOT_DIR / "lambda_catalog" / "write_sheet_diagnostic_guide.py",
        _CONTENT_CONSTANTS["write_sheet_diagnostic_guide.py"],
    )
    assert guide.interpolated, "no interpolated segments collected from the Diagnostic Guide"
    assert guide.checked_segments, "no interpolated segment was long enough to check"


# --------------------------------------------------------------------------
# The extraction and comparison rules, on synthetic input.
# --------------------------------------------------------------------------


def test_static_and_interpolated_literals_are_separated(tmp_path: Path) -> None:
    """A plain literal is exact; an interpolated one is only checkable in parts.

    Adjacent constant parts are *folded* by the parser — ``"plain " f"mid {a}"``
    yields a leading ``Constant("plain mid ")``, not two — so the segment list is
    the interpolation's static text in runs, which is what the substring tier
    wants anyway.
    """
    module = tmp_path / "writer.py"
    module.write_text(
        'SHEET_NAME = "S"\n'
        '_ROWS: list = [\n'
        '    (1, "plain text here", "heading"),\n'
        '    (2, "plain " f"mid {addr}" "tail", "body"),\n'
        "]\n"
        "def _write_template_sheet(workbook):\n"
        '    _heading(sheet, 1, "helper literal")\n'
        '    sheet.range((2, 1)).value = "assigned literal"\n'
        '    sheet.range((3, 1)).value = "=FORMULA(A1)"\n',
        encoding="utf-8",
    )
    literals = _module_literals(module, ("_ROWS",))
    assert literals.sheet_name == "S"
    assert "plain text here" in literals.exact
    assert "helper literal" in literals.exact
    assert "assigned literal" in literals.exact
    assert not any("FORMULA" in text for text in literals.exact)
    assert literals.interpolated == ("plain mid ", "tail")


def test_a_rows_entry_with_no_style_is_not_written(tmp_path: Path) -> None:
    """The writer ``continue``s on ``style is None``, so its text is never a cell."""
    module = tmp_path / "writer.py"
    module.write_text(
        'SHEET_NAME = "S"\n'
        "_ROWS = [\n"
        '    (1, "written", "heading"),\n'
        '    (2, "skipped", None),\n'
        "]\n"
        "def _write_template_sheet(workbook):\n"
        "    pass\n",
        encoding="utf-8",
    )
    literals = _module_literals(module, ("_ROWS",))
    assert list(literals.exact) == ["written"]


def test_a_row_helper_argument_bound_to_a_local_is_resolved(tmp_path: Path) -> None:
    """The local-name path — the Diagnostic Guide's FE-residual table uses it.

    The name the helper receives is the *loop* variable, so this also pins the
    resolution chain: ``vals`` → ``block`` → the literal.
    """
    module = tmp_path / "writer.py"
    module.write_text(
        'SHEET_NAME = "S"\n'
        "def _write_template_sheet(workbook):\n"
        '    block = [["first cell", "second cell"]]\n'
        "    for vals in block:\n"
        "        _row(sheet, r, vals)\n",
        encoding="utf-8",
    )
    literals = _module_literals(module, ())
    assert "first cell" in literals.exact and "second cell" in literals.exact


def test_a_loop_target_reused_across_loops_unions_its_iterables(tmp_path: Path) -> None:
    """A reused loop variable is ordinary, so every loop's rows are collected."""
    module = tmp_path / "writer.py"
    module.write_text(
        'SHEET_NAME = "S"\n'
        "def _write_template_sheet(workbook):\n"
        "    for vals in [[\"row one cell\"]]:\n"
        "        _row(sheet, r, vals)\n"
        "    for vals in [[\"row two cell\"]]:\n"
        "        _row(sheet, r, vals)\n",
        encoding="utf-8",
    )
    literals = _module_literals(module, ())
    assert "row one cell" in literals.exact and "row two cell" in literals.exact


def test_a_local_rebound_more_than_once_is_not_guessed_at(tmp_path: Path) -> None:
    """Two bindings means two possible values, so neither is claimed."""
    module = tmp_path / "writer.py"
    module.write_text(
        'SHEET_NAME = "S"\n'
        "def _write_template_sheet(workbook):\n"
        '    block = ["first value"]\n'
        '    block = ["second value"]\n'
        "    _row(sheet, r, block)\n",
        encoding="utf-8",
    )
    assert _module_literals(module, ()).exact == ()


def test_a_literal_absent_from_the_sheet_is_reported(tmp_path: Path) -> None:
    """The comparison reports the missing text, which is the whole point."""
    module = tmp_path / "writer.py"
    module.write_text(
        'SHEET_NAME = "S"\n_ROWS = [(1, "present", "heading"), (2, "absent", "body")]\n'
        "def _write_template_sheet(workbook):\n    pass\n",
        encoding="utf-8",
    )
    literals = _module_literals(module, ("_ROWS",))
    assert _missing(literals, {"present"}) == ["absent"]


def test_whitespace_differences_do_not_read_as_drift() -> None:
    """Wrapping and stored newlines must not be reported as content drift."""
    module_literals = ModuleLiterals("S", ("a wrapped\n  sentence",), ())
    assert _missing(module_literals, {"a wrapped sentence"}) == []


def test_a_short_interpolated_segment_is_not_evidence() -> None:
    """Boilerplate around an interpolation would match almost anything."""
    literals = ModuleLiterals("S", (), ("Cell ", ", Diagnostics"))
    assert literals.checked_segments == ()

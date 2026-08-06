"""Every relative markdown link in the repo's documentation resolves — file and anchor.

These are checks 1 and 2 of the documentation drift proposal in CONTRIBUTING.md.
Both exist because the failure mode is silent: a link that points at a file which
is not there, or at a heading that has been renamed, still renders as a link and
only fails for whoever clicks it.

**The path half.** Two breakages it would have caught, both real:

* **Deleted targets.** ``docs/REVIEW.md`` was removed while four documents still
  linked to it.
* **Wrong-depth prefixes.** A pass that moved the planning docs under ``docs/``
  prefixed every relative link with ``docs/`` — including links *inside*
  ``docs/``, where the prefix is one level too many. ``docs/TODOs.md`` pointing
  at ``docs/ROADMAP.md`` resolves to ``docs/docs/ROADMAP.md``. That single
  commit broke 171 links, and nothing failed.

**The anchor half.** A heading rename breaks every link into it with no error
anywhere — the ``ARCHITECTURE.md`` §4 renames from ``(A–L)`` to ``(A–N)`` to
``(A–O)`` each broke at least one ``ROADMAP.md`` link this way, and the commit
that retitled CONTRIBUTING's own *Documentation drift* section broke the
``TODOs.md`` link into it. Checking anchors means reproducing GitHub's slug
rules, which is why it is the larger of the two: lowercase the heading, drop
everything outside ``[a-z0-9 _-]`` (**underscores survive** — ``Model_Context``
headings depend on it), collapse spaces to hyphens, and suffix a repeated slug
with ``-1``, ``-2``, … in document order.

Headings inside fenced code blocks are not headings. A block that shows shell
comments or a markdown example would otherwise register ``# some comment`` as a
target, so slug collection tracks fences line by line rather than reusing the
whole-text stripper.

Code spans are stripped before scanning for links. Documentation that
*describes* link syntax writes things like ``](FILE.md#anchor)`` inside
backticks, and those are examples of a notation, not links to a document —
treating them as links would make the docs unable to talk about their own
conventions.

Pure ``re`` + ``pathlib``: no Excel, so it runs in the Linux CI job.
"""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

import pytest

ROOT_DIR = Path(__file__).resolve().parents[1]

# ``](target.md)`` or ``](target.md#anchor)`` — capture the path half only.
_MARKDOWN_LINK = re.compile(r"\]\(([^)\s#]+\.md)(?:#[^)]*)?\)")

# ``](target.md#anchor)`` or, with the path half absent, ``](#anchor)`` — a link
# into the linking document itself. A URL fragment (``](https://…#readme)``)
# cannot match: the optional path group only accepts a ``.md`` suffix, so the
# ``#`` has to sit immediately after ``](`` for the pathless form to apply.
_ANCHORED_LINK = re.compile(r"\]\(([^)\s#]*\.md)?#([^)\s]+)\)")

# An ATX heading, and the fence that means the next one is not a heading at all.
_ATX_HEADING = re.compile(r"#{1,6}\s+(.*)")
_FENCE = re.compile(r"^\s*(```|~~~)")

# A markdown link inside a heading contributes its TEXT to the slug, not its URL.
_LINK_IN_HEADING = re.compile(r"\[([^\]]*)\]\([^)]*\)")

# Fenced blocks (``` / ~~~) and inline spans (`...`, ``...``). Fences first so a
# stray backtick inside a block cannot pair with one outside it.
_FENCED_BLOCK = re.compile(r"^(?P<fence>```|~~~).*?^(?P=fence)", re.DOTALL | re.MULTILINE)
_INLINE_CODE = re.compile(r"(?P<ticks>`+)(?:(?!(?P=ticks)).)*(?P=ticks)", re.DOTALL)

# Directories that are not ours to police: dependency trees and build output.
_SKIP_DIRS = {".venv", "venv", "node_modules", ".git", "build", "dist"}


def _strip_code(text: str) -> str:
    """Remove fenced blocks and inline code spans; link syntax inside is prose."""
    return _INLINE_CODE.sub(" ", _FENCED_BLOCK.sub(" ", text))


def _markdown_files() -> list[Path]:
    """Every tracked markdown file, dependency and build trees excluded."""
    return sorted(
        path
        for path in ROOT_DIR.rglob("*.md")
        if not _SKIP_DIRS & set(path.relative_to(ROOT_DIR).parts)
    )


def _display(path: Path) -> str:
    """Render a resolved path for the failure message, repo-relative when it can be.

    A broken link can resolve *outside* the repo — ``](/docs/ROADMAP.md)``
    written as if the root were the web root (pathlib treats the leading slash
    as absolute and discards the prefix), or enough ``../`` to climb out. Those
    are exactly the cases this test exists to report, so the display path must
    not be the thing that raises: a bare ``relative_to`` would turn a reportable
    broken link into a ValueError and lose every other finding in the file.
    """
    try:
        return str(path.relative_to(ROOT_DIR))
    except ValueError:
        return str(path)


def _slug(heading: str) -> str:
    """Render a heading the way GitHub renders it into an ``#anchor``.

    Lowercase, drop every character outside ``[a-z0-9 _-]``, spaces to hyphens.
    The underscore is the one piece of punctuation that survives, which is not a
    detail: half the headings in DECISIONS.md name a function or a defined range
    (``Model_Context``, ``SS_Total``, ``Sample_Include``), and a stripper that
    ate the underscore would report every link into them as broken.
    """
    text = _LINK_IN_HEADING.sub(r"\1", heading)
    text = re.sub(r"[`*~]", "", text).lower()
    return re.sub(r"[^a-z0-9 _-]", "", text).strip().replace(" ", "-")


def _heading_slugs(text: str) -> set[str]:
    """Every anchor ``text`` defines, including GitHub's ``-1``/``-2`` duplicates."""
    seen: Counter[str] = Counter()
    slugs: set[str] = set()
    fence: str | None = None

    for line in text.splitlines():
        if fence is not None:
            if line.strip().startswith(fence):
                fence = None
            continue

        opening = _FENCE.match(line)
        if opening:
            fence = opening.group(1)
            continue

        heading = _ATX_HEADING.match(line)
        if heading:
            base = _slug(heading.group(1))
            seen[base] += 1
            slugs.add(base if seen[base] == 1 else f"{base}-{seen[base] - 1}")

    return slugs


def _broken_anchors(text: str, source_path: Path) -> list[str]:
    """Return a report line per ``#anchor`` with no matching heading in its target.

    A link whose *file* half is broken is not reported here — that is the other
    test's finding, and reporting it twice would just make one typo fail two
    tests with two different messages.
    """
    broken: list[str] = []
    for target, anchor in _ANCHORED_LINK.findall(_strip_code(text)):
        resolved = (source_path.parent / target).resolve() if target else source_path
        if not resolved.is_file():
            continue
        if anchor.lower() not in _heading_slugs(resolved.read_text(encoding="utf-8")):
            broken.append(f"{target or '(this file)'}#{anchor}")
    return broken


def _broken_links(text: str, source_dir: Path) -> list[str]:
    """Return a report line per unresolvable relative markdown link in ``text``."""
    broken: list[str] = []
    for target in _MARKDOWN_LINK.findall(_strip_code(text)):
        if target.startswith(("http://", "https://")):
            continue
        # Relative to the LINKING FILE's directory, which is the whole point:
        # a bare filename in docs/ means docs/, not the repo root.
        resolved = (source_dir / target).resolve()
        if not resolved.is_file():
            broken.append(f"{target} -> {_display(resolved)}")
    return broken


def test_there_are_markdown_files_to_check() -> None:
    """Guard against the glob silently matching nothing and vacuously passing."""
    assert len(_markdown_files()) >= 10


@pytest.mark.parametrize(
    "escaping_link",
    ["/docs/ROADMAP.md", "../../../outside.md"],
    ids=["repo-root-absolute", "dotdot-escape"],
)
def test_links_resolving_outside_the_repo_are_reported_not_raised(
    escaping_link: str,
) -> None:
    """A link that escapes the repo is a finding, never a ValueError."""
    broken = _broken_links(f"see [it]({escaping_link})", ROOT_DIR / "docs")
    assert len(broken) == 1
    assert broken[0].startswith(f"{escaping_link} -> ")


@pytest.mark.parametrize(
    "markdown_path", _markdown_files(), ids=lambda p: str(p.relative_to(ROOT_DIR))
)
def test_relative_markdown_links_resolve(markdown_path: Path) -> None:
    """Each ``](*.md)`` target exists, resolved relative to the linking file."""
    source = markdown_path.relative_to(ROOT_DIR)
    broken = _broken_links(
        markdown_path.read_text(encoding="utf-8"), markdown_path.parent
    )

    assert not broken, f"{source} has unresolvable markdown link(s): " + ", ".join(
        broken
    )


def test_slug_keeps_underscores_and_drops_other_punctuation() -> None:
    """The rule the whole anchor half rests on, pinned on a real heading."""
    assert (
        _slug("### `Model_Context` — a bounded, materialized cache of spec-derived scalars")
        == "model_context--a-bounded-materialized-cache-of-spec-derived-scalars"
    )
    # The em dash is removed rather than replaced, so the spaces on either side
    # of it collapse to two hyphens. Getting this wrong reports every such
    # heading — most of DECISIONS.md — as a broken target.
    assert _slug("v2.1 — Sequence") == "v21--sequence"


def test_slug_uses_a_headings_link_text_not_its_url() -> None:
    assert _slug("See [the ladder](ROADMAP.md#ladder)") == "see-the-ladder"


def test_repeated_headings_get_githubs_numeric_suffixes() -> None:
    slugs = _heading_slugs("## Notes\n\n## Notes\n\n## Notes\n")
    assert slugs == {"notes", "notes-1", "notes-2"}


def test_headings_inside_fenced_blocks_are_not_anchors() -> None:
    """A shell comment in a fenced block is not a heading, and neither is a
    markdown example that shows one."""
    slugs = _heading_slugs(
        "# Real Heading\n\n```bash\n# not a heading\n```\n\n~~~md\n## also not\n~~~\n"
    )
    assert slugs == {"real-heading"}


def test_an_anchor_into_a_missing_file_is_left_to_the_path_check(tmp_path: Path) -> None:
    """One typo should fail one test, not two with two different messages."""
    source = tmp_path / "doc.md"
    source.write_text("see [it](nowhere.md#whatever)", encoding="utf-8")
    assert _broken_anchors(source.read_text(encoding="utf-8"), source) == []


def test_there_are_anchored_links_to_check() -> None:
    """Guard against the anchor half vacuously passing on zero links."""
    total = sum(
        len(_ANCHORED_LINK.findall(_strip_code(path.read_text(encoding="utf-8"))))
        for path in _markdown_files()
    )
    assert total >= 100


@pytest.mark.parametrize(
    "markdown_path", _markdown_files(), ids=lambda p: str(p.relative_to(ROOT_DIR))
)
def test_link_anchors_resolve_to_a_heading(markdown_path: Path) -> None:
    """Each ``#anchor`` names a heading that exists in the document it points at."""
    source = markdown_path.relative_to(ROOT_DIR)
    broken = _broken_anchors(markdown_path.read_text(encoding="utf-8"), markdown_path)

    assert not broken, f"{source} links to missing heading(s): " + ", ".join(broken)

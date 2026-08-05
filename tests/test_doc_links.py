"""Every relative markdown link in the repo's documentation resolves to a real file.

This is the cheaper half of the documentation-drift check CONTRIBUTING.md has
long proposed. It exists because the failure mode is silent: a link that points
at a file which is not there still renders as a link, and only 404s for whoever
clicks it.

The two breakages it would have caught, both real:

* **Deleted targets.** ``docs/REVIEW.md`` was removed while four documents still
  linked to it.
* **Wrong-depth prefixes.** A pass that moved the planning docs under ``docs/``
  prefixed every relative link with ``docs/`` — including links *inside*
  ``docs/``, where the prefix is one level too many. ``docs/TODOs.md`` pointing
  at ``docs/ROADMAP.md`` resolves to ``docs/docs/ROADMAP.md``. That single
  commit broke 171 links, and nothing failed.

Anchors (the ``#section`` half) are deliberately not checked here — that is the
other, more involved half of the proposal. This one only asks whether the file
on the other end exists.

Code spans are stripped before scanning. Documentation that *describes* link
syntax writes things like ``](FILE.md#anchor)`` inside backticks, and those are
examples of a notation, not links to a document — treating them as links would
make the docs unable to talk about their own conventions.

Pure ``re`` + ``pathlib``: no Excel, so it runs in the Linux CI job.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest

ROOT_DIR = Path(__file__).resolve().parents[1]

# ``](target.md)`` or ``](target.md#anchor)`` — capture the path half only.
_MARKDOWN_LINK = re.compile(r"\]\(([^)\s#]+\.md)(?:#[^)]*)?\)")

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

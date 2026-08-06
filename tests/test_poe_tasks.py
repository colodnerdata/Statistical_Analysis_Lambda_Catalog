"""Pins the poethepoet task table in pyproject.toml.

The tasks under ``[tool.poe.tasks]`` are the repo's command surface: CI invokes
``test-cov`` and ``lint`` by name, and CONTRIBUTING.md documents the rest. A
silent rename or deletion would break the workflow file or the docs without any
test noticing, so the task names are pinned here the same way
``_EXPECTED_CASE_NAMES`` pins the regression spec cases — adding, renaming, or
dropping a task means editing this list in the same commit.

Pure TOML reads; no Excel, no poethepoet import, no subprocess.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest

# tomllib is stdlib from 3.11; this project supports 3.10, where the tomli
# backport (dev group, same API) stands in for it.
if sys.version_info >= (3, 11):
    import tomllib
else:
    import tomli as tomllib

_PYPROJECT = Path(__file__).resolve().parent.parent / "pyproject.toml"

# Ordered exactly as they appear in pyproject.toml.
_EXPECTED_TASK_NAMES = [
    "test",
    "test-cov",
    "test-excel",
    "lint",
    "check",
    "verify-headless",
    "verify-deep",
    "verify-deep-univariate",
    "verify-test-models",
    "verify",
    "build",
    "build-univariate",
    "static-sheets",
    "resync-names",
]

# The tasks .github/workflows/ci.yml runs by name. Renaming either without
# editing the workflow turns every CI run red at the "poe: unrecognised task"
# stage, which is a slower way to learn it than this assertion.
_CI_TASK_NAMES = ["test-cov", "lint"]


@pytest.fixture(scope="module")
def poe_config() -> dict[str, Any]:
    with _PYPROJECT.open("rb") as handle:
        return tomllib.load(handle)["tool"]["poe"]


@pytest.fixture(scope="module")
def tasks(poe_config: dict[str, Any]) -> dict[str, Any]:
    return poe_config["tasks"]


def test_task_names_match_the_pinned_list(tasks: dict[str, Any]) -> None:
    assert list(tasks) == _EXPECTED_TASK_NAMES


def test_ci_invoked_tasks_exist(tasks: dict[str, Any]) -> None:
    for name in _CI_TASK_NAMES:
        assert name in tasks


def test_every_task_has_help_text(tasks: dict[str, Any]) -> None:
    """Bare ``poe`` lists tasks by their help string; a task without one is
    invisible in exactly the way the old Makefile's comments were."""
    missing = [name for name, task in tasks.items() if not task.get("help")]
    assert missing == []


def test_no_task_re_enters_uv(tasks: dict[str, Any]) -> None:
    """poe's uv executor already runs each task inside the project environment.
    A task that spelled ``uv run`` itself would nest one uv invocation inside
    another and silently reintroduce the prefix the migration removed."""
    for name, task in tasks.items():
        command = task.get("cmd")
        if command is not None:
            assert not command.startswith("uv "), name


def test_lockfile_strictness_is_configured(poe_config: dict[str, Any]) -> None:
    """The Makefile passed ``--frozen`` on every recipe. UV_FROZEN carries that
    intent for every task at once; dropping it would let a stale lockfile
    resolve silently instead of failing."""
    assert poe_config["env"]["UV_FROZEN"] == "1"


def test_verify_builds_the_artifacts_before_screening_them(
    tasks: dict[str, Any],
) -> None:
    """The screen has to run LAST.

    ``verify-headless`` reads whatever is sitting in ``dist/``; the three deep
    tasks rewrite those files. The original sequence put the screen first, so
    it validated the previously committed artifacts and never looked at the
    ones the run had just built — a rebuild that broke a defined name or
    orphaned a chart relationship passed ``verify`` clean.
    """
    steps = tasks["verify"]["sequence"]

    assert steps[-1] == "verify-headless"
    assert steps[0]["parallel"] == [
        "verify-deep",
        "verify-deep-univariate",
        "verify-test-models",
    ]


def test_parallel_excel_output_is_buffered_not_interleaved(
    tasks: dict[str, Any],
) -> None:
    """Streaming three concurrent drivers line-by-line would shred exactly the
    transcripts ``excel-only-runs/`` exists to preserve."""
    assert tasks["verify"]["sequence"][0]["output_mode"] == "buffer"


def test_parallel_tasks_need_a_poe_that_has_them(poe_config: dict[str, Any]) -> None:
    """``parallel`` is not in every poethepoet. An older one reads the key as an
    unknown task type and refuses to run ``verify`` at all, so the floor in the
    dev dependency group is load-bearing, not cosmetic."""
    del poe_config  # the pin lives in [dependency-groups], read separately
    with _PYPROJECT.open("rb") as handle:
        dev_group = tomllib.load(handle)["dependency-groups"]["dev"]

    poe_pin = next(spec for spec in dev_group if spec.startswith("poethepoet"))
    floor = poe_pin.split(">=")[1].split(",")[0]
    assert tuple(int(part) for part in floor.split(".")) >= (0, 48)

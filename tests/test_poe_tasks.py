"""Pins the poethepoet task table in pyproject.toml.

The tasks under ``[tool.poe.tasks]`` are the repo's command surface: CI invokes
``test-cov`` and ``lint`` by name, and CONTRIBUTING.md documents the rest. A
silent rename or deletion would break the workflow file or the docs without any
test noticing, so the task names are pinned here the same way
``_EXPECTED_CASE_NAMES`` pins the regression spec cases — adding, renaming, or
dropping a task means editing this list in the same commit.

Pure ``tomllib``; no Excel, no poethepoet import, no subprocess.
"""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import Any

import pytest

_PYPROJECT = Path(__file__).resolve().parent.parent / "pyproject.toml"

# Ordered exactly as they appear in pyproject.toml.
_EXPECTED_TASK_NAMES = [
    "test",
    "test-cov",
    "test-excel",
    "lint",
    "verify-headless",
    "verify-deep",
    "verify-deep-univariate",
    "verify-test-models",
    "verify",
    "build",
    "build-univariate",
    "qc",
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


def test_verify_sequences_both_layers(tasks: dict[str, Any]) -> None:
    assert tasks["verify"]["sequence"] == [
        "verify-headless",
        "verify-deep",
        "verify-deep-univariate",
        "verify-test-models",
    ]

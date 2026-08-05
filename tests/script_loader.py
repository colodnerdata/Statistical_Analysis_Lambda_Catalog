"""Import a ``scripts/`` build driver by path, for tests that exercise it directly.

The ``build_*`` entry points live in ``scripts/`` so they sit alongside the
package rather than spilling into the repo root. They are standalone scripts,
not modules of ``lambda_catalog``. Tests that replace a driver's attributes or
call its ``main()`` resolve it explicitly here instead — one helper, so every
test file loads a driver the same way and a missing script produces the same
clear error.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"


def load_script_module(name: str) -> ModuleType:
    """Load ``scripts/<name>.py`` and register it in ``sys.modules`` as ``name``.

    Raises ``ImportError`` naming the path when the script is missing or cannot
    produce a loader, rather than letting the failure surface later as an
    ``AttributeError`` on ``spec.loader``.

    The module object is bound into ``sys.modules[name]`` *before* it is
    executed — the loader protocol's own contract — so a module that imports
    itself, or a later ``import <name>`` anywhere in the suite, sees exactly the
    object this call executed. A previously registered module under the same
    name is replaced, so the returned module is always the freshly executed one.
    """
    path = SCRIPTS_DIR / f"{name}.py"
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load {name} from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except BaseException:
        sys.modules.pop(name, None)
        raise
    return module

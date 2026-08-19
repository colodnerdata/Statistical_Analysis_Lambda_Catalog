"""Regenerate every sheet in ``templates/static_sheets.xlsx`` in one shot.

The static reference sheets (Regression Instructions, Diagnostic Guide) are
dataset-independent, so ``build_production.py`` copies them out of this
template instead of re-running their COM writes on every build
(see CONTRIBUTING.md -> "Static reference sheets"). Each sheet's authored
content lives in its own module -- ``_ROWS`` in
``write_sheet_regression_instructions.py``, the body of
``_write_template_sheet`` in ``write_sheet_diagnostic_guide.py``. Editing one sheet's content and
forgetting the regenerate step -- or regenerating only one of two sheets
edited together -- silently ships stale template text (see DECISIONS.md ->
"Static template drift").

This script opens the template once, rebuilds every static sheet from its
current Python source, and saves once -- there is exactly one command to
run (and to remember) after editing any static sheet's content. The
per-module CLIs (``python -m lambda_catalog.write_sheet_regression_instructions``,
``python -m lambda_catalog.write_sheet_diagnostic_guide``) still work
standalone for targeted debugging of a single sheet.
"""
from __future__ import annotations

import argparse
from collections.abc import Callable
from pathlib import Path

import xlwings as xw

from lambda_catalog.workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    open_or_create_workbook,
    raise_excel_access_error,
)
from lambda_catalog.write_sheet_diagnostic_guide import (
    SHEET_NAME as _DIAGNOSTIC_GUIDE_SHEET_NAME,
)
from lambda_catalog.write_sheet_diagnostic_guide import (
    TEMPLATE_PATH as _DEFAULT_TEMPLATE_PATH,
)
from lambda_catalog.write_sheet_diagnostic_guide import (
    _write_template_sheet as _write_diagnostic_guide_template,
)
from lambda_catalog.write_sheet_regression_instructions import (
    SHEET_NAME as _REGRESSION_INSTRUCTIONS_SHEET_NAME,
)
from lambda_catalog.write_sheet_regression_instructions import (
    _write_template_sheet as _write_regression_instructions_template,
)

# Each entry's writer is that module's own ``_write_template_sheet`` -- the
# same function its individual CLI calls -- so this script can never drift
# from what running the per-module CLIs would produce.
_STATIC_SHEET_WRITERS: tuple[tuple[str, Callable[[xw.Book], None]], ...] = (
    (_REGRESSION_INSTRUCTIONS_SHEET_NAME, _write_regression_instructions_template),
    (_DIAGNOSTIC_GUIDE_SHEET_NAME, _write_diagnostic_guide_template),
)


def _main() -> None:
    parser = argparse.ArgumentParser(
        description="Rebuild every static reference sheet inside templates/static_sheets.xlsx."
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=_DEFAULT_TEMPLATE_PATH,
        help="Path to the static template workbook (default: templates/static_sheets.xlsx).",
    )
    args = parser.parse_args()

    template_path = args.template.resolve()
    try:
        with xw.App(visible=True, add_book=False) as app:
            workbook, _ = open_or_create_workbook(app, template_path)
            try:
                for sheet_name, write_template in _STATIC_SHEET_WRITERS:
                    write_template(workbook)
                    print(f"Template sheet updated: {sheet_name}")
                workbook.save(str(template_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(template_path, "open or save", exc)

    print(f"Template workbook: {template_path}")


if __name__ == "__main__":
    _main()

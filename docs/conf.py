"""Sphinx configuration for the Lambda Catalog tutorial site.

Run order matters: ``uv run --group docs poe docs`` regenerates the pages
under ``docs/generated/`` from the repo's authored sheet lists
(``scripts/generate_docs_pages.py``) and then runs ``sphinx-build -W``.
The generated pages import the same module-level content lists that write
the workbook's static sheets, so the docs cannot drift from the sheets.

The hand-written narrative lives at the ``docs/`` root; the existing
planning documents (ARCHITECTURE.md, ROADMAP.md, ...) are contributor
reference, not tutorial pages, and are excluded from the site's ToC.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from excel_lexer import ExcelFormulaLexer  # noqa: E402

project = "Statistical Analysis Lambda Catalog"
author = "Statistical Analysis Lambda Catalog contributors"
copyright = "2026, Statistical Analysis Lambda Catalog contributors"  # noqa: A001

# -- Extensions ----------------------------------------------------------------

extensions = [
    "myst_parser",
    "sphinx_copybutton",
    "sphinx_design",
]

# MyST: allow headers to be linked by number depth, and permit raw targets
myst_heading_anchors = 3

# -- Theme ---------------------------------------------------------------------

html_theme = "sphinx_book_theme"
html_static_path = ["_static"]
html_css_files = ["custom.css"]
html_title = "Lambda Catalog — a guided tour for Excel users"

# Sidebar stays usable at tutorial length; the site is one linear story.
html_theme_options = {
    "repository_url": "https://github.com/colodnerdata/Statistical_Analysis_Lambda_Catalog",
    "use_repository_button": True,
    "use_issues_button": True,
    "home_page_in_toc": True,
}

# -- Source layout --------------------------------------------------------------

# The repo's planning docs are contributor reference, not tutorial pages.
exclude_patterns = [
    "_build",
    "excel_lexer.py",
    "ARCHITECTURE.md",
    "DECISIONS.md",
    "MODEL_TESTING_ASSETS.md",
    "ROADMAP.md",
    "TODOs.md",
]

# -- Highlighting ---------------------------------------------------------------

# The custom Excel-formula lexer: code fences tagged ``excel`` render the
# workbook's actual formula vocabulary instead of a generic look.
from sphinx.highlighting import lexers  # noqa: E402

_excel = ExcelFormulaLexer()
lexers["excel"] = _excel
lexers["excel-formula"] = _excel
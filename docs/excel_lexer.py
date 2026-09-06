"""A Pygments lexer for Excel worksheet formulas.

Registered in ``conf.py`` as the ``excel`` language, so MyST fences like

    ```{code-block} excel
    =Multiple_R(Fit_Design_Columns(), Design_Response())
    ```

highlight the workbook's formula vocabulary — LAMBDA/LET/BYROW and the
dynamic-array functions as keywords, Name-Manager function names as
functions, A1-style cell references as labels — instead of falling back to
a generic look.
"""
from __future__ import annotations

from pygments.lexer import RegexLexer, words
from pygments import token

# The formula-language keywords and dynamic-array functions the workbook
# leans on. Matched only when followed by "(" so ordinary words sharing a
# prefix stay ordinary; words() matches longest-first, which keeps IF
# from stealing IFERROR.
_KEYWORDS = (
    "LAMBDA", "LET",
    "BYROW", "BYCOL", "MAP", "REDUCE", "SCAN",
    "SEQUENCE", "TAKE", "DROP", "EXPAND", "CHOOSEROWS", "CHOOSECOLS",
    "HSTACK", "VSTACK", "TOCOL", "TOROW", "WRAPROWS",
    "FILTER", "SORT", "SORTBY", "UNIQUE",
    "XMATCH", "XLOOKUP",
    "IFERROR", "IFNA", "IF",
    "TRANSPOSE", "SUMPRODUCT",
)

_TRUE_FALSE = ("TRUE", "FALSE")


class ExcelFormulaLexer(RegexLexer):
    """Tokenize a single Excel formula (the leading ``=`` included)."""

    name = "Excel formula"
    aliases = ("excel", "excel-formula", "xlsx")
    filenames = ()
    tokens = {
        "root": [
            # Keywords must be tried BEFORE the generic function-name rule.
            (words(_KEYWORDS, suffix=r"(?=\s*\()"), token.Keyword),
            (words(_TRUE_FALSE, suffix=r"\b"), token.Keyword.Constant),
            # Sheet-qualified references: 'Regression'!$B$4
            (r"'[^']*'!", token.Name.Namespace),
            # A1-style cell refs with optional $ anchors ($AB$9). The
            # lookahead stops a ref from matching the head of a longer
            # identifier (NLL_Beta's leading "N" must not eat "NLL").
            (r"\$?[A-Z]{1,3}\$?\d+(?![A-Za-z0-9_])", token.Name.Label),
            # Function call: any identifier directly followed by "(" —
            # covers Name-Manager functions (Multiple_R, NLL_Beta) and the
            # built-ins alike.
            (r"[A-Za-z_][A-Za-z0-9_.]*(?=\s*\()", token.Name.Function),
            (r'"[^"]*"', token.String.Double),
            (r"\b\d+(\.\d+)?([Ee][+-]?\d+)?\b", token.Number),
            # Bare identifiers — named ranges and bare arguments that are
            # neither calls nor cell refs (Spec_Role, cutoff, z).
            (r"[A-Za-z_][A-Za-z0-9_.]*", token.Name),
            (r"[(),:=+\-*/^&<>{}\[\];!#]", token.Operator),
            (r"\s+", token.Whitespace),
        ],
    }
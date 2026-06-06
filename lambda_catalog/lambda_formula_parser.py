"""Parse and translate LAMBDA formula display text into workbook.xml token syntax."""
from __future__ import annotations


XML_FUNCTION_PREFIXES = {
    "BYROW": "_xlfn.BYROW",
    "CHOOSECOLS": "_xlfn.CHOOSECOLS",
    "CHOOSEROWS": "_xlfn.CHOOSEROWS",
    "DROP": "_xlfn.DROP",
    "FILTER": "_xlfn._xlws.FILTER",
    "HSTACK": "_xlfn.HSTACK",
    "ISOMITTED": "_xlfn.ISOMITTED",
    "LAMBDA": "_xlfn.LAMBDA",
    "LET": "_xlfn.LET",
    "MAKEARRAY": "_xlfn.MAKEARRAY",
    "F.DIST.RT": "_xlfn.F.DIST.RT",
    "NORM.S.INV": "_xlfn.NORM.S.INV",
    "RANK.AVG": "_xlfn.RANK.AVG",
    "SEQUENCE": "_xlfn.SEQUENCE",
    "SORT": "_xlfn._xlws.SORT",
    "T.DIST.2T": "_xlfn.T.DIST.2T",
    "T.INV.2T": "_xlfn.T.INV.2T",
    "TAKE": "_xlfn.TAKE",
    "VSTACK": "_xlfn.VSTACK",
}


def _normalize_user_formula(formula: str) -> str:
    """Strip a leading ``=`` sign from a user-facing Excel formula.

    Parameters
    ----------
    formula : str
        Raw formula string, optionally starting with ``=``.

    Returns
    -------
    str
        The formula with a leading ``=`` removed and surrounding whitespace
        stripped.
    """
    formula = formula.strip()
    if formula.startswith("="):
        return formula[1:]
    return formula


def to_workbook_xml_formula(formula: str) -> str:
    """Convert a user-facing LAMBDA formula into Excel's workbook.xml syntax.

    Parameters
    ----------
    formula : str
        A user-facing LAMBDA formula string, with or without a leading ``=``.

    Returns
    -------
    str
        The formula rewritten with ``_xlfn.LAMBDA``, ``_xlpm.*`` parameter
        prefixes, and ``_xlfn.*`` function prefixes as required by
        workbook.xml.

    Raises
    ------
    ValueError
        If the formula does not conform to the expected LAMBDA structure.
    """
    formula = _normalize_user_formula(formula)

    parameters, body = _split_lambda_signature(formula)
    bound_names: dict[str, str] = {}
    xml_parameters: list[str] = []

    for parameter in parameters:
        parameter = parameter.strip()
        is_optional = parameter.startswith("[") and parameter.endswith("]")
        parameter_name = parameter[1:-1].strip() if is_optional else parameter
        if not parameter_name:
            raise ValueError(f"Invalid LAMBDA parameter in formula: {formula}")

        parameter_prefix = "_xlop" if is_optional else "_xlpm"
        xml_parameters.append(f"{parameter_prefix}.{parameter_name}")
        bound_names[parameter_name] = f"_xlpm.{parameter_name}"

    xml_body = _render_xml_expression(body, bound_names)
    return f"_xlfn.LAMBDA({','.join(xml_parameters + [xml_body])})"


def to_workbook_xml_formula_from_display(formula_display: str) -> str:
    """Convert a multi-line display formula into workbook.xml syntax.

    Parameters
    ----------
    formula_display : str
        Human-readable LAMBDA formula, potentially containing indentation
        and newlines for readability.

    Returns
    -------
    str
        The compacted formula rewritten in workbook.xml token syntax.
    """
    return to_workbook_xml_formula(
        _strip_non_string_whitespace(_normalize_user_formula(formula_display))
    )


def _strip_non_string_whitespace(text: str) -> str:
    """Remove formatting whitespace while preserving spaces inside string literals.

    Parameters
    ----------
    text : str
        Formula text that may contain string literals and formatting whitespace.

    Returns
    -------
    str
        The text with all whitespace outside string literals removed.
    """
    compacted: list[str] = []
    index = 0

    while index < len(text):
        character = text[index]
        if character == '"':
            string_end = _consume_string_literal(text, index)
            compacted.append(text[index:string_end])
            index = string_end
            continue

        if character.isspace():
            index += 1
            continue

        compacted.append(character)
        index += 1

    return "".join(compacted)


def _split_lambda_signature(formula: str) -> tuple[list[str], str]:
    """Split a top-level LAMBDA formula into its parameter list and body.

    Parameters
    ----------
    formula : str
        A whitespace-free LAMBDA formula starting with ``LAMBDA(`` and
        ending with ``)``.

    Returns
    -------
    tuple[list[str], str]
        A 2-tuple of (parameters, body) where parameters is the list of
        parameter tokens and body is the final expression.

    Raises
    ------
    ValueError
        If the formula does not start with ``LAMBDA(...)`` or contains
        fewer than one parameter and a body.
    """
    if not formula.startswith("LAMBDA(") or not formula.endswith(")"):
        raise ValueError(f"Expected formula to start with LAMBDA(...): {formula}")

    inner = formula[len("LAMBDA("):-1]
    parts: list[str] = []
    current: list[str] = []
    depth = 0

    index = 0
    while index < len(inner):
        character = inner[index]

        if character == '"':
            string_end = _consume_string_literal(inner, index)
            current.append(inner[index:string_end])
            index = string_end
            continue

        if character == "(":
            depth += 1
        elif character == ")":
            depth -= 1

        if character == "," and depth == 0:
            parts.append("".join(current).strip())
            current = []
            index += 1
            continue

        current.append(character)
        index += 1

    parts.append("".join(current).strip())
    if len(parts) < 2:
        raise ValueError(
            f"LAMBDA formula must contain at least one parameter and a body: {formula}"
        )

    return parts[:-1], parts[-1]


def _render_xml_expression(expression: str, bound_names: dict[str, str]) -> str:
    """Rewrite a formula expression using workbook.xml tokens and bound-name references.

    Parameters
    ----------
    expression : str
        A single formula expression, possibly containing function calls,
        identifiers, and string literals.
    bound_names : dict[str, str]
        Mapping from plain parameter names to their ``_xlpm.*`` prefixed
        XML counterparts.

    Returns
    -------
    str
        The expression rewritten with XML prefixes and bound-name substitutions.
    """
    rendered: list[str] = []
    index = 0

    while index < len(expression):
        character = expression[index]

        if character == '"':
            string_end = _consume_string_literal(expression, index)
            rendered.append(expression[index:string_end])
            index = string_end
            continue

        if character.isalpha() or character == "_":
            token_end = index + 1
            while token_end < len(expression) and (
                expression[token_end].isalnum() or expression[token_end] in "_."
            ):
                token_end += 1

            token = expression[index:token_end]
            next_non_space = token_end
            while next_non_space < len(expression) and expression[next_non_space].isspace():
                next_non_space += 1

            if (
                token == "LET"
                and next_non_space < len(expression)
                and expression[next_non_space] == "("
            ):
                call_end = _find_matching_paren(expression, next_non_space)
                rendered.append(
                    _render_let_call(expression[next_non_space + 1:call_end], bound_names)
                )
                index = call_end + 1
                continue

            if (
                token == "LAMBDA"
                and next_non_space < len(expression)
                and expression[next_non_space] == "("
            ):
                call_end = _find_matching_paren(expression, next_non_space)
                rendered.append(
                    _render_inner_lambda_call(
                        expression[next_non_space + 1:call_end], bound_names
                    )
                )
                index = call_end + 1
                continue

            if token in bound_names:
                rendered.append(bound_names[token])
                index = token_end
                continue

            if (
                token in XML_FUNCTION_PREFIXES
                and next_non_space < len(expression)
                and expression[next_non_space] == "("
            ):
                rendered.append(XML_FUNCTION_PREFIXES[token])
                index = token_end
                continue

            rendered.append(token)
            index = token_end
            continue

        rendered.append(character)
        index += 1

    return "".join(rendered)


def _render_inner_lambda_call(arguments_text: str, bound_names: dict[str, str]) -> str:
    """Render a LAMBDA(...) call that appears inside a larger expression body.

    Parameters
    ----------
    arguments_text : str
        The raw argument text inside the inner LAMBDA(...) call, excluding the
        outer parentheses.
    bound_names : dict[str, str]
        Mapping of currently in-scope names to their ``_xlpm.*`` prefixed
        XML counterparts (from the enclosing LAMBDA and any enclosing LET bindings).

    Returns
    -------
    str
        The fully translated ``_xlfn.LAMBDA(...)`` expression with the inner
        parameters added to the bound-name scope for the body.
    """
    arguments = _split_top_level_arguments(arguments_text)
    if len(arguments) < 2:
        raise ValueError(
            f"LAMBDA requires at least one parameter and a body: LAMBDA({arguments_text})"
        )

    inner_bound_names = dict(bound_names)
    xml_parameters: list[str] = []

    for param in arguments[:-1]:
        param = param.strip()
        is_optional = param.startswith("[") and param.endswith("]")
        param_name = param[1:-1].strip() if is_optional else param
        if not param_name:
            raise ValueError(f"Invalid inner LAMBDA parameter: '{param}'")
        prefix = "_xlop" if is_optional else "_xlpm"
        xml_parameters.append(f"{prefix}.{param_name}")
        inner_bound_names[param_name] = f"_xlpm.{param_name}"

    body = _render_xml_expression(arguments[-1], inner_bound_names)
    return f"_xlfn.LAMBDA({','.join(xml_parameters + [body])})"


def _render_let_call(arguments_text: str, bound_names: dict[str, str]) -> str:
    """Translate a LET call, adding each local binding before later references use it.

    Parameters
    ----------
    arguments_text : str
        The raw argument text inside the LET(...) call, excluding the outer
        parentheses.
    bound_names : dict[str, str]
        Mapping of currently in-scope names to their XML counterparts.

    Returns
    -------
    str
        The fully translated ``_xlfn.LET(...)`` expression.

    Raises
    ------
    ValueError
        If the LET call contains fewer than three arguments or has malformed
        name/value binding pairs.
    """
    arguments = _split_top_level_arguments(arguments_text)
    if len(arguments) < 3:
        raise ValueError(
            f"LET requires at least one binding and a result expression: LET({arguments_text})"
        )

    rendered_arguments: list[str] = []
    local_bound_names = dict(bound_names)

    for index in range(0, len(arguments) - 1, 2):
        if index + 1 >= len(arguments) - 1:
            raise ValueError(
                f"LET bindings must be name/value pairs: LET({arguments_text})"
            )

        binding_name = arguments[index].strip()
        if not binding_name or not (binding_name[0].isalpha() or binding_name[0] == "_"):
            raise ValueError(f"Invalid LET binding name: {binding_name}")

        rendered_arguments.append(f"_xlpm.{binding_name}")
        rendered_arguments.append(
            _render_xml_expression(arguments[index + 1], local_bound_names)
        )
        local_bound_names[binding_name] = f"_xlpm.{binding_name}"

    rendered_arguments.append(_render_xml_expression(arguments[-1], local_bound_names))
    return f"_xlfn.LET({','.join(rendered_arguments)})"


def _split_top_level_arguments(text: str) -> list[str]:
    """Split comma-separated arguments while respecting nested calls and string literals.

    Parameters
    ----------
    text : str
        Comma-separated argument text at the top level of a function call.

    Returns
    -------
    list[str]
        List of individual argument strings, stripped of surrounding whitespace.
    """
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    index = 0

    while index < len(text):
        character = text[index]

        if character == '"':
            string_end = _consume_string_literal(text, index)
            current.append(text[index:string_end])
            index = string_end
            continue

        if character == "(":
            depth += 1
        elif character == ")":
            depth -= 1

        if character == "," and depth == 0:
            parts.append("".join(current).strip())
            current = []
            index += 1
            continue

        current.append(character)
        index += 1

    parts.append("".join(current).strip())
    return parts


def _consume_string_literal(text: str, start_index: int) -> int:
    """Return the index just past a quoted Excel string literal.

    Handles Excel's doubled-quote escape convention (``""`` inside a string).

    Parameters
    ----------
    text : str
        The full expression text.
    start_index : int
        Index of the opening double-quote character.

    Returns
    -------
    int
        Index of the first character after the closing double-quote.

    Raises
    ------
    ValueError
        If the string literal is not properly terminated.
    """
    index = start_index + 1
    while index < len(text):
        if text[index] == '"':
            if index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            return index + 1
        index += 1

    raise ValueError(f"Unterminated string literal in expression: {text}")


def _find_matching_paren(text: str, open_index: int) -> int:
    """Find the closing parenthesis that matches the opening one at open_index.

    Parameters
    ----------
    text : str
        The full expression text.
    open_index : int
        Index of the opening parenthesis character.

    Returns
    -------
    int
        Index of the matching closing parenthesis.

    Raises
    ------
    ValueError
        If no matching closing parenthesis is found.
    """
    depth = 0
    index = open_index

    while index < len(text):
        character = text[index]
        if character == '"':
            index = _consume_string_literal(text, index)
            continue
        if character == "(":
            depth += 1
        elif character == ")":
            depth -= 1
            if depth == 0:
                return index
        index += 1

    raise ValueError(f"Unmatched parenthesis in expression: {text}")

"""Parse and translate LAMBDA formula display text into workbook.xml token syntax."""
from __future__ import annotations


XML_FUNCTION_PREFIXES = {
    # ── Excel 365 dynamic-array / meta-functions ─────────────────────────────
    "ANCHORARRAY": "_xlfn.ANCHORARRAY",
    "BYCOL":       "_xlfn.BYCOL",
    "BYROW":       "_xlfn.BYROW",
    "CHOOSECOLS":  "_xlfn.CHOOSECOLS",
    "CHOOSEROWS":  "_xlfn.CHOOSEROWS",
    "DROP":        "_xlfn.DROP",
    "FILTER":      "_xlfn._xlws.FILTER",
    "HSTACK":      "_xlfn.HSTACK",
    "ISOMITTED":   "_xlfn.ISOMITTED",
    "LAMBDA":      "_xlfn.LAMBDA",
    "LET":         "_xlfn.LET",
    "MAKEARRAY":   "_xlfn.MAKEARRAY",
    "MAP":         "_xlfn.MAP",
    "RANDARRAY":   "_xlfn.RANDARRAY",
    "REDUCE":      "_xlfn.REDUCE",
    "SCAN":        "_xlfn.SCAN",
    "SEQUENCE":    "_xlfn.SEQUENCE",
    "SORT":        "_xlfn._xlws.SORT",
    "SORTBY":      "_xlfn.SORTBY",
    "TAKE":        "_xlfn.TAKE",
    "UNIQUE":       "_xlfn.UNIQUE",
    "VSTACK":       "_xlfn.VSTACK",

    # ── Excel 365 text / array utilities ─────────────────────────────────────
    "ARRAYTOTEXT": "_xlfn.ARRAYTOTEXT",
    "EXPAND":      "_xlfn.EXPAND",
    "TEXTAFTER":   "_xlfn.TEXTAFTER",
    "TEXTBEFORE":  "_xlfn.TEXTBEFORE",
    "TEXTSPLIT":   "_xlfn.TEXTSPLIT",
    "TOCOL":       "_xlfn.TOCOL",
    "TOROW":       "_xlfn.TOROW",
    "VALUETOTEXT": "_xlfn.VALUETOTEXT",
    "WRAPCOLS":    "_xlfn.WRAPCOLS",
    "WRAPROWS":    "_xlfn.WRAPROWS",

    # ── Excel 2019 / 365 new functions ───────────────────────────────────────
    "CONCAT":   "_xlfn.CONCAT",
    "IFS":      "_xlfn.IFS",
    "MAXIFS":   "_xlfn.MAXIFS",
    "MINIFS":   "_xlfn.MINIFS",
    "TEXTJOIN": "_xlfn.TEXTJOIN",
    "XLOOKUP":  "_xlfn.XLOOKUP",
    "XMATCH":   "_xlfn.XMATCH",

    # ── Excel 2013 functions ─────────────────────────────────────────────────
    "ARABIC":       "_xlfn.ARABIC",
    "BASE":         "_xlfn.BASE",
    "BITAND":       "_xlfn.BITAND",
    "BITLSHIFT":    "_xlfn.BITLSHIFT",
    "BITOR":        "_xlfn.BITOR",
    "BITRSHIFT":    "_xlfn.BITRSHIFT",
    "BITXOR":       "_xlfn.BITXOR",
    "COMBINA":      "_xlfn.COMBINA",
    "DAYS":         "_xlfn.DAYS",
    "DECIMAL":      "_xlfn.DECIMAL",
    "ENCODEURL":    "_xlfn.ENCODEURL",
    "FILTERXML":    "_xlfn.FILTERXML",
    "FORMULATEXT":  "_xlfn.FORMULATEXT",
    "IFNA":         "_xlfn.IFNA",
    "IMCOSH":       "_xlfn.IMCOSH",
    "IMCOT":        "_xlfn.IMCOT",
    "IMCSC":        "_xlfn.IMCSC",
    "IMCSCH":       "_xlfn.IMCSCH",
    "IMSEC":        "_xlfn.IMSEC",
    "IMSECH":       "_xlfn.IMSECH",
    "IMSINH":       "_xlfn.IMSINH",
    "IMTAN":        "_xlfn.IMTAN",
    "ISFORMULA":    "_xlfn.ISFORMULA",
    "MUNIT":        "_xlfn.MUNIT",
    "NUMBERVALUE":  "_xlfn.NUMBERVALUE",
    "PDURATION":    "_xlfn.PDURATION",
    "PERMUTATIONA": "_xlfn.PERMUTATIONA",
    "PHI":          "_xlfn.PHI",
    "RRI":          "_xlfn.RRI",
    "SHEET":        "_xlfn.SHEET",
    "SHEETS":       "_xlfn.SHEETS",
    "SKEW.P":       "_xlfn.SKEW.P",
    "UNICHAR":      "_xlfn.UNICHAR",
    "UNICODE":      "_xlfn.UNICODE",
    "WEBSERVICE":   "_xlfn.WEBSERVICE",
    "XOR":          "_xlfn.XOR",

    # ── Excel 2016 functions ────────────────────────────────────────────
    "FORECAST.ETS":             "_xlfn.FORECAST.ETS",
    "FORECAST.ETS.CONFINT":     "_xlfn.FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY": "_xlfn.FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT":        "_xlfn.FORECAST.ETS.STAT",
    "FORECAST.LINEAR":          "_xlfn.FORECAST.LINEAR",
    "SWITCH":                   "_xlfn.SWITCH",

    # ── Excel 2010 statistical dot-notation functions ─────────────────────────
    # All prefixes confirmed by inspecting an Excel-serialized workbook.xml.
    # Do not infer prefixes from rules of thumb — always verify against a real Excel file.
    "BETA.DIST":        "_xlfn.BETA.DIST",
    "BETA.INV":         "_xlfn.BETA.INV",
    "BINOM.DIST":       "_xlfn.BINOM.DIST",
    "BINOM.DIST.RANGE": "_xlfn.BINOM.DIST.RANGE",
    "BINOM.INV":        "_xlfn.BINOM.INV",
    "CHISQ.DIST":       "_xlfn.CHISQ.DIST",
    "CHISQ.DIST.RT":    "_xlfn.CHISQ.DIST.RT",
    "CHISQ.INV":        "_xlfn.CHISQ.INV",
    "CHISQ.INV.RT":     "_xlfn.CHISQ.INV.RT",
    "CHISQ.TEST":       "_xlfn.CHISQ.TEST",
    "CONFIDENCE.NORM":  "_xlfn.CONFIDENCE.NORM",
    "CONFIDENCE.T":     "_xlfn.CONFIDENCE.T",
    "COVARIANCE.P":     "_xlfn.COVARIANCE.P",
    "COVARIANCE.S":     "_xlfn.COVARIANCE.S",
    "EXPON.DIST":       "_xlfn.EXPON.DIST",
    "F.DIST":           "_xlfn.F.DIST",
    "F.DIST.RT":        "_xlfn.F.DIST.RT",   # confirmed _xlfn in workbook.xml
    "F.INV":            "_xlfn.F.INV",
    "F.INV.RT":         "_xlfn.F.INV.RT",
    "F.TEST":           "_xlfn.F.TEST",
    "GAMMA.DIST":       "_xlfn.GAMMA.DIST",
    "GAMMA.INV":        "_xlfn.GAMMA.INV",
    "GAMMALN.PRECISE":  "_xlfn.GAMMALN.PRECISE",
    "HYPGEOM.DIST":     "_xlfn.HYPGEOM.DIST",
    "LOGNORM.DIST":     "_xlfn.LOGNORM.DIST",
    "LOGNORM.INV":      "_xlfn.LOGNORM.INV",
    "MODE.MULT":        "_xlfn.MODE.MULT",
    "MODE.SNGL":        "_xlfn.MODE.SNGL",
    "NEGBINOM.DIST":    "_xlfn.NEGBINOM.DIST",
    "NORM.DIST":        "_xlfn.NORM.DIST",
    "NORM.INV":         "_xlfn.NORM.INV",
    "NORM.S.DIST":      "_xlfn.NORM.S.DIST",
    "NORM.S.INV":       "_xlfn.NORM.S.INV",  # confirmed _xlfn in workbook.xml
    "PERCENTILE.EXC":   "_xlfn.PERCENTILE.EXC",
    "PERCENTILE.INC":   "_xlfn.PERCENTILE.INC",
    "PERCENTRANK.EXC":  "_xlfn.PERCENTRANK.EXC",
    "PERCENTRANK.INC":  "_xlfn.PERCENTRANK.INC",
    "POISSON.DIST":     "_xlfn.POISSON.DIST",
    "QUARTILE.EXC":     "_xlfn.QUARTILE.EXC",
    "QUARTILE.INC":     "_xlfn.QUARTILE.INC",
    "RANK.AVG":         "_xlfn.RANK.AVG",    # confirmed _xlfn in workbook.xml
    "RANK.EQ":          "_xlfn.RANK.EQ",
    "STDEV.P":          "_xlfn.STDEV.P",
    "STDEV.S":          "_xlfn.STDEV.S",
    "T.DIST":           "_xlfn.T.DIST",
    "T.DIST.2T":        "_xlfn.T.DIST.2T",  # confirmed _xlfn in workbook.xml
    "T.DIST.RT":        "_xlfn.T.DIST.RT",
    "T.INV":            "_xlfn.T.INV",
    "T.INV.2T":         "_xlfn.T.INV.2T",   # confirmed _xlfn in workbook.xml
    "T.TEST":           "_xlfn.T.TEST",
    "VAR.P":            "_xlfn.VAR.P",
    "VAR.S":            "_xlfn.VAR.S",
    "WEIBULL.DIST":     "_xlfn.WEIBULL.DIST",
    "Z.TEST":           "_xlfn.Z.TEST",

    # ── Excel 2010 math / date precision variants ────────────────────────────
    "CEILING.MATH":     "_xlfn.CEILING.MATH",
    "CEILING.PRECISE":  "_xlfn.CEILING.PRECISE",
    "ERF.PRECISE":      "_xlfn.ERF.PRECISE",
    "ERFC.PRECISE":     "_xlfn.ERFC.PRECISE",
    "FLOOR.MATH":       "_xlfn.FLOOR.MATH",
    "FLOOR.PRECISE":    "_xlfn.FLOOR.PRECISE",
    "ISO.CEILING":      "_xlfn.ISO.CEILING",
    "NETWORKDAYS.INTL": "_xlfn.NETWORKDAYS.INTL",
    "WORKDAY.INTL":     "_xlfn.WORKDAY.INTL",
}


# Excel built-in functions that need no prefix in workbook.xml definedName formulas.
# Together with XML_FUNCTION_PREFIXES this forms the complete registry of known Excel
# built-ins; anything else encountered in a LAMBDA body is treated as a workbook-defined
# name (another LAMBDA or named range) and emitted as-is, which is also correct.
CLASSIC_FUNCTIONS: frozenset[str] = frozenset({
    # ── Math & Trigonometry ──────────────────────────────────────────────────
    "ABS", "ACOS", "ACOSH", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH",
    "CEILING", "COMBIN", "COS", "COSH", "DEGREES", "EVEN", "EXP",
    "FACT", "FACTDOUBLE", "FLOOR", "GCD", "INT", "LCM", "LN", "LOG", "LOG10",
    "MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND", "MULTINOMIAL",
    "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS", "RAND", "RANDBETWEEN",
    "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "SERIESSUM", "SIGN",
    "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL",
    "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ",
    "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "TAN", "TANH", "TRUNC",
    # ── Statistical — pre-2010 names (dot-notation replacements are in XML_FUNCTION_PREFIXES) ──
    "AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS",
    "BETADIST", "BETAINV", "BINOMDIST",
    "CHIDIST", "CHIINV", "CHITEST",
    "CONFIDENCE", "CORREL",
    "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
    "COVAR", "CRITBINOM", "DEVSQ",
    "EXPONDIST", "FDIST", "FINV", "FISHER", "FISHERINV",
    "FORECAST", "FTEST",
    "GAMMADIST", "GAMMAINV", "GAMMALN",
    "GEOMEAN", "GROWTH", "HARMEAN", "HYPGEOMDIST",
    "INTERCEPT", "KURT", "LARGE",
    "LINEST", "LOGEST", "LOGINV", "LOGNORMDIST",
    "MAX", "MAXA", "MEDIAN", "MIN", "MINA", "MODE",
    "NEGBINOMDIST", "NORMDIST", "NORMINV", "NORMSDIST", "NORMSINV",
    "PEARSON", "PERCENTILE", "PERCENTRANK",
    "PERMUT", "POISSON", "PROB", "QUARTILE", "RANK",
    "RSQ", "SKEW", "SLOPE", "SMALL",
    "STANDARDIZE", "STDEV", "STDEVA", "STDEVP", "STDEVPA", "STEYX",
    "TDIST", "TINV", "TREND", "TRIMMEAN", "TTEST",
    "VAR", "VARA", "VARP", "VARPA", "WEIBULL", "ZTEST",
    # ── Lookup & Reference ────────────────────────────────────────────────────
    "ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS",
    "GETPIVOTDATA", "HLOOKUP", "HYPERLINK", "INDEX", "INDIRECT",
    "LOOKUP", "MATCH", "OFFSET", "ROW", "ROWS", "RTD", "TRANSPOSE", "VLOOKUP",
    # ── Date & Time ───────────────────────────────────────────────────────────
    "DATE", "DATEVALUE", "DAY", "DAYS360",
    "EDATE", "EOMONTH", "HOUR", "MINUTE", "MONTH",
    "NETWORKDAYS", "NOW", "SECOND", "TIME", "TIMEVALUE", "TODAY",
    "WEEKDAY", "WEEKNUM", "WORKDAY", "YEAR", "YEARFRAC",
    # ── Text ──────────────────────────────────────────────────────────────────
    "BAHTTEXT", "CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR",
    "EXACT", "FIND", "FINDB", "FIXED",
    "LEFT", "LEFTB", "LEN", "LENB", "LOWER",
    "MID", "MIDB", "PHONETIC", "PROPER",
    "REPLACE", "REPLACEB", "REPT",
    "RIGHT", "RIGHTB", "SEARCH", "SEARCHB",
    "SUBSTITUTE", "T", "TEXT", "TRIM", "UPPER", "VALUE",
    # ── Logical ───────────────────────────────────────────────────────────────
    "AND", "FALSE", "IF", "IFERROR", "NOT", "OR", "TRUE",
    # ── Information ───────────────────────────────────────────────────────────
    "CELL", "ERROR.TYPE", "INFO",
    "ISBLANK", "ISERR", "ISERROR", "ISEVEN",
    "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD",
    "ISREF", "ISTEXT", "N", "NA", "TYPE",
    # ── Financial ─────────────────────────────────────────────────────────────
    "ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC",
    "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD",
    "CUMIPMT", "CUMPRINC", "DB", "DDB", "DISC",
    "DOLLARDE", "DOLLARFR", "DURATION", "EFFECT",
    "FV", "FVSCHEDULE", "INTRATE", "IPMT", "IRR", "ISPMT",
    "MDURATION", "MIRR", "NOMINAL", "NPER", "NPV",
    "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD",
    "PMT", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT",
    "PV", "RATE", "RECEIVED", "SLN", "SYD",
    "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "VDB",
    "XIRR", "XNPV", "YIELD", "YIELDDISC", "YIELDMAT",
    # ── Engineering (pre-2013 functions) ──────────────────────────────────────
    "BESSELI", "BESSELJ", "BESSELK", "BESSELY",
    "BIN2DEC", "BIN2HEX", "BIN2OCT",
    "COMPLEX", "CONVERT",
    "DEC2BIN", "DEC2HEX", "DEC2OCT",
    "DELTA", "ERF", "ERFC", "GESTEP",
    "HEX2BIN", "HEX2DEC", "HEX2OCT",
    "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE",
    "IMCOS", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2",
    "IMPOWER", "IMPRODUCT", "IMREAL",
    "IMSIN", "IMSQRT", "IMSUB", "IMSUM",
    "OCT2BIN", "OCT2DEC", "OCT2HEX",
    # ── Cube ──────────────────────────────────────────────────────────────────
    "CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER", "CUBESET", "CUBESETCOUNT", "CUBEVALUE",
    # ── Database ──────────────────────────────────────────────────────────────
    "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN",
    "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP",
})


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

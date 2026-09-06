#!/usr/bin/env python
"""Generate the tutorial site's derived pages into docs/generated/.

Run as `uv run --group docs poe docs-generate` (or via `poe docs`, which
generates then builds). The workbook's user-facing content is authored in
module-level Python lists — the same lists that write the static reference
sheets — so the docs can be rendered FROM that single source of truth and
cannot drift from the sheets. If a source move breaks an extraction, the
assertions here fail loudly at generation time instead of the docs quietly
shipping a stale formula.

Pages produced:

- regression-instructions.md  — from write_sheet_regression_instructions._ROWS
- modeling-concepts.md        — from write_sheet_modeling_concepts._FEATURES /
                                _PLANNED_FEATURES / _COLUMNS
- diagnostic-guide.md         — from write_sheet_diagnostic_guide._TIER1 /
                                _TIER2 / _THRESHOLDS / _GUIDANCE
- spec-block.md               — headers from write_spec_block's spec header
                                pairs; allowed values from spec_layout's
                                _*_VALIDATION_LIST constants
- lambda-reference.md         — every entry of lambda_functions.json
- formula-review.md           — formulas imported or source-extracted; the
                                narrative annotation for each sample lives
                                in _FORMULA_SAMPLES below

No Excel is needed: importing the writer modules performs no COM calls.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from lambda_catalog import spec_layout as _sl
from lambda_catalog import write_sheet_diagnostic_guide as _dg
from lambda_catalog import write_sheet_modeling_concepts as _mc
from lambda_catalog import write_sheet_regression_instructions as _ri
from lambda_catalog.write_spec_block import _PERIOD_IN_USE_SPILL_FORMULA, _spec_band
from lambda_catalog.write_sheet_univariate import _weibull_profile_scale

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
GENERATED = DOCS / "generated"

GENERATED_BANNER = (
    "<!-- GENERATED FILE — do not edit. Regenerate: "
    "uv run --group docs poe docs-generate -->\n"
)


def _md_escape(text: str) -> str:
    """Escape the MyST-significant characters a sheet body may contain."""
    return text.replace("|", "\\|").replace("\n", " ")


def _write(name: str, body: str) -> None:
    GENERATED.mkdir(parents=True, exist_ok=True)
    (GENERATED / name).write_text(GENERATED_BANNER + body, encoding="utf-8")
    print(f"generated: docs/generated/{name}")


# ── Regression Instructions ───────────────────────────────────────────────────


def write_regression_instructions() -> None:
    lines = [
        "# Regression Instructions",
        "",
        "The same content as the workbook's **Regression Instructions** sheet,",
        "rendered from the authored rows (`_ROWS` in",
        "`lambda_catalog/write_sheet_regression_instructions.py`).",
        "",
    ]
    for _rownum, text, kind in _ri._ROWS:
        if kind == "heading":
            lines += [f"## {text}", ""]
        elif kind == "body":
            lines += [text, ""]
    _write("regression-instructions.md", "\n".join(lines))


# ── Modeling Concepts ─────────────────────────────────────────────────────────


def _concept_table(rows: list[list[str]], headers: list[str]) -> list[str]:
    out = ["| " + " | ".join(headers) + " |",
           "|" + "|".join("---" for _ in headers) + "|"]
    for row in rows:
        assert len(row) == len(headers), f"row width {len(row)} != {len(headers)}"
        out.append("| " + " | ".join(_md_escape(c) for c in row) + " |")
    out.append("")
    return out


def write_modeling_concepts() -> None:
    headers = [c.header for c in _mc._COLUMNS]
    lines = [
        "# Modeling Concepts",
        "",
        "The same table as the workbook's **Modeling Concepts** sheet, rendered",
        "from the authored lists in",
        "`lambda_catalog/write_sheet_modeling_concepts.py`.",
        "",
        "## Shipped features",
        "",
    ]
    lines += _concept_table(_mc._FEATURES, headers)
    lines += [
        "## Planned features",
        "",
        "Declared in the sheet's structure but not built in this release — see",
        "the workbook sheet for the same rows.",
        "",
    ]
    lines += _concept_table(_mc._PLANNED_FEATURES, headers)
    _write("modeling-concepts.md", "\n".join(lines))


# ── Diagnostic Guide ──────────────────────────────────────────────────────────


def write_diagnostic_guide() -> None:
    lines = [
        "# Diagnostic Guide",
        "",
        "The same tiers and thresholds as the workbook's **Diagnostic Guide**",
        "sheet, rendered from the authored lists in",
        "`lambda_catalog/write_sheet_diagnostic_guide.py`.",
        "",
        "Use the charts and flagged cells on the Regression sheet to assess",
        "model assumptions. Work through Tier 1 first; investigate Tier 2",
        "only when a Tier 1 plot raises a concern.",
        "",
        "## Tier 1 — review for every model",
        "",
    ]
    plot_headers = ["Plot", "X-axis", "Y-axis", "What to look for"]
    lines += _concept_table(_dg._TIER1, plot_headers)
    lines += ["## Tier 2 — investigate when Tier 1 raises a concern", ""]
    lines += _concept_table(_dg._TIER2, plot_headers)
    lines += ["## Diagnostic threshold reference", ""]
    lines += _concept_table(
        _dg._THRESHOLDS,
        ["Diagnostic", "Location on sheet", "Yellow threshold", "Red threshold"],
    )
    lines += ["## Common patterns and next steps", ""]
    lines += _concept_table(_dg._GUIDANCE, ["Pattern", "Symptom", "Next step"])
    _write("diagnostic-guide.md", "\n".join(lines))


# ── Spec-block reference ──────────────────────────────────────────────────────


def _extract_spec_headers() -> list[tuple[int, str]]:
    """The A–O header pairs, extracted from write_spec_block's header loop.

    The pairs are an inline literal in ``_write_spec_block``; pulling them
    from the source (rather than hand-copying) means a renamed header
    updates the docs or fails the single-match assertion below.
    """
    src = (ROOT / "lambda_catalog" / "write_spec_block.py").read_text(encoding="utf-8")
    m = re.search(r"for col, header in \((.*?)\):\s*\n\s*val\(", src, flags=re.S)
    assert m, "spec header pairs loop not found in write_spec_block.py"
    pairs = re.findall(r"\(_C_([A-Z_]+),\s*\"([^\"]+)\"\)", m.group(1))
    assert len(pairs) == 15, f"expected 15 spec header pairs, found {len(pairs)}"
    return [(int(getattr(_sl, f"_C_{name}")), header) for name, header in pairs]


def write_spec_block() -> None:
    headers = _extract_spec_headers()

    def _values(v: str | list[str]) -> str:
        """Validation lists are strings for xlValidateList; some may be lists."""
        return v if isinstance(v, str) else ", ".join(v)

    role = _values(_sl._ROLE_VALIDATION_LIST)
    include = _values(_sl._INCLUDE_VALIDATION_LIST)
    type_ = _values(_sl._TYPE_VALIDATION_LIST)
    transform = _values(_sl._TRANSFORM_VALIDATION_LIST)
    sequence = _values(_sl._SEQUENCE_VALIDATION_LIST)
    operation = _values(_sl._INTERACTION_OPERATION_VALIDATION_LIST)

    # Column → how the user interacts with it (authored map; the header
    # names above are extracted, the interaction note is explanatory).
    usage = {
        1: "Computed — spills the table's header names.",
        2: f"Dropdown: {role}. Exactly one Response; at most one Fixed Effects.",
        3: f"Dropdown: {include}.",
        4: f"Dropdown: {type_} (Predictor rows only).",
        5: "Typed (Categorical rows) — override the reference level; red if absent from the sample.",
        6: "Reserved — hidden, width 0, read by nothing in this release.",
        7: f"Dropdown: {transform} (Response and Continuous Predictor rows).",
        8: f"Dropdown: {sequence} — mark at most one ordering axis.",
        9: "Typed — Δ override; blank = the computed candidate.",
        10: "Computed — the override if typed, else the candidate.",
        11: "Computed — distinct level count in the analysis sample.",
        12: "Computed — the reference actually in effect.",
        13: "Dropdown over variable names — the OTHER operand of an interaction.",
        14: f"Dropdown: {operation}.",
        15: "Computed — design-matrix columns this row contributes.",
    }
    lines = [
        "# The MODEL SPECIFICATION block (columns A–O)",
        "",
        "One row per source-table column, header row 3, spec rows from row 4.",
        "Orange cells are yours to type; everything else computes. The block",
        "sizes itself from `Source_Table`, so retargeting the one name resizes",
        "every band.",
        "",
        "| Col | Header | How you use it |",
        "|---|---|---|",
    ]
    for idx, header in sorted(headers):
        lines.append(f"| {chr(ord('A') + idx - 1)} | {_md_escape(header)} | {usage[idx]} |")
    lines += ["", "## The self-sizing band every column is built on", "",
              "Each spec column is a sheet-scoped named range over a fixed",
              "16000-row band, trimmed to the table's width at calculation:",
              "", "```excel", _spec_band("Regression", _sl._C_ROLE), "```", "",
              "and the computed columns are single spills — Period In Use (J):",
              "", "```excel", _PERIOD_IN_USE_SPILL_FORMULA, "```", ""]
    _write("spec-block.md", "\n".join(lines))


# ── LAMBDA reference ──────────────────────────────────────────────────────────


def write_lambda_reference() -> None:
    data = json.loads((ROOT / "lambda_functions.json").read_text(encoding="utf-8"))
    functions = data["functions"]
    lines = [
        "# LAMBDA function reference",
        "",
        f"Every catalog entry in `lambda_functions.json` — {len(functions)}",
        "functions. Workbook-scoped names work on any sheet; sheet-scoped names",
        "(marked) are defined per-sheet.",
        "",
    ]
    for fn in sorted(functions, key=lambda f: f["name"].lower()):
        scope = fn.get("scope")
        scope_note = f" *(sheet-scoped: {scope})*" if scope else ""
        lines += [f"## `{fn['name']}`{scope_note}", ""]
        if fn.get("plain_language_summary"):
            lines += [f"**{fn['plain_language_summary']}**", ""]
        args = fn.get("arguments") or []
        if args:
            lines += ["Arguments:", ""]
            for a in args:
                lines += [f"- **{a['name']}** — {a['description']}"]
            lines += [""]
        if fn.get("description"):
            lines += [fn["description"], ""]
        if fn.get("yields"):
            lines += [f"Returns: {fn['yields']}", ""]
        if fn.get("notes"):
            lines += [fn["notes"], ""]
        if fn.get("formula_display"):
            lines += ["```excel", fn["formula_display"], "```", ""]
    _write("lambda-reference.md", "\n".join(lines))


# ── Formula review (Part 2) ───────────────────────────────────────────────────


def _extract_once(relpath: str, pattern: str, flags: int = 0) -> str:
    """Single-match regex extraction from a repo source file, fail-loud."""
    src = (ROOT / relpath).read_text(encoding="utf-8")
    m = re.search(pattern, src, flags=flags)
    assert m, f"pattern not found in {relpath}: {pattern[:60]!r}"
    return m.group(1)


def _extract_concatenated_literals(source_block: str) -> str:
    """Join the '...' fragments of a Python implicit-concatenation block.

    F-string fragments come back with their ``{...}`` placeholders intact;
    callers substitute the values the writer interpolates.
    """
    return "".join(re.findall(r'"((?:[^"\\]|\\.)*)"', source_block))


def _beta_nll_formula() -> str:
    """The Beta fit's BYROW NLL spill, from _stage_nll's authored literal.

    The anchor (the stage's grid spill cell) is an f-string interpolation;
    the shipped default layout puts Stage 1's grid at BY33, so the docs show
    `$BY$33#` and say so in the annotation.
    """
    block = _extract_once(
        "lambda_catalog/write_sheet_univariate.py",
        r"def _stage_nll\(grid_anchor: str\) -> str:\n(.*?)\n        \)",
        flags=re.S,
    )
    body = _extract_concatenated_literals(block)
    body = body.replace("{grid_anchor}", "$BY$33")
    assert body.startswith("=LET(d,FILTER"), f"unexpected Beta NLL body: {body[:40]}"
    return body


def _weibull_nll_formula() -> str:
    """The Weibull fit's BYROW profile-NLL spill (Stage 1 anchor $BP$33#).

    Composed from the same authored pieces the writer uses: the NLL call is
    the `_write_profile_fit` Weibull ``nll_formula``, the profiled scale is
    the actual `_weibull_profile_scale` closed form (both importable), and
    the surrounding template is source-extracted with the load-bearing
    fragments asserted so a writer change fails generation, not the docs.
    """
    template = _extract_once(
        "lambda_catalog/write_sheet_univariate.py",
        r"def _stage_nll\(axis_anchor: str\) -> str:\n(.*?)\n        \)",
        flags=re.S,
    )
    # Guard the composition against the template drifting: the fragments
    # below are the pieces the formula is hand-assembled from.
    for fragment in (
        "LET(x,", "BYROW(", "INDEX(r,1,1)", "IFERROR(", "1E+15",
        "nll_formula('x', 'p', partner_formula('p', 'x'))",
    ):
        assert fragment in template, f"Weibull NLL template lost {fragment!r}"
    scale = _weibull_profile_scale("p", "x")
    return (
        "=LET(x,FILTER(UV_Data,UV_Include),"
        "BYROW($BP$33#,LAMBDA(r,LET(p,INDEX(r,1,1),"
        f"IFERROR(NLL_Weibull(x,p,{scale}),1E+15)))))"
    )


def _unit_space_r_squared_formula() -> str:
    """The Unit-Space Fit block's R² — the formula that reads the AH5 toggle.

    The writer interpolates the ``_A_BACK_TRANSFORM_METHOD`` anchor (the
    `$AH$5` address, via ``_abs_ref``), so the literal fragments are joined
    and the placeholder replaced with the imported anchor.
    """
    from lambda_catalog.regression_layout import _A_BACK_TRANSFORM_METHOD

    block = _extract_once(
        "lambda_catalog/write_sheet_regression.py",
        r'("=Unit_Space_R_Squared\(.*?f"\{_A_BACK_TRANSFORM_METHOD\}\)")',
        flags=re.S,
    )
    body = _extract_concatenated_literals(block)
    body = body.replace("{_A_BACK_TRANSFORM_METHOD}", _A_BACK_TRANSFORM_METHOD)
    assert body.startswith("=Unit_Space_R_Squared(") and "$AH$5)" in body, (
        f"unexpected Unit_Space_R_Squared body: {body[:60]}"
    )
    return body


# The narrative annotation for each sample lives here; the formula itself is
# always imported or extracted, never hand-copied.
_FORMULA_SAMPLES: list[dict[str, str]] = [
    {
        "title": "The self-sizing spec band",
        "source": "lambda_catalog/write_spec_block.py — `_spec_band` (the `Spec_Role` name)",
        "formula": lambda: _spec_band("Regression", _sl._C_ROLE),
        "plain": (
            "Take the first `COLUMNS(Source_Data)` rows of the fixed 16000-row "
            "band under column B. Because it is `TAKE` (not the volatile "
            "`OFFSET`), the name costs nothing until the table is retargeted — "
            "and retargeting `Source_Table` resizes every spec column at once. "
            "This one formula is the whole \"one edit\" promise."
        ),
    },
    {
        "title": "A status cell is a call, not text",
        "source": "lambda_catalog/write_spec_block.py — the B2 cell; body from lambda_functions.json",
        "formula": lambda: (
            "=Role_Status()\n\n"
            + json.loads((ROOT / "lambda_functions.json").read_text(encoding="utf-8"))
            ["functions"][[f["name"] for f in json.loads(
                (ROOT / "lambda_functions.json").read_text(encoding="utf-8")
            )["functions"]].index("Role_Status")]["formula_display"]
        ),
        "plain": (
            "Click B2 and the formula bar shows `=Role_Status()` — one name. "
            "The check itself lives in the Name Manager: exactly one Response "
            "row, at most one Fixed Effects row, and a plain-English error "
            "message in the cell when a rule is broken. Red fill comes from "
            "a conditional format keyed on the cell being non-blank."
        ),
    },
    {
        "title": "One formula, many rows — the computed spec columns",
        "source": "lambda_catalog/write_spec_block.py — `_PERIOD_IN_USE_SPILL_FORMULA`",
        "formula": lambda: _PERIOD_IN_USE_SPILL_FORMULA,
        "plain": (
            "Column J (Period In Use) is a single `MAP(SEQUENCE(nc), ...)` "
            "spill: for each spec row, show the typed override (column I) if "
            "one was entered, otherwise the computed candidate — the most "
            "common gap between consecutive periods. One formula covers "
            "every row and resizes with the table; `nc` is just the table's "
            "column count."
        ),
    },
    {
        "title": "Every statistic is a named function over the fitted model",
        "source": "lambda_catalog/write_sheet_regression.py — the Regression Statistics block",
        "formula": lambda: _extract_once(
            "lambda_catalog/write_sheet_regression.py",
            r'("=Multiple_R\([^"]*\)")',
        ).strip('"'),
        "plain": (
            "Multiple R is not a bespoke formula — it is the catalog's "
            "`Multiple_R` LAMBDA called with the materialized model: the "
            "design matrix, the response, the row mask, and the fit context. "
            "Every cell in the Regression Statistics block reads the same "
            "way, which is why a spec edit updates all of them at once."
        ),
    },
    {
        "title": "The Duan/Naive toggle drives the back-transformation",
        "source": "lambda_catalog/write_sheet_regression.py — the Unit-Space Fit block",
        "formula": _unit_space_r_squared_formula,
        "plain": (
            "When the response is Log-transformed, the fit runs in log space "
            "and `EXP(ŷ)` is the *median* prediction, not the mean. The last "
            "argument here is the AH5 toggle: Duan multiplies by the smearing "
            "factor (the mean of EXP(residuals)) to recover the conditional "
            "mean; Naive is plain EXP. One toggle re-points R², RMSE, the "
            "prediction bounds and the residual columns together."
        ),
    },
    {
        "title": "A distribution fit is one formula — Weibull's profile NLL",
        "source": "lambda_catalog/write_sheet_univariate.py — `_stage_nll` (Weibull)",
        "formula": _weibull_nll_formula,
        "plain": (
            "The Univariate sheet's Weibull fit evaluates N points, not N², "
            "because the scale parameter is profiled out in closed form "
            "(λ̂ = (mean of xᵏ)^(1/k)). One `BYROW` walks the grid axis; "
            "`IFERROR` sits INSIDE the `LAMBDA` so one non-evaluable trial "
            "costs only its own row; `INDEX(r,1,1)` scalarizes the 1×1 row "
            "`BYROW` hands the callback. The anchor `$BP$33#` is the Stage-1 "
            "grid spill — the `#` operator reads the whole spilled axis "
            "whatever its height."
        ),
    },
    {
        "title": "The Beta fit's 2-D grid — N² evaluations in one spill pair",
        "source": "lambda_catalog/write_sheet_univariate.py — `_stage_nll` (Beta)",
        "formula": _beta_nll_formula,
        "plain": (
            "Beta has no closed-form partner, so each stage is a Cartesian "
            "product: an N²×2 `Full_Factorial` grid spill (Alpha | Beta) and "
            "this `BYROW` NLL column reading it via `#`. The sample is "
            "rescaled once into `z` on [pad, 1] — `NLL_Beta` needs a bounded "
            "support — and `COUNT(d)*LN(scale_)` is the Jacobian that puts "
            "the rescaled NLL back on the original scale."
        ),
    },
    {
        "title": "Finding the optimum: `Grid_Argument_Minimum`",
        "source": "lambda_functions.json — the Grid_Argument_Minimum entry",
        "formula": lambda: next(
            f["formula_display"]
            for f in json.loads(
                (ROOT / "lambda_functions.json").read_text(encoding="utf-8")
            )["functions"]
            if f["name"] == "Grid_Argument_Minimum"
        ),
        "plain": (
            "The whole search-recovery trick in one LAMBDA: find the minimum "
            "of a grid, then recover WHERE it sits. `TOCOL` flattens the grid "
            "row-major, `XMATCH` finds the first minimum's flat position, "
            "and `QUOTIENT`/`MOD` convert it back to (row, column). The "
            "boundary guard reads the column: if the best shape is the first "
            "or last grid point, the optimum is on the edge and the Min/Max "
            "bounds should be widened — the sheet turns that cell red."
        ),
    },
]


def write_formula_review() -> None:
    lines = [
        "# Code review — the formulas, annotated",
        "",
        "Every formula below is pulled from the repo at generation time",
        "(imported constants, the JSON catalog, or a single-match source",
        "extraction), so the docs cannot drift from the workbook. Each is",
        "annotated in plain English.",
        "",
    ]
    for sample in _FORMULA_SAMPLES:
        lines += [f"## {sample['title']}", "",
                  f"*Source: {sample['source']}*", "", "```excel",
                  sample["formula"](), "```", "", sample["plain"], ""]
    _write("formula-review.md", "\n".join(lines))


# ── Workbook tour ──────────────────────────────────────────────────────────────


def write_workbook_tour() -> None:
    """The tab order, from build_production.py's ordered_front.

    The ORDER lives in a function-local list, so it is regex-extracted from
    the source (fail-loud); the NAMES resolve by importing the module —
    some are plain strings, some are constants imported from the writer
    modules, and getattr handles both without a second regex.
    """
    import sys

    scripts_dir = str(ROOT / "scripts")
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import build_production as _bp  # noqa: E402  (main-guarded, import-safe)

    src = (ROOT / "scripts" / "build_production.py").read_text(encoding="utf-8")
    m = re.search(r"ordered_front = \[(.*?)\]", src, flags=re.S)
    assert m, "ordered_front list not found in build_production.py"
    consts = re.findall(r"_SHEET_NAME_([A-Z_]+)", m.group(1))
    assert len(consts) == 10, f"expected 10 sheet names, found {len(consts)}"
    names = []
    for cname in consts:
        attr = f"_SHEET_NAME_{cname}"
        assert hasattr(_bp, attr), f"{attr} not found on build_production"
        name = getattr(_bp, attr)
        assert isinstance(name, str), f"{attr} did not resolve to a sheet name"
        names.append(name)
    lines = [
        "# The workbook, tab by tab",
        "",
        "The shipped `dist/Lambda_Library.xlsx` presents ten tabs in this",
        "order (extracted from `scripts/build_production.py` at generation",
        "time):",
        "",
    ]
    for i, name in enumerate(names, start=1):
        lines.append(f"{i}. {name}")
    lines += [
        "",
        "- **Regression** — the working sheet: MODEL SPECIFICATION (A–O),",
        "  Regression Outputs, Prediction Outputs, Residual Output, and the",
        "  seven diagnostic charts.",
        "- **Regression Instructions / Modeling Concepts / Diagnostic Guide** —",
        "  the built-in manual (each has a generated page in this site).",
        "- **Univariate** — descriptive statistics, histograms, distribution",
        "  fitting, and Q-Q plots for one column of data.",
        "- **LAMBDA_functions** — the catalog itself, one row per function.",
        "- **Version History** — what shipped when.",
        "- **Production Lots / Life Expectancy Data / Mileage Data** — three",
        "  Excel Tables you can practice retargeting `Source_Table` against.",
        "",
    ]
    _write("workbook-tour.md", "\n".join(lines))


def main() -> None:
    write_workbook_tour()
    write_regression_instructions()
    write_modeling_concepts()
    write_diagnostic_guide()
    write_spec_block()
    write_lambda_reference()
    write_formula_review()
    print(f"done: {len(list(GENERATED.glob('*.md')))} pages in docs/generated/")


if __name__ == "__main__":
    main()
from pathlib import Path
import json
import re

# Reuse the already-reviewed structural transformation staged in the first
# temporary workflow, then align contract tests with the intentionally changed
# API. This helper is deleted by the successful commit and never lands in the PR.
workflow = Path('.github/workflows/refactor-log-drop-sample-mask.yml').read_text(encoding='utf-8')
marker = "          python - <<'PY'\n"
start = workflow.index(marker) + len(marker)
end = workflow.index("\n          PY\n", start)
lines = workflow[start:end].splitlines()
code = "\n".join(line[10:] if line.startswith('          ') else line for line in lines)
exec(compile(code, '<staged-refactor>', 'exec'), {'__name__': '__main__'})

catalog_path = Path('lambda_functions.json')
document = json.loads(catalog_path.read_text(encoding='utf-8'))
by_name = {item['name']: item for item in document['functions']}
for name in ('Sample_Include_Calc', 'Log_Drop_Sample_Include_Calc', 'Log_Domain_Status'):
    formula = by_name[name]['formula_display'].rstrip()
    by_name[name]['formula_display'] = formula.rsplit('\n', 1)[0] + '\n)'
catalog_path.write_text(json.dumps(document, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def replace_test(text: str, name: str, replacement: str) -> str:
    pattern = rf'def {re.escape(name)}\(\) -> None:\n.*?(?=\ndef |\Z)'
    updated, count = re.subn(pattern, replacement.rstrip() + '\n\n', text, count=1, flags=re.S)
    if count != 1:
        raise RuntimeError(f'could not replace {name}: {count}')
    return updated


path = Path('tests/test_spec_block_writer.py')
text = path.read_text(encoding='utf-8')
text = text.replace(
    '    "Sample_Include_Calc",\n    "Sample_Include",\n    "Response_Column",',
    '    "Sample_Include_Calc",\n    "Sample_Include",\n    "Log_Drop_Sample_Include_Calc",\n    "Response_Column",',
    1,
)
text = replace_test(text, 'test_sample_include_is_the_reduce_product_mask', '''def test_sample_include_is_the_reduce_product_mask() -> None:
    sheet = _named_sheet()
    base = _refers_to(sheet, "Sample_Include_Calc")
    final = _refers_to(sheet, "Log_Drop_Sample_Include_Calc")

    # Ordinary eligibility owns Filters and numeric completeness, but no
    # transform semantics at all.
    assert base.startswith("=LAMBDA(LET(")
    assert 'INDEX(Spec_Roles,Column_Number)="Filter"' in base
    assert 'Accumulated_Mask*N(ISNUMBER(Source_Column))' in base
    assert "Spec_Transform" not in base
    assert "Log (drop ≤ 0)" not in base
    assert '="Log"' not in base
    assert "REDUCE(" in base
    assert "BYROW(" not in base
    assert base.endswith("Sample_Mask=1))")

    # The specialized final-mask leaf starts from ordinary eligibility and
    # adds positivity only for the explicit row-dropping transform.
    assert "Base_Sample_Mask,Sample_Include_Calc()" in final
    assert 'INDEX(Spec_Transforms,Column_Number)="Log (drop ≤ 0)"' in final
    assert "Accumulated_Mask*--IFERROR((Source_Column+0)>0,FALSE)" in final
    assert 'INDEX(Spec_Transforms,Column_Number)="Log"' not in final
    assert "Ln_Positive" not in final
    assert final.endswith("Log_Drop_Sample_Mask=1))")

    # The public Sample_Include name is intentionally the ordinary mask.
    assert _refers_to(sheet, "Sample_Include") == "=LAMBDA(Sample_Include_Calc())"''')

text = replace_test(text, 'test_spec_transform_is_read_only_by_the_transform_aware_constructors', '''def test_spec_transform_is_read_only_by_the_transform_aware_constructors() -> None:
    sheet = _named_sheet()
    readers = sorted(
        item.Name.split("!", 1)[-1]
        for item in sheet.api.Names.items
        if "Spec_Transform" in item.RefersTo
        and item.Name.split("!", 1)[-1] != "Spec_Transform"
    )
    assert readers == [
        "Constructed_Column_Names",
        "Constructed_Column_Transforms",
        "Log_Domain_Status",
        "Log_Drop_Sample_Include_Calc",
        "Model_Formula",
        "Predictor_Columns",
        "Response_Column",
    ]
    assert "Spec_Transform" not in _refers_to(sheet, "Row_Labels")
    assert "Spec_Transform" not in _refers_to(sheet, "Sample_Include")
    assert "Spec_Transform" not in _refers_to(sheet, "Sample_Include_Calc")

    mask = _refers_to(sheet, "Log_Drop_Sample_Include_Calc")
    assert 'INDEX(Spec_Transforms,Column_Number)="Log (drop ≤ 0)"' in mask
    assert "Ln_Positive" not in mask
    assert 'INDEX(Spec_Transforms,Column_Number)="Log"' not in mask''')

text = replace_test(text, 'test_log_domain_status_reports_the_poisoned_column_then_the_dropped_count', '''def test_log_domain_status_reports_the_poisoned_column_then_the_dropped_count() -> None:
    """G2 — strict Log failure outranks intentional Log-drop sample shrinkage."""
    sheet = _feedback_sheet()

    assert (
        cast(str, sheet.cell(_FEEDBACK_STATUS_ROW, _C_TRANSFORM).api.Formula2)
        == "=Log_Domain_Status()"
    )
    formula = _catalog_body("Log_Domain_Status")
    compact = formula.replace("\n", "").replace(" ", "")

    # RED: only strict Log is tested, and only on rows that actually survive
    # into the fit.
    assert 'INDEX(Spec_Transforms,Column_Number)="Log"' in formula
    assert "--Fitted_Sample_Include" in formula
    assert "Use Log (drop ≤ 0)." in formula
    assert (
        'Eligible_Columns,((Spec_Roles="Response(y)")+'
        '((Spec_Roles="Predictor(x)")*(Spec_Includes=TRUE)'
        '*(Spec_Types="Continuous")))>0'
    ) in compact

    # AMBER: ordinary eligibility minus the final fit mask is exactly the
    # distinct rows removed by the explicit Log-drop layer.
    assert "Base_Sample_Include,Sample_Include()" in formula
    assert "Fitted_Sample_Include,Fit_Sample_Include()" in formula
    assert (
        "Log_Drop_Excluded_Row_Count,SUMPRODUCT(--Base_Sample_Include)-"
        "SUMPRODUCT(--Fitted_Sample_Include)"
    ) in compact
    assert '" excluded: Log of ≤ 0"' in formula
    assert "N(Fit_Sample_Include())" not in formula

    conditions = sheet.range("$G$2").api.FormatConditions.items
    assert [c.Formula1 for c in conditions] == [
        '=ISNUMBER(SEARCH("ERROR",$G$2))',
        '=$G$2<>""',
    ]
    red, amber = conditions
    assert red.Interior.Color == excel_color(CF_LIGHT_RED_FILL)
    assert red.Font.Color == excel_color(CF_DARK_RED_TEXT)
    assert red.StopIfTrue is True
    assert amber.Interior.Color == excel_color(CF_YELLOW_FILL)
    assert amber.Font.Color == excel_color(CF_DARK_YELLOW_TEXT)
    assert amber.StopIfTrue is False''')

text = text.replace(
    '    assert \'" rows excluded: Log of ≤ 0"\' in log',
    '    assert \'" excluded: Log of ≤ 0"\' in log',
    1,
)

text = replace_test(text, 'test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform', '''def test_both_log_tokens_reach_every_catalog_body_that_reads_spec_transform() -> None:
    import json
    from pathlib import Path

    from lambda_catalog.write_spec_block import _TRANSFORM_LOG, _TRANSFORM_LOG_DROP

    document = json.loads(
        (Path(__file__).resolve().parents[1] / "lambda_functions.json").read_text(
            encoding="utf-8"
        )
    )
    bodies = {
        entry["name"]: entry["formula_display"]
        for entry in document["functions"]
        if "Spec_Transform" in entry.get("formula_display", "")
    }

    assert set(bodies) == {
        "Constructed_Column_Names",
        "Constructed_Column_Transforms",
        "Log_Domain_Status",
        "Log_Drop_Sample_Include_Calc",
        "Model_Formula",
        "Predictor_Columns",
        "Response_Column",
    }

    for name, body in bodies.items():
        compact = body.replace(" ", "").replace("\n", "")
        if name == "Log_Drop_Sample_Include_Calc":
            assert _TRANSFORM_LOG_DROP in body
            assert f'="{_TRANSFORM_LOG}"' not in compact, name
            continue
        if name == "Log_Domain_Status":
            assert f'="{_TRANSFORM_LOG}"' in compact
            assert f"Use {_TRANSFORM_LOG_DROP}." in body
            assert f'="{_TRANSFORM_LOG_DROP.replace(" ", "")}"' not in compact, name
            continue
        assert f'="{_TRANSFORM_LOG}"' in compact, name
        assert compact.count(f'="{_TRANSFORM_LOG}"') == compact.count(
            f'="{_TRANSFORM_LOG_DROP.replace(" ", "")}"'
        ), name''')
path.write_text(text, encoding='utf-8')

path = Path('tests/test_within_estimator.py')
text = path.read_text(encoding='utf-8')
text = replace_test(text, 'test_design_columns_and_sample_include_are_readers_over_their_spills', '''def test_design_columns_and_sample_include_are_readers_over_their_spills() -> None:
    # Sample_Include is now deliberately the ordinary pre-drop eligibility
    # mask. The final fitted mask remains materialized through
    # Fit_Sample_Include; Design_Columns remains a reader over its own spill.
    si = _formula("Sample_Include")
    log_drop = _formula("Log_Drop_Sample_Include_Calc")
    dc = _formula("Design_Columns")
    assert si == "LAMBDA(Sample_Include_Calc())"
    assert "Sample_Include_Calc()" in log_drop
    assert 'Log (drop ≤ 0)' in log_drop
    assert dc == "LAMBDA(Fit_Design_Columns())"''')
path.write_text(text, encoding='utf-8')

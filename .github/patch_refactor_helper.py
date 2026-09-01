from pathlib import Path

path = Path('.github/refactor_log_drop_retry.py')
text = path.read_text(encoding='utf-8')
text = text.replace(
    r'compact = formula.replace("\n", "").replace(" ", "")',
    'compact = "".join(formula.splitlines()).replace(" ", "")',
)
text = text.replace(
    r'compact = body.replace(" ", "").replace("\n", "")',
    'compact = "".join(body.splitlines()).replace(" ", "")',
)
path.write_text(text, encoding='utf-8')

"""One-shot audit: count catalog functions in lambda_functions.json."""
import json
from pathlib import Path

CATALOG = Path(__file__).parent.parent / "lambda_functions.json"
doc = json.loads(CATALOG.read_text(encoding="utf-8"))
# catalog_schema: a CatalogDocument has .functions
functions = doc.get("functions", doc) if isinstance(doc, dict) else doc
print(f"Total functions: {len(functions)}")
print(f"Workbook-scoped: {sum(1 for f in functions if f.get('scope', 'workbook') == 'workbook')}")
print(f"Sheet-scoped:    {sum(1 for f in functions if f.get('scope', 'workbook') != 'workbook')}")
print("---")
for f in functions:
    if f.get("scope", "workbook") != "workbook":
        print(f"  {f['name']}  scope={f.get('scope')}")
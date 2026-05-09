"""
build_workbook.py — Create/update Lambda_Library.xlsx with workbook-scoped LAMBDA names.

Reads lambda_functions.json, validates each entry, then syncs every function into
the Excel Name Manager as a workbook-level name.  Existing names are overwritten.
Requires desktop Excel through xlwings (Windows or macOS).

Usage:
    python build_workbook.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

WORKBOOK = "Lambda_Library.xlsx"
JSON_SRC = Path(__file__).parent / "lambda_functions.json"
STARTER_SHEET = "MLR"


# ── JSON loading & validation ─────────────────────────────────────────────────

def _load_functions() -> tuple[list[dict], int]:
    try:
        raw = json.loads(JSON_SRC.read_text(encoding="utf-8"))
    except FileNotFoundError:
        sys.exit(f"ERROR: {JSON_SRC} not found")
    except json.JSONDecodeError as exc:
        sys.exit(f"ERROR: {JSON_SRC} is not valid JSON — {exc}")

    if not isinstance(raw, dict) or "functions" not in raw:
        sys.exit("ERROR: lambda_functions.json must contain a top-level 'functions' array")

    entries = raw["functions"]
    if not isinstance(entries, list):
        sys.exit("ERROR: 'functions' must be a list")

    valid: list[dict] = []
    seen: set[str] = set()
    invalid = 0

    for i, entry in enumerate(entries):
        name = (entry.get("name") or "").strip()
        formula = (entry.get("formula_compact") or "").strip()

        if not name or not formula:
            missing = "name" if not name else "formula_compact"
            print(f"  SKIP entry[{i}]: missing '{missing}'")
            invalid += 1
            continue

        if name in seen:
            sys.exit(f"ERROR: duplicate name '{name}' in lambda_functions.json")

        seen.add(name)
        valid.append({"name": name, "formula": formula})

    return valid, invalid


# ── Workbook helpers ──────────────────────────────────────────────────────────

def _ensure_starter_sheet(wb: xw.Book, is_new: bool) -> None:
    sheet_names = [s.name for s in wb.sheets]
    if STARTER_SHEET not in sheet_names:
        if is_new and len(wb.sheets) == 1:
            wb.sheets[0].name = STARTER_SHEET          # rename default Sheet1
        else:
            wb.sheets.add(STARTER_SHEET)

    ws = wb.sheets[STARTER_SHEET]
    if ws["A1"].value is None:
        ws["A1"].value = "MLR Lambda Library"
        ws["A2"].value = (
            "Multilinear regression helper functions — "
            "use Name Manager (Ctrl+F3) to browse definitions."
        )
        ws["A4"].value = "Function"
        ws["B4"].value = "Arguments"
        ws["C4"].value = "Returns"
        ws["D4"].value = "Formula (compact)"
        ws["E4"].value = "Description"


def _sync_names(wb: xw.Book, functions: list[dict]) -> tuple[int, int, int]:
    # Index existing workbook-scoped names (sheet-scoped names contain '!')
    existing: dict[str, object] = {
        n.Name: n
        for n in wb.api.Names
        if "!" not in n.Name
    }

    created = updated = failed = 0
    for fn in functions:
        fname, formula = fn["name"], fn["formula"]
        try:
            if fname in existing:
                existing[fname].RefersTo = formula
                updated += 1
            else:
                wb.api.Names.Add(fname, formula)   # positional: Name, RefersTo
                created += 1
        except Exception as exc:
            # Excel rejects names that look like cell references (e.g. R2 = R1C1 row 2)
            print(f"  FAIL {fname}: {exc}")
            failed += 1

    return created, updated, failed


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    wb_path = Path(__file__).parent / WORKBOOK
    functions, invalid = _load_functions()

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    wb: xw.Book | None = None
    created = updated = failed = 0
    try:
        is_new = not wb_path.exists()
        if is_new:
            wb = app.books.add()
            wb.save(str(wb_path))
        else:
            wb = app.books.open(str(wb_path))

        _ensure_starter_sheet(wb, is_new)
        created, updated, failed = _sync_names(wb, functions)
        wb.save()
    finally:
        if wb is not None:
            wb.close()
        app.quit()

    print(f"Workbook : {wb_path.resolve()}")
    print(f"Created  : {created}")
    print(f"Updated  : {updated}")
    if failed:
        print(f"Failed   : {failed}  (Excel rejected — see FAIL lines above)")
    print(f"Invalid  : {invalid}  (malformed JSON entries)")


if __name__ == "__main__":
    main()

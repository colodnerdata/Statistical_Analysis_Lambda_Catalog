"""
inspect_xlsx.py
Inspect the ZIP contents and key XML files inside an Excel workbook.

Usage:
    python inspect_xlsx.py [FILE]

    FILE  Path to the .xlsx file to inspect (default: Lambda_Library.xlsx)
"""
from __future__ import annotations
import argparse
import zipfile


def inspect(file_path: str) -> None:
    with zipfile.ZipFile(file_path, "r") as z:
        print("--- File List ---")
        for name in z.namelist():
            print(f"  {name}")

        print("\n--- [Content_Types].xml ---")
        with z.open("[Content_Types].xml") as f:
            print(f.read().decode("utf-8"))

        print("\n--- xl/_rels/workbook.xml.rels ---")
        try:
            with z.open("xl/_rels/workbook.xml.rels") as f:
                print(f.read().decode("utf-8"))
        except KeyError:
            print("  xl/_rels/workbook.xml.rels not found")


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect the contents of an Excel workbook.")
    parser.add_argument("file", nargs="?", default="Lambda_Library.xlsx",
                        help="Path to the .xlsx file (default: Lambda_Library.xlsx)")
    args = parser.parse_args()

    try:
        inspect(args.file)
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()

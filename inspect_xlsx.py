import zipfile
import lxml.etree as ET

import sys; file_path = sys.argv[1] if len(sys.argv) > 1 else "Lambda_Library.xlsx"

try:
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

except Exception as e:
    print(f"Error: {e}")

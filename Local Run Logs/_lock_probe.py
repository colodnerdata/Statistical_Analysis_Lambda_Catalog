"""Probe: is the test-models workbook currently locked, and by whom?"""
import os
import sys
from pathlib import Path

WORKBOOK = Path(r"C:\Users\Stephen_Colodner\windows code\Statistical_Analysis_Lambda_Catalog\Lambda_Library_TestModels.xlsx")
if not WORKBOOK.exists():
    print(f"NOT FOUND: {WORKBOOK}")
    sys.exit(1)

# Probe 1: try to open the file with share mode = 0 (exclusive). If sharing
# violation, the file is locked by something else.
import msvcrt
try:
    fd = os.open(str(WORKBOOK), os.O_RDWR | os.O_BINARY)
    os.close(fd)
    print("OPEN exclusive OK -- file is NOT locked")
except PermissionError as exc:
    print(f"LOCKED: {exc}")
    print("  Windows reports another process has the file open with a share mode that excludes this exclusive open.")

# Probe 2: read-only with sharing=0 also fails if file is locked.
try:
    fd = os.open(str(WORKBOOK), os.O_RDONLY | os.O_BINARY)
    os.close(fd)
    print("OPEN read-only OK -- no lock at all")
except PermissionError as exc:
    print(f"LOCKED (read-only also failed): {exc}")

# Probe 3: enumerate all processes; filter to those likely to hold the lock.
print()
print("Suspicious processes (excel/python related):")
import win32com.client  # pywin32 ships with the project
wmi = win32com.client.GetObject("winmgmts:")
for p in wmi.InstancesOf("Win32_Process"):
    name = (p.Name or "").lower()
    if name in ("excel.exe", "python.exe"):
        try:
            cmd = p.CommandLine or ""
            short = cmd[:200].replace("\n", " ")
            print(f"  PID={p.ProcessId:>6}  {p.Name:<15}  {short}")
        except Exception:
            pass
"""Probe: is the test-models workbook currently locked, and by whom?"""
import os
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
WORKBOOK = Path(sys.argv[1]) if len(sys.argv) > 1 else (ROOT_DIR / "Lambda_Library_TestModels.xlsx")
if not WORKBOOK.exists():
    print(f"NOT FOUND: {WORKBOOK}")
    sys.exit(1)

# Probe 1: try to open the file with dwShareMode = 0 (exclusive). If this fails
# with a sharing/lock violation, another process has it open.
import pywintypes
import win32con
import win32file

try:
    h = win32file.CreateFile(
        str(WORKBOOK),
        win32con.GENERIC_READ,
        0,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None,
    )
    win32file.CloseHandle(h)
    print("OPEN exclusive OK -- file is NOT locked")
except pywintypes.error as exc:
    if getattr(exc, "winerror", None) in (32, 33):  # sharing/lock violation
        print(f"LOCKED: {exc}")
    else:
        raise
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
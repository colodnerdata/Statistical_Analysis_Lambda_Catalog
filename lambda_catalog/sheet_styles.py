"""Shared cell-formatting constants for all workbook sheet writers.

All sheet writers import their color constants from here so the palette
stays consistent across every sheet in Lambda_Library.xlsx.
"""
from __future__ import annotations

# ── Palette ───────────────────────────────────────────────────────────────────
HEADER_COLOR = (202, 237, 251)   # section / zone headings — light blue
SUBHDR_COLOR = (220, 230, 241)   # column sub-headers — slightly darker blue
INPUT_COLOR  = (251, 226, 213)   # user-editable input cells — light orange

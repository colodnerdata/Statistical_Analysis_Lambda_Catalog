"""Shared cell-formatting constants for all workbook sheet writers.

All sheet writers import their color constants from here so the palette
stays consistent across every sheet in Lambda_Library.xlsx.  Includes
both cell-fill colors (HEADER, SUBHDR, INPUT) and conditional-formatting
colors (CF_LIGHT_RED_FILL, CF_DARK_RED_TEXT, CF_YELLOW_FILL, CF_DARK_YELLOW_TEXT).
"""
from __future__ import annotations

# ── Palette ───────────────────────────────────────────────────────────────────
HEADER_COLOR = (202, 237, 251)   # section / zone headings — light blue
SUBHDR_COLOR = (220, 230, 241)   # column sub-headers — slightly darker blue
INPUT_COLOR  = (251, 226, 213)   # user-editable input cells — light orange
CF_LIGHT_RED_FILL = (255, 199, 206)    # failed significance / diagnostic flag
CF_DARK_RED_TEXT = (156, 0, 6)
CF_YELLOW_FILL = (255, 235, 156)       # borderline diagnostic flag
CF_DARK_YELLOW_TEXT = (156, 101, 0)

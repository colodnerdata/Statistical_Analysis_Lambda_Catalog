"""
format_workbook.py
Visual formatting for MLR_Lambda_Library.xlsx.
Not wired up yet — call format_workbook(path) after build_workbook() when ready.

Dependencies: xlwings
"""
from __future__ import annotations
import xlwings as xw


# ── Color helpers ─────────────────────────────────────────────────────────────

def _hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


BLUE  = _hex_to_rgb("2F5496")   # (47, 84, 150)  — header fill / title font
WHITE = _hex_to_rgb("FFFFFF")   # (255, 255, 255) — header font
GRAY  = _hex_to_rgb("595959")   # (89, 89, 89)    — description font


# ── COM alignment constants ───────────────────────────────────────────────────
# Integer values of Excel's xlHAlign / xlVAlign enum members.
# Avoids importing win32com.client constants.

XL_HALIGN_CENTER = -4108
XL_HALIGN_LEFT   = -4131
XL_VALIGN_CENTER = -4108
XL_VALIGN_TOP    = -4160


# ── Style helper ──────────────────────────────────────────────────────────────

def _style(
    rng,
    *,
    font_name:   str | None   = None,
    font_size:   float | None = None,
    font_bold:   bool | None  = None,
    font_italic: bool | None  = None,
    font_color:  tuple | None = None,  # (R, G, B)
    bg_color:    tuple | None = None,  # (R, G, B)
    h_align:     int | None   = None,  # XL_HALIGN_* constant
    v_align:     int | None   = None,  # XL_VALIGN_* constant
    wrap_text:   bool | None  = None,
) -> None:
    if font_name   is not None: rng.font.name   = font_name
    if font_size   is not None: rng.font.size   = font_size
    if font_bold   is not None: rng.font.bold   = font_bold
    if font_italic is not None: rng.font.italic = font_italic
    if font_color  is not None: rng.font.color  = font_color
    if bg_color    is not None: rng.color        = bg_color
    if h_align     is not None: rng.api.HorizontalAlignment = h_align
    if v_align     is not None: rng.api.VerticalAlignment   = v_align
    if wrap_text   is not None: rng.api.WrapText            = wrap_text


# ── Main entry point ──────────────────────────────────────────────────────────

def format_workbook(path: str) -> None:
    """
    Open an existing MLR_Lambda_Library.xlsx and apply all visual formatting.
    Expects the workbook to have been created by build_workbook() first.
    """
    app = xw.App(visible=True)
    wb  = app.books.open(path)
    ws  = wb.sheets["MLR"]

    last_data_row = ws.range("A4").current_region.last_cell.row

    # Row 1 — title
    _style(ws["A1"],
           font_name="Arial", font_bold=True, font_size=14,
           font_color=BLUE,
           h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
    ws.range("1:1").row_height = 28

    # Row 2 — description
    _style(ws["A2"],
           font_name="Arial", font_size=9, font_italic=True,
           font_color=GRAY,
           h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER, wrap_text=True)
    ws.range("2:2").row_height = 30

    # Row 4 — table headers
    for col_idx in range(1, 6):
        _style(ws.range(4, col_idx),
               font_name="Arial", font_bold=True, font_size=11,
               font_color=WHITE, bg_color=BLUE,
               h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)

    # Data rows
    for r in range(5, last_data_row + 1):
        _style(ws.range(r, 1),
               font_name="Arial", font_bold=True, font_size=10,
               v_align=XL_VALIGN_TOP)
        _style(ws.range(r, 2),
               font_name="Courier New", font_size=9,
               v_align=XL_VALIGN_TOP, wrap_text=True)
        _style(ws.range(r, 3),
               font_name="Courier New", font_size=9,
               v_align=XL_VALIGN_TOP, wrap_text=True)
        _style(ws.range(r, 4),
               font_name="Arial", font_size=10,
               v_align=XL_VALIGN_TOP, wrap_text=True)
        _style(ws.range(r, 5),
               font_name="Arial", font_size=10,
               v_align=XL_VALIGN_TOP, wrap_text=True)

    # Table visual style
    lo = ws.api.ListObjects("Lambda_Definitions")
    lo.TableStyle = "TableStyleMedium2"
    lo.ShowTableStyleRowStripes    = True
    lo.ShowFirstColumn             = False
    lo.ShowLastColumn              = False
    lo.ShowTableStyleColumnStripes = False

    wb.save()
    print(f"Formatted: {path}")

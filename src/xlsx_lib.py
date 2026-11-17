"""Shared openpyxl styling. Financial colour convention: blue = input, black = formula, green = link to other sheet."""
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F2A5A")
SUB_FILL = PatternFill("solid", fgColor="E8ECF5")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
NOTE_FILL = PatternFill("solid", fgColor="F2F4F8")
THIN = Side(style="thin", color="C9CDD6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLUE = "0000FF"
GREEN = "008000"
LAKH = '#,##0.00;(#,##0.00);"-"'
LAKH1 = '#,##0.0;(#,##0.0);"-"'
INT = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'


def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    return Font(**kw)


def title(ws, text, sub=None, row=1):
    ws.cell(row, 1, text).font = font(size=14, bold=True, color="1F2A5A")
    if sub:
        ws.cell(row + 1, 1, sub).font = font(size=9, italic=True, color="52514E")


def header(ws, row, headers, col=1, height=30):
    for i, h in enumerate(headers):
        c = ws.cell(row, col + i, h)
        c.font = font(bold=True, color="FFFFFF", size=9.5)
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX
    ws.row_dimensions[row].height = height


def cell(ws, row, col, value, fmt=None, bold=False, color=None, fill=None, wrap=False, comment=None, italic=False,
         size=10, align=None):
    c = ws.cell(row, col, value)
    if color is None:
        color = BLUE if (value is not None and not (isinstance(value, str) and value.startswith("=")) and
                         isinstance(value, (int, float))) else "000000"
    c.font = font(bold=bold, color=color, italic=italic, size=size)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    c.border = BOX
    c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    if comment:
        c.comment = Comment(comment, "Group 4")
    return c


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def note(ws, row, text, col=1, span=8):
    c = ws.cell(row, col, text)
    c.font = font(size=9, italic=True, color="52514E")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    ws.row_dimensions[row].height = max(15, 13 * (len(text) // (span * 11) + 1))

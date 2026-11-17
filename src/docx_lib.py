"""Tiny house-style wrapper around python-docx used by every Word deliverable."""
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import facts

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "data" / "figures"

INDIGO = RGBColor(0x1F, 0x2A, 0x5A)
TEAL = RGBColor(0x0F, 0x7C, 0x78)
INK2 = RGBColor(0x52, 0x51, 0x4E)
HEAD_FILL = "1F2A5A"
ALT_FILL = "F2F4F8"
CALLOUT = {"info": "E8F1FB", "warn": "FFF4DE", "ok": "E6F6EF", "ai": "F0EEFB"}
FONT = "Calibri"


def _shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _cell_margins(table, top=60, bottom=60, left=90, right=90):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)


def _field(run, instr):
    for t, txt in (("begin", None), (None, instr), ("end", None)):
        if t:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), t)
            run._r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            run._r.append(it)


def add_runs(par, text, size=None, color=None, italic=False):
    """Supports **bold** inline markup; strips markdown backticks."""
    text = text.replace("`", "")
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        bold = part.startswith("**") and part.endswith("**")
        r = par.add_run(part[2:-2] if bold else part)
        r.bold = bold or None
        r.italic = italic or None
        if size:
            r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    return par


class Doc:
    def __init__(self, title, subtitle, owner, code, landscape=False):
        self.d = Document()
        self.code = code
        sec = self.d.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        if landscape:
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
        for side in ("left_margin", "right_margin"):
            setattr(sec, side, Cm(2.0))
        sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)
        self.width_cm = (29.7 if landscape else 21.0) - 4.0
        st = self.d.styles
        st["Normal"].font.name = FONT
        st["Normal"].font.size = Pt(10.5)
        st["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        st["Normal"].paragraph_format.space_after = Pt(5)
        st["Normal"].paragraph_format.line_spacing = 1.12
        for lvl, size, col in ((1, 16, INDIGO), (2, 12.5, TEAL), (3, 11, INDIGO)):
            h = st[f"Heading {lvl}"]
            h.font.name = FONT
            rf = h._element.rPr.rFonts
            for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rf.set(qn(a), FONT)
            for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if rf.get(qn(a)) is not None:
                    del rf.attrib[qn(a)]
            h.font.size = Pt(size)
            h.font.bold = True
            h.font.color.rgb = col
            h.paragraph_format.space_before = Pt(14 if lvl == 1 else 10)
            h.paragraph_format.space_after = Pt(5)
            h.paragraph_format.keep_with_next = True
        for name in ("List Bullet", "List Number"):
            st[name].font.name = FONT
            st[name].font.size = Pt(10.5)
            st[name].paragraph_format.space_after = Pt(2)
        # header / footer
        hp = sec.header.paragraphs[0]
        add_runs(hp, f"{facts.CLIENT}  |  AI Transformation Consultancy Challenge  |  {facts.GROUP}", size=8, color=INK2)
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp = sec.footer.paragraphs[0]
        add_runs(fp, f"{code}  |  Synthetic/assumed data unless cited  |  Page ", size=8, color=INK2)
        r = fp.add_run()
        r.font.size = Pt(8)
        _field(r, "PAGE")
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # title block
        p = self.d.add_paragraph()
        add_runs(p, title, size=22, color=INDIGO).runs[0].bold = True
        p.paragraph_format.space_after = Pt(2)
        p = self.d.add_paragraph()
        add_runs(p, subtitle, size=12, color=TEAL)
        meta = [("Client", f"{facts.CLIENT} (fictional; Tiruppur textile exporter, Rs 120 crore revenue)"),
                ("Team", f"{facts.GROUP} (4 members); shock card {facts.SHOCK_CARD}"),
                ("Owner", owner), ("Date / version", f"{facts.TODAY}  /  v1.0")]
        self.table(["", ""], meta, widths=[3.2, self.width_cm - 3.2], show_header=False, size=9, first_col_bold=True)

    # ------------------------------------------------------------ blocks
    def h1(self, t):
        self.d.add_heading(t, 1)

    def h2(self, t):
        self.d.add_heading(t, 2)

    def h3(self, t):
        self.d.add_heading(t, 3)

    def p(self, t, size=None, italic=False, color=None, align=None, after=None):
        par = self.d.add_paragraph()
        add_runs(par, t, size=size, italic=italic, color=color)
        if align == "center":
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if after is not None:
            par.paragraph_format.space_after = Pt(after)
        return par

    def bullets(self, items, style="List Bullet", size=None):
        for it in items:
            par = self.d.add_paragraph(style=style)
            add_runs(par, it, size=size)

    def numbered(self, items, size=None):
        """Manual numbering so every list restarts at 1."""
        for i, it in enumerate(items, 1):
            par = self.d.add_paragraph()
            par.paragraph_format.left_indent = Cm(0.75)
            par.paragraph_format.first_line_indent = Cm(-0.55)
            par.paragraph_format.space_after = Pt(2)
            add_runs(par, f"{i}.  " + it, size=size)

    def table(self, header, rows, widths=None, show_header=True, size=9, first_col_bold=False,
              col_fill=None, align_right=()):
        ncol = len(rows[0]) if rows else len(header)
        t = self.d.add_table(rows=0, cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.style = "Table Grid"
        _cell_margins(t)
        widths = widths or [self.width_cm / ncol] * ncol
        if show_header:
            cells = t.add_row().cells
            for i, h in enumerate(header):
                cells[i].text = ""
                add_runs(cells[i].paragraphs[0], str(h), size=size, color=RGBColor(0xFF, 0xFF, 0xFF))
                cells[i].paragraphs[0].runs[0].bold = True
                _shade(cells[i], HEAD_FILL)
            # repeat header row
            trPr = t.rows[0]._tr.get_or_add_trPr()
            th = OxmlElement("w:tblHeader")
            th.set(qn("w:val"), "true")
            trPr.append(th)
        for ri, row in enumerate(rows):
            new_row = t.add_row()
            trPr = new_row._tr.get_or_add_trPr()
            cs = OxmlElement("w:cantSplit")
            cs.set(qn("w:val"), "true")
            trPr.append(cs)
            cells = new_row.cells
            for i, v in enumerate(row):
                cells[i].text = ""
                lines = str(v).split("\n")
                for li, line in enumerate(lines):
                    par = cells[i].paragraphs[0] if li == 0 else cells[i].add_paragraph()
                    par.paragraph_format.space_after = Pt(0)
                    add_runs(par, line, size=size)
                    if i in align_right:
                        par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    if first_col_bold and i == 0:
                        for r in par.runs:
                            r.bold = True
                if ri % 2 == 1 and show_header:
                    _shade(cells[i], ALT_FILL)
                if col_fill and i in col_fill:
                    _shade(cells[i], col_fill[i])
        t.autofit = False
        for i, w in enumerate(widths):
            t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
        self.d.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    def callout(self, text, kind="info", title=None):
        t = self.d.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _cell_margins(t, 100, 100, 160, 160)
        c = t.rows[0].cells[0]
        c.width = Cm(self.width_cm)
        _shade(c, CALLOUT[kind])
        par = c.paragraphs[0]
        if title:
            r = par.add_run(title + "  ")
            r.bold = True
            r.font.color.rgb = INDIGO
        lines = text if isinstance(text, list) else [text]
        add_runs(par, lines[0], size=10)
        for line in lines[1:]:
            add_runs(c.add_paragraph(), line, size=10)
        self.d.add_paragraph().paragraph_format.space_after = Pt(2)

    def image(self, path, width_cm=None, caption=None):
        self.d.add_picture(str(path), width=Cm(width_cm or self.width_cm))
        self.d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            self.p(caption, size=8.5, italic=True, color=INK2, align="center")

    def prompt_log(self, entries):
        """entries: list of dict(title, tool, prompt, raw, edited, why) - the brief's 'prompt, raw output, edited version'."""
        for i, e in enumerate(entries, 1):
            self.h3(f"Log {i}: {e['title']}")
            self.table(["Step", "Content"], [
                ("Tool / date", e.get("tool", f"Claude (Anthropic), {facts.TODAY}")),
                ("Prompt", e["prompt"]),
                ("Raw AI output (verbatim extract)", e["raw"]),
                ("Human-edited version", e["edited"]),
                ("What we changed and why", e["why"]),
            ], widths=[3.6, self.width_cm - 3.6], size=8.5, first_col_bold=True)

    def page_break(self):
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def sources(self, keys, extra=()):
        self.h2("Sources and labelling")
        items = [facts.SOURCES[k] for k in keys] + list(extra)
        items.append("All figures about WeaveWorks Exports are ASSUMED (fictional client). All customer, buyer, "
                     "order and employee data are SYNTHETIC, generated by src/gen_data.py.")
        self.bullets(items, size=9)

    def save(self, path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.d.save(path)
        return path

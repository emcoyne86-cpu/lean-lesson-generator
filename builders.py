"""
builders.py — Document builders for all 4 outputs.
All functions return BytesIO objects ready for st.download_button.
"""

import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn as pqn
from lxml import etree
import copy


# ══════════════════════════════════════════════════════════════════════════════
# DOCX BUILDER — Lean Lesson + Scaffolded Lesson
# ══════════════════════════════════════════════════════════════════════════════

# Colors
NAVY   = RGBColor(0x1F, 0x38, 0x64)
CORAL  = RGBColor(0xFF, 0x6B, 0x6B)
TEAL   = RGBColor(0x4E, 0xCE, 0xC8)
PURPLE = RGBColor(0x9B, 0x59, 0xB6)
AMBER  = RGBColor(0xF3, 0x9C, 0x12)
GREEN  = RGBColor(0x27, 0xAE, 0x60)
LTGRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DKTEXT = RGBColor(0x1A, 0x1A, 0x2E)

# Population colors
ALL_BG   = RGBColor(0xD5, 0xF5, 0xE3)
ELL_BG   = RGBColor(0xD6, 0xEA, 0xF8)
SCD_BG   = RGBColor(0xE8, 0xDA, 0xEF)
ALL_TEXT = RGBColor(0x1E, 0x84, 0x49)
ELL_TEXT = RGBColor(0x19, 0x56, 0x99)
SCD_TEXT = RGBColor(0x6C, 0x34, 0x83)

TAG_COLORS = {
    "BACKGROUND":   CORAL,
    "OBJECTIVE":    TEAL,
    "KEY QUESTION": PURPLE,
    "VOCAB":        AMBER,
    "MODEL":        GREEN,
    "READ ALOUD":   RGBColor(0x85, 0x29, 0x9E),
    "PRACTICE":     RGBColor(0xE6, 0x7E, 0x22),
    "PROTOCOL":     RGBColor(0x16, 0xA0, 0x85),
    "CHECK":        RGBColor(0xC0, 0x39, 0x2B),
    "CLOSE":        RGBColor(0x21, 0x8C, 0x74),
    "HW":           RGBColor(0x7F, 0x8C, 0x8D),
}

COL_A = Inches(0.7)
COL_B = Inches(3.9)
COL_C = Inches(2.4)

POP_LABEL = {"ALL": "ALL STUDENTS", "ELL": "ELL", "SCD": "SCD"}
POP_BG    = {"ALL": ALL_BG,  "ELL": ELL_BG,  "SCD": SCD_BG}
POP_TEXT  = {"ALL": ALL_TEXT, "ELL": ELL_TEXT, "SCD": SCD_TEXT}


def _set_cell_bg(cell, rgb: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
    existing = tcPr.find(qn("w:shd"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(shd)


def _set_col_width(cell, width):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(int(width.twips)))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def _cell_text(cell, text, bold=False, italic=False, color=None, size=10, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def _make_header_table(doc, lesson_data):
    tbl = doc.add_table(rows=4, cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    meta = [
        ("Curriculum", f"{lesson_data.get('curriculum','')} | {lesson_data.get('unit','')} | {lesson_data.get('lesson','')}"),
        ("Grade / Subject", f"Grade {lesson_data.get('grade','')} | {lesson_data.get('subject','')} | {lesson_data.get('duration','')}"),
        ("Standards", " · ".join(lesson_data.get("standards", []))),
        ("Objective", lesson_data.get("objective", "")),
    ]
    for i, (label, val) in enumerate(meta):
        row = tbl.rows[i]
        row.cells[0].merge(row.cells[0])
        _set_cell_bg(row.cells[0], NAVY)
        _cell_text(row.cells[0], label, bold=True, color=WHITE, size=9)
        _set_col_width(row.cells[0], Inches(1.2))
        _cell_text(row.cells[1], val, size=9)
        _set_col_width(row.cells[1], Inches(5.8))

    doc.add_paragraph()


def _section_header(doc, name, duration):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(f"{name.upper()}  ·  {duration}")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = NAVY


def _step_row(doc, step, show_scaffold=False):
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Prevent Word from auto-adjusting column widths
    tblPr = tbl._tbl.tblPr
    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    row = tbl.rows[0]
    tag_cell = row.cells[0]
    content_cell = row.cells[1]
    time_cell = row.cells[2]

    tag = step.get("tag", "")
    tag_color = TAG_COLORS.get(tag.upper(), NAVY)
    _set_cell_bg(tag_cell, tag_color)
    _cell_text(tag_cell, f"[{tag}]", bold=True, color=WHITE, size=8,
               align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_col_width(tag_cell, COL_A)

    _cell_text(content_cell, step.get("content", ""), size=9)
    _set_col_width(content_cell, COL_B)

    timing = step.get("timing", "")
    _cell_text(time_cell, f"{timing} min" if timing else "", size=9,
               align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_col_width(time_cell, COL_C)

    # Scaffold rows
    if show_scaffold:
        for scaf in step.get("scaffolds", []):
            pop = scaf.get("population", "ALL")
            bg = POP_BG.get(pop, ALL_BG)
            txt = POP_TEXT.get(pop, ALL_TEXT)

            scaf_row = tbl.add_row()
            label_cell = scaf_row.cells[0]
            scaf_cell  = scaf_row.cells[1]
            note_cell  = scaf_row.cells[2]

            _set_cell_bg(label_cell, bg)
            _cell_text(label_cell, POP_LABEL.get(pop, pop), bold=True, color=txt,
                       size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_col_width(label_cell, COL_A)

            scaf_cell.text = ""
            p = scaf_cell.paragraphs[0]
            r1 = p.add_run(f"[{scaf.get('type','')}]  ")
            r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = txt
            r2 = p.add_run(scaf.get("content", ""))
            r2.font.size = Pt(9)
            _set_col_width(scaf_cell, COL_B)

            _cell_text(note_cell, scaf.get("teacher_note", ""), size=8, italic=True)
            _set_col_width(note_cell, COL_C)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def build_lean_docx(lesson_data: dict) -> bytes:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin   = Inches(0.75)
        section.right_margin  = Inches(0.75)

    # Title
    title_p = doc.add_paragraph()
    title_r = title_p.add_run(lesson_data.get("title", "Lean Lesson"))
    title_r.bold = True
    title_r.font.size = Pt(16)
    title_r.font.color.rgb = NAVY
    title_p.paragraph_format.space_after = Pt(4)

    sub_p = doc.add_paragraph()
    sub_r = sub_p.add_run("LEAN LESSON")
    sub_r.bold = True; sub_r.font.size = Pt(9); sub_r.font.color.rgb = CORAL
    sub_p.paragraph_format.space_after = Pt(6)

    _make_header_table(doc, lesson_data)

    for section_data in lesson_data.get("sections", []):
        _section_header(doc, section_data["name"], section_data.get("duration", ""))
        for step in section_data.get("steps", []):
            _step_row(doc, step, show_scaffold=False)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def build_scaffolded_docx(lesson_data: dict) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin   = Inches(0.75)
        section.right_margin  = Inches(0.75)

    title_p = doc.add_paragraph()
    title_r = title_p.add_run(lesson_data.get("title", "Scaffolded Lesson"))
    title_r.bold = True; title_r.font.size = Pt(16); title_r.font.color.rgb = NAVY
    title_p.paragraph_format.space_after = Pt(4)

    sub_p = doc.add_paragraph()
    sub_r = sub_p.add_run("SCAFFOLDED LESSON")
    sub_r.bold = True; sub_r.font.size = Pt(9); sub_r.font.color.rgb = CORAL
    sub_p.paragraph_format.space_after = Pt(6)

    _make_header_table(doc, lesson_data)

    for section_data in lesson_data.get("sections", []):
        _section_header(doc, section_data["name"], section_data.get("duration", ""))
        for step in section_data.get("steps", []):
            _step_row(doc, step, show_scaffold=True)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
# XLSX BUILDER — Scaffold Selector
# ══════════════════════════════════════════════════════════════════════════════

def _xfill(hex_): return PatternFill("solid", fgColor=hex_)
def _xfont(hex_, bold=False, size=10, italic=False):
    return Font(color=hex_, bold=bold, size=size, italic=italic, name="Calibri")
def _xborder():
    s = Side(border_style="thin", color="CCCCDD")
    return Border(top=s, bottom=s, left=s, right=s)


def build_selector_xlsx(selector_rows: list[dict], lesson_data: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scaffold Selector"

    NAVY_F  = _xfill("1F3864"); ALL_F = _xfill("1E8449"); ELL_F = _xfill("195699"); SCD_F = _xfill("6C3483")
    ALL_R   = _xfill("D5F5E3"); ELL_R = _xfill("D6EAF8"); SCD_R = _xfill("E8DAEF")
    ALT_ALL = _xfill("EAF9F0"); ALT_ELL = _xfill("EBF5FB"); ALT_SCD = _xfill("F4ECF7")

    headers = ["Step #", "Tag", "Step Description", "Scaffold?", "ALL", "ELL", "SCD", "Teacher Directive (optional)"]
    col_ws  = [7, 18, 56, 12, 7, 7, 7, 44]

    for c, (h, w) in enumerate(zip(headers, col_ws), 1):
        cell = ws.cell(row=1, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = w
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if c == 5:
            cell.fill = ALL_F; cell.font = _xfont("FFFFFF", bold=True)
        elif c == 6:
            cell.fill = ELL_F; cell.font = _xfont("FFFFFF", bold=True)
        elif c == 7:
            cell.fill = SCD_F; cell.font = _xfont("FFFFFF", bold=True)
        else:
            cell.fill = NAVY_F; cell.font = _xfont("FFFFFF", bold=True)
    ws.row_dimensions[1].height = 28

    # Lesson info in a merged cell above the table
    ws.insert_rows(1)
    title_cell = ws.cell(row=1, column=1, value=f"Scaffold Selector — {lesson_data.get('title','')}  |  Grade {lesson_data.get('grade','')}  |  {lesson_data.get('subject','')}")
    title_cell.fill = _xfill("1F3864"); title_cell.font = _xfont("FFFFFF", bold=True, size=12)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.row_dimensions[1].height = 24

    # Data validation
    dv_yn  = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv_pop = DataValidation(type="list", formula1='"Y,N"',    allow_blank=True)
    ws.add_data_validation(dv_yn)
    ws.add_data_validation(dv_pop)

    for i, row in enumerate(selector_rows):
        r = i + 3  # offset for header rows
        alt = i % 2 == 1
        vals = [row["step_num"], row["tag"], row["content"], row.get("scaffold",""),
                row.get("all",""), row.get("ell",""), row.get("scd",""), row.get("teacher_directive","")]
        for c, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = _xborder()
            cell.alignment = Alignment(horizontal="left" if c > 2 else "center",
                                       vertical="top", wrap_text=True)
            cell.font = _xfont("1A1A2E", bold=(c == 2))
        ws.row_dimensions[r].height = 40

        # Color-code population cells
        for c, (fill_yes, fill_alt) in enumerate([(ALL_R, ALT_ALL), (ELL_R, ALT_ELL), (SCD_R, ALT_SCD)], 5):
            ws.cell(row=r, column=c).fill = fill_alt if alt else fill_yes
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")

        # Apply validations
        dv_yn.add(ws.cell(row=r, column=4))
        for c in [5, 6, 7]:
            dv_pop.add(ws.cell(row=r, column=c))

    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def read_selector_xlsx(xlsx_bytes: bytes) -> list[dict]:
    """Parse a filled selector back into a list of row dicts."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    rows = []
    # Find the header row (contains "Step #")
    header_row = None
    for r in ws.iter_rows():
        for cell in r:
            if str(cell.value or "").strip() == "Step #":
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        return rows

    headers = [str(ws.cell(row=header_row, column=c).value or "").strip()
               for c in range(1, 9)]

    for r in range(header_row + 1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 9)]
        if not any(vals):
            continue
        rows.append({
            "step_num":         vals[0],
            "tag":              str(vals[1] or "").strip(),
            "content":          str(vals[2] or "").strip(),
            "scaffold":         str(vals[3] or "").strip(),
            "all":              str(vals[4] or "").strip(),
            "ell":              str(vals[5] or "").strip(),
            "scd":              str(vals[6] or "").strip(),
            "teacher_directive": str(vals[7] or "").strip(),
        })
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# PPTX BUILDER — Student-Facing Slides
# ══════════════════════════════════════════════════════════════════════════════

SLIDE_W = PInches(10)
SLIDE_H = PInches(5.625)

BG_MAP = {
    "CORAL":  PRGBColor(0xFF, 0x6B, 0x6B),
    "TEAL":   PRGBColor(0x4E, 0xCE, 0xC8),
    "PURPLE": PRGBColor(0x9B, 0x59, 0xB6),
    "AMBER":  PRGBColor(0xF3, 0x9C, 0x12),
    "GREEN":  PRGBColor(0x27, 0xAE, 0x60),
    "PEACH":  PRGBColor(0xFF, 0xD9, 0xC8),
    "WHITE":  PRGBColor(0xFF, 0xFF, 0xFF),
    "NAVY":   PRGBColor(0x1F, 0x38, 0x64),
}
TEXT_WHITE = PRGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK  = PRGBColor(0x1A, 0x1A, 0x2E)

LIGHT_BG = {"PEACH", "WHITE", "AMBER"}


def _slide_bg(slide, color_name: str):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    rgb = BG_MAP.get(color_name.upper(), BG_MAP["TEAL"])
    fill.fore_color.rgb = rgb
    return rgb


def _tb(slide, text, x, y, w, h, size=18, bold=False, color=None,
        align=PP_ALIGN.LEFT, italic=False):
    txb = slide.shapes.add_textbox(PInches(x), PInches(y), PInches(w), PInches(h))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = PPt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    return txb


def _add_rounded_rect(slide, x, y, w, h, color: PRGBColor, radius=0.1):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE — use rounded via XML
        PInches(x), PInches(y), PInches(w), PInches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _text_color(bg_name: str) -> PRGBColor:
    return TEXT_DARK if bg_name.upper() in LIGHT_BG else TEXT_WHITE


def build_slides_pptx(slides_data: list[dict], lesson_data: dict) -> bytes:
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[6]

    tc = _text_color  # alias

    for slide_info in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        bg_name = slide_info.get("bg_color", "TEAL")
        bg_rgb  = _slide_bg(slide, bg_name)
        txt_col = tc(bg_name)
        title_text = slide_info.get("title", "")
        bullets    = slide_info.get("bullets", [])
        featured   = slide_info.get("featured_text", "")
        stype      = slide_info.get("type", "content")
        notes_text = slide_info.get("speaker_notes", "")

        # Decorative accent oval (top-right)
        accent_rgb = PRGBColor(0xFF, 0xFF, 0xFF) if bg_name.upper() not in LIGHT_BG else PRGBColor(0x1F, 0x38, 0x64)
        oval_shape = slide.shapes.add_shape(9, PInches(8.8), PInches(0), PInches(2.5), PInches(2.5))
        oval_shape.fill.solid()
        oval_shape.fill.fore_color.rgb = accent_rgb
        oval_shape.fill.fore_color.transparency = 0.85
        oval_shape.line.fill.background()

        if stype == "title":
            # Big centered title
            _tb(slide, lesson_data.get("grade","") + " · " + lesson_data.get("subject","") + " · " + lesson_data.get("lesson",""),
                0.4, 0.3, 9.2, 0.5, size=13, bold=False, color=txt_col, align=PP_ALIGN.CENTER)
            _tb(slide, title_text, 0.4, 1.0, 9.2, 2.5, size=40, bold=True, color=txt_col, align=PP_ALIGN.CENTER)
            if lesson_data.get("curriculum"):
                _tb(slide, lesson_data["curriculum"], 0.4, 4.9, 9.2, 0.5, size=11, color=txt_col, align=PP_ALIGN.CENTER)

        elif stype == "target":
            _tb(slide, "Learning Target", 0.4, 0.2, 9.2, 0.55, size=14, bold=True, color=txt_col)
            # White card
            card = _add_rounded_rect(slide, 0.5, 0.9, 9.0, 3.8, PRGBColor(0xFF, 0xFF, 0xFF))
            card.fill.fore_color.transparency = 0.1
            _tb(slide, title_text, 0.7, 1.0, 8.6, 3.5, size=26, bold=True,
                color=PRGBColor(0x1A, 0x1A, 0x2E), align=PP_ALIGN.CENTER)
            if bullets:
                _tb(slide, "By the end of today, you will be able to:", 0.4, 4.85, 9.2, 0.5,
                    size=11, italic=True, color=txt_col)

        elif stype == "agenda":
            _tb(slide, title_text, 0.4, 0.2, 9.2, 0.55, size=22, bold=True, color=txt_col)
            y = 1.0
            for b in bullets[:6]:
                _tb(slide, "▸  " + b, 0.6, y, 8.8, 0.6, size=16, color=txt_col)
                y += 0.65

        elif stype in ("content", "materials"):
            _tb(slide, title_text, 0.4, 0.15, 9.2, 0.65, size=22, bold=True, color=txt_col)
            if featured:
                card = _add_rounded_rect(slide, 0.5, 0.95, 9.0, 2.4,
                                         PRGBColor(0xFF, 0xFF, 0xFF))
                card.fill.fore_color.transparency = 0.15
                _tb(slide, featured, 0.7, 1.0, 8.6, 2.3, size=20, bold=True,
                    color=TEXT_DARK, align=PP_ALIGN.CENTER)
                y = 3.5
            else:
                y = 0.95
            for b in bullets[:5]:
                _tb(slide, "▸  " + b, 0.6, y, 8.8, 0.65, size=15, color=txt_col)
                y += 0.7

        elif stype == "scaffold":
            # Peach background override for scaffold slides
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = PRGBColor(0xFF, 0xD9, 0xC8)
            txt_col = TEXT_DARK
            _tb(slide, title_text, 0.4, 0.15, 9.2, 0.65, size=20, bold=True, color=txt_col)
            if featured:
                card = _add_rounded_rect(slide, 0.5, 0.95, 9.0, 2.8, PRGBColor(0xFF, 0xFF, 0xFF))
                _tb(slide, featured, 0.7, 1.05, 8.6, 2.6, size=22, bold=True,
                    color=TEXT_DARK, align=PP_ALIGN.CENTER)
                y = 3.95
            else:
                y = 0.95
            for b in bullets[:4]:
                _tb(slide, "▸  " + b, 0.6, y, 8.8, 0.65, size=15, color=txt_col)
                y += 0.7

        elif stype == "closing":
            _tb(slide, title_text, 0.4, 0.2, 9.2, 0.65, size=22, bold=True, color=txt_col)
            if featured:
                card = _add_rounded_rect(slide, 0.5, 1.0, 9.0, 2.6, PRGBColor(0xFF, 0xFF, 0xFF))
                card.fill.fore_color.transparency = 0.1
                _tb(slide, featured, 0.7, 1.1, 8.6, 2.4, size=20, color=TEXT_DARK, align=PP_ALIGN.CENTER)
                y = 3.75
            else:
                y = 1.0
            for b in bullets[:4]:
                _tb(slide, "▸  " + b, 0.6, y, 8.8, 0.65, size=15, color=txt_col)
                y += 0.7

        # Speaker notes
        if notes_text:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = notes_text

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

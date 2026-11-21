import base64, hashlib, json, uuid, zlib
from datetime import datetime, timedelta, timezone
import html as _html
import io, os, re, time
import streamlit as st
import streamlit.components.v1 as components
from typing import List, Optional
from pydantic import BaseModel
from google import genai
from google.genai import types
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# set_page_config має бути першою командою Streamlit
st.set_page_config(
    page_title="English Teacher's Assistant",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="expanded",
)

MODEL = "gemini-3.5-flash-lite"
FALLBACK_MODEL = "gemini-3.1-flash-lite"


# ---------- Структури відповіді моделі ----------
class Section(BaseModel):
    heading: str
    paragraphs: List[str]


class Slide(BaseModel):
    title: str
    bullets: List[str]
    layout: str = "content"          # content | two_column | section
    bullets_right: List[str] = []    # права колонка для two_column
    notes: str = ""


class DocxData(BaseModel):
    title: str
    sections: List[Section]


class PptxData(BaseModel):
    title: str
    subtitle: str = ""
    theme: str = ""                  # ocean | forest | violet | sunset | slate | teal | rose | mono
    style: str = ""                  # modern | bold | editorial | soft (empty = the app picks)
    slides: List[Slide]


class Answer(BaseModel):
    reply: str                       # коротка відповідь у чаті
    docx: Optional[DocxData] = None  # заповнити, якщо потрібен Word
    pptx: Optional[PptxData] = None  # заповнити, якщо потрібна презентація


SYSTEM = """You are an assistant for a university teacher of English and linguistics.
Tasks: academic articles (structure, editing), linguistics lectures, EN<->UA translation
that preserves terminology, exercises with answer keys, and presentations.
Language: reply in the language the user writes in (English or Ukrainian). Write teaching
materials in English unless the user asks for another language.
Put lectures, exercises, articles and translations in the docx field, presentations in the pptx field.
In the reply field write a short chat message (1-3 sentences), not the whole document.
Exception: if the user wants a chat-only answer (a short explanation, or a translation
shown in chat), put the full answer in reply and leave docx and pptx empty.

Presentation rules:
- The title slide is created automatically from title and subtitle, so do not add it to slides.
- theme: ocean (blue), forest (green), violet (purple), sunset (warm), slate (strict academic),
  teal, rose, mono (black and white). Choose by topic mood or the user's request, and vary
  the choice between presentations: do not always pick ocean.
- style: leave empty unless the user asks for a design style (modern, bold, editorial, soft).
- Each slide: 3-5 short bullet points, up to 12 words each. No long paragraphs.
- layout "content" for regular slides; "two_column" for comparisons (fill bullets and
  bullets_right); "section" for section dividers (title and 0-1 bullet in bullets).
- The last slide must be a summary (Summary / Key takeaways).
- In notes write speaker hints for the teacher (1-3 sentences per slide).

If the request is clear, do not ask questions: use sensible defaults (level B1-B2,
10 exercises, examples with Ukrainian translation when relevant) and briefly state in reply
which assumptions you made. Ask a clarifying question only if the topic or purpose is
completely unclear. Never invent sources or quotations."""

api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))
client = genai.Client(
    api_key=api_key,
    http_options=types.HttpOptions(timeout=40_000),  # 40 секунд
)


# ---------- Word ----------
def build_docx(d):
    doc = Document()
    doc.add_heading(d["title"], 0)
    for s in d["sections"]:
        doc.add_heading(s["heading"], 1)
        for p in s["paragraphs"]:
            doc.add_paragraph(p)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ===== PPTX BUILDER START =====
THEMES = {
    #           dark      accent    accent2   light     text
    "ocean":  ("1F3A5F", "2E86AB", "F6AE2D", "F4F8FB", "1F2937"),
    "forest": ("1E4D3A", "3FA175", "E9B44C", "F3F8F5", "1F2937"),
    "violet": ("3B2F63", "7C5CBF", "F2A65A", "F7F5FB", "1F2937"),
    "sunset": ("5C2A2E", "C8553D", "F0A202", "FBF6F3", "2B1B1B"),
    "slate":  ("2B2D42", "4A6FA5", "D4A373", "F5F6F8", "1F2937"),
    "teal":   ("12433F", "14A38B", "F4B860", "F2F9F8", "1F2937"),
    "rose":   ("4A1F3D", "D6457A", "F7B538", "FCF4F7", "2B1B24"),
    "mono":   ("111827", "374151", "F59E0B", "F9FAFB", "111827"),
}
HEAD_FONT = "Georgia"
BODY_FONT = "Calibri"
SW, SH = 13.333, 7.5  # розмір слайда (16:9), дюйми


def _rgb(h):
    return RGBColor.from_string(h)


def _bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(color)


def _shape(slide, kind, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = _rgb(fill)
    if line:
        s.line.color.rgb = _rgb(line)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def _text(slide, x, y, w, h, text, size, color, bold=False, font=BODY_FONT,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = font
    r.font.color.rgb = _rgb(color)
    return tb


def _bullets(slide, x, y, w, h, items, size, color, bullet_color, font=BODY_FONT, char="●"):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(size * 0.7)
        r = p.add_run()
        r.text = item
        r.font.size = Pt(size)
        r.font.name = font
        r.font.color.rgb = _rgb(color)
        # справжній маркер списку з кольором теми
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Emu(Inches(0.38))))
        pPr.set("indent", str(-Emu(Inches(0.38))))
        buClr = etree.SubElement(pPr, qn("a:buClr"))
        clr = etree.SubElement(buClr, qn("a:srgbClr"))
        clr.set("val", bullet_color)
        buFont = etree.SubElement(pPr, qn("a:buFont"))
        buFont.set("typeface", "Arial")
        buChar = etree.SubElement(pPr, qn("a:buChar"))
        buChar.set("char", char)
    return tb


def _font_size(items, two_col=False):
    n = len(items)
    total = sum(len(i) for i in items)
    if n <= 4 and total < 260:
        size = 26
    elif n <= 5 and total < 380:
        size = 24
    elif total < 520:
        size = 22
    else:
        size = 20
    return size - 3 if two_col else size


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


def _title_modern(prs, d, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, dark)
    _shape(s, MSO_SHAPE.OVAL, 9.2, -1.8, 6.6, 6.6, accent)
    _shape(s, MSO_SHAPE.OVAL, 11.3, 4.9, 1.9, 1.9, accent2)
    _shape(s, MSO_SHAPE.OVAL, 8.6, 5.6, 0.7, 0.7, accent)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.9, 2.35, 1.4, 0.09, accent2)
    _text(s, 0.85, 2.6, 8.4, 2.2, d["title"], 44, "FFFFFF", bold=True,
          font=HEAD_FONT, anchor=MSO_ANCHOR.TOP)
    if d.get("subtitle"):
        _text(s, 0.85, 4.85, 8.0, 1.0, d["subtitle"], 22, light)


def _section_modern(prs, sl, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, accent)
    _shape(s, MSO_SHAPE.OVAL, 10.2, 3.6, 5.4, 5.4, dark)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.9, 2.9, 1.4, 0.09, accent2)
    _text(s, 0.85, 3.15, 9.0, 1.6, sl["title"], 40, "FFFFFF", bold=True, font=HEAD_FONT)
    if sl.get("bullets"):
        _text(s, 0.85, 4.65, 8.5, 1.2, sl["bullets"][0], 22, "FFFFFF")
    _notes(s, sl.get("notes"))


def _content_modern(prs, sl, t, num, deck_title):
    dark, accent, accent2, light, text = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, light)
    _shape(s, MSO_SHAPE.RECTANGLE, 0, 0, 0.28, SH, accent)
    tsize = 32 if len(sl["title"]) <= 42 else 26
    _text(s, 0.8, 0.4, 11.8, 1.05, sl["title"], tsize, dark, bold=True,
          font=HEAD_FONT, anchor=MSO_ANCHOR.MIDDLE)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.85, 1.5, 1.3, 0.07, accent2)

    if sl.get("layout") == "two_column" and sl.get("bullets_right"):
        for cx, items in ((0.8, sl["bullets"]), (6.85, sl["bullets_right"])):
            card = _shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 1.9, 5.7, 4.7,
                          "FFFFFF", line="DDE3EA")
            card.adjustments[0] = 0.04
            _shape(s, MSO_SHAPE.RECTANGLE, cx + 0.3, 2.1, 0.9, 0.06, accent)
            _bullets(s, cx + 0.3, 2.35, 5.1, 4.1, items,
                     _font_size(items, True), text, accent)
    else:
        _bullets(s, 0.85, 1.95, 11.6, 4.7, sl["bullets"],
                 _font_size(sl["bullets"]), text, accent)

    # футер: назва презентації і номер слайда
    _text(s, 0.8, 6.95, 8.0, 0.35, deck_title, 11, "8A94A6")
    badge = _shape(s, MSO_SHAPE.OVAL, 12.2, 6.85, 0.5, 0.5, accent)
    tf = badge.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(num)
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.name = BODY_FONT
    r.font.color.rgb = _rgb("FFFFFF")
    _notes(s, sl.get("notes"))


# ---------- Стилі презентацій: modern | bold | editorial | soft ----------
STYLES = ["modern", "bold", "editorial", "soft"]
BOLD_FONT, SOFT_FONT = "Arial", "Trebuchet MS"


def _mix(a, b, k):
    """Змішує два кольори (hex): k=0 дає a, k=1 дає b."""
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * k):02X}" for x, y in zip(ca, cb))


def _h(d, shift):
    return int(hashlib.md5((d.get("title") or "").encode("utf-8")).hexdigest(), 16) >> shift


def resolve_style(d):
    s = (d.get("style") or "").lower()
    return s if s in STYLES else STYLES[_h(d, 0) % len(STYLES)]


def resolve_theme(d):
    k = (d.get("theme") or "").lower()
    keys = list(THEMES)
    return THEMES[k if k in THEMES else keys[_h(d, 8) % len(keys)]]


def _two_col(sl):
    return sl.get("layout") == "two_column" and sl.get("bullets_right")


def _num_badge(s, x, y, w, h, num, fill, fg):
    b = _shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill)
    b.adjustments[0] = 0.3
    tf = b.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(num)
    r.font.size, r.font.bold, r.font.name = Pt(13), True, BODY_FONT
    r.font.color.rgb = _rgb(fg)


# ----- bold: темний фон, великі заголовки -----
def _title_bold(prs, d, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, dark)
    _shape(s, MSO_SHAPE.RECTANGLE, 10.2, 0, SW - 10.2, SH, accent)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.9, 1.5, 2.2, 0.16, accent2)
    _text(s, 0.85, 1.9, 9.0, 3.3, d["title"], 52, "FFFFFF", bold=True, font=BOLD_FONT)
    if d.get("subtitle"):
        _text(s, 0.85, 5.3, 8.8, 1.0, d["subtitle"], 22, accent2)


def _section_bold(prs, sl, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, accent)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.9, 2.75, 2.2, 0.16, accent2)
    _text(s, 0.85, 3.05, 10.5, 1.8, sl["title"], 46, "FFFFFF", bold=True, font=BOLD_FONT)
    if sl.get("bullets"):
        _text(s, 0.85, 4.9, 9.0, 1.0, sl["bullets"][0], 22, "FFFFFF")
    _notes(s, sl.get("notes"))


def _content_bold(prs, sl, t, num, deck_title):
    dark, accent, accent2, light, text = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, dark)
    _text(s, 0.8, 0.4, 11.8, 1.05, sl["title"], 32 if len(sl["title"]) <= 42 else 26,
          "FFFFFF", bold=True, font=BOLD_FONT, anchor=MSO_ANCHOR.MIDDLE)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.85, 1.5, 1.6, 0.1, accent2)
    if _two_col(sl):
        for cx, items in ((0.8, sl["bullets"]), (6.85, sl["bullets_right"])):
            c = _shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 1.9, 5.7, 4.7, _mix(dark, "FFFFFF", 0.1))
            c.adjustments[0] = 0.04
            _bullets(s, cx + 0.3, 2.2, 5.1, 4.2, items, _font_size(items, True), "FFFFFF", accent2)
    else:
        _bullets(s, 0.85, 1.95, 11.6, 4.7, sl["bullets"], _font_size(sl["bullets"]), "FFFFFF", accent2)
    _text(s, 11.0, 6.55, 1.9, 0.8, f"{num:02d}", 36, accent, bold=True, font=BOLD_FONT, align=PP_ALIGN.RIGHT)
    _notes(s, sl.get("notes"))


# ----- editorial: білий аркуш, серифи, тонкі лінії -----
def _title_editorial(prs, d, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, light)
    _shape(s, MSO_SHAPE.RECTANGLE, 0.9, 1.3, 0.07, 4.5, accent)
    _text(s, 1.35, 1.3, 9.8, 3.3, d["title"], 48, dark, bold=True, font=HEAD_FONT)
    if d.get("subtitle"):
        _text(s, 1.35, 4.7, 9.5, 1.0, d["subtitle"], 20, "6B7280")
    _shape(s, MSO_SHAPE.RECTANGLE, 1.35, 6.6, 10.8, 0.02, "CBD2D9")


def _section_editorial(prs, sl, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, light)
    _shape(s, MSO_SHAPE.RECTANGLE, 1.35, 2.9, 1.0, 0.07, accent)
    _text(s, 1.3, 3.15, 10.5, 1.6, sl["title"], 42, dark, bold=True, font=HEAD_FONT)
    if sl.get("bullets"):
        _text(s, 1.3, 4.7, 9.5, 1.0, sl["bullets"][0], 20, "6B7280")
    _notes(s, sl.get("notes"))


def _content_editorial(prs, sl, t, num, deck_title):
    dark, accent, accent2, light, text = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, "FFFFFF")
    _text(s, 1.2, 0.5, 11.0, 0.95, sl["title"], 30 if len(sl["title"]) <= 42 else 25, dark,
          bold=True, font=HEAD_FONT, anchor=MSO_ANCHOR.MIDDLE)
    _shape(s, MSO_SHAPE.RECTANGLE, 1.25, 1.55, 10.85, 0.015, "CBD2D9")
    if _two_col(sl):
        _bullets(s, 1.2, 1.95, 5.1, 4.6, sl["bullets"], _font_size(sl["bullets"], True), text, accent, char="–")
        _shape(s, MSO_SHAPE.RECTANGLE, 6.75, 2.0, 0.015, 4.3, "CBD2D9")
        _bullets(s, 7.1, 1.95, 5.1, 4.6, sl["bullets_right"], _font_size(sl["bullets_right"], True),
                 text, accent, char="–")
    else:
        _bullets(s, 1.2, 1.95, 10.9, 4.7, sl["bullets"], _font_size(sl["bullets"]), text, accent, char="–")
    _text(s, 1.2, 6.95, 8.0, 0.35, deck_title, 11, "8A94A6")
    _text(s, 11.5, 6.95, 0.7, 0.35, str(num), 12, "8A94A6", font=HEAD_FONT, align=PP_ALIGN.RIGHT)
    _notes(s, sl.get("notes"))


# ----- soft: пастельні заокруглені форми -----
def _title_soft(prs, d, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, light)
    _shape(s, MSO_SHAPE.OVAL, -1.6, -1.8, 5.2, 5.2, _mix(accent, "FFFFFF", 0.7))
    _shape(s, MSO_SHAPE.OVAL, 10.6, 4.7, 3.6, 3.6, _mix(accent2, "FFFFFF", 0.55))
    _shape(s, MSO_SHAPE.OVAL, 11.4, 0.9, 0.8, 0.8, accent)
    _text(s, 1.6, 2.4, 10.1, 2.2, d["title"], 44, dark, bold=True, font=SOFT_FONT,
          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if d.get("subtitle"):
        _text(s, 1.6, 4.75, 10.1, 1.0, d["subtitle"], 22, accent, align=PP_ALIGN.CENTER)


def _section_soft(prs, sl, t):
    dark, accent, accent2, light, _ = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, _mix(accent, "FFFFFF", 0.78))
    _shape(s, MSO_SHAPE.OVAL, 9.6, 3.4, 5.6, 5.6, accent)
    _shape(s, MSO_SHAPE.OVAL, 8.6, 1.2, 1.0, 1.0, accent2)
    _text(s, 0.9, 2.9, 8.2, 1.8, sl["title"], 40, dark, bold=True, font=SOFT_FONT)
    if sl.get("bullets"):
        _text(s, 0.9, 4.7, 7.8, 1.0, sl["bullets"][0], 22, dark)
    _notes(s, sl.get("notes"))


def _content_soft(prs, sl, t, num, deck_title):
    dark, accent, accent2, light, text = t
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, light)
    pill = _shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.8, 0.5, 11.7, 1.05, accent)
    pill.adjustments[0] = 0.5
    _text(s, 1.25, 0.5, 10.8, 1.05, sl["title"], 28 if len(sl["title"]) <= 42 else 24, "FFFFFF",
          bold=True, font=SOFT_FONT, anchor=MSO_ANCHOR.MIDDLE)
    edge = _mix(accent, "FFFFFF", 0.6)
    if _two_col(sl):
        for cx, items in ((0.8, sl["bullets"]), (6.85, sl["bullets_right"])):
            c = _shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 1.9, 5.7, 4.75, "FFFFFF", line=edge)
            c.adjustments[0] = 0.05
            _bullets(s, cx + 0.35, 2.25, 5.0, 4.2, items, _font_size(items, True), text, accent)
    else:
        c = _shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.8, 1.9, 11.7, 4.75, "FFFFFF", line=edge)
        c.adjustments[0] = 0.04
        _bullets(s, 1.25, 2.3, 10.8, 4.2, sl["bullets"], _font_size(sl["bullets"]), text, accent)
    _num_badge(s, 11.9, 6.85, 0.6, 0.42, num, accent2, dark)
    _notes(s, sl.get("notes"))


STYLE_FUNCS = {
    "modern": (_title_modern, _section_modern, _content_modern),
    "bold": (_title_bold, _section_bold, _content_bold),
    "editorial": (_title_editorial, _section_editorial, _content_editorial),
    "soft": (_title_soft, _section_soft, _content_soft),
}


def build_pptx(d, template_path=None):
    t = resolve_theme(d)
    f_title, f_section, f_content = STYLE_FUNCS[resolve_style(d)]
    prs = Presentation(template_path) if template_path else Presentation()
    if not template_path:
        prs.slide_width = Inches(SW)
        prs.slide_height = Inches(SH)
    f_title(prs, d, t)
    for i, sl in enumerate(d["slides"], start=2):
        if sl.get("layout") == "section":
            f_section(prs, sl, t)
        else:
            f_content(prs, sl, t, i, d["title"])
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
# ===== PPTX BUILDER END =====


# ---------- Перегляд у браузері (без завантаження) ----------
def _e(t):
    return _html.escape(str(t or ""))


def docx_preview_html(d):
    parts = ['<div class="doc-page">', f'<h1>{_e(d["title"])}</h1>']
    for sec in d["sections"]:
        parts.append(f'<h2>{_e(sec["heading"])}</h2>')
        for p in sec["paragraphs"]:
            parts.append("<p>" + _e(p).replace(chr(10), "<br>") + "</p>")
    parts.append("</div>")
    return "".join(parts)


def _ul(items, size_pt, color, marker):
    lis = "".join(f"<li>{_e(i)}</li>" for i in items)
    return (f'<ul class="s-ul" style="font-size:{size_pt / 9.6:.2f}cqw;color:#{color};'
            f'--c:#{marker}">{lis}</ul>')


def pptx_preview_html(d, large=False):
    dark, accent, accent2, light, text = resolve_theme(d)
    sty = resolve_style(d)
    vs = f"--dark:#{dark};--accent:#{accent};--accent2:#{accent2};--light:#{light};--text:#{text}"
    uc, um = ("FFFFFF", accent2) if sty == "bold" else (text, accent)

    def wrap(kind, inner, note=""):
        return (f'<div class="slide-wrap"><div class="slide st-{sty} k-{kind}" style="{vs}">'
                f'{inner}</div>{note}</div>')

    def head(title, subtitle):
        sub = f'<div class="s-sub">{_e(subtitle)}</div>' if subtitle else ""
        return f'<div class="s-body"><div class="s-bar"></div><div class="s-big">{_e(title)}</div>{sub}</div>'

    cards = [wrap("title", head(d["title"], d.get("subtitle")))]
    for num, sl in enumerate(d["slides"], start=2):
        note = f'<div class="s-note">🗒 {_e(sl.get("notes"))}</div>' if sl.get("notes") else ""
        if sl.get("layout") == "section":
            cards.append(wrap("section", head(sl["title"], (sl.get("bullets") or [""])[0]), note))
            continue
        if _two_col(sl):
            cols = "".join(
                f'<div class="s-card">{_ul(items, _font_size(items, True), uc, um)}</div>'
                for items in (sl["bullets"], sl["bullets_right"]))
            body = f'<div class="s-cols">{cols}</div>'
        else:
            body = _ul(sl["bullets"], _font_size(sl["bullets"]), uc, um)
        n = f"{num:02d}" if sty == "bold" else num
        cards.append(wrap(
            "content",
            f'<div class="s-body"><div class="s-h">{_e(sl["title"])}</div><div class="s-line"></div>{body}</div>'
            f'<div class="s-num">{n}</div>', note))

    cls = "slides large" if large else "slides"
    return f'<div class="{cls}">' + "".join(cards) + "</div>"


def render_previews(previews, key_prefix, expanded):
    for k, (kind, data) in enumerate(previews):
        if kind == "docx":
            with st.expander("👁 Preview document", expanded=expanded):
                st.markdown(docx_preview_html(data), unsafe_allow_html=True)
                st.caption("Preview is approximate. Download the file to get the final Word formatting.")
        else:
            with st.expander("👁 Preview presentation", expanded=expanded):
                mode = st.radio("View", ["Grid", "Large"], horizontal=True,
                                key=f"{key_prefix}_vm{k}", label_visibility="collapsed")
                st.markdown(pptx_preview_html(data, mode == "Large"), unsafe_allow_html=True)
                st.caption("Preview is approximate. Download the file to get the final PowerPoint design.")




def file_name(title, ext):
    """Безпечна назва файлу з заголовка документа."""
    base = re.sub(r"[^\w\- ]+", "", title or "", flags=re.UNICODE).strip().replace(" ", "_")
    return f"{base[:60] or 'document'}.{ext}"


def render_files(files, key_prefix):
    for j, (name, content, mime) in enumerate(files):
        label = "📊 Download presentation (PowerPoint)" if name.endswith(".pptx") \
            else "📄 Download document (Word)"
        st.download_button(label, content, file_name=name, mime=mime, key=f"{key_prefix}_{j}")



THEME_OPTIONS = {
    "Auto": "",
    "🔵 Blue": "ocean",
    "🟢 Green": "forest",
    "🟣 Purple": "violet",
    "🟠 Warm": "sunset",
    "⚫ Academic": "slate",
    "🩵 Teal": "teal",
    "🌸 Rose": "rose",
    "⬜ Black & white": "mono",
}

STYLE_OPTIONS = {
    "Auto (varies)": "",
    "Modern": "modern",
    "Bold (dark)": "bold",
    "Editorial (clean)": "editorial",
    "Soft (rounded)": "soft",
}

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&family=Inter:wght@400;500;600&family=Lora:wght@600;700&display=swap');

:root {
    --navy: #0B1226; --cloud: #F4F7FB; --lime: #C8F535;
    --ink: #0B1226; --muted: #5B6785; --line: #DEE4F0; --r: 16px;
}
html, body, [class*="css"], .stApp { font-family: 'Inter', sans-serif; color: var(--ink); }
.stApp { background: var(--cloud); }
.block-container { padding-top: 1.6rem; padding-bottom: 6rem; max-width: 980px; }
#MainMenu, footer, [data-testid="stDeployButton"] { visibility: hidden; display: none; }
header[data-testid="stHeader"] { background: transparent; }
.cat { display: none; }
.stMarkdown h4, .stMarkdown h5 { font-family: 'Sora', sans-serif; letter-spacing: -.01em; color: var(--ink); }

/* шапка: компактна смуга */
.hero { display: flex; align-items: center; justify-content: space-between; gap: 14px; flex-wrap: wrap;
        background: var(--navy); border-radius: 20px; padding: 18px 26px; margin: 0 0 1.5rem; }
.hero h1 { font-family: 'Sora', sans-serif; font-weight: 700; font-size: 1.4rem; letter-spacing: -.02em; color: #fff;
           margin: 0; padding: 0; display: flex; align-items: center; gap: 10px; }
.hero h1::before { content: ""; width: 12px; height: 12px; border-radius: 50%; background: var(--lime); flex: none; }
.hero p, .hero-badge { display: none; }
.chips { display: flex; gap: 8px; flex-wrap: wrap; }
.chips span { font-size: .78rem; font-weight: 500; color: var(--lime); border: 1px solid rgba(200,245,53,.4); padding: 4px 12px; border-radius: 999px; }

/* чат */
[data-testid="stChatMessage"] { background: transparent; border: none; box-shadow: none; padding: .6rem 0; }
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) { background: var(--navy); border-radius: var(--r); padding: .9rem 1.2rem; }
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) * { color: #fff; }
[data-testid="stChatInput"] { border-radius: var(--r); border: 1px solid var(--line); background: #fff; }
[data-testid="stChatInput"] button { background: var(--lime); color: var(--navy); border-radius: 10px; }

/* поля */
[data-baseweb="input"], [data-baseweb="textarea"], [data-baseweb="select"] > div { border-radius: 12px; }
[data-testid="stExpander"] { border-radius: var(--r); border-color: var(--line); background: #fff; }

/* кнопки */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
    border-radius: 12px; font-weight: 600; font-size: .92rem; box-shadow: none;
    background: #fff; color: var(--ink); border: 1px solid var(--line);
    transition: box-shadow .12s ease, transform .12s ease, border-color .12s ease, background .12s ease;
}
.stButton > button:hover { border-color: var(--navy); color: var(--navy); background: #fff; }
.stDownloadButton > button, .stFormSubmitButton > button { background: var(--lime); color: var(--navy); border: 1px solid var(--lime); }
.stDownloadButton > button p, .stFormSubmitButton > button p { color: var(--navy); }
.stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
    background: #D6FF4D; border-color: var(--navy); color: var(--navy); box-shadow: 3px 3px 0 var(--navy); transform: translate(-1px, -1px);
}
.stButton > button:focus-visible, .stDownloadButton > button:focus-visible { outline: 2px solid var(--navy); outline-offset: 2px; }

/* картки */
[data-testid="stVerticalBlockBorderWrapper"]:has(.cat) {
    background: #fff; border: 1px solid var(--line); border-radius: 20px; box-shadow: none;
    transition: box-shadow .15s ease, transform .15s ease, border-color .15s ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.card-ico):hover {
    border-color: var(--navy); box-shadow: 5px 5px 0 var(--lime); transform: translate(-2px, -2px);
}
.card-ico { width: 44px; height: 44px; border-radius: 12px; background: var(--navy); display: flex; align-items: center; justify-content: center; font-size: 1.25rem; margin-bottom: 12px; }
.card-title { font-family: 'Sora', sans-serif; font-weight: 600; font-size: 1.08rem; color: var(--ink); margin-bottom: 3px; }
.card-desc { font-size: .86rem; color: var(--muted); margin-bottom: 12px; }

/* бічна панель */
[data-testid="stSidebar"] { background: var(--navy); border-right: none; }
[data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color: #AEB9D6; }
[data-testid="stSidebar"] .stMarkdown strong { color: #8591B3; font-size: .78rem; font-weight: 600; letter-spacing: .04em; }
[data-testid="stSidebarCollapseButton"] * { color: #AEB9D6; }
.side-title { font-family: 'Sora', sans-serif; font-weight: 700; font-size: 1.3rem; color: #fff; display: flex; align-items: center; gap: 9px; }
.side-title .dot { width: 12px; height: 12px; border-radius: 50%; background: var(--lime); display: inline-block; }
.side-sub { font-size: .82rem; color: #8591B3; margin-bottom: 14px; }
[data-testid="stSidebar"] .stButton > button {
    justify-content: flex-start; text-align: left; font-weight: 500; padding: .4rem .8rem;
    border-color: transparent; background: transparent !important; color: #E8EEF9 !important; border-radius: 10px;
}
[data-testid="stSidebar"] .stButton > button p { text-align: left; }
[data-testid="stSidebar"] .stButton > button * { color: inherit !important; }
[data-testid="stSidebar"] button:hover, [data-testid="stSidebar"] button:focus,
[data-testid="stSidebar"] button:active, [data-testid="stSidebar"] button:focus-visible {
    background: transparent !important; background-color: transparent !important;
    color: var(--lime) !important; border-color: transparent !important; box-shadow: none !important; outline: none !important;
}
[data-testid="stSidebar"] button:hover *, [data-testid="stSidebar"] button:focus *,
[data-testid="stSidebar"] button:active * { color: var(--lime) !important; }
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: .35rem; }
[data-testid="stSidebar"] [class*="st-key-d_"] button { padding: 0 !important; justify-content: center; min-width: 0; }
[data-testid="stSidebar"] [class*="st-key-d_"] button:hover, [data-testid="stSidebar"] [class*="st-key-d_"] button:hover * { color: #FF8A8A !important; }
[data-testid="stSidebar"] .st-key-new_chat button { background: var(--lime) !important; color: var(--navy) !important; font-weight: 700; justify-content: center; margin-bottom: .6rem; }
[data-testid="stSidebar"] .st-key-new_chat button p { color: var(--navy) !important; }
[data-testid="stSidebar"] .st-key-new_chat button:hover, [data-testid="stSidebar"] .st-key-new_chat button:focus,
[data-testid="stSidebar"] .st-key-new_chat button:active {
    background: #D6FF4D !important; color: var(--navy) !important; }
[data-testid="stSidebar"] .st-key-new_chat button:hover p, [data-testid="stSidebar"] .st-key-new_chat button:focus p,
[data-testid="stSidebar"] .st-key-new_chat button:active p { color: var(--navy) !important; }
[data-testid="stSidebar"] [data-testid="stExpander"] { background: rgba(255,255,255,.05); border-color: rgba(255,255,255,.12); }
[data-testid="stSidebar"] [data-testid="stExpander"] summary, [data-testid="stSidebar"] [data-testid="stExpander"] summary * { color: #E8EEF9; }
[data-testid="stSidebar"] [data-testid="stAlert"] { background: #FFF6D6; }
[data-testid="stSidebar"] [data-testid="stAlert"] p { color: #4A3A00; }

/* бічна панель: Tips і історія */
[data-testid="stSidebar"] [data-testid="stExpander"] { border: 1px solid rgba(255,255,255,.1); border-radius: 14px; background: rgba(255,255,255,.04); }
[data-testid="stSidebar"] [data-testid="stExpander"] summary,
[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover,
[data-testid="stSidebar"] [data-testid="stExpander"] summary:focus,
[data-testid="stSidebar"] [data-testid="stExpander"] summary:focus-visible { background: transparent !important; outline: none !important; box-shadow: none !important; }
[data-testid="stSidebar"] [data-testid="stExpanderDetails"], [data-testid="stSidebar"] [data-testid="stExpanderDetails"] * { color: #C9D3EC !important; }
[data-testid="stSidebar"] [data-testid="stExpanderDetails"] li { margin-bottom: .4rem; line-height: 1.5; font-size: .88rem; }
[data-testid="stSidebar"] [data-testid="stExpanderDetails"] code, [data-testid="stSidebar"] [data-testid="stExpanderDetails"] pre { background: rgba(255,255,255,.08); }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"], [data-testid="stSidebar"] [data-testid="stCaptionContainer"] * { color: #8F9BBD !important; font-size: .8rem; line-height: 1.5; }

[data-testid="stSidebar"] [class*="st-key-h_"] .stButton > button, [data-testid="stSidebar"] [class*="st-key-hc_"] .stButton > button {
    font-size: .9rem; font-weight: 400; border-radius: 10px; padding: .5rem .8rem; border-left: 3px solid transparent;
}
[data-testid="stSidebar"] [class*="st-key-h_"] .stButton > button { color: #C9D3EC !important; }
[data-testid="stSidebar"] [class*="st-key-hc_"] .stButton > button {
    background: rgba(200,245,53,.12) !important; color: var(--lime) !important; font-weight: 600; border-left-color: var(--lime);
}
[data-testid="stSidebar"] [class*="st-key-hc_"] .stButton > button:hover { background: rgba(200,245,53,.18) !important; }
[data-testid="stSidebar"] [class*="st-key-d_"] .stButton > button { color: #6F7BA0 !important; }

/* перегляд документа */
.doc-page {
    background: #fff; border: 1px solid var(--line); border-radius: var(--r); padding: 38px 46px;
    max-height: 620px; overflow-y: auto; font-family: 'Lora', Georgia, serif; color: #1F2937; line-height: 1.65;
}
.doc-page h1 { font-size: 1.7rem; color: var(--ink); margin: 0 0 18px; padding: 0; }
.doc-page h2 { font-size: 1.2rem; color: var(--ink); margin: 22px 0 8px; padding: 0; }
.doc-page p { margin: 0 0 10px; font-size: .98rem; }

/* перегляд презентації: кольори беруться зі змінних слайда */
.slides { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.slides.large { grid-template-columns: 1fr; }
.slide {
    position: relative; aspect-ratio: 16 / 9; border-radius: 10px; overflow: hidden;
    container-type: inline-size; border: 1px solid var(--line); background: var(--light);
}
.slide::before, .slide::after { content: ""; position: absolute; pointer-events: none; }
.s-body { position: absolute; inset: 0; padding: 4.5cqw 5cqw 4cqw 6.5cqw; display: flex; flex-direction: column; }
.k-title .s-body, .k-section .s-body { justify-content: center; }
.s-bar { width: 10cqw; height: .7cqw; margin-bottom: 2cqw; background: var(--accent2); }
.s-big { font-family: 'Lora', serif; font-weight: 700; color: #fff; font-size: 4.6cqw; line-height: 1.15; max-width: 62%; }
.s-sub { color: #E8EEF4; font-size: 2.1cqw; margin-top: 2cqw; max-width: 60%; }
.s-h { font-family: 'Lora', serif; font-weight: 700; font-size: 3.3cqw; line-height: 1.2; color: var(--dark); }
.s-line { width: 9cqw; height: .5cqw; margin: 1.2cqw 0 2.2cqw; background: var(--accent2); }
.s-ul { margin: 0; padding-left: 2.6cqw; line-height: 1.3; }
.s-ul li { margin-bottom: .9cqw; }
.s-ul li::marker { color: var(--c); }
.s-cols { display: grid; grid-template-columns: 1fr 1fr; gap: 2cqw; }
.s-card { background: #fff; border: 1px solid #DDE3EA; border-radius: 1.2cqw; padding: 1.8cqw 1.6cqw 1cqw; }
.s-num {
    position: absolute; right: 2.5cqw; bottom: 2cqw; width: 3.6cqw; height: 3.6cqw; border-radius: 50%;
    background: var(--accent); color: #fff; font-size: 1.6cqw; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
}
.s-note { font-size: .78rem; color: var(--muted); margin: 6px 4px 4px; }

/* modern */
.st-modern.k-title { background: var(--dark); }
.st-modern.k-title::after { width: 42%; aspect-ratio: 1; border-radius: 50%; background: var(--accent); right: -6%; top: -32%; }
.st-modern.k-title::before { width: 13%; aspect-ratio: 1; border-radius: 50%; background: var(--accent2); right: 4%; bottom: 6%; }
.st-modern.k-section { background: var(--accent); }
.st-modern.k-section::after { width: 38%; aspect-ratio: 1; border-radius: 50%; background: var(--dark); right: -8%; bottom: -40%; }
.st-modern.k-content::before { left: 0; top: 0; bottom: 0; width: 2%; background: var(--accent); }

/* bold */
.st-bold { background: var(--dark); }
.st-bold.k-section { background: var(--accent); }
.st-bold.k-title::before { right: 0; top: 0; bottom: 0; width: 23%; background: var(--accent); }
.st-bold .s-big, .st-bold .s-h { font-family: Arial, Helvetica, sans-serif; font-weight: 800; color: #fff; }
.st-bold .s-big { font-size: 5.2cqw; max-width: 68%; }
.st-bold .s-bar { width: 16cqw; height: 1.2cqw; }
.st-bold .s-line { height: .8cqw; width: 12cqw; }
.st-bold .s-sub { color: var(--accent2); }
.st-bold .s-card { background: rgba(255,255,255,.09); border-color: transparent; }
.st-bold .s-num { background: none; color: var(--accent); font-size: 4.2cqw; width: auto; height: auto; right: 3cqw; bottom: 1.6cqw; }

/* editorial */
.st-editorial { background: #fff; }
.st-editorial.k-title, .st-editorial.k-section { background: var(--light); }
.st-editorial.k-title::before, .st-editorial.k-section::before { left: 8%; top: 18%; bottom: 18%; width: .5%; background: var(--accent); }
.st-editorial .s-body { padding-left: 13cqw; }
.st-editorial .s-big { color: var(--dark); font-weight: 600; }
.st-editorial .s-sub { color: #6B7280; }
.st-editorial .s-bar { display: none; }
.st-editorial .s-h { font-weight: 600; }
.st-editorial .s-line { width: 100%; height: 1px; background: #CBD2D9; margin: 1.4cqw 0 2.4cqw; }
.st-editorial .s-num { background: none; color: #8A94A6; font-size: 1.6cqw; font-weight: 400; font-style: italic; }
.st-editorial .s-cols { gap: 5cqw; }
.st-editorial .s-card { background: none; border: none; border-left: 1px solid #CBD2D9; border-radius: 0; }

/* soft */
.st-soft.k-title::before { width: 32%; aspect-ratio: 1; border-radius: 50%; background: color-mix(in srgb, var(--accent) 30%, #fff); left: -8%; top: -24%; }
.st-soft.k-title::after { width: 24%; aspect-ratio: 1; border-radius: 50%; background: color-mix(in srgb, var(--accent2) 45%, #fff); right: -4%; bottom: -14%; }
.st-soft.k-section { background: color-mix(in srgb, var(--accent) 22%, #fff); }
.st-soft.k-section::after { width: 34%; aspect-ratio: 1; border-radius: 50%; background: var(--accent); right: -6%; bottom: -30%; }
.st-soft.k-title .s-body { align-items: center; text-align: center; }
.st-soft .s-big { color: var(--dark); font-family: 'Inter', sans-serif; font-weight: 700; }
.st-soft .s-sub { color: var(--accent); }
.st-soft.k-section .s-sub { color: var(--dark); }
.st-soft .s-bar, .st-soft .s-line { display: none; }
.st-soft .s-body { padding-top: 4cqw; }
.st-soft .s-h { background: var(--accent); color: #fff; font-family: 'Inter', sans-serif; font-size: 2.8cqw; padding: 1.2cqw 2.6cqw; border-radius: 99px; align-self: flex-start; }
.st-soft .s-ul { background: #fff; border: 1px solid color-mix(in srgb, var(--accent) 35%, #fff); border-radius: 2cqw; padding: 2cqw 2.4cqw 1.2cqw 5.4cqw; margin-top: 2.2cqw; }
.st-soft .s-cols { margin-top: 2.2cqw; }
.st-soft .s-card { border-radius: 2cqw; border-color: color-mix(in srgb, var(--accent) 35%, #fff); }
.st-soft .s-card .s-ul { background: none; border: none; padding: 0 0 0 2.6cqw; margin: 0; }
.st-soft .s-num { border-radius: 1cqw; background: var(--accent2); color: var(--dark); }
</style>
"""

HERO = """
<div class="hero">
    <h1>English Teacher's Assistant</h1>
    <div class="chips"><span>Word</span><span>PowerPoint</span><span>EN ↔ UA</span></div>
</div>
"""


# ---------- Моделі, помилки, запити до AI ----------
MODELS = []
for _m in (MODEL, FALLBACK_MODEL, "gemini-3.5-flash"):  # основна, запасна, ще одна
    if _m not in MODELS:
        MODELS.append(_m)

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

ERRORS = {
    "quota_day": (
        "Daily free limit reached",
        "The AI's free daily request limit has been used up. It resets automatically, "
        "usually around 10:00 Kyiv time. Everything you created is saved in History on the "
        "left, so you can still open and download it. To work without daily limits, the app "
        "owner can switch the app to a paid API key."),
    "quota_min": (
        "Too many requests at once",
        "The per-minute limit was reached. Wait about a minute and press Try again."),
    "key": (
        "Access problem",
        "The AI service rejected the app's API key. Please ask the app owner to check or renew it."),
    "busy": (
        "The AI service is busy",
        "Google's servers are overloaded right now. This is usually temporary: try again in a minute."),
    "timeout": (
        "The request took too long",
        "Long requests can time out. Press Try again, or ask for something smaller "
        "(fewer slides or exercises, or one section at a time)."),
    "model": (
        "The AI model is unavailable",
        "The model the app uses is not available at the moment. Please ask the app owner to update it."),
    "blocked": (
        "The AI declined this request",
        "The answer was blocked by the AI's safety filters. Try rephrasing the request."),
    "parse": (
        "The answer came back incomplete",
        "The AI's answer could not be read. Press Try again, this usually helps."),
    "network": (
        "No connection",
        "The app could not reach the AI service. Check the internet connection and try again."),
    "other": (
        "Something went wrong",
        "An unexpected error occurred. Press Try again. If it keeps happening, send the "
        "technical details below to the app owner."),
}
PRIORITY = ["quota_day", "quota_min", "key", "busy", "timeout", "model",
            "blocked", "parse", "network", "other"]


def classify_error(e):
    msg = str(e)
    low = msg.lower()

    def has(pattern):
        return re.search(pattern, msg) is not None

    if has(r"\b429\b") or "resource_exhausted" in low or "quota" in low or "rate limit" in low:
        daily = any(x in low for x in ("perday", "per day", "per_day", "daily"))
        return "quota_day" if daily else "quota_min"
    if "api key" in low or "api_key" in low or "permission_denied" in low or has(r"\b40[13]\b"):
        return "key"
    if has(r"\b503\b") or any(x in low for x in ("unavailable", "overloaded", "high demand")):
        return "busy"
    if has(r"\b504\b") or any(x in low for x in ("deadline", "timed out", "timeout")):
        return "timeout"
    if has(r"\b404\b") or "not_found" in low or "not found" in low:
        return "model"
    if any(x in low for x in ("safety", "blocked", "prohibited", "recitation")):
        return "blocked"
    if type(e).__name__ in ("ValidationError", "JSONDecodeError") or any(
            x in low for x in ("validation error", "empty response", "expecting value", "unterminated")):
        return "parse"
    if any(x in low for x in ("connection", "getaddrinfo", "name or service", "network", "ssl")):
        return "network"
    return "other"


class AllModelsFailed(Exception):
    def __init__(self, errors):
        self.errors = errors
        super().__init__("; ".join(str(x)[:200] for x in errors))


def ask_model(contents):
    """Запит до моделі. При проблемах (ліміт, перевантаження, тайм-аут) пробує наступну модель."""
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM,
        response_mime_type="application/json",
        response_schema=Answer,
    )
    errs = []
    for model in MODELS:
        for attempt in range(2):
            try:
                resp = client.models.generate_content(
                    model=model, contents=contents, config=config)
                if not getattr(resp, "text", None):
                    raise ValueError("Empty response from the model (possibly blocked)")
                return resp
            except Exception as e:
                code = classify_error(e)
                if code == "busy" and attempt == 0:
                    time.sleep(3)          # одна повторна спроба тієї ж моделі
                    continue
                if code in ("quota_day", "quota_min", "busy", "timeout", "model", "parse"):
                    errs.append(e)         # переходимо до наступної моделі
                    break
                raise                      # ключ, мережа, фільтри тощо: далі немає сенсу
    raise AllModelsFailed(errs)


def describe_failure(e):
    """Повертає (дружній текст, технічні деталі)."""
    errs = e.errors if isinstance(e, AllModelsFailed) else [e]
    if not errs:
        errs = [e]
    code = min((classify_error(x) for x in errs), key=PRIORITY.index)
    title, hint = ERRORS[code]
    details = "\n\n".join(f"{type(x).__name__}: {str(x)[:700]}" for x in errs)
    return f"**⚠️ {title}**\n\n{hint}", details


# ---------- Історія розмов (зберігається у файлі) ----------
COLLECTION = "conversations"
RETENTION_DAYS = 30   # чати видаляються через стільки днів після останньої зміни


def _expiry():
    return datetime.now(timezone.utc) + timedelta(days=RETENTION_DAYS)

try:
    _BASE = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _BASE = os.getcwd()
DATA_FILE = os.path.join(_BASE, "chat_history.json")


def _pack(messages):
    """Стискає повідомлення в короткий текст (щоб вкластись у ліміт документа бази)."""
    raw = json.dumps(messages, ensure_ascii=False).encode("utf-8")
    return base64.b64encode(zlib.compress(raw)).decode("ascii")


def _unpack(text):
    return json.loads(zlib.decompress(base64.b64decode(text)).decode("utf-8"))


def cloud_configured():
    try:
        return bool(st.secrets.get("firebase"))
    except Exception:
        return False


@st.cache_resource
def _connect():
    """Повертає (клієнт Firestore або None, текст помилки або порожній рядок)."""
    try:
        info = st.secrets.get("firebase")
        if not info:
            return None, ""
        from google.oauth2 import service_account
        from google.cloud import firestore
        creds = service_account.Credentials.from_service_account_info(dict(info))
        return firestore.Client(credentials=creds, project=creds.project_id), ""
    except Exception as e:
        return None, f"{type(e).__name__}: {str(e)[:200]}"


def get_db():
    return _connect()[0]


def db_error():
    return _connect()[1]


def _load_file():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_file(convs):
    try:
        tmp = DATA_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(convs, f, ensure_ascii=False)
        os.replace(tmp, DATA_FILE)
    except Exception:
        pass  # файлова система може бути тільки для читання


def load_convs():
    db = get_db()
    if db is None:
        if cloud_configured():          # база налаштована, але не підключилась
            st.session_state["storage_error"] = "connect"
            return []
        return _load_file()
    try:
        out = []
        for doc in db.collection(COLLECTION).stream():
            d = doc.to_dict() or {}
            if "expireAt" not in d:      # старі чати: рахуємо 30 днів від першого відкриття
                try:
                    doc.reference.update({"expireAt": _expiry()})
                except Exception:
                    pass
            try:
                msgs = _unpack(d.get("messages", ""))
            except Exception:
                msgs = []
            out.append({"id": doc.id, "title": d.get("title", "Untitled"),
                        "created": d.get("created", ""), "updated": d.get("updated", ""),
                        "messages": msgs})
        out.sort(key=lambda c: c.get("updated") or c.get("created") or "", reverse=True)
        return out
    except Exception as e:
        st.session_state["storage_error"] = str(e)[:300]
        return []


def save_conv(c):
    """Зберігає одну розмову (хмара) або всі розмови (локальний файл)."""
    c["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    db = get_db()
    if db is None:
        if cloud_configured():
            st.session_state["storage_error"] = "connect"
            return False
        _save_file(st.session_state["convs"])
        return True
    try:
        packed = _pack(c["messages"])
        if len(packed) > 900_000:       # ліміт документа 1 МБ: прибираємо службові дані зі старих повідомлень
            slim = [dict(m) for m in c["messages"]]
            keep = [i for i, m in enumerate(slim) if m.get("raw")][-2:]
            for i, m in enumerate(slim):
                if i not in keep:
                    m.pop("raw", None)
            packed = _pack(slim)
        db.collection(COLLECTION).document(c["id"]).set({
            "title": c["title"], "created": c.get("created", ""),
            "updated": c["updated"], "messages": packed,
            "expireAt": _expiry()})
        st.session_state["storage_error"] = None
        return True
    except Exception as e:
        st.session_state["storage_error"] = str(e)[:300]
        return False


def delete_conv(cid):
    """Видаляє розмову назавжди. Викликається лише коли користувачка підтвердила видалення."""
    db = get_db()
    if db is None:
        if not cloud_configured():
            _save_file(st.session_state["convs"])
        return
    try:
        db.collection(COLLECTION).document(cid).delete()
    except Exception as e:
        st.session_state["storage_error"] = str(e)[:300]


def current_conv():
    cid = st.session_state.get("current")
    for c in st.session_state["convs"]:
        if c["id"] == cid:
            return c
    return None


def build_contents(msgs):
    """Історія для моделі: тільки успішні пари запит-відповідь."""
    out = []
    i = 0
    while i < len(msgs) - 1:
        u, a = msgs[i], msgs[i + 1]
        if u["role"] == "user" and a["role"] == "assistant" and a.get("raw"):
            out.append(types.Content(role="user", parts=[types.Part(text=u["text"])]))
            out.append(types.Content(role="model", parts=[types.Part(text=a["raw"])]))
            i += 2
        else:
            i += 1
    return out[-12:]  # останні 6 пар, щоб не роздувати запит


def files_for(docs):
    """Будує файли Word/PowerPoint із збережених даних (з кешем на сеанс)."""
    cache = st.session_state.setdefault("_filecache", {})
    out = []
    for kind, data in docs:
        key = hashlib.md5(json.dumps([kind, data], sort_keys=True,
                                     ensure_ascii=False).encode("utf-8")).hexdigest()
        if key not in cache:
            if kind == "docx":
                cache[key] = (file_name(data["title"], "docx"), build_docx(data), DOCX_MIME)
            else:
                cache[key] = (file_name(data["title"], "pptx"), build_pptx(data), PPTX_MIME)
        out.append(cache[key])
    return out


# ---------- Категорії: форми з полями ----------
LEVEL_OPTS = ["A1", "A2", "B1", "B2", "C1", "C2"]
EX_TYPES = ["Gap-fill", "Multiple choice", "Matching", "Error correction",
            "True / False", "Sentence transformation", "Short writing tasks"]
LECTURE_PARTS = ["Learning objectives", "Theory", "Examples", "Discussion questions",
                 "Practice tasks", "Homework", "Glossary of key terms"]


def _extra(text):
    return f"Extra instructions: {text.strip()}\n" if text and text.strip() else ""


def read_upload(f):
    if f is None:
        return ""
    try:
        data = f.getvalue()
        if f.name.lower().endswith(".docx"):
            d = Document(io.BytesIO(data))
            return "\n".join(p.text for p in d.paragraphs if p.text.strip())[:40000]
        return data.decode("utf-8", errors="ignore")[:40000]
    except Exception:
        return ""


def form_exercises():
    with st.form("form_exercises"):
        topic = st.text_input("Topic", placeholder="e.g. Present Perfect vs Past Simple")
        c1, c2 = st.columns(2)
        level = c1.selectbox("Student level (CEFR)", LEVEL_OPTS, index=2)
        n = c2.slider("Number of exercises", 5, 30, 10)
        kinds = st.multiselect("Exercise types", EX_TYPES, default=["Gap-fill", "Multiple choice"])
        key = st.checkbox("Include answer key", value=True)
        extra = st.text_area("Extra instructions (optional)", height=80,
                             placeholder="e.g. use everyday situations, add 2 short writing tasks")
        go = st.form_submit_button("Generate ✨")
    if not go:
        return None
    if not topic.strip():
        st.warning("Please enter a topic.")
        return None
    prompt = (f"Create a Word document with {n} exercises on: {topic.strip()}.\n"
              f"Student level: CEFR {level}.\n"
              f"Exercise types: {', '.join(kinds) if kinds else 'a mix of types'}.\n"
              + ("Add a separate answer key at the end.\n" if key else "Do not include an answer key.\n")
              + _extra(extra))
    return prompt, f"📝 {topic.strip()}"


def form_presentation():
    with st.form("form_presentation"):
        topic = st.text_input("Topic", placeholder="e.g. Phonological processes: assimilation, elision, epenthesis")
        c1, c2 = st.columns(2)
        level = c1.selectbox("Audience level (CEFR)", LEVEL_OPTS, index=4)
        n = c2.slider("Number of slides", 4, 20, 8)
        c3, c4 = st.columns(2)
        style_label = c3.selectbox("Design style", list(STYLE_OPTIONS.keys()))
        theme_label = c4.selectbox("Colour theme", ["Auto"] + [k for k in THEME_OPTIONS if k != "Auto"])
        notes = st.checkbox("Add speaker notes", value=True)
        extra = st.text_area("Extra instructions (optional)", height=80,
                             placeholder="e.g. include a two-column comparison and a practice slide")
        go = st.form_submit_button("Generate ✨")
    if not go:
        return None
    if not topic.strip():
        st.warning("Please enter a topic.")
        return None
    theme = THEME_OPTIONS[theme_label]
    prompt = (f"Make a presentation on: {topic.strip()}.\n"
              f"Audience level: CEFR {level}.\n"
              f"Number of slides: {n} (not counting the title slide).\n"
              + (f"Use theme: {theme}.\n" if theme else "")
              + (f"Use design style: {STYLE_OPTIONS[style_label]}.\n" if STYLE_OPTIONS[style_label] else "")
              + ("Add speaker notes to the slides.\n" if notes else "Leave the speaker notes empty.\n")
              + _extra(extra))
    return prompt, f"📊 {topic.strip()}"


def form_lecture():
    with st.form("form_lecture"):
        topic = st.text_input("Topic", placeholder="e.g. Morphology: word formation processes")
        c1, c2 = st.columns(2)
        level = c1.selectbox("Audience level (CEFR)", LEVEL_OPTS, index=4)
        minutes = c2.slider("Duration (minutes)", 20, 120, 80, step=10)
        parts = st.multiselect("Include", LECTURE_PARTS,
                               default=["Learning objectives", "Theory", "Examples",
                                        "Discussion questions", "Homework"])
        extra = st.text_area("Extra instructions (optional)", height=80)
        go = st.form_submit_button("Generate ✨")
    if not go:
        return None
    if not topic.strip():
        st.warning("Please enter a topic.")
        return None
    prompt = (f"Prepare a Word document with a {minutes}-minute lecture on: {topic.strip()}.\n"
              f"Audience level: CEFR {level}.\n"
              f"Include: {', '.join(parts) if parts else 'a sensible lecture structure'}.\n"
              + _extra(extra))
    return prompt, f"🎓 {topic.strip()}"


def form_article():
    with st.form("form_article"):
        topic = st.text_input("Topic or working title",
                              placeholder="e.g. Code-switching in Ukrainian-English bilingual speech")
        c1, c2 = st.columns(2)
        kind = c1.selectbox("Type", ["Academic article", "Literature review",
                                     "Abstract and keywords only", "Outline only"])
        words = c2.number_input("Length (words)", min_value=200, max_value=5000, value=1500, step=100)
        c3, c4 = st.columns(2)
        lang = c3.selectbox("Language", ["English", "Ukrainian"])
        abstract = c4.checkbox("Include Abstract and Keywords", value=True)
        extra = st.text_area("Extra instructions (optional)", height=80,
                             placeholder="e.g. focus on classroom implications, formal register")
        go = st.form_submit_button("Generate ✨")
    if not go:
        return None
    if not topic.strip():
        st.warning("Please enter a topic.")
        return None
    prompt = (f"Write a Word document: {kind.lower()} on: {topic.strip()}.\n"
              f"Language: {lang}. Length: about {int(words)} words.\n"
              + ("Include an Abstract and Keywords.\n" if abstract else "")
              + "Do not invent sources or citations; where a reference is needed, write [citation needed].\n"
              + _extra(extra))
    return prompt, f"✍️ {topic.strip()}"


def form_translation():
    with st.form("form_translation"):
        c1, c2 = st.columns(2)
        direction = c1.selectbox("Direction", ["English → Ukrainian", "Ukrainian → English"])
        domain = c2.selectbox("Style", ["Linguistics / academic", "General", "Teaching materials"])
        text = st.text_area("Text to translate", height=160, placeholder="Paste your text here")
        upload = st.file_uploader("...or upload a file (.docx, .txt)", type=["docx", "txt", "md"])
        terms = st.checkbox("List key terms with the chosen translations", value=True)
        output = st.selectbox("Result", ["Word document", "Show in chat only"])
        go = st.form_submit_button("Translate ✨")
    if not go:
        return None
    body = (text or "").strip() or read_upload(upload).strip()
    if not body:
        st.warning("Please paste a text or upload a file.")
        return None
    src, dst = direction.split(" → ")
    prompt = (f"Translate the following text from {src} to {dst}. Style: {domain}.\n"
              + ("Keep terminology consistent and add a short list of key terms with the chosen translations.\n" if terms else "")
              + ("Put the translation in a Word document.\n" if output == "Word document"
                 else "Do not create a document: put the full translation in the reply.\n")
              + "\nTEXT:\n" + body)
    return prompt, f"🌐 {body[:40].strip()}"


def form_edit():
    with st.form("form_edit"):
        text = st.text_area("Your text", height=180, placeholder="Paste the text you want edited")
        upload = st.file_uploader("...or upload a file (.docx, .txt)", type=["docx", "txt", "md"])
        goal = st.selectbox("What should be improved?", [
            "Academic style and grammar", "Grammar and spelling only",
            "Make it shorter and more concise", "Clarity and flow"])
        go = st.form_submit_button("Edit ✨")
    if not go:
        return None
    body = (text or "").strip() or read_upload(upload).strip()
    if not body:
        st.warning("Please paste a text or upload a file.")
        return None
    prompt = (f"Edit the following text. Goal: {goal}.\n"
              "Keep my meaning and my voice, and keep the original language.\n"
              "Return the edited text in a Word document, followed by a section "
              "'Summary of changes' that lists the main edits and why you made them.\n"
              "\nTEXT:\n" + body)
    return prompt, f"✏️ {body[:40].strip()}"


# (ключ, емодзі, назва, опис, функція форми)
CATS = [
    ("exercises", "📝", "Exercises", "Gap-fill, multiple choice, answer keys", form_exercises),
    ("presentation", "📊", "Presentation", "Styled slides with speaker notes", form_presentation),
    ("lecture", "🎓", "Lecture", "Objectives, theory, discussion questions", form_lecture),
    ("article", "✍️", "Article", "Academic article, review or outline", form_article),
    ("translation", "🌐", "Translation", "EN ↔ UA with terminology", form_translation),
    ("edit", "✏️", "Edit text", "Proofread and polish your own text", form_edit),
]


st.markdown(CSS, unsafe_allow_html=True)

USER_AV = "🧑‍🏫"
BOT_AV = "🎓"

# ---------- Вхід за паролем ----------
LOGIN_HTML = """
<style>
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display: none; }
.block-container { padding-top: 16vh; }
[data-testid="stForm"] { background: #fff; border: 1px solid var(--line); border-radius: 20px; padding: 1.4rem; }
.login-head { text-align: center; margin-bottom: 1.3rem; }
.login-logo { width: 48px; height: 48px; border-radius: 14px; background: var(--navy); display: inline-flex;
              align-items: center; justify-content: center; margin-bottom: .9rem; }
.login-logo::after { content: ""; width: 15px; height: 15px; border-radius: 50%; background: var(--lime); }
.login-title { font-family: 'Sora', sans-serif; font-weight: 700; font-size: 1.5rem; letter-spacing: -.02em; color: var(--ink); }
.login-sub { color: var(--muted); font-size: .92rem; margin-top: .3rem; }
</style>
<div class="login-head"><div class="login-logo"></div>
<div class="login-title">English Teacher's Assistant</div>
<div class="login-sub">Enter the password to continue</div></div>
"""

pwd = st.secrets.get("APP_PASSWORD")
if pwd and not st.session_state.get("auth"):
    _l, _mid, _r = st.columns([1, 1.5, 1])
    with _mid:
        st.markdown(LOGIN_HTML, unsafe_allow_html=True)
        with st.form("login"):
            entered = st.text_input("Password", type="password", placeholder="Enter password",
                                    label_visibility="collapsed")
            ok = st.form_submit_button("Log in")
    if ok:
        if entered == pwd:
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("Incorrect password")
    st.stop()

if "convs" not in st.session_state:
    st.session_state["convs"] = load_convs()
st.session_state.setdefault("current", None)
st.session_state.setdefault("mode", None)

# ---------- Бічна панель (верхня частина) ----------
with st.sidebar:
    st.markdown('<div class="side-title"><span class="dot"></span>English</div>'
                '<div class="side-sub">Teacher\'s assistant</div>', unsafe_allow_html=True)
    if st.button(":material/add: New chat", key="new_chat"):
        st.session_state["current"] = None
        st.session_state["mode"] = None
        st.rerun()
    st.markdown("**Create**")
    for key, emoji, title, desc, _fn in CATS:
        if st.button(f"{emoji} {title}", key=f"cat_{key}", help=desc):
            st.session_state["mode"] = key
            st.rerun()


# ---------- Показ повідомлень ----------
def render_body(m, cid, i, is_last):
    st.markdown(m["text"])
    if m.get("details"):
        with st.expander("Technical details"):
            st.code(m["details"])
    docs = m.get("docs", [])
    if docs:
        render_previews(docs, f"pv_{cid}_{i}", expanded=is_last)
        render_files(files_for(docs), f"dl_{cid}_{i}")
    if m.get("error") and is_last:
        if st.button("🔁 Try again", key=f"retry_{cid}_{i}"):
            c = current_conv()
            if c and i >= 1:
                again = c["messages"][i - 1]["text"]
                del c["messages"][i - 1:]
                save_conv(c)
                st.session_state["queued"] = {"prompt": again}
                st.rerun()


def render_message(m, cid, i, is_last):
    with st.chat_message(m["role"], avatar=USER_AV if m["role"] == "user" else BOT_AV):
        render_body(m, cid, i, is_last)


def render_panel(mode):
    cat = next(c for c in CATS if c[0] == mode)
    with st.container(border=True):
        st.markdown(f'<span class="cat cat-{mode}"></span>', unsafe_allow_html=True)
        st.markdown(f"#### {cat[1]} {cat[2]}")
        st.caption(cat[3])
        result = cat[4]()
        if result:
            prompt, title = result
            st.session_state["queued"] = {"prompt": prompt, "title": title}
            st.session_state["mode"] = None
        if st.button("✕ Close", key="close_panel"):
            st.session_state["mode"] = None
            st.rerun()


# ---------- Головна частина ----------
st.markdown(HERO, unsafe_allow_html=True)

conv = current_conv()
msgs = conv["messages"] if conv else []
cid = conv["id"] if conv else "new"

if not msgs:   # стартовий екран завжди відкривається з верху сторінки
    try:
        components.html(
            "<script>const d=window.parent.document;"
            "const m=d.querySelector('[data-testid=\"stMain\"]')||d.querySelector('section.main');"
            "if(m){m.scrollTo(0,0);}window.parent.scrollTo(0,0);</script>", height=0)
    except Exception:
        pass

for i, m in enumerate(msgs):
    render_message(m, cid, i, i == len(msgs) - 1)

# стартовий екран із категоріями
welcome = st.empty()
if not msgs and not st.session_state["mode"]:
    with welcome.container():
        st.markdown("##### What would you like to create?")
        st.caption("Pick a category, fill in a few details and generate. "
                   "Or just type your own request below.")
        cols = st.columns(3)
        for idx, (key, emoji, title, desc, _fn) in enumerate(CATS):
            with cols[idx % 3]:
                with st.container(border=True):
                    st.markdown(f'<span class="cat cat-{key}"></span><div class="card-ico">{emoji}</div><div class="card-title">{title}</div>'
                                f'<div class="card-desc">{desc}</div>', unsafe_allow_html=True)
                    if st.button("Open", key=f"card_{key}"):
                        st.session_state["mode"] = key
                        st.rerun()

# панель вибраної категорії (форма з полями)
panel = st.empty()
if st.session_state["mode"]:
    with panel.container():
        render_panel(st.session_state["mode"])

prompt = st.chat_input("Describe what to create…")
queued = st.session_state.pop("queued", None)
title_hint = None
if queued:
    prompt = queued["prompt"]
    title_hint = queued.get("title")

if prompt:
    welcome.empty()
    panel.empty()

    # нова розмова створюється при першому повідомленні
    if conv is None:
        conv = {"id": uuid.uuid4().hex[:8],
                "title": (title_hint or prompt).strip().replace("\n", " ")[:60],
                "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "messages": []}
        st.session_state["convs"].insert(0, conv)
        st.session_state["current"] = conv["id"]
    msgs = conv["messages"]
    cid = conv["id"]

    contents = build_contents(msgs) + [
        types.Content(role="user", parts=[types.Part(text=prompt)])]

    user_msg = {"role": "user", "text": prompt}
    render_message(user_msg, cid, len(msgs), False)
    msgs.append(user_msg)

    with st.chat_message("assistant", avatar=BOT_AV):
        with st.spinner("Preparing your materials…"):
            try:
                resp = ask_model(contents)
                ans = Answer.model_validate_json(resp.text)
                docs = []
                if ans.docx:
                    docs.append(["docx", ans.docx.model_dump()])
                if ans.pptx:
                    docs.append(["pptx", ans.pptx.model_dump()])
                new_msg = {"role": "assistant", "text": ans.reply,
                           "raw": resp.text, "docs": docs}
            except Exception as e:
                friendly, details = describe_failure(e)
                new_msg = {"role": "assistant", "text": friendly,
                           "error": True, "details": details}
        render_body(new_msg, cid, len(msgs), True)

    msgs.append(new_msg)
    save_conv(conv)
    convs_ = st.session_state["convs"]          # свіжа розмова піднімається вгору списку
    if conv in convs_:
        convs_.remove(conv)
        convs_.insert(0, conv)

# ---------- Бічна панель (історія) ----------
with st.sidebar:
    st.markdown("**History**")
    convs = st.session_state["convs"]
    if not convs:
        st.caption("Your conversations will appear here.")
    for c in list(convs):
        if st.session_state.get("confirm_delete") == c["id"]:
            st.warning(f"Delete “{c['title'][:30]}”? This cannot be undone.")
            b1, b2 = st.columns(2)
            if b1.button("Yes, delete", key=f"yes_{c['id']}"):
                st.session_state["convs"] = [x for x in st.session_state["convs"] if x["id"] != c["id"]]
                if st.session_state.get("current") == c["id"]:
                    st.session_state["current"] = None
                st.session_state["confirm_delete"] = None
                delete_conv(c["id"])
                st.rerun()
            if b2.button("Cancel", key=f"no_{c['id']}"):
                st.session_state["confirm_delete"] = None
                st.rerun()
            continue
        cols = st.columns([6, 1.3])
        is_cur = c["id"] == st.session_state.get("current")
        label = c["title"]
        if cols[0].button(label[:30], key=f"{'hc' if is_cur else 'h'}_{c['id']}", help=f"{c['title']} ({c.get('created', '')})"):
            st.session_state["current"] = c["id"]
            st.session_state["mode"] = None
            st.rerun()
        if cols[1].button(":material/delete:", key=f"d_{c['id']}", help="Delete this conversation"):
            st.session_state["confirm_delete"] = c["id"]
            st.rerun()

    if get_db() is not None:
        st.caption(f"Saved in the cloud. Chats are deleted automatically {RETENTION_DAYS} days "
                   "after the last change. Download a backup to keep them longer.")
    elif cloud_configured():
        st.warning("Cloud storage is set up but could not connect. Your chats are not deleted; "
                   "ask the app owner to check the Firebase key.")
        if db_error():
            with st.expander("Technical details"):
                st.code(db_error())
    else:
        st.caption("⚠️ Saved on this server only. Chats may be lost when the app restarts.")
    if st.session_state.get("storage_error") and get_db() is not None:
        st.warning("The last save did not go through. Please try again in a moment.")
    if convs:
        st.download_button("Download all chats (backup)",
                           json.dumps(st.session_state["convs"], ensure_ascii=False).encode("utf-8"),
                           file_name="chats_backup.json", mime="application/json", key="backup")

    with st.expander("Tips"):
        st.markdown(
            "- Pick a category and fill in the details. Nothing is generated until you press the button.\n"
            "- After a file is ready you can ask in the chat: *add 5 more exercises*, *shorten the slides*.\n"
            "- Check examples, answer keys and terms before using them in class.")
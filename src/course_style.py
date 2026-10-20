# -*- coding: utf-8 -*-
"""Общий стиль PDF-уроков курса «ИИ без воды»: шрифты, палитра, плашки, таблицы, схемы, обложка."""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                Table, TableStyle, KeepTogether, PageBreak, Preformatted,
                                NextPageTemplate, CondPageBreak)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Circle, Polygon, Group
import math

FD = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DV", FD + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DV-B", FD + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("DV-I", FD + "DejaVuSans-Oblique.ttf"))
pdfmetrics.registerFont(TTFont("DV-M", FD + "DejaVuSansMono.ttf"))
addMapping("DV", 0, 0, "DV"); addMapping("DV", 1, 0, "DV-B")
addMapping("DV", 0, 1, "DV-I"); addMapping("DV", 1, 1, "DV-B")

NAVY = HexColor("#1F3A5F"); TEAL = HexColor("#2A9D8F"); GOLD = HexColor("#E9C46A")
CORAL = HexColor("#E76F51"); PURPLE = HexColor("#6A4C93"); BROWN = HexColor("#8B6B1A")
DTEAL = HexColor("#0F766E"); INK = HexColor("#2B2D33"); GREY = HexColor("#6B7280")
LGREY = HexColor("#F1F3F5"); LINE = HexColor("#D1D5DB")

# цвета шести блоков (единые на всех схемах)
B_ROLE, B_CTX, B_TASK, B_FMT, B_LIM, B_EX = PURPLE, TEAL, NAVY, HexColor("#C9971C"), CORAL, BROWN

W, H = A4
M = 52
CW = W - 2 * M

# ---------- стили ----------
body = ParagraphStyle("body", fontName="DV", fontSize=9.6, leading=14.6, textColor=INK, spaceAfter=6)
h1 = ParagraphStyle("h1", fontName="DV-B", fontSize=17, leading=21, textColor=NAVY, spaceBefore=12, spaceAfter=8)
h2 = ParagraphStyle("h2", fontName="DV-B", fontSize=12.2, leading=16, textColor=NAVY, spaceBefore=8, spaceAfter=4)
bul = ParagraphStyle("bul", parent=body, leftIndent=14, bulletIndent=3, spaceAfter=2.5)
cap = ParagraphStyle("cap", fontName="DV-I", fontSize=7.8, leading=10, textColor=GREY, alignment=TA_CENTER, spaceBefore=3, spaceAfter=10)
cbody = ParagraphStyle("cbody", parent=body, fontSize=9, leading=13.4, spaceAfter=3)
cbul = ParagraphStyle("cbul", parent=cbody, leftIndent=12, bulletIndent=2, spaceAfter=1.5)
chead = ParagraphStyle("chead", fontName="DV-B", fontSize=8.6, leading=11, textColor=white)
tcell = ParagraphStyle("tcell", fontName="DV", fontSize=8.4, leading=11.4, textColor=INK)
thead = ParagraphStyle("thead", fontName="DV-B", fontSize=8.4, leading=11, textColor=white)
small = ParagraphStyle("small", fontName="DV", fontSize=7.6, leading=10, textColor=GREY)
mono = ParagraphStyle("mono", fontName="DV-M", fontSize=7.9, leading=11.2, textColor=INK)

P = lambda t, s=body: Paragraph(t, s)
def bullets(items, s=bul):
    return [Paragraph(i, s, bulletText="•") for i in items]
def numbered(items, s=bul):
    return [Paragraph(i, s, bulletText=f"{n}.") for n, i in enumerate(items, 1)]

CALL = {
    "Главное": (NAVY, HexColor("#E8EEF6")),
    "Пример": (TEAL, HexColor("#E6F4F1")),
    "Аналогия": (BROWN, HexColor("#FBF4E0")),
    "Осторожно": (CORAL, HexColor("#FCEAE4")),
    "Быстро устаревает — проверяйте актуальное": (PURPLE, HexColor("#EFE9F6")),
    "Практика": (DTEAL, HexColor("#E6F3F1")),
    "Ответы": (NAVY, HexColor("#E8EEF6")),
}
def callout(kind, content, title=None):
    hc, bc = CALL[kind]
    if isinstance(content, str):
        content = [P(content, cbody)]
    t = Table([[P(title or kind, chead)], [content]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hc), ("BACKGROUND", (0, 1), (-1, 1), bc),
        ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, 0), 4), ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
        ("TOPPADDING", (0, 1), (-1, 1), 6), ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
    ]))
    return KeepTogether([Spacer(1, 4), t, Spacer(1, 8)])

def code(text, width=None):
    pre = Preformatted(text, mono)
    t = Table([[pre]], colWidths=[width or CW])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LGREY), ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    if width:
        return t
    return KeepTogether([Spacer(1, 2), t, Spacer(1, 8)])

def table(rows, widths, head_color=NAVY):
    data = [[P(c, thead) for c in rows[0]]] + [[P(c, tcell) for c in r] for r in rows[1:]]
    t = Table(data, colWidths=[CW * w for w in widths], repeatRows=1)
    st = [("BACKGROUND", (0, 0), (-1, 0), head_color),
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("LINEBELOW", (0, 1), (-1, -1), 0.4, LINE),
          ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
          ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]
    for i in range(1, len(data)):
        if i % 2 == 0:
            st.append(("BACKGROUND", (0, i), (-1, i), HexColor("#F7F8FA")))
    t.setStyle(TableStyle(st))
    return KeepTogether([Spacer(1, 2), t, Spacer(1, 9)])

def figure(drawing, caption):
    return KeepTogether([Spacer(1, 4), drawing, P(caption, cap)])

# ---------- примитивы для схем ----------
def rbox(d, x, y, w, h, fill, lines, fs=8, fc=white, stroke=None, r=6, bold_first=True, lead=None):
    d.add(Rect(x, y, w, h, rx=r, ry=r, fillColor=fill, strokeColor=stroke or fill, strokeWidth=0.8))
    lead = lead or fs * 1.3
    n = len(lines)
    top = y + h / 2 + (n - 1) * lead / 2 - fs * 0.35
    for i, ln in enumerate(lines):
        font = "DV-B" if (bold_first and i == 0) else "DV"
        d.add(String(x + w / 2, top - i * lead, ln, fontName=font, fontSize=fs, fillColor=fc, textAnchor="middle"))

def arrow(d, x1, y1, x2, y2, c=GREY, w=1.2, head=6):
    d.add(Line(x1, y1, x2, y2, strokeColor=c, strokeWidth=w))
    a = math.atan2(y2 - y1, x2 - x1)
    p1 = (x2 - head * math.cos(a - 0.4), y2 - head * math.sin(a - 0.4))
    p2 = (x2 - head * math.cos(a + 0.4), y2 - head * math.sin(a + 0.4))
    d.add(Polygon([x2, y2, p1[0], p1[1], p2[0], p2[1]], fillColor=c, strokeColor=c, strokeWidth=0.5))

def txt(d, x, y, s, fs=7.5, c=INK, font="DV", anchor="start"):
    d.add(String(x, y, s, fontName=font, fontSize=fs, fillColor=c, textAnchor=anchor))


# ---------- плашка «Частое сомнение» ----------
DOUBT_HEAD, DOUBT_BG = HexColor("#475569"), HexColor("#EEF2F6")
CALL["Частое сомнение"] = (DOUBT_HEAD, DOUBT_BG)
dq = ParagraphStyle("dq", parent=cbody, fontName="DV-B", fontSize=9.6, leading=13.6, textColor=HexColor("#334155"), spaceAfter=4)
dsrc = "<font color='#6B7280' size='7.6'><i>— {}</i></font>"

def doubt(question, answer, facts=(), fair=None, phrase=None):
    """Плашка «Частое сомнение»: сомнение → ответ → факты с источниками → что справедливо → фраза для обсуждения."""
    c = [Paragraph(f"«{question}»", dq), Paragraph(answer, cbody)]
    for text, source in facts:
        c.append(Paragraph(f"{text} " + dsrc.format(source), cbul, bulletText="•"))
    if fair:
        c.append(Paragraph(f"<font color='#C2410C'><b>Что здесь справедливо.</b></font> {fair}", cbody))
    if phrase:
        c.append(Paragraph(f"<b>В разговоре с коллегой:</b> <i>«{phrase}»</i>", cbody))
    return callout("Частое сомнение", c)

srcst = ParagraphStyle("srcst", fontName="DV", fontSize=7.2, leading=9.6, textColor=GREY, leftIndent=10, bulletIndent=0, spaceAfter=1.5)
def sources_block(items):
    """Источники фактов из плашек «Частое сомнение» — в конце урока."""
    from reportlab.platypus import KeepTogether as KT
    return KT([Paragraph("Источники фактов", h2)] + [Paragraph(i, srcst, bulletText="·") for i in items])

# ---------- оформление страниц и сборка ----------
def _network(c, ox, oy, layers=(3, 5, 5, 2), lx=46, ly=34):
    pts = []
    for li, n in enumerate(layers):
        pts.append([(ox + li * lx, oy + (max(layers) - 1) * ly / 2 - (k - (n - 1) / 2) * ly) for k in range(n)])
    c.setStrokeColor(HexColor("#5C7BA3")); c.setLineWidth(0.5)
    for a, b in zip(pts, pts[1:]):
        for p in a:
            for q in b:
                c.line(p[0], p[1], q[0], q[1])
    for li, col in enumerate(pts):
        for p in col:
            c.setFillColor(GOLD if li in (0, len(layers) - 1) else TEAL)
            c.circle(p[0], p[1], 7, stroke=0, fill=1)

def build(path, title, kicker, title_lines, subtitle_lines, header, story,
          version="октябрь 2026"):
    """Собрать урок: обложка + страницы с колонтитулом. story начинается с первой страницы."""
    def cover(c, doc):
        c.saveState()
        top_h = 300
        c.setFillColor(NAVY); c.rect(0, H - top_h, W, top_h, stroke=0, fill=1)
        c.setFillColor(TEAL); c.rect(0, H - top_h - 6, W, 6, stroke=0, fill=1)
        _network(c, 375, H - 225)
        c.setFillColor(HexColor("#A9C1E0")); c.setFont("DV-B", 8.5)
        c.drawString(M, H - 78, kicker)
        c.setFillColor(white); c.setFont("DV-B", 25)
        for i, ln in enumerate(title_lines):
            c.drawString(M, H - 130 - i * 31, ln)
        c.setFillColor(HexColor("#D6E2F0")); c.setFont("DV", 10)
        y0 = H - 130 - (len(title_lines) - 1) * 31 - 45
        for i, ln in enumerate(subtitle_lines):
            c.drawString(M, y0 - i * 15, ln)
        c.setFillColor(GREY); c.setFont("DV", 7.2)
        c.drawString(M, 30, f"Версия: {version} · для самостоятельного повторения")
        c.restoreState()

    def later(c, doc):
        c.saveState()
        c.setFont("DV", 7.2); c.setFillColor(GREY)
        c.drawString(M, H - 34, header)
        c.setStrokeColor(LINE); c.setLineWidth(0.5); c.line(M, H - 40, W - M, H - 40)
        c.drawRightString(W - M, 28, str(doc.page))
        c.restoreState()

    doc = BaseDocTemplate(path, pagesize=A4, title=title, author="Курс «ИИ без воды»",
                          leftMargin=M, rightMargin=M, topMargin=58, bottomMargin=50)
    doc.addPageTemplates([
        PageTemplate("cover", [Frame(M, 50, CW, H - 306 - 50 - 20, id="c")], onPage=cover),
        PageTemplate("body", [Frame(M, 46, CW, H - 58 - 46, id="b")], onPage=later)])
    doc.build([NextPageTemplate("body"), Spacer(1, 6)] + story)

h1c = ParagraphStyle("h1c", parent=h1, spaceBefore=0)

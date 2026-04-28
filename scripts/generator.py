"""
مذكرتي Pro — محرك التحويل البصري v3.0
Visual Rendering Engine — transforms raw content to world-class slides
نفس المحتوى · تجربة بصرية مختلفة تماماً
"""
import json, sys, datetime, math
from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

W, H = 33.867, 19.05

def cm(v): return Cm(v)
def r(r,g,b): return RGBColor(r,g,b)
def safe(v, fb=""): return str(v).strip() if v else fb

# ══════════════════════════════════════════════════════════════
# COLOUR PALETTES
# ══════════════════════════════════════════════════════════════
THEMES = {
    "navy_gold": dict(
        D=r(0x07,0x17,0x2F), M=r(0x0E,0x27,0x4D), L=r(0xF4,0xF6,0xFB),
        A=r(0xC6,0xA0,0x3C), A2=r(0xE8,0xC9,0x7B),
        TL=r(0xFF,0xFF,0xFF), TD=r(0x0E,0x27,0x4D), TM=r(0x64,0x74,0x8B),
        CB=r(0xFF,0xFF,0xFF), CE=r(0xE2,0xE8,0xF0),
        HF="Georgia", BF="Cairo",
        SC=[r(0xC6,0xA0,0x3C),r(0x0E,0x27,0x4D),r(0x1A,0x40,0x72),
            r(0xE8,0xC9,0x7B),r(0x07,0x17,0x2F),r(0x2A,0x55,0x98)],
    ),
    "dark_teal": dict(
        D=r(0x06,0x1A,0x28), M=r(0x04,0x4E,0x6E), L=r(0xEF,0xFD,0xFA),
        A=r(0x00,0xD4,0xAA), A2=r(0x67,0xE8,0xD3),
        TL=r(0xFF,0xFF,0xFF), TD=r(0x06,0x1A,0x28), TM=r(0x52,0x73,0x84),
        CB=r(0xFF,0xFF,0xFF), CE=r(0xCC,0xF5,0xED),
        HF="Trebuchet MS", BF="Cairo",
        SC=[r(0x00,0xD4,0xAA),r(0x04,0x4E,0x6E),r(0x06,0x7A,0x9F),
            r(0x67,0xE8,0xD3),r(0x03,0x30,0x44),r(0x09,0x93,0xC3)],
    ),
    "burgundy": dict(
        D=r(0x3A,0x00,0x18), M=r(0x6B,0x15,0x37), L=r(0xFD,0xF0,0xF4),
        A=r(0xF0,0xB8,0xCC), A2=r(0xFA,0xD4,0xE2),
        TL=r(0xFF,0xFF,0xFF), TD=r(0x3A,0x00,0x18), TM=r(0x78,0x55,0x63),
        CB=r(0xFF,0xFF,0xFF), CE=r(0xF5,0xD8,0xE4),
        HF="Georgia", BF="Cairo",
        SC=[r(0xF0,0xB8,0xCC),r(0x6B,0x15,0x37),r(0x9A,0x20,0x50),
            r(0xFA,0xD4,0xE2),r(0x3A,0x00,0x18),r(0xCC,0x3D,0x73)],
    ),
    "forest": dict(
        D=r(0x0F,0x2D,0x1E), M=r(0x1E,0x4D,0x36), L=r(0xF0,0xFB,0xF4),
        A=r(0x86,0xBB,0x56), A2=r(0xB4,0xD9,0x84),
        TL=r(0xFF,0xFF,0xFF), TD=r(0x0F,0x2D,0x1E), TM=r(0x4A,0x6B,0x56),
        CB=r(0xFF,0xFF,0xFF), CE=r(0xD1,0xF0,0xDA),
        HF="Cambria", BF="Cairo",
        SC=[r(0x86,0xBB,0x56),r(0x1E,0x4D,0x36),r(0x2E,0x7A,0x56),
            r(0xB4,0xD9,0x84),r(0x0F,0x2D,0x1E),r(0x4A,0x99,0x6B)],
    ),
}

# ══════════════════════════════════════════════════════════════
# PRIMITIVES
# ══════════════════════════════════════════════════════════════
def blank(prs): return prs.slides.add_slide(prs.slide_layouts[6])

def bg(s, c):
    f = s.background.fill; f.solid(); f.fore_color.rgb = c

def rect(slide, x, y, w, h, fill, line=None, alpha=0):
    sp = slide.shapes.add_shape(1, cm(x), cm(y), cm(w), cm(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if alpha: sp.fill.fore_color.transparency = alpha
    if line: sp.line.color.rgb = line; sp.line.width = Pt(0.5)
    else: sp.line.fill.background()
    return sp

def oval(slide, x, y, w, h, fill, alpha=0):
    sp = slide.shapes.add_shape(9, cm(x), cm(y), cm(w), cm(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if alpha: sp.fill.fore_color.transparency = alpha
    sp.line.fill.background()
    return sp

def txt(slide, text, x, y, w, h,
        font="Cairo", size=13, bold=False, italic=False,
        color=None, align=PP_ALIGN.RIGHT, mg=0.1):
    tb = slide.shapes.add_textbox(cm(x), cm(y), cm(w), cm(h))
    tb.word_wrap = True
    tf = tb.text_frame; tf.word_wrap = True
    m = cm(mg)
    tf.margin_left = m; tf.margin_right = m
    tf.margin_top = cm(0.04); tf.margin_bottom = cm(0.04)
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run()
    run.text = str(text) if text is not None else ""
    run.font.name = font; run.font.size = Pt(size)
    run.font.bold = bold; run.font.italic = italic
    if color: run.font.color.rgb = color
    return tb

def ln(slide, x, y, w, color, h=0.06):
    rect(slide, x, y, w, h, color)

# ══════════════════════════════════════════════════════════════
# LAYOUT COMPONENTS
# ══════════════════════════════════════════════════════════════
def dark_header(slide, T, title_ar, sub_en=None):
    rect(slide, 0, 0, W, 0.45, T["A"])
    txt(slide, title_ar, 1.2, 0.65, W-2.4, 1.2,
        font=T["HF"], size=30, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
    if sub_en:
        txt(slide, sub_en, 1.2, 1.88, W-2.4, 0.65,
            font="Calibri", size=13, italic=True, color=T["A"], align=PP_ALIGN.RIGHT)
    return 2.75

def light_header(slide, T, title_ar, sub_en=None):
    rect(slide, 0, 0, W, 0.45, T["D"])
    txt(slide, title_ar, 1.2, 0.65, W-2.4, 1.2,
        font=T["HF"], size=30, bold=True, color=T["TD"], align=PP_ALIGN.RIGHT)
    if sub_en:
        txt(slide, sub_en, 1.2, 1.88, W-2.4, 0.65,
            font="Calibri", size=13, italic=True, color=T["TM"], align=PP_ALIGN.RIGHT)
    ln(slide, 1.2, 2.72, W-2.4, T["A"])
    return 2.88

def kpi_card(slide, T, x, y, w, h, value, label):
    rect(slide, x, y, w, h, T["M"])
    rect(slide, x, y, w, 0.32, T["A"])
    rect(slide, x, y+h-0.32, w, 0.32, T["A"])
    # interior region between accent bars
    interior_y = y + 0.36
    interior_h = h - 0.72
    # value size scales with string length
    vs = max(30, min(58, 60 - max(0, len(str(value))-4)*5))
    # estimate rendered line height from font size
    line_h = vs * 0.0353  # pt → cm approx
    label_h = 0.7
    total_content = line_h + 0.25 + label_h
    start = interior_y + max(0, (interior_h - total_content) / 2)
    txt(slide, str(value), x+0.1, start, w-0.2, line_h + 0.3,
        font="Calibri", size=vs, bold=True, color=T["A"], align=PP_ALIGN.CENTER)
    txt(slide, str(label), x+0.1, start+line_h+0.3, w-0.2, label_h,
        font=T["BF"], size=12, color=T["TL"], align=PP_ALIGN.CENTER)

def quote_box(slide, T, x, y, w, h, text):
    rect(slide, x, y, w, h, T["M"])
    rect(slide, x, y, 0.38, h, T["A"])
    txt(slide, "\u275d", x+0.55, y+0.18, 2.0, 1.1,
        font="Georgia", size=46, color=T["A"], align=PP_ALIGN.RIGHT)
    txt(slide, text, x+0.65, y+1.15, w-1.0, h-1.3,
        font=T["BF"], size=13, color=T["TL"], align=PP_ALIGN.RIGHT)

def icon_block(slide, T, x, y, w, h, icon, label, value):
    rect(slide, x+0.08, y+0.08, w, h, T["CE"])
    rect(slide, x, y, w, h, T["CB"], line=T["CE"])
    rect(slide, x, y, w, 0.82, T["D"])
    txt(slide, f"{icon}  {label}", x+0.15, y+0.1, w-0.3, 0.62,
        font=T["BF"], size=13, bold=True, color=T["A"], align=PP_ALIGN.RIGHT)
    txt(slide, str(value), x+0.15, y+1.0, w-0.3, h-1.15,
        font=T["BF"], size=13, bold=True, color=T["TD"], align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════
# SLIDE 01 — COVER
# ══════════════════════════════════════════════════════════════
def slide_cover(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    oval(slide, W-12, -5, 16, 16, T["M"], 0.38)
    oval(slide, -3, H-8, 10, 10, T["M"], 0.55)
    rect(slide, W-0.65, 0, 0.65, H, T["A"])
    rect(slide, 0, H-0.5, W-0.65, 0.5, T["A"])

    rect(slide, 0, 0, W-0.65, 2.6, T["M"])
    txt(slide, safe(data.get("university")), 0.9, 0.28, W-2.3, 1.05,
        font=T["BF"], size=17, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
    fac = " | ".join(filter(None,[safe(data.get("faculty")),safe(data.get("department"))]))
    if fac.strip(" |"):
        txt(slide, fac, 0.9, 1.38, W-2.3, 0.82,
            font=T["BF"], size=12, color=T["A"], align=PP_ALIGN.RIGHT)

    rect(slide, 0.9, 3.1, 8.2, 0.88, T["A"])
    txt(slide, f"مذكرة تخرج — {safe(data.get('level'),'ماستر')}",
        0.9, 3.1, 8.2, 0.88, font=T["BF"], size=14, bold=True,
        color=T["D"], align=PP_ALIGN.CENTER)

    txt(slide, safe(data.get("titleAr")), 0.9, 4.22, W-2.3, 3.5,
        font=T["BF"], size=21, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)

    if data.get("titleFr"):
        txt(slide, data["titleFr"], 0.9, 7.92, W-2.3, 0.82,
            font="Calibri", size=11, italic=True, color=T["A2"], align=PP_ALIGN.LEFT)

    my = H * 0.56
    rect(slide, 0.9, my, W-2.3, 0.06, T["A"])
    for di in range(3):
        oval(slide, 0.9+di*1.3, my-0.23, 0.45, 0.45, T["A"], 0.5)

    sy = H * 0.79
    ln(slide, 0.9, sy, W-2.3, T["A"], 0.08)
    cw = (W-2.3)/2 - 0.6
    txt(slide, "إعداد الطالب", 0.9, sy+0.28, cw, 0.5,
        font=T["BF"], size=11, color=T["A"], align=PP_ALIGN.RIGHT)
    txt(slide, safe(data.get("studentName")), 0.9, sy+0.82, cw, 0.88,
        font=T["BF"], size=18, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
    rx = 0.9 + cw + 1.1
    sw = W - rx - 1.3
    txt(slide, "إشراف الأستاذ", rx, sy+0.28, sw, 0.5,
        font=T["BF"], size=11, color=T["A"], align=PP_ALIGN.RIGHT)
    txt(slide, safe(data.get("supervisor")), rx, sy+0.82, sw, 0.88,
        font=T["BF"], size=18, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)

    meta = []
    if data.get("major"):  meta.append(f"التخصص: {data['major']}")
    if data.get("year"):   meta.append(f"السنة: {data['year']}")
    if data.get("defenseDate"):
        try:
            d = datetime.date.fromisoformat(data["defenseDate"])
            mo = ["","يناير","فبراير","مارس","أبريل","مايو","يونيو",
                  "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]
            meta.append(f"تاريخ المناقشة: {d.day} {mo[d.month]} {d.year}")
        except: meta.append(f"تاريخ المناقشة: {data['defenseDate']}")
    if meta:
        txt(slide, "   ·   ".join(meta), 0.9, sy+1.92, W-2.3, 0.62,
            font=T["BF"], size=11, color=T["TM"], align=PP_ALIGN.RIGHT)
    if data.get("keywords"):
        txt(slide, f"كلمات مفتاحية: {data['keywords']}", 0.9, H-1.22, W-2.3, 0.58,
            font=T["BF"], size=10, color=T["TM"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 02 — TABLE OF CONTENTS
# ══════════════════════════════════════════════════════════════
def slide_toc(prs, data, T, chapters):
    slide = blank(prs)
    bg(slide, T["L"])
    rect(slide, 0, 0, 0.55, H, T["D"])
    rect(slide, 0.55, 0, 0.22, H, T["A"])
    rect(slide, 0.77, 0, W-0.77, 3.15, T["D"])
    txt(slide, "المحتويات", W-1.8, 0.45, W-2.6, 1.5,
        font=T["HF"], size=36, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
    txt(slide, "Table of Contents", 1.5, 2.0, W-3.0, 0.72,
        font="Calibri", size=14, italic=True, color=T["A"], align=PP_ALIGN.LEFT)

    cw = (W-3.2)/2; gap = 0.42
    ch = (H-4.1)/3 - 0.22
    for i, chi in enumerate(chapters[:6]):
        col = i%2; row = i//2
        cx = 0.97 + col*(cw+gap)
        cy = 3.48 + row*(ch+0.26)
        rect(slide, cx+0.1, cy+0.1, cw, ch, T["CE"])
        rect(slide, cx, cy, cw, ch, T["CB"], line=T["CE"])
        rect(slide, cx, cy, 0.22, ch, T["A"])
        txt(slide, f"{i+1:02d}", cx+0.32, cy+0.1, 1.65, ch*0.5,
            font="Calibri", size=32, bold=True, color=T["A"], align=PP_ALIGN.LEFT)
        txt(slide, chi["title"], cx+0.32, cy+ch*0.52, cw-0.5, ch*0.36,
            font=T["BF"], size=14, bold=True, color=T["TD"], align=PP_ALIGN.RIGHT)
        if chi.get("sub"):
            txt(slide, chi["sub"], cx+0.32, cy+ch*0.84, cw-0.5, 0.48,
                font="Calibri", size=10, italic=True, color=T["TM"], align=PP_ALIGN.LEFT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 03 — RESEARCH PROBLEM
# Quote box + numbered step sub-questions
# ══════════════════════════════════════════════════════════════
def slide_problem(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    oval(slide, W-11, -3, 14, 14, T["M"], 0.45)
    oval(slide, -3, H-8, 10, 10, T["A"], 0.88)
    dark_header(slide, T, "إشكالية البحث", "Research Problem")

    problem = safe(data.get("mainProblem"))
    subs = [s for s in data.get("subQuestions",[]) if s]
    n_s  = min(len(subs), 4)
    qh   = 4.0 if n_s else 6.5

    quote_box(slide, T, 1.2, 2.82, W-2.4, qh, problem)

    if subs:
        sy = 2.82 + qh + 0.35
        txt(slide, "التساؤلات الفرعية", 1.2, sy, W-2.4, 0.72,
            font=T["HF"], size=18, bold=True, color=T["A"], align=PP_ALIGN.RIGHT)
        avail = H - sy - 0.82 - 0.18*n_s
        rh = max(1.3, avail/n_s)
        for i, q in enumerate(subs[:4]):
            ry   = sy + 0.82 + i*(rh+0.18)
            bgc  = T["M"] if i%2==0 else T["D"]
            sc   = T["SC"][i % len(T["SC"])]
            rect(slide, 1.2, ry, W-2.4, rh, bgc)
            rect(slide, 1.2, ry, 0.25, rh, sc)
            bs = min(rh-0.25, 1.05)
            bx = W-3.65; by = ry+(rh-bs)/2
            rect(slide, bx, by, bs, bs, T["A"])
            txt(slide, str(i+1), bx, by, bs, bs,
                font="Calibri", size=22, bold=True, color=T["D"], align=PP_ALIGN.CENTER)
            txt(slide, q, 1.55, ry+0.12, bx-1.85, rh-0.24,
                font=T["BF"], size=13, color=T["TL"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 04 — OBJECTIVES & HYPOTHESES
# Left = step cards | Right = H-badge cards
# ══════════════════════════════════════════════════════════════
def slide_objectives(prs, data, T):
    slide = blank(prs)
    bg(slide, T["L"])
    cy0 = light_header(slide, T, "أهداف البحث والفرضيات", "Objectives & Hypotheses")

    objs  = [o for o in data.get("objectives",[])  if o]
    hypos = [h for h in data.get("hypotheses",[]) if h]
    n     = max(len(objs[:6]), len(hypos[:6]), 1)
    cw    = (W-3.2)/2
    avail = H - cy0 - 0.82
    ch    = max(1.3, (avail - 0.2*n) / n)

    rect(slide, 1.2, cy0, cw, 0.78, T["D"])
    txt(slide, "🎯  الأهداف", 1.2, cy0, cw, 0.78,
        font=T["BF"], size=15, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
    for i, obj in enumerate(objs[:6]):
        oy  = cy0+0.8 + i*(ch+0.2)
        sc  = T["SC"][i%len(T["SC"])]
        rect(slide, 1.3, oy+0.07, cw, ch, T["CE"])
        rect(slide, 1.2, oy, cw, ch, T["CB"], line=T["CE"])
        rect(slide, 1.2, oy, 0.22, ch, sc)
        # Step number
        rect(slide, 1.2, oy, 1.8, ch, T["D"])
        txt(slide, str(i+1), 1.2, oy+ch/2-0.65, 1.8, 1.3,
            font="Calibri", size=28, bold=True, color=T["A"], align=PP_ALIGN.CENTER)
        txt(slide, obj, 3.2, oy+0.1, cw-2.2, ch-0.2,
            font=T["BF"], size=12, color=T["TD"], align=PP_ALIGN.RIGHT)

    rx = 1.2+cw+0.8
    rect(slide, rx, cy0, cw, 0.78, T["M"])
    txt(slide, "💡  الفرضيات", rx, cy0, cw, 0.78,
        font=T["BF"], size=15, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
    for i, hy in enumerate(hypos[:6]):
        hy_y = cy0+0.8 + i*(ch+0.2)
        sc   = T["SC"][i%len(T["SC"])]
        rect(slide, rx+0.1, hy_y+0.08, cw, ch, T["CE"])
        rect(slide, rx, hy_y, cw, ch, T["CB"], line=T["CE"])
        rect(slide, rx, hy_y, 0.22, ch, sc)
        txt(slide, f"H{i+1}", rx+0.35, hy_y+ch*0.1, 1.1, ch*0.62,
            font="Calibri", size=18, bold=True, color=T["A"], align=PP_ALIGN.RIGHT)
        txt(slide, hy, rx+1.35, hy_y+0.1, cw-1.55, ch-0.2,
            font=T["BF"], size=12, color=T["TD"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 05 — IMPORTANCE
# ══════════════════════════════════════════════════════════════
def slide_importance(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    oval(slide, W-10, -2, 13, 13, T["M"], 0.5)
    dark_header(slide, T, "أهمية البحث وأسباب اختياره", "Significance & Motivation")

    panels = [
        ("importance", "الأهمية العلمية والعملية", "⭐"),
        ("reasons",    "أسباب اختيار الموضوع",     "🔍"),
    ]
    ph = (H-3.4)/2 - 0.3
    for i, (key, lbl, icon) in enumerate(panels):
        py = 2.95 + i*(ph+0.45)
        rect(slide, 1.2, py, W-2.4, ph, T["M"])
        rect(slide, 1.2, py, 0.35, ph, T["A"])
        txt(slide, f"{icon}  {lbl}", 1.7, py+0.18, W-3.5, 0.72,
            font=T["BF"], size=16, bold=True, color=T["A"], align=PP_ALIGN.RIGHT)
        ln(slide, 1.7, py+1.06, W-3.5, T["A"], 0.05)
        txt(slide, safe(data.get(key)), 1.7, py+1.2, W-3.5, ph-1.35,
            font=T["BF"], size=13, color=T["TL"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 06 — THEORETICAL FRAMEWORK
# Concept cards with coloured header band per card
# ══════════════════════════════════════════════════════════════
def slide_theory(prs, data, T, concepts):
    slide = blank(prs)
    bg(slide, T["L"])
    cy0 = light_header(slide, T, "الإطار النظري والمفاهيمي",
                        "Theoretical & Conceptual Framework")

    n = min(len(concepts), 6)
    if not n: return slide
    cols = 3 if n >= 3 else n
    rows = math.ceil(n/cols)
    gx, gy = 0.32, 0.3
    cw = (W-2.4-gx*(cols-1))/cols
    avail = H-cy0-gy*(rows-1)
    ch = avail/rows
    if rows == 1:
        ch = min(ch, 9.0)
        grid_y = cy0 + (H-cy0-ch)/2
    else:
        grid_y = cy0

    for i, c in enumerate(concepts[:6]):
        col = i%cols; row = i//cols
        cx  = 1.2 + col*(cw+gx)
        cy  = grid_y + row*(ch+gy)
        sc  = T["SC"][i%len(T["SC"])]

        rect(slide, cx+0.1, cy+0.1, cw, ch, T["CE"])
        rect(slide, cx, cy, cw, ch, T["CB"], line=T["CE"])
        rect(slide, cx, cy, cw, 0.82, sc)
        rect(slide, cx, cy, 0.2, ch, T["A"])
        txt(slide, safe(c.get("name")), cx+0.3, cy+0.1, cw-0.5, 0.62,
            font=T["BF"], size=14, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)
        txt(slide, safe(c.get("def")), cx+0.3, cy+0.98, cw-0.5, ch-1.12,
            font=T["BF"], size=12, color=T["TD"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 07 — LITERATURE REVIEW
# Structured table with colour header + alternating rows
# ══════════════════════════════════════════════════════════════
def slide_literature(prs, data, T, lits):
    slide = blank(prs)
    bg(slide, T["D"])
    dark_header(slide, T, "مراجعة الأدبيات والدراسات السابقة", "Literature Review")

    col_defs = [
        ("الباحث / المؤلف", 4.5), ("السنة", 2.0),
        ("عنوان الدراسة", 9.5),   ("أبرز النتائج", 15.7),
    ]
    xs = [1.2]
    for _, cw in col_defs[:-1]: xs.append(xs[-1]+cw+0.12)

    hy, hh = 2.8, 0.88
    for (lbl,cw),x in zip(col_defs, xs):
        rect(slide, x, hy, cw, hh, T["A"])
        txt(slide, lbl, x+0.1, hy+0.05, cw-0.2, hh-0.1,
            font=T["BF"], size=12, bold=True, color=T["D"], align=PP_ALIGN.RIGHT)

    n  = min(len(lits), 5)
    rh = max(1.5, (H-hy-hh-0.4)/max(n,1)-0.14)
    for ri, lit in enumerate(lits[:5]):
        ry  = hy+hh+0.12+ri*(rh+0.14)
        bgc = T["M"] if ri%2==0 else T["D"]
        vals = [safe(lit.get("author")), safe(lit.get("year")),
                safe(lit.get("title")), safe(lit.get("findings"))]
        for (_, cw),x,val in zip(col_defs, xs, vals):
            rect(slide, x, ry, cw, rh, bgc)
            rect(slide, x, ry, 0.06, rh, T["A"])
            txt(slide, val, x+0.14, ry+0.1, cw-0.24, rh-0.2,
                font=T["BF"], size=11, color=T["TL"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 08 — METHODOLOGY
# 4 icon-blocks + test badges
# ══════════════════════════════════════════════════════════════
def slide_methodology(prs, data, T):
    slide = blank(prs)
    bg(slide, T["L"])
    light_header(slide, T, "المنهجية والأدوات", "Methodology & Tools")

    tests     = [t for t in data.get("statisticalTests",[]) if t]
    has_tests = bool(tests)
    bh = 3.35 if has_tests else (H-3.2)/2-0.3
    bw = (W-3.0)/2-0.2

    boxes = [
        ("🔧","المنهج المتبع",   safe(data.get("methodology"))),
        ("📊","مصدر البيانات",   safe(data.get("dataSource"))),
        ("📅","الفترة الزمنية",  safe(data.get("timePeriod"))),
        ("💻","برنامج التحليل",  safe(data.get("software"))),
    ]
    for i,(icon,lbl,val) in enumerate(boxes):
        bx = 1.2 + (i%2)*(bw+0.45)
        by = 3.05 + (i//2)*(bh+0.32)
        icon_block(slide, T, bx, by, bw, bh, icon, lbl, val)

    if has_tests:
        ty    = 3.05 + 2*(bh+0.32)
        avail = H - ty - 0.08
        txt(slide, "الاختبارات الإحصائية المستخدمة", 1.2, ty, W-2.4, 0.7,
            font=T["BF"], size=14, bold=True, color=T["TD"], align=PP_ALIGN.RIGHT)
        n  = min(len(tests), 5)
        tw = (W-2.4-0.18*(n-1))/n
        th = avail - 0.8
        for i, t in enumerate(tests[:5]):
            tx = 1.2 + i*(tw+0.18)
            sc = T["SC"][i%len(T["SC"])]
            rect(slide, tx, ty+0.8, tw, th, T["D"])
            rect(slide, tx, ty+0.8, tw, 0.36, sc)
            lines = max(1, len(t)//16 + 1)
            text_h = lines * 0.6
            text_y = ty + 0.8 + 0.42 + (th - 0.42 - text_h) / 2
            txt(slide, t, tx+0.1, max(text_y, ty+0.85), tw-0.2, text_h + 0.3,
                font=T["BF"], size=11, bold=True, color=T["TL"], align=PP_ALIGN.CENTER)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 09 — KPI DASHBOARD
# ══════════════════════════════════════════════════════════════
def slide_stats(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    dark_header(slide, T, "النتائج الكمية والإحصائية",
                "Key Statistical Results — KPI Dashboard")

    stats = [s for s in data.get("stats",[]) if s.get("label") and s.get("value")]
    if not stats: return slide

    n    = min(len(stats), 8)
    cols = min(n, 4)
    rows = math.ceil(n/cols)
    gx, gy = 0.28, 0.4
    cw   = (W-2.4-gx*(cols-1))/cols
    # Cap height at 7.0 cm — never let cards stretch into empty space
    raw_ch = (H-2.9-gy*(rows-1))/rows
    ch     = min(raw_ch, 7.0)
    # Centre the grid vertically
    total_grid_h = rows*ch + (rows-1)*gy
    grid_y = 2.9 + max(0, (H-2.9-total_grid_h)/2)

    for i, s in enumerate(stats[:8]):
        col = i%cols; row = i//cols
        cx  = 1.2 + col*(cw+gx)
        cy  = grid_y + row*(ch+gy)
        kpi_card(slide, T, cx, cy, cw, ch, s["value"], s["label"])
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 10 — RESULTS
# Each result = card with coloured left stripe + ✓ badge
# ══════════════════════════════════════════════════════════════
def slide_results(prs, data, T):
    slide = blank(prs)
    bg(slide, T["L"])
    cy0 = light_header(slide, T, "نتائج البحث التفصيلية", "Research Findings")

    results = [r for r in data.get("mainResults",[]) if r]
    n   = min(len(results), 7)
    gap = 0.2
    rh  = max(1.45, (H-cy0-gap*n)/max(n,1))

    for i, res in enumerate(results[:7]):
        ry = cy0 + i*(rh+gap)
        sc = T["SC"][i%len(T["SC"])]
        rect(slide, 1.3, ry+0.07, W-2.4, rh, T["CE"])
        rect(slide, 1.2, ry, W-2.4, rh, T["CB"], line=T["CE"])
        rect(slide, 1.2, ry, 0.28, rh, sc)
        bs = min(rh-0.22, 1.0)
        bx = W-2.75; by = ry+(rh-bs)/2
        rect(slide, bx, by, bs, bs, T["D"])
        txt(slide, "✓", bx, by, bs, bs,
            font="Calibri", size=16, bold=True, color=T["A"], align=PP_ALIGN.CENTER)
        txt(slide, res, 1.65, ry+0.1, bx-1.92, rh-0.2,
            font=T["BF"], size=13, color=T["TD"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 11 — RECOMMENDATIONS
# 2-col numbered cards with unique stripe per card
# ══════════════════════════════════════════════════════════════
def slide_recommendations(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    oval(slide, W-12, -4, 16, 16, T["M"], 0.45)
    dark_header(slide, T, "التوصيات", "Recommendations")

    recs = [r for r in data.get("recommendations",[]) if r]
    n    = min(len(recs), 6)
    if not n: return slide

    cols = 2; cw = (W-3.2)/cols
    gapx = 0.8; gapy = 0.32
    rows = math.ceil(n/cols)
    ch   = (H-3.1-gapy*rows)/rows

    for i, rec in enumerate(recs[:6]):
        col = i%cols; row = i//cols
        cx  = 1.2 + col*(cw+gapx)
        cy  = 3.1 + row*(ch+gapy)
        sc  = T["SC"][i%len(T["SC"])]
        rect(slide, cx, cy, cw, ch, T["M"])
        rect(slide, cx, cy, 0.28, ch, sc)
        txt(slide, f"{i+1:02d}", cx+0.38, cy+0.12, 1.5, ch*0.42,
            font="Calibri", size=30, bold=True, color=T["A"], align=PP_ALIGN.LEFT)
        ln(slide, cx+0.38, cy+ch*0.44, cw-0.55, T["A"], 0.05)
        txt(slide, rec, cx+0.38, cy+ch*0.5, cw-0.55, ch*0.46,
            font=T["BF"], size=12, color=T["TL"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 12 — FUTURE PERSPECTIVES  (vertical timeline)
# ══════════════════════════════════════════════════════════════
def slide_future(prs, data, T):
    slide = blank(prs)
    bg(slide, T["L"])
    cy0 = light_header(slide, T, "آفاق وامتدادات البحث",
                        "Future Research Perspectives")

    futures = [f for f in data.get("futureWork",[]) if f]
    n = min(len(futures), 5)
    if not n: return slide

    tlx = W-3.2
    rect(slide, tlx-0.06, cy0, 0.12, H-cy0-0.4, T["A"])
    gap = 0.28
    fh  = (H-cy0-0.4-gap*n)/n

    for i, fw in enumerate(futures[:5]):
        fy  = cy0 + i*(fh+gap)
        ncy = fy + fh/2 - 0.46
        sc  = T["SC"][i%len(T["SC"])]
        oval(slide, tlx-0.47, ncy, 0.94, 0.94, T["D"])
        oval(slide, tlx-0.37, ncy+0.1, 0.74, 0.74, sc)
        txt(slide, str(i+1), tlx-0.37, ncy+0.1, 0.74, 0.74,
            font="Calibri", size=13, bold=True, color=T["D"], align=PP_ALIGN.CENTER)
        rect(slide, tlx-1.72, ncy+0.39, 1.28, 0.1, T["A"])
        cw = tlx-2.85
        rect(slide, 1.35, fy+0.06, cw, fh, T["CB"], line=T["CE"])
        rect(slide, 1.35, fy+0.06, 0.22, fh, sc)
        txt(slide, f"آفق بحثي {i+1}", 1.7, fy+0.1, 4.5, 0.52,
            font=T["BF"], size=10, bold=True, color=T["A"], align=PP_ALIGN.RIGHT)
        txt(slide, fw, 1.7, fy+0.62, cw-0.55, fh-0.72,
            font=T["BF"], size=12, color=T["TD"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 13 — CONCLUSION
# ══════════════════════════════════════════════════════════════
def slide_conclusion(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    oval(slide, W-13, -4, 18, 18, T["M"], 0.45)
    oval(slide, -4, H-9, 12, 12, T["M"], 0.55)
    rect(slide, 0, 0, W, 0.45, T["A"])
    txt(slide, "الخاتمة والاستنتاجات", 1.2, 0.65, W-2.4, 1.2,
        font=T["HF"], size=28, bold=True, color=T["TL"], align=PP_ALIGN.RIGHT)

    conc = safe(data.get("generalConclusion"))
    qw   = W-5.0; qx = (W-qw)/2
    qy, qh = 2.45, 6.65
    quote_box(slide, T, qx, qy, qw, qh, conc)

    sy = qy+qh+0.3; sh = H-sy-0.18
    rect(slide, 1.2, sy, W-2.4, sh, T["M"])
    rect(slide, 1.2, sy, W-2.4, 0.42, T["A"])
    txt(slide, "أبرز ما توصلت إليه الدراسة", 1.4, sy+0.04, W-2.8, 0.38,
        font=T["BF"], size=12, bold=True, color=T["D"], align=PP_ALIGN.RIGHT)

    results = [r for r in data.get("mainResults",[]) if r]
    top3 = results[:3]
    if top3 and sh > 1.15:
        ch = sh - 0.5
        tw = (W-2.8-0.22*(len(top3)-1))/len(top3)
        for i, r in enumerate(top3):
            tx = 1.35 + i*(tw+0.22)
            sc = T["SC"][i%len(T["SC"])]
            rect(slide, tx, sy+0.48, tw, ch, T["D"])
            rect(slide, tx, sy+0.48, 0.2, ch, sc)
            # accent tag top
            rect(slide, tx, sy+0.48, tw, 0.28, sc)
            txt(slide, r, tx+0.28, sy+0.85, tw-0.38, ch-0.45,
                font=T["BF"], size=10, color=T["TL"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# SLIDE 14 — THANK YOU
# ══════════════════════════════════════════════════════════════
def slide_final(prs, data, T):
    slide = blank(prs)
    bg(slide, T["D"])
    oval(slide, W/2-8, H/2-8, 16, 16, T["M"], 0.38)
    oval(slide, W/2-5.5, H/2-5.5, 11, 11, T["M"], 0.55)
    rect(slide, 0, 0, W, 0.45, T["A"])
    rect(slide, 0, H-0.45, W, 0.45, T["A"])

    ty = H/2-3.2
    txt(slide, "شكراً لحسن استماعكم", 0, ty, W, 1.95,
        font=T["HF"], size=40, bold=True, color=T["TL"], align=PP_ALIGN.CENTER)
    txt(slide, "Merci pour votre attention", 0, ty+2.0, W, 1.1,
        font="Calibri", size=20, italic=True, color=T["A"], align=PP_ALIGN.CENTER)
    txt(slide, safe(data.get("studentName")), 0, ty+3.35, W, 0.9,
        font=T["BF"], size=18, bold=True, color=T["TL"], align=PP_ALIGN.CENTER)
    ln(slide, W/2-4, ty+4.45, 8, T["A"], 0.08)
    if data.get("titleAr"):
        txt(slide, data["titleAr"], 2.0, ty+4.75, W-4.0, 1.0,
            font=T["BF"], size=11, color=T["TM"], align=PP_ALIGN.CENTER)

    refs = [r for r in data.get("references",[]) if r]
    if refs:
        ry0 = H-4.95
        rect(slide, 1.2, ry0, W-2.4, 0.55, T["M"])
        txt(slide, "أبرز المراجع", 1.4, ry0+0.06, W-2.8, 0.44,
            font=T["BF"], size=13, bold=True, color=T["A"], align=PP_ALIGN.RIGHT)
        for i, ref in enumerate(refs[:3]):
            txt(slide, f"[{i+1}]  {ref}", 1.4, ry0+0.65+i*0.95, W-2.8, 0.85,
                font=T["BF"], size=11, color=T["TL"], align=PP_ALIGN.RIGHT)
    return slide

# ══════════════════════════════════════════════════════════════
# ORCHESTRATOR
# ══════════════════════════════════════════════════════════════
def generate_presentation(data: dict, output_path: str) -> None:
    T   = THEMES.get(data.get("theme","navy_gold"), THEMES["navy_gold"])
    prs = Presentation()
    prs.slide_width  = Cm(W)
    prs.slide_height = Cm(H)

    chapters = [
        {"title": "الإشكالية والتساؤلات",  "sub": "Research Problem"},
        {"title": "الأهداف والفرضيات",      "sub": "Objectives & Hypotheses"},
        {"title": "الإطار النظري",           "sub": "Theoretical Framework"},
        {"title": "الدراسات السابقة",        "sub": "Literature Review"},
        {"title": "المنهجية والأدوات",       "sub": "Methodology & Tools"},
        {"title": "النتائج والتوصيات",       "sub": "Results & Recommendations"},
    ]
    def fl(k): return [x for x in data.get(k,[]) if x]

    slide_cover(prs, data, T)
    slide_toc(prs, data, T, chapters)
    slide_problem(prs, data, T)
    slide_objectives(prs, data, T)
    if data.get("importance") or data.get("reasons"):
        slide_importance(prs, data, T)
    concepts = [c for c in data.get("concepts",[]) if c.get("name")]
    if concepts: slide_theory(prs, data, T, concepts)
    lits = [l for l in data.get("literatures",[]) if l.get("title") or l.get("author")]
    if lits: slide_literature(prs, data, T, lits)
    slide_methodology(prs, data, T)
    if [s for s in data.get("stats",[]) if s.get("label") and s.get("value")]:
        slide_stats(prs, data, T)
    if fl("mainResults"):     slide_results(prs, data, T)
    if fl("recommendations"): slide_recommendations(prs, data, T)
    if fl("futureWork"):      slide_future(prs, data, T)
    slide_conclusion(prs, data, T)
    slide_final(prs, data, T)

    prs.save(output_path)
    print(f"✅  {len(prs.slides._sldIdLst)} slides → {output_path}", file=sys.stderr)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generator.py input.json output.pptx", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        payload = json.load(f)
    generate_presentation(payload, sys.argv[2])

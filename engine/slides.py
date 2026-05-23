"""
Canva Engine — مذكرتي Pro v17.2
تخطيط Canva: بطاقات عائمة، تدرجات متعددة، زخارف هندسية

v17.2: Typography System — هرمية الخطوط المتسقة
"""
from __future__ import annotations
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow,
    set_solid_alpha, multi_stop_gradient,
    glow, gradient_oval, diamond, hexagon,
    decorative_dots, card_3d,
    icon_circle, number_badge,
    slide_number,
    txt, blank_slide,
)
from engine.typography import TS, FONT_TITLE, FONT_BODY, FONT_NUM, FONT_EN, display_size, stat_size
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"

def set_font(font_name: str):
    global _FONT
    _FONT = font_name

# ══════════════════════════════════════════════════════════════
# BACKGROUND CANVAS
# ══════════════════════════════════════════════════════════════
def _bg_canvas(slide, T: Theme, style='default'):
    bg(slide, T.bg_rgb)
    if style == 'default':
        gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)
        oval(slide, -4, -4, 13, 13, T.accent_rgb, alpha=5)
        oval(slide, W-11, H-10, 18, 18, T.bg2_rgb, alpha=50)
        oval(slide, W-6, -3, 10, 10, T.accent_rgb, alpha=4)
        decorative_dots(slide, 1.5, H-4.5, 6, 3, 0.18, 0.45, T.accent_rgb, alpha=12)
    elif style == 'split':
        gradient_rect(slide, 0, 0, W, H*0.55, T.grad2, T.grad1, angle=90)
        gradient_rect(slide, 0, H*0.55, W, H*0.45, T.bg2, T.bg, angle=90)
        div = rect(slide, 0, H*0.55-0.08, W, 0.16, T.accent_rgb)
        if div: multi_stop_gradient(div, [(0,T.bg),(50,T.accent),(100,T.bg)], angle=0)
        oval(slide, -3, -3, 12, 12, T.accent_rgb, alpha=5)
        oval(slide, W-8, H-8, 14, 14, T.accent_rgb, alpha=4)
        decorative_dots(slide, W-7, 1.5, 5, 4, 0.16, 0.4, T.accent_rgb, alpha=10)
    elif style == 'geometric':
        gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)
        diamond(slide, W-8, -2.5, 7, 7, T.accent_rgb, alpha=6)
        diamond(slide, -2, H-5, 5, 5, T.accent_rgb, alpha=5)
        diamond(slide, W*0.4, H-4, 4, 4, T.bg2_rgb, alpha=40)
        hexagon(slide, W-5, H*0.3, 3, 3, T.accent_rgb, alpha=7)
        decorative_dots(slide, 1.2, 2.0, 4, 5, 0.16, 0.38, T.accent_rgb, alpha=10)
    elif style == 'radial':
        gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=45)
        cx2, cy2 = W*0.5, H*0.5
        for r_size, alpha in [(28,4),(22,5),(16,6),(10,8)]:
            oval(slide, cx2-r_size/2, cy2-r_size/2, r_size, r_size, T.accent_rgb, alpha=alpha)
        decorative_dots(slide, 2.0, H-4, 5, 2, 0.2, 0.5, T.accent_rgb, alpha=12)

# ══════════════════════════════════════════════════════════════
# HEADERS
# ══════════════════════════════════════════════════════════════
def _header_wave(slide, T, title, subtitle=""):
    gradient_rect(slide, 0, 0, W, 3.2, T.grad2, T.grad1, angle=135)
    acc = rect(slide, 0, 3.2-0.18, W, 0.18, T.accent_rgb)
    if acc: multi_stop_gradient(acc, [(0,T.bg),(40,T.accent),(60,T.accent2),(100,T.bg)], 0)
    oval(slide, W-4.5, -2.0, 6.5, 6.5, T.accent_rgb, alpha=10)
    oval(slide, -2, -1.5, 5, 5, T.bg2_rgb, alpha=60)
    txt(slide, title, 0.8, 0.3, W-1.6, 1.6, font=FONT_TITLE, size=TS.H1, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
    if subtitle:
        txt(slide, subtitle, 0.8, 1.95, W-1.6, 1.0, font=FONT_BODY, size=TS.BODY_LG,
            bold=False, italic=True, color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

def _header_bar(slide, T, title, subtitle=""):
    gradient_rect(slide, 0, 0, W, 3.0, T.grad1, T.grad2, angle=90)
    acc_v = rect(slide, 0, 0, 0.5, 3.0, T.accent_rgb)
    if acc_v: gradient_fill(acc_v, T.accent_grad1, T.accent_grad2, 90)
    acc_h = rect(slide, 0, 3.0-0.1, W, 0.1, T.accent_rgb)
    if acc_h: gradient_fill(acc_h, T.accent_grad1, T.accent_grad2, 0)
    oval(slide, W-5, -1.8, 6, 6, T.accent_rgb, alpha=8)
    txt(slide, title, 0.8, 0.35, W-1.6, 1.55, font=FONT_TITLE, size=TS.H1, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
    if subtitle:
        txt(slide, subtitle, 0.8, 1.98, W-1.6, 0.85, font=FONT_BODY, size=TS.BODY_LG,
            bold=False, color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

def _section_header(slide, T, title, subtitle="", style=0):
    if style % 2 == 0:
        _header_wave(slide, T, title, subtitle)
    else:
        _header_bar(slide, T, title, subtitle)

# ══════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'geometric')

    top = rect(slide, 0, 0, W, 0.45, T.accent_rgb)
    if top: multi_stop_gradient(top, [(0,T.bg),(30,T.accent),(70,T.accent2),(100,T.bg)], 0)
    bot = rect(slide, 0, H-0.3, W, 0.3, T.accent_rgb)
    if bot: gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)

    if req.institution:
        ib = rrect(slide, W/2-9, 0.65, 18, 0.75, T.card_rgb, radius_pct=40)
        if ib: set_solid_alpha(ib, 85)
        txt(slide, req.institution, W/2-9, 0.65, 18, 0.75, font=FONT_BODY, size=TS.COVER_INST,
            bold=False, color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    cy = H*0.2; cw = W-4.0; cx = 2.0; ch = H*0.42
    mc = rrect(slide, cx, cy, cw, ch, T.card_rgb, radius_pct=14)
    if mc:
        multi_stop_gradient(mc, [(0,T.card),(100,T.bg2)], angle=135)
        shadow(mc, blur=28, dist=8, alpha=0.5)

    ct = rrect(slide, cx, cy, cw, 0.38, T.accent_rgb, radius_pct=0)
    if ct: multi_stop_gradient(ct, [(0,T.accent),(50,T.accent2),(100,T.accent)], 0)
    vline(slide, cx+cw-0.25, cy+0.38, ch-0.38, T.accent_rgb, thickness=0.25)

    ts = 26 if len(req.title_ar)<45 else 21 if len(req.title_ar)<75 else 17
    txt(slide, req.title_ar, cx+0.5, cy+0.55, cw-1.0, ch*0.62,
        font=FONT_TITLE, size=ts, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    if req.title_en:
        txt(slide, req.title_en, cx+0.5, cy+ch*0.66, cw-1.0, 0.85,
            font=FONT_EN, size=TS.BODY_SM, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    div_y = cy+ch*0.76
    dl = rect(slide, cx+cw*0.1, div_y, cw*0.8, 0.05, T.accent_rgb)
    if dl: multi_stop_gradient(dl, [(0,T.bg2),(50,T.accent),(100,T.bg2)], 0)

    info_y = div_y+0.25; rh = 0.62

    def _row(label, value, y):
        rb = rrect(slide, cx+0.4, y, cw-0.8, rh, T.bg_rgb, radius_pct=8)
        if rb: set_solid_alpha(rb, 60)
        txt(slide, label, cx+0.6, y, 4.5, rh, font=FONT_TITLE, size=TS.COVER_LABEL, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        vline(slide, cx+5.3, y+0.08, rh-0.16, T.muted_rgb, thickness=0.04)
        txt(slide, value, cx+5.5, y, cw-6.3, rh, font=FONT_BODY, size=TS.COVER_VALUE,
            bold=False, color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    r = 0
    _row("الطالب :", req.student_name, info_y+r*rh); r+=1
    if req.supervisor: _row("المشرف :", req.supervisor, info_y+r*rh); r+=1
    if req.co_supervisor: _row("المشرف المساعد :", req.co_supervisor, info_y+r*rh); r+=1
    if req.specialization: _row("التخصص :", req.specialization, info_y+r*rh)

    if req.year:
        yb = rrect(slide, W/2-2.5, H-1.55, 5.0, 0.6, T.accent_rgb, radius_pct=50)
        if yb:
            multi_stop_gradient(yb, [(0,T.accent),(100,T.accent2)], 0)
            shadow(yb, blur=12, dist=3, alpha=0.35)
            glow(yb, T.accent.lstrip('#'), radius=15, alpha=0.25)
        txt(slide, req.year, W/2-2.5, H-1.55, 5.0, 0.6, font=FONT_NUM, size=TS.COVER_YEAR,
            bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)
    return slide

# ══════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'split')
    _section_header(slide, T, "مقدمة البحث", "نظرة عامة على الدراسة", style=0)

    cy = 3.4
    items = []
    if req.intro_overview: items.append(("📖", "نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("🔬", "المنهج المتبع", req.intro_approach))

    avail_h = H-cy-0.7
    n = max(len(items), 1)
    col_w = (W-2.6-0.3*(n-1)) / n

    for i, (icon, lbl, val) in enumerate(items[:2]):
        x = 1.3 + i*(col_w+0.3)
        c = rrect(slide, x, cy, col_w, avail_h, T.card_rgb, radius_pct=12)
        if c:
            multi_stop_gradient(c, [(0,T.card),(60,T.bg2),(100,T.bg)], angle=150)
            shadow(c, blur=18, dist=5, alpha=0.4)
        tp = rrect(slide, x, cy, col_w, 0.32, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp, [(0,T.accent),(100,T.accent2)], 0)

        icon_circle(slide, x+col_w/2-1.0, cy+0.5, 2.0,
                    T.accent_grad1, T.accent_grad2, icon, icon_size=22, T=T)

        txt(slide, lbl, x+0.3, cy+2.65, col_w-0.6, 0.75,
            font=FONT_TITLE, size=TS.H2, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=True)
        hline(slide, x+col_w*0.15, cy+3.45, col_w*0.7, T.accent_rgb, thickness=0.04)
        txt(slide, val, x+0.3, cy+3.6, col_w-0.6, avail_h-4.0,
            font=FONT_BODY, size=TS.BODY, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 1, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'default')
    _section_header(slide, T, "خطة البحث",
                    f"يتكون البحث من {len(req.chapters)} فصول", style=1)

    cy = 3.2
    chapters = req.chapters[:8]
    avail_h = H-cy-0.6
    row_h = min(avail_h/max(len(chapters),1)-0.12, 1.65)

    for i, ch in enumerate(chapters):
        y = cy + i*(row_h+0.12)
        even = i%2==0
        row = rrect(slide, 1.3, y, W-2.6, row_h,
                    T.card_rgb if even else T.bg2_rgb, radius_pct=9)
        if row:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(row, stops, angle=0 if even else 180)
            shadow(row, blur=7, dist=2, alpha=0.22)

        num_y = y+(row_h-0.75)/2
        nc = oval(slide, W-3.2, num_y, 0.75, 0.75, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc, [(0,T.accent),(100,T.accent2)], 135)
            shadow(nc, blur=8, dist=2, alpha=0.3)
        txt(slide, str(i+1), W-3.2, num_y, 0.75, 0.75, font=FONT_NUM, size=TS.H3,
            bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, ch.title, 1.6, y, W-5.3, row_h, font=FONT_BODY, size=TS.H3,
            bold=False, color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        if ch.pages:
            pg = rrect(slide, 1.5, y+(row_h-0.42)/2, 1.8, 0.42, T.bg_rgb, radius_pct=40)
            if pg: set_solid_alpha(pg, 60)
            txt(slide, ch.pages, 1.5, y+(row_h-0.42)/2, 1.8, 0.42,
                font=FONT_NUM, size=TS.LABEL, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=False)

    slide_number(slide, 2, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# PROBLEM
# ══════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'geometric')
    _section_header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية", style=0)

    cy = 3.4
    if req.main_problem:
        ph = 2.6
        pc = rrect(slide, 1.3, cy, W-2.6, ph, T.card_rgb, radius_pct=12)
        if pc:
            multi_stop_gradient(pc, [(0,T.card),(100,T.bg2)], 135)
            shadow(pc, blur=20, dist=6, alpha=0.45)
            glow(pc, T.accent.lstrip('#'), radius=25, alpha=0.12)
        lb = rrect(slide, W-7.0, cy, 5.5, 0.55, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb, [(0,T.accent),(100,T.accent2)], 0)
        txt(slide, "◆ الإشكالية الرئيسية", W-7.0, cy, 5.5, 0.55, font=FONT_TITLE,
            size=TS.H2, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, "❝", 1.6, cy+0.65, 1.5, 1.2, font=FONT_EN, size=34,
            bold=False, color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)
        txt(slide, req.main_problem, 3.2, cy+0.65, W-5.0, ph-0.85,
            font=FONT_BODY, size=TS.BODY_LG, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        cy += ph+0.25

    if req.main_question:
        qh = 1.7
        qc = rrect(slide, 1.3, cy, W-2.6, qh, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=10, dist=3, alpha=0.3)
        vline(slide, W-1.55, cy, qh, T.accent_rgb, thickness=0.25)
        dot = oval(slide, W-3.5, cy+qh/2-0.22, 0.44, 0.44, T.accent_rgb)
        if dot: multi_stop_gradient(dot, [(0,T.accent),(100,T.accent2)], 135)
        txt(slide, req.main_question, 1.6, cy, W-4.0, qh, font=FONT_BODY,
            size=TS.BODY_LG, bold=True, italic=True, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        cy += qh+0.2

    if req.sub_questions:
        avail = H-cy-0.4
        sub_h = min(avail/max(len(req.sub_questions),1), 0.95)
        for i, q in enumerate(req.sub_questions[:6]):
            y = cy+i*sub_h
            nc = oval(slide, W-2.8, y+(sub_h-0.38)/2, 0.38, 0.38, T.accent_rgb)
            if nc: set_solid_alpha(nc, 70)
            txt(slide, str(i+1), W-2.8, y+(sub_h-0.38)/2, 0.38, 0.38,
                font=FONT_NUM, size=TS.LABEL, bold=True, color=T.accent_rgb,
                align=PP_ALIGN.CENTER, rtl=False)
            txt(slide, q, 1.6, y, W-3.6, sub_h, font=FONT_BODY, size=TS.BODY,
                bold=False, color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 3, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# OBJECTIVES
# ══════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'split')
    _section_header(slide, T, "أهداف البحث وفرضياته", "", style=1)

    cy = 3.3
    cols = []
    if req.objectives: cols.append(("🎯 الأهداف", req.objectives))
    if req.hypotheses:  cols.append(("💡 الفرضيات", req.hypotheses))
    if not cols: return slide

    col_w = (W-2.6-0.35*(len(cols)-1)) / len(cols)
    avail_h = H-cy-0.6

    for i, (lbl, items) in enumerate(cols[:2]):
        x = 1.3+i*(col_w+0.35)
        c = rrect(slide, x, cy, col_w, avail_h, T.card_rgb, radius_pct=12)
        if c:
            multi_stop_gradient(c, [(0,T.card),(100,T.bg2)], 150)
            shadow(c, blur=18, dist=5, alpha=0.4)
        hdr = rrect(slide, x, cy, col_w, 0.75, T.accent_rgb, radius_pct=0)
        if hdr: multi_stop_gradient(hdr, [(0,T.accent2),(100,T.accent)], 0)
        txt(slide, lbl, x+0.2, cy, col_w-0.4, 0.75, font=FONT_TITLE, size=TS.H2,
            bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        item_avail = avail_h-0.9
        item_h = min(item_avail/max(len(items),1), 1.2)-0.1

        for j, item in enumerate(items[:8]):
            iy = cy+0.88+j*(item_h+0.1)
            if iy+item_h > H-0.5: break
            rb = rrect(slide, x+0.15, iy, col_w-0.3, item_h,
                       T.bg2_rgb if j%2==0 else T.bg_rgb, radius_pct=7)
            if rb: set_solid_alpha(rb, 80)
            number_badge(slide, x+col_w-0.85, iy+(item_h-0.55)/2, 0.55, j+1, T)
            txt(slide, item, x+0.28, iy+0.06, col_w-1.35, item_h-0.12,
                font=FONT_BODY, size=TS.BODY_SM, bold=False, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 4, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# IMPORTANCE
# ══════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'geometric')
    _section_header(slide, T, "أهمية البحث", "الأثر والقيمة المضافة", style=0)

    cy = 3.3
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐","🔑","📌","🌟","💎","🏆"]
    cols = 2 if len(items) > 3 else 1
    col_w = (W-2.6-0.3*(cols-1)) / cols
    rows = (len(items)+cols-1)//cols
    avail_h = H-cy-0.5
    card_h = min(avail_h/rows-0.2, 2.4)

    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = 1.3+ci*(col_w+0.3)
        y = cy+ri*(card_h+0.2)

        c = card_3d(slide, x, y, col_w, card_h, T, radius=10)
        acc = rrect(slide, x+col_w-0.3, y, 0.3, card_h, T.accent_rgb, radius_pct=0)
        if acc: multi_stop_gradient(acc, [(0,T.accent2),(100,T.accent)], 90)

        icon_circle(slide, x+0.25, y+(card_h-1.3)/2, 1.3,
                    T.accent_grad1, T.accent_grad2,
                    icons[i%len(icons)], icon_size=18, T=T)

        txt(slide, item, x+1.75, y+0.15, col_w-2.35, card_h-0.3,
            font=FONT_BODY, size=TS.BODY, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 5, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'radial')
    _section_header(slide, T, "منهجية البحث", "الإجراءات والأدوات المستخدمة", style=1)

    cy = 3.3
    icons_map = {"المنهج":"📊","العينة":"👥","حجم العينة":"📏","الأداة":"🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج", req.methodology))
    if req.sample_type:  fields.append(("العينة", req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة", req.tool))

    cols = 2 if len(fields)>2 else len(fields)
    col_w = (W-2.6-0.3*(cols-1)) / max(cols,1)
    rows = (len(fields)+cols-1)//cols
    avail_h = H-cy-0.5
    card_h = min(avail_h/rows-0.2, 3.8)

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i%cols; ri = i//cols
        x = 1.3+ci*(col_w+0.3)
        y = cy+ri*(card_h+0.2)

        c = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=12)
        if c:
            multi_stop_gradient(c, [(0,T.card),(100,T.bg2)], angle=145)
            shadow(c, blur=16, dist=5, alpha=0.42)

        ic_x = x+col_w/2-1.1
        ic = oval(slide, ic_x, y+0.35, 2.2, 2.2, T.accent_rgb)
        if ic:
            multi_stop_gradient(ic, [(0,T.accent),(100,T.accent2)], 135)
            shadow(ic, blur=12, dist=3, alpha=0.35)
            glow(ic, T.accent.lstrip('#'), radius=18, alpha=0.2)
        txt(slide, icons_map.get(lbl,"📌"), ic_x, y+0.45, 2.2, 1.9,
            font=FONT_NUM, size=TS.STAT_SM, bold=False, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, lbl, x+0.25, y+2.68, col_w-0.5, 0.72, font=FONT_TITLE,
            size=TS.H3, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)
        hline(slide, x+col_w*0.15, y+3.45, col_w*0.7, T.muted_rgb, thickness=0.04)
        txt(slide, val, x+0.25, y+3.58, col_w-0.5, card_h-3.75,
            font=FONT_BODY, size=TS.BODY, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 6, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# STATS
# ══════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'default')
    _section_header(slide, T, "الأرقام والإحصاءات الرئيسية", "", style=0)

    cy = 3.3
    stats = req.stats[:6]
    if not stats: return slide

    cols = 3 if len(stats)>=3 else len(stats)
    rows = (len(stats)+cols-1)//cols
    gap = 0.3
    col_w = (W-2.6-gap*(cols-1)) / cols
    avail_h = H-cy-0.5
    card_h = min(avail_h/rows-gap, 4.2)

    for i, stat in enumerate(stats):
        ci = i%cols; ri = i//cols
        x = 1.3+ci*(col_w+gap)
        y = cy+ri*(card_h+gap)

        c = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=14)
        if c:
            multi_stop_gradient(c, [(0,T.bg2),(50,T.card),(100,T.bg2)], angle=135)
            shadow(c, blur=20, dist=6, alpha=0.45)

        tp = rrect(slide, x, y, col_w, 0.35, T.accent_rgb, radius_pct=0)
        if tp:
            multi_stop_gradient(tp, [(0,T.accent2),(50,T.accent),(100,T.accent2)], 0)
            glow(tp, T.accent.lstrip('#'), radius=12, alpha=0.3)

        bp = rrect(slide, x, y+card_h-0.25, col_w, 0.25, T.accent_rgb, radius_pct=0)
        if bp: set_solid_alpha(bp, 40)

        vs = 38 if len(stat.value)<=4 else 28 if len(stat.value)<=8 else 20
        txt(slide, stat.value, x+0.2, y+0.5, col_w-0.4, card_h*0.5,
            font=FONT_NUM, size=vs, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

        if stat.unit:
            ub = rrect(slide, x+col_w/2-1.5, y+card_h*0.52+0.15, 3.0, 0.5,
                       T.bg_rgb, radius_pct=40)
            if ub: set_solid_alpha(ub, 60)
            txt(slide, stat.unit, x+col_w/2-1.5, y+card_h*0.52+0.15, 3.0, 0.5,
                font=FONT_BODY, size=TS.LABEL, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=True)

        hline(slide, x+col_w*0.15, y+card_h*0.68, col_w*0.7, T.muted_rgb, thickness=0.04)
        txt(slide, stat.label, x+0.2, y+card_h*0.7, col_w-0.4, card_h*0.27,
            font=FONT_BODY, size=TS.BODY, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 7, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'split')
    _section_header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة", style=1)

    cy = 3.2
    results = req.main_results[:8]
    avail_h = H-cy-0.5
    item_h = min(avail_h/max(len(results),1)-0.12, 1.55)

    for i, result in enumerate(results):
        y = cy+i*(item_h+0.12)
        even = i%2==0
        c = rrect(slide, 1.3, y, W-2.6, item_h,
                  T.card_rgb if even else T.bg2_rgb, radius_pct=9)
        if c:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(c, stops, angle=0)
            shadow(c, blur=6, dist=2, alpha=0.2)

        acc = rrect(slide, W-1.55, y, 0.22, item_h, T.accent_rgb, radius_pct=0)
        if acc:
            gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(acc, max(20, 60-i*8))

        number_badge(slide, W-3.1, y+(item_h-0.6)/2, 0.6, i+1, T)

        txt(slide, result, 1.55, y+0.1, W-5.1, item_h-0.2,
            font=FONT_BODY, size=TS.BODY_LG, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 8, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'radial')
    _section_header(slide, T, "خاتمة البحث", "الاستنتاج العام", style=0)

    card_y = 3.3; card_h = H-card_y-0.7; cw = W-3.0
    c = rrect(slide, 1.5, card_y, cw, card_h, T.card_rgb, radius_pct=16)
    if c:
        multi_stop_gradient(c, [(0,T.card),(50,T.bg2),(100,T.card)], angle=135)
        shadow(c, blur=28, dist=8, alpha=0.5)
        glow(c, T.accent.lstrip('#'), radius=30, alpha=0.1)

    tp = rrect(slide, 1.5, card_y, cw, 0.38, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp, [(0,T.accent2),(50,T.accent),(100,T.accent2)], 0)
        glow(tp, T.accent.lstrip('#'), radius=15, alpha=0.35)

    diamond(slide, 1.8, card_y+0.55, 1.2, 1.2, T.accent_rgb, alpha=15)
    diamond(slide, W-3.2, card_y+card_h-1.8, 1.0, 1.0, T.accent_rgb, alpha=10)

    txt(slide, "❝", 2.2, card_y+0.5, 2.0, 1.8, font=FONT_EN, size=52,
        bold=False, color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)
    txt(slide, req.general_conclusion, 2.2, card_y+1.4, cw-1.4, card_h-2.2,
        font=FONT_BODY, size=TS.H2, bold=False, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, rtl=True)

    div_y = card_y+card_h-1.2
    hl = rect(slide, 1.5+cw*0.2, div_y, cw*0.6, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0,T.bg2),(50,T.accent),(100,T.bg2)], 0)
    txt(slide, req.student_name, 1.5, div_y+0.15, cw, 0.78,
        font=FONT_TITLE, size=TS.H2, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 9, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'geometric')
    _section_header(slide, T, "توصيات البحث", "", style=1)

    cy = 3.2
    recs = req.recommendations[:8]
    avail_h = H-cy-0.5
    item_h = min(avail_h/max(len(recs),1)-0.12, 1.5)

    for i, rec in enumerate(recs):
        y = cy+i*(item_h+0.12)
        c = rrect(slide, 1.3, y, W-2.6, item_h, T.card_rgb, radius_pct=10)
        if c:
            multi_stop_gradient(c, [(0,T.card),(100,T.bg2)], angle=0)
            shadow(c, blur=8, dist=2, alpha=0.25)

        dot = oval(slide, W-1.8, y+(item_h-0.42)/2, 0.42, 0.42, T.accent_rgb)
        if dot:
            multi_stop_gradient(dot, [(0,T.accent),(100,T.accent2)], 135)
            shadow(dot, blur=6, dist=1, alpha=0.3)

        acc = rect(slide, W-1.55, y, 0.2, item_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        txt(slide, rec, 1.55, y+0.1, W-3.5, item_h-0.2,
            font=FONT_BODY, size=TS.BODY, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 10, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# FUTURE
# ══════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'default')
    _section_header(slide, T, "آفاق البحث المستقبلية", "", style=0)

    cy = 3.3
    items = req.future_work[:6]
    cols = 2 if len(items)>3 else 1
    col_w = (W-2.6-0.3*(cols-1)) / cols
    rows = (len(items)+cols-1)//cols
    avail_h = H-cy-0.5
    card_h = min(avail_h/rows-0.2, 3.5)

    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = 1.3+ci*(col_w+0.3)
        y = cy+ri*(card_h+0.2)

        c = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=12)
        if c:
            multi_stop_gradient(c, [(0,T.card),(70,T.bg2),(100,T.bg)], angle=160)
            shadow(c, blur=16, dist=4, alpha=0.38)

        tp = rrect(slide, x, y, col_w, 0.3, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp, [(0,T.accent),(100,T.accent2)], 0)

        number_badge(slide, x+col_w/2-0.45, y+0.45, 0.9, i+1, T)
        hline(slide, x+col_w*0.2, y+1.5, col_w*0.6, T.muted_rgb, thickness=0.04)
        txt(slide, item, x+0.3, y+1.65, col_w-0.6, card_h-1.85,
            font=FONT_BODY, size=TS.BODY_LG, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 11, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg_canvas(slide, T, 'split')
    _section_header(slide, T, "قائمة المراجع والمصادر", "", style=1)

    cy = 3.2
    refs = req.references[:12]
    avail_h = H-cy-0.4
    item_h = max(min(avail_h/max(len(refs),1)-0.1, 1.2), 0.55)

    for i, ref in enumerate(refs):
        y = cy+i*(item_h+0.1)
        if y+item_h > H-0.25: break
        even = i%2==0
        c = rrect(slide, 1.3, y, W-2.6, item_h,
                  T.card_rgb if even else T.bg2_rgb, radius_pct=5)
        if c:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(c, stops, 0)

        acc = rect(slide, W-1.55, y, 0.22, item_h, T.accent_rgb)
        if acc: set_solid_alpha(acc, 60)

        nb = rrect(slide, 1.45, y+(item_h-0.42)/2, 0.72, 0.42, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 70)
        txt(slide, f"[{i+1}]", 1.45, y+(item_h-0.42)/2, 0.72, 0.42,
            font=FONT_NUM, size=TS.LABEL, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False)
        txt(slide, ref, 2.35, y+0.05, W-4.2, item_h-0.1,
            font=FONT_BODY, size=TS.BODY_SM, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 12, 13, T)
    return slide

# ══════════════════════════════════════════════════════════════
# FINAL / THANK YOU
# ══════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    oval(slide, -5, -5, 16, 16, T.accent_rgb, alpha=5)
    oval(slide, W-12, H-12, 18, 18, T.accent_rgb, alpha=4)
    oval(slide, W-7, -3, 11, 11, T.bg2_rgb, alpha=35)
    oval(slide, -3, H-7, 10, 10, T.bg2_rgb, alpha=30)
    diamond(slide, W*0.3, H*0.05, 2.5, 2.5, T.accent_rgb, alpha=8)
    diamond(slide, W*0.65, H*0.75, 2.0, 2.0, T.accent_rgb, alpha=6)
    decorative_dots(slide, 1.5, H-5, 7, 3, 0.18, 0.42, T.accent_rgb, alpha=12)
    decorative_dots(slide, W-6, 1.5, 5, 4, 0.16, 0.38, T.accent_rgb, alpha=10)

    cw = 23; ch = 11
    cx = (W-cw)/2; cy2 = (H-ch)/2

    c = rrect(slide, cx, cy2, cw, ch, T.card_rgb, radius_pct=16)
    if c:
        multi_stop_gradient(c, [(0,T.card),(50,T.bg2),(100,T.card)], angle=135)
        shadow(c, blur=32, dist=10, alpha=0.55)
        glow(c, T.accent.lstrip('#'), radius=40, alpha=0.12)

    tp = rrect(slide, cx, cy2, cw, 0.45, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp,
            [(0,T.bg),(30,T.accent2),(50,T.accent),(70,T.accent2),(100,T.bg)], 0)
        glow(tp, T.accent.lstrip('#'), radius=20, alpha=0.4)

    bp = rrect(slide, cx, cy2+ch-0.3, cw, 0.3, T.accent_rgb, radius_pct=0)
    if bp: set_solid_alpha(bp, 50)

    txt(slide, "✦", cx+cw/2-0.8, cy2+0.55, 1.6, 1.5, font=FONT_EN,
        size=28, bold=False, color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

    txt(slide, "شكراً وتقديراً", cx+0.8, cy2+1.2, cw-1.6, 2.8,
        font=FONT_TITLE, size=TS.DISPLAY_MD, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    div_y = cy2+4.2
    dm = rect(slide, cx+cw*0.15, div_y, cw*0.7, 0.06, T.accent_rgb)
    if dm: multi_stop_gradient(dm, [(0,T.bg2),(50,T.accent),(100,T.bg2)], 0)
    rect(slide, cx+cw*0.25, div_y+0.13, cw*0.5, 0.03, T.muted_rgb)

    txt(slide, req.student_name, cx+0.8, div_y+0.3, cw-1.6, 1.4,
        font=FONT_TITLE, size=TS.H1, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    ts = req.title_ar[:72]+("..." if len(req.title_ar)>72 else "")
    txt(slide, ts, cx+1.2, div_y+1.85, cw-2.4, 2.0,
        font=FONT_BODY, size=TS.BODY, bold=False, italic=True, color=T.muted_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    footer = []
    if req.institution: footer.append(req.institution)
    if req.year: footer.append(req.year)
    if footer:
        fb = rrect(slide, cx+cw*0.1, cy2+ch-1.35, cw*0.8, 0.62, T.bg_rgb, radius_pct=40)
        if fb: set_solid_alpha(fb, 55)
        txt(slide, "  ·  ".join(footer), cx+0.8, cy2+ch-1.35, cw-1.6, 0.62,
            font=FONT_BODY, size=TS.BODY, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True)

    bottom = rect(slide, 0, H-0.28, W, 0.28, T.accent_rgb)
    if bottom:
        multi_stop_gradient(bottom,
            [(0,T.bg),(30,T.accent),(70,T.accent2),(100,T.bg)], 0)

    return slide

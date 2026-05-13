"""
Slide Builder — مذكرتي Pro v17
Each make_* function builds exactly one slide type.
Pure functions: (prs, request, theme) → slide
No global state. No file I/O. No threading.
"""
from __future__ import annotations

import sys
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from engine.primitives import (
    W, H, cm, pt,
    rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow,
    set_solid_alpha,
    txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

# Default font — will be overridden at runtime if Cairo not available
_FONT = "Cairo"


def set_font(font_name: str):
    global _FONT
    _FONT = font_name


def _hx(h: str) -> RGBColor:
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ══════════════════════════════════════════════════════════════════════
# COVER SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Background gradient overlay
    s = gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    # Decorative geometric elements
    # Top-left corner accent
    s1 = rrect(slide, -0.5, -0.5, 8, 8, T.accent_rgb, radius_pct=0)
    if s1:
        gradient_fill(s1, T.accent_grad1, T.accent_grad2, angle=45)
        set_solid_alpha(s1, 8)

    # Bottom-right circle
    s2 = oval(slide, W - 10, H - 10, 14, 14, T.accent_rgb, alpha=6)

    # Top accent bar
    accent_bar = rrect(slide, 0, 0, W, 0.55, T.accent_rgb, radius_pct=0)
    if accent_bar:
        gradient_fill(accent_bar, T.accent_grad1, T.accent_grad2, angle=0)

    # Institution badge (top-left area)
    if req.institution:
        institution_card = rrect(slide, 1.2, 0.9, 14, 0.7, T.card_rgb, radius_pct=35)
        if institution_card:
            shadow(institution_card, blur=10, dist=3, alpha=0.3)
        txt(slide, req.institution, 1.4, 0.9, 13.6, 0.7,
            font=_FONT, size=11, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    # Main title area — centered
    title_y = H * 0.28
    title_h = H * 0.32

    # Title card background
    title_card = rrect(slide, 1.5, title_y - 0.3, W - 3.0, title_h + 0.6,
                       T.card_rgb, radius_pct=12)
    if title_card:
        shadow(title_card, blur=20, dist=6, alpha=0.4)

    # Accent top stripe on card
    stripe = rrect(slide, 1.5, title_y - 0.3, W - 3.0, 0.25, T.accent_rgb, radius_pct=0)
    if stripe:
        gradient_fill(stripe, T.accent_grad1, T.accent_grad2, angle=0)

    # Title text
    title_size = 28 if len(req.title_ar) < 50 else 22 if len(req.title_ar) < 80 else 18
    txt(slide, req.title_ar,
        2.0, title_y, W - 4.0, title_h,
        font=_FONT, size=title_size, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # English title if provided
    if req.title_en:
        txt(slide, req.title_en,
            2.0, title_y + title_h - 0.8, W - 4.0, 0.9,
            font="Calibri", size=12, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # Divider line
    div_y = title_y + title_h + 0.5
    hline(slide, W * 0.25, div_y, W * 0.5, T.accent_rgb, thickness=0.06)

    # Student info block
    info_y = div_y + 0.3
    row_h = 0.65

    # Student name
    _info_row(slide, T, "الطالب :", req.student_name, info_y)

    # Supervisor
    if req.supervisor:
        _info_row(slide, T, "المشرف :", req.supervisor, info_y + row_h)

    if req.co_supervisor:
        _info_row(slide, T, "المشرف المساعد :", req.co_supervisor, info_y + row_h * 2)

    if req.specialization:
        extra_y = info_y + row_h * (3 if req.co_supervisor else 2 if req.supervisor else 1)
        _info_row(slide, T, "التخصص :", req.specialization, extra_y)

    # Year badge
    if req.year:
        yr_y = H - 0.85
        yr_card = rrect(slide, W / 2 - 2.2, yr_y, 4.4, 0.52, T.accent_rgb, radius_pct=50)
        if yr_card:
            gradient_fill(yr_card, T.accent_grad1, T.accent_grad2, angle=0)
            shadow(yr_card, blur=10, dist=2, alpha=0.3)
        txt(slide, req.year, W / 2 - 2.2, yr_y, 4.4, 0.52,
            font="Calibri", size=12, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # Bottom accent bar
    bottom_bar = rect(slide, 0, H - 0.22, W, 0.22, T.accent_rgb)
    if bottom_bar:
        gradient_fill(bottom_bar, T.accent_grad1, T.accent_grad2, angle=0)

    return slide


def _info_row(slide, T: Theme, label: str, value: str, y: float):
    """Render a label:value row for cover slide."""
    row_bg = rrect(slide, 1.5, y - 0.05, W - 3.0, 0.58, T.bg2_rgb, radius_pct=8)
    # Label
    txt(slide, label, 1.7, y, 5.5, 0.58,
        font=_FONT, size=11, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
    # Value
    txt(slide, value, 7.5, y, W - 9.2, 0.58,
        font=_FONT, size=12, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)


# _set_alpha removed — use set_solid_alpha from primitives


# ══════════════════════════════════════════════════════════════════════
# SECTION HEADER (reusable)
# ══════════════════════════════════════════════════════════════════════
def _section_header(slide, T: Theme, title: str, subtitle: str = ""):
    """Standard section header at top of slide."""
    # Header strip
    header = gradient_rect(slide, 0, 0, W, 2.8, T.grad1, T.grad2, angle=135)
    # Accent left bar
    accent_line = rect(slide, 0, 0, 0.35, 2.8, T.accent_rgb)
    if accent_line:
        gradient_fill(accent_line, T.accent_grad1, T.accent_grad2, angle=90)

    # Decorative circle
    oval(slide, W - 3.5, -1.2, 4.8, 4.8, T.accent_rgb, alpha=8)

    # Title
    txt(slide, title, 0.6, 0.35, W - 1.2, 1.35,
        font=_FONT, size=22, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    if subtitle:
        txt(slide, subtitle, 0.6, 1.6, W - 1.2, 0.9,
            font=_FONT, size=13, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)


# ══════════════════════════════════════════════════════════════════════
# INTRO SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")

    content_y = 3.1
    col_w = (W - 2.8) / 2
    gap = 0.3

    items = []
    if req.intro_overview:
        items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:
        items.append(("المنهج المتبع", req.intro_approach))

    for i, (lbl, val) in enumerate(items[:2]):
        x = 1.2 + i * (col_w + gap)
        card = rrect(slide, x, content_y, col_w, H - content_y - 0.8, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=14, dist=4, alpha=0.35)
        # Card accent top
        card_stripe = rrect(slide, x, content_y, col_w, 0.22, T.accent_rgb, radius_pct=0)
        if card_stripe:
            gradient_fill(card_stripe, T.accent_grad1, T.accent_grad2, angle=0)
        txt(slide, lbl, x + 0.2, content_y + 0.28, col_w - 0.4, 0.7,
            font=_FONT, size=13, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, x + 0.2, content_y + 1.05, col_w - 0.4, H - content_y - 2.2,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS / PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول")

    content_y = 3.1
    avail_h = H - content_y - 0.6
    max_ch = 8
    chapters = req.chapters[:max_ch]
    ch_h = avail_h / max(len(chapters), 1)
    row_h = min(ch_h, 1.5) - 0.12

    for i, ch in enumerate(chapters):
        y = content_y + i * (row_h + 0.12)
        row_bg = rrect(slide, 1.2, y, W - 2.4, row_h, T.card_rgb, radius_pct=8)
        if row_bg:
            shadow(row_bg, blur=8, dist=2, alpha=0.25)

        # Number badge
        num_badge = oval(slide, 1.5, y + (row_h - 0.7) / 2, 0.7, 0.7, T.accent_rgb)
        if num_badge:
            gradient_fill(num_badge, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i + 1), 1.5, y + (row_h - 0.7) / 2, 0.7, 0.7,
            font="Calibri", size=11, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # Chapter title
        txt(slide, ch.title, 2.5, y, W - 5.8, row_h,
            font=_FONT, size=13, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        # Pages if provided
        if ch.pages:
            txt(slide, ch.pages, W - 3.8, y, 2.4, row_h,
                font="Calibri", size=10, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")

    content_y = 3.1
    cur_y = content_y

    if req.main_problem:
        card = rrect(slide, 1.2, cur_y, W - 2.4, 2.0, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=12, dist=4, alpha=0.35)
        lbl = rrect(slide, 1.2, cur_y, 3.8, 0.55, T.accent_rgb, radius_pct=0)
        if lbl:
            gradient_fill(lbl, T.accent_grad1, T.accent_grad2, angle=0)
        txt(slide, "الإشكالية الرئيسية", 1.3, cur_y, 3.6, 0.55,
            font=_FONT, size=11, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem, 1.4, cur_y + 0.62, W - 2.8, 1.3,
            font=_FONT, size=12, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cur_y += 2.15

    if req.main_question:
        q_card = rrect(slide, 1.2, cur_y, W - 2.4, 1.5, T.bg2_rgb, radius_pct=8)
        if q_card:
            shadow(q_card, blur=8, dist=3, alpha=0.2)
        hline(slide, 1.2, cur_y, 0.3, T.accent_rgb, thickness=1.5)
        txt(slide, req.main_question, 1.7, cur_y, W - 3.2, 1.5,
            font=_FONT, size=12.5, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cur_y += 1.65

    if req.sub_questions:
        avail = H - cur_y - 0.5
        sub_h = min(avail / max(len(req.sub_questions), 1), 0.9)
        for i, q in enumerate(req.sub_questions[:6]):
            y = cur_y + i * sub_h
            dot = oval(slide, W - 2.6, y + sub_h * 0.3, 0.28, 0.28, T.accent_rgb)
            txt(slide, q, 1.2, y, W - 3.2, sub_h,
                font=_FONT, size=11, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهداف البحث وفرضياته", "")

    content_y = 3.1
    col_w = (W - 2.8) / 2
    gap = 0.3

    cols = []
    if req.objectives:
        cols.append(("الأهداف", req.objectives))
    if req.hypotheses:
        cols.append(("الفرضيات", req.hypotheses))

    for i, (lbl, items) in enumerate(cols[:2]):
        x = 1.2 + i * (col_w + gap)
        card = rrect(slide, x, content_y, col_w, H - content_y - 0.6, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=14, dist=4, alpha=0.35)
        # Header
        hdr = rrect(slide, x, content_y, col_w, 0.65, T.accent_rgb, radius_pct=0)
        if hdr:
            gradient_fill(hdr, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x + 0.15, content_y, col_w - 0.3, 0.65,
            font=_FONT, size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        # Items
        avail_h = H - content_y - 1.5
        item_h = min(avail_h / max(len(items), 1), 1.1)
        for j, item in enumerate(items[:7]):
            iy = content_y + 0.75 + j * item_h
            num = oval(slide, x + col_w - 0.9, iy + 0.05, 0.5, 0.5, T.bg_rgb)
            txt(slide, str(j + 1), x + col_w - 0.9, iy + 0.05, 0.5, 0.5,
                font="Calibri", size=9, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
            txt(slide, item, x + 0.2, iy, col_w - 1.2, item_h,
                font=_FONT, size=10.5, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهمية البحث", "")

    content_y = 3.1
    items = req.importance[:6] if req.importance else []
    if req.reasons and len(items) < 6:
        items.append(req.reasons)

    if not items:
        return slide

    cols = 2 if len(items) > 3 else 1
    col_w = (W - 2.4 - (cols - 1) * 0.3) / cols
    rows = (len(items) + cols - 1) // cols
    avail_h = H - content_y - 0.5
    card_h = min(avail_h / rows - 0.2, 2.0)

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.2 + col_idx * (col_w + 0.3)
        y = content_y + row_idx * (card_h + 0.2)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)
        # Left accent
        vline(slide, x, y, card_h, T.accent_rgb, thickness=0.2)

        txt(slide, item, x + 0.4, y + 0.12, col_w - 0.6, card_h - 0.24,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "منهجية البحث", "الإجراءات والأدوات")

    content_y = 3.1
    fields = []
    if req.methodology:
        fields.append(("المنهج", req.methodology))
    if req.sample_type:
        fields.append(("العينة", req.sample_type))
    if req.sample_size:
        fields.append(("حجم العينة", req.sample_size))
    if req.tool:
        fields.append(("الأداة", req.tool))

    cols = 2 if len(fields) > 2 else 1
    col_w = (W - 2.4 - (cols - 1) * 0.3) / cols
    rows = (len(fields) + cols - 1) // cols
    avail_h = H - content_y - 0.5
    card_h = min(avail_h / rows - 0.2, 2.2)

    for i, (lbl, val) in enumerate(fields[:4]):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.2 + col_idx * (col_w + 0.3)
        y = content_y + row_idx * (card_h + 0.2)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)

        lbl_card = rrect(slide, x, y, col_w, 0.55, T.accent_rgb, radius_pct=0)
        if lbl_card:
            gradient_fill(lbl_card, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x + 0.2, y, col_w - 0.4, 0.55,
            font=_FONT, size=12, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, x + 0.2, y + 0.62, col_w - 0.4, card_h - 0.75,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# KPI / STATS SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "الإحصاءات والأرقام الرئيسية", "")

    content_y = 3.1
    stats = req.stats[:6]
    n = len(stats)
    if n == 0:
        return slide

    cols = 3 if n >= 3 else n
    rows = (n + cols - 1) // cols
    gap = 0.3
    card_w = (W - 2.4 - (cols - 1) * gap) / cols
    avail_h = H - content_y - 0.5
    card_h = min(avail_h / rows - gap, 3.5)

    for i, stat in enumerate(stats):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.2 + col_idx * (card_w + gap)
        y = content_y + row_idx * (card_h + gap)

        # Card
        card = rrect(slide, x, y, card_w, card_h, T.card_rgb, radius_pct=12)
        if card:
            shadow(card, blur=14, dist=5, alpha=0.4)

        # Top accent band
        band = rrect(slide, x, y, card_w, 0.22, T.accent_rgb, radius_pct=0)
        if band:
            gradient_fill(band, T.accent_grad1, T.accent_grad2, 0)

        # Value (big)
        val_size = 32 if len(stat.value) <= 4 else 24 if len(stat.value) <= 8 else 18
        txt(slide, stat.value, x + 0.2, y + 0.4, card_w - 0.4, card_h * 0.52,
            font="Calibri", size=val_size, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # Unit
        if stat.unit:
            txt(slide, stat.unit, x + 0.2, y + card_h * 0.52 + 0.3,
                card_w - 0.4, 0.5,
                font=_FONT, size=10, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

        # Label
        txt(slide, stat.label, x + 0.2, y + card_h - 0.8, card_w - 0.4, 0.7,
            font=_FONT, size=11, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_results(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")

    content_y = 3.1
    results = req.main_results[:8]
    avail_h = H - content_y - 0.5
    item_h = min(avail_h / max(len(results), 1), 1.4) - 0.12

    for i, result in enumerate(results):
        y = content_y + i * (item_h + 0.12)
        row_bg = rrect(slide, 1.2, y, W - 2.4, item_h, T.card_rgb, radius_pct=8)
        if row_bg:
            shadow(row_bg, blur=6, dist=2, alpha=0.2)

        # Number badge
        badge = rrect(slide, W - 3.0, y + (item_h - 0.58) / 2, 1.5, 0.58,
                      T.accent_rgb, radius_pct=50)
        if badge:
            gradient_fill(badge, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i + 1), W - 3.0, y + (item_h - 0.58) / 2, 1.5, 0.58,
            font="Calibri", size=11, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, result, 1.5, y + 0.1, W - 5.0, item_h - 0.2,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خاتمة البحث", "الاستنتاج العام")

    # Large centered card
    card_y = 3.1
    card_h = H - card_y - 0.8
    card = rrect(slide, 1.5, card_y, W - 3.0, card_h, T.card_rgb, radius_pct=14)
    if card:
        shadow(card, blur=20, dist=6, alpha=0.4)

    # Top accent
    top_stripe = rrect(slide, 1.5, card_y, W - 3.0, 0.28, T.accent_rgb, radius_pct=0)
    if top_stripe:
        gradient_fill(top_stripe, T.accent_grad1, T.accent_grad2, 0)

    # Quote mark decoration
    txt(slide, "❝", 2.5, card_y + 0.4, 1.5, 1.5,
        font="Calibri", size=40, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

    txt(slide, req.general_conclusion,
        2.0, card_y + 1.0, W - 4.5, card_h - 1.5,
        font=_FONT, size=14, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "توصيات البحث", "")

    content_y = 3.1
    recs = req.recommendations[:8]
    avail_h = H - content_y - 0.5
    item_h = min(avail_h / max(len(recs), 1), 1.4) - 0.1

    for i, rec in enumerate(recs):
        y = content_y + i * (item_h + 0.1)
        row_bg = rrect(slide, 1.2, y, W - 2.4, item_h, T.card_rgb, radius_pct=8)
        if row_bg:
            shadow(row_bg, blur=6, dist=2, alpha=0.2)

        # Bullet accent
        dot = oval(slide, W - 2.4, y + (item_h - 0.32) / 2, 0.32, 0.32, T.accent_rgb)

        txt(slide, rec, 1.5, y + 0.08, W - 3.2, item_h - 0.16,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE WORK SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_future(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "آفاق البحث المستقبلية", "")

    content_y = 3.1
    items = req.future_work[:6]
    cols = 2 if len(items) > 3 else 1
    col_w = (W - 2.4 - (cols - 1) * 0.3) / cols
    rows = (len(items) + cols - 1) // cols
    avail_h = H - content_y - 0.5
    card_h = min(avail_h / rows - 0.2, 2.2)

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.2 + col_idx * (col_w + 0.3)
        y = content_y + row_idx * (card_h + 0.2)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)

        # Icon area
        icon_card = rrect(slide, x, y, 0.6, card_h, T.accent_rgb, radius_pct=0)
        if icon_card:
            gradient_fill(icon_card, T.accent_grad1, T.accent_grad2, 90)

        txt(slide, item, x + 0.75, y + 0.15, col_w - 0.95, card_h - 0.3,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_references(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "المراجع والمصادر", "")

    content_y = 3.1
    refs = req.references[:12]
    avail_h = H - content_y - 0.5
    item_h = max(min(avail_h / max(len(refs), 1) - 0.1, 1.1), 0.5)

    for i, ref in enumerate(refs):
        y = content_y + i * (item_h + 0.1)
        if y + item_h > H - 0.3:
            break
        # Subtle row bg (alternating)
        if i % 2 == 0:
            row_bg = rrect(slide, 1.2, y, W - 2.4, item_h, T.card_rgb, radius_pct=4)

        # Number
        txt(slide, f"[{i + 1}]", W - 2.8, y, 1.4, item_h,
            font="Calibri", size=9, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

        txt(slide, ref, 1.5, y + 0.04, W - 4.5, item_h - 0.08,
            font=_FONT, size=10, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# THANK YOU / FINAL SLIDE
# ══════════════════════════════════════════════════════════════════════
def make_final(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Full gradient background
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    # Decorative circles
    oval(slide, -3, -3, 12, 12, T.accent_rgb, alpha=6)
    oval(slide, W - 9, H - 9, 14, 14, T.accent_rgb, alpha=5)
    oval(slide, W - 5, -2, 8, 8, T.bg2_rgb, alpha=30)

    # Center card
    card_w, card_h = 22, 10
    card_x = (W - card_w) / 2
    card_y = (H - card_h) / 2
    card = rrect(slide, card_x, card_y, card_w, card_h, T.card_rgb, radius_pct=14)
    if card:
        shadow(card, blur=24, dist=8, alpha=0.45)

    # Top stripe
    top_s = rrect(slide, card_x, card_y, card_w, 0.35, T.accent_rgb, radius_pct=0)
    if top_s:
        gradient_fill(top_s, T.accent_grad1, T.accent_grad2, 0)

    # Thank you text
    txt(slide, "شكراً وتقديراً",
        card_x + 1.0, card_y + 0.6, card_w - 2.0, 2.5,
        font=_FONT, size=36, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Divider
    hline(slide, card_x + card_w * 0.2, card_y + 3.2, card_w * 0.6, T.accent_rgb, thickness=0.05)

    # Name
    txt(slide, req.student_name,
        card_x + 1.0, card_y + 3.5, card_w - 2.0, 1.2,
        font=_FONT, size=18, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Title (truncated)
    title_display = req.title_ar[:70] + ("..." if len(req.title_ar) > 70 else "")
    txt(slide, title_display,
        card_x + 1.0, card_y + 4.8, card_w - 2.0, 2.0,
        font=_FONT, size=12, bold=False, italic=True,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Bottom institution + year
    footer_parts = []
    if req.institution:
        footer_parts.append(req.institution)
    if req.year:
        footer_parts.append(req.year)
    if footer_parts:
        txt(slide, "  ·  ".join(footer_parts),
            card_x + 1.0, card_y + card_h - 1.0, card_w - 2.0, 0.8,
            font=_FONT, size=11, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Bottom bar
    bottom_bar = rect(slide, 0, H - 0.25, W, 0.25, T.accent_rgb)
    if bottom_bar:
        gradient_fill(bottom_bar, T.accent_grad1, T.accent_grad2, 0)

    return slide

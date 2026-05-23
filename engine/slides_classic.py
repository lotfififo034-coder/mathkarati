"""
Classic Engine Slides — مذكرتي Pro v17
تخطيط أكاديمي كلاسيكي: نظيف، هادئ، رسمي
مختلف بصرياً عن Canva (بطاقات) وPremium (شريط جانبي)
"""
from __future__ import annotations
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha, txt, blank_slide,
)
from engine.typography import TS, FONT_TITLE, FONT_BODY, FONT_NUM, FONT_EN, display_size, stat_size
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"

HEADER_H = 2.2      # ارتفاع شريط العنوان العلوي
FOOTER_H = 0.55     # ارتفاع الشريط السفلي
MARGIN_X = 1.8      # هامش أفقي


def set_font(font_name: str):
    global _FONT
    _FONT = font_name


# ══════════════════════════════════════════════════════════════════════
# HEADER — شريط علوي بسيط في كل الشرائح
# ══════════════════════════════════════════════════════════════════════
def _header(slide, T: Theme, title: str, page_num: int = 0):
    """شريط علوي أكاديمي: خلفية فاتحة + خط accent + عنوان"""
    bg(slide, T.bg_rgb)

    # خلفية الهيدر
    hdr_bg = rect(slide, 0, 0, W, HEADER_H, T.bg2_rgb)

    # خط accent سميك في الأسفل
    accent_line = rect(slide, 0, HEADER_H - 0.12, W, 0.12, T.accent_rgb)
    if accent_line:
        gradient_fill(accent_line, T.accent_grad1, T.accent_grad2, angle=0)

    # خط رفيع فوق accent
    rect(slide, 0, HEADER_H - 0.22, W, 0.06, T.muted_rgb)

    # عنوان الشريحة
    txt(slide, title,
        MARGIN_X, 0.3, W - MARGIN_X * 2, HEADER_H - 0.5,
        font=FONT_TITLE, size=TS.H1, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    # رقم الصفحة يسار
    if page_num > 0:
        txt(slide, str(page_num),
            0.3, 0.3, 1.2, HEADER_H - 0.5,
            font=FONT_NUM, size=TS.STAT_MD, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

    # شريط سفلي
    footer_bg = rect(slide, 0, H - FOOTER_H, W, FOOTER_H, T.bg2_rgb)
    footer_line = rect(slide, 0, H - FOOTER_H, W, 0.06, T.accent_rgb)
    if footer_line:
        gradient_fill(footer_line, T.accent_grad1, T.accent_grad2, 0)


def _content_y():
    return HEADER_H + 0.5


def _content_h():
    return H - HEADER_H - FOOTER_H - 1.0


# ══════════════════════════════════════════════════════════════════════
# COVER — Classic
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # طبقة خلفية بتدرج خفيف
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)

    # الإطار الخارجي — خطوط accent على الحواف
    rect(slide, 0, 0, W, 0.18, T.accent_rgb)                  # أعلى
    r_top = rect(slide, 0, 0, W, 0.18, T.accent_rgb)
    if r_top: gradient_fill(r_top, T.accent_grad1, T.accent_grad2, 0)

    r_bot = rect(slide, 0, H - 0.18, W, 0.18, T.accent_rgb)   # أسفل
    if r_bot: gradient_fill(r_bot, T.accent_grad1, T.accent_grad2, 0)

    vline(slide, 0.18, 0.18, H - 0.36, T.accent_rgb, thickness=0.06)     # يسار
    vline(slide, W - 0.24, 0.18, H - 0.36, T.accent_rgb, thickness=0.06) # يمين

    # المؤسسة
    if req.institution:
        txt(slide, req.institution,
            2.0, 0.9, W - 4.0, 0.9,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # خط فاصل تحت المؤسسة
    hline(slide, W * 0.2, 1.85, W * 0.6, T.accent_rgb, thickness=0.04)

    # العنوان الرئيسي — في وسط الشريحة
    title_y = H * 0.28
    title_size = 26 if len(req.title_ar) < 50 else 20 if len(req.title_ar) < 80 else 16

    txt(slide, req.title_ar,
        2.5, title_y, W - 5.0, H * 0.28,
        font=FONT_TITLE, size=title_size, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    if req.title_en:
        txt(slide, req.title_en,
            2.5, title_y + H * 0.27, W - 5.0, 0.9,
            font=FONT_EN, size=TS.BODY, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # خط فاصل تحت العنوان
    div_y = title_y + H * 0.32
    hline(slide, W * 0.15, div_y, W * 0.7, T.accent_rgb, thickness=0.06)

    # جدول المعلومات — شبكة نظيفة
    info_y = div_y + 0.45
    row_h = 0.72

    def info_row(label, value, y, is_bold=False):
        # خلفية الصف
        row_bg = rect(slide, MARGIN_X, y, W - MARGIN_X * 2, row_h - 0.06, T.bg2_rgb)
        # خط يميني accent
        vline(slide, W - MARGIN_X - 0.08, y, row_h - 0.06, T.accent_rgb, thickness=0.08)
        # التسمية
        txt(slide, label, MARGIN_X + 0.2, y, 4.5, row_h,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        # الفاصل
        vline(slide, W / 2, y + 0.1, row_h - 0.26, T.muted_rgb, thickness=0.04)
        # القيمة
        txt(slide, value, W / 2 + 0.3, y, W / 2 - MARGIN_X - 0.5, row_h,
            font=FONT_BODY, size=TS.BODY_LG, bold=is_bold,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    info_row("اسم الطالب", req.student_name, info_y, is_bold=True)
    r = 1
    if req.supervisor:
        info_row("المشرف", req.supervisor, info_y + row_h * r); r += 1
    if req.co_supervisor:
        info_row("المشرف المساعد", req.co_supervisor, info_y + row_h * r); r += 1
    if req.specialization:
        info_row("التخصص", req.specialization, info_y + row_h * r); r += 1
    if req.year:
        info_row("السنة الجامعية", req.year, info_y + row_h * r)

    return slide


# ══════════════════════════════════════════════════════════════════════
# SECTION SLIDES — كلاسيك: جدول/قائمة بأرقام + خطوط فاصلة
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "مقدمة البحث", 1)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()

    items = []
    if req.intro_overview:  items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("المنهج المتبع", req.intro_approach))

    avail_h = _content_h()
    card_h = avail_h / max(len(items), 1) - 0.25

    for i, (lbl, val) in enumerate(items[:2]):
        y = cy + i * (card_h + 0.25)
        # خط accent يميني
        vline(slide, W - MARGIN_X - 0.1, y, card_h, T.accent_rgb, thickness=0.1)
        # تسمية
        txt(slide, lbl, cx, y, cw - 0.4, 0.65,
            font=FONT_TITLE, size=TS.H2, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        # خط تحت التسمية
        hline(slide, cx, y + 0.7, cw - 0.4, T.muted_rgb, thickness=0.03)
        # المحتوى
        txt(slide, val, cx, y + 0.8, cw - 0.4, card_h - 0.9,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "خطة البحث", 2)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    chapters = req.chapters[:8]
    avail_h = _content_h()
    row_h = min(avail_h / max(len(chapters), 1) - 0.1, 1.8)

    for i, ch in enumerate(chapters):
        y = cy + i * (row_h + 0.1)

        # خلفية متناوبة
        if i % 2 == 0:
            rect(slide, cx, y, cw, row_h, T.bg2_rgb)
        else:
            rect(slide, cx, y, cw, row_h, T.card_rgb)

        # خط accent يميني
        vline(slide, W - MARGIN_X - 0.1, y, row_h, T.accent_rgb, thickness=0.1)

        # رقم الفصل
        txt(slide, f"الفصل {i + 1}", cx, y, 3.0, row_h,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        # خط فاصل عمودي
        vline(slide, cx + 3.2, y + 0.1, row_h - 0.2, T.muted_rgb, thickness=0.04)

        # عنوان الفصل
        txt(slide, ch.title, cx + 3.4, y, cw - 4.0, row_h,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        # الصفحات
        if ch.pages:
            txt(slide, ch.pages, cx, y, 1.5, row_h,
                font=FONT_NUM, size=TS.LABEL, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)

        # خط فاصل أفقي سفلي
        hline(slide, cx, y + row_h, cw, T.bg_rgb, thickness=0.1)

    return slide


def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "إشكالية البحث", 3)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()

    if req.main_problem:
        # بطاقة الإشكالية
        vline(slide, W - MARGIN_X - 0.14, cy, 2.6, T.accent_rgb, thickness=0.14)
        txt(slide, "الإشكالية الرئيسية", cx, cy, cw - 0.3, 0.65,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        hline(slide, cx, cy + 0.68, cw - 0.3, T.muted_rgb, thickness=0.03)
        txt(slide, req.main_problem, cx, cy + 0.78, cw - 0.3, 1.7,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 2.75

    if req.main_question:
        hline(slide, cx, cy, cw, T.accent_rgb, thickness=0.06)
        cy += 0.15
        txt(slide, "التساؤل الرئيسي", cx, cy, cw, 0.6,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_question, cx, cy + 0.65, cw, 1.3,
            font=FONT_BODY, size=TS.BODY_LG, bold=False, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 2.1

    if req.sub_questions:
        hline(slide, cx, cy, cw, T.muted_rgb, thickness=0.03)
        cy += 0.2
        txt(slide, "التساؤلات الفرعية", cx, cy, cw, 0.55,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 0.6
        avail = H - cy - FOOTER_H - 0.5
        sub_h = min(avail / max(len(req.sub_questions), 1), 0.85)
        for i, q in enumerate(req.sub_questions[:6]):
            y = cy + i * sub_h
            txt(slide, f"{'─'} {q}", cx, y, cw, sub_h,
                font=FONT_BODY, size=TS.BODY, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "أهداف البحث وفرضياته", 4)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()

    cols_data = []
    if req.objectives:  cols_data.append(("الأهداف", req.objectives))
    if req.hypotheses:  cols_data.append(("الفرضيات", req.hypotheses))
    if not cols_data:   return slide

    col_w = (cw - 0.4 * (len(cols_data) - 1)) / len(cols_data)

    for i, (lbl, items) in enumerate(cols_data[:2]):
        x = cx + i * (col_w + 0.4)

        # عنوان العمود
        rect(slide, x, cy, col_w, 0.65, T.bg2_rgb)
        vline(slide, x + col_w - 0.1, cy, 0.65, T.accent_rgb, thickness=0.1)
        txt(slide, lbl, x, cy, col_w - 0.2, 0.65,
            font=FONT_TITLE, size=TS.H2, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        hline(slide, x, cy + 0.65, col_w, T.accent_rgb, thickness=0.06)

        avail_h = _content_h() - 0.8
        item_h = min(avail_h / max(len(items), 1), 1.2) - 0.08

        for j, item in enumerate(items[:8]):
            iy = cy + 0.78 + j * (item_h + 0.08)

            # خلفية متناوبة
            fill = T.bg2_rgb if j % 2 == 0 else T.card_rgb
            rect(slide, x, iy, col_w, item_h, fill)

            # رقم
            txt(slide, str(j + 1), x, iy, 0.7, item_h,
                font=FONT_NUM, size=TS.H3, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
            vline(slide, x + 0.75, iy + 0.05, item_h - 0.1, T.muted_rgb, thickness=0.03)

            txt(slide, item, x + 0.85, iy + 0.08, col_w - 1.1, item_h - 0.16,
                font=FONT_BODY, size=TS.BODY_SM, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

            hline(slide, x, iy + item_h, col_w, T.bg_rgb, thickness=0.08)

    return slide


def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "أهمية البحث", 5)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    items = req.importance[:6]
    avail_h = _content_h()
    item_h = min(avail_h / max(len(items), 1) - 0.12, 2.0)

    for i, item in enumerate(items):
        y = cy + i * (item_h + 0.12)

        # خلفية
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, item_h, fill)

        # خط accent يميني
        vline(slide, W - MARGIN_X - 0.12, y, item_h, T.accent_rgb, thickness=0.12)

        # رقم كبير
        txt(slide, f"{i + 1:02d}", cx, y, 1.5, item_h,
            font=FONT_NUM, size=TS.STAT_SM, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        vline(slide, cx + 1.6, y + 0.1, item_h - 0.2, T.muted_rgb, thickness=0.04)

        txt(slide, item, cx + 1.8, y + 0.1, cw - 2.2, item_h - 0.2,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "منهجية البحث", 6)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()

    fields = []
    if req.methodology:  fields.append(("المنهج المتبع", req.methodology))
    if req.sample_type:  fields.append(("نوع العينة", req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("أداة الدراسة", req.tool))

    avail_h = _content_h()
    row_h = min(avail_h / max(len(fields), 1) - 0.15, 2.5)

    for i, (lbl, val) in enumerate(fields[:4]):
        y = cy + i * (row_h + 0.15)

        # الصف
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, row_h, fill)

        # خط accent يميني
        vline(slide, W - MARGIN_X - 0.12, y, row_h, T.accent_rgb, thickness=0.12)

        # التسمية
        rect(slide, cx, y, 5.0, row_h, T.bg2_rgb if i % 2 != 0 else T.card_rgb)
        txt(slide, lbl, cx, y, 4.8, row_h,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        # فاصل
        vline(slide, cx + 5.1, y + 0.1, row_h - 0.2, T.muted_rgb, thickness=0.04)

        # القيمة
        txt(slide, val, cx + 5.3, y + 0.1, cw - 5.7, row_h - 0.2,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "الإحصاءات والأرقام", 7)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    stats = req.stats[:6]
    n = len(stats)
    if n == 0:
        return slide

    cols = 3 if n >= 3 else n
    rows = (n + cols - 1) // cols
    gap = 0.35
    card_w = (cw - gap * (cols - 1)) / cols
    avail_h = _content_h()
    card_h = min(avail_h / rows - gap, 4.0)

    for i, stat in enumerate(stats):
        col_idx = i % cols
        row_idx = i // cols
        x = cx + col_idx * (card_w + gap)
        y = cy + row_idx * (card_h + gap)

        # إطار البطاقة
        card_bg = rect(slide, x, y, card_w, card_h, T.bg2_rgb)

        # خط accent أعلى
        top = rect(slide, x, y, card_w, 0.12, T.accent_rgb)
        if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

        # خط accent يميني
        vline(slide, x + card_w - 0.1, y, card_h, T.accent_rgb, thickness=0.1)

        # القيمة الرئيسية
        val_size = 34 if len(stat.value) <= 5 else 24 if len(stat.value) <= 9 else 18
        txt(slide, stat.value, x + 0.2, y + 0.3, card_w - 0.5, card_h * 0.52,
            font=FONT_NUM, size=val_size, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # خط فاصل
        hline(slide, x + 0.3, y + card_h * 0.56, card_w - 0.6,
              T.muted_rgb, thickness=0.04)

        if stat.unit:
            txt(slide, stat.unit, x + 0.2, y + card_h * 0.58, card_w - 0.5, 0.5,
                font=FONT_BODY, size=TS.LABEL, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

        txt(slide, stat.label, x + 0.2, y + card_h - 0.9, card_w - 0.5, 0.75,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "نتائج البحث", 8)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    results = req.main_results[:8]
    avail_h = _content_h()
    item_h = min(avail_h / max(len(results), 1) - 0.1, 1.5)

    for i, result in enumerate(results):
        y = cy + i * (item_h + 0.1)

        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, item_h, fill)

        # خط accent يميني
        vline(slide, W - MARGIN_X - 0.12, y, item_h, T.accent_rgb, thickness=0.12)

        # رقم
        txt(slide, str(i + 1), cx, y, 0.9, item_h,
            font=FONT_NUM, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        vline(slide, cx + 1.0, y + 0.08, item_h - 0.16, T.muted_rgb, thickness=0.04)

        txt(slide, result, cx + 1.1, y + 0.08, cw - 1.5, item_h - 0.16,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "خاتمة البحث", 9)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    avail_h = _content_h()

    # إطار كامل
    rect(slide, cx, cy, cw, avail_h, T.bg2_rgb)
    vline(slide, W - MARGIN_X - 0.14, cy, avail_h, T.accent_rgb, thickness=0.14)
    vline(slide, cx, cy, avail_h, T.bg2_rgb, thickness=0.14)

    # عنوان فرعي
    txt(slide, "الاستنتاج العام", cx, cy + 0.2, cw - 0.3, 0.7,
        font=FONT_TITLE, size=TS.H3, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    hline(slide, cx, cy + 0.95, cw - 0.3, T.accent_rgb, thickness=0.06)

    txt(slide, req.general_conclusion,
        cx, cy + 1.1, cw - 0.3, avail_h - 1.3,
        font=FONT_BODY, size=TS.BODY_LG, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "توصيات البحث", 10)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    recs = req.recommendations[:8]
    avail_h = _content_h()
    item_h = min(avail_h / max(len(recs), 1) - 0.1, 1.4)

    for i, rec in enumerate(recs):
        y = cy + i * (item_h + 0.1)
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, item_h, fill)

        vline(slide, W - MARGIN_X - 0.12, y, item_h, T.accent_rgb, thickness=0.12)

        # نقطة
        oval(slide, cx + 0.25, y + (item_h - 0.38) / 2,
             0.38, 0.38, T.accent_rgb)

        txt(slide, rec, cx + 0.8, y + 0.08, cw - 1.2, item_h - 0.16,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "آفاق البحث المستقبلية", 11)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    items = req.future_work[:6]
    avail_h = _content_h()
    item_h = min(avail_h / max(len(items), 1) - 0.12, 2.0)

    for i, item in enumerate(items):
        y = cy + i * (item_h + 0.12)
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, item_h, fill)

        vline(slide, W - MARGIN_X - 0.12, y, item_h, T.accent_rgb, thickness=0.12)

        # رقم + فاصل
        txt(slide, str(i + 1), cx, y, 0.9, item_h,
            font=FONT_NUM, size=TS.H1, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
        vline(slide, cx + 1.0, y + 0.1, item_h - 0.2, T.muted_rgb, thickness=0.04)

        txt(slide, item, cx + 1.2, y + 0.1, cw - 1.6, item_h - 0.2,
            font=FONT_BODY, size=TS.BODY_LG, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "قائمة المراجع", 12)

    cx = MARGIN_X
    cw = W - MARGIN_X * 2
    cy = _content_y()
    refs = req.references[:12]
    avail_h = _content_h()
    item_h = max(min(avail_h / max(len(refs), 1) - 0.08, 1.1), 0.52)

    for i, ref in enumerate(refs):
        y = cy + i * (item_h + 0.08)
        if y + item_h > H - FOOTER_H - 0.2:
            break

        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, item_h, fill)

        vline(slide, W - MARGIN_X - 0.1, y, item_h, T.accent_rgb, thickness=0.1)

        txt(slide, f"[{i + 1}]", cx, y, 1.0, item_h,
            font=FONT_NUM, size=TS.LABEL, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        vline(slide, cx + 1.1, y + 0.05, item_h - 0.1, T.muted_rgb, thickness=0.03)

        txt(slide, ref, cx + 1.2, y + 0.05, cw - 1.5, item_h - 0.1,
            font=FONT_BODY, size=TS.LABEL, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)

    # إطار خارجي
    r_top = rect(slide, 0, 0, W, 0.2, T.accent_rgb)
    if r_top: gradient_fill(r_top, T.accent_grad1, T.accent_grad2, 0)
    r_bot = rect(slide, 0, H - 0.2, W, 0.2, T.accent_rgb)
    if r_bot: gradient_fill(r_bot, T.accent_grad1, T.accent_grad2, 0)
    vline(slide, 0.2, 0.2, H - 0.4, T.accent_rgb, thickness=0.06)
    vline(slide, W - 0.26, 0.2, H - 0.4, T.accent_rgb, thickness=0.06)

    center_x = W / 2
    # شكر وتقدير
    txt(slide, "شكراً وتقديراً",
        2.0, H * 0.22, W - 4.0, 2.5,
        font=FONT_TITLE, size=TS.H1, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # خط مزدوج
    hline(slide, W * 0.15, H * 0.46, W * 0.7, T.accent_rgb, thickness=0.06)
    hline(slide, W * 0.2,  H * 0.47 + 0.12, W * 0.6, T.muted_rgb, thickness=0.03)

    # اسم الطالب
    txt(slide, req.student_name,
        2.0, H * 0.5, W - 4.0, 1.3,
        font=FONT_TITLE, size=TS.H1, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # عنوان المذكرة
    title_short = req.title_ar[:70] + ("..." if len(req.title_ar) > 70 else "")
    txt(slide, title_short,
        2.5, H * 0.62, W - 5.0, 1.8,
        font=FONT_BODY, size=TS.BODY_LG, bold=False, italic=True,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # المؤسسة والسنة
    footer = []
    if req.institution: footer.append(req.institution)
    if req.year:        footer.append(req.year)
    if footer:
        txt(slide, " · ".join(footer),
            2.0, H * 0.8, W - 4.0, 0.8,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide

"""
Premium Engine Slides — مذكرتي Pro v17.2
تخطيط سينمائي: شريط جانبي عريض + محتوى يميني
مختلف بصرياً تماماً عن Canva (بطاقات) وClassic (أكاديمي)

v17.2: Typography System — هرمية الخطوط المتسقة
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
SIDEBAR_W = 7.5   # عرض الشريط الجانبي الأيسر


def set_font(font_name: str):
    global _FONT
    _FONT = font_name


def _hx(h):
    from pptx.dml.color import RGBColor
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ══════════════════════════════════════════════════════════════════════
# SIDEBAR LAYOUT: الشريط الجانبي الأيسر ثابت في كل الشرائح
# ══════════════════════════════════════════════════════════════════════
def _sidebar(slide, T: Theme, icon: str, section_title: str):
    """شريط جانبي أيسر بتدرج + أيقونة + عنوان القسم"""
    # الخلفية الكاملة
    bg(slide, T.bg_rgb)

    # الشريط الجانبي
    sidebar = gradient_rect(slide, 0, 0, SIDEBAR_W, H, T.grad1, T.grad2, angle=180)

    # خط فاصل بين الشريط والمحتوى
    sep = rect(slide, SIDEBAR_W, 0, 0.06, H, T.accent_rgb)

    # دائرة أيقونة في المنتصف الجانبي
    icon_x = SIDEBAR_W / 2 - 1.8
    icon_y = H / 2 - 2.2
    icon_circle = oval(slide, icon_x, icon_y, 3.6, 3.6, T.accent_rgb, alpha=15)
    txt(slide, icon,
        icon_x, icon_y + 0.6, 3.6, 2.4,
        font=FONT_EN, size=TS.SIDEBAR_ICON, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # عنوان القسم (عمودي في الشريط) — H2 size واضح ومميز
    txt(slide, section_title,
        0.3, H / 2 + 1.8, SIDEBAR_W - 0.6, 2.5,
        font=FONT_TITLE, size=TS.SIDEBAR_LABEL, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # زخارف سفلى وعلوية في الشريط
    oval(slide, -1.5, -1.5, 5, 5, T.accent_rgb, alpha=6)
    oval(slide, 1, H - 4, 4, 4, T.bg2_rgb, alpha=40)

    # خلفية منطقة المحتوى
    rect(slide, SIDEBAR_W + 0.06, 0, W - SIDEBAR_W - 0.06, H, T.bg2_rgb)


def _content_area_x():
    return SIDEBAR_W + 0.5


def _content_w():
    return W - SIDEBAR_W - 1.0


# ══════════════════════════════════════════════════════════════════════
# COVER — Premium
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # خلفية بتدرج كامل
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=120)

    # شريط جانبي أيسر واسع
    sidebar = gradient_rect(slide, 0, 0, SIDEBAR_W + 1, H, T.grad2, T.grad1, angle=90)

    # خط فاصل ذهبي
    sep = rect(slide, SIDEBAR_W + 1, 0, 0.1, H, T.accent_rgb)
    if sep:
        gradient_fill(sep, T.accent_grad1, T.accent_grad2, angle=90)

    # شعار/أيقونة مذكرتي في الشريط
    logo_y = H * 0.15
    oval(slide, 1.5, logo_y, 4.5, 4.5, T.accent_rgb, alpha=20)
    txt(slide, "🎓", 1.5, logo_y + 0.8, 4.5, 3.0,
        font=FONT_EN, size=52, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # اسم المؤسسة في الشريط — CAPTION حجم ثانوي واضح
    if req.institution:
        txt(slide, req.institution,
            0.3, H * 0.62, SIDEBAR_W + 0.4, 1.2,
            font=FONT_BODY, size=TS.COVER_INST, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # السنة — رقم بارز في الـ badge
    if req.year:
        yr = rrect(slide, 0.8, H - 1.4, SIDEBAR_W - 0.6, 0.65,
                   T.accent_rgb, radius_pct=50)
        if yr:
            gradient_fill(yr, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, req.year, 0.8, H - 1.4, SIDEBAR_W - 0.6, 0.65,
            font=FONT_NUM, size=TS.COVER_YEAR, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # منطقة المحتوى اليمنى
    cx = SIDEBAR_W + 1.5
    cw = W - cx - 0.8

    # العنوان الرئيسي — DISPLAY scale حسب طول النص
    title_y = H * 0.18
    title_size = display_size(req.title_ar)

    # خلفية شفافة للعنوان
    title_bg = rrect(slide, cx - 0.3, title_y - 0.3, cw + 0.3, H * 0.35,
                     T.card_rgb, radius_pct=10)
    if title_bg:
        shadow(title_bg, blur=20, dist=6, alpha=0.4)

    # شريط accent أعلى البطاقة
    top_s = rrect(slide, cx - 0.3, title_y - 0.3, cw + 0.3, 0.28,
                  T.accent_rgb, radius_pct=0)
    if top_s:
        gradient_fill(top_s, T.accent_grad1, T.accent_grad2, 0)

    # العنوان الرئيسي — أكبر خط في العرض
    txt(slide, req.title_ar,
        cx, title_y, cw, H * 0.32,
        font=FONT_TITLE, size=title_size, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # العنوان بالإنجليزية — ثانوي وأصغر بوضوح
    if req.title_en:
        txt(slide, req.title_en,
            cx, title_y + H * 0.30, cw, 0.8,
            font=FONT_EN, size=TS.BODY_SM, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # خط فاصل
    div_y = title_y + H * 0.38
    hline(slide, cx + cw * 0.1, div_y, cw * 0.8, T.accent_rgb, thickness=0.05)

    # معلومات الطالب والمشرف — تسمية أصغر + قيمة أوضح
    info_y = div_y + 0.4
    row_h = 0.7

    def info_row(label, val, y):
        # التسمية (الطالب:) — LABEL صغير ملون
        txt(slide, label, cx, y, 3.5, row_h,
            font=FONT_TITLE, size=TS.COVER_LABEL, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        # القيمة (الاسم) — BODY_LG أوضح وأكبر
        txt(slide, val, cx + 3.8, y, cw - 3.8, row_h,
            font=FONT_BODY, size=TS.COVER_VALUE, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    info_row("الطالب :", req.student_name, info_y)
    if req.supervisor:
        info_row("المشرف :", req.supervisor, info_y + row_h)
    if req.co_supervisor:
        info_row("المشرف المساعد :", req.co_supervisor, info_y + row_h * 2)

    return slide


# ══════════════════════════════════════════════════════════════════════
# SECTION SLIDES — كلها تستخدم _sidebar
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "📖", "مقدمة\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 1.2

    items = []
    if req.intro_overview:
        items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:
        items.append(("المنهج المتبع", req.intro_approach))

    avail_h = H - cy - 0.8
    card_h = avail_h / max(len(items), 1) - 0.3

    for i, (lbl, val) in enumerate(items[:2]):
        y = cy + i * (card_h + 0.3)
        card = rrect(slide, cx, y, cw, card_h, T.card_rgb, radius_pct=8)
        if card:
            shadow(card, blur=12, dist=4, alpha=0.3)

        # الخط الجانبي الأيمن
        accent_v = rect(slide, cx + cw - 0.25, y, 0.25, card_h, T.accent_rgb)
        if accent_v:
            gradient_fill(accent_v, T.accent_grad1, T.accent_grad2, 90)

        txt(slide, lbl, cx + 0.3, y + 0.2, cw - 0.8, 0.65,
            font=FONT_TITLE, size=TS.H2, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, cx + 0.3, y + 0.95, cw - 0.8, card_h - 1.1,
            font=FONT_BODY, size=TS.BODY,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "📋", "خطة\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 1.0
    chapters = req.chapters[:8]
    avail_h = H - cy - 0.6
    ch_h = avail_h / max(len(chapters), 1)
    row_h = min(ch_h, 1.6) - 0.15

    for i, ch in enumerate(chapters):
        y = cy + i * (row_h + 0.15)

        # خلفية الصف - alternating
        fill = T.card_rgb if i % 2 == 0 else T.bg2_rgb
        row_bg = rrect(slide, cx, y, cw, row_h, fill, radius_pct=6)

        # شريط الرقم على اليسار من المحتوى
        num_strip = rrect(slide, cx, y, 1.2, row_h, T.accent_rgb, radius_pct=0)
        if num_strip:
            gradient_fill(num_strip, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, str(i + 1), cx, y, 1.2, row_h,
            font=FONT_NUM, size=TS.STAT_SM, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, ch.title, cx + 1.4, y, cw - 2.0, row_h,
            font=FONT_BODY, size=TS.H3, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        if ch.pages:
            txt(slide, ch.pages, cx + cw - 2.0, y, 1.8, row_h,
                font=FONT_NUM, size=TS.LABEL, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)

    return slide


def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "❓", "إشكالية\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8

    if req.main_problem:
        # بطاقة الإشكالية الكبيرة
        card = rrect(slide, cx, cy, cw, 2.8, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=16, dist=5, alpha=0.4)

        # شريط أعلى ملوّن
        top = rrect(slide, cx, cy, cw, 0.3, T.accent_rgb, radius_pct=0)
        if top:
            gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

        txt(slide, "الإشكالية الرئيسية", cx + 0.3, cy + 0.35, cw - 0.6, 0.65,
            font=FONT_TITLE, size=TS.H2, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem, cx + 0.3, cy + 1.05, cw - 0.6, 1.55,
            font=FONT_BODY, size=TS.BODY_LG, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        cy += 3.1

    if req.main_question:
        q_bg = rrect(slide, cx, cy, cw, 1.6, T.bg_rgb, radius_pct=8)
        # خط accent يميني
        vline(slide, cx + cw - 0.2, cy, 1.6, T.accent_rgb, thickness=0.2)
        txt(slide, req.main_question, cx + 0.3, cy, cw - 0.7, 1.6,
            font=FONT_BODY, size=TS.BODY_LG, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 1.8

    if req.sub_questions:
        avail = H - cy - 0.4
        sub_h = min(avail / max(len(req.sub_questions), 1), 1.0)
        for i, q in enumerate(req.sub_questions[:5]):
            y = cy + i * sub_h
            # رقم صغير
            num_c = rrect(slide, cx + cw - 0.9, y + 0.1, 0.65, 0.65,
                          T.accent_rgb, radius_pct=50)
            if num_c:
                gradient_fill(num_c, T.accent_grad1, T.accent_grad2, 0)
            txt(slide, str(i + 1), cx + cw - 0.9, y + 0.1, 0.65, 0.65,
                font=FONT_NUM, size=TS.LABEL, bold=True,
                color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)
            txt(slide, q, cx + 0.3, y, cw - 1.5, sub_h,
                font=FONT_BODY, size=TS.BODY, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "🎯", "أهداف\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8

    cols_data = []
    if req.objectives:
        cols_data.append(("الأهداف", req.objectives))
    if req.hypotheses:
        cols_data.append(("الفرضيات", req.hypotheses))

    if not cols_data:
        return slide

    col_w = (cw - 0.3 * (len(cols_data) - 1)) / len(cols_data)

    for i, (lbl, items) in enumerate(cols_data[:2]):
        x = cx + i * (col_w + 0.3)

        # عنوان العمود بتدرج
        hdr = rrect(slide, x, cy, col_w, 0.75, T.accent_rgb, radius_pct=8)
        if hdr:
            gradient_fill(hdr, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x + 0.2, cy, col_w - 0.4, 0.75,
            font=FONT_TITLE, size=TS.H2, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        avail_h = H - cy - 1.1
        item_h = min(avail_h / max(len(items), 1), 1.3) - 0.1

        for j, item in enumerate(items[:8]):
            iy = cy + 0.85 + j * (item_h + 0.1)
            row = rrect(slide, x, iy, col_w, item_h, T.card_rgb, radius_pct=6)

            # رقم يساري
            num_bg = rect(slide, x, iy, 0.55, item_h, T.bg_rgb)
            txt(slide, str(j + 1), x, iy, 0.55, item_h,
                font=FONT_NUM, size=TS.LABEL, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

            txt(slide, item, x + 0.65, iy + 0.08, col_w - 0.85, item_h - 0.16,
                font=FONT_BODY, size=TS.BODY_SM, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "⭐", "أهمية\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    items = req.importance[:6]

    avail_h = H - cy - 0.5
    item_h = min(avail_h / max(len(items), 1), 1.8) - 0.2

    for i, item in enumerate(items):
        y = cy + i * (item_h + 0.2)
        card = rrect(slide, cx, y, cw, item_h, T.card_rgb, radius_pct=8)
        if card:
            shadow(card, blur=8, dist=3, alpha=0.25)

        # شريط أيمن ملوّن
        v = rect(slide, cx + cw - 0.28, y, 0.28, item_h, T.accent_rgb)
        if v:
            gradient_fill(v, T.accent_grad1, T.accent_grad2, 90)

        # رقم كبير شفاف خلفية
        txt(slide, str(i + 1), cx + 0.1, y, 1.2, item_h,
            font=FONT_NUM, size=TS.STAT_SM, bold=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, item, cx + 1.4, y + 0.1, cw - 2.0, item_h - 0.2,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "🔬", "منهجية\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    fields = []
    if req.methodology: fields.append(("المنهج", req.methodology, "📊"))
    if req.sample_type: fields.append(("نوع العينة", req.sample_type, "👥"))
    if req.sample_size: fields.append(("حجم العينة", req.sample_size, "📏"))
    if req.tool:        fields.append(("أداة البحث", req.tool, "🛠️"))

    cols = 2 if len(fields) > 2 else 1
    col_w = (cw - 0.3 * (cols - 1)) / cols
    rows = (len(fields) + cols - 1) // cols
    avail_h = H - cy - 0.5
    card_h = min(avail_h / rows - 0.2, 3.5)

    for i, (lbl, val, icon) in enumerate(fields[:4]):
        col_idx = i % cols
        row_idx = i // cols
        x = cx + col_idx * (col_w + 0.3)
        y = cy + row_idx * (card_h + 0.2)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=12, dist=4, alpha=0.3)

        # أيقونة دائرة
        ic_c = oval(slide, x + col_w / 2 - 0.75, y + 0.2, 1.5, 1.5, T.accent_rgb, alpha=20)
        txt(slide, icon, x + col_w / 2 - 0.75, y + 0.25, 1.5, 1.2,
            font=FONT_NUM, size=TS.STAT_SM, bold=False,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, lbl, x + 0.2, y + 1.8, col_w - 0.4, 0.65,
            font=FONT_TITLE, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

        txt(slide, val, x + 0.2, y + 2.5, col_w - 0.4, card_h - 2.7,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "📈", "الأرقام\nالرئيسية")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    stats = req.stats[:6]
    n = len(stats)
    if n == 0:
        return slide

    cols = 3 if n >= 3 else n
    rows = (n + cols - 1) // cols
    gap = 0.3
    card_w = (cw - gap * (cols - 1)) / cols
    avail_h = H - cy - 0.5
    card_h = min(avail_h / rows - gap, 4.0)

    for i, stat in enumerate(stats):
        col_idx = i % cols
        row_idx = i // cols
        x = cx + col_idx * (card_w + gap)
        y = cy + row_idx * (card_h + gap)

        card = rrect(slide, x, y, card_w, card_h, T.card_rgb, radius_pct=12)
        if card:
            shadow(card, blur=14, dist=5, alpha=0.4)

        # Gradient fill على الكارد
        gradient_fill(card, T.bg, T.card, angle=135)

        # شريط سفلي accent
        bottom = rrect(slide, x, y + card_h - 0.25, card_w, 0.25,
                       T.accent_rgb, radius_pct=0)
        if bottom:
            gradient_fill(bottom, T.accent_grad1, T.accent_grad2, 0)

        # القيمة الكبيرة
        val_size = 36 if len(stat.value) <= 4 else 26 if len(stat.value) <= 8 else 20
        txt(slide, stat.value, x + 0.2, y + 0.4, card_w - 0.4, card_h * 0.5,
            font=FONT_NUM, size=val_size, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        if stat.unit:
            txt(slide, stat.unit, x + 0.2, y + card_h * 0.5 + 0.2,
                card_w - 0.4, 0.55,
                font=FONT_BODY, size=TS.LABEL, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

        txt(slide, stat.label, x + 0.2, y + card_h - 0.85, card_w - 0.4, 0.7,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "📊", "نتائج\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    results = req.main_results[:8]
    avail_h = H - cy - 0.5
    item_h = min(avail_h / max(len(results), 1), 1.5) - 0.15

    for i, result in enumerate(results):
        y = cy + i * (item_h + 0.15)

        # خلفية الصف
        row_bg = rrect(slide, cx, y, cw, item_h, T.card_rgb, radius_pct=6)
        if row_bg and i % 2 == 1:
            gradient_fill(row_bg, T.bg, T.card, angle=0)

        # شريط يميني ملوّن بحجم متناسب مع الترتيب
        intensity = max(3, 10 - i)
        accent_strip = rrect(slide, cx + cw - 0.35, y, 0.35, item_h,
                             T.accent_rgb, radius_pct=0)
        if accent_strip:
            gradient_fill(accent_strip, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(accent_strip, intensity * 10)

        # رقم
        txt(slide, f"{i + 1:02d}", cx + 0.2, y, 0.9, item_h,
            font=FONT_NUM, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, result, cx + 1.2, y + 0.08, cw - 1.9, item_h - 0.16,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "💡", "خاتمة\nالبحث")

    cx, cw = _content_area_x(), _content_w()

    # بطاقة خاتمة كبيرة
    card_h = H - 1.5
    card = rrect(slide, cx, 0.8, cw, card_h, T.card_rgb, radius_pct=12)
    if card:
        shadow(card, blur=22, dist=7, alpha=0.45)

    # تدرج داخل البطاقة
    gradient_fill(card, T.bg, T.card, angle=135)

    # شريط أعلى
    top = rrect(slide, cx, 0.8, cw, 0.32, T.accent_rgb, radius_pct=0)
    if top:
        gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

    # علامة اقتباس كبيرة
    txt(slide, "❝", cx + 0.5, 1.3, 2.5, 2.0,
        font=FONT_EN, size=52, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

    txt(slide, req.general_conclusion,
        cx + 0.5, 2.5, cw - 1.0, card_h - 2.5,
        font=FONT_BODY, size=TS.H3, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    # اسم الطالب في الأسفل
    hline(slide, cx + cw * 0.2, 0.8 + card_h - 1.2, cw * 0.6, T.accent_rgb, thickness=0.04)
    txt(slide, req.student_name,
        cx + 0.5, 0.8 + card_h - 1.0, cw - 1.0, 0.7,
        font=FONT_TITLE, size=TS.H3, bold=True,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "✅", "توصيات\nالبحث")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    recs = req.recommendations[:8]
    avail_h = H - cy - 0.5
    item_h = min(avail_h / max(len(recs), 1), 1.4) - 0.12

    for i, rec in enumerate(recs):
        y = cy + i * (item_h + 0.12)
        row_bg = rrect(slide, cx, y, cw, item_h, T.card_rgb, radius_pct=8)
        if row_bg:
            shadow(row_bg, blur=6, dist=2, alpha=0.2)

        # نقطة accent يمينية
        dot = oval(slide, cx + cw - 0.7, y + (item_h - 0.45) / 2,
                   0.45, 0.45, T.accent_rgb)
        if dot:
            gradient_fill(dot, T.accent_grad1, T.accent_grad2, 0)

        txt(slide, rec, cx + 0.3, y + 0.08, cw - 1.2, item_h - 0.16,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "🔭", "آفاق\nمستقبلية")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    items = req.future_work[:6]
    cols = 2 if len(items) > 3 else 1
    col_w = (cw - 0.3 * (cols - 1)) / cols
    rows = (len(items) + cols - 1) // cols
    avail_h = H - cy - 0.5
    card_h = min(avail_h / rows - 0.2, 3.0)

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = cx + col_idx * (col_w + 0.3)
        y = cy + row_idx * (card_h + 0.2)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=10)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)

        # شريط علوي
        top = rrect(slide, x, y, col_w, 0.24, T.accent_rgb, radius_pct=0)
        if top:
            gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

        # رقم دائري
        num_c = oval(slide, x + col_w / 2 - 0.45, y + 0.4, 0.9, 0.9, T.accent_rgb, alpha=25)
        txt(slide, str(i + 1), x + col_w / 2 - 0.45, y + 0.42, 0.9, 0.85,
            font=FONT_NUM, size=TS.H3, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, item, x + 0.25, y + 1.45, col_w - 0.5, card_h - 1.65,
            font=FONT_BODY, size=TS.BODY, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _sidebar(slide, T, "📚", "المراجع\nوالمصادر")

    cx, cw = _content_area_x(), _content_w()
    cy = 0.8
    refs = req.references[:12]
    avail_h = H - cy - 0.5
    item_h = max(min(avail_h / max(len(refs), 1) - 0.1, 1.2), 0.55)

    for i, ref in enumerate(refs):
        y = cy + i * (item_h + 0.1)
        if y + item_h > H - 0.3:
            break

        if i % 2 == 0:
            row_bg = rrect(slide, cx, y, cw, item_h, T.card_rgb, radius_pct=4)

        # شريط رقم يساري
        num_s = rect(slide, cx, y, 0.7, item_h, T.bg_rgb)
        txt(slide, f"{i + 1}", cx, y, 0.7, item_h,
            font=FONT_NUM, size=TS.LABEL, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, ref, cx + 0.8, y + 0.05, cw - 1.0, item_h - 0.1,
            font=FONT_BODY, size=TS.LABEL, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=120)

    # شريط جانبي عريض
    sidebar = gradient_rect(slide, 0, 0, SIDEBAR_W + 1, H, T.grad2, T.grad1, angle=90)
    sep = rect(slide, SIDEBAR_W + 1, 0, 0.1, H, T.accent_rgb)
    if sep:
        gradient_fill(sep, T.accent_grad1, T.accent_grad2, 90)

    # أيقونة شكراً في الشريط
    txt(slide, "🌟", 0.5, H / 2 - 3.5, SIDEBAR_W, 3.0,
        font=FONT_EN, size=60, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # منطقة المحتوى اليمنى
    cx = SIDEBAR_W + 1.5
    cw = W - cx - 1.0
    center_y = H / 2 - 3.5

    # بطاقة مركزية
    card = rrect(slide, cx, center_y, cw, 7.0, T.card_rgb, radius_pct=14)
    if card:
        shadow(card, blur=24, dist=8, alpha=0.45)
        gradient_fill(card, T.bg, T.card, angle=135)

    top_s = rrect(slide, cx, center_y, cw, 0.32, T.accent_rgb, radius_pct=0)
    if top_s:
        gradient_fill(top_s, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, "شكراً وتقديراً",
        cx + 0.5, center_y + 0.5, cw - 1.0, 2.5,
        font=FONT_TITLE, size=TS.H1, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    hline(slide, cx + cw * 0.15, center_y + 3.1, cw * 0.7, T.accent_rgb, thickness=0.05)

    txt(slide, req.student_name,
        cx + 0.5, center_y + 3.4, cw - 1.0, 1.2,
        font=FONT_TITLE, size=TS.H2, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    title_short = req.title_ar[:65] + ("..." if len(req.title_ar) > 65 else "")
    txt(slide, title_short,
        cx + 0.5, center_y + 4.7, cw - 1.0, 1.8,
        font=FONT_BODY, size=TS.BODY, bold=False, italic=True,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    footer = []
    if req.institution: footer.append(req.institution)
    if req.year: footer.append(req.year)
    if footer:
        txt(slide, " · ".join(footer),
            cx + 0.5, center_y + 6.4, cw - 1.0, 0.7,
            font=FONT_BODY, size=TS.LABEL, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    bottom_bar = rect(slide, 0, H - 0.25, W, 0.25, T.accent_rgb)
    if bottom_bar:
        gradient_fill(bottom_bar, T.accent_grad1, T.accent_grad2, 0)

    return slide

"""
Slide Builder — مذكرتي Pro v19 (Adaptive Text Engine)
═══════════════════════════════════════════════════════
v19: نظام تكيّف ذكي — النص القصير والطويل يملأ المساحة تلقائياً
  • estimate_lines()  — يحسب عدد الأسطر الفعلية بدقة
  • auto_font_size()  — يختار أكبر حجم خط يناسب المساحة
  • adaptive_rows()   — يوزّع الصفوف بلا فراغات ولا قطع
  • smart_card_h()    — يحسب ارتفاع البطاقة من محتواها
"""
from __future__ import annotations
import math
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from engine.primitives import (
    W, H, cm, pt,
    rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow,
    set_solid_alpha, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"

def set_font(font_name: str):
    global _FONT
    _FONT = font_name

def _hx(h: str) -> RGBColor:
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))


# ══════════════════════════════════════════════════════════════════════
# ADAPTIVE TEXT ENGINE
# ══════════════════════════════════════════════════════════════════════

_WIDE  = set("مشسظضطصغفقكلآأإ")
_NARROW= set("ريزونإأءة.،")

def _cw(c):
    if c in _WIDE:   return 1.18
    if c in _NARROW: return 0.72
    if c == ' ':     return 0.45
    if c.isascii():  return 0.65
    return 1.0

def estimate_lines(text: str, box_w: float, fs: float) -> int:
    if not text or box_w <= 0: return 1
    avg = fs * 0.55 * 0.0353
    cpl = max(1, box_w / avg)
    lines = 0
    for raw in text.split('\n'):
        if not raw: lines += 1; continue
        w = sum(_cw(c) for c in raw)
        lines += max(1, math.ceil(w / cpl))
    return max(1, lines)

def lh(fs: float, sp: float = 1.35) -> float:
    return fs * 0.0353 * sp

def text_h(text: str, bw: float, fs: float, sp: float = 1.35, pad: float = 0.15) -> float:
    return estimate_lines(text, bw, fs) * lh(fs, sp) + pad * 2

def auto_fs(text: str, bw: float, bh: float,
            smax: float = 16, smin: float = 7, step: float = 0.5) -> float:
    s = smax
    while s > smin:
        if text_h(text, bw, s) <= bh: return s
        s -= step
    return smin

def smart_card_h(text: str, bw: float, fs: float,
                 hdr: float = 0, pad: float = 0.3,
                 mn: float = 0.7, mx: float = 99) -> float:
    return max(mn, min(mx, hdr + text_h(text, bw, fs) + pad * 2))

def adaptive_rows(items, cy, bw, avh,
                  spref=11.5, smin=7.5, gap=0.14):
    if not items: return []
    fs = spref
    while fs >= smin:
        hs = [text_h(it, bw, fs) + gap for it in items]
        if sum(hs) <= avh: break
        fs -= 0.5
    total = sum(hs)
    stretch = (avh - total) / len(items) if total < avh else 0
    res, cy_ = [], cy
    for it, ih in zip(items, hs):
        fh = ih + stretch
        res.append((it, cy_, fh - gap, fs))
        cy_ += fh
    return res


# ══════════════════════════════════════════════════════════════════════
# SHARED HELPERS
# ══════════════════════════════════════════════════════════════════════

def _section_header(slide, T, title, subtitle=""):
    gradient_rect(slide, 0, 0, W, 2.8, T.grad1, T.grad2, angle=135)
    al = rect(slide, 0, 0, 0.35, 2.8, T.accent_rgb)
    if al: gradient_fill(al, T.accent_grad1, T.accent_grad2, angle=90)
    oval(slide, W-3.5, -1.2, 4.8, 4.8, T.accent_rgb, alpha=8)
    ts = auto_fs(title, W-1.5, 1.35, smax=22, smin=14)
    txt(slide, title, 0.6, 0.35, W-1.2, 1.35,
        font=_FONT, size=ts, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, rtl=True)
    if subtitle:
        ss = auto_fs(subtitle, W-1.2, 0.85, smax=13, smin=9)
        txt(slide, subtitle, 0.6, 1.6, W-1.2, 0.9,
            font=_FONT, size=ss, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

def _info_row(slide, T, label, value, y, rh=0.6):
    rrect(slide, 1.5, y-0.04, W-3.0, rh, T.bg2_rgb, radius_pct=8)
    txt(slide, label, 1.7, y, 5.5, rh,
        font=_FONT, size=11, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.RIGHT, rtl=True)
    vs = auto_fs(value, W-9.5, rh, smax=12, smin=7)
    txt(slide, value, 7.5, y, W-9.2, rh,
        font=_FONT, size=vs, bold=False, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, rtl=True)


# ══════════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    s1 = rrect(slide, -0.5, -0.5, 8, 8, T.accent_rgb, radius_pct=0)
    if s1: gradient_fill(s1, T.accent_grad1, T.accent_grad2, 45); set_solid_alpha(s1, 8)
    oval(slide, W-10, H-10, 14, 14, T.accent_rgb, alpha=6)
    ab = rrect(slide, 0, 0, W, 0.55, T.accent_rgb, radius_pct=0)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 0)

    if req.institution:
        rrect(slide, 1.2, 0.9, 14, 0.7, T.card_rgb, radius_pct=35)
        ins = auto_fs(req.institution, 13.6, 0.7, smax=11, smin=7)
        txt(slide, req.institution, 1.4, 0.9, 13.6, 0.7,
            font=_FONT, size=ins, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    TBW = W - 4.0
    ts  = auto_fs(req.title_ar, TBW, 5.5, smax=28, smin=12)
    thn = text_h(req.title_ar, TBW, ts, sp=1.4)
    enh = 0.0
    if req.title_en:
        es  = min(ts*0.5, 13)
        enh = text_h(req.title_en, TBW, es) + 0.2
    card_h = max(thn + enh + 1.0, 2.2)

    info_fields = []
    if req.student_name:   info_fields.append(("الطالب :",         req.student_name))
    if req.supervisor:     info_fields.append(("المشرف :",         req.supervisor))
    if req.co_supervisor:  info_fields.append(("المشرف المساعد :", req.co_supervisor))
    if req.specialization: info_fields.append(("التخصص :",        req.specialization))

    rh = max(max(
        smart_card_h(v, W-9.2, 12, mn=0.52, mx=1.1) + 0.12
        for _,v in info_fields
    ) if info_fields else 0.6, 0.6)

    info_h  = len(info_fields) * (rh + 0.05) + 0.1
    yr_h    = 0.6 if req.year else 0
    total   = card_h + 0.5 + info_h + yr_h + 0.3
    top_m   = max(1.6, (H - total) / 2)

    ty = top_m
    tc = rrect(slide, 1.5, ty-0.3, W-3.0, card_h+0.6, T.card_rgb, radius_pct=12)
    if tc: shadow(tc, blur=20, dist=6, alpha=0.4)
    st = rrect(slide, 1.5, ty-0.3, W-3.0, 0.25, T.accent_rgb, radius_pct=0)
    if st: gradient_fill(st, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, req.title_ar, 2.0, ty, TBW, thn,
        font=_FONT, size=ts, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True)
    if req.title_en and enh > 0:
        es = min(ts*0.5, 13)
        txt(slide, req.title_en, 2.0, ty+thn, TBW, enh,
            font="Calibri", size=es, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    div_y = ty + card_h + 0.3 + 0.25
    hline(slide, W*0.25, div_y, W*0.5, T.accent_rgb, thickness=0.06)

    iy = div_y + 0.2
    for i,(label,value) in enumerate(info_fields):
        _info_row(slide, T, label, value, iy + i*(rh+0.05), rh)

    if req.year:
        yy = iy + len(info_fields)*(rh+0.05) + 0.2
        yc = rrect(slide, W/2-2.2, yy, 4.4, 0.52, T.accent_rgb, radius_pct=50)
        if yc: gradient_fill(yc, T.accent_grad1, T.accent_grad2, 0); shadow(yc,blur=10,dist=2,alpha=0.3)
        txt(slide, req.year, W/2-2.2, yy, 4.4, 0.52,
            font="Calibri", size=12, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

    bb = rect(slide, 0, H-0.22, W, 0.22, T.accent_rgb)
    if bb: gradient_fill(bb, T.accent_grad1, T.accent_grad2, 0)
    return slide


# ══════════════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")

    cy = 3.1
    avh = H - cy - 0.6
    items = []
    if req.intro_overview: items.append(("نظرة عامة",    req.intro_overview))
    if req.intro_approach: items.append(("المنهج المتبع", req.intro_approach))
    if not items: return slide

    cols = min(len(items), 2)
    cw = (W - 2.4 - (cols-1)*0.3) / cols
    HH, PAD = 0.65, 0.25

    for i,(lbl,val) in enumerate(items[:2]):
        x = 1.2 + i*(cw+0.3)
        cah = avh - HH - PAD*2
        fs  = auto_fs(val, cw-0.4, cah, smax=14, smin=7.5)
        # Card always fills full available height for visual balance
        c   = rect(slide, x, cy, cw, avh, T.card_rgb)
        if c: shadow(c, blur=14, dist=4, alpha=0.35)
        cs  = rrect(slide, x, cy, cw, HH, T.accent_rgb, radius_pct=0)
        if cs: gradient_fill(cs, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.2, cy+0.05, cw-0.4, HH,
            font=_FONT, size=13, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        # Vertically center text inside card body
        actual_th = text_h(val, cw-0.4, fs)
        text_area_h = avh - HH
        text_y = cy + HH + max(PAD, (text_area_h - actual_th) / 2)
        txt(slide, val, x+0.2, text_y, cw-0.4, actual_th + 0.1,
            font=_FONT, size=fs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    n = len(req.chapters)
    _section_header(slide, T, "خطة البحث",
                    f"يتكون البحث من {n} فصل{'اً' if n!=2 else 'ين'}")

    cy = 3.1; avh = H - cy - 0.5
    chs = req.chapters[:9]
    if not chs: return slide

    RW   = W - 5.8
    long = max(chs, key=lambda c: len(c.title)).title
    fs   = auto_fs(long, RW, avh/len(chs)-0.12, smax=13, smin=7.5)
    ml   = max(estimate_lines(c.title, RW, fs) for c in chs)
    rh   = max(ml * lh(fs) + 0.22, 0.65)
    tot  = len(chs)*(rh+0.12) - 0.12
    if tot < avh: rh += (avh - tot) / len(chs)

    for i,ch in enumerate(chs):
        y = cy + i*(rh+0.12)
        if y+rh > H-0.3: break
        rb = rrect(slide, 1.2, y, W-2.4, rh, T.card_rgb, radius_pct=8)
        if rb: shadow(rb, blur=8, dist=2, alpha=0.25)
        bs = min(rh*0.7, 0.75)
        by = y + (rh-bs)/2
        nb = oval(slide, 1.5, by, bs, bs, T.accent_rgb)
        if nb: gradient_fill(nb, T.accent_grad1, T.accent_grad2, 0)
        num_fs = max(8, bs*18)
        num_lh = num_fs * 0.0353 * 1.25
        num_y  = by + (bs - num_lh) / 2
        txt(slide, str(i+1), 1.5, num_y, bs, num_lh,
            font="Calibri", size=num_fs, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False, margin=0)
        pages_w = 2.0 if ch.pages else 0
        title_w = RW - pages_w - 0.3
        title_th = text_h(ch.title, title_w, fs)
        title_y  = y + max(0.05, (rh - title_th) / 2)
        txt(slide, ch.title, 2.5, title_y, title_w, title_th + 0.1,
            font=_FONT, size=fs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        if ch.pages:
            txt(slide, ch.pages, 1.55, y, pages_w, rh,
                font="Calibri", size=10, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.LEFT, rtl=False)
    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")

    cy = 3.1; avh = H - cy - 0.5; BW = W - 2.8

    if req.main_problem:
        mfs = auto_fs(req.main_problem, BW, avh*0.38, smax=13, smin=7.5)
        mh  = smart_card_h(req.main_problem, BW, mfs, hdr=0.6, pad=0.2, mn=1.4, mx=avh*0.45)
        c   = rrect(slide, 1.2, cy, W-2.4, mh, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=12, dist=4, alpha=0.35)
        lb  = rrect(slide, 1.2, cy, 3.8, 0.55, T.accent_rgb, radius_pct=0)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, "الإشكالية الرئيسية", 1.3, cy, 3.6, 0.55,
            font=_FONT, size=11, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem, 1.4, cy+0.62, BW, mh-0.75,
            font=_FONT, size=mfs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        cy += mh + 0.2

    if req.main_question and cy < H-1.0:
        rem = H - cy - 0.5
        qfs = auto_fs(req.main_question, BW, min(rem*0.4,1.8), smax=13, smin=7.5)
        qh  = smart_card_h(req.main_question, BW, qfs, pad=0.2, mn=1.0, mx=min(rem*0.42,2.0))
        qc  = rrect(slide, 1.2, cy, W-2.4, qh, T.bg2_rgb, radius_pct=8)
        if qc: shadow(qc, blur=8, dist=3, alpha=0.2)
        hline(slide, 1.2, cy, 0.3, T.accent_rgb, thickness=1.5)
        txt(slide, req.main_question, 1.7, cy+0.12, W-3.2, qh-0.24,
            font=_FONT, size=qfs, bold=True, italic=True, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        cy += qh + 0.2

    if req.sub_questions and cy < H-0.8:
        layout = adaptive_rows(req.sub_questions[:7], cy, W-3.2, H-cy-0.4,
                               spref=11, smin=7.0, gap=0.1)
        for st, sy, sh, sfs in layout:
            oval(slide, W-2.6, sy+sh*0.35, 0.26, 0.26, T.accent_rgb)
            txt(slide, st, 1.2, sy, W-3.2, sh,
                font=_FONT, size=sfs, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهداف البحث وفرضياته", "")

    cy = 3.1; avh = H - cy - 0.5
    cols_data = []
    if req.objectives: cols_data.append(("الأهداف",  req.objectives))
    if req.hypotheses:  cols_data.append(("الفرضيات", req.hypotheses))
    if not cols_data: return slide

    cols = min(len(cols_data), 2)
    cw   = (W - 2.4 - (cols-1)*0.3) / cols
    HH   = 0.65

    for i,(lbl,items) in enumerate(cols_data[:2]):
        x   = 1.2 + i*(cw+0.3)
        lng = max(items, key=len) if items else ""
        ia  = (avh - HH) / max(len(items),1) - 0.1
        fs  = auto_fs(lng, cw-1.2, ia, smax=11.5, smin=7.0)
        c   = rrect(slide, x, cy, cw, avh, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=14, dist=4, alpha=0.35)
        h   = rrect(slide, x, cy, cw, HH, T.accent_rgb, radius_pct=0)
        if h: gradient_fill(h, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.15, cy, cw-0.3, HH,
            font=_FONT, size=14, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True)
        layout = adaptive_rows(items, cy+HH+0.1, cw-1.2,
                               avh-HH-0.2, spref=fs, smin=7.0, gap=0.08)
        for j,(it,iy,ih,ifs) in enumerate(layout):
            bs = min(ih*0.6, 0.52); by = iy+(ih-bs)/2
            oval(slide, x+cw-0.9, by, bs, bs, T.bg_rgb)
            txt(slide, str(j+1), x+cw-0.9, by, bs, bs,
                font="Calibri", size=max(7,bs*18), bold=True, color=T.accent_rgb,
                align=PP_ALIGN.CENTER, rtl=False)
            txt(slide, it, x+0.2, iy, cw-1.2, ih,
                font=_FONT, size=ifs, bold=False, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهمية البحث", "")

    cy = 3.1; avh = H - cy - 0.4
    items = list(req.importance[:6]) if req.importance else []
    if req.reasons and len(items) < 6: items.append(req.reasons)
    if not items: return slide

    cols = 2 if len(items) > 3 else 1
    cw   = (W - 2.4 - (cols-1)*0.3) / cols
    rows = math.ceil(len(items)/cols)
    AW   = 0.2; ctw = cw - AW - 0.6
    lng  = max(items, key=len)
    cah  = avh/rows - 0.2
    fs   = auto_fs(lng, ctw, cah-0.3, smax=12, smin=7.5)
    mch  = max(text_h(it, ctw, fs) for it in items)
    ch   = min(max(mch+0.3, 0.9), cah)

    for i,item in enumerate(items):
        ci = i%cols; ri = i//cols
        x  = 1.2 + ci*(cw+0.3)
        y  = cy + ri*(ch+0.2)
        if y+ch > H-0.2: break
        c  = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=10, dist=3, alpha=0.3)
        vline(slide, x, y, ch, T.accent_rgb, thickness=AW)
        ifs = auto_fs(item, ctw, ch-0.2, smax=fs, smin=7.5)
        txt(slide, item, x+AW+0.2, y+0.1, ctw, ch-0.2,
            font=_FONT, size=ifs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "منهجية البحث", "الإجراءات والأدوات")

    cy = 3.1; avh = H - cy - 0.4
    fields = []
    if req.methodology:  fields.append(("المنهج",      req.methodology))
    if req.sample_type:  fields.append(("العينة",       req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة",  req.sample_size))
    if req.tool:         fields.append(("الأداة",       req.tool))
    if not fields: return slide

    cols = 2 if len(fields) > 2 else 1
    cw   = (W - 2.4 - (cols-1)*0.3) / cols
    rows = math.ceil(len(fields)/cols)
    HH   = 0.55; ctw = cw - 0.4
    lv   = max(v for _,v in fields)
    cah  = avh/rows - 0.2
    fs   = auto_fs(lv, ctw, cah-HH-0.3, smax=12.5, smin=7.5)
    mvh  = max(text_h(v, ctw, fs) for _,v in fields)
    ch   = min(max(mvh+HH+0.4, 1.2), cah)

    for i,(lbl,val) in enumerate(fields[:4]):
        ci = i%cols; ri = i//cols
        x  = 1.2 + ci*(cw+0.3)
        y  = cy + ri*(ch+0.2)
        if y+ch > H-0.2: break
        c  = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=10, dist=3, alpha=0.3)
        lc = rrect(slide, x, y, cw, HH, T.accent_rgb, radius_pct=0)
        if lc: gradient_fill(lc, T.accent_grad1, T.accent_grad2, 0)
        ls = auto_fs(lbl, ctw, HH, smax=12, smin=8)
        txt(slide, lbl, x+0.2, y, ctw, HH,
            font=_FONT, size=ls, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
        vfs = auto_fs(val, ctw, ch-HH-0.15, smax=fs, smin=7.5)
        txt(slide, val, x+0.2, y+HH+0.1, ctw, ch-HH-0.15,
            font=_FONT, size=vfs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "الإحصاءات والأرقام الرئيسية", "")

    cy = 3.1; stats = req.stats[:6]; n = len(stats)
    if n == 0: return slide

    cols = 3 if n >= 3 else n
    rows = math.ceil(n/cols)
    gap  = 0.3
    cw   = (W - 2.4 - (cols-1)*gap) / cols
    avh  = H - cy - 0.4
    ch   = min(avh/rows - gap, 3.8)

    for i,stat in enumerate(stats):
        ci = i%cols; ri = i//cols
        x  = 1.2 + ci*(cw+gap)
        y  = cy + ri*(ch+gap)
        c  = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if c: shadow(c, blur=14, dist=5, alpha=0.4)
        b  = rrect(slide, x, y, cw, 0.22, T.accent_rgb, radius_pct=0)
        if b: gradient_fill(b, T.accent_grad1, T.accent_grad2, 0)
        vs = auto_fs(stat.value, cw-0.4, ch*0.5, smax=34, smin=14)
        txt(slide, stat.value, x+0.2, y+0.38, cw-0.4, ch*0.5,
            font="Calibri", size=vs, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False)
        if stat.unit:
            txt(slide, stat.unit, x+0.2, y+0.38+ch*0.5, cw-0.4, 0.55,
                font=_FONT, size=10, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=True)
        ls = auto_fs(stat.label, cw-0.4, 0.75, smax=11, smin=7)
        txt(slide, stat.label, x+0.2, y+ch-0.85, cw-0.4, 0.75,
            font=_FONT, size=ls, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS  (مع تقسيم تلقائي لشريحتين)
# ══════════════════════════════════════════════════════════════════════
def _results_slide(prs, req, T, results, suffix=""):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, f"نتائج البحث{suffix}", "أبرز ما توصلت إليه الدراسة")
    cy = 3.1; avh = H - cy - 0.4; BW = W - 5.0
    layout = adaptive_rows(results, cy, BW, avh, spref=11.5, smin=7.5, gap=0.12)
    for i,(res,ry,rh,rfs) in enumerate(layout):
        rb = rrect(slide, 1.2, ry, W-2.4, rh, T.card_rgb, radius_pct=8)
        if rb: shadow(rb, blur=6, dist=2, alpha=0.2)
        bh = min(rh*0.65, 0.6); by = ry+(rh-bh)/2
        bd = rrect(slide, W-3.0, by, 1.5, bh, T.accent_rgb, radius_pct=50)
        if bd: gradient_fill(bd, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i+1), W-3.0, by, 1.5, bh,
            font="Calibri", size=max(8,bh*18), bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)
        txt(slide, res, 1.5, ry+0.06, BW, rh-0.12,
            font=_FONT, size=rfs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

def make_results(prs, req, T):
    results = req.main_results[:12]
    if not results: return _results_slide(prs, req, T, [])
    cy = 3.1; avh = H - cy - 0.4; BW = W - 5.0
    total = sum(text_h(r, BW, 7.5)+0.12 for r in results)
    if total <= avh or len(results) <= 6:
        return _results_slide(prs, req, T, results)
    mid = len(results)//2
    _results_slide(prs, req, T, results[:mid], " (1/2)")
    return _results_slide(prs, req, T, results[mid:], " (2/2)")


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خاتمة البحث", "الاستنتاج العام")

    cy = 3.1; ch = H - cy - 0.7
    c  = rrect(slide, 1.5, cy, W-3.0, ch, T.card_rgb, radius_pct=14)
    if c: shadow(c, blur=20, dist=6, alpha=0.4)
    ts = rrect(slide, 1.5, cy, W-3.0, 0.28, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)

    qs = min(40, ch*8)
    txt(slide, "❝", 2.5, cy+0.35, 1.5, 1.5,
        font="Calibri", size=qs, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.LEFT, rtl=False)

    ctw = W - 4.5; cth = ch - 1.5
    fs  = auto_fs(req.general_conclusion, ctw, cth, smax=16, smin=8)
    txt(slide, req.general_conclusion, 2.0, cy+1.0, ctw, cth,
        font=_FONT, size=fs, bold=False, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "توصيات البحث", "")

    cy = 3.1; avh = H - cy - 0.4
    recs = req.recommendations[:9]
    if not recs: return slide

    DW = 0.35; BW = W - 3.2 - DW
    layout = adaptive_rows(recs, cy, BW, avh, spref=11.5, smin=7.5, gap=0.1)
    for rec,ry,rh,rfs in layout:
        rb = rrect(slide, 1.2, ry, W-2.4, rh, T.card_rgb, radius_pct=8)
        if rb: shadow(rb, blur=6, dist=2, alpha=0.2)
        ds = min(rh*0.45, 0.35)
        oval(slide, W-2.4, ry+(rh-ds)/2, ds, ds, T.accent_rgb)
        txt(slide, rec, 1.5, ry+0.06, BW, rh-0.12,
            font=_FONT, size=rfs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE WORK
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "آفاق البحث المستقبلية", "")

    cy = 3.1; avh = H - cy - 0.4
    items = req.future_work[:6]
    if not items: return slide

    cols = 2 if len(items) > 3 else 1
    cw   = (W - 2.4 - (cols-1)*0.3) / cols
    rows = math.ceil(len(items)/cols)
    IW   = 0.6; ctw = cw - IW - 0.35
    lng  = max(items, key=len)
    cah  = avh/rows - 0.2
    fs   = auto_fs(lng, ctw, cah-0.3, smax=12, smin=7.5)
    mch  = max(text_h(it, ctw, fs) for it in items)
    ch   = min(max(mch+0.3, 1.0), cah)

    for i,item in enumerate(items):
        ci = i%cols; ri = i//cols
        x  = 1.2 + ci*(cw+0.3)
        y  = cy + ri*(ch+0.2)
        if y+ch > H-0.2: break
        c  = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=10, dist=3, alpha=0.3)
        ic = rrect(slide, x, y, IW, ch, T.accent_rgb, radius_pct=0)
        if ic: gradient_fill(ic, T.accent_grad1, T.accent_grad2, 90)
        ifs = auto_fs(item, ctw, ch-0.25, smax=fs, smin=7.5)
        txt(slide, item, x+IW+0.15, y+0.12, ctw, ch-0.25,
            font=_FONT, size=ifs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "المراجع والمصادر", "")

    cy = 3.1; avh = H - cy - 0.4
    refs = req.references[:14]
    if not refs: return slide

    BW  = W - 4.5
    lng = max(refs, key=len)
    fs  = auto_fs(lng, BW, avh/len(refs)-0.1, smax=11, smin=7.0)
    layout = adaptive_rows(refs, cy, BW, avh, spref=fs, smin=7.0, gap=0.1)
    for i,(ref,ry,rh,rfs) in enumerate(layout):
        if i%2 == 0:
            rrect(slide, 1.2, ry, W-2.4, rh, T.card_rgb, radius_pct=4)
        txt(slide, f"[{i+1}]", W-2.8, ry, 1.4, rh,
            font="Calibri", size=max(7, rfs*0.85), bold=True, color=T.accent_rgb,
            align=PP_ALIGN.LEFT, rtl=False)
        txt(slide, ref, 1.5, ry+0.03, BW, rh-0.06,
            font=_FONT, size=rfs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL / THANK YOU
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    oval(slide, -3,  -3, 12, 12, T.accent_rgb, alpha=6)
    oval(slide, W-9, H-9, 14, 14, T.accent_rgb, alpha=5)
    oval(slide, W-5, -2,  8,  8, T.bg2_rgb,    alpha=30)

    cw, ch = 22, 10
    cx, cy = (W-cw)/2, (H-ch)/2
    c  = rrect(slide, cx, cy, cw, ch, T.card_rgb, radius_pct=14)
    if c: shadow(c, blur=24, dist=8, alpha=0.45)
    ts = rrect(slide, cx, cy, cw, 0.35, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, "شكراً وتقديراً", cx+1.0, cy+0.5, cw-2.0, 2.5,
        font=_FONT, size=36, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    hline(slide, cx+cw*0.2, cy+3.2, cw*0.6, T.accent_rgb, thickness=0.05)

    ns = auto_fs(req.student_name, cw-2.0, 1.2, smax=20, smin=12)
    txt(slide, req.student_name, cx+1.0, cy+3.4, cw-2.0, 1.2,
        font=_FONT, size=ns, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    tih = 2.2
    tis = auto_fs(req.title_ar, cw-2.0, tih, smax=13, smin=7.5)
    txt(slide, req.title_ar, cx+1.0, cy+4.7, cw-2.0, tih,
        font=_FONT, size=tis, bold=False, italic=True, color=T.muted_rgb,
        align=PP_ALIGN.CENTER, rtl=True)

    fps = [p for p in [req.institution, req.year] if p]
    if fps:
        footer = "  ·  ".join(fps)
        fts = auto_fs(footer, cw-2.0, 0.8, smax=11, smin=7)
        txt(slide, footer, cx+1.0, cy+ch-1.0, cw-2.0, 0.8,
            font=_FONT, size=fts, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True)

    bb = rect(slide, 0, H-0.25, W, 0.25, T.accent_rgb)
    if bb: gradient_fill(bb, T.accent_grad1, T.accent_grad2, 0)
    return slide

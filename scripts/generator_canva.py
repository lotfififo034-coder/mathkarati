"""
مذكرتي Pro — ULTRA VISUAL ENGINE v10
========================================
تحسينات جذرية في الجودة البصرية:
- نظام تدرج لوني حقيقي عبر XML
- ظلال متعددة الطبقات
- هندسة متطورة بتأثيرات مميزة
- بطاقات بعمق بصري حقيقي
- غلاف بمستوى Behance/Awwwards
"""

import sys, json, math, datetime
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import shutil, os, subprocess

W, H   = 33.867, 19.05   # cm — 16:9 = 1920×1080 mapping
MX, MY = 1.40, 0.88      # margins

# ── Font availability check ─────────────────────────────────────
def _font_available(name):
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(["fc-list", f":family={name}"],
                                 capture_output=True, text=True, timeout=3)
            if name.lower() in out.stdout.lower():
                return True
        except: pass
    for d in ["/usr/share/fonts", "/usr/local/share/fonts",
              os.path.expanduser("~/.fonts"), "C:/Windows/Fonts"]:
        if os.path.isdir(d):
            for root, _, files in os.walk(d):
                for f in files:
                    if name.lower() in f.lower() and f.lower().endswith((".ttf",".otf")):
                        return True
    return False

_CAIRO_OK = _font_available("Cairo")
HF = "Cairo" if _CAIRO_OK else "Calibri"   # heading font
BF = "Cairo" if _CAIRO_OK else "Arial"     # body font

# ── Core helpers ────────────────────────────────────────────────
def rgb(r,g,b): return RGBColor(r,g,b)
def hx(h):      return RGBColor.from_string(h.lstrip('#'))
def safe(v,fb=""): return str(v).strip() if v else fb
def clamp(v,lo,hi): return max(lo,min(hi,v))
def blank(prs): return prs.slides.add_slide(prs.slide_layouts[6])
def cm(v): return Cm(v)
def emu(v): return int(Cm(v))

# ── Drawing primitives ─────────────────────────────────────────
def rect(slide, x,y,w,h, fill, line_color=None, line_w=0.5):
    if w<=0 or h<=0: return None
    s = slide.shapes.add_shape(1, cm(x),cm(y),cm(w),cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line_color: s.line.color.rgb=line_color; s.line.width=Pt(line_w)
    else: s.line.fill.background()
    return s

def rrect(slide, x,y,w,h, fill, r_pct=8, line_color=None, line_w=0.5):
    if w<=0 or h<=0: return None
    s = slide.shapes.add_shape(5, cm(x),cm(y),cm(w),cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line_color: s.line.color.rgb=line_color; s.line.width=Pt(line_w)
    else: s.line.fill.background()
    try:
        adj = s.adjustments
        if adj and len(adj)>0: adj[0] = clamp(r_pct,0,50)*1000
    except: pass
    return s

def oval(slide, x,y,w,h, fill, alpha=100):
    if w<=0 or h<=0: return None
    s = slide.shapes.add_shape(9, cm(x),cm(y),cm(w),cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        _set_fill_alpha(s, alpha)
    return s

def _set_fill_alpha(shape, alpha_pct):
    try:
        sp = shape._element
        spPr = sp.find(qn('p:spPr'))
        fld = spPr.find('.//' + qn('a:solidFill'))
        if fld is not None:
            srgb = fld.find(qn('a:srgbClr'))
            if srgb is not None:
                for e in srgb.findall(qn('a:alpha')): srgb.remove(e)
                alp = etree.SubElement(srgb, qn('a:alpha'))
                alp.set('val', str(int(alpha_pct * 1000)))
    except: pass

def bg(slide, color): rect(slide,0,0,W,H,color)
def lh(slide,x,y,w,color,h=0.08): rect(slide,x,y,w,h,color)
def lv(slide,x,y,h2,color,w=0.08): rect(slide,x,y,w,h2,color)

# ── XML gradient fill ────────────────────────────────────────────
def gradient_fill(shape, c1_hex, c2_hex, angle=90):
    """Applies a real linear gradient via XML"""
    try:
        sp = shape._element
        spPr = sp.find(qn('p:spPr'))
        # Remove old fill
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'),
                    qn('a:pattFill'), qn('a:blipFill')]:
            for el in spPr.findall(tag): spPr.remove(el)
        # Build gradFill
        grad = etree.SubElement(spPr, qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))
        gs0 = etree.SubElement(gsLst, qn('a:gs')); gs0.set('pos','0')
        sc0 = etree.SubElement(gs0, qn('a:srgbClr')); sc0.set('val', c1_hex.lstrip('#'))
        gs1 = etree.SubElement(gsLst, qn('a:gs')); gs1.set('pos','100000')
        sc1 = etree.SubElement(gs1, qn('a:srgbClr')); sc1.set('val', c2_hex.lstrip('#'))
        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000)))
        lin.set('scaled', '0')
    except: pass

def gradient_rect(slide, x,y,w,h, c1_hex, c2_hex, angle=0):
    """Rectangle with real gradient"""
    if w<=0 or h<=0: return None
    s = rect(slide,x,y,w,h, hx(c1_hex))
    gradient_fill(s, c1_hex, c2_hex, angle)
    s.line.fill.background()
    return s

# ── Real shadow via XML ─────────────────────────────────────────
def shadow(shape, blur=16, dist=5, angle=135, alpha=0.22, color="000000"):
    try:
        sp = shape._element
        spPr = sp.find(qn('p:spPr'))
        for old in spPr.findall('.//' + qn('a:effectLst')): spPr.remove(old)
        eLst = etree.SubElement(spPr, qn('a:effectLst'))
        shdw = etree.SubElement(eLst, qn('a:outerShdw'))
        shdw.set('blurRad', str(int(blur*12700)))
        shdw.set('dist',    str(int(dist*12700)))
        shdw.set('dir',     str(int(angle*60000)))
        shdw.set('algn',    'tl')
        srgb = etree.SubElement(shdw, qn('a:srgbClr'))
        srgb.set('val', color)
        alp  = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha*100000)))
    except: pass

def glow(shape, color_hex, radius=8, alpha=0.40):
    """Glow effect"""
    try:
        sp = shape._element
        spPr = sp.find(qn('p:spPr'))
        eLst = spPr.find(qn('a:effectLst'))
        if eLst is None:
            eLst = etree.SubElement(spPr, qn('a:effectLst'))
        g = etree.SubElement(eLst, qn('a:glow'))
        g.set('rad', str(int(radius*12700)))
        srgb = etree.SubElement(g, qn('a:srgbClr'))
        srgb.set('val', color_hex.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha*100000)))
    except: pass

# ── Text ─────────────────────────────────────────────────────────
def txt(slide, text, x,y,w,h,
        font=None, size=14, bold=False, italic=False,
        color=None, align=PP_ALIGN.RIGHT, mg=0.12,
        rtl=True, spacing=None):
    if w<=0 or h<=0 or not text: return None
    tb = slide.shapes.add_textbox(cm(x),cm(y),cm(w),cm(h))
    tb.word_wrap = True
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=cm(mg); tf.margin_right=cm(mg)
    tf.margin_top=cm(0.04); tf.margin_bottom=cm(0.04)
    p = tf.paragraphs[0]; p.alignment = align
    if spacing:
        try: p.line_spacing = Pt(spacing)
        except: pass
    run = p.add_run()
    run.text = str(text)
    run.font.name = font or BF
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color: run.font.color.rgb = color
    return tb

def overline(slide, text, x,y,w, color, align=PP_ALIGN.LEFT):
    return txt(slide, text.upper(), x,y,w,0.45,
               font="Calibri", size=10, bold=True,
               color=color, align=align, rtl=False)

# ═══════════════════════════════════════════════════════════════════
# COLOR PALETTES — 9 Premium Palettes
# ═══════════════════════════════════════════════════════════════════
def PAL(d,m,l, a,a2, tl,td,tm, g1,g2, sc_list, hf,bf, fam,
        card_d, card_l, border):
    return {
        "D":hx(d), "M":hx(m), "L":hx(l),
        "A":hx(a), "A2":hx(a2),
        "TL":hx(tl), "TD":hx(td), "TM":hx(tm),
        "G1":g1, "G2":g2,           # hex strings for gradient
        "SC":[hx(c) for c in sc_list],
        "SCH": sc_list,              # hex strings
        "HF":hf, "BF":bf, "FAM":fam,
        "CD":hx(card_d), "CL":hx(card_l),
        "BO":hx(border),
        "CB":hx("FFFFFF"), "CE":hx("E8EDF5"),
        "DH":d, "MH":m, "AH":a, "A2H":a2,  # hex strings
    }

PALETTES = {
# ── NOIR family ─────────────────────────────────────────────────
"navy_gold": PAL(
    "06111F","0E2040","EEF3FA",
    "C8921A","F0C84A",
    "FFFFFF","06111F","7A90AD",
    "0E2040","C8921A",
    ["C8921A","4AB8FF","F0C84A","1A5A9A","FF6B6B","00D4AA","A78BFA"],
    HF,BF,"NOIR",
    "0D1B35","FFFFFF","1E3A5F"),

"midnight_purple": PAL(
    "0E0520","28096B","F5F3FF",
    "BF7FFF","DDB3FF",
    "FFFFFF","0E0520","9CA3AF",
    "28096B","BF7FFF",
    ["BF7FFF","FF6B6B","DDB3FF","7B3FE0","00D4AA","F0C84A","4AB8FF"],
    HF,BF,"NOIR",
    "200B52","FFFFFF","3D1A8A"),

"forest": PAL(
    "0B2418","1A4030","F0FDF6",
    "78C850","B8E890",
    "FFFFFF","0B2418","6B8F75",
    "1A4030","78C850",
    ["78C850","4AB8FF","B8E890","1A4030","FF6B6B","F0C84A","BF7FFF"],
    HF,BF,"NOIR",
    "162E20","FFFFFF","2D6040"),

"sand_gold": PAL(
    "2A1508","5A3515","FFF9F0",
    "C8821A","E8B860",
    "FFFFFF","2A1508","8A6A40",
    "5A3515","C8821A",
    ["C8821A","4AB8FF","E8B860","5A3515","FF6B6B","00D4AA","78C850"],
    HF,BF,"NOIR",
    "3A2010","FFFFFF","7A4A25"),

# ── VIVID family ────────────────────────────────────────────────
"dark_teal": PAL(
    "081E2E","0E3D52","E8F8FF",
    "00C8A0","67E8D0",
    "FFFFFF","061A28","4A8BA0",
    "0E3D52","00C8A0",
    ["00C8A0","FF6535","67E8D0","0A80A0","BF7FFF","F0C84A","FF6B6B"],
    HF,BF,"VIVID",
    "0A2E44","FFFFFF","1A5A72"),

"charcoal_orange": PAL(
    "181826","2C2C42","FFF8F5",
    "FF6535","FFA070",
    "FFFFFF","181826","7A7A9A",
    "2C2C42","FF6535",
    ["FF6535","00C8A0","FFA070","2C2C42","BF7FFF","4AB8FF","F0C84A"],
    HF,BF,"VIVID",
    "242438","FFFFFF","404058"),

"burgundy": PAL(
    "350014","681230","FFF5F8",
    "E8409A","FFB0D8",
    "FFFFFF","350014","8A4060",
    "681230","E8409A",
    ["E8409A","4AB8FF","FFB0D8","681230","00C8A0","F0C84A","BF7FFF"],
    HF,BF,"VIVID",
    "4A0820","FFFFFF","8A2050"),

# ── MINIMAL family ───────────────────────────────────────────────
"ice_blue": PAL(
    "08224A","184888","EEF5FF",
    "0058CC","4AB0FF",
    "FFFFFF","08224A","3A5A80",
    "08224A","0058CC",
    ["0058CC","FF6535","4AB0FF","184888","00C8A0","BF7FFF","F0C84A"],
    HF,BF,"MINIMAL",
    "0C3060","FFFFFF","1A3C6E"),

"slate_crimson": PAL(
    "1A1A2E","2D2D4A","F8F9FF",
    "E63946","FF8FA3",
    "FFFFFF","1A1A2E","6B6B8A",
    "2D2D4A","E63946",
    ["E63946","4AB8FF","FF8FA3","2D2D4A","00C8A0","F0C84A","BF7FFF"],
    HF,BF,"MINIMAL",
    "26263C","FFFFFF","3D3D58"),
}

# ═══════════════════════════════════════════════════════════════════
# DECORATIVE UTILITIES
# ═══════════════════════════════════════════════════════════════════
def deco_arc_corner(slide, T, corner='tr', size=6.0):
    """Large decorative arcs in slide corners"""
    c_main = T["A"]
    c_sec  = T["M"]
    if corner == 'tr':
        oval(slide, W-size*0.9, -size*0.7, size*1.8, size*1.8, c_main, 12)
        oval(slide, W-size*0.55, -size*0.4, size*1.1, size*1.1, c_sec,  8)
    elif corner == 'bl':
        oval(slide, -size*0.6, H-size*0.9, size*1.8, size*1.8, c_main, 10)
        oval(slide, -size*0.2, H-size*0.5, size*1.0, size*1.0, c_sec, 7)
    elif corner == 'br':
        oval(slide, W-size*0.8, H-size*0.8, size*1.6, size*1.6, c_main, 12)
    elif corner == 'tl':
        oval(slide, -size*0.7, -size*0.5, size*1.8, size*1.8, c_main, 10)

def deco_dots_grid(slide, x, y, w, h, color, rows=4, cols=6, alpha=30):
    gx = w/cols; gy = h/rows
    r = 0.12
    for ri in range(rows):
        for ci in range(cols):
            dx = x + ci*gx + gx/2 - r/2
            dy = y + ri*gy + gy/2 - r/2
            s = oval(slide, dx, dy, r, r, color)
            if s: _set_fill_alpha(s, alpha)

def deco_wave_lines(slide, x, y, w, color, n=5, spacing=0.28):
    """Thin horizontal decorative lines"""
    for i in range(n):
        alpha = max(8, 22 - i*4)
        s = rect(slide, x, y + i*spacing, w*(1 - i*0.12), 0.05, color)
        if s: _set_fill_alpha(s, alpha)

def accent_pill(slide, x, y, w, h, T, text, text_color=None):
    """Rounded pill badge"""
    s = rrect(slide, x,y,w,h, T["A"], r_pct=50)
    if s: shadow(s, blur=8, dist=2, alpha=0.20)
    tc = text_color or T["D"]
    txt(slide, text, x,y,w,h, font=BF, size=10, bold=True,
        color=tc, align=PP_ALIGN.CENTER)
    return s

def card_premium(slide, x,y,w,h, T, accent_color,
                  style='dark', corner_r=10, shadow_alpha=0.20):
    """Premium card with gradient top bar + shadow"""
    bg_c = T["CD"] if style=='dark' else T["CB"]
    s = rrect(slide, x,y,w,h, bg_c, r_pct=corner_r,
              line_color=None if style=='dark' else hx("E2E8F0"),
              line_w=0.6)
    if s: shadow(s, blur=18, dist=5, alpha=shadow_alpha)
    # gradient top strip
    gs = gradient_rect(slide, x, y, w, 0.16, T["AH"], T["A2H"], angle=0)
    return s

# ═══════════════════════════════════════════════════════════════════
# COVER SLIDES
# ═══════════════════════════════════════════════════════════════════
def cover_noir(slide, T, data):
    """NOIR — Cinematic dark cover"""
    # Background gradient
    s_bg = rect(slide, 0,0,W,H, T["D"])
    gradient_fill(s_bg, T["DH"], T["MH"], angle=135)

    # Big decorative arcs (background layer)
    deco_arc_corner(slide, T, 'tr', size=9.0)
    deco_arc_corner(slide, T, 'bl', size=8.0)

    # Dots grid top-right
    deco_dots_grid(slide, W*0.55, 0.3, W*0.42, H*0.45, T["A"], 5, 7, 18)

    # Vertical accent bar — RIGHT
    bar_w = 1.20
    s_bar = rect(slide, W-bar_w, 0, bar_w, H, T["A"])
    gradient_fill(s_bar, T["AH"], T["A2H"], angle=90)
    lv(slide, W-bar_w, 0, H, T["A2"], w=0.06)

    # Top accent strip
    s_top = rect(slide, 0, 0, W-bar_w, 0.16, T["A"])
    gradient_fill(s_top, T["AH"], T["DH"], angle=0)

    # University band
    band_h = 3.20
    s_band = rect(slide, 0, 0.16, W-bar_w, band_h-0.16, T["M"])
    _set_fill_alpha(s_band, 85)
    lh(slide, 0, band_h, W-bar_w, T["A"], 0.10)
    lh(slide, 0, band_h+0.10, W-bar_w*0.5, T["A2"], 0.04)

    # Year — subtle background digit
    yr = safe(data.get("year","")).split("–")[-1].strip().split("-")[-1].strip()
    if yr and yr.isdigit():
        txt(slide, yr, 0.4, H*0.32, W*0.55, H*0.58,
            font="Calibri", size=240, bold=True,
            color=T["M"], align=PP_ALIGN.LEFT, rtl=False)

    # University & Faculty
    uni = safe(data.get("university",""))
    if uni:
        txt(slide, uni, MX, 0.30, W-bar_w-MX*2, 1.20,
            font=HF, size=18, bold=True,
            color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
    fac = " · ".join(filter(None,[safe(data.get("faculty","")),
                                   safe(data.get("department",""))]))
    if fac:
        txt(slide, fac, MX, 1.56, W-bar_w-MX*2, 0.72,
            font=BF, size=13, color=T["A"],
            align=PP_ALIGN.RIGHT, rtl=True)

    # Level pill
    lvl = safe(data.get("level","ماستر 2"))
    accent_pill(slide, MX, 2.50, 4.20, 0.72, T,
                "مذكرة تخرج  ·  " + lvl)

    # Divider line + field overline
    lh(slide, MX, band_h+0.28, W-bar_w-MX*2, hx("FFFFFF"), 0.03)
    _set_fill_alpha(slide.shapes[-1], 20)
    overline(slide, safe(data.get("fieldEn","Research Thesis")),
             MX, band_h+0.42, W-bar_w-MX*2, T["A"], align=PP_ALIGN.LEFT)
    # accent dots on overline
    for i in range(3):
        oval(slide, MX+0.06+i*0.36, band_h+0.44, 0.16, 0.16, T["A"])

    # Title
    lh(slide, MX, band_h+0.96, 3.0, T["A"], 0.12)
    lh(slide, MX+3.06, band_h+0.96, W-bar_w-MX*2-3.06, T["M"], 0.06)
    txt(slide, safe(data.get("titleAr","")),
        MX, band_h+1.18, W-bar_w-MX*2, 5.60,
        font=HF, size=28, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)

    # French title
    if data.get("titleFr"):
        lh(slide, MX, H-1.70, W-bar_w-MX*2, hx("FFFFFF"), 0.03)
        _set_fill_alpha(slide.shapes[-1], 18)
        txt(slide, safe(data.get("titleFr","")),
            MX, H-1.60, W-bar_w-MX*2, 0.72,
            font="Calibri", size=13, italic=True,
            color=T["A"], align=PP_ALIGN.LEFT, rtl=False)

    # Bottom info bar
    bar_y = H-0.96
    s_bot = rect(slide, 0, bar_y, W-bar_w, 0.96, T["M"])
    gradient_fill(s_bot, T["MH"], T["DH"], angle=0)
    lh(slide, 0, bar_y, W-bar_w, T["A"], 0.08)

    hw = (W-bar_w-MX*2) / 2
    txt(slide, "إعداد: " + safe(data.get("studentName","")),
        MX, bar_y+0.18, hw, 0.64,
        font=BF, size=15, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
    txt(slide, "إشراف: " + safe(data.get("supervisor","")),
        MX+hw+0.30, bar_y+0.18, hw-0.30, 0.64,
        font=BF, size=15, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)

    # Year tag on bar
    if data.get("year"):
        accent_pill(slide, W-bar_w-4.20, bar_y+0.24, 3.60, 0.50, T,
                    safe(data.get("year","")), text_color=T["D"])


def cover_vivid(slide, T, data):
    """VIVID — Bold split + strong typography"""
    bg(slide, T["D"])

    SPLIT = W * 0.44

    # Left panel gradient
    s_left = rect(slide, 0,0, SPLIT, H, T["M"])
    gradient_fill(s_left, T["MH"], T["DH"], angle=160)

    # Right panel
    s_right = rect(slide, SPLIT,0, W-SPLIT, H, T["D"])
    gradient_fill(s_right, T["DH"], T["MH"], angle=340)

    # Vertical accent divider
    s_div = rect(slide, SPLIT-0.12, 0, 0.24, H, T["A"])
    gradient_fill(s_div, T["AH"], T["A2H"], angle=90)
    shadow(s_div, blur=20, dist=0, angle=0, alpha=0.35, color=T["AH"].lstrip('#'))

    # Decorative arcs
    deco_arc_corner(slide, T, 'tl', size=7.0)
    deco_arc_corner(slide, T, 'br', size=6.0)

    # Dots on right side
    deco_dots_grid(slide, SPLIT+0.40, H*0.04, W-SPLIT-0.60, H*0.40, T["A"], 4,5, 20)

    # Top accent bar
    s_top = rect(slide, 0,0, W, 0.18, T["A"])
    gradient_fill(s_top, T["AH"], T["A2H"], angle=0)

    # ── LEFT: info ──────────────────────────────────────
    uni = safe(data.get("university",""))
    if uni:
        txt(slide, uni, MX, 0.38, SPLIT-MX-0.36, 2.0,
            font=HF, size=16, bold=True,
            color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
    fac = " · ".join(filter(None,[safe(data.get("faculty","")),
                                   safe(data.get("department",""))]))
    if fac:
        txt(slide, fac, MX, 2.46, SPLIT-MX-0.36, 0.72,
            font=BF, size=12, color=T["A"],
            align=PP_ALIGN.RIGHT, rtl=True)

    # Level pill
    lvl = safe(data.get("level","ماستر 2"))
    accent_pill(slide, MX, 3.40, min(3.8, SPLIT-MX-0.36), 0.72, T, lvl)

    # Wave lines decoration
    deco_wave_lines(slide, MX, H*0.54, SPLIT-MX-0.36, T["A"])

    # Student name LARGE
    sname = safe(data.get("studentName",""))
    if sname:
        txt(slide, sname, MX, H-3.00, SPLIT-MX-0.36, 1.0,
            font=HF, size=24, bold=True,
            color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
    lh(slide, MX, H-1.90, SPLIT-MX-0.36, T["A"], 0.06)
    txt(slide, "إشراف: " + safe(data.get("supervisor","")),
        MX, H-1.76, SPLIT-MX-0.36, 0.68,
        font=BF, size=13, color=T["A"],
        align=PP_ALIGN.RIGHT, rtl=True)
    yr = safe(data.get("year",""))
    if yr:
        accent_pill(slide, MX, H-0.96, 3.0, 0.60, T, yr, text_color=T["D"])

    # ── RIGHT: title ────────────────────────────────────
    RX = SPLIT + 0.50
    RW = W - RX - MX*0.6

    overline(slide, safe(data.get("fieldEn","Research Thesis")),
             RX, 0.32, RW, T["A"], align=PP_ALIGN.LEFT)
    lh(slide, RX, 0.82, RW, T["A"], 0.12)
    lh(slide, RX, 0.94, RW*0.50, T["A2"], 0.06)

    txt(slide, safe(data.get("titleAr","")),
        RX, 1.18, RW, 6.0,
        font=HF, size=26, bold=True,
        color=T["TD"], align=PP_ALIGN.RIGHT, rtl=True)

    if data.get("titleFr"):
        lh(slide, RX, H-1.30, RW, hx("C8D8E8"), 0.04)
        txt(slide, safe(data.get("titleFr","")),
            RX, H-1.18, RW, 0.80,
            font="Calibri", size=12, italic=True,
            color=T["D"], align=PP_ALIGN.LEFT, rtl=False)

    kw = safe(data.get("keywords",""))
    if kw:
        txt(slide, "🔑 " + kw, RX, H-0.76, RW, 0.60,
            font=BF, size=10, italic=True,
            color=hx("7A9AB8"), align=PP_ALIGN.LEFT, rtl=False)


def cover_minimal(slide, T, data):
    """MINIMAL — White premium + strong left sidebar"""
    bg(slide, hx("F7F9FD"))

    # Gradient left sidebar
    s_side = rect(slide, 0, 0, 1.10, H, T["D"])
    gradient_fill(s_side, T["DH"], T["MH"], angle=90)

    # Accent stripe next to sidebar
    s_acc = rect(slide, 1.10, 0, 0.18, H, T["A"])
    gradient_fill(s_acc, T["AH"], T["A2H"], angle=90)

    # Decorative arcs (light mode)
    deco_arc_corner(slide, T, 'br', size=7.0)
    deco_arc_corner(slide, T, 'tl', size=5.0)

    # Dots grid
    deco_dots_grid(slide, W*0.58, H*0.08, W*0.38, H*0.72, T["CE"], 5,7, 50)

    # Bottom accent bar
    s_bot = rect(slide, 1.28, H-1.50, W-1.28, 1.50, T["D"])
    gradient_fill(s_bot, T["DH"], T["MH"], angle=0)
    lh(slide, 1.28, H-1.50, W-1.28, T["A"], 0.12)

    # Rotated text in sidebar
    txt(slide, "MÉMOIRE DE MASTER", 0.18, 1.5, H-3.0, 0.62,
        font="Calibri", size=9, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.CENTER, rtl=False)

    # Content area
    CX = 1.60
    CW = W - CX - 1.20

    # University
    uni = safe(data.get("university",""))
    if uni:
        txt(slide, uni, CX, 0.30, CW, 1.40,
            font=HF, size=16, bold=True,
            color=T["D"], align=PP_ALIGN.RIGHT, rtl=True)
    fac = " · ".join(filter(None,[safe(data.get("faculty","")),
                                   safe(data.get("department",""))]))
    if fac:
        txt(slide, fac, CX, 1.76, CW, 0.68,
            font=BF, size=13, color=T["A"],
            align=PP_ALIGN.RIGHT, rtl=True)

    # Accent divider with gradient
    s_div = gradient_rect(slide, CX, 2.66, 5.60, 0.16, T["AH"], T["A2H"], angle=0)
    s_div2 = rect(slide, CX+5.60, 2.66, CW-5.60, 0.08, T["CE"])

    # Title LARGE
    txt(slide, safe(data.get("titleAr","")),
        CX, 3.00, CW, 5.20,
        font=HF, size=32, bold=True,
        color=T["D"], align=PP_ALIGN.RIGHT, rtl=True)

    if data.get("titleFr"):
        lh(slide, CX, H-2.00, CW, T["CE"], 0.05)
        txt(slide, safe(data.get("titleFr","")),
            CX, H-1.88, CW, 0.80,
            font="Calibri", size=13, italic=True,
            color=T["A"], align=PP_ALIGN.LEFT, rtl=False)

    # Bottom bar content
    hw = (W - 1.28 - CX*2) / 2
    txt(slide, "إعداد: " + safe(data.get("studentName","")),
        CX, H-1.30, hw, 0.72,
        font=BF, size=15, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
    txt(slide, "إشراف: " + safe(data.get("supervisor","")),
        CX+hw+0.36, H-1.30, hw-0.36, 0.72,
        font=BF, size=15, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)

    # Level + year pills
    accent_pill(slide, W-5.80, H-0.70, 2.40, 0.48, T,
                safe(data.get("level","ماستر 2")), text_color=T["D"])
    if data.get("year"):
        accent_pill(slide, W-3.20, H-0.70, 2.0, 0.48, T,
                    safe(data.get("year","")), text_color=T["D"])


def make_cover(prs, data, T):
    slide = blank(prs)
    fam = T["FAM"]
    if fam == "VIVID": cover_vivid(slide, T, data)
    elif fam == "MINIMAL": cover_minimal(slide, T, data)
    else: cover_noir(slide, T, data)
    return slide

# ═══════════════════════════════════════════════════════════════════
# SLIDE HEADER — unified premium header
# ═══════════════════════════════════════════════════════════════════
def slide_header(slide, T, title_ar, sub_en="", dark=True):
    H_HDR = 2.20

    if dark:
        # Gradient background header area
        s_bg = rect(slide, 0,0, W, H_HDR, T["D"])
        gradient_fill(s_bg, T["DH"], T["MH"], angle=175)
        # full bg gradient
        s_full = rect(slide, 0, H_HDR, W, H-H_HDR, T["D"])
        gradient_fill(s_full, T["DH"], T["MH"]+"22", angle=180)

        # Top accent strip
        s_acc = rect(slide, 0,0, W, 0.18, T["A"])
        gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
        rect(slide, 0, 0.18, W*0.28, 0.06, T["A2"])

        # Decorative arcs
        deco_arc_corner(slide, T, 'tr', size=6.0)
        deco_dots_grid(slide, W*0.72, 0.26, W*0.26, H_HDR*0.80, T["A"], 3,5, 15)

        # Overline
        overline(slide, sub_en, MX, 0.30, W-MX*2, T["A"], align=PP_ALIGN.LEFT)

        # Title
        txt(slide, title_ar, MX, 0.68, W-MX*2, 1.22,
            font=HF, size=34, bold=True,
            color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)

        # Bottom divider of header
        s_dl = gradient_rect(slide, MX, H_HDR-0.14, 3.60, 0.12, T["AH"], T["DH"], 0)
        rect(slide, MX+3.60, H_HDR-0.14, W-MX*2-3.60, 0.06, T["M"])

    else:
        bg(slide, hx("F7F9FD"))
        s_acc = rect(slide, 0,0, W, 0.18, T["A"])
        gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
        rect(slide, 0, 0.18, W*0.40, 0.06, T["A2"])

        overline(slide, sub_en, MX, 0.30, W-MX*2, T["A"], align=PP_ALIGN.LEFT)
        txt(slide, title_ar, MX, 0.68, W-MX*2, 1.22,
            font=HF, size=34, bold=True,
            color=T["TD"], align=PP_ALIGN.RIGHT, rtl=True)

        s_dl = gradient_rect(slide, MX, H_HDR-0.14, 3.60, 0.12, T["AH"], T["DH"], 0)
        rect(slide, MX+3.60, H_HDR-0.14, W-MX*2-3.60, 0.06, T["CE"])

    return H_HDR

# ═══════════════════════════════════════════════════════════════════
# INTRO SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_intro(prs, data, T):
    slide = blank(prs)

    # BG
    s_bg = rect(slide, 0,0,W,H, T["D"])
    gradient_fill(s_bg, T["DH"], T["MH"], angle=145)

    # Accent bar top
    s_top = rect(slide, 0,0,W,0.18, T["A"])
    gradient_fill(s_top, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.22, 0.06, T["A2"])

    LW = W * 0.34
    # Left panel
    s_left = rect(slide, 0,0, LW, H, T["M"])
    gradient_fill(s_left, T["MH"], T["DH"], angle=160)

    # Left divider accent
    s_ld = rect(slide, LW-0.18, 0, 0.18, H, T["A"])
    gradient_fill(s_ld, T["AH"], T["A2H"], angle=90)
    shadow(s_ld, blur=16, dist=0, alpha=0.30, color=T["AH"])

    # Decorative circles in left panel
    oval(slide, LW*0.10, H*0.10, LW*1.20, LW*1.20, T["D"], alpha=12)
    oval(slide, LW*0.55, H*0.58, LW*0.70, LW*0.70, T["A"], alpha=8)

    # Dots grid
    deco_dots_grid(slide, LW*0.08, H*0.08, LW*0.86, H*0.40, T["A"], 4,4, 20)

    # Icon large
    txt(slide, "📖", LW*0.10, H*0.22, LW*0.80, H*0.40,
        font="Segoe UI Emoji", size=64, align=PP_ALIGN.CENTER)

    # "مقدمة" title
    txt(slide, "مقدمة", LW*0.04, H*0.64, LW*0.92, H*0.22,
        font=HF, size=30, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.CENTER)
    lh(slide, LW*0.20, H*0.88, LW*0.60, T["A"], 0.08)
    overline(slide, "INTRODUCTION", LW*0.04, H*0.90, LW*0.92, T["A"], align=PP_ALIGN.CENTER)

    # Right column content
    RX = LW + 0.40
    RW = W - RX - MX

    overview = safe(data.get("introOverview",""))
    approach = safe(data.get("introApproach",""))

    if overview:
        s_lbl = rrect(slide, RX, 0.32, 3.60, 0.60, T["A"], r_pct=50)
        if s_lbl: shadow(s_lbl, blur=10, dist=2, alpha=0.25)
        txt(slide, "نظرة عامة", RX, 0.32, 3.60, 0.60,
            font=BF, size=12, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        # Big quote mark
        txt(slide, "\u201c", RX, 1.08, 1.10, 1.10,
            font="Georgia", size=60, bold=True, color=T["A"],
            align=PP_ALIGN.LEFT, rtl=False)

        oh = H*0.44 if approach else H-2.0
        txt(slide, overview, RX, 1.94, RW, oh,
            font=BF, size=16, color=hx("FFFFFF"),
            align=PP_ALIGN.RIGHT, rtl=True, spacing=22)

    if approach:
        lh(slide, RX, H*0.58, RW, T["A"], 0.04)
        s_lbl2 = rrect(slide, RX, H*0.61, 3.60, 0.60, T["M"], r_pct=50,
                       line_color=T["A"], line_w=1.0)
        txt(slide, "المقاربة النظرية", RX, H*0.61, 3.60, 0.60,
            font=BF, size=12, bold=True,
            color=T["A"], align=PP_ALIGN.CENTER)
        txt(slide, approach, RX, H*0.61+0.72, RW, H*0.28,
            font=BF, size=14, italic=True,
            color=hx("B8CCE0"), align=PP_ALIGN.RIGHT, rtl=True)

    return slide

# ═══════════════════════════════════════════════════════════════════
# PLAN SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_plan(prs, data, T, chapters_data):
    slide = blank(prs)

    # Full bg gradient
    s_bg = rect(slide,0,0,W,H, T["D"])
    gradient_fill(s_bg, T["DH"], T["MH"], angle=145)

    # Narrow header strip — no big arcs
    s_hdr = rect(slide, 0, 0, W, 2.10, T["M"])
    gradient_fill(s_hdr, T["MH"], T["DH"], angle=0)
    s_acc = rect(slide, 0, 0, W, 0.18, T["A"])
    gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.22, 0.06, T["A2"])

    # Overline + title
    overline(slide, "PLAN D'ÉTUDE · STUDY PLAN", MX, 0.30, W-MX*2, T["A"], align=PP_ALIGN.LEFT)
    txt(slide, "خطة الدراسة", MX, 0.60, W-MX*2, 1.10,
        font=HF, size=32, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
    gradient_rect(slide, MX, 1.82, 3.60, 0.12, T["AH"], T["DH"], 0)
    rect(slide, MX+3.60, 1.82, W-MX*2-3.60, 0.06, T["BO"])

    cy0 = 2.10

    chs = chapters_data[:4]
    n   = len(chs)
    if not n: return slide

    gx   = 0.36
    cw   = (W - MX*2 - gx*(n-1)) / n
    ch   = H - cy0 - 0.40

    for i, chap in enumerate(chs):
        cx  = MX + i*(cw+gx)
        sc  = T["SC"][i % len(T["SC"])]
        sch = T["SCH"][i % len(T["SCH"])]

        # Card bg
        s = rrect(slide, cx, cy0+0.20, cw, ch, T["CD"], r_pct=4)
        if s: shadow(s, blur=20, dist=6, alpha=0.25)

        # Gradient top bar
        gradient_rect(slide, cx, cy0+0.20, cw, 0.20, sch, T["DH"], 0)

        # Left accent
        s_lft = rect(slide, cx, cy0+0.20, 0.18, ch, sc)
        gradient_fill(s_lft, sch, T["MH"], angle=90)

        # Chapter number badge
        s_num = rrect(slide, cx+0.28, cy0+0.30, 1.30, 0.62, sc, r_pct=8)
        if s_num:
            gradient_fill(s_num, sch, T["DH"], angle=135)
            shadow(s_num, blur=10, dist=2, alpha=0.26)
        txt(slide, "F%d" % (i+1), cx+0.28, cy0+0.30, 1.30, 0.62,
            font="Calibri", size=20, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        # Chapter title
        lh(slide, cx+0.28, cy0+1.08, cw-0.42, sc, 0.04)
        txt(slide, safe(chap.get("title","")),
            cx+0.28, cy0+1.20, cw-0.42, 1.50,
            font=BF, size=14, bold=True,
            color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)

        # Sections
        secs = [s for s in chap.get("sections",[]) if s][:5]
        if secs:
            avail_h = ch - 2.90
            sh = max(avail_h / len(secs), 0.58)
            for j, sec in enumerate(secs):
                sy = cy0 + 0.20 + 2.80 + j*sh
                if sy + sh > cy0 + 0.20 + ch - 0.10: break
                oval(slide, cx+0.30, sy+sh*0.38, 0.16, 0.16, sc)
                txt(slide, safe(sec),
                    cx+0.54, sy+0.04, cw-0.70, sh-0.08,
                    font=BF, size=12, color=hx("B8CCE0"),
                    align=PP_ALIGN.RIGHT, rtl=True)

    return slide

# ═══════════════════════════════════════════════════════════════════
# PROBLEM SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_problem(prs, data, T):
    slide  = blank(prs)
    cy0    = slide_header(slide, T, "إشكالية البحث والتساؤلات",
                          "RESEARCH PROBLEM & QUESTIONS", dark=True)

    problem = safe(data.get("mainProblem",""))
    main_q  = safe(data.get("mainQuestion",""))
    subs    = [s for s in data.get("subQuestions",[]) if s][:5]

    if subs:
        lw = W*0.455 - MX - 0.14
        rw = W - MX*2 - lw - 0.36

        # Problem card
        qh = H - cy0 - (2.0 if main_q else 0.44) - 0.30
        s  = rrect(slide, MX, cy0+0.24, lw, qh, T["CD"], r_pct=10)
        if s: shadow(s, blur=20, dist=6, alpha=0.28)
        s_lft = rect(slide, MX, cy0+0.24, 0.18, qh, T["A"])
        gradient_fill(s_lft, T["AH"], T["MH"], angle=90)

        txt(slide, "\u201c", MX+0.30, cy0+0.30, 1.40, 1.10,
            font="Georgia", size=56, bold=True, color=T["A"])
        txt(slide, problem, MX+0.30, cy0+1.26, lw-0.44, qh-1.38,
            font=BF, size=14, color=hx("FFFFFF"),
            align=PP_ALIGN.RIGHT, rtl=True, spacing=20)

        # Main question card
        if main_q:
            mq_y = cy0+0.24+qh+0.20
            mq_h = H - mq_y - 0.28
            s2   = rrect(slide, MX, mq_y, lw, mq_h, T["A"], r_pct=10)
            if s2: shadow(s2, blur=16, dist=4, alpha=0.30)
            gradient_fill(s2, T["AH"], T["A2H"], angle=135)
            txt(slide, "؟", MX+0.14, mq_y+0.12, 1.0, mq_h-0.24,
                font="Georgia", size=56, bold=True,
                color=T["D"], align=PP_ALIGN.CENTER)
            txt(slide, main_q, MX+1.22, mq_y+0.18, lw-1.36, mq_h-0.36,
                font=BF, size=15, bold=True,
                color=T["D"], align=PP_ALIGN.RIGHT, rtl=True, spacing=20)

        # Sub-questions
        rx    = MX + lw + 0.36
        avail = H - cy0 - 0.44
        rh    = max(1.10, (avail - 0.16*(len(subs)-1)) / len(subs))

        for i, q in enumerate(subs):
            ry  = cy0+0.24 + i*(rh+0.16)
            sc  = T["SC"][i % len(T["SC"])]
            sch = T["SCH"][i % len(T["SCH"])]

            s  = rrect(slide, rx, ry, rw, rh, T["CD"], r_pct=9)
            if s: shadow(s, blur=14, dist=4, alpha=0.22)
            # top gradient strip
            gradient_rect(slide, rx, ry, rw, 0.16, sch, T["DH"], 0)

            # Numbered circle
            oval(slide, rx+0.24, ry+(rh-0.80)/2, 0.80, 0.80, sc)
            shadow(slide.shapes[-1], blur=10, dist=2, alpha=0.25)
            txt(slide, str(i+1), rx+0.24, ry+(rh-0.80)/2, 0.80, 0.80,
                font="Calibri", size=18, bold=True,
                color=T["D"], align=PP_ALIGN.CENTER)
            txt(slide, q, rx+1.18, ry+0.14, rw-1.32, rh-0.28,
                font=BF, size=14, color=hx("FFFFFF"),
                align=PP_ALIGN.RIGHT, rtl=True)
    else:
        qh = H - cy0 - (2.0 if main_q else 0.44) - 0.28
        s  = rrect(slide, MX, cy0+0.24, W-MX*2, qh, T["CD"], r_pct=10)
        if s: shadow(s, blur=20, dist=6, alpha=0.28)
        rect(slide, MX, cy0+0.24, 0.20, qh, T["A"])
        txt(slide, "\u201c", MX+0.30, cy0+0.28, 1.60, 1.20,
            font="Georgia", size=64, bold=True, color=T["A"])
        txt(slide, problem, MX+0.30, cy0+1.36, W-MX*2-0.44, qh-1.48,
            font=BF, size=15, color=hx("FFFFFF"),
            align=PP_ALIGN.RIGHT, rtl=True, spacing=22)
        if main_q:
            mq_y = cy0+0.24+qh+0.20
            s2   = rrect(slide, MX, mq_y, W-MX*2, H-mq_y-0.26, T["A"], r_pct=9)
            if s2:
                shadow(s2, blur=16, dist=4, alpha=0.30)
                gradient_fill(s2, T["AH"], T["A2H"], angle=135)
            txt(slide, "؟", MX+0.16, mq_y+0.14, 1.10, H-mq_y-0.40,
                font="Georgia", size=54, bold=True, color=T["D"])
            txt(slide, main_q, MX+1.36, mq_y+0.18, W-MX*2-1.52, H-mq_y-0.36,
                font=BF, size=17, bold=True, color=T["D"],
                align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ═══════════════════════════════════════════════════════════════════
# OBJECTIVES SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_objectives(prs, data, T):
    slide = blank(prs)
    dark  = (T["FAM"] != "MINIMAL")

    if dark:
        s_bg = rect(slide, 0,0,W,H, T["D"])
        gradient_fill(s_bg, T["DH"], T["MH"], angle=145)
    else:
        bg(slide, hx("F7F9FD"))

    s_acc = rect(slide, 0,0,W,0.18, T["A"])
    gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.28, 0.06, T["A2"])

    objs  = [o for o in data.get("objectives",[]) if o][:4]
    hypos = [h for h in data.get("hypotheses",[]) if h][:3]

    overline(slide, "OBJECTIVES & HYPOTHESES", MX, 0.32, W-MX*2, T["A"], align=PP_ALIGN.LEFT)
    txt(slide, "أهداف البحث والفرضيات", MX, 0.66, W-MX*2, 1.10,
        font=HF, size=32, bold=True,
        color=hx("FFFFFF") if dark else T["TD"], align=PP_ALIGN.RIGHT, rtl=True)
    gradient_rect(slide, MX, 1.84, 4.20, 0.14, T["AH"], T["DH"], 0)
    rect(slide, MX+4.20, 1.84, W-MX*2-4.20, 0.07, T["BO"] if dark else T["CE"])

    cy0   = 2.10
    BODY_H= H - cy0 - 0.28
    OW    = W * 0.60

    n = len(objs)
    if n:
        line_y = cy0 + 0.70
        # Timeline line gradient
        gradient_rect(slide, MX, line_y, OW-MX, 0.08, T["AH"], T["DH"], 0)

        card_w = (OW - MX - 0.20*(n-1)) / max(n,1)
        card_h = BODY_H - 1.10

        for i, obj in enumerate(objs):
            cx = MX + i*(card_w+0.20)
            sc  = T["SC"][i % len(T["SC"])]
            sch = T["SCH"][i % len(T["SCH"])]

            # Timeline circle
            oval(slide, cx+card_w/2-0.56, line_y-0.48, 1.12, 1.12, sc)
            shadow(slide.shapes[-1], blur=14, dist=3, alpha=0.30)
            txt(slide, str(i+1),
                cx+card_w/2-0.56, line_y-0.48, 1.12, 1.12,
                font="Calibri", size=24, bold=True,
                color=T["D"], align=PP_ALIGN.CENTER)

            # Connector line
            lv(slide, cx+card_w/2-0.04, line_y+0.60, 0.40, sc, 0.08)

            # Card
            card_y = line_y + 1.0
            bg_c = T["CD"] if dark else hx("FFFFFF")
            s = rrect(slide, cx, card_y, card_w, card_h, bg_c, r_pct=10,
                      line_color=None if dark else hx("E2E8F0"))
            if s: shadow(s, blur=18, dist=5, alpha=0.22)
            gradient_rect(slide, cx, card_y, card_w, 0.18, sch, T["DH"], 0)

            icons = ["🎯","📊","🔍","💡"]
            txt(slide, icons[i%4], cx+card_w/2-0.38, card_y+0.24, 0.76, 0.68,
                font="Segoe UI Emoji", size=22, align=PP_ALIGN.CENTER)

            txt(slide, safe(obj),
                cx+0.16, card_y+1.06, card_w-0.32, card_h-1.20,
                font=BF, size=13,
                color=hx("FFFFFF") if dark else T["TD"],
                align=PP_ALIGN.CENTER, rtl=True, spacing=18)

    # Hypotheses panel
    HX_start = OW + 0.28
    HW = W - HX_start - MX*0.5

    s_panel = rrect(slide, HX_start, cy0+0.10, HW, BODY_H, T["M"], r_pct=10)
    if s_panel:
        shadow(s_panel, blur=20, dist=6, alpha=0.28)
        gradient_fill(s_panel, T["MH"], T["DH"], angle=160)
    gradient_rect(slide, HX_start, cy0+0.10, HW, 0.20, T["AH"], T["A2H"], 0)

    txt(slide, "💡", HX_start+HW/2-0.38, cy0+0.22, 0.76, 0.68,
        font="Segoe UI Emoji", size=22, align=PP_ALIGN.CENTER)
    txt(slide, "الفرضيات", HX_start+0.16, cy0+0.24, HW-0.32, 0.60,
        font=BF, size=16, bold=True,
        color=T["A"], align=PP_ALIGN.CENTER)
    lh(slide, HX_start+0.26, cy0+0.96, HW-0.52, T["A"], 0.06)

    nh = len(hypos)
    if nh:
        item_h = (BODY_H - 1.10) / nh
        for i, hy in enumerate(hypos):
            iy  = cy0 + 1.12 + i*item_h
            sc  = T["SC"][(i+2) % len(T["SC"])]
            sch = T["SCH"][(i+2) % len(T["SCH"])]
            s2  = rrect(slide, HX_start+0.20, iy+item_h*0.12, 0.64, 0.44, sc, r_pct=50)
            shadow(slide.shapes[-1], blur=8, dist=2, alpha=0.22)
            txt(slide, "H%d"%(i+1),
                HX_start+0.20, iy+item_h*0.12, 0.64, 0.44,
                font="Calibri", size=13, bold=True,
                color=T["D"], align=PP_ALIGN.CENTER)
            txt(slide, safe(hy),
                HX_start+0.16, iy+item_h*0.12+0.52,
                HW-0.32, item_h-0.68,
                font=BF, size=13,
                color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)
            if i < nh-1:
                lh(slide, HX_start+0.26, iy+item_h-0.10, HW-0.52, T["BO"], 0.03)

    return slide

# ═══════════════════════════════════════════════════════════════════
# IMPORTANCE SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_importance(prs, data, T):
    slide = blank(prs)
    dark  = (T["FAM"] == "NOIR")
    cy0   = slide_header(slide, T, "أهمية الدراسة وأسباب اختيارها",
                         "RESEARCH SIGNIFICANCE", dark=dark)

    items = [x for x in data.get("importance",[]) if x]
    if data.get("reasons"): items.append(data["reasons"])
    items = items[:6]
    if not items: return slide

    n    = len(items)
    cols = 2 if n > 3 else 1
    rows = math.ceil(n/cols)
    gx,gy= 0.36, 0.28
    aw  = W - MX*2
    ah  = H - cy0 - 0.40
    cw  = (aw - gx*(cols-1)) / cols
    ch  = (ah - gy*(rows-1)) / rows
    icons= ["🔬","💡","📊","🎯","🌐","⚡"]

    for i, item in enumerate(items):
        col = i%cols; row = i//cols
        cx  = MX + col*(cw+gx)
        cy  = cy0+0.22 + row*(ch+gy)
        sc  = T["SC"][i % len(T["SC"])]
        sch = T["SCH"][i % len(T["SCH"])]

        bg_c = T["CD"] if dark else hx("FFFFFF")
        s = rrect(slide, cx,cy,cw,ch, bg_c, r_pct=10,
                  line_color=None if dark else hx("E2E8F0"))
        if s: shadow(s, blur=18, dist=5, alpha=0.24)

        # Gradient top
        gradient_rect(slide, cx, cy, cw, 0.20, sch, T["DH"], 0)
        # Left accent
        s_l = rect(slide, cx, cy, 0.18, ch, sc)
        gradient_fill(s_l, sch, T["MH"], angle=90)

        # Icon
        txt(slide, icons[i%6], cx+0.28, cy+0.24, 1.0, 0.90,
            font="Segoe UI Emoji", size=28, align=PP_ALIGN.CENTER)

        # Number circle
        oval(slide, cx+cw-1.10, cy+0.22, 0.84, 0.84, sc)
        shadow(slide.shapes[-1], blur=10, dist=2, alpha=0.26)
        txt(slide, "%02d"%(i+1), cx+cw-1.10, cy+0.22, 0.84, 0.84,
            font="Calibri", size=20, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        lh(slide, cx+0.28, cy+1.22, cw-0.44, sc, 0.04)
        txt(slide, safe(item), cx+0.28, cy+1.36, cw-0.44, ch-1.50,
            font=BF, size=14,
            color=hx("FFFFFF") if dark else T["TD"],
            align=PP_ALIGN.RIGHT, rtl=True)

    return slide

# ═══════════════════════════════════════════════════════════════════
# METHODOLOGY SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_methodology(prs, data, T):
    slide = blank(prs)
    dark  = (T["FAM"] != "MINIMAL")
    cy0   = slide_header(slide, T, "المنهجية والعينة والمجالات",
                         "METHODOLOGY & SAMPLE", dark=dark)

    meth    = safe(data.get("methodology",""))
    stype   = safe(data.get("sampleType",""))
    ssize   = safe(data.get("sampleSize",""))
    tool_v  = safe(data.get("tool",""))
    axes    = [a for a in data.get("toolAxes",[]) if a][:4]
    spatial = safe(data.get("spatialScope",""))
    temporal= safe(data.get("temporalScope",""))
    human_s = safe(data.get("humanScope",""))
    sw      = safe(data.get("software",""))
    tests   = [t for t in data.get("statisticalTests",[]) if t][:4]

    boxes = []
    if meth: boxes.append(("🔬","المنهج المتبع",meth))
    if stype or ssize:
        boxes.append(("👥","العينة"," · ".join(filter(None,[stype,ssize]))))
    if tool_v:
        tv = tool_v+("\n"+" · ".join(axes) if axes else "")
        boxes.append(("📋","أداة الدراسة",tv))
    if spatial or temporal or human_s:
        scope = "\n".join(filter(None,[
            "📍 "+spatial if spatial else "",
            "🕐 "+temporal if temporal else "",
            "👤 "+human_s if human_s else ""]))
        boxes.append(("🌐","مجالات الدراسة",scope))
    if sw:
        swv = sw+(" · "+" · ".join(tests) if tests else "")
        boxes.append(("⚙️","البرنامج والاختبارات",swv))
    if data.get("dataSource"):
        boxes.append(("📂","مصدر البيانات",safe(data.get("dataSource",""))))

    if not boxes: return slide
    n    = len(boxes)
    cols = min(n,3)
    rows = math.ceil(n/cols)
    gx,gy= 0.32, 0.28
    bw  = (W-MX*2 - gx*(cols-1)) / cols
    bh  = (H-cy0-0.42 - gy*(rows-1)) / rows

    for i,(icon,lbl,val) in enumerate(boxes):
        col = i%cols; row = i//cols
        bx  = MX+col*(bw+gx); by = cy0+0.24+row*(bh+gy)
        sc  = T["SC"][i%len(T["SC"])]
        sch = T["SCH"][i%len(T["SCH"])]
        bg_c= T["CD"] if dark else hx("FFFFFF")

        s = rrect(slide, bx,by,bw,bh, bg_c, r_pct=10,
                  line_color=None if dark else hx("E2E8F0"))
        if s: shadow(s, blur=18, dist=5, alpha=0.24)
        gradient_rect(slide, bx, by, bw, 0.20, sch, T["DH"], 0)
        s_l = rect(slide, bx, by, 0.18, bh, sc)
        gradient_fill(s_l, sch, T["MH"], angle=90)

        txt(slide, icon, bx+0.28, by+0.24, 0.90, 0.76,
            font="Segoe UI Emoji", size=24, align=PP_ALIGN.CENTER)
        txt(slide, lbl, bx+1.26, by+0.20, bw-1.42, 0.64,
            font=BF, size=15, bold=True, color=sc,
            align=PP_ALIGN.RIGHT, rtl=True)
        lh(slide, bx+0.28, by+0.94, bw-0.44, sc, 0.04)
        txt(slide, safe(val), bx+0.28, by+1.08, bw-0.44, bh-1.22,
            font=BF, size=13,
            color=hx("FFFFFF") if dark else T["TD"],
            align=PP_ALIGN.RIGHT, rtl=True)

    return slide

# ═══════════════════════════════════════════════════════════════════
# KPI DASHBOARD
# ═══════════════════════════════════════════════════════════════════
def make_stats(prs, data, T):
    slide = blank(prs)
    cy0   = slide_header(slide, T,
                         "لوحة المؤشرات الإحصائية الرئيسية",
                         "KEY INDICATORS · DASHBOARD", dark=True)

    raw_stats = data.get("stats",[])
    stats = [
        {"label": str(s.get("label","")).strip()[:60],
         "value": str(s.get("value","")).strip()[:40],
         "sub":   str(s.get("sub","")).strip()[:50]}
        for s in raw_stats
        if str(s.get("label","")).strip() and str(s.get("value","")).strip()
    ]
    if not stats: return slide

    n    = min(len(stats),8)
    cols = min(n,4)
    rows = math.ceil(n/cols)
    gx,gy= 0.32, 0.30
    cw  = (W-MX*2 - gx*(cols-1)) / cols
    ch  = (H-cy0-0.40 - gy*(rows-1)) / rows
    y0  = cy0+0.24

    for i, s in enumerate(stats[:8]):
        col = i%cols; row = i//cols
        cx  = MX+col*(cw+gx); cy = y0+row*(ch+gy)
        sc  = T["SC"][i%len(T["SC"])]
        sch = T["SCH"][i%len(T["SCH"])]

        # Main card
        s_card = rrect(slide, cx, cy, cw, ch, T["CD"], r_pct=10)
        if s_card: shadow(s_card, blur=22, dist=6, alpha=0.28)

        # Gradient top bar
        gradient_rect(slide, cx, cy, cw, 0.22, sch, T["DH"], 0)

        # Wave decoration inside card
        deco_wave_lines(slide, cx+0.16, cy+ch-1.10, cw-0.32,
                        sc, n=3, spacing=0.22)

        # Value — LARGE
        v    = safe(s["value"])
        vsz  = clamp(52 - max(0,len(v)-4)*8, 28, 52)
        txt(slide, v, cx+0.16, cy+0.28, cw-0.32, ch*0.52,
            font="Calibri", size=vsz, bold=True, color=sc,
            align=PP_ALIGN.CENTER)

        # Divider
        lh(slide, cx+0.22, cy+ch*0.68, cw-0.44, sc, 0.04)

        # Label
        txt(slide, safe(s["label"]), cx+0.16, cy+ch*0.72, cw-0.32, ch*0.22,
            font=BF, size=13, color=hx("B8CCE0"),
            align=PP_ALIGN.CENTER)
        if s.get("sub"):
            txt(slide, safe(s["sub"]), cx+0.16, cy+ch*0.90, cw-0.32, ch*0.10,
                font="Calibri", size=10, italic=True, color=hx("6A8AB0"),
                align=PP_ALIGN.CENTER)

    return slide

# ═══════════════════════════════════════════════════════════════════
# RESULTS SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_results(prs, data, T):
    slide = blank(prs)
    dark  = (T["FAM"] != "MINIMAL")

    if dark:
        s_bg = rect(slide, 0,0,W,H, T["D"])
        gradient_fill(s_bg, T["DH"], T["MH"], angle=145)
    else:
        bg(slide, hx("F7F9FD"))

    s_acc = rect(slide, 0,0,W,0.18, T["A"])
    gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.28, 0.06, T["A2"])

    results = [r for r in data.get("mainResults",[]) if r][:6]
    if not results: return slide

    overline(slide, "RESEARCH FINDINGS", MX, 0.32, W-MX*2, T["A"], align=PP_ALIGN.LEFT)
    txt(slide, "أهم نتائج البحث", MX, 0.66, W-MX*2, 1.10,
        font=HF, size=32, bold=True,
        color=hx("FFFFFF") if dark else T["TD"], align=PP_ALIGN.RIGHT, rtl=True)
    gradient_rect(slide, MX, 1.84, 4.20, 0.14, T["AH"], T["DH"], 0)
    rect(slide, MX+4.20, 1.84, W-MX*2-4.20, 0.07, T["BO"] if dark else T["CE"])

    cy0  = 2.10
    n    = len(results)
    avail= H - cy0 - 0.24
    rh   = max(0.96, (avail - 0.14*(n-1)) / n)
    NW   = 2.0

    for i, res in enumerate(results):
        ry  = cy0 + i*(rh+0.14)
        sc  = T["SC"][i % len(T["SC"])]
        sch = T["SCH"][i % len(T["SCH"])]
        alt = (i % 2 == 0)

        bg_c = (T["CD"] if alt else T["M"]) if dark else \
               (hx("FFFFFF") if alt else hx("EEF4FF"))

        s = rrect(slide, MX, ry, W-MX*2, rh, bg_c, r_pct=7,
                  line_color=None if dark else hx("E2E8F0"), line_w=0.6)
        if s: shadow(s, blur=10, dist=3, alpha=0.16)

        # Number column — gradient bg
        s_num = rect(slide, MX, ry, NW, rh, sc)
        gradient_fill(s_num, sch, T["DH"], angle=90)
        num_sz = clamp(42 - max(0,n-4)*5, 26, 44)
        txt(slide, str(i+1), MX, ry, NW, rh,
            font="Calibri", size=num_sz, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        # Vertical separator
        lv(slide, MX+NW, ry+0.14, rh-0.28, sc, 0.04)

        txt(slide, safe(res),
            MX+NW+0.30, ry+0.16, W-MX*2-NW-0.44, rh-0.32,
            font=BF, size=15,
            color=hx("FFFFFF") if dark else T["TD"],
            align=PP_ALIGN.RIGHT, rtl=True, spacing=20)

    return slide

# ═══════════════════════════════════════════════════════════════════
# CONCLUSION SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_conclusion(prs, data, T):
    slide = blank(prs)

    # BG gradient
    s_bg = rect(slide, 0,0,W,H, T["D"])
    gradient_fill(s_bg, T["DH"], T["MH"], angle=145)

    # Decorative arcs
    oval(slide, W*0.60, -3.0, W*0.72, W*0.72, T["M"], alpha=16)
    oval(slide, -4.0,   H*0.42, W*0.62, W*0.62, T["M"], alpha=12)
    deco_dots_grid(slide, W*0.56, H*0.04, W*0.40, H*0.55, T["A"], 4,6, 14)

    # Top accent bar
    s_top = rect(slide, 0,0,W,0.18, T["A"])
    gradient_fill(s_top, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.24, 0.06, T["A2"])

    conclusion = safe(data.get("generalConclusion",""))

    # ── Left accent panel (30% width) ──────────────────────────
    LW = W * 0.28
    s_lp = rect(slide, 0, 0.24, LW, H-1.20, T["M"])
    gradient_fill(s_lp, T["MH"], T["DH"], angle=160)
    s_ld = rect(slide, LW-0.18, 0.24, 0.18, H-1.20, T["A"])
    gradient_fill(s_ld, T["AH"], T["A2H"], angle=90)
    shadow(s_ld, blur=14, dist=0, alpha=0.28, color=T["AH"])

    # Decorative in left panel
    oval(slide, LW*0.10, H*0.12, LW*1.10, LW*1.10, T["D"], alpha=14)
    deco_dots_grid(slide, LW*0.06, H*0.06, LW*0.88, H*0.46, T["A"], 4,3, 20)

    txt(slide, "الخاتمة", LW*0.04, H*0.42, LW*0.92, H*0.28,
        font=HF, size=26, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.CENTER)
    lh(slide, LW*0.18, H*0.72, LW*0.64, T["A"], 0.08)
    overline(slide, "CONCLUSION", LW*0.04, H*0.76, LW*0.92, T["A"], align=PP_ALIGN.CENTER)

    # Pill label
    s_pill = rrect(slide, LW*0.10, H*0.82, LW*0.80, 0.60, T["A"], r_pct=50)
    if s_pill:
        gradient_fill(s_pill, T["AH"], T["A2H"], angle=0)
        shadow(s_pill, blur=10, dist=2, alpha=0.28)
    txt(slide, "✦  استنتاجات  ✦", LW*0.10, H*0.82, LW*0.80, 0.60,
        font=BF, size=11, bold=True,
        color=T["D"], align=PP_ALIGN.CENTER)

    # ── Right content (70% width) ───────────────────────────────
    RX = LW + 0.30
    RW = W - RX - MX*0.6

    # Large opening quote
    txt(slide, "\u201c", RX, 0.32, 1.40, 1.10,
        font="Georgia", size=80, bold=True, color=T["A"],
        align=PP_ALIGN.LEFT, rtl=False)

    # Text fills the right column from quote to bottom bar
    txt_y = 1.30
    txt_h = H - txt_y - 1.20

    # Subtle panel
    s_card = rrect(slide, RX, txt_y, RW, txt_h, T["M"], r_pct=10)
    if s_card:
        _set_fill_alpha(s_card, 55)
        shadow(s_card, blur=18, dist=5, alpha=0.22)
    # Top gradient strip
    gradient_rect(slide, RX, txt_y, RW, 0.18, T["AH"], T["A2H"], 0)
    # Left accent
    s_cl = rect(slide, RX, txt_y, 0.20, txt_h, T["A"])
    gradient_fill(s_cl, T["AH"], T["A2H"], angle=90)

    txt(slide, conclusion,
        RX+0.40, txt_y+0.30, RW-0.56, txt_h-0.52,
        font=BF, size=19, color=hx("FFFFFF"),
        align=PP_ALIGN.RIGHT, rtl=True, spacing=28)

    # Closing quote bottom right
    txt(slide, "\u201d", RX+RW-1.50, txt_y+txt_h-1.0, 1.40, 1.0,
        font="Georgia", size=72, bold=True, color=T["A"],
        align=PP_ALIGN.LEFT, rtl=False)

    # ── Bottom info bar ──────────────────────────────────────────
    bar_y = H - 1.0
    s_bot = rect(slide, 0, bar_y, W, 1.0, T["A"])
    gradient_fill(s_bot, T["AH"], T["A2H"], angle=0)
    lh(slide, 0, bar_y, W, T["A2"], 0.08)

    student = safe(data.get("studentName",""))
    sup     = safe(data.get("supervisor",""))
    yr      = safe(data.get("year",""))
    info    = "  ·  ".join(filter(None,[
        "إعداد: "+student if student else "",
        "إشراف: "+sup if sup else "",
        yr]))
    txt(slide, info, MX, bar_y+0.12, W-MX*2, 0.76,
        font=BF, size=14, bold=True,
        color=T["D"], align=PP_ALIGN.CENTER)

    return slide

# ═══════════════════════════════════════════════════════════════════
# RECOMMENDATIONS SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_recommendations(prs, data, T):
    slide = blank(prs)
    dark  = (T["FAM"] != "MINIMAL")

    if dark:
        s_bg = rect(slide, 0,0,W,H, T["D"])
        gradient_fill(s_bg, T["DH"], T["MH"], angle=145)
    else:
        bg(slide, hx("F7F9FD"))

    s_acc = rect(slide, 0,0,W,0.18, T["A"])
    gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.28, 0.06, T["A2"])

    recs = [r for r in data.get("recommendations",[]) if r][:6]
    if not recs: return slide

    overline(slide, "RECOMMENDATIONS", MX, 0.32, W-MX*2, T["A"], align=PP_ALIGN.LEFT)
    txt(slide, "توصيات البحث", MX, 0.66, W-MX*2, 1.10,
        font=HF, size=32, bold=True,
        color=hx("FFFFFF") if dark else T["TD"], align=PP_ALIGN.RIGHT, rtl=True)
    gradient_rect(slide, MX, 1.84, 4.20, 0.14, T["AH"], T["DH"], 0)
    rect(slide, MX+4.20, 1.84, W-MX*2-4.20, 0.07, T["BO"] if dark else T["CE"])

    cy0  = 2.10
    n    = len(recs)
    cols = min(n,3)
    rows = math.ceil(n/cols)
    gx,gy= 0.30, 0.28
    cw   = (W-MX*2 - gx*(cols-1)) / cols
    ch   = (H-cy0-0.20 - gy*(rows-1)) / rows
    ICONS= ["💡","📌","🔧","🌐","📈","🤝"]

    for i, rec in enumerate(recs):
        col  = i%cols; row = i//cols
        cx   = MX+col*(cw+gx); cy = cy0+row*(ch+gy)
        sc   = T["SC"][i%len(T["SC"])]
        sch  = T["SCH"][i%len(T["SCH"])]
        bg_c = T["CD"] if dark else hx("FFFFFF")

        s = rrect(slide, cx,cy,cw,ch, bg_c, r_pct=4,
                  line_color=None if dark else hx("E2E8F0"), line_w=0.7)
        if s: shadow(s, blur=18, dist=5, alpha=0.24)

        gradient_rect(slide, cx, cy, cw, 0.22, sch, T["DH"], 0)
        # Left accent
        s_la = rect(slide, cx, cy, 0.18, ch, sc)
        gradient_fill(s_la, sch, T["MH"], angle=90)

        # Number badge
        s_num = rrect(slide, cx+0.28, cy+0.30, 0.96, 0.84, sc, r_pct=8)
        if s_num: shadow(s_num, blur=10, dist=2, alpha=0.26)
        gradient_fill(s_num, sch, T["DH"], angle=135)
        txt(slide, str(i+1), cx+0.28, cy+0.30, 0.96, 0.84,
            font="Calibri", size=24, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        # Icon top right
        txt(slide, ICONS[i%6], cx+cw-1.10, cy+0.28, 0.86, 0.76,
            font="Segoe UI Emoji", size=22, align=PP_ALIGN.CENTER)

        lh(slide, cx+0.28, cy+1.24, cw-0.46, sc, 0.04)

        txt(slide, safe(rec),
            cx+0.28, cy+1.36, cw-0.46, ch-1.50,
            font=BF, size=14,
            color=hx("FFFFFF") if dark else T["TD"],
            align=PP_ALIGN.RIGHT, rtl=True, spacing=18)

        # Arrow between cards
        if col < cols-1 and i < n-1:
            ax = cx+cw+gx*0.18
            ay = cy+ch/2-0.28
            txt(slide, "←", ax, ay, gx*0.64, 0.56,
                font="Calibri", size=16, bold=True,
                color=T["A"], align=PP_ALIGN.CENTER)

    return slide

# ═══════════════════════════════════════════════════════════════════
# FUTURE SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_future(prs, data, T):
    slide = blank(prs)
    s_bg  = rect(slide, 0,0,W,H, T["D"])
    gradient_fill(s_bg, T["DH"], T["MH"], angle=145)

    items = [f for f in data.get("futureWork",[]) if f][:4]
    if not items: return slide

    HEADER_H = H * 0.36

    # Arcs
    oval(slide, W*0.60, HEADER_H*0.08, HEADER_H*1.6, HEADER_H*1.6, T["M"], alpha=16)
    oval(slide, W*0.88, HEADER_H*0.80, HEADER_H*0.9, HEADER_H*0.9, T["A"], alpha=9)

    # Top bar
    s_acc = rect(slide, 0,0,W,0.18, T["A"])
    gradient_fill(s_acc, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0.18, W*0.26, 0.06, T["A2"])

    # Header text
    txt(slide, "🔭", MX, HEADER_H*0.14, 1.80, HEADER_H*0.76,
        font="Segoe UI Emoji", size=54, align=PP_ALIGN.CENTER)
    overline(slide, "FUTURE RESEARCH PERSPECTIVES",
             MX+1.96, 0.28, W-MX*2-1.96, T["A"], align=PP_ALIGN.LEFT)
    txt(slide, "آفاق البحث المستقبلية",
        MX+1.96, 0.62, W-MX*2-1.96, HEADER_H-0.76,
        font=HF, size=32, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.RIGHT, rtl=True)

    # Divider
    gradient_rect(slide, 0, HEADER_H, W, 0.14, T["AH"], T["DH"], 0)
    rect(slide, 0, HEADER_H+0.14, W*0.36, 0.06, T["A2"])

    # Cards
    cy0  = HEADER_H + 0.28
    n    = len(items)
    avail= H - cy0 - 0.20
    rh   = max(0.96, (avail - 0.16*(n-1)) / n)
    ICONS= ["🚀","🌐","💻","🔬"]

    for i, fut in enumerate(items):
        ry  = cy0 + i*(rh+0.16)
        sc  = T["SC"][i%len(T["SC"])]
        sch = T["SCH"][i%len(T["SCH"])]
        alt = (i%2==0)
        bg_c= T["CD"] if alt else T["M"]

        s = rrect(slide, MX, ry, W-MX*2, rh, bg_c, r_pct=8)
        if s:
            shadow(s, blur=12, dist=4, alpha=0.20)
            gradient_fill(s, T["MH"] if alt else T["DH"], T["DH"], angle=0)

        gradient_rect(slide, MX, ry, W-MX*2, 0.14, sch, T["DH"], 0)
        rect(slide, MX, ry, 0.18, rh, sc)

        txt(slide, ICONS[i%4], MX+0.28, ry+(rh-0.70)/2, 0.70, 0.70,
            font="Segoe UI Emoji", size=22, align=PP_ALIGN.CENTER)

        s2 = rrect(slide, W-MX-1.10, ry+(rh-0.46)/2, 0.90, 0.46, sc, r_pct=50)
        txt(slide, "0%d"%(i+1), W-MX-1.10, ry+(rh-0.46)/2, 0.90, 0.46,
            font="Calibri", size=13, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        txt(slide, safe(fut),
            MX+1.10, ry+0.14, W-MX*2-2.20, rh-0.28,
            font=BF, size=15, color=hx("FFFFFF"),
            align=PP_ALIGN.RIGHT, rtl=True)

    return slide

# ═══════════════════════════════════════════════════════════════════
# REFERENCES SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_references(prs, data, T):
    refs = [r for r in data.get("references",[]) if r][:6]
    if not refs: return
    slide = blank(prs)
    dark  = (T["FAM"]=="NOIR")
    cy0   = slide_header(slide, T, "أبرز المراجع والمصادر",
                         "KEY REFERENCES", dark=dark)
    n     = len(refs)
    avail = H - cy0 - 0.42
    rh    = max(1.0, (avail - 0.16*(n-1)) / n)

    for i, ref in enumerate(refs):
        ry  = cy0+0.24 + i*(rh+0.16)
        sc  = T["SC"][i%len(T["SC"])]
        sch = T["SCH"][i%len(T["SCH"])]
        alt = i%2==0
        bg_c= (T["CD"] if alt else T["M"]) if dark else \
              (hx("FFFFFF") if alt else hx("F4F7FF"))

        s = rrect(slide, MX,ry, W-MX*2,rh, bg_c, r_pct=7,
                  line_color=None if dark else hx("E2E8F0"))
        if s: shadow(s, blur=10, dist=3, alpha=0.18)
        rect(slide, MX,ry,0.18,rh, sc)

        # Number circle
        oval(slide, MX+0.28, ry+(rh-0.72)/2, 0.72, 0.72, sc)
        shadow(slide.shapes[-1], blur=10, dist=2, alpha=0.24)
        txt(slide, str(i+1), MX+0.28, ry+(rh-0.72)/2, 0.72, 0.72,
            font="Calibri", size=17, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER)

        txt(slide, safe(ref), MX+1.16, ry+0.14, W-MX*2-1.30, rh-0.28,
            font=BF, size=13,
            color=hx("FFFFFF") if dark else T["TD"],
            align=PP_ALIGN.RIGHT, rtl=True)

    return slide

# ═══════════════════════════════════════════════════════════════════
# THANK YOU SLIDE
# ═══════════════════════════════════════════════════════════════════
def make_final(prs, data, T):
    slide = blank(prs)

    s_bg = rect(slide, 0,0,W,H, T["D"])
    gradient_fill(s_bg, T["DH"], T["MH"], angle=145)

    # Big background arcs
    oval(slide, W*0.48, -W*0.30, W*0.80, W*0.80, T["M"], alpha=18)
    oval(slide, -W*0.28, H*0.30, W*0.72, W*0.72, T["A"], alpha=9)
    oval(slide, W*0.26, H*0.14, W*0.40, W*0.40, T["M"], alpha=14)
    oval(slide, W*0.82, H*0.65, W*0.30, W*0.30, hx(T["G2"]), alpha=16)

    deco_dots_grid(slide, W*0.52, H*0.04, W*0.44, H*0.56, T["A"], 4,7, 14)

    # Accent strips
    s_top = rect(slide, 0,0,W,0.84, T["A"])
    gradient_fill(s_top, T["AH"], T["A2H"], angle=0)
    rect(slide, 0, 0, W*0.30, 0.06, T["A2"])

    s_bot = rect(slide, 0, H-0.84, W, 0.84, T["A"])
    gradient_fill(s_bot, T["AH"], T["A2H"], angle=0)

    # Thank you — 3 languages
    txt(slide, "شكراً لحسن استماعكم",
        MX, H*0.22, W-MX*2, 2.0,
        font=HF, size=54, bold=True,
        color=hx("FFFFFF"), align=PP_ALIGN.CENTER, rtl=True)
    txt(slide, "Merci pour votre attention",
        MX, H*0.22+2.14, W-MX*2, 0.94,
        font="Calibri", size=26, italic=True,
        color=T["A"], align=PP_ALIGN.CENTER, rtl=False)
    txt(slide, "Thank you for your kind attention",
        MX, H*0.22+3.18, W-MX*2, 0.68,
        font="Calibri", size=16, italic=True,
        color=hx("7A9AB8"), align=PP_ALIGN.CENTER, rtl=False)

    # Decorative divider
    gradient_rect(slide, W*0.28, H*0.70, W*0.20, 0.08, T["AH"], T["DH"], 0)
    oval(slide, W/2-0.28, H*0.69-0.18, 0.56, 0.56, T["A"])
    shadow(slide.shapes[-1], blur=12, dist=2, alpha=0.30)
    gradient_rect(slide, W*0.52, H*0.70, W*0.20, 0.08, T["AH"], T["DH"], 0)

    # Student info
    student = safe(data.get("studentName",""))
    sup     = safe(data.get("supervisor",""))
    yr      = safe(data.get("year",""))
    info    = "  ·  ".join(filter(None,[
        "إعداد: "+student if student else "",
        "إشراف: "+sup if sup else ""]))
    if info:
        txt(slide, info, MX, H*0.75, W-MX*2, 0.70,
            font=BF, size=15, color=hx("B8CCE0"),
            align=PP_ALIGN.CENTER, rtl=True)
    if data.get("university"):
        txt(slide, safe(data.get("university","")),
            MX, H*0.82, W-MX*2, 0.60,
            font=BF, size=13, italic=True,
            color=T["A"], align=PP_ALIGN.CENTER, rtl=True)
    if yr:
        s_yr = rrect(slide, W/2-1.60, H-0.68, 3.20, 0.48, T["A"], r_pct=50)
        if s_yr:
            shadow(s_yr, blur=12, dist=2, alpha=0.28)
            gradient_fill(s_yr, T["AH"], T["A2H"], angle=0)
        txt(slide, yr, W/2-1.60, H-0.68, 3.20, 0.48,
            font="Calibri", size=13, bold=True,
            color=T["D"], align=PP_ALIGN.CENTER, rtl=False)

    return slide

# ═══════════════════════════════════════════════════════════════════
# ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════
def generate_presentation(data: dict, output_path: str) -> None:
    key = data.get("theme","navy_gold")
    T   = PALETTES.get(key, PALETTES["navy_gold"])

    prs = Presentation()
    prs.slide_width  = Cm(W)
    prs.slide_height = Cm(H)

    cfg  = data.get("slides",{})
    def show(k): return cfg.get(k, True)
    def fl(k):   return [x for x in data.get(k,[]) if x]

    make_cover(prs, data, T)

    if show("intro") and (data.get("introOverview") or data.get("introApproach")):
        make_intro(prs, data, T)

    chs = [c for c in data.get("chapters",[]) if c.get("title")]
    if show("plan") and chs:
        make_plan(prs, data, T, chs)

    if show("problem") and (data.get("mainProblem") or data.get("mainQuestion") or fl("subQuestions")):
        make_problem(prs, data, T)

    if show("objectives") and (fl("objectives") or fl("hypotheses")):
        make_objectives(prs, data, T)

    if show("importance") and (fl("importance") or data.get("reasons")):
        make_importance(prs, data, T)

    if show("methodology") and (data.get("methodology") or data.get("sampleType") or data.get("tool")):
        make_methodology(prs, data, T)

    stats = [s for s in data.get("stats",[]) if s.get("label") and s.get("value")]
    if show("kpi") and stats:
        make_stats(prs, data, T)

    if show("results") and fl("mainResults"):
        make_results(prs, data, T)

    if show("conclusion") and data.get("generalConclusion"):
        make_conclusion(prs, data, T)

    if show("recommendations") and fl("recommendations"):
        make_recommendations(prs, data, T)

    if show("future") and fl("futureWork"):
        make_future(prs, data, T)

    if show("references") and fl("references"):
        make_references(prs, data, T)

    if show("thankyou"):
        make_final(prs, data, T)

    prs.save(output_path)
    n = len(prs.slides._sldIdLst)
    print(f"✅  {n} slides [v10·{T['FAM']}·{key}] → {output_path}", file=sys.stderr)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generator_v10.py input.json output.pptx", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        payload = json.load(f)
    generate_presentation(payload, sys.argv[2])

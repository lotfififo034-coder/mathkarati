"""
Drawing Primitives — مذكرتي Pro v17.1
Low-level, deterministic shape/text builders.

FIXED BUGS vs v17.0:
- gradient_fill() now inserts gradFill in CORRECT OOXML position (before <a:ln>)
- shadow() now inserts effectLst in CORRECT position (after <a:ln>)
- _sort_spPr() enforces strict OOXML child order on every shape
- set_solid_alpha() exposed (no duplication with slides.py)
"""
from __future__ import annotations

from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# Slide dimensions (cm) — 16:9
W, H = 33.867, 19.05

# OOXML spPr required child order
_SPPR_ORDER = [
    qn('a:xfrm'), qn('a:prstGeom'), qn('a:custGeom'),
    qn('a:noFill'), qn('a:solidFill'), qn('a:gradFill'),
    qn('a:blipFill'), qn('a:pattFill'), qn('a:grpFill'),
    qn('a:ln'),
    qn('a:effectLst'), qn('a:effectDag'),
    qn('a:scene3d'), qn('a:sp3d'), qn('a:extLst'),
]
_SPPR_RANK = {tag: i for i, tag in enumerate(_SPPR_ORDER)}


def _sort_spPr(spPr) -> None:
    """Reorder <p:spPr> children to comply with OOXML schema."""
    children = list(spPr)
    children.sort(key=lambda el: _SPPR_RANK.get(el.tag, 99))
    for child in children:
        spPr.remove(child)
    for child in children:
        spPr.append(child)


def _get_spPr(shape):
    return shape._element.find(qn('p:spPr'))


# ── Unit helpers ─────────────────────────────────────────────────────
def cm(v: float) -> int:
    return int(Cm(v))

def pt(v: float) -> int:
    return int(Pt(v))


# ── Shape builders ───────────────────────────────────────────────────
def rect(slide, x, y, w, h, fill: RGBColor, line=None, line_w=0.5):
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(1, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = pt(line_w)
    else:
        s.line.fill.background()
    return s


def rrect(slide, x, y, w, h, fill: RGBColor, radius_pct=8, line=None, line_w=0.5):
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(5, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = pt(line_w)
    else:
        s.line.fill.background()
    try:
        adj = s.adjustments
        if adj and len(adj) > 0:
            adj[0] = max(0, min(50, radius_pct)) * 1000
    except Exception:
        pass
    return s


def oval(slide, x, y, w, h, fill: RGBColor, alpha=100):
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(9, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def bg(slide, color: RGBColor):
    rect(slide, 0, 0, W, H, color)

def hline(slide, x, y, w, color: RGBColor, thickness=0.08):
    rect(slide, x, y, w, thickness, color)

def vline(slide, x, y, h2, color: RGBColor, thickness=0.08):
    rect(slide, x, y, thickness, h2, color)


# ── XML fill helpers ─────────────────────────────────────────────────
def set_solid_alpha(shape, alpha_pct: int) -> None:
    """Set transparency on a solidFill shape (0=transparent, 100=opaque)."""
    try:
        spPr = _get_spPr(shape)
        srgb = spPr.find('.//' + qn('a:srgbClr'))
        if srgb is not None:
            for e in srgb.findall(qn('a:alpha')):
                srgb.remove(e)
            alp = etree.SubElement(srgb, qn('a:alpha'))
            alp.set('val', str(int(alpha_pct * 1000)))
    except Exception:
        pass


def gradient_fill(shape, c1: str, c2: str, angle: float = 90) -> None:
    """
    Apply linear gradient via XML, with correct OOXML element ordering.
    BUG FIX: previously appended gradFill AFTER <a:ln> — now sorted correctly.
    """
    try:
        spPr = _get_spPr(shape)
        # Remove all existing fill variants
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'),
                    qn('a:pattFill'), qn('a:blipFill'), qn('a:grpFill')]:
            for el in spPr.findall(tag):
                spPr.remove(el)

        # Build gradFill
        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))

        gs0 = etree.SubElement(gsLst, qn('a:gs'))
        gs0.set('pos', '0')
        etree.SubElement(gs0, qn('a:srgbClr')).set('val', c1.lstrip('#'))

        gs1 = etree.SubElement(gsLst, qn('a:gs'))
        gs1.set('pos', '100000')
        etree.SubElement(gs1, qn('a:srgbClr')).set('val', c2.lstrip('#'))

        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000)))
        lin.set('scaled', '0')

        spPr.append(grad)
        _sort_spPr(spPr)  # ← enforce correct order
    except Exception:
        pass


def gradient_rect(slide, x, y, w, h, c1: str, c2: str, angle=0):
    c1h = c1.lstrip('#')
    fill_color = RGBColor(int(c1h[0:2], 16), int(c1h[2:4], 16), int(c1h[4:6], 16))
    s = rect(slide, x, y, w, h, fill_color)
    if s:
        gradient_fill(s, c1, c2, angle)
    return s


def shadow(shape, blur=16, dist=5, angle=135, alpha=0.22, color="000000") -> None:
    """
    Add outer drop shadow via XML.
    BUG FIX: effectLst now inserted in correct OOXML position (after <a:ln>).
    """
    try:
        spPr = _get_spPr(shape)
        for old in spPr.findall(qn('a:effectLst')):
            spPr.remove(old)

        eLst = etree.Element(qn('a:effectLst'))
        shdw = etree.SubElement(eLst, qn('a:outerShdw'))
        shdw.set('blurRad', str(int(blur * 12700)))
        shdw.set('dist', str(int(dist * 12700)))
        shdw.set('dir', str(int(angle * 60000)))
        shdw.set('algn', 'tl')
        srgb = etree.SubElement(shdw, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))

        spPr.append(eLst)
        _sort_spPr(spPr)  # ← enforce correct order
    except Exception:
        pass


# ── Text ─────────────────────────────────────────────────────────────
def txt(slide, text: str, x, y, w, h,
        font="Cairo", size=14, bold=False, italic=False,
        color: RGBColor | None = None,
        align=PP_ALIGN.RIGHT,
        margin=0.12, rtl=True, spacing=None):
    if not text or w <= 0 or h <= 0:
        return None
    tb = slide.shapes.add_textbox(cm(x), cm(y), cm(w), cm(h))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = cm(margin)
    tf.margin_right = cm(margin)
    tf.margin_top = cm(0.04)
    tf.margin_bottom = cm(0.04)

    p = tf.paragraphs[0]
    p.alignment = align
    if spacing is not None:
        try:
            p.line_spacing = Pt(spacing)
        except Exception:
            pass

    run = p.add_run()
    run.text = str(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    return tb


def blank_slide(prs):
    """Add a completely blank slide (layout 6 = Blank)."""
    return prs.slides.add_slide(prs.slide_layouts[6])

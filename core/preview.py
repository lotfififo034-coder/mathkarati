"""
core/preview.py — توليد صور المعاينة من ملفات PPTX مع علامة مائية
"""
import base64
import io
import logging
import threading
from pathlib import Path
from typing import List, Optional

log = logging.getLogger(__name__)

# ── In-memory cache (presentation_id → list of base64 JPEG strings) ──────────
_cache: dict = {}
_cache_lock = threading.Lock()


def get_cached_preview(presentation_id: str) -> Optional[List[str]]:
    with _cache_lock:
        return _cache.get(presentation_id)


def set_cached_preview(presentation_id: str, slides: List[str]) -> None:
    with _cache_lock:
        _cache[presentation_id] = slides


def pptx_to_preview_images(pptx_path: str, watermark: bool = True) -> List[str]:
    """
    يحوّل ملف PPTX إلى قائمة صور JPEG مُرمَّزة بـ base64.
    يستخدم python-pptx + Pillow لرسم كل شريحة.
    إذا فشل التحويل الكامل يرجع قائمة فارغة.
    """
    try:
        return _render_via_pptx(pptx_path, watermark)
    except Exception as exc:
        log.warning(f"pptx_to_preview_images failed: {exc}")
        return []


def _render_via_pptx(pptx_path: str, watermark: bool) -> List[str]:
    """
    يرسم كل شريحة PPTX على صورة Pillow ويضيف علامة مائية.
    """
    from pptx import Presentation
    from pptx.util import Pt
    from PIL import Image, ImageDraw, ImageFont

    prs = Presentation(pptx_path)

    # أبعاد الشريحة بالـ EMU → بكسل (96 dpi)
    EMU_PER_INCH = 914400
    DPI = 96
    slide_w = int(prs.slide_width  / EMU_PER_INCH * DPI)
    slide_h = int(prs.slide_height / EMU_PER_INCH * DPI)

    # تأكد من حجم معقول
    if slide_w < 100: slide_w = 960
    if slide_h < 100: slide_h = 540

    results = []

    for slide_idx, slide in enumerate(prs.slides):
        try:
            img = _render_slide(slide, slide_w, slide_h)
            if watermark:
                img = _add_watermark(img)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75, optimize=True)
            results.append(base64.b64encode(buf.getvalue()).decode("ascii"))
        except Exception as exc:
            log.warning(f"Slide {slide_idx} render failed: {exc}")
            # أضف صورة placeholder بدلاً من إيقاف العملية
            results.append(_placeholder_slide(slide_w, slide_h, slide_idx + 1))

    return results


def _render_slide(slide, width: int, height: int):
    """يرسم شريحة واحدة على صورة Pillow."""
    from PIL import Image, ImageDraw, ImageFont
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor

    # لون خلفية الشريحة
    bg_color = _get_slide_bg_color(slide)
    img = Image.new("RGB", (width, height), bg_color)
    draw = ImageDraw.Draw(img)

    # ارسم كل شكل نصي
    for shape in slide.shapes:
        try:
            if not shape.has_text_frame:
                continue
            # موضع الشكل
            left   = int(shape.left   / 914400 * 96) if shape.left   else 0
            top    = int(shape.top    / 914400 * 96) if shape.top    else 0
            s_w    = int(shape.width  / 914400 * 96) if shape.width  else width
            s_h    = int(shape.height / 914400 * 96) if shape.height else 40

            y = top
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if not text:
                    y += 14
                    continue
                # حجم الخط
                font_size = 16
                for run in para.runs:
                    if run.font and run.font.size:
                        font_size = max(8, min(int(run.font.size / 12700), 72))
                        break
                # لون النص
                txt_color = (30, 30, 30)
                for run in para.runs:
                    try:
                        if run.font and run.font.color and run.font.color.rgb:
                            c = run.font.color.rgb
                            txt_color = (c.r, c.g, c.b)
                            break
                    except Exception:
                        pass
                try:
                    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
                except Exception:
                    font = ImageFont.load_default()
                draw.text((left + 4, y), text, fill=txt_color, font=font)
                y += font_size + 4
        except Exception:
            continue

    return img


def _get_slide_bg_color(slide) -> tuple:
    """يستخرج لون خلفية الشريحة أو يرجع أبيض."""
    try:
        bg = slide.background
        fill = bg.fill
        fill.fore_color  # trigger
        if fill.type is not None:
            rgb = fill.fore_color.rgb
            return (rgb.r, rgb.g, rgb.b)
    except Exception:
        pass
    return (255, 255, 255)


def _add_watermark(img) -> "Image":
    """يضيف علامة مائية نصية قطرية."""
    from PIL import Image, ImageDraw, ImageFont
    import math

    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    text = "مذكرتي Pro — معاينة فقط"
    w, h = img.size

    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", max(18, w // 25))
    except Exception:
        font = ImageFont.load_default()

    # ارسم العلامة عدة مرات بشكل قطري
    for i in range(-2, 4):
        x = w // 4 + i * (w // 4)
        y = h // 3 + i * (h // 6)
        draw.text((x, y), text, fill=(80, 80, 80, 55), font=font)

    # دمج
    base = img.convert("RGBA")
    combined = Image.alpha_composite(base, overlay)
    return combined.convert("RGB")


def _placeholder_slide(width: int, height: int, num: int) -> str:
    """يُنشئ صورة placeholder رمادية لشريحة فشل تحويلها."""
    from PIL import Image, ImageDraw
    import io, base64

    img = Image.new("RGB", (width, height), (30, 30, 50))
    draw = ImageDraw.Draw(img)
    draw.text((width // 2 - 40, height // 2 - 10), f"شريحة {num}", fill=(200, 200, 200))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=60)
    return base64.b64encode(buf.getvalue()).decode("ascii")

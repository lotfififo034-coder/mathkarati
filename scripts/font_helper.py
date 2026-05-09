"""
font_helper.py — اكتشاف الخطوط العربية المتاحة على الخادم
نسخة مُحسَّنة: تستخدم fc-list فقط عند الاستيراد (سريع جداً)
ومسح المجلدات فقط عند الحاجة وبعد timeout.
"""
import os, subprocess, shutil, threading

# ── Cache للنتائج حتى لا نفحص مرتين ──────────────────────────────────
_font_cache: dict = {}
_cache_lock = threading.Lock()

def _font_available_fast(name: str) -> bool:
    """
    فحص سريع عبر fc-list فقط — timeout قصير جداً (1 ثانية).
    هذا هو الفحص الوحيد الذي يُنفَّذ عند الاستيراد.
    """
    with _cache_lock:
        if name in _font_cache:
            return _font_cache[name]

    result = False

    # المحاولة الأولى: fc-list — سريع جداً (< 100ms عادةً)
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(
                ["fc-list", f":family={name}"],
                capture_output=True, text=True, timeout=2
            )
            if name.lower() in out.stdout.lower():
                result = True
        except Exception:
            pass

    # المحاولة الثانية: مسح ~/.fonts فقط (لا os.walk عام)
    if not result:
        home_fonts = os.path.expanduser("~/.fonts")
        tmp_fonts  = "/tmp/fonts"
        for base_dir in [home_fonts, tmp_fonts]:
            if os.path.isdir(base_dir):
                try:
                    for root, _, files in os.walk(base_dir):
                        for f in files:
                            if name.lower() in f.lower() and f.lower().endswith((".ttf", ".otf")):
                                result = True
                                break
                        if result:
                            break
                except Exception:
                    pass
            if result:
                break

    with _cache_lock:
        _font_cache[name] = result

    return result


def best_arabic_font() -> str:
    """يُعيد أفضل خط عربي متاح على الخادم."""
    for font in ["Cairo", "Amiri", "Scheherazade", "Noto Naskh Arabic", "Noto Sans Arabic"]:
        if _font_available_fast(font):
            return font
    return "Calibri"   # fallback أخير


def best_body_font() -> str:
    """يُعيد أفضل خط للنصوص العامة."""
    for font in ["Cairo", "Amiri", "Arial", "Calibri"]:
        if _font_available_fast(font):
            return font
    return "Arial"


# ── تقييم عند الاستيراد — سريع الآن ────────────────────────────────
CAIRO_OK    = _font_available_fast("Cairo")
AMIRI_OK    = _font_available_fast("Amiri")
ARABIC_FONT = best_arabic_font()
BODY_FONT   = best_body_font()

if __name__ == "__main__":
    print(f"Cairo:  {CAIRO_OK}")
    print(f"Amiri:  {AMIRI_OK}")
    print(f"Best Arabic font: {ARABIC_FONT}")
    print(f"Best body font:   {BODY_FONT}")

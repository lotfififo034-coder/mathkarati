"""
font_helper.py — اكتشاف الخطوط العربية المتاحة على الخادم
يُستخدم من generator_canva.py و generator_classic.py
"""
import os, subprocess, shutil

def _font_available(name: str) -> bool:
    """فحص إذا كان الخط متاحاً عبر fc-list أو مسح المجلدات."""
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(
                ["fc-list", f":family={name}"],
                capture_output=True, text=True, timeout=5
            )
            if name.lower() in out.stdout.lower():
                return True
        except Exception:
            pass

    # مجلدات إضافية — تشمل $HOME/.fonts الذي نثبّت فيه الخط
    extra_dirs = [
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        os.path.expanduser("~/.fonts"),
        "/tmp/fonts",
        "C:/Windows/Fonts",
    ]
    for d in extra_dirs:
        if os.path.isdir(d):
            for root, _, files in os.walk(d):
                for f in files:
                    if name.lower() in f.lower() and f.lower().endswith((".ttf", ".otf")):
                        return True
    return False


def best_arabic_font() -> str:
    """يُعيد أفضل خط عربي متاح على الخادم."""
    for font in ["Cairo", "Amiri", "Scheherazade", "Noto Naskh Arabic", "Noto Sans Arabic"]:
        if _font_available(font):
            return font
    return "Calibri"   # fallback أخير


def best_body_font() -> str:
    """يُعيد أفضل خط للنصوص العامة."""
    for font in ["Cairo", "Amiri", "Arial", "Calibri"]:
        if _font_available(font):
            return font
    return "Arial"


# ── تقييم عند الاستيراد ──────────────────────────────────────────────
CAIRO_OK  = _font_available("Cairo")
AMIRI_OK  = _font_available("Amiri")
ARABIC_FONT = best_arabic_font()
BODY_FONT   = best_body_font()

if __name__ == "__main__":
    print(f"Cairo:  {CAIRO_OK}")
    print(f"Amiri:  {AMIRI_OK}")
    print(f"Best Arabic font: {ARABIC_FONT}")
    print(f"Best body font:   {BODY_FONT}")

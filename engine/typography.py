"""
Typography System — مذكرتي Pro v17.2
نظام هرمية الخطوط — Type Scale للعروض الأكاديمية

DESIGN PHILOSOPHY (من منظور مشاهد):
─────────────────────────────────────
عندما ينظر الجمهور للشريحة يجب أن:
  1. يُدرك العنوان الرئيسي فوراً (أكبر + أثقل)
  2. يميّز عناوين الأقسام بسهولة
  3. يقرأ النص العادي بارتياح
  4. يلاحظ الأرقام الإحصائية بشكل مختلف تماماً
  5. يرى الملاحظات والتسميات كعنصر ثانوي

SCALE (pt):
  DISPLAY   → 32-40pt  — العنوان الرئيسي في غلاف
  H1        → 22-26pt  — عناوين الشرائح الرئيسية
  H2        → 16-18pt  — عناوين الأقسام والبطاقات
  H3        → 13-14pt  — عناوين الفقرات الفرعية
  BODY      → 11-12pt  — النص الشارح العادي
  SMALL     → 9-10pt   — الملاحظات والتسميات الثانوية
  STAT      → 28-48pt  — الأرقام الإحصائية (Calibri - Latin)
  CAPTION   → 8-9pt    — التسميات الرقمية تحت الإحصاء

FONTS:
  Arabic text  → Cairo (عناوين bold), Tajawal (نص عادي light)
  Numbers/Stats → Calibri Bold (حاد وواضح)
  English/mixed → Calibri

الفارق الصحيح:
  عنوان الشريحة يجب أن يكون ضعف حجم النص العادي على الأقل
  الأرقام الإحصائية يجب أن تكون 3-4x حجم النص العادي
"""

# ── Font families ──────────────────────────────────────────────────────
FONT_TITLE   = "Cairo"      # العناوين الرئيسية — ثقيل وواضح
FONT_HEADING = "Cairo"      # عناوين الأقسام
FONT_BODY    = "Cairo"      # نص شارح عربي
FONT_NUM     = "Calibri"    # أرقام وإحصاءات — Latin crisp
FONT_EN      = "Calibri"    # نص إنجليزي

# ── Type Scale (pt) ────────────────────────────────────────────────────
class TS:
    """Type Scale — استخدم هذه الثوابت بدلاً من الأرقام المتفرقة"""

    # ── Display — الغلاف فقط
    DISPLAY_LG  = 34   # عنوان رسالة قصير (< 40 حرف)
    DISPLAY_MD  = 26   # عنوان رسالة متوسط
    DISPLAY_SM  = 20   # عنوان رسالة طويل (> 80 حرف)

    # ── Headings — عناوين الشرائح
    H1          = 22   # عنوان الشريحة الرئيسي (أعلى الشريحة)
    H2          = 16   # عنوان القسم / عنوان البطاقة
    H3          = 13   # عنوان الفقرة الفرعية

    # ── Body — نص شارح
    BODY_LG     = 12   # نص رئيسي (إشكالية، ملخص)
    BODY        = 11   # نص عادي (قوائم، شروح)
    BODY_SM     = 10   # نص مضغوط (جداول كثيفة)

    # ── Supporting
    LABEL       = 9    # تسميات الصفوف / الأعمدة
    CAPTION     = 8    # الملاحظات الصغيرة، المصادر

    # ── Statistics — أرقام إحصائية بارزة
    STAT_XL     = 48   # رقم بطل (stat hero, slide 8 مثلاً)
    STAT_LG     = 36   # رقم بارز
    STAT_MD     = 28   # رقم متوسط
    STAT_SM     = 22   # رقم صغير في جدول

    # ── Sidebar
    SIDEBAR_LABEL = 13  # عنوان القسم في الشريط الجانبي
    SIDEBAR_ICON  = 38  # أيقونة إيموجي في الشريط

    # ── Cover info rows
    COVER_LABEL   = 10  # تسمية (الطالب:، المشرف:)
    COVER_VALUE   = 13  # قيمة (اسم الطالب، اسم المشرف)
    COVER_INST    = 10  # اسم المؤسسة
    COVER_YEAR    = 14  # السنة في الـ badge

    # ── Slide index / watermark
    SLIDE_NUM     = 9
    WATERMARK     = 8


# ── Dynamic title size (based on text length) ─────────────────────────
def display_size(text: str) -> int:
    """اختر حجم عنوان الغلاف بناءً على طول النص"""
    n = len(text)
    if n < 40:
        return TS.DISPLAY_LG
    elif n < 70:
        return TS.DISPLAY_MD
    else:
        return TS.DISPLAY_SM


def h1_size(text: str) -> int:
    """اختر حجم عنوان الشريحة بناءً على طول النص"""
    n = len(text)
    if n < 30:
        return TS.H1
    elif n < 55:
        return 19
    else:
        return 16


def stat_size(value: str) -> int:
    """اختر حجم الرقم الإحصائي بناءً على طول القيمة"""
    n = len(str(value))
    if n <= 4:
        return TS.STAT_LG
    elif n <= 6:
        return TS.STAT_MD
    else:
        return TS.STAT_SM


# ── Typography helpers ────────────────────────────────────────────────
def get_title_font(arabic_font: str) -> str:
    """Font للعناوين — يستخدم Cairo دائماً بغض النظر عن _FONT"""
    return "Cairo"  # Cairo أوضح في العناوين


def get_body_font(arabic_font: str) -> str:
    """Font للنص العادي — يستخدم _FONT المحدد"""
    return arabic_font


# ── Spacing ratios ────────────────────────────────────────────────────
# نسب الارتفاع الموصى بها لكل عنصر
class Heights:
    SLIDE_TITLE_ROW = 0.95   # ارتفاع صف عنوان الشريحة
    SECTION_HEADER  = 0.75   # ارتفاع عنوان القسم/البطاقة
    BODY_LINE       = 0.55   # ارتفاع سطر نص عادي
    STAT_ROW        = 2.0    # ارتفاع صف إحصاء رئيسي
    LABEL_ROW       = 0.45   # ارتفاع تسمية صغيرة


# ── Quick reference (للمطور) ──────────────────────────────────────────
"""
استخدام سريع:

from engine.typography import TS, FONT_TITLE, FONT_BODY, FONT_NUM, display_size

# عنوان الغلاف:
txt(..., font=FONT_TITLE, size=display_size(req.title_ar), bold=True, ...)

# عنوان الشريحة (H1):
txt(..., font=FONT_TITLE, size=TS.H1, bold=True, ...)

# عنوان بطاقة (H2):
txt(..., font=FONT_TITLE, size=TS.H2, bold=True, ...)

# نص شارح:
txt(..., font=FONT_BODY, size=TS.BODY, bold=False, ...)

# رقم إحصائي:
txt(..., font=FONT_NUM, size=TS.STAT_LG, bold=True, ...)

# تسمية تحت الرقم:
txt(..., font=FONT_BODY, size=TS.CAPTION, bold=False, ...)
"""

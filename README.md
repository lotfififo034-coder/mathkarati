# مذكرتي Pro v10 — Ultra Visual Engine 🎓

منشئ عروض PowerPoint أكاديمية احترافية للجامعات الجزائرية  
**محرك v10 · 9 ألوان · 14 شريحة · 3 لغات · تدرجات لونية حقيقية**

---

## 🏗️ هيكل المشروع

```
mathkarati-pro/
├── app.py                      ← Flask server (router للمحركات)
├── requirements.txt            ← Python deps (flask, gunicorn, python-pptx, lxml)
├── Procfile                    ← gunicorn start command
├── render.yaml                 ← Render.com config
├── build.sh                    ← Build script (تثبيت Cairo font + npm)
├── scripts/
│   ├── generator_canva.py      ← ✨ Ultra v10 engine (9 palettes × 3 families)
│   └── generator_classic.py   ← Classic engine (fallback)
├── node_scripts/
│   ├── generator_api.js        ← Premium engine (Node.js / pptxgenjs)
│   └── package.json            ← pptxgenjs dependency
└── public/
    └── index.html              ← واجهة 6 خطوات · 3 لغات · Toggle للشرائح
```

---

## 🚀 النشر على Render.com

### الطريقة السريعة:
1. ارفع المشروع على GitHub
2. في Render: **New → Web Service** → اختر الـ repo
3. الإعدادات:
   - **Runtime:** Python 3
   - **Build Command:** `bash build.sh`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
4. Environment Variables:
   - `FLASK_ENV` = `production`
   - `NODE_VERSION` = `20.11.0`

### أو استخدم render.yaml (تلقائي):
render.yaml موجود بالفعل في المشروع ويُطبَّق تلقائياً.

---

## ⚠️ ملاحظات مهمة للإنتاج

### خط Cairo:
- `build.sh` يُحاول تنزيل خط Cairo تلقائياً
- إذا فشل: يستخدم **Calibri** تلقائياً (fallback مضمون)
- العرض يعمل بشكل طبيعي في كلتا الحالتين

### المحركات:
| المحرك | المتطلب | Fallback |
|--------|---------|---------|
| Canva (v10) | Python + lxml | — |
| Classic | Python | — |
| Premium | Node.js | → Canva تلقائياً |

### الأداء على Render Free Tier:
- توليد شريحة: ~5–15 ثانية
- Timeout مضبوط على 120 ثانية (كافٍ)
- `--max-requests 50` يمنع تراكم الذاكرة

---

## 🎨 الألوان المدعومة (v10)

| الاسم | العائلة | الطابع |
|-------|---------|--------|
| `navy_gold` | NOIR | أزرق ملكي ذهبي |
| `midnight_purple` | NOIR | بنفسجي ليلي |
| `forest` | NOIR | أخضر غابي |
| `sand_gold` | NOIR | ذهبي رملي |
| `dark_teal` | VIVID | تيل حيوي |
| `charcoal_orange` | VIVID | فحمي برتقالي |
| `burgundy` | VIVID | بورغندي |
| `ice_blue` | MINIMAL | أزرق ثلجي |
| `slate_crimson` | MINIMAL | فولاذي قرمزي |

---

## 💻 التشغيل محلياً

```bash
git clone <repo-url>
cd mathkarati-v2

# تثبيت Python
pip install -r requirements.txt

# تثبيت Node (اختياري - للمحرك Premium)
cd node_scripts && npm install && cd ..

# تشغيل
python app.py
```

ثم: http://localhost:5000

---

## 🔧 استكشاف الأخطاء

| المشكلة | الحل |
|---------|------|
| `ImportError: lxml` | `pip install lxml==5.3.0` |
| `ImportError: pptx` | `pip install python-pptx==1.0.2` |
| الخط عربي لا يظهر | Cairo غير مثبت — شغّل `build.sh` أو ثبّت يدوياً |
| `500 Internal Server Error` | راجع `/health` للتشخيص |
| ملف PPTX فارغ | تأكد من إرسال `studentName` و`titleAr` |

---

*مذكرتي Pro v10 — Ultra Visual Engine · 2024–2025*

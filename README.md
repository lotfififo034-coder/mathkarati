# مذكرتي Pro v7 — Canva Level 🎓

منشئ عروض PowerPoint أكاديمية احترافية للجامعات الجزائرية  
**3 محركات · 11 لوناً · 14 شريحة · 3 لغات**

---

## 🏗️ هيكل المشروع

```
mathkarati-pro/
├── app.py                      ← Flask server (3 engines router)
├── requirements.txt            ← Python deps (flask, gunicorn, python-pptx, lxml)
├── Procfile                    ← gunicorn start command
├── render.yaml                 ← Render.com config
├── build.sh                    ← Build script
├── scripts/
│   ├── generator_canva.py      ← ✨ Canva Level engine (8 palettes × 3 families)
│   └── generator_classic.py   ← Classic engine (8 palettes × 3 layouts)
├── node_scripts/
│   ├── generator_api.js        ← Premium engine (3 styles: Noir/Atlas/Sakura)
│   └── package.json            ← pptxgenjs dependency
└── public/
    └── index.html              ← واجهة 6 خطوات · 3 لغات · Toggle للشرائح
```

---

## 🚀 النشر على Render

### Build Command:
```
pip install -r requirements.txt && cd node_scripts && npm install --production
```

### Start Command:
```
gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
```

### Runtime: **Python**

> ملاحظة: Render يثبّت Node.js تلقائياً مع Python services.

---

## 🎨 المحركات والألوان

### ✨ Canva Level (الافتراضي) — python-pptx
| اللون | العائلة | الطابع |
|-------|---------|--------|
| navy_gold | NOIR | أزرق ملكي ذهبي |
| midnight_purple | NOIR | بنفسجي ليلي |
| forest | NOIR | أخضر غابي |
| sand_gold | NOIR | ذهبي رملي |
| dark_teal | VIVID | تيل حيوي |
| charcoal_orange | VIVID | فحمي برتقالي |
| burgundy | VIVID | بورغندي |
| ice_blue | MINIMAL | أزرق ثلجي نظيف |

### Classic — python-pptx
نفس الألوان الـ 8 بتخطيطات Classic/Bold/Minimal

### Premium — PptxGenJS (Node.js)
| النمط | الطابع |
|-------|--------|
| Noir Académique | أكاديمي فاخر داكن |
| Atlas Corporate | استشاري McKinsey |
| Sakura Créative | إبداعي طوكيو |

---

## 📊 الشرائح (14 شريحة قابلة للتحكم)

1. الغلاف (دائماً)
2. المقدمة + المقاربة
3. خطة الدراسة (الفصول والمباحث)
4. الإشكالية + التساؤل الرئيسي + الفرعية
5. الأهداف والفرضيات
6. أهمية الدراسة
7. المنهجية + العينة + المجالات
8. لوحة KPI
9. النتائج
10. الخاتمة
11. التوصيات
12. الآفاق البحثية
13. المراجع
14. شريحة الشكر (ثلاث لغات)

---

## 💻 التشغيل محلياً

```bash
pip install -r requirements.txt
cd node_scripts && npm install && cd ..
python app.py
```

ثم: http://localhost:5000

---

*مذكرتي Pro v7 — Canva Level · 2024–2025*

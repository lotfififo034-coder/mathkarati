# مذكرتي Pro v3 🎓

منشئ عروض أكاديمية احترافية للجامعات الجزائرية  
**محرك التصميم:** MathKarati PRO v3 — PptxGenJS

---

## 🏗️ البنية التقنية

```
mathkarati-final/
├── app.py                    ← Flask server (Python)
├── requirements.txt          ← Python dependencies
├── Procfile                  ← gunicorn start command
├── build.sh                  ← build script for Render
├── render.yaml               ← Render config
├── public/
│   └── index.html            ← الواجهة الكاملة (6 خطوات)
└── node_scripts/
    ├── package.json          ← pptxgenjs dependency
    └── generator_api.js      ← محرك PPTX (MathKarati PRO v3)
```

---

## 🚀 النشر على Render

### الطريقة الموصى بها

1. ارفع المشروع على GitHub
2. أنشئ **Web Service** جديد على [render.com](https://render.com)
3. اختر **Python** كـ runtime
4. في **Build Command**:
   ```bash
   pip install -r requirements.txt && cd node_scripts && npm install --production
   ```
5. في **Start Command**:
   ```bash
   gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
   ```

> **مهم:** Render يثبّت Node.js تلقائياً مع Python services — لا تحتاج لإضافته يدوياً.

---

## 💻 التشغيل المحلي

```bash
# 1. تثبيت Python dependencies
pip install -r requirements.txt

# 2. تثبيت Node.js dependencies
cd node_scripts
npm install
cd ..

# 3. تشغيل الخادم
python app.py
```

ثم افتح: http://localhost:5000

---

## 🎨 أنماط التصميم

| النمط | الروح | الألوان | الخط |
|-------|-------|---------|------|
| **Noir Académique** | أكاديمي فاخر — باريسي | أسود · ذهبي · فضي | Palatino + Cairo |
| **Atlas Corporate** | استشاري — McKinsey style | أزرق عميق · سيان · برتقالي | Trebuchet + Cairo |
| **Sakura Créative** | إبداعي — توكيو ستوديو | بنفسجي · مرجاني · بنفسجي فاتح | Georgia + Cairo |

---

## 📊 الشرائح المُنتَجة (حتى 12 شريحة)

1. الغلاف السينمائي
2. جدول المحتويات
3. الإشكالية والتساؤلات
4. الأهداف والفرضيات
5. أهمية البحث *(إذا أُدخلت)*
6. الإطار النظري *(إذا أُدخل)*
7. المنهجية والأدوات
8. لوحة المؤشرات KPI *(إذا أُدخلت)*
9. نتائج البحث *(إذا أُدخلت)*
10. مراجعة الدراسات السابقة *(إذا أُدخلت)*
11. التوصيات *(إذا أُدخلت)*
12. الخاتمة وشكر

---

*مذكرتي Pro v3 — 2024–2025*

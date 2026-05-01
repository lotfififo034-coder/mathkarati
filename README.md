# مذكرتي Pro 🎓

**منشئ العروض الأكاديمية الاحترافية للجامعات الجزائرية**

> يحوّل بيانات مذكرتك إلى عرض PowerPoint احترافي — نفس المحتوى، تجربة بصرية عالمية

---

## 🚀 النشر على Render (الطريقة الموصى بها)

### الخطوة 1 — رفع المشروع على GitHub

```bash
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/USERNAME/mathkarati-pro.git
git push -u origin main
```

### الخطوة 2 — إنشاء Web Service على Render

1. اذهب إلى [render.com](https://render.com) → **New** → **Web Service**
2. اربط حسابك بـ GitHub واختر الـ Repository
3. اضبط الإعدادات:

| الحقل | القيمة |
|---|---|
| **Name** | `mathkarati-pro` |
| **Runtime** | `Python 3` |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `gunicorn app:app --workers 2 --timeout 120 --bind 0.0.0.0:$PORT` |
| **Instance Type** | `Free` (أو Starter للإنتاج) |

4. اضغط **Create Web Service**
5. انتظر 2-3 دقائق حتى ينتهي البناء
6. الرابط سيكون: `https://mathkarati-pro.onrender.com`

---

## 💻 التشغيل المحلي (للتطوير)

```bash
# تثبيت المتطلبات
pip install -r requirements.txt

# تشغيل السيرفر
python app.py

# افتح المتصفح على
http://localhost:5000
```

---

## 📁 هيكل المشروع

```
mathkarati-pro/
├── app.py              ← Flask server (نقطة الدخول)
├── requirements.txt    ← Python dependencies
├── Procfile            ← Render/Heroku start command
├── render.yaml         ← Render configuration
├── .gitignore
├── public/
│   └── index.html      ← واجهة المستخدم (6 خطوات)
└── scripts/
    └── generator.py    ← محرك توليد PPTX
```

---

## 🎨 القوالب المتاحة

| الكود | الاسم | الألوان |
|---|---|---|
| `navy_gold` | كلاسيك أزرق ذهبي | أزرق داكن + ذهبي |
| `dark_teal` | نيل أخضر داكن | رمادي داكن + أخضر فيروزي |
| `burgundy` | بوردو فاخر | خمري + وردي |
| `forest` | غابة ملكي | أخضر داكن + زيتوني |

---

## ⚙️ المتطلبات

- Python 3.10+
- flask
- gunicorn
- python-pptx

---

## 📌 ملاحظات مهمة

- **Render Free Tier**: السيرفر "ينام" بعد 15 دقيقة من عدم الاستخدام — أول طلب بعد النوم يأخذ 30 ثانية
- **Render Starter ($7/شهر)**: لا ينام — موصى به للاستخدام الجاد
- **حجم الملفات**: كل PPTX بين 40-80 KB — لا مشكلة في الأداء
- **لا قاعدة بيانات**: كل شيء يتم في الذاكرة — لا حاجة لـ PostgreSQL

---

## 🔧 متغيرات البيئة (اختيارية)

| المتغير | القيمة الافتراضية | الوصف |
|---|---|---|
| `PORT` | `5000` | يُعيَّن تلقائياً من Render |
| `FLASK_ENV` | `production` | وضع الإنتاج |

---

صُنع بـ 🇩🇿 للطلاب الجزائريين

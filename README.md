# إتحاف الأخيار في الحفظ بالأعشار

منصة عربية RTL لاختبارات الحفظ الشفوي للقرآن الكريم، مع معاينة النماذج وتوليد Word وPDF محدثين من داخل الموقع.

## النسخة الحالية

- **46 مسارًا**.
- **1,263 اختبارًا**.
- **25 نموذجًا لكل اختبار**.
- الإجمالي: **31,575 نموذجًا**.
- البيانات المدمجة موجودة في `public/route-files/route-01.json` إلى `route-46.json`.

كل ملف مسار يحتوي اختبارات المسار كاملة، وكل اختبار يحتوي نماذجه الخمسة والعشرين. يعتمد الموقع على هذه الملفات مباشرة عبر `/route-files/`، وليس على ملفات `ranges` القديمة.

## أهم الملفات

- `src/main.jsx`: اختيار الاتجاه والفئة والمسار والاختبار، عرض النماذج، التعديل، Word وPDF والطباعة.
- `src/styles.css`: تصميم RTL والألوان والتنسيق الموحد.
- `public/route-catalog.json`: فهرس المسارات والاختبارات.
- `public/route-models-index.json`: فهرس الاختبارات وعدد النماذج.
- `public/route-files/`: 46 ملف JSON مدمجًا.
- `scripts/generate_route_models.py`: مولد النماذج والبيانات.
- `QA_REPORT.json`: تقرير فحص الأعداد والقواعد.

## التشغيل والبناء

```bash
npm install
npm run dev
npm run build
```

لإعادة توليد ملفات البيانات:

```bash
python3 scripts/generate_route_models.py
```

## النشر على GitHub وRender

ارفع `src` و`public/route-files` و`public/route-catalog.json` و`public/route-models-index.json` مع `package.json` وملف القفل وباقي ملفات المشروع. في Render استخدم:

- Build Command: `npm ci --silent && npm run build`
- Publish Directory: `dist`

بعد الرفع نفّذ **Clear build cache & deploy**.

## Supabase

ضع متغيرات البيئة في إعدادات Render فقط:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

لا تضع `service_role` في GitHub أو داخل الواجهة.

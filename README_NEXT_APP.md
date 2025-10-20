# مولّد خطة الدرس (Next.js 14 + TypeScript)

تطبيق ويب بسيط لإنشاء خطة درس وفق قالب الوزارة (DOCX) باستخدام Docxtemplater + PizZip.

## المتطلبات
- Node.js 18+ (مستخدَم هنا Node 22)

## التثبيت والتشغيل
```bash
# من جذر المشروع
npm install

# ضع قالب الوزارة:
#   templates/ministry_template.docx
# انظر أدناه لقائمة الحقول (Placeholders)

# تشغيل التطوير
npm run dev

# بناء الإنتاج
npm run build
npm start
```

## المسارات
- الواجهة: `/` — نموذج عربي RTL لجمع بيانات الدرس
- API: `/api/generate-docx` — يولّد ملف DOCX ويعيده للتحميل

## الحقول (Placeholders) في القالب
ضع الملف في `templates/ministry_template.docx`. يدعم القالب الحقول التالية:

- `{{teacher_name}}`
- `{{subject}}`
- `{{grade}}`
- `{{week}}`
- `{{lesson_title}}`
- `{{assessment}}`
- `{{homework}}`
- `{{today}}` (تاريخ اليوم، اختياري)

قوائم (صفيفات) — استخدم تكرار Docxtemplater:

أهداف الدرس:
```
{#objectives}
- { . }
{/objectives}
```

الوسائل/المواد:
```
{#materials}
- { . }
{/materials}
```

خطوات الدرس:
```
{#steps}
- { . }
{/steps}
```

ملاحظات:
- النقطة `{ . }` تطبع عنصر الصفيف. يمكن استخدام `{this}`.
- تأكد من أن القالب يستخدم خطوط تدعم العربية.
- يجب أن يكون اسم الملف: `ministry_template.docx`.

## الأمان
- لا توجد أسرار/مفاتيح API مطلوبة. أي متغيرات حساسة يجب أن تحفظ في الخادم فقط (بيئة Next.js Server) إن وجدت لاحقاً.

## البنية
- `app/page.tsx`: نموذج عربي RTL مع تحقّق أساسي وحالة تحميل
- `app/api/generate-docx/route.ts`: يقرأ القالب، يملأ الحقول، ويعيد DOCX عبر `Content-Disposition: attachment`
- `templates/`: ضع القالب هنا

## المدخلات من النموذج
- اسم المعلم: `teacher_name`
- المادة: `subject`
- الصف/المرحلة: `grade`
- الأسبوع: `week`
- عنوان الدرس: `lesson_title`
- الأهداف: `objectives[]` (مدخل نصي متعدد الأسطر يتم تحويله إلى مصفوفة)
- الوسائل/المواد: `materials[]`
- الخطوات: `steps[]`
- أساليب التقويم: `assessment`
- الواجب المنزلي: `homework`

## ملاحظات تشغيلية
- لا يتم حفظ أي بيانات على الخادم أو قاعدة بيانات.
- الواجهة ترسل JSON إلى `/api/generate-docx` وتستلم ملفًا قابلًا للتحميل.

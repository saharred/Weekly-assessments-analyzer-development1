"use client";

import { useState } from "react";

type LessonPlanForm = {
  teacher_name: string;
  subject: string;
  grade: string;
  week: string;
  lesson_title: string;
  objectives: string;
  materials: string;
  steps: string;
  assessment: string;
  homework: string;
};

function splitLinesToArray(value: string): string[] {
  return value
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter(Boolean);
}

export default function HomePage() {
  const [form, setForm] = useState<LessonPlanForm>({
    teacher_name: "",
    subject: "",
    grade: "",
    week: "",
    lesson_title: "",
    objectives: "",
    materials: "",
    steps: "",
    assessment: "",
    homework: "",
  });
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(false);

  function validate(): boolean {
    const e: Record<string, string> = {};
    if (!form.teacher_name.trim()) e.teacher_name = "مطلوب";
    if (!form.subject.trim()) e.subject = "مطلوب";
    if (!form.grade.trim()) e.grade = "مطلوب";
    if (!form.week.trim()) e.week = "مطلوب";
    if (!form.lesson_title.trim()) e.lesson_title = "مطلوب";
    if (!form.objectives.trim()) e.objectives = "مطلوب";
    if (!form.steps.trim()) e.steps = "مطلوب";
    setErrors(e);
    return Object.keys(e).length === 0;
  }

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!validate()) return;

    setLoading(true);
    try {
      const payload = {
        teacher_name: form.teacher_name.trim(),
        subject: form.subject.trim(),
        grade: form.grade.trim(),
        week: form.week.trim(),
        lesson_title: form.lesson_title.trim(),
        objectives: splitLinesToArray(form.objectives),
        materials: splitLinesToArray(form.materials),
        steps: splitLinesToArray(form.steps),
        assessment: form.assessment.trim(),
        homework: form.homework.trim(),
      };

      const res = await fetch("/api/generate-docx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(text || "فشل إنشاء الملف");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `خطة_${payload.subject}_${payload.lesson_title}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err: any) {
      alert(err?.message || "حدث خطأ غير متوقع");
    } finally {
      setLoading(false);
    }
  }

  function bind<K extends keyof LessonPlanForm>(key: K) {
    return {
      value: form[key],
      onChange: (ev: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) =>
        setForm((f) => ({ ...f, [key]: ev.target.value })),
    };
  }

  return (
    <div className="container">
      <div className="card">
        <h1>مولّد خطة الدرس (DOCX)</h1>
        <p className="muted">املأ الحقول التالية، وسيتم إنشاء ملف DOCX وفق قالب الوزارة.</p>

        <form onSubmit={onSubmit} className="form-grid">
          <div>
            <label>اسم المعلم</label>
            <input placeholder="مثال: أ. أحمد" {...bind("teacher_name")} />
            {errors.teacher_name && <div className="error">{errors.teacher_name}</div>}
          </div>

          <div>
            <label>المادة</label>
            <input placeholder="مثال: الرياضيات" {...bind("subject")} />
            {errors.subject && <div className="error">{errors.subject}</div>}
          </div>

          <div>
            <label>الصف/المرحلة</label>
            <input placeholder="مثال: الصف السادس" {...bind("grade")} />
            {errors.grade && <div className="error">{errors.grade}</div>}
          </div>

          <div>
            <label>الأسبوع</label>
            <input placeholder="مثال: الأسبوع الثالث" {...bind("week")} />
            {errors.week && <div className="error">{errors.week}</div>}
          </div>

          <div className="full">
            <label>عنوان الدرس</label>
            <input placeholder="مثال: الكسور العشرية" {...bind("lesson_title")} />
            {errors.lesson_title && <div className="error">{errors.lesson_title}</div>}
          </div>

          <div className="full">
            <label>الأهداف (سطر لكل هدف)</label>
            <textarea rows={4} placeholder={"مثال:\n- يعرّف مفهوم الكسور العشرية\n- يحوّل بين الكسر العادي والعشري"} {...bind("objectives")} />
            {errors.objectives && <div className="error">{errors.objectives}</div>}
          </div>

          <div className="full">
            <label>المواد/الوسائل (سطر لكل عنصر)</label>
            <textarea rows={3} placeholder={"مثال:\n- كتاب الطالب\n- سبورة ذكية"} {...bind("materials")} />
          </div>

          <div className="full">
            <label>خطوات الدرس (سطر لكل خطوة)</label>
            <textarea rows={5} placeholder={"مثال:\n- تمهيد ونشاط افتتاحي\n- شرح المفهوم\n- تدريب جماعي\n- تطبيق فردي"} {...bind("steps")} />
            {errors.steps && <div className="error">{errors.steps}</div>}
          </div>

          <div className="full">
            <label>أساليب التقويم</label>
            <textarea rows={3} placeholder="مثال: أسئلة شفهية، ورقة عمل" {...bind("assessment")} />
          </div>

          <div className="full">
            <label>الواجب المنزلي</label>
            <textarea rows={3} placeholder="مثال: حل التمارين 3-5 صفحة 21" {...bind("homework")} />
          </div>

          <div className="full actions">
            <button type="submit" className="primary" disabled={loading}>
              {loading ? "جاري الإنشاء..." : "إنشاء ملف DOCX"}
            </button>
            <span className="small">لن يتم حفظ أي بيانات على الخادم.</span>
          </div>
        </form>
      </div>
    </div>
  );
}

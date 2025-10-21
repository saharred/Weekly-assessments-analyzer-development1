# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io
from datetime import datetime

# ============== 🎨 تنسيقات CSS ==============
st.markdown("""
<style>
/* زر Browse Files بالعنابي */
div[data-testid="stFileUploader"] button {
    background-color: #8A1538 !important;
    color: white !important;
    border: 2px solid #C9A646 !important;
    font-weight: 700 !important;
}
div[data-testid="stFileUploader"] button:hover {
    background-color: #6B1029 !important;
}

/* نص الضمان بالذهبي الغامق */
.guarantee {
    color: #B8860B;
    font-weight: 800;
    text-shadow: 0 0 10px rgba(184,134,11,0.3);
    letter-spacing: 0.5px;
    font-size: 14px;
}

/* ترويسة أقسام */
.section-title{
    font-weight:800;font-size:20px;margin:6px 0 10px;color:#8A1538
}
.subtle{
    color:#555;font-size:13px
}
</style>
""", unsafe_allow_html=True)

st.title("إنجاز — إدارة القوائم الحرجة وإرسالها للمعلم")

# ============== ⚙️ دوال مساعدة ==============
def detect_percent_column(df: pd.DataFrame) -> str | None:
    """يحاول اكتشاف اسم عمود النسبة (الإنجاز الكلي) داخل الملف."""
    candidates = [
        "solve_pct", "النسبة", "النسبة الكلية", "متوسط الإنجاز", "المتوسط",
        "percent", "Percentage", "إنجاز", "إنجاز %"
    ]
    lower_cols = {c.lower(): c for c in df.columns}
    for name in candidates:
        if name.lower() in lower_cols:
            return lower_cols[name.lower()]
    return None

def categorize(percent: float) -> str:
    """تصنيف الأداء."""
    if pd.isna(percent) or percent == 0:
        return "لا يستفيد"
    if percent >= 90:
        return "بلاتيني 🥇"
    if percent >= 80:
        return "ذهبي 🥈"
    if percent >= 70:
        return "فضي 🥉"
    if percent >= 60:
        return "برونزي"
    return "بحاجة لتحسين"

def risk_score_from_row(row: pd.Series) -> int:
    """نسخة خفيفة لحساب الخطر (اختياري)."""
    base = 100 - float(row.get("المتوسط", row.get("النسبة_الكلية", 0)) or 0)
    pending_cols = [c for c in row.index if "متبقي" in str(c)]
    def _to_int(x):
        try:
            return int(str(x).split()[0])
        except:
            return 0
    pending_total = sum(_to_int(row[c]) for c in pending_cols if pd.notna(row[c]))
    score = base + 5 * pending_total
    return max(0, min(100, int(score)))

def build_email_body_html(subject_name: str, coordinator_name: str, reco_text: str) -> str:
    """يبني نص رسالة HTML محترف."""
    reco_html = (reco_text or "").replace("\n", "<br>")
    return f"""
    <div dir="rtl" style="font-family:'Tahoma',Arial;line-height:1.8">
      <p>السلام عليكم ورحمة الله وبركاته،</p>
      <p>نُرفق لكم قائمة الطلاب المطلوب متابعتهم لمادة <b>{subject_name or '—'}</b>.</p>
      <p><b>توصية منسّق المشاريع:</b><br>{reco_html or '—'}</p>
      <p>مع خالص الشكر والتقدير،<br>
         منسّق/ة المشاريع: <b>{coordinator_name or '—'}</b></p>
    </div>
    """

def to_csv_bytes(df: pd.DataFrame, fname_prefix: str) -> tuple[bytes, str]:
    """تحويل DataFrame إلى CSV بايت + اسم ملف."""
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    fname = f"{fname_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return csv_bytes, fname

def try_send_email(
    smtp_enabled: bool,
    smtp_server: str, smtp_port: int,
    smtp_user: str, smtp_pass: str,
    mail_from: str, mail_to: str,
    mail_subject: str, mail_html: str,
    attach_name: str, attach_bytes: bytes
):
    """
    إرسال فعلي عبر SMTP (اختياري).
    إذا smtp_enabled=False → معاينة فقط بدون إرسال.
    """
    if not smtp_enabled:
        st.info("📧 المعاينة فقط: لم يتم إرسال بريد (أزل علامة المعاينة لتفعيل الإرسال الفعلي).")
        return False, "Preview only"

    # إرسال فعلي (يلزم بيئة تدعم الشبكة)
    try:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.application import MIMEApplication

        msg = MIMEMultipart()
        msg["From"] = mail_from
        msg["To"] = mail_to
        msg["Subject"] = mail_subject

        msg.attach(MIMEText(mail_html, "html", "utf-8"))

        if attach_bytes:
            part = MIMEApplication(attach_bytes, Name=attach_name)
            part['Content-Disposition'] = f'attachment; filename="{attach_name}"'
            msg.attach(part)

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            if smtp_user and smtp_pass:
                server.login(smtp_user, smtp_pass)
            server.send_message(msg)

        return True, "Sent"
    except Exception as e:
        return False, f"Failed: {e}"

# ============== 📂 رفع الملف ==============
uploaded = st.file_uploader("📂 اختر ملف Excel للطلاب", type=["xlsx", "xls", "csv"])

if uploaded:
    # قراءة الملف
    try:
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"❌ خطأ في قراءة الملف: {e}")
        st.stop()

    st.success("✅ تم رفع الملف بنجاح")

    # محاولة اكتشاف عمود النسبة
    percent_col = detect_percent_column(df)
    if percent_col is None:
        st.warning("⚠️ لم أجد عمود النسبة تلقائيًا. اختاري/اختر العمود يدويًا من القائمة أدناه.")
        chosen = st.selectbox("🧮 اختر عمود النسبة (٪):", options=df.columns.tolist())
        percent_col = chosen

    # توحيد عمود النسبة تحت اسم موحد
    df = df.copy()
    df["النسبة_الكلية"] = pd.to_numeric(df[percent_col], errors="coerce").fillna(0).astype(float)

    # إضافة عمود الفئة إذا غير موجود
    if "الفئة" not in df.columns:
        df["الفئة"] = df["النسبة_الكلية"].apply(categorize)

    # أسماء مفاتيح محتملة للأسماء/الصف/الشعبة (للتصدير الجميل)
    name_col = None
    for c in ["الطالب", "اسم الطالب", "student", "Student", "student_name"]:
        if c in df.columns:
            name_col = c
            break

    grade_col = next((c for c in ["الصف", "Grade", "level", "class"] if c in df.columns), None)
    section_col = next((c for c in ["الشعبة", "Section", "section"] if c in df.columns), None)
    subject_col = next((c for c in ["المادة", "subject", "Subject"] if c in df.columns), None)

    # ============== 🧮 إنشاء القوائم ==============
    df_zero = df[df["النسبة_الكلية"] == 0].copy()
    df_need = df[df["الفئة"].astype(str).str.contains("بحاجة لتحسين", na=False)].copy()

    st.markdown("<div class='section-title'>📌 القوائم الحرِجة الجاهزة</div>", unsafe_allow_html=True)
    tabs = st.tabs(["📍 نسبة الإنجاز = 0%", "🛠️ فئة «بحاجة لتحسين»", "📦 تصدير/إرسال"])
    with tabs[0]:
        st.write("هذه قائمة الطلاب الذين لم ينجزوا أي تقييم (0%).")
        st.dataframe(df_zero, use_container_width=True, height=320)
        csv0, name0 = to_csv_bytes(
            df_zero[[c for c in [name_col, grade_col, section_col, subject_col, "النسبة_الكلية", "الفئة"] if c in df_zero.columns]],
            "طلاب_0_انجاز"
        )
        st.download_button("⬇️ تحميل CSV - 0%", csv0, file_name=name0, mime="text/csv")

    with tabs[1]:
        st.write("هذه قائمة الطلاب المصنّفين «بحاجة لتحسين».")
        st.dataframe(df_need, use_container_width=True, height=320)
        csvN, nameN = to_csv_bytes(
            df_need[[c for c in [name_col, grade_col, section_col, subject_col, "النسبة_الكلية", "الفئة"] if c in df_need.columns]],
            "طلاب_بحاجة_لتحسين"
        )
        st.download_button("⬇️ تحميل CSV - «بحاجة لتحسين»", csvN, file_name=nameN, mime="text/csv")

    with tabs[2]:
        st.write("اختاري/اختر أي قائمة لإرسالها عبر البريد للمعلم.")
        group_choice = st.radio(
            "المجموعة المراد إرسالها:",
            ["طلاب بنسبة 0%", "طلاب بحاجة لتحسين"],
            horizontal=True
        )
        target_df = df_zero if group_choice == "طلاب بنسبة 0%" else df_need
        if target_df.empty:
            st.warning("لا توجد بيانات في هذه المجموعة حاليًا.")
        else:
            st.success(f"عدد الطلاب في القائمة المختارة: {len(target_df)}")

        # توصية المنسق + معلومات المعلم
        st.markdown("<div class='section-title'>✍️ توصية منسّق المشاريع</div>", unsafe_allow_html=True)
        coordinator_name = st.text_input("اسم منسّق/ة المشاريع (يظهر في البريد):", value="")
        subject_name = st.text_input("اسم المادة (اختياري يظهر في البريد):", value="" if subject_col is None else str(df[subject_col].dropna().unique()[0]) if df[subject_col].notna().any() else "")
        reco_text = st.text_area("اكتب/ي التوصية هنا لتُضمَّن في البريد:", height=160, placeholder="مثال: يُرجى تعزيز المتابعة للطلاب المذكورين وتذكيرهم بحل التقييمات...")

        st.markdown("<div class='section-title'>📧 إعدادات الإرسال</div>", unsafe_allow_html=True)
        colA, colB = st.columns(2)
        with colA:
            teacher_email = st.text_input("بريد المعلم/ة:", value="")
            mail_subject = st.text_input("عنوان الرسالة:", value=f"قائمة المتابعة — {group_choice}")
        with colB:
            smtp_enabled = st.checkbox("أرسل فعليًا عبر SMTP (اختياري)", value=False, help="إذا لم تُفعِّل هذا الخيار سيتم إنشاء معاينة فقط بدون إرسال.")
            smtp_server = st.text_input("SMTP Server", value="smtp.education.qa")
            smtp_port = st.number_input("SMTP Port", value=587, step=1)
            smtp_user = st.text_input("SMTP User (اختياري)", value="")
            smtp_pass = st.text_input("SMTP Password (اختياري)", value="", type="password")
            mail_from = st.text_input("المرسل (From):", value=smtp_user or "noreply@education.qa")

        # معاينة النص
        mail_html = build_email_body_html(subject_name=subject_name, coordinator_name=coordinator_name, reco_text=reco_text)
        st.markdown("<div class='section-title'>👀 معاينة المحتوى</div>", unsafe_allow_html=True)
        st.components.v1.html(mail_html, height=220, scrolling=True)

        # تحضير المرفق
        attach_df = target_df[[c for c in [name_col, grade_col, section_col, subject_col, "النسبة_الكلية", "الفئة"] if c in target_df.columns]].copy()
        attach_csv, attach_name = to_csv_bytes(attach_df, "قائمة_متابعة")

        colbtn1, colbtn2 = st.columns(2)
        with colbtn1:
            st.download_button("⬇️ تحميل المرفق (CSV)", attach_csv, file_name=attach_name, mime="text/csv")
        with colbtn2:
            disabled = not teacher_email or attach_csv is None
            if st.button("📧 إرسال الآن", type="primary", disabled=disabled):
                ok, msg = try_send_email(
                    smtp_enabled=smtp_enabled,
                    smtp_server=smtp_server, smtp_port=int(smtp_port),
                    smtp_user=smtp_user, smtp_pass=smtp_pass,
                    mail_from=mail_from, mail_to=teacher_email,
                    mail_subject=mail_subject, mail_html=mail_html,
                    attach_name=attach_name, attach_bytes=attach_csv
                )
                if ok:
                    st.success("✅ تم الإرسال بنجاح." if smtp_enabled else "✅ تمت المعاينة. (لم يتم الإرسال الفعلي)")
                else:
                    st.error(f"❌ تعذّر الإرسال: {msg}")

# ============== 📝 Footer ==============
st.markdown("""
<hr style="border: 1px solid #C9A646;">
<div style="text-align:center;">
  <p><b>إنجاز — نظام ذكي لتحليل ومتابعة التقييمات الأسبوعية على نظام قطر للتعليم</b></p>
  <p class="guarantee">✨ ضمان تنمية رقمية مستدامة ✨</p>
  <p>© 2025 جميع الحقوق محفوظة — مدرسة عثمان بن عفّان النموذجية للبنين<br>
     All rights reserved © 2025 — Othman Bin Affan Model School for Boys</p>
  <p>تطوير وتنفيذ: منسّق المشاريع الإلكترونية / سحر عثمان<br>
     Developed & Implemented by: E-Projects Coordinator / Sahar Osman</p>
  <p>📧 للتواصل | Contact:
     <a href="mailto:Sahar.Osman@education.qa" style="color:#E8D4A0;">Sahar.Osman@education.qa</a></p>
  <p>🎯 رؤيتنا: متعلم ريادي لتنمية مستدامة</p>
</div>
<hr style="border: 1px solid #C9A646;">
""", unsafe_allow_html=True)

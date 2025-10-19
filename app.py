import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
import re
import unicodedata
from datetime import datetime

# إعداد الصفحة
st.set_page_config(
    page_title="📊 محلل التقييمات الأسبوعية",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================== HELPER FUNCTIONS ==================

def normalize_cell_token(cell) -> str:
    """يطبع رمز الخلية لغايات العد"""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    s = str(cell).strip()

    # إزالة محارف الاتجاه/الخفيّة
    s = re.sub(r'[\u200e\u200f\u202a-\u202e\u2066-\u2069\u200b\u200c\u200d]', '', s)

    # توحيد حرف الميم العربي 'م' إلى 'M'
    s_nfkc = unicodedata.normalize("NFKC", s)
    s_nfkc = s_nfkc.replace("م", "M").replace("ﻡ", "M")

    # تحويل للحروف الكبيرة
    s_nfkc = s_nfkc.strip().upper()

    # اعتبر "-" أو "—" أو فارغ = غير مسند
    if s_nfkc in {"", "-", "—"}:
        return ""

    return s_nfkc


def analyze_excel_file(file, sheet_name):
    """تحليل شيت واحد"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)

        subject = str(df.iloc[0, 0]) if pd.notna(df.iloc[0, 0]) else sheet_name
        results = []

        # أول صفوف لعنوان المادة/الصفوف نطنشها
        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue

            student_name_clean = " ".join(str(student_name).strip().split())

            total_assessments = df.shape[1] - 7  # عدد الأعمدة بعد العمود G
            assigned_count = 0
            m_count = 0

            for i in range(total_assessments):
                col_idx = 7 + i
                if col_idx >= df.shape[1]:
                    break

                raw_val = df.iloc[idx, col_idx]
                token = normalize_cell_token(raw_val)

                # استبعاد "-" أو فارغ
                if token == "":
                    continue

                # أي قيمة غير فارغة = تقييم مسند
                assigned_count += 1

                # لو "M" → لم يسلم
                if token == "M":
                    m_count += 1

            completed_count = assigned_count - m_count
            solve_pct = (completed_count / assigned_count * 100) if assigned_count > 0 else 0.0

            results.append({
                "student_name": student_name_clean,
                "subject": subject,
                "completed_count": completed_count,
                "assigned_count": assigned_count,
                "solve_pct": solve_pct
            })

        return results

    except Exception as e:
        st.error(f"خطأ في تحليل الورقة {sheet_name}: {str(e)}")
        return []


# ================== MAIN APP ==================

st.title("📊 محلل التقييمات الأسبوعية")
st.markdown("---")

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

with st.sidebar:
    st.header("⚙️ الإعدادات")

    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")

        try:
            xls = pd.ExcelFile(uploaded_files[0])
            sheets = xls.sheet_names
            selected_sheets = st.multiselect("اختر الأوراق", sheets, default=sheets)
        except Exception as e:
            st.error(f"خطأ: {e}")
            selected_sheets = []
    else:
        selected_sheets = []

    run_analysis = st.button(
        "🚀 تشغيل التحليل",
        use_container_width=True,
        type="primary",
        disabled=not (uploaded_files and selected_sheets)
    )

if not uploaded_files:
    st.info("👈 الرجاء رفع ملفات Excel من الشريط الجانبي")

elif run_analysis:
    with st.spinner("⏳ جاري التحليل..."):
        try:
            all_results = []
            for file in uploaded_files:
                for sheet in selected_sheets:
                    results = analyze_excel_file(file, sheet)
                    all_results.extend(results)

            if all_results:
                df = pd.DataFrame(all_results)
                st.session_state.analysis_results = df
                st.success(f"✅ تم تحليل {len(df)} سجل طالب!")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")
        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")

# عرض النتائج
if st.session_state.analysis_results is not None:
    df = st.session_state.analysis_results

    st.subheader("📊 النتائج")
    display_df = df.copy()
    display_df['solve_pct'] = display_df['solve_pct'].apply(lambda x: f"{x:.2f}%")
    st.dataframe(display_df, use_container_width=True, height=400)

    # تحميل CSV
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        "📥 تحميل CSV",
        csv,
        f"analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        "text/csv"
    )

    st.divider()

    # رسم بياني للفئات
    st.subheader("📈 توزيع الإنجاز")

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.hist(df['solve_pct'], bins=10, color="#800000", edgecolor="black")
    ax.set_xlabel("نسبة الإنجاز %")
    ax.set_ylabel("عدد الطلاب")
    ax.set_title("توزيع نسب الإنجاز")
    st.pyplot(fig)

# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile, io, re, unicodedata
from datetime import datetime

# ============ إعداد الصفحة ============
st.set_page_config(page_title="محلل التقييمات الأسبوعية", page_icon="📊", layout="wide")

# ============ تنسيق بسيط ============
st.markdown("""
<style>
html, body, [class*="css"] { font-family: "Tajawal","Cairo","DejaVu Sans",Arial,sans-serif !important; }
thead tr th { background:#8A1538 !important; color:#fff !important; font-weight:700 !important; }
.stButton>button { background:#8A1538; color:#fff; border:1px solid #6b0f2b; border-radius:10px; }
.stButton>button:hover { background:#6b0f2b; }
</style>
""", unsafe_allow_html=True)

import matplotlib
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['axes.unicode_minus'] = False

# ============ أدوات مساعدة ============

def normalize_ar_name(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r'[\u200e\u200f\u202a-\u202e\u2066-\u2069]', '', s)  # محارف اتجاه
    s = s.replace("ـ", "")
    s = re.sub(r"[إأآٱ]", "ا", s)
    s = s.replace("ى", "ي")
    s = re.sub(r"[ًٌٍَُِّْ]", "", s)  # تشكيل
    s = " ".join(s.split())
    return s.strip()

def normalize_cell_token(cell) -> str:
    """
    تطبيع قيمة الخلية لعد التقييمات:
    - إزالة محارف خفيّة
    - توحيد 'م' العربية إلى 'M'
    - تحويل إلى حروف كبيرة
    - اعتبار '-' و '—' و الفراغ = غير مُسنّد
    - إرجاع نص غير فارغ للأرقام والدرجات
    """
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    s = str(cell).strip()
    s = re.sub(r'[\u200e\u200f\u202a-\u202e\u2066-\u2069\u200b\u200c\u200d]', '', s)  # محارف خفية
    s = unicodedata.normalize("NFKC", s).replace("م", "M").replace("ﻡ", "M").upper().strip()
    if s in {"", "-", "—"}:
        return ""
    return s

def analyze_excel_sheet(file, sheet_name):
    """
    يحلل ورقة واحدة:
    - المادة = اسم الورقة (ثابت)
    - التقييمات تبدأ من العمود H (index 7)، ونعدّ الأعمدة التي لها عنوان غير فارغ في الصف 0
    - المسند = خلايا غير فارغة بعد التطبيع، ما عدا '-'/'—'
    - M فقط = لم يُسلّم
    - المنجز = المسند - M
    """
    df = pd.read_excel(file, sheet_name=sheet_name, header=None)

    # حدد عدد أعمدة التقييم الفعلية من صف العناوين (الصف 0) بعد H
    total_cols = 0
    for col in range(7, df.shape[1]):
        title = df.iloc[0, col]
        if pd.notna(title) and str(title).strip():
            total_cols += 1

    results, seen = [], set()
    subject = sheet_name  # << المهم: المادة = اسم الورقة

    for r in range(4, len(df)):  # أسماء الطلاب من الصف 5 (index 4)
        raw_name = df.iloc[r, 0]
        if pd.isna(raw_name) or str(raw_name).strip() == "":
            continue
        student = normalize_ar_name(raw_name)
        if not student or (student, subject) in seen:
            continue
        seen.add((student, subject))

        assigned = 0
        m_count = 0

        for i in range(total_cols):
            c = 7 + i
            if c >= df.shape[1]: break
            tok = normalize_cell_token(df.iloc[r, c])

            if tok == "":        # فارغ / '-' / '—' => غير مُسنّد
                continue

            # أي قيمة غير فارغة = مُسنّد
            assigned += 1

            # فقط M = لم يُسلّم
            if tok == "M":
                m_count += 1

        completed = max(assigned - m_count, 0)
        pct = (completed / assigned * 100) if assigned > 0 else 0.0

        results.append({
            "student_name": student,
            "subject": subject,
            "assigned_count": assigned,
            "completed_count": completed,
            "solve_pct": pct
        })
    return results

def create_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """
    جدول محوري: صف واحد لكل طالب، والمواد كرؤوس أعمدة:
    [المادة - إجمالي]، [المادة - منجز]، [المادة - نسبة %] + نسبة عامة وفئة.
    """
    df = df.copy()
    df["student_name"] = df["student_name"].apply(normalize_ar_name)

    meta = df[["student_name"]].drop_duplicates().copy()
    result = meta.copy()

    for sub in sorted(df["subject"].unique()):
        part = (df[df["subject"] == sub]
                .drop_duplicates(subset=["student_name"], keep="first")
                [["student_name", "assigned_count", "completed_count", "solve_pct"]]
                .rename(columns={
                    "assigned_count": f"{sub} - إجمالي",
                    "completed_count": f"{sub} - منجز",
                    "solve_pct": f"{sub} - نسبة %"
                }))
        result = result.merge(part, on="student_name", how="left")

    pct_cols = [c for c in result.columns if c.endswith("نسبة %")]
    if pct_cols:
        result["النسبة الكلية %"] = result[pct_cols].mean(axis=1, skipna=True)

        def cat(p):
            if pd.isna(p): return "-"
            if p == 0: return "لا يستفيد من النظام 🚫"
            if p >= 90: return "البلاتينية 🥇"
            if p >= 80: return "الذهبي 🥈"
            if p >= 70: return "الفضي 🥉"
            if p >= 60: return "البرونزي"
            return "يحتاج تحسين ⚠️"
        result["الفئة"] = result["النسبة الكلية %"].apply(cat)

    result = result.rename(columns={"student_name": "اسم الطالب"})
    return result

# ============ الواجهة الرئيسية ============

st.title("📊 محلل التقييمات الأسبوعية")
st.markdown("---")

if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "pivot" not in st.session_state:
    st.session_state.pivot = None

with st.sidebar:
    st.header("⚙️ الإعدادات")
    uploaded = st.file_uploader("اختر ملفات Excel", type=["xlsx", "xls"], accept_multiple_files=True)

    # تجميع كل أسماء الأوراق من جميع الملفات
    if uploaded:
        all_sheets = set()
        per_file_sheets = {}
        for f in uploaded:
            try:
                xls = pd.ExcelFile(f)
                per_file_sheets[f.name] = set(xls.sheet_names)
                all_sheets.update(xls.sheet_names)
            except Exception as e:
                st.warning(f"تعذّر قراءة الأوراق من {f.name}: {e}")
        all_sheets = sorted(list(all_sheets))
        selected_sheets = st.multiselect("اختر الأوراق (المواد)", all_sheets, default=all_sheets)
    else:
        selected_sheets, per_file_sheets = [], {}

    run = st.button("🚀 تشغيل التحليل", use_container_width=True, type="primary",
                    disabled=not (uploaded and selected_sheets))

if not uploaded:
    st.info("👈 ارفع ملف/ملفات Excel من الشريط الجانبي.")
elif run:
    with st.spinner("جاري التحليل…"):
        all_rows = []
        missing = []
        # نحلّل لكل ملف فقط الأوراق الموجودة فعليًا بداخله
        steps = sum(len(set(selected_sheets) & per_file_sheets.get(f.name, set())) for f in uploaded)
        done = 0
        prog = st.progress(0.0)

        for f in uploaded:
            available = per_file_sheets.get(f.name, set())
            target = list(set(selected_sheets) & available)
            missing_in_file = list(set(selected_sheets) - available)
            if missing_in_file:
                for m in sorted(missing_in_file):
                    missing.append(f"{f.name} → {m}")

            for sh in sorted(target):
                try:
                    rows = analyze_excel_sheet(f, sh)
                    all_rows.extend(rows)
                except Exception as e:
                    missing.append(f"{f.name} → {sh} (خطأ: {e})")
                finally:
                    done += 1
                    prog.progress(min(1.0, done / max(1, steps)))

        if missing:
            with st.expander("الأوراق التي لم تُحلَّل (غير موجودة في بعض الملفات أو أخطاء قراءة)"):
                for m in missing:
                    st.write("• ", m)

        if all_rows:
            raw_df = pd.DataFrame(all_rows)
            st.session_state.raw_df = raw_df
            pivot = create_pivot(raw_df)
            st.session_state.pivot = pivot
            st.success(f"✅ تم تحليل {len(raw_df)} سجل عبر {raw_df['subject'].nunique()} مادة و{len(uploaded)} ملف.")
        else:
            st.warning("لم يتم العثور على بيانات صالحة.")

# ============ عرض النتائج ============

if st.session_state.raw_df is not None:
    raw_df = st.session_state.raw_df
    st.subheader("📋 النتائج التفصيلية (صف لكل مادة)")
    show = raw_df.copy()
    show["solve_pct"] = show["solve_pct"].apply(lambda x: f"{x:.2f}%")
    st.dataframe(show, use_container_width=True, height=420)

    csv = raw_df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("📥 تحميل CSV التفصيلي", csv,
                       f"raw_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                       "text/csv")

if st.session_state.pivot is not None:
    pivot = st.session_state.pivot
    st.markdown("## 📊 الجدول المحوري (المواد كرؤوس)")
    nice = pivot.copy()
    for c in nice.columns:
        if c.endswith("نسبة %") or c == "النسبة الكلية %":
            nice[c] = nice[c].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else "-")
    st.dataframe(nice, use_container_width=True, height=520)

    # تنزيلات
    col1, col2 = st.columns(2)
    with col1:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            pivot.to_excel(w, index=False, sheet_name="Pivot")
        st.download_button("📊 تحميل Excel (Pivot)", out.getvalue(),
                           f"pivot_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with col2:
        csvp = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("📄 تحميل CSV (Pivot)", csvp,
                           f"pivot_{datetime.now().strftime('%Y%m%d')}.csv",
                           "text/csv", use_container_width=True)

    st.markdown("---")
    st.subheader("📈 متوسط الإنجاز حسب المادة")
    fig, ax = plt.subplots(figsize=(10,5))
    sub_avg = raw_df.groupby("subject")["solve_pct"].mean().sort_values()
    ax.barh(range(len(sub_avg)), sub_avg.values, edgecolor="black")
    ax.set_yticks(range(len(sub_avg)))
    ax.set_yticklabels(list(sub_avg.index))
    ax.set_xlabel("نسبة الإنجاز (%)")
    ax.set_xlim(0, 100)
    st.pyplot(fig)

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import io
from datetime import datetime, date
import unicodedata, re
from typing import Optional

# ============== إعدادات الصفحة ==============
st.set_page_config(
    page_title="إنجاز - تحليل القييمات الأسبوعية على نظام قطر للتعليم",
    page_icon="https://i.imgur.com/XLef7tS.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============== دوال مساعدة للتواريخ ==============

def _normalize_arabic_digits(s: str) -> str:
    """تحويل الأرقام العربية-الهندية إلى إنجليزية"""
    return s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

def _strip_invisible_and_diacritics(s: str) -> str:
    """إزالة المحارف غير المرئية والتشكيل والتطويل"""
    INVISIBLES = {
        "\u200f", "\u200e", "\u202a", "\u202b", "\u202c",
        "\u202d", "\u202e", "\u2066", "\u2067", "\u2069",
        "\u00a0", "\ufeff", "ـ"
    }
    for ch in INVISIBLES:
        s = s.replace(ch, " ")
    s = "".join(c for c in unicodedata.normalize("NFKD", s)
                if not unicodedata.combining(c))
    return " ".join(s.split()).strip()

def parse_due_date_cell(cell, default_year: int = None) -> Optional[date]:
    """
    يحوّل قيمة خلية H3 إلى datetime.date:
    - يدعم Timestamp/Datetime, رقم تسلسلي Excel, نص عربي أو إنجليزي.
    - يعيد None إذا فشل.
    """
    if default_year is None:
        default_year = date.today().year

    # Timestamp
    if isinstance(cell, (pd.Timestamp, )):
        return cell.date()
    if hasattr(cell, "date"):
        try: return cell.date()
        except: pass

    # رقم تسلسلي Excel
    try:
        if isinstance(cell, (int, float)) and not pd.isna(cell):
            base = pd.to_datetime("1899-12-30")
            d = base + pd.to_timedelta(float(cell), unit="D")
            if 2000 <= d.year <= 2100:
                return d.date()
    except: pass

    # نص عربي/إنجليزي
    try:
        s = str(cell or "").strip()
        if not s: return None
        s = _normalize_arabic_digits(_strip_invisible_and_diacritics(s))

        ar_months = {
            "يناير":1,"فبراير":2,"مارس":3,
            "ابريل":4,"أبريل":4,"نيسان":4,
            "مايو":5,"يونيو":6,"يونيه":6,
            "يوليو":7,"يوليه":7,
            "اغسطس":8,"أغسطس":8,"آب":8,
            "سبتمبر":9,"ايلول":9,
            "اكتوبر":10,"أكتوبر":10,"تشرين الاول":10,"تشرين الأول":10,
            "نوفمبر":11,"تشرين الثاني":11,
            "ديسمبر":12,"كانون الاول":12,"كانون الأول":12
        }

        s_norm = (s.replace("أ","ا").replace("إ","ا").replace("آ","ا").replace("ـ",""))
        m = re.search(r"(\d{1,2})\s*[-/ ]*\s*([^\d\s]+)", s_norm)
        if m:
            day = int(m.group(1))
            mon_name = m.group(2).strip()
            month = ar_months.get(mon_name)
            if not month:
                mon_name = re.sub(r"[^ء-ي]+","",mon_name)
                month = ar_months.get(mon_name)
            if month:
                try:
                    return pd.Timestamp(year=default_year, month=month, day=day).date()
                except:
                    return date(default_year, month, min(day,28))

        # محاولة أخيرة لو مكتوب بالإنجليزي
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt): return dt.date()
    except: pass

    return None

def in_range(d: Optional[date], start: Optional[date], end: Optional[date]) -> bool:
    """يتحقق إذا التاريخ d داخل المدى (start–end)."""
    if d is None: return True   # إذا لم نفهم التاريخ نعتبره داخل المدى
    if start and d < start: return False
    if end and d > end: return False
    return True

# ============== دالة تحليل ملف الإكسل ==============

def analyze_excel_file(file, sheet_name, due_start=None, due_end=None):
    df = pd.read_excel(file, sheet_name=sheet_name, header=None)
    results = []

    filter_active = (due_start is not None and due_end is not None)

    # الأعمدة تبدأ من H (index 7)
    for c in range(7, df.shape[1]):
        title = df.iloc[0, c]
        if pd.isna(title): continue

        # تاريخ الاستحقاق في الصف الثالث (H3)
        due_cell = df.iloc[2, c] if c < df.shape[1] else None
        due_dt   = parse_due_date_cell(due_cell, default_year=date.today().year)

        # فلترة التاريخ: إذا تعرّفنا عليه وكان خارج المدى → تجاهل
        if filter_active and (due_dt is not None) and (not in_range(due_dt, due_start, due_end)):
            continue

        # مثال مبسّط: حساب المنجز/الإجمالي لكل طالب
        for r in range(4, df.shape[0]):
            student = df.iloc[r, 0]
            if pd.isna(student): continue

            val = str(df.iloc[r, c]).strip()
            if val in ["-", "M", "I", "AB", "X", ""]: 
                done = 0
            else:
                try:
                    num = float(val)
                    done = 1 if num > 0 else 0
                except:
                    done = 1

            results.append({
                "student": student,
                "assessment": str(title),
                "due_date": due_dt,
                "done": done
            })
    return pd.DataFrame(results)

# ============== واجهة التطبيق ==============

st.title("📊 إنجاز - تحليل القييمات الأسبوعية على نظام قطر للتعليم")

uploaded_files = st.file_uploader("📂 ارفع ملفات Excel", type=["xlsx","xls"], accept_multiple_files=True)
run_btn = st.button("▶️ تشغيل التحليل")

if run_btn and uploaded_files:
    all_results = []
    for f in uploaded_files:
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            df = analyze_excel_file(f, sheet)
            all_results.append(df)
    if all_results:
        df_all = pd.concat(all_results, ignore_index=True)
        st.success(f"✅ تم تحليل {len(df_all)} سجل")
        st.dataframe(df_all, use_container_width=True)

        # تحميل CSV
        csv = df_all.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("📥 تحميل CSV", csv, "results.csv", "text/csv")
    else:
        st.warning("⚠️ لم يتم العثور على بيانات")

# ============== الفوتر ==============

st.markdown("""
<div class="footer" style="margin-top:50px; text-align:center; padding:15px; background:linear-gradient(135deg,#8A1538,#6B1029); color:white; border-radius:10px;">
  <div class="school" style="font-weight:800; font-size:16px;">مدرسة عُثمان بن عفّان النموذجية للبنين</div>
  <div class="rights" style="font-weight:700; font-size:12px;">جميع الحقوق محفوظة © 2025</div>
  <div class="contact" style="font-size:12px;">للتواصل: <a href="mailto:S.mahgoub0101@education.qa" style="color:#E8D4A0; text-decoration:none;">S.mahgoub0101@education.qa</a></div>
  <div class="credit" style="font-size:11px;">إشراف وتنفيذ: منسّق المشاريع الإلكترونية / سحر عثمان</div>
</div>
""", unsafe_allow_html=True)

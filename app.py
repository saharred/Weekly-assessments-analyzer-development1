# -*- coding: utf-8 -*-
import os, io, re, zipfile, logging, unicodedata
from datetime import datetime, date
from typing import Tuple, Optional, List

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ========= PDF (fpdf2) + Arabic RTL =========
from fpdf import FPDF
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    AR_OK = True
except Exception:
    AR_OK = False

QATAR_MAROON = (138, 21, 56)
QATAR_GOLD   = (201, 166, 70)

# ============== دوال معالجة التاريخ المحسّنة ==============

def _normalize_arabic_digits(s: str) -> str:
    """
    تحويل الأرقام العربية-الهندية (٠-٩) إلى أرقام إنجليزية (0-9)
    """
    if not isinstance(s, str):
        return str(s) if s is not None else ""
    return s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

def _strip_invisible_and_diacritics(s: str) -> str:
    """
    إزالة المحارف غير المرئية والتشكيل والمحارف الاتجاهية
    - RTL/LTR marks
    - Zero-width characters
    - Arabic diacritics (تشكيل)
    - Tatweel (تطويل)
    """
    if not isinstance(s, str):
        return ""
    
    # إزالة المحارف غير المرئية والاتجاهية
    invisible_chars = [
        '\u200e',  # LRM (Left-to-Right Mark)
        '\u200f',  # RLM (Right-to-Left Mark)
        '\u202a',  # LRE
        '\u202b',  # RLE
        '\u202c',  # PDF
        '\u202d',  # LRO
        '\u202e',  # RLO
        '\u2066',  # LRI
        '\u2067',  # RLI
        '\u2068',  # FSI
        '\u2069',  # PDI
        '\u200b',  # ZWSP (Zero Width Space)
        '\u200c',  # ZWNJ
        '\u200d',  # ZWJ
        '\ufeff',  # ZWNBSP / BOM
        '\xa0',    # NBSP
        '\u0640',  # Tatweel (تطويل)
    ]
    
    for char in invisible_chars:
        s = s.replace(char, '')
    
    # إزالة التشكيل العربي باستخدام unicodedata
    s = ''.join(c for c in s if not unicodedata.combining(c))
    
    # تطبيع المسافات المتعددة
    s = ' '.join(s.split())
    
    return s.strip()

def parse_due_date_cell(cell, default_year: int = None) -> Optional[date]:
    """
    تحويل خلية تاريخ الاستحقاق (H3) إلى datetime.date
    
    المدخلات:
        cell: قد يكون Timestamp، datetime، رقم تسلسلي Excel، أو نص (عربي/إنجليزي)
        default_year: السنة الافتراضية عند غياب السنة في النص (افتراضي: سنة اليوم)
    
    المخرجات:
        datetime.date عند النجاح، None عند الفشل
    
    يدعم:
        - Timestamps/Datetime
        - أرقام Excel التسلسلية
        - نصوص عربية: "2 أكتوبر", "٢ اكتوبر", "2-أكتوبر"
        - نصوص إنجليزية: "2 Oct", "02 Oct 2025"
    """
    # السنة الافتراضية
    if default_year is None:
        default_year = date.today().year
    
    # 1) التحقق من القيم الفارغة
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    
    # 2) معالجة Timestamp/Datetime
    if isinstance(cell, (pd.Timestamp, datetime)):
        try:
            return cell.date() if hasattr(cell, 'date') else cell
        except Exception:
            pass
    
    # 3) معالجة الرقم التسلسلي Excel
    try:
        if isinstance(cell, (int, float)) and not pd.isna(cell):
            # التحقق من أن الرقم في نطاق معقول
            if 1 <= cell <= 100000:  # تقريباً 1900-2173
                base = pd.to_datetime("1899-12-30")
                result = base + pd.to_timedelta(float(cell), unit="D")
                # التحقق من أن السنة معقولة
                if 1900 <= result.year <= 2200:
                    return result.date()
    except Exception:
        pass
    
    # 4) معالجة النصوص (عربي/إنجليزي)
    try:
        s = str(cell).strip()
        if not s or s.lower() in ['nan', 'none', 'nat']:
            return None
        
        # تنظيف النص
        s = _strip_invisible_and_diacritics(s)
        s = _normalize_arabic_digits(s)
        
        if not s:
            return None
        
        # خريطة الأشهر العربية الشاملة
        arabic_months = {
            # يناير
            "يناير": 1, "كانون الثاني": 1, "جانفي": 1,
            # فبراير
            "فبراير": 2, "شباط": 2, "فيفري": 2,
            # مارس
            "مارس": 3, "اذار": 3, "آذار": 3,
            # أبريل
            "ابريل": 4, "أبريل": 4, "نيسان": 4, "افريل": 4,
            # مايو
            "مايو": 5, "ماي": 5, "ايار": 5, "أيار": 5,
            # يونيو
            "يونيو": 6, "يونيه": 6, "حزيران": 6, "جوان": 6,
            # يوليو
            "يوليو": 7, "يوليه": 7, "تموز": 7, "جويلية": 7,
            # أغسطس
            "اغسطس": 8, "أغسطس": 8, "اب": 8, "آب": 8, "اوت": 8,
            # سبتمبر
            "سبتمبر": 9, "ايلول": 9, "أيلول": 9, "سيبتمبر": 9,
            # أكتوبر
            "اكتوبر": 10, "أكتوبر": 10, "تشرين الاول": 10, "تشرين الأول": 10, "اكتوبر": 10,
            # نوفمبر
            "نوفمبر": 11, "تشرين الثاني": 11, "نونبر": 11,
            # ديسمبر
            "ديسمبر": 12, "كانون الاول": 12, "كانون الأول": 12, "دجنبر": 12,
        }
        
        # تطبيع الهمزات للبحث الأفضل
        def normalize_hamza(text):
            return text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ـ", "")
        
        # محاولة استخراج اليوم + الشهر بالعربية
        # Pattern: رقم + فاصل اختياري + نص الشهر
        pattern = r"(\d{1,2})\s*[-/\s]*\s*([^\d\s]+)"
        match = re.search(pattern, s)
        
        if match:
            try:
                day = int(match.group(1))
                month_name = match.group(2).strip()
                
                # البحث في خريطة الأشهر
                month = None
                
                # بحث مباشر
                if month_name in arabic_months:
                    month = arabic_months[month_name]
                else:
                    # بحث بعد التطبيع
                    normalized_name = normalize_hamza(month_name)
                    for key, val in arabic_months.items():
                        if normalize_hamza(key) == normalized_name:
                            month = val
                            break
                
                if month:
                    # محاولة إنشاء التاريخ
                    try:
                        # محاولة 1: استخدام التاريخ مباشرة
                        return date(default_year, month, day)
                    except ValueError:
                        # محاولة 2: تقليم اليوم إذا كان خارج النطاق
                        try:
                            safe_day = min(day, 28)  # أقل عدد أيام مضمون في أي شهر
                            return date(default_year, month, safe_day)
                        except ValueError:
                            pass
            except (ValueError, AttributeError):
                pass
        
        # 5) محاولة pandas.to_datetime للنصوص الإنجليزية أو التنسيقات القياسية
        try:
            parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(parsed):
                result_date = parsed.date()
                # إذا لم تكن هناك سنة في النص، استبدلها بالسنة الافتراضية
                if parsed.year < 1900:  # سنة غير معقولة
                    result_date = result_date.replace(year=default_year)
                return result_date
        except Exception:
            pass
    
    except Exception:
        pass
    
    # الفشل في كل المحاولات
    return None

# ============== اختبارات الدالة ==============
if __name__ == "__main__":
    print("=" * 60)
    print("اختبار دالة parse_due_date_cell")
    print("=" * 60)
    
    test_cases = [
        # حالات عربية
        ("2 أكتوبر", "نص عربي بسيط"),
        ("٢ اكتوبر", "أرقام عربية-هندية"),
        ("  2 أكتوبر  ", "مسافات إضافية"),
        ("15-أكتوبر", "فاصلة شرطة"),
        ("15/اكتوبر", "فاصلة شرطة مائلة"),
        ("2 تشرين الأول", "اسم شهر بديل"),
        
        # حالات إنجليزية
        ("2 Oct", "نص إنجليزي قصير"),
        ("02 Oct 2025", "نص إنجليزي مع سنة"),
        
        # Timestamp
        (pd.Timestamp("2025-10-02"), "Pandas Timestamp"),
        (datetime(2025, 10, 2), "Python datetime"),
        
        # رقم Excel
        (45200, "رقم تسلسلي Excel"),
        
        # حالات حدّية
        (None, "قيمة فارغة"),
        ("", "نص فارغ"),
        ("أكتوبر", "شهر فقط بدون يوم"),
        ("32 أكتوبر", "يوم غير صالح"),
        (float('nan'), "NaN"),
    ]
    
    for cell_value, description in test_cases:
        result = parse_due_date_cell(cell_value, default_year=2025)
        print(f"\n{description}:")
        print(f"  المدخل: {repr(cell_value)}")
        print(f"  النتيجة: {result}")
    
    print("\n" + "=" * 60)
    print("✅ انتهى الاختبار")
    print("=" * 60)

# ============== تابع بقية الكود الأصلي مع الإصلاحات ==============

def in_range(d: Optional[date], start: Optional[date], end: Optional[date]) -> bool:
    """التحقق من أن التاريخ ضمن النطاق المحدد"""
    if not (start and end):
        return True
    if d is None:
        return False
    # تصحيح الترتيب إذا كان مقلوباً
    if start > end:
        start, end = end, start
    return start <= d <= end

# ---------------- Foundation ----------------
def setup_app():
    APP_TITLE = "إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم"

    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="https://i.imgur.com/XLef7tS.png",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("ingaz-app")

    defaults = {
        "analysis_results": None,
        "pivot_table": None,
        "font_info": None,
        "logo_path": None,
        "selected_sheets": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    if st.session_state.font_info is None:
        st.session_state.font_info = prepare_default_font()

    # ---------- CSS ----------
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
    * { font-family: 'Cairo','Segoe UI',-apple-system,sans-serif }
    .main, body, .stApp { background:#fff; }
    .header-container{
      background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
      padding:44px 36px;color:#fff;text-align:center;margin-bottom:18px;
      border-bottom:4px solid #C9A646;box-shadow:0 6px 20px rgba(138,21,56,.25);position:relative
    }
    .header-container::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;
      background:linear-gradient(90deg,#C9A646 0%,#E8D4A0 50%,#C9A646 100%)}
    .header-container h1{margin:0 0 6px 0;font-size:32px;font-weight:800}
    .header-container .subtitle{font-size:15px;font-weight:700;margin:0 0 4px}
    .header-container .accent-line{font-size:12px;color:#C9A646;font-weight:700;margin:0 0 6px}
    .header-container .description{font-size:12px;opacity:.95;margin:0}

    [data-testid="stSidebar"]{
      background:linear-gradient(180deg,#8A1538 0%,#6B1029 100%)!important;
      border-right:2px solid #C9A646;box-shadow:4px 0 16px rgba(0,0,0,.15)
    }
    [data-testid="stSidebar"] *{ color:#fff !important; }

    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] select {
      color:#000 !important; background:#fff !important; caret-color:#000 !important;
    }
    [data-testid="stSidebar"] div[role="combobox"] input{ color:#000 !important; background:#fff !important; }
    [data-testid="stSidebar"] .stTextInput input,
    [data-testid="stSidebar"] .stNumberInput input{ color:#000 !important; background:#fff !important; }
    [data-testid="stSidebar"] ::placeholder{ color:#444 !important; opacity:1 !important; }
    [data-testid="stSidebar"] .stTextInput > div > div,
    [data-testid="stSidebar"] .stNumberInput > div > div{ border:1px solid rgba(0,0,0,.2) !important; box-shadow:none !important; }

    .chart-container{background:#fff;border:2px solid #E5E7EB;border-right:5px solid #8A1538;
      border-radius:12px;padding:16px;margin:12px 0;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .chart-title{font-size:20px;font-weight:800;color:#8A1538;text-align:center;margin-bottom:10px}

    .footer{margin-top:22px;background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
      color:#fff;border-radius:10px;padding:12px 10px;text-align:center;box-shadow:0 6px 18px rgba(138,21,56,.20);position:relative}
    .footer .line{width:100%;height:3px;background:linear-gradient(90deg,#C9A646 0%,#E8D4A0 50%,#C9A646 100%);
      position:absolute;top:0;left:0}
    .footer .school{font-weight:800;font-size:15px;margin:2px 0 4px}
    .footer .rights{font-weight:700;font-size:12px;margin:0 0 4px;opacity:.95}
    .footer .contact{font-size:12px;margin-top:2px}
    .footer a{color:#E8D4A0;font-weight:700;text-decoration:none;border-bottom:1px solid #C9A646}
    .footer .credit{margin-top:6px;font-size:11px;opacity:.85}
    </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class='header-container'>
      <div style='display:flex; align-items:center; justify-content:center; gap:16px; margin-bottom: 10px;'>
        <svg width="44" height="44" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect x="4" y="4" width="40" height="40" rx="4" fill="#C9A646" opacity="0.15"/>
          <path d="M12 32V24M18 32V20M24 32V16M30 32V22M36 32V18" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
          <circle cx="12" cy="24" r="2.5" fill="#C9A646"/><circle cx="18" cy="20" r="2.5" fill="#C9A646"/>
          <circle cx="24" cy="16" r="2.5" fill="#C9A646"/><circle cx="30" cy="22" r="2.5" fill="#C9A646"/>
          <circle cx="36" cy="18" r="2.5" fill="#C9A646"/>
          <path d="M12 24L18 20L24 16L30 22L36 18" stroke="#C9A646" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <h1>{APP_TITLE}</h1>
      </div>
      <p class='subtitle'>لوحة مهنية لقياس التقدم وتحليل النتائج</p>
      <p class='accent-line'>هوية إنجاز • دعم العربية الكامل</p>
      <p class='description'>فلترة الأعمدة حسب تاريخ الاستحقاق (يُقرأ من الخلية H3 لكل عمود)</p>
    </div>
    """, unsafe_allow_html=True)

    return logger

# ---------- Utilities ----------
def rtl(text: str) -> str:
    """تحويل النص العربي إلى RTL للعرض الصحيح"""
    if not isinstance(text, str):
        text = str(text)
    if AR_OK:
        return get_display(arabic_reshaper.reshape(text))
    return text

def prepare_default_font() -> Tuple[str, Optional[str]]:
    """تحضير الخط الافتراضي للـ PDF"""
    font_name = "ARFont"
    candidate = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    if os.path.exists(candidate):
        return font_name, candidate
    return "", None

def prepare_logo_file(logo_file) -> Optional[str]:
    """حفظ شعار المدرسة مؤقتاً"""
    if logo_file is None:
        return None
    try:
        ext = os.path.splitext(logo_file.name)[1].lower()
        if ext not in [".png", ".jpg", ".jpeg"]:
            return None
        path = f"/tmp/school_logo{ext}"
        with open(path, "wb") as f:
            f.write(logo_file.read())
        return path
    except Exception:
        return None

# ---------- PDF ----------
def make_student_pdf_fpdf(
    school_name: str,
    student_name: str,
    grade: str,
    section: str,
    table_df: pd.DataFrame,
    overall_avg: float,
    reco_text: str,
    coordinator_name: str,
    academic_deputy: str,
    admin_deputy: str,
    principal_name: str,
    font_info: Tuple[str, Optional[str]],
    logo_path: Optional[str] = None,
) -> bytes:
    """إنشاء تقرير PDF فردي للطالب"""
    font_name, font_path = font_info
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    if font_path:
        try:
            pdf.add_font(font_name, "", font_path, uni=True)
        except Exception:
            font_name = ""

    def set_font(size=12, color=(0,0,0)):
        if font_name:
            pdf.set_font(font_name, size=size)
        else:
            pdf.set_font("Helvetica", size=size)
        pdf.set_text_color(*color)

    # شريط أعلى + شعار
    pdf.set_fill_color(*QATAR_MAROON)
    pdf.rect(0, 0, 210, 20, style="F")
    if logo_path:
        try:
            pdf.image(logo_path, x=185, y=2.5, w=20)
        except Exception:
            pass
    
    set_font(14, (255,255,255))
    pdf.set_xy(10,7)
    pdf.cell(0,8, rtl("إنجاز - تقرير أداء الطالب"), align="R")

    # عنوان
    set_font(18, QATAR_MAROON)
    pdf.set_y(28)
    pdf.cell(0,10, rtl("تقرير أداء الطالب - نظام قطر للتعليم"), ln=1, align="R")
    pdf.set_draw_color(*QATAR_GOLD)
    pdf.set_line_width(0.6)
    pdf.line(30,38,200,38)

    # معلومات الطالب
    set_font(12, (0,0,0))
    pdf.ln(6)
    pdf.cell(0,8, rtl(f"اسم المدرسة: {school_name or '—'}"), ln=1, align="R")
    pdf.cell(0,8, rtl(f"اسم الطالب: {student_name}"), ln=1, align="R")
    pdf.cell(0,8, rtl(f"الصف: {grade or '—'}     الشعبة: {section or '—'}"), ln=1, align="R")
    pdf.ln(2)

    # جدول المواد
    headers = [rtl("المادة"), rtl("عدد التقييمات الإجمالي"), rtl("عدد التقييمات المنجزة"), rtl("عدد التقييمات المتبقية")]
    widths  = [70, 45, 45, 40]
    
    pdf.set_fill_color(*QATAR_MAROON)
    set_font(12, (255,255,255))
    pdf.set_y(pdf.get_y()+4)
    
    for w, h in zip(widths, headers):
        pdf.cell(w,9,h,border=0,align="C",fill=True)
    pdf.ln(9)

    set_font(11, (0,0,0))
    total_total = 0
    total_solved = 0
    
    for _, r in table_df.iterrows():
        sub = rtl(str(r['المادة']))
        tot = int(r['إجمالي'])
        solv = int(r['منجز'])
        rem = int(max(tot-solv, 0))
        
        total_total += tot
        total_solved += solv
        
        pdf.set_fill_color(247,247,247)
        pdf.cell(widths[0],8, sub, 0, 0, "C", True)
        pdf.cell(widths[1],8, str(tot), 0, 0, "C", True)
        pdf.cell(widths[2],8, str(solv), 0, 0, "C", True)
        pdf.cell(widths[3],8, str(rem), 0, 1, "C", True)

    # إحصائيات
    pdf.ln(3)
    set_font(12, QATAR_MAROON)
    pdf.cell(0,8, rtl("الإحصائيات"), ln=1, align="R")
    
    set_font(12, (0,0,0))
    pdf.cell(0,8, rtl(f"منجز: {total_solved}    متبقي: {max(total_total-total_solved,0)}    نسبة حل التقييمات: {overall_avg:.1f}%"), ln=1, align="R")

    # توصية
    pdf.ln(2)
    set_font(12, QATAR_MAROON)
    pdf.cell(0,8, rtl("توصية منسق المشاريع:"), ln=1, align="R")
    
    set_font(11, (0,0,0))
    for line in (reco_text or "—").splitlines() or ["—"]:
        pdf.multi_cell(0,7, rtl(line), align="R")

    # روابط مهمة
    pdf.ln(2)
    set_font(12, QATAR_MAROON)
    pdf.cell(0,8, rtl("روابط مهمة:"), ln=1, align="R")
    
    set_font(11, (0,0,0))
    pdf.cell(0,7, rtl("رابط نظام قطر: https://portal.education.qa"), ln=1, align="R")
    pdf.cell(0,7, rtl("استعادة كلمة المرور: https://password.education.qa"), ln=1, align="R")
    pdf.cell(0,7, rtl("قناة قطر للتعليم: https://edu.tv.qa"), ln=1, align="R")

    # توقيعات
    pdf.ln(4)
    set_font(12, QATAR_MAROON)
    pdf.cell(0,8, rtl("التوقيعات"), ln=1, align="R")
    
    set_font(11, (0,0,0))
    pdf.set_draw_color(*QATAR_GOLD)
    
    boxes = [
        ("منسق المشاريع", coordinator_name),
        ("النائب الأكاديمي", academic_deputy),
        ("النائب الإداري", admin_deputy),
        ("مدير المدرسة", principal_name)
    ]
    
    x_left, x_right = 10, 110
    y0 = pdf.get_y() + 2
    w, h = 90, 18
    
    for i, (title, name) in enumerate(boxes):
        row = i // 2
        col = i % 2
        x = x_right if col == 0 else x_left
        yb = y0 + row * (h + 6)
        
        pdf.rect(x, yb, w, h)
        set_font(11, (0,0,0))
        pdf.set_xy(x, yb+3)
        pdf.cell(w-4, 6, rtl(f"{title} / {name or '—'}"), align="R")
        pdf.set_xy(x, yb+10)
        pdf.cell(w-4, 6, rtl("التوقيع: __________________    التاريخ: __________"), align="R")

    # إخراج PDF
    try:
        out = pdf.output(dest="S")
        return out if isinstance(out, bytes) else out.encode("utf-8", "ignore")
    except Exception:
        # محاولة بديلة
        out = pdf.output(dest="S")
        return bytes(out) if not isinstance(out, bytes) else out

# ---------- Data Logic ----------
CATEGORY_COLORS = {
    'بلاتيني 🥇': '#E5E4E2',
    'ذهبي 🥈': '#C9A646',
    'فضي 🥉': '#C0C0C0',
    'برونزي': '#CD7F32',
    'بحاجة لتحسين': '#8A1538'
}
CATEGORY_ORDER = ['بلاتيني 🥇', 'ذهبي 🥈', 'فضي 🥉', 'برونزي', 'بحاجة لتحسين']

def parse_sheet_name(sheet_name: str):
    """استخراج المادة والصف والشعبة من اسم الورقة"""
    try:
        parts = sheet_name.strip().split()
        if len(parts) < 3:
            return sheet_name.strip(), "", ""
        
        section = parts[-1]
        level = parts[-2]
        subject = " ".join(parts[:-2])
        
        if not (level.isdigit() or (level.startswith('0') and len(level) <= 2)):
            subject = " ".join(parts[:-1])
            level = parts[-1]
            section = ""
        
        return subject, level, section
    except Exception:
        return sheet_name, "", ""

@st.cache_data(ttl=3600, max_entries=10)
def analyze_excel_file(file, sheet_name, due_start: Optional[date]=None, due_end: Optional[date]=None):
    """
    تحليل ورقة Excel واستخراج بيانات الطلاب
    
    - فلترة بالتاريخ باستخدام صف تاريخ الاستحقاق H3: df.iloc[2, col]
    - تجاهل أي عمود عنوانه يحتوي على شرطة '-' أو '—' أو '–'
    - تجاهل الأعمدة الفارغة تماماً
    - تجاهل الخلايا الفارغة أو التي تحتوي على شرطات
    - 'M' = مستحق غير منجز (يزيد الإجمالي ويُعد متبقّي)
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)

        filter_active = (due_start is not None and due_end is not None)
        if filter_active and due_start > due_end:
            due_start, due_end = due_end, due_start

        assessment_columns = []
        
        # البحث عن أعمدة التقييمات (بدءاً من H = العمود 7)
        for c in range(7, df.shape[1]):
            # 1) قراءة عنوان العمود
            title = df.iloc[0, c] if c < df.shape[1] else None
            if pd.isna(title):
                break
            
            t = str(title).strip()

            # 2) تجاهل العناوين التي تحتوي على شرطات
            if any(ch in t for ch in ['-', '—', '–']):
                continue

            # 3) قراءة تاريخ الاستحقاق من H3 (الصف 2، index=2)
            due_cell = df.iloc[2, c] if c < df.shape[1] else None
            due_dt = parse_due_date_cell(due_cell, default_year=date.today().year)
            
            # 4) فلترة حسب النطاق الزمني
            if filter_active and not in_range(due_dt, due_start, due_end):
                continue

            # 5) تجاهل الأعمدة الفارغة تماماً (فحص جميع الصفوف)
            all_dash = True
            for r in range(4, len(df)):  # ✅ إصلاح: فحص كل الصفوف
                if r >= df.shape[0]:
                    break
                val = df.iloc[r, c]
                if pd.notna(val):
                    s = str(val).strip().upper()
                    if s not in ['-', '—', '–', '', 'NAN', 'NONE']:
                        all_dash = False
                        break
            
            if all_dash:
                continue

            # إضافة العمود للتحليل
            assessment_columns.append({'index': c, 'title': t})

        if not assessment_columns:
            return []

        # معالجة بيانات الطلاب
        results = []
        IGNORE = {'-', '—', '–', '', 'NAN', 'NONE'}
        
        for r in range(4, len(df)):
            student = df.iloc[r, 0]
            if pd.isna(student) or str(student).strip() == "":
                continue
            
            name = " ".join(str(student).strip().split())

            total = 0
            done = 0
            pending = []
            
            for col in assessment_columns:
                c = col['index']
                title = col['title']
                
                if c >= df.shape[1]:
                    continue
                
                raw = df.iloc[r, c]
                s = "" if pd.isna(raw) else str(raw).strip().upper()

                # تجاهل الخلايا الفارغة والشرطات
                if s in IGNORE:
                    continue
                
                # معالجة 'M' = مستحق غير منجز
                if s == 'M':
                    total += 1
                    pending.append(title)
                    continue
                
                # تقييم منجز
                total += 1
                done += 1

            # حساب النسبة المئوية
            pct = (done / total * 100) if total > 0 else 0.0
            
            results.append({
                "student_name": name,
                "subject": subject,
                "level": str(level_from_name).strip(),
                "section": str(section_from_name).strip(),
                "solve_pct": round(pct, 1),
                "completed_count": int(done),
                "total_count": int(total),
                "pending_titles": ", ".join(pending) if pending else "-"
            })
        
        return results

    except Exception as e:
        st.error(f"❌ خطأ في تحليل الملف '{sheet_name}': {e}")
        return []

@st.cache_data
def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    """إنشاء جدول محوري من بيانات التحليل"""
    try:
        if df.empty:
            return pd.DataFrame()
        
        # إزالة التكرارات
        dfc = df.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'])
        
        # الطلاب الفريدين
        unq = dfc[['student_name', 'level', 'section']].drop_duplicates()
        unq = unq.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
        res = unq.copy()
        
        # إضافة بيانات كل مادة
        for subject in sorted(dfc['subject'].unique()):
            sub = dfc[dfc['subject'] == subject].copy()
            
            # ملء القيم الفارغة بصفر
            sub[['total_count', 'completed_count', 'solve_pct']] = sub[['total_count', 'completed_count', 'solve_pct']].fillna(0)
            
            # دمج الأعمدة الرقمية
            block = sub[['student_name', 'level', 'section', 'total_count', 'completed_count', 'solve_pct']].rename(columns={
                'total_count': f'{subject} - إجمالي',
                'completed_count': f'{subject} - منجز',
                'solve_pct': f'{subject} - النسبة'
            }).drop_duplicates(subset=['student_name', 'level', 'section'])
            
            res = res.merge(block, on=['student_name', 'level', 'section'], how='left')
            
            # دمج التقييمات المتبقية
            pend = sub[['student_name', 'level', 'section', 'pending_titles']].drop_duplicates(
                subset=['student_name', 'level', 'section']
            ).rename(columns={'pending_titles': f'{subject} - متبقي'})
            
            res = res.merge(pend, on=['student_name', 'level', 'section'], how='left')

        # حساب المتوسط والفئة
        pct_cols = [c for c in res.columns if 'النسبة' in c]
        if pct_cols:
            # ✅ إصلاح: معالجة NaN بشكل صحيح
            res['المتوسط'] = res[pct_cols].apply(lambda row: row.dropna().mean() if row.notna().any() else 0, axis=1)
            
            def cat(p):
                if pd.isna(p) or p == 0:
                    return 'بحاجة لتحسين'
                elif p >= 90:
                    return 'بلاتيني 🥇'
                elif p >= 80:
                    return 'ذهبي 🥈'
                elif p >= 70:
                    return 'فضي 🥉'
                elif p >= 60:
                    return 'برونزي'
                else:
                    return 'بحاجة لتحسين'
            
            res['الفئة'] = res['المتوسط'].apply(cat)

        # إعادة تسمية الأعمدة
        res = res.rename(columns={
            'student_name': 'الطالب',
            'level': 'الصف',
            'section': 'الشعبة'
        })
        
        # تنسيق الأعمدة
        for c in res.columns:
            if ('إجمالي' in c) or ('منجز' in c):
                res[c] = res[c].fillna(0).astype(int)
            elif ('النسبة' in c) or (c == 'المتوسط'):
                res[c] = res[c].fillna(0).round(1)
            elif 'متبقي' in c:
                res[c] = res[c].fillna('-')
        
        return res.drop_duplicates(subset=['الطالب', 'الصف', 'الشعبة']).reset_index(drop=True)
    
    except Exception as e:
        st.error(f"❌ خطأ في معالجة البيانات: {e}")
        return pd.DataFrame()

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """تطبيع البيانات وإضافة الفئات"""
    out = df.rename(columns={'solve_pct': 'percent', 'student_name': 'student'})
    
    def cat(p):
        if pd.isna(p):
            return 'بحاجة لتحسين'
        elif p >= 90:
            return 'بلاتيني 🥇'
        elif p >= 80:
            return 'ذهبي 🥈'
        elif p >= 70:
            return 'فضي 🥉'
        elif p >= 60:
            return 'برونزي'
        else:
            return 'بحاجة لتحسين'
    
    out['category'] = out['percent'].apply(cat)
    return out

def aggregate_by_subject(df: pd.DataFrame) -> pd.DataFrame:
    """تجميع البيانات حسب المادة"""
    rows = []
    
    for s in df['subject'].dropna().unique():
        sub = df[df['subject'] == s]
        n = len(sub)
        # ✅ إصلاح: معالجة NaN بشكل صحيح
        avg = sub['percent'].mean() if n > 0 and sub['percent'].notna().any() else 0.0
        
        for cat in CATEGORY_ORDER:
            c = (sub['category'] == cat).sum()
            pct = (c / n * 100) if n > 0 else 0.0
            
            rows.append({
                'subject': s,
                'category': cat,
                'count': int(c),
                'percent_share': round(pct, 1),
                'avg_completion': round(avg, 1)
            })
    
    agg = pd.DataFrame(rows)
    if agg.empty:
        return agg
    
    # ترتيب المواد حسب متوسط الإنجاز
    order = agg.groupby('subject')['avg_completion'].first().sort_values(ascending=False).index.tolist()
    agg['subject'] = pd.Categorical(agg['subject'], categories=order, ordered=True)
    
    return agg.sort_values('subject')

def chart_stacked_by_subject(agg_df: pd.DataFrame, mode='percent') -> go.Figure:
    """رسم بياني مكدس حسب المادة"""
    fig = go.Figure()
    colors = [CATEGORY_COLORS[c] for c in CATEGORY_ORDER]
    
    for i, cat in enumerate(CATEGORY_ORDER):
        d = agg_df[agg_df['category'] == cat]
        vals = d['percent_share'] if mode == 'percent' else d['count']
        text = [(f"{v:.1f}%" if mode == 'percent' else str(v)) if v > 0 else "" for v in vals]
        hover = "<b>%{y}</b><br>الفئة: " + cat + "<br>" + (
            "النسبة: %{x:.1f}%<extra></extra>" if mode == 'percent' else "العدد: %{x}<extra></extra>"
        )
        
        fig.add_trace(go.Bar(
            name=cat,
            x=vals,
            y=d['subject'],
            orientation='h',
            marker=dict(color=colors[i], line=dict(color='white', width=1)),
            text=text,
            textposition='inside',
            textfont=dict(size=11, family='Cairo'),
            hovertemplate=hover
        ))
    
    fig.update_layout(
        title=dict(
            text="توزيع الفئات حسب المادة",
            font=dict(size=20, family='Cairo', color='#8A1538'),
            x=0.5
        ),
        xaxis=dict(
            title="النسبة المئوية (%)" if mode == 'percent' else "عدد الطلاب",
            tickfont=dict(size=12, family='Cairo'),
            gridcolor='#E5E7EB',
            range=[0, 100] if mode == 'percent' else None
        ),
        yaxis=dict(
            title="المادة",
            tickfont=dict(size=12, family='Cairo'),
            autorange='reversed'
        ),
        barmode='stack',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Cairo')
    )
    
    return fig

def chart_overall_donut(pivot: pd.DataFrame) -> go.Figure:
    """رسم دائري للتوزيع العام"""
    if 'الفئة' not in pivot.columns or pivot.empty:
        return go.Figure()
    
    counts = pivot['الفئة'].value_counts().reindex(CATEGORY_ORDER, fill_value=0)
    
    fig = go.Figure([go.Pie(
        labels=counts.index,
        values=counts.values,
        hole=0.55,
        marker=dict(colors=[CATEGORY_COLORS[k] for k in counts.index]),
        textinfo='label+value',
        hovertemplate="%{label}: %{value} طالب<extra></extra>"
    )])
    
    fig.update_layout(
        title=dict(
            text="توزيع عام للفئات",
            font=dict(size=20, family='Cairo', color='#8A1538'),
            x=0.5
        ),
        showlegend=False,
        font=dict(family='Cairo')
    )
    
    return fig

def chart_overall_gauge(pivot: pd.DataFrame) -> go.Figure:
    """مؤشر متوسط الإنجاز"""
    avg = 0.0
    if 'المتوسط' in pivot.columns and not pivot.empty:
        avg = float(pivot['المتوسط'].mean())
        if pd.isna(avg):
            avg = 0.0
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=avg,
        number={'suffix': "%", 'font': {'family': 'Cairo'}},
        gauge={'axis': {'range': [0, 100]}, 'bar': {'color': '#8A1538'}}
    ))
    
    fig.update_layout(
        title=dict(
            text="متوسط الإنجاز العام",
            font=dict(size=20, family='Cairo', color='#8A1538'),
            x=0.5
        ),
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(family='Cairo'),
        height=320
    )
    
    return fig

# ================== Run App ==================
logger = setup_app()

# Sidebar
with st.sidebar:
    st.image("https://i.imgur.com/XLef7tS.png", width=110)
    st.markdown("---")
    st.header("⚙️ الإعدادات")

    # تحميل الملفات + فلترة الأوراق
    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    selected_sheets: List[tuple] = []
    all_sheets = []
    sheet_file_map = {}
    
    if uploaded_files:
        for file_idx, file in enumerate(uploaded_files):
            try:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    label = f"[ملف {file_idx+1}] {sheet}"
                    all_sheets.append(label)
                    sheet_file_map[label] = (file, sheet)
            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف: {e}")

        if all_sheets:
            st.info(f"📋 وُجدت {len(all_sheets)} ورقة في {len(uploaded_files)} ملف")
            select_all = st.checkbox("✔️ اختر الجميع", value=True, key="select_all_sheets")
            
            if select_all:
                chosen = all_sheets
            else:
                chosen = st.multiselect(
                    "اختر الأوراق للتحليل",
                    all_sheets,
                    default=all_sheets[:1] if all_sheets else []
                )
            
            selected_sheets = [sheet_file_map[c] for c in chosen]

    st.session_state.selected_sheets = selected_sheets

    # فلتر تاريخ الاستحقاق
    st.subheader("⏳ فلترة تاريخ الاستحقاق (من — إلى)")
    default_start = date.today().replace(day=1)
    default_end = date.today()
    
    range_val = st.date_input(
        "اختر المدى",
        value=(default_start, default_end),
        format="YYYY-MM-DD",
        key="due_range"
    )
    
    if isinstance(range_val, (list, tuple)) and len(range_val) >= 2:
        due_start, due_end = range_val[0], range_val[1]
    else:
        due_start, due_end = None, None

    # شعار المدرسة
    st.subheader("🖼️ شعار المدرسة (اختياري)")
    logo_file = st.file_uploader(
        "ارفع شعار PNG/JPG",
        type=["png", "jpg", "jpeg"],
        key="logo_file"
    )
    st.session_state.logo_path = prepare_logo_file(logo_file)

    st.markdown("---")
    st.subheader("🏫 معلومات المدرسة")
    school_name = st.text_input("اسم المدرسة", placeholder="مدرسة قطر النموذجية")
    
    st.subheader("✍️ التوقيعات")
    coordinator_name = st.text_input("منسق/ة المشاريع")
    academic_deputy = st.text_input("النائب الأكاديمي")
    admin_deputy = st.text_input("النائب الإداري")
    principal_name = st.text_input("مدير/ة المدرسة")

    st.markdown("---")
    run_analysis = st.button(
        "▶️ تشغيل التحليل",
        use_container_width=True,
        type="primary",
        disabled=not uploaded_files
    )

# تحليل
if not uploaded_files:
    st.info("📤 من الشريط الجانبي ارفع ملفات Excel للبدء في التحليل")
elif run_analysis:
    sheets_to_use = st.session_state.selected_sheets
    
    if not sheets_to_use:
        tmp = []
        for file in uploaded_files:
            try:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    tmp.append((file, sheet))
            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف: {e}")
        sheets_to_use = tmp

    if not sheets_to_use:
        st.warning("⚠️ لم يتم العثور على أوراق داخل الملفات المرفوعة.")
    else:
        with st.spinner("⏳ جاري التحليل..."):
            rows = []
            for file, sheet in sheets_to_use:
                rows.extend(analyze_excel_file(file, sheet, due_start, due_end))
            
            if rows:
                df = pd.DataFrame(rows)
                st.session_state.analysis_results = df
                st.session_state.pivot_table = create_pivot_table(df)
                st.success(
                    f"✅ تم تحليل {len(st.session_state.pivot_table)} طالب عبر "
                    f"{df['subject'].nunique()} مادة"
                )
            else:
                st.warning(
                    "⚠️ لم يتم استخراج بيانات من الأوراق المحددة. "
                    "تأكد من تنسيق الجداول وتواريخ الاستحقاق."
                )

# عرض النتائج
pivot = st.session_state.pivot_table
df = st.session_state.analysis_results

if pivot is not None and not pivot.empty and df is not None:
    st.subheader("📈 ملخص النتائج")
    
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        st.metric("👥 إجمالي الطلاب", len(pivot))
    
    with c2:
        st.metric("📚 عدد المواد", df['subject'].nunique())
    
    with c3:
        avg = 0.0
        if 'المتوسط' in pivot.columns:
            avg = float(pivot['المتوسط'].mean())
            if pd.isna(avg):
                avg = 0.0
        st.metric("📊 متوسط الإنجاز", f"{avg:.1f}%")
    
    with c4:
        platinum_count = int((pivot['الفئة'] == 'بلاتيني 🥇').sum())
        st.metric("🥇 فئة بلاتيني", platinum_count)
    
    with c5:
        zero = 0
        if 'المتوسط' in pivot.columns:
            zero = int((pivot['المتوسط'] == 0).sum())
        st.metric("⚠️ بدون إنجاز", zero)

    st.divider()
    st.subheader("📋 جدول النتائج التفصيلي")
    st.dataframe(pivot, use_container_width=True, height=420)

    st.divider()
    
    # الرسومات البيانية
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">🍩 التوزيع العام للفئات</h2>', unsafe_allow_html=True)
    st.plotly_chart(chart_overall_donut(pivot), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">🎯 مؤشر متوسط الإنجاز</h2>', unsafe_allow_html=True)
    st.plotly_chart(chart_overall_gauge(pivot), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">📊 توزيع الفئات حسب المادة الدراسية</h2>', unsafe_allow_html=True)
    
    try:
        normalized = normalize_dataframe(df)
        mode_choice = st.radio(
            'نوع العرض',
            ['النسبة المئوية (%)', 'العدد المطلق'],
            horizontal=True,
            key="chart_mode"
        )
        mode = 'percent' if mode_choice == 'النسبة المئوية (%)' else 'count'
        agg_df = aggregate_by_subject(normalized)
        st.plotly_chart(chart_stacked_by_subject(agg_df, mode=mode), use_container_width=True)
    except Exception as e:
        st.error(f"خطأ في الرسم: {e}")
    
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # التقارير الفردية
    st.subheader("📑 التقارير الفردية (PDF)")
    
    students = sorted(pivot['الطالب'].dropna().astype(str).unique().tolist())
    
    if students:
        csel, crec = st.columns([2, 3])
        
        with csel:
            sel = st.selectbox("اختر الطالب", students, index=0)
            row = pivot[pivot['الطالب'] == sel].head(1)
            g = str(row['الصف'].iloc[0]) if not row.empty else ''
            s = str(row['الشعبة'].iloc[0]) if not row.empty else ''
        
        with crec:
            reco = st.text_area(
                "توصية منسق المشاريع",
                value="",
                height=120,
                placeholder="اكتب التوصيات هنا..."
            )

        sdata = df[df['student_name'].str.strip().eq(sel.strip())].copy()
        
        table = sdata[['subject', 'total_count', 'completed_count']].rename(columns={
            'subject': 'المادة',
            'total_count': 'إجمالي',
            'completed_count': 'منجز'
        })
        
        table['متبقي'] = (table['إجمالي'] - table['منجز']).clip(lower=0).astype(int)
        avg_stu = float(sdata['solve_pct'].mean()) if not sdata.empty else 0.0

        st.markdown("### معاينة سريعة")
        st.dataframe(table, use_container_width=True, height=260)

        pdf_one = make_student_pdf_fpdf(
            school_name=school_name or "",
            student_name=sel,
            grade=g,
            section=s,
            table_df=table[['المادة', 'إجمالي', 'منجز', 'متبقي']],
            overall_avg=avg_stu,
            reco_text=reco,
            coordinator_name=coordinator_name or "",
            academic_deputy=academic_deputy or "",
            admin_deputy=admin_deputy or "",
            principal_name=principal_name or "",
            font_info=st.session_state.font_info,
            logo_path=st.session_state.logo_path
        )
        
        if not isinstance(pdf_one, bytes):
            pdf_one = bytes(pdf_one)

        st.download_button(
            "📥 تحميل تقرير الطالب (PDF)",
            pdf_one,
            file_name=f"student_report_{sel}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

        st.markdown("---")
        st.subheader("📦 تصدير جميع التقارير (ZIP)")
        
        same_reco = st.checkbox("استخدم نفس التوصية لكل الطلاب", value=True)
        
        if st.button("إنشاء ملف ZIP لكل التقارير", type="primary"):
            with st.spinner("جاري إنشاء حزمة التقارير..."):
                buf = io.BytesIO()
                
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                    for stu in students:
                        r = pivot[pivot['الطالب'] == stu].head(1)
                        g = str(r['الصف'].iloc[0]) if not r.empty else ''
                        s = str(r['الشعبة'].iloc[0]) if not r.empty else ''
                        
                        sd = df[df['student_name'].str.strip().eq(stu.strip())].copy()
                        
                        t = sd[['subject', 'total_count', 'completed_count']].rename(columns={
                            'subject': 'المادة',
                            'total_count': 'إجمالي',
                            'completed_count': 'منجز'
                        })
                        
                        t['متبقي'] = (t['إجمالي'] - t['منجز']).clip(lower=0).astype(int)
                        av = float(sd['solve_pct'].mean()) if not sd.empty else 0.0
                        
                        rtext = reco if same_reco else ""
                        
                        pdfb = make_student_pdf_fpdf(
                            school_name=school_name or "",
                            student_name=stu,
                            grade=g,
                            section=s,
                            table_df=t[['المادة', 'إجمالي', 'منجز', 'متبقي']],
                            overall_avg=av,
                            reco_text=rtext,
                            coordinator_name=coordinator_name or "",
                            academic_deputy=academic_deputy or "",
                            admin_deputy=admin_deputy or "",
                            principal_name=principal_name or "",
                            font_info=st.session_state.font_info,
                            logo_path=st.session_state.logo_path
                        )
                        
                        if not isinstance(pdfb, bytes):
                            pdfb = bytes(pdfb)
                        
                        safe = re.sub(r"[^\w\-]+", "_", str(stu))
                        z.writestr(f"{safe}.pdf", pdfb)
                
                buf.seek(0)
                
                st.download_button(
                    "⬇️ تحميل الحزمة (ZIP)",
                    buf.getvalue(),
                    file_name=f"student_reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

# Footer
st.markdown(f"""
<div class="footer">
  <div class="line"></div>
  <div class="school">مدرسة عثمان بن عفان النموذجية للبنين</div>
  <div class="rights">© {datetime.now().year} جميع الحقوق محفوظة</div>
  <div class="contact">للتواصل:
    <a href="mailto:S.mahgoub0101@education.qa">S.mahgoub0101@education.qa</a>
  </div>
  <div class="credit">تطوير وتصميم: قسم التحول الرقمي</div>
</div>
""", unsafe_allow_html=True)

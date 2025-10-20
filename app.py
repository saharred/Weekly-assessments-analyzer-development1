# -*- coding: utf-8 -*-
"""
تطبيق إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم
النسخة المحسّنة 2.0
"""

import os
import io
import re
import zipfile
import logging
import unicodedata
import warnings
from datetime import datetime, date
from typing import Tuple, Optional, List, Dict, Any
from functools import wraps
import time

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from fpdf import FPDF

# استيراد مكتبات العربية مع معالجة أفضل
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    AR_OK = True
except ImportError:
    AR_OK = False
    warnings.warn(
        "⚠️ مكتبات arabic_reshaper غير متوفرة - قد يتأثر عرض النصوص العربية في PDF\n"
        "للتثبيت: pip install arabic-reshaper python-bidi"
    )

# ============== الإعدادات والثوابت ==============

QATAR_MAROON = (138, 21, 56)
QATAR_GOLD = (201, 166, 70)

CATEGORY_COLORS = {
    'بلاتيني 🥇': '#E5E4E2',
    'ذهبي 🥈': '#C9A646',
    'فضي 🥉': '#C0C0C0',
    'برونزي': '#CD7F32',
    'بحاجة لتحسين': '#8A1538'
}

CATEGORY_ORDER = ['بلاتيني 🥇', 'ذهبي 🥈', 'فضي 🥉', 'برونزي', 'بحاجة لتحسين']

# إعداد Logging احترافي
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger("ingaz-app")

# ============== ديكورايتورات مساعدة ==============

def log_performance(func):
    """ديكورايتور لقياس أداء الدوال المهمة"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start = time.time()
        func_name = func.__name__
        logger.info(f"🔄 بدء تنفيذ: {func_name}")
        
        try:
            result = func(*args, **kwargs)
            duration = time.time() - start
            logger.info(f"✅ اكتمل {func_name} في {duration:.2f} ثانية")
            return result
        except Exception as e:
            duration = time.time() - start
            logger.error(f"❌ فشل {func_name} بعد {duration:.2f} ثانية: {e}")
            raise
    
    return wrapper

def safe_execute(default_return=None, error_message="حدث خطأ"):
    """ديكورايتور لمعالجة الأخطاء بشكل آمن"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger.error(f"{error_message} في {func.__name__}: {e}")
                if st:
                    st.error(f"{error_message}: {str(e)}")
                return default_return
        return wrapper
    return decorator

# ============== دوال معالجة التاريخ ==============

def _normalize_arabic_digits(s: str) -> str:
    """تحويل الأرقام العربية إلى إنجليزية"""
    if not isinstance(s, str):
        return str(s) if s is not None else ""
    return s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

def _strip_invisible_and_diacritics(s: str) -> str:
    """إزالة الأحرف غير المرئية والتشكيل"""
    if not isinstance(s, str):
        return ""
    
    # الأحرف غير المرئية الشائعة
    invisible_chars = [
        '\u200e', '\u200f', '\u202a', '\u202b', '\u202c', '\u202d', '\u202e',
        '\u2066', '\u2067', '\u2068', '\u2069', '\u200b', '\u200c', '\u200d',
        '\ufeff', '\xa0', '\u0640',
    ]
    
    for char in invisible_chars:
        s = s.replace(char, '')
    
    # إزالة التشكيل
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = ' '.join(s.split())
    
    return s.strip()

def parse_due_date_cell(cell, default_year: int = None) -> Optional[date]:
    """
    معالجة محسّنة للتواريخ من خلايا Excel
    ✅ يدعم: "2 أكتوبر"، "أكتوبر 19"، "19-10"، إلخ
    
    Args:
        cell: قيمة الخلية (قد تكون نص، رقم، تاريخ)
        default_year: السنة الافتراضية إذا لم تُحدد
    
    Returns:
        Optional[date]: التاريخ المعالج أو None
    """
    if default_year is None:
        default_year = date.today().year
    
    # معالجة القيم الفارغة
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    
    # معالجة التواريخ الجاهزة
    if isinstance(cell, (pd.Timestamp, datetime)):
        try:
            return cell.date() if hasattr(cell, 'date') else cell
        except (ValueError, AttributeError) as e:
            logger.warning(f"خطأ في تحويل timestamp: {e}")
            return None
    
    # معالجة الأرقام (Excel serial dates)
    if isinstance(cell, (int, float)) and not pd.isna(cell):
        try:
            if 1 <= cell <= 100000:
                base = pd.to_datetime("1899-12-30")
                result = base + pd.to_timedelta(float(cell), unit="D")
                if 1900 <= result.year <= 2200:
                    return result.date()
        except (ValueError, OverflowError) as e:
            logger.warning(f"خطأ في معالجة رقم تاريخ: {e}")
    
    # معالجة النصوص
    try:
        s = str(cell).strip()
        if not s or s.lower() in ['nan', 'none', 'nat', 'null']:
            return None
        
        s = _strip_invisible_and_diacritics(s)
        s = _normalize_arabic_digits(s)
        
        if not s:
            return None
        
        # قاموس الأشهر العربية الموسع
        arabic_months = {
            # يناير
            "يناير": 1, "كانون الثاني": 1, "جانفي": 1, "ينايرJanuary": 1,
            "jan": 1, "january": 1, "يناير": 1,
            # فبراير
            "فبراير": 2, "شباط": 2, "فيفري": 2, "فبرايرFebruary": 2,
            "feb": 2, "february": 2, "فبراير": 2,
            # مارس
            "مارس": 3, "اذار": 3, "آذار": 3, "مارسMarch": 3,
            "mar": 3, "march": 3, "مارس": 3,
            # أبريل
            "ابريل": 4, "أبريل": 4, "نيسان": 4, "افريل": 4, "ابريلApril": 4,
            "apr": 4, "april": 4, "ابريل": 4,
            # مايو
            "مايو": 5, "ماي": 5, "ايار": 5, "أيار": 5, "مايوMay": 5,
            "may": 5, "مايو": 5,
            # يونيو
            "يونيو": 6, "يونيه": 6, "حزيران": 6, "جوان": 6, "يونيوJune": 6,
            "jun": 6, "june": 6, "يونيو": 6,
            # يوليو
            "يوليو": 7, "يوليه": 7, "تموز": 7, "جويلية": 7, "يوليوJuly": 7,
            "jul": 7, "july": 7, "يوليو": 7,
            # أغسطس
            "اغسطس": 8, "أغسطس": 8, "اب": 8, "آب": 8, "اوت": 8, "اغسطسAugust": 8,
            "aug": 8, "august": 8, "اغسطس": 8,
            # سبتمبر
            "سبتمبر": 9, "ايلول": 9, "أيلول": 9, "سبتمبرSeptember": 9,
            "sep": 9, "sept": 9, "september": 9, "سبتمبر": 9,
            # أكتوبر
            "اكتوبر": 10, "أكتوبر": 10, "تشرين الاول": 10, "تشرين الأول": 10, "اكتوبرOctober": 10,
            "oct": 10, "october": 10, "اكتوبر": 10,
            # نوفمبر
            "نوفمبر": 11, "تشرين الثاني": 11, "نونبر": 11, "نوفمبرNovember": 11,
            "nov": 11, "november": 11, "نوفمبر": 11,
            # ديسمبر
            "ديسمبر": 12, "كانون الاول": 12, "كانون الأول": 12, "دجنبر": 12, "ديسمبرDecember": 12,
            "dec": 12, "december": 12, "ديسمبر": 12,
        }
        
        def normalize_hamza(text):
            """توحيد الهمزات والألفات"""
            text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
            text = text.replace("ة", "ه").replace("ـ", "")
            return text
        
        # ✅ نمط 1: "أكتوبر 19" أو "19 أكتوبر"
        # يدعم: "اكتوبر 19"، "19 اكتوبر"، "October 19"
        pattern1 = r"(\d{1,2})\s*[-/،,\s]*\s*([^\d\s]+)"
        match1 = re.search(pattern1, s)
        
        if match1:
            day = int(match1.group(1))
            month_name = match1.group(2).strip()
            
            # البحث عن الشهر
            month = arabic_months.get(month_name)
            
            if not month:
                # محاولة مع تطبيع الهمزات
                normalized_name = normalize_hamza(month_name)
                for key, val in arabic_months.items():
                    if normalize_hamza(key) == normalized_name:
                        month = val
                        break
            
            if month:
                try:
                    result_date = date(default_year, month, day)
                    logger.debug(f"✅ تم تحليل التاريخ: '{s}' → {result_date}")
                    return result_date
                except ValueError:
                    # إذا كان اليوم غير صالح، استخدم يوم آمن
                    safe_day = min(day, 28)
                    try:
                        return date(default_year, month, safe_day)
                    except ValueError:
                        logger.warning(f"تاريخ غير صالح: {day}/{month}/{default_year}")
        
        # ✅ نمط 2: "أكتوبر 19" (عكس)
        pattern2 = r"([^\d\s]+)\s*[-/،,\s]*\s*(\d{1,2})"
        match2 = re.search(pattern2, s)
        
        if match2:
            month_name = match2.group(1).strip()
            day = int(match2.group(2))
            
            month = arabic_months.get(month_name)
            
            if not month:
                normalized_name = normalize_hamza(month_name)
                for key, val in arabic_months.items():
                    if normalize_hamza(key) == normalized_name:
                        month = val
                        break
            
            if month:
                try:
                    result_date = date(default_year, month, day)
                    logger.debug(f"✅ تم تحليل التاريخ: '{s}' → {result_date}")
                    return result_date
                except ValueError:
                    safe_day = min(day, 28)
                    try:
                        return date(default_year, month, safe_day)
                    except ValueError:
                        logger.warning(f"تاريخ غير صالح: {day}/{month}/{default_year}")
        
        # ✅ نمط 3: محاولة pandas (للتواريخ الرقمية)
        try:
            parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(parsed):
                result_date = parsed.date()
                if parsed.year < 1900:
                    result_date = result_date.replace(year=default_year)
                logger.debug(f"✅ تم تحليل التاريخ (pandas): '{s}' → {result_date}")
                return result_date
        except Exception:
            pass
    
    except Exception as e:
        logger.warning(f"فشل في معالجة التاريخ '{cell}': {e}")
    
    # إذا فشل كل شيء
    logger.debug(f"⚠️ لم يتم التعرف على التاريخ: '{cell}'")
    return None

def in_range(d: Optional[date], start: Optional[date], end: Optional[date]) -> bool:
    """التحقق من وقوع التاريخ في النطاق المحدد"""
    if not (start and end):
        return True
    if d is None:
        return False
    if start > end:
        start, end = end, start
    return start <= d <= end

# ============== دوال التحقق من البيانات ==============

def validate_excel_structure(df: pd.DataFrame, sheet_name: str) -> Tuple[bool, str]:
    """
    التحقق من صحة بنية ملف Excel
    
    Returns:
        Tuple[bool, str]: (هل صالح؟, رسالة الخطأ إن وُجد)
    """
    if df is None or df.empty:
        return False, "الملف فارغ"
    
    if df.shape[0] < 4:
        return False, f"عدد الصفوف قليل جداً ({df.shape[0]} صف)"
    
    if df.shape[1] < 8:
        return False, f"عدد الأعمدة قليل جداً ({df.shape[1]} عمود)"
    
    # التحقق من وجود أسماء الطلاب
    student_col = df.iloc[4:, 0].dropna()
    if len(student_col) == 0:
        return False, "لا توجد أسماء طلاب في العمود الأول"
    
    return True, ""

# ============== إعداد التطبيق ==============

def setup_app():
    """إعداد واجهة Streamlit"""
    APP_TITLE = "إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم"

    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="https://i.imgur.com/XLef7tS.png",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # إعداد session state
    defaults = {
        "analysis_results": None,
        "pivot_table": None,
        "font_info": None,
        "logo_path": None,
        "selected_sheets": [],
        "analysis_stats": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    if st.session_state.font_info is None:
        st.session_state.font_info = prepare_default_font()

    # تطبيق الأنماط CSS
    apply_custom_styles()
    
    # عرض الهيدر
    render_header(APP_TITLE)

def apply_custom_styles():
    """تطبيق أنماط CSS المخصصة"""
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
    * { font-family: 'Cairo','Segoe UI',-apple-system,sans-serif; direction: rtl; }
    .main, body, .stApp { background:#fff; direction: rtl; }
    
    section[data-testid="stSidebar"] {
        right: 0 !important;
        left: auto !important;
    }
    
    .main .block-container {
        padding-right: 5rem !important;
        padding-left: 1rem !important;
    }
    
    .header-container{
      background:linear-gradient(135deg, #8A1538 0%, #6B1029 100%);
      padding:44px 36px;color:#fff;text-align:center;margin-bottom:18px;
      border-bottom:4px solid #C9A646;box-shadow:0 6px 20px rgba(138,21,56,.25);position:relative;
      direction: rtl;
    }
    .header-container::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;
      background:linear-gradient(90deg, #C9A646 0%, #E8D4A0 50%, #C9A646 100%)}
    .header-container h1{margin:0 0 6px 0;font-size:32px;font-weight:800}
    .header-container .subtitle{font-size:15px;font-weight:700;margin:0 0 4px}
    .header-container .accent-line{font-size:12px;color:#C9A646;font-weight:700;margin:0 0 6px}
    .header-container .description{font-size:12px;opacity:.95;margin:0}

    [data-testid="stSidebar"]{
      background:linear-gradient(180deg, #8A1538 0%, #6B1029 100%)!important;
      border-left:2px solid #C9A646;box-shadow:-4px 0 16px rgba(0,0,0,.15);
      direction: rtl;
    }
    [data-testid="stSidebar"] *{ color:#fff !important; }
    [data-testid="stSidebar"] > div:first-child { direction: rtl; }

    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] select {
      color:#000 !important; background:#fff !important; caret-color:#000 !important;
      text-align: right;
    }
    [data-testid="stSidebar"] div[role="combobox"] input{ color:#000 !important; background:#fff !important; }
    [data-testid="stSidebar"] .stTextInput input,
    [data-testid="stSidebar"] .stNumberInput input{ color:#000 !important; background:#fff !important; text-align: right; }
    [data-testid="stSidebar"] ::placeholder{ color:#444 !important; opacity:1 !important; }
    [data-testid="stSidebar"] .stTextInput > div > div,
    [data-testid="stSidebar"] .stNumberInput > div > div{ border:1px solid rgba(0,0,0,.2) !important; box-shadow:none !important; }

    .chart-container{background:#fff;border:2px solid #E5E7EB;border-right:5px solid #8A1538;
      border-radius:12px;padding:16px;margin:12px 0;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .chart-title{font-size:20px;font-weight:800;color:#8A1538;text-align:center;margin-bottom:10px}

    .footer{margin-top:22px;background:linear-gradient(135deg, #8A1538 0%, #6B1029 100%);
      color:#fff;border-radius:10px;padding:12px 10px;text-align:center;box-shadow:0 6px 18px rgba(138,21,56,.20);position:relative}
    .footer .line{width:100%;height:3px;background:linear-gradient(90deg, #C9A646 0%, #E8D4A0 50%, #C9A646 100%);
      position:absolute;top:0;left:0}
    .footer .school{font-weight:800;font-size:15px;margin:2px 0 4px}
    .footer .rights{font-weight:700;font-size:12px;margin:0 0 4px;opacity:.95}
    .footer .contact{font-size:12px;margin-top:2px}
    .footer a{color:#E8D4A0;font-weight:700;text-decoration:none;border-bottom:1px solid #C9A646}
    .footer .credit{margin-top:6px;font-size:11px;opacity:.85}
    
    .stRadio > div { direction: rtl; justify-content: flex-end; }
    .stCheckbox > label { direction: rtl; }
    .stSelectbox > div > div { direction: rtl; text-align: right; }
    
    /* إضافة: مؤشر التحميل */
    .stProgress > div > div { background: #8A1538 !important; }
    </style>
    """, unsafe_allow_html=True)

def render_header(title: str):
    """عرض الهيدر الرئيسي"""
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
        <h1>{title}</h1>
      </div>
      <p class='subtitle'>لوحة مهنية لقياس التقدم وتحليل النتائج - النسخة المحسّنة 2.0</p>
      <p class='accent-line'>هوية إنجاز • دعم العربية الكامل • أداء محسّن</p>
      <p class='description'>المنطق الذكي: الشرطة = غير مستحق | M = متبقي | القيمة = منجز</p>
    </div>
    """, unsafe_allow_html=True)

# ============== دوال معالجة النصوص العربية ==============

def rtl(text: str) -> str:
    """تحويل النص للعرض من اليمين لليسار"""
    if not isinstance(text, str):
        text = str(text)
    if AR_OK:
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception as e:
            logger.warning(f"خطأ في معالجة RTL: {e}")
            return text
    return text

# ============== دوال معالجة الملفات ==============

@safe_execute(default_return=("", None), error_message="خطأ في إعداد الخط")
def prepare_default_font() -> Tuple[str, Optional[str]]:
    """إعداد الخط الافتراضي لـ PDF"""
    font_name = "ARFont"
    
    # قائمة الخطوط المحتملة
    font_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
    ]
    
    for candidate in font_candidates:
        if os.path.exists(candidate):
            logger.info(f"✅ تم العثور على خط: {candidate}")
            return font_name, candidate
    
    logger.warning("⚠️ لم يتم العثور على خط مناسب")
    return "", None

@safe_execute(default_return=None, error_message="خطأ في معالجة الشعار")
def prepare_logo_file(logo_file) -> Optional[str]:
    """حفظ شعار المدرسة مؤقتاً"""
    if logo_file is None:
        return None
    
    # التحقق من نوع الملف
    ext = os.path.splitext(logo_file.name)[1].lower()
    if ext not in [".png", ".jpg", ".jpeg"]:
        st.warning("⚠️ يرجى رفع شعار بصيغة PNG أو JPG")
        return None
    
    # حفظ الملف
    path = f"/tmp/school_logo{ext}"
    logo_file.seek(0)  # ✅ إعادة تعيين مؤشر الملف
    
    with open(path, "wb") as f:
        f.write(logo_file.read())
    
    logger.info(f"✅ تم حفظ الشعار: {path}")
    return path

# ============== دوال تحليل Excel ==============

def parse_sheet_name(sheet_name: str) -> Tuple[str, str, str]:
    """استخراج المادة والصف والشعبة من اسم الورقة"""
    try:
        parts = sheet_name.strip().split()
        if len(parts) < 3:
            return sheet_name.strip(), "", ""
        
        section = parts[-1]
        level = parts[-2]
        subject = " ".join(parts[:-2])
        
        # التحقق من صحة الصف
        if not (level.isdigit() or (level.startswith('0') and len(level) <= 2)):
            subject = " ".join(parts[:-1])
            level = parts[-1]
            section = ""
        
        return subject, level, section
    except Exception as e:
        logger.warning(f"خطأ في معالجة اسم الورقة '{sheet_name}': {e}")
        return sheet_name, "", ""

@st.cache_data(ttl=3600, max_entries=10, show_spinner=False)
@log_performance
def analyze_excel_file(
    file, 
    sheet_name: str, 
    due_start: Optional[date] = None, 
    due_end: Optional[date] = None
) -> List[Dict[str, Any]]:
    """
    تحليل محسّن لورقة Excel مع معالجة أخطاء أفضل
    ✅ حل مشكلة تكرار الطلاب في نفس المادة
    
    Returns:
        List[Dict]: قائمة ببيانات الطلاب
    """
    try:
        # قراءة الملف
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # التحقق من البنية
        is_valid, error_msg = validate_excel_structure(df, sheet_name)
        if not is_valid:
            st.error(f"❌ الورقة '{sheet_name}': {error_msg}")
            return []
        
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        
        # إعداد الفلترة
        filter_active = (due_start is not None and due_end is not None)
        if filter_active and due_start > due_end:
            due_start, due_end = due_end, due_start
        
        # استخراج أعمدة التقييم
        assessment_columns = []
        skipped_reasons = []
        columns_without_dates = 0
        
        for c in range(7, df.shape[1]):
            # العنوان
            title = df.iloc[0, c] if c < df.shape[1] else None
            if pd.isna(title):
                break
            
            t = str(title).strip()
            
            # تجاهل الأعمدة بعنوان فارغ فقط
            if not t or t in ['_', 'Unnamed']:
                skipped_reasons.append(f"عمود {c+1} - عنوان فارغ")
                continue
            
            # معالجة التاريخ
            due_dt = None
            if filter_active:
                due_cell = df.iloc[2, c] if 2 < df.shape[0] and c < df.shape[1] else None
                due_dt = parse_due_date_cell(due_cell, default_year=date.today().year)
                
                if due_dt is None:
                    columns_without_dates += 1
                else:
                    if not in_range(due_dt, due_start, due_end):
                        skipped_reasons.append(f"'{t}' - خارج النطاق ({due_dt})")
                        continue
            
            # ✅ التحقق المحسّن: العمود صالح حتى لو كان كله "-"
            # نتحقق فقط من أن العمود ليس فارغاً تماماً (NaN فقط)
            has_any_value = False
            for r in range(4, min(len(df), 50)):
                if r >= df.shape[0] or c >= df.shape[1]:
                    break
                val = df.iloc[r, c]
                if pd.notna(val):  # ✅ أي قيمة (حتى "-") تعتبر صالحة
                    has_any_value = True
                    break
            
            if not has_any_value:
                skipped_reasons.append(f"'{t}' - عمود فارغ تماماً (NaN)")
                continue
            
            # ✅ نضيف العمود حتى لو كان كله "-"
            assessment_columns.append({
                'index': c,
                'title': t,
                'due_date': due_dt,
                'has_date': due_dt is not None
            })
        
        # رسائل المعلومات
        if not assessment_columns:
            st.warning(f"⚠️ الورقة '{sheet_name}': لم يتم العثور على أعمدة تقييم صالحة")
            if skipped_reasons:
                with st.expander(f"📋 الأعمدة المتجاهلة ({len(skipped_reasons)})"):
                    for reason in skipped_reasons[:15]:
                        st.text(f"  • {reason}")
            return []
        
        cols_with_dates = sum(1 for c in assessment_columns if c['has_date'])
        
        info_msg = f"✅ الورقة '{sheet_name}': وُجد {len(assessment_columns)} عمود تقييم"
        if filter_active:
            info_msg += f" ({cols_with_dates} ضمن النطاق"
            if columns_without_dates > 0:
                info_msg += f"، {columns_without_dates} بدون تاريخ"
            info_msg += ")"
        
        st.success(info_msg)
        
        # ✅ عرض التواريخ المكتشفة للتأكد
        if filter_active and cols_with_dates > 0:
            with st.expander(f"📅 التواريخ المكتشفة في '{sheet_name}'"):
                for col in assessment_columns[:10]:  # عرض أول 10
                    if col['has_date']:
                        st.text(f"  ✅ {col['title']}: {col['due_date']}")
                    else:
                        st.text(f"  ⚠️ {col['title']}: لا يوجد تاريخ")
        
        if skipped_reasons and len(skipped_reasons) > 0:
            with st.expander(f"ℹ️ تم تجاهل {len(skipped_reasons)} عمود"):
                for reason in skipped_reasons[:10]:
                    st.text(f"  • {reason}")
        
        # ✅ تحليل بيانات الطلاب مع دمج الصفوف المكررة
        student_data_dict = {}  # {student_name: {total, done, pending}}
        NOT_DUE = {'-', '—', '–', '', 'NAN', 'NONE'}
        
        students_count = 0
        rows_processed = 0
        
        for r in range(4, len(df)):
            student = df.iloc[r, 0]
            if pd.isna(student) or str(student).strip() == "":
                continue
            
            # تنظيف اسم الطالب
            name = " ".join(str(student).strip().split())
            rows_processed += 1
            
            # ✅ إنشاء مفتاح فريد للطالب (الاسم فقط، لأن المادة واحدة)
            if name not in student_data_dict:
                student_data_dict[name] = {
                    'total': 0,
                    'done': 0,
                    'pending': []
                }
                students_count += 1
            
            # معالجة التقييمات لهذا الصف
            for col in assessment_columns:
                c = col['index']
                title = col['title']
                
                if c >= df.shape[1]:
                    continue
                
                raw = df.iloc[r, c]
                s = "" if pd.isna(raw) else str(raw).strip().upper()
                
                # الشرطة = غير مستحق (لا يُحسب)
                if s in NOT_DUE:
                    continue
                
                # M = مستحق غير منجز
                if s == 'M':
                    student_data_dict[name]['total'] += 1
                    if title not in student_data_dict[name]['pending']:
                        student_data_dict[name]['pending'].append(title)
                    continue
                
                # أي قيمة أخرى = منجز
                student_data_dict[name]['total'] += 1
                student_data_dict[name]['done'] += 1
        
        # ✅ تحويل القاموس إلى قائمة نتائج
        results = []
        for name, data in student_data_dict.items():
            total = data['total']
            done = data['done']
            pending = data['pending']
            
            if total == 0:
                continue
            
            pct = (done / total * 100) if total > 0 else 0.0
            
            results.append({
                "student_name": name,
                "subject": subject,
                "level": str(level_from_name).strip(),
                "section": str(section_from_name).strip(),
                "solve_pct": round(pct, 1),
                "completed_count": int(done),
                "total_count": int(total),
                "pending_titles": ", ".join(pending) if pending else "-",
                "sheet_name": sheet_name
            })
        
        # رسائل المعلومات النهائية
        if results:
            if rows_processed > students_count:
                st.info(
                    f"📊 تم معالجة {rows_processed} صف، "
                    f"دُمجت البيانات إلى {students_count} طالب فريد"
                )
            else:
                st.info(f"📊 تم تحليل {len(results)} طالب")
        else:
            st.warning(f"⚠️ الورقة '{sheet_name}': لم يتم العثور على طلاب بتقييمات مستحقة")
        
        return results
    
    except Exception as e:
        st.error(f"❌ خطأ في تحليل الملف '{sheet_name}': {e}")
        import traceback
        with st.expander("🔍 تفاصيل الخطأ التقنية"):
            st.code(traceback.format_exc())
        return []

# ============== دوال الجداول المحورية ==============

def categorize_performance(percent: float) -> str:
    """تصنيف الأداء بناءً على النسبة - نسخة محسّنة"""
    if pd.isna(percent) or percent == 0:
        return 'بحاجة لتحسين'
    elif percent >= 90:
        return 'بلاتيني 🥇'
    elif percent >= 80:
        return 'ذهبي 🥈'
    elif percent >= 70:
        return 'فضي 🥉'
    elif percent >= 60:
        return 'برونزي'
    else:
        return 'بحاجة لتحسين'

def categorize_vectorized(series: pd.Series) -> pd.Series:
    """تصنيف سريع باستخدام numpy - vectorized"""
    conditions = [
        series >= 90,
        (series >= 80) & (series < 90),
        (series >= 70) & (series < 80),
        (series >= 60) & (series < 70),
        series < 60
    ]
    
    choices = [
        'بلاتيني 🥇',
        'ذهبي 🥈',
        'فضي 🥉',
        'برونزي',
        'بحاجة لتحسين'
    ]
    
    return pd.Series(
        np.select(conditions, choices, default='بحاجة لتحسين'),
        index=series.index
    )

@st.cache_data(show_spinner=False)
@log_performance
def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    إنشاء جدول محوري محسّن من بيانات التحليل
    ✅ حل مشكلة التكرار نهائياً
    """
    try:
        if df is None or df.empty:
            logger.warning("لا توجد بيانات لإنشاء جدول محوري")
            return pd.DataFrame()
        
        logger.info(f"🔄 بدء معالجة {len(df)} سجل")
        
        # ✅ خطوة 1: حذف التكرارات الأولية
        initial_count = len(df)
        dfc = df.drop_duplicates(
            subset=['student_name', 'level', 'section', 'subject'],
            keep='last'
        )
        
        if len(dfc) < initial_count:
            logger.info(f"🧹 تم حذف {initial_count - len(dfc)} سجل مكرر (نفس الطالب + المادة)")
        
        # ✅ خطوة 2: حذف تكرار الأوراق
        if 'sheet_name' in dfc.columns:
            before = len(dfc)
            dfc = dfc.drop_duplicates(
                subset=['student_name', 'level', 'section', 'sheet_name'],
                keep='last'
            )
            if len(dfc) < before:
                logger.info(f"🧹 تم حذف {before - len(dfc)} سجل مكرر (نفس الورقة)")
        
        # ✅ خطوة 3: استخراج الطلاب الفريدين
        unique_students = dfc[['student_name', 'level', 'section']].drop_duplicates()
        unique_students = unique_students.sort_values(
            ['level', 'section', 'student_name']
        ).reset_index(drop=True)
        
        st.info(f"👥 عدد الطلاب الفريدين: {len(unique_students)}")
        
        result = unique_students.copy()
        
        # ✅ خطوة 4: معالجة كل مادة على حدة
        subjects = sorted(dfc['subject'].dropna().unique())
        st.info(f"📚 المواد المكتشفة ({len(subjects)}): {', '.join(subjects)}")
        
        for subject in subjects:
            # بيانات المادة
            subject_data = dfc[dfc['subject'] == subject].copy()
            
            # حذف التكرار داخل المادة
            subject_data = subject_data.drop_duplicates(
                subset=['student_name', 'level', 'section'],
                keep='last'
            )
            
            # تأكد من القيم الرقمية
            numeric_cols = ['total_count', 'completed_count', 'solve_pct']
            for col in numeric_cols:
                if col in subject_data.columns:
                    subject_data[col] = pd.to_numeric(subject_data[col], errors='coerce').fillna(0)
            
            # إنشاء أعمدة المادة
            subject_cols = subject_data[[
                'student_name', 'level', 'section',
                'total_count', 'completed_count', 'solve_pct'
            ]].copy()
            
            subject_cols = subject_cols.rename(columns={
                'total_count': f'{subject} - إجمالي',
                'completed_count': f'{subject} - منجز',
                'solve_pct': f'{subject} - النسبة'
            })
            
            # دمج البيانات
            result = result.merge(
                subject_cols,
                on=['student_name', 'level', 'section'],
                how='left'
            )
            
            # إضافة عمود المتبقي
            pending_data = subject_data[[
                'student_name', 'level', 'section', 'pending_titles'
            ]].copy()
            pending_data = pending_data.rename(columns={
                'pending_titles': f'{subject} - متبقي'
            })
            
            result = result.merge(
                pending_data,
                on=['student_name', 'level', 'section'],
                how='left'
            )
        
        # ✅ خطوة 5: حساب المتوسطات
        pct_cols = [c for c in result.columns if 'النسبة' in c]
        
        if pct_cols:
            # حساب متوسط فقط للقيم الموجودة
            def calc_average(row):
                values = row[pct_cols].replace(0, np.nan).dropna()
                return values.mean() if len(values) > 0 else 0
            
            result['المتوسط'] = result.apply(calc_average, axis=1)
            
            # التصنيف السريع باستخدام vectorization
            result['الفئة'] = categorize_vectorized(result['المتوسط'])
        
        # ✅ خطوة 6: إعادة تسمية الأعمدة
        result = result.rename(columns={
            'student_name': 'الطالب',
            'level': 'الصف',
            'section': 'الشعبة'
        })
        
        # ✅ خطوة 7: تنظيف القيم
        for c in result.columns:
            if ('إجمالي' in c) or ('منجز' in c):
                result[c] = result[c].fillna(0).astype(int)
            elif ('النسبة' in c) or (c == 'المتوسط'):
                result[c] = result[c].fillna(0).round(1)
            elif 'متبقي' in c:
                result[c] = result[c].fillna('-')
        
        # ✅ خطوة 8: التحقق النهائي من التكرار
        before_final = len(result)
        result = result.drop_duplicates(
            subset=['الطالب', 'الصف', 'الشعبة'],
            keep='first'
        ).reset_index(drop=True)
        
        if before_final != len(result):
            logger.warning(
                f"⚠️ تم حذف {before_final - len(result)} صف مكرر في المرحلة النهائية"
            )
        
        logger.info(f"✅ الجدول النهائي: {len(result)} طالب × {len(result.columns)} عمود")
        
        return result
    
    except Exception as e:
        logger.error(f"❌ خطأ في معالجة البيانات: {e}")
        import traceback
        with st.expander("🔍 تفاصيل الخطأ"):
            st.code(traceback.format_exc())
        return pd.DataFrame()

# ============== دوال المعالجة والتجميع ==============

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """تطبيع البيانات للرسوم البيانية"""
    out = df.copy()
    
    # إعادة تسمية
    if 'solve_pct' in out.columns:
        out = out.rename(columns={'solve_pct': 'percent'})
    if 'student_name' in out.columns:
        out = out.rename(columns={'student_name': 'student'})
    
    # التصنيف
    if 'percent' in out.columns:
        out['category'] = categorize_vectorized(out['percent'])
    
    return out

def aggregate_by_subject(df: pd.DataFrame) -> pd.DataFrame:
    """تجميع البيانات حسب المادة والفئة"""
    if df.empty:
        return pd.DataFrame()
    
    rows = []
    
    for s in df['subject'].dropna().unique():
        sub = df[df['subject'] == s]
        n = len(sub)
        
        # حساب المتوسط
        avg = 0.0
        if 'percent' in sub.columns:
            avg = sub['percent'].mean() if n > 0 else 0.0
            if pd.isna(avg):
                avg = 0.0
        
        # حساب توزيع الفئات
        for cat in CATEGORY_ORDER:
            count = (sub['category'] == cat).sum() if 'category' in sub.columns else 0
            pct = (count / n * 100) if n > 0 else 0.0
            
            rows.append({
                'subject': s,
                'category': cat,
                'count': int(count),
                'percent_share': round(pct, 1),
                'avg_completion': round(avg, 1)
            })
    
    agg = pd.DataFrame(rows)
    
    if agg.empty:
        return agg
    
    # ترتيب المواد حسب متوسط الإنجاز
    order = agg.groupby('subject')['avg_completion'].first().sort_values(
        ascending=False
    ).index.tolist()
    
    agg['subject'] = pd.Categorical(agg['subject'], categories=order, ordered=True)
    return agg.sort_values('subject')

# ============== دوال الرسوم البيانية ==============

def chart_stacked_by_subject(agg_df: pd.DataFrame, mode='percent') -> go.Figure:
    """رسم بياني مكدس حسب المادة"""
    fig = go.Figure()
    
    colors = [CATEGORY_COLORS[c] for c in CATEGORY_ORDER]
    
    for i, cat in enumerate(CATEGORY_ORDER):
        d = agg_df[agg_df['category'] == cat]
        
        vals = d['percent_share'] if mode == 'percent' else d['count']
        
        # النصوص على الأعمدة
        text = [
            (f"{v:.1f}%" if mode == 'percent' else str(int(v))) if v > 0 else ""
            for v in vals
        ]
        
        # Hover info
        hover = (
            "<b>%{y}</b><br>الفئة: " + cat + "<br>" +
            ("النسبة: %{x:.1f}%<extra></extra>" if mode == 'percent' else "العدد: %{x}<extra></extra>")
        )
        
        fig.add_trace(go.Bar(
            name=cat,
            x=vals,
            y=d['subject'],
            orientation='h',
            marker=dict(
                color=colors[i],
                line=dict(color='white', width=1)
            ),
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
        font=dict(family='Cairo'),
        height=max(400, len(agg_df['subject'].unique()) * 40)
    )
    
    return fig

def chart_overall_donut(pivot: pd.DataFrame) -> go.Figure:
    """رسم دائري للتوزيع العام"""
    if pivot.empty or 'الفئة' not in pivot.columns:
        return go.Figure()
    
    counts = pivot['الفئة'].value_counts().reindex(CATEGORY_ORDER, fill_value=0)
    
    fig = go.Figure([
        go.Pie(
            labels=counts.index,
            values=counts.values,
            hole=0.55,
            marker=dict(colors=[CATEGORY_COLORS[k] for k in counts.index]),
            textinfo='label+value',
            textfont=dict(size=13, family='Cairo'),
            hovertemplate="%{label}: %{value} طالب<extra></extra>"
        )
    ])
    
    fig.update_layout(
        title=dict(
            text="توزيع عام للفئات",
            font=dict(size=20, family='Cairo', color='#8A1538'),
            x=0.5
        ),
        showlegend=False,
        font=dict(family='Cairo'),
        height=400
    )
    
    return fig

def chart_overall_gauge(pivot: pd.DataFrame) -> go.Figure:
    """مؤشر قياس متوسط الإنجاز العام"""
    avg = 0.0
    
    if not pivot.empty and 'المتوسط' in pivot.columns:
        mean_val = pivot['المتوسط'].mean()
        avg = float(mean_val) if pd.notna(mean_val) else 0.0
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=avg,
        number={'suffix': "%", 'font': {'family': 'Cairo', 'size': 40}},
        gauge={
            'axis': {'range': [0, 100], 'tickfont': {'family': 'Cairo'}},
            'bar': {'color': '#8A1538'},
            'steps': [
                {'range': [0, 60], 'color': '#ffebee'},
                {'range': [60, 70], 'color': '#fff3e0'},
                {'range': [70, 80], 'color': '#f1f8e9'},
                {'range': [80, 90], 'color': '#e8f5e9'},
                {'range': [90, 100], 'color': '#e0f7fa'}
            ],
            'threshold': {
                'line': {'color': CATEGORY_COLORS['ذهبي 🥈'], 'width': 4},
                'thickness': 0.75,
                'value': 80
            }
        }
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
        height=350
    )
    
    return fig

# ============== دوال PDF ==============

@safe_execute(default_return=b"", error_message="خطأ في إنشاء PDF")
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
    """إنشاء تقرير PDF للطالب - نسخة محسّنة"""
    
    font_name, font_path = font_info
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    # إضافة الخط
    if font_path and os.path.exists(font_path):
        try:
            pdf.add_font(font_name, "", font_path, uni=True)
        except Exception as e:
            logger.warning(f"فشل في إضافة الخط: {e}")
            font_name = ""
    
    def set_font(size=12, color=(0, 0, 0)):
        if font_name:
            pdf.set_font(font_name, size=size)
        else:
            pdf.set_font("Helvetica", size=size)
        pdf.set_text_color(*color)
    
    # الهيدر
    pdf.set_fill_color(*QATAR_MAROON)
    pdf.rect(0, 0, 210, 20, style="F")
    
    # الشعار
    if logo_path and os.path.exists(logo_path):
        try:
            pdf.image(logo_path, x=185, y=2.5, w=20)
        except Exception as e:
            logger.warning(f"فشل في إضافة الشعار: {e}")
    
    set_font(14, (255, 255, 255))
    pdf.set_xy(10, 7)
    pdf.cell(0, 8, rtl("إنجاز - تقرير أداء الطالب"), align="R")
    
    # العنوان
    set_font(18, QATAR_MAROON)
    pdf.set_y(28)
    pdf.cell(0, 10, rtl("تقرير أداء الطالب - نظام قطر للتعليم"), ln=1, align="R")
    pdf.set_draw_color(*QATAR_GOLD)
    pdf.set_line_width(0.6)
    pdf.line(30, 38, 200, 38)
    
    # بيانات الطالب
    set_font(12, (0, 0, 0))
    pdf.ln(6)
    pdf.cell(0, 8, rtl(f"اسم المدرسة: {school_name or '—'}"), ln=1, align="R")
    pdf.cell(0, 8, rtl(f"اسم الطالب: {student_name}"), ln=1, align="R")
    pdf.cell(0, 8, rtl(f"الصف: {grade or '—'}     الشعبة: {section or '—'}"), ln=1, align="R")
    pdf.ln(2)
    
    # جدول المواد
    headers = [
        rtl("المادة"),
        rtl("عدد التقييمات الإجمالي"),
        rtl("عدد التقييمات المنجزة"),
        rtl("عدد التقييمات المتبقية")
    ]
    widths = [70, 45, 45, 40]
    
    pdf.set_fill_color(*QATAR_MAROON)
    set_font(12, (255, 255, 255))
    pdf.set_y(pdf.get_y() + 4)
    
    for w, h in zip(widths, headers):
        pdf.cell(w, 9, h, border=0, align="C", fill=True)
    pdf.ln(9)
    
    # بيانات الجدول
    set_font(11, (0, 0, 0))
    total_total = 0
    total_solved = 0
    
    for _, r in table_df.iterrows():
        sub = rtl(str(r['المادة']))
        tot = int(r['إجمالي'])
        solv = int(r['منجز'])
        rem = int(max(tot - solv, 0))
        
        total_total += tot
        total_solved += solv
        
        pdf.set_fill_color(247, 247, 247)
        pdf.cell(widths[0], 8, sub, 0, 0, "C", True)
        pdf.cell(widths[1], 8, str(tot), 0, 0, "C", True)
        pdf.cell(widths[2], 8, str(solv), 0, 0, "C", True)
        pdf.cell(widths[3], 8, str(rem), 0, 1, "C", True)
    
    # الإحصائيات
    pdf.ln(3)
    set_font(12, QATAR_MAROON)
    pdf.cell(0, 8, rtl("الإحصائيات"), ln=1, align="R")
    
    set_font(12, (0, 0, 0))
    remaining = max(total_total - total_solved, 0)
    pdf.cell(
        0, 8,
        rtl(f"منجز: {total_solved}    متبقي: {remaining}    نسبة حل التقييمات: {overall_avg:.1f}%"),
        ln=1, align="R"
    )
    
    # التوصية
    pdf.ln(2)
    set_font(12, QATAR_MAROON)
    pdf.cell(0, 8, rtl("توصية منسق المشاريع:"), ln=1, align="R")
    
    set_font(11, (0, 0, 0))
    reco_lines = (reco_text or "—").splitlines() if reco_text else ["—"]
    for line in reco_lines:
        pdf.multi_cell(0, 7, rtl(line), align="R")
    
    # الروابط
    pdf.ln(2)
    set_font(12, QATAR_MAROON)
    pdf.cell(0, 8, rtl("روابط مهمة:"), ln=1, align="R")
    
    set_font(11, (0, 0, 0))
    pdf.cell(0, 7, rtl("رابط نظام قطر: https://portal.education.qa"), ln=1, align="R")
    pdf.cell(0, 7, rtl("استعادة كلمة المرور: https://password.education.qa"), ln=1, align="R")
    pdf.cell(0, 7, rtl("قناة قطر للتعليم: https://edu.tv.qa"), ln=1, align="R")
    
    # التوقيعات
    pdf.ln(4)
    set_font(12, QATAR_MAROON)
    pdf.cell(0, 8, rtl("التوقيعات"), ln=1, align="R")
    
    set_font(11, (0, 0, 0))
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
        set_font(11, (0, 0, 0))
        pdf.set_xy(x, yb + 3)
        pdf.cell(w - 4, 6, rtl(f"{title} / {name or '—'}"), align="R")
        pdf.set_xy(x, yb + 10)
        pdf.cell(w - 4, 6, rtl("التوقيع: __________________    التاريخ: __________"), align="R")
    
    # إخراج PDF
    try:
        out = pdf.output(dest="S")
        if isinstance(out, bytes):
            return out
        elif isinstance(out, str):
            return out.encode("latin-1")
        else:
            # محاولة تحويل أي نوع آخر
            return bytes(out)
    except Exception as e:
        logger.error(f"خطأ في إخراج PDF: {e}")
        raise

# ============== التطبيق الرئيسي ==============

def main():
    """الدالة الرئيسية للتطبيق"""
    
    # إعداد التطبيق
    setup_app()
    
    # الشريط الجانبي
    with st.sidebar:
        st.image("https://i.imgur.com/XLef7tS.png", width=110)
        st.markdown("---")
        
        st.header("⚙️ الإعدادات")
        
        # 1. تحميل الملفات
        st.subheader("📁 تحميل الملفات")
        uploaded_files = st.file_uploader(
            "اختر ملفات Excel",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="يمكنك رفع عدة ملفات في آن واحد"
        )
        
        # معالجة الأوراق
        selected_sheets = []
        all_sheets = []
        sheet_file_map = {}
        
        if uploaded_files:
            for file_idx, file in enumerate(uploaded_files):
                try:
                    file.seek(0)  # ✅ إعادة تعيين المؤشر
                    xls = pd.ExcelFile(file)
                    for sheet in xls.sheet_names:
                        label = f"[ملف {file_idx + 1}] {sheet}"
                        all_sheets.append(label)
                        sheet_file_map[label] = (file, sheet)
                except Exception as e:
                    st.error(f"❌ خطأ في قراءة الملف: {e}")
            
            if all_sheets:
                st.info(f"📋 وُجدت {len(all_sheets)} ورقة في {len(uploaded_files)} ملف")
                
                select_all = st.checkbox(
                    "✔️ اختر الجميع",
                    value=True,
                    key="select_all_sheets"
                )
                
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
        
        # 2. فلترة التواريخ
        st.subheader("⏳ فلترة الأعمدة حسب تاريخ الاستحقاق")
        enable_date_filter = st.checkbox(
            "تفعيل فلتر التاريخ",
            value=False,
            help="يقرأ التاريخ من H3 لكل عمود. الأعمدة خارج النطاق الزمني يتم تجاهلها بالكامل.",
            key="enable_date_filter"
        )
        
        if enable_date_filter:
            default_start = date.today().replace(day=1)
            default_end = date.today()
            
            st.info("ℹ️ سيتم تحليل الأعمدة التي تواريخها (H3) ضمن النطاق فقط")
            
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
        else:
            due_start, due_end = None, None
            st.success(
                "✅ **المنطق الذكي مفعّل:**\n"
                "- الخلية `-` أو فارغة = تقييم غير مستحق (لا يُحسب)\n"
                "- الخلية `M` = تقييم مستحق غير منجز (يُحسب متبقي)\n"
                "- الخلية بها قيمة = تقييم منجز (يُحسب منجز)"
            )
        
        # 3. شعار المدرسة
        st.subheader("🖼️ شعار المدرسة (اختياري)")
        logo_file = st.file_uploader(
            "ارفع شعار PNG/JPG",
            type=["png", "jpg", "jpeg"],
            key="logo_file"
        )
        st.session_state.logo_path = prepare_logo_file(logo_file)
        
        st.markdown("---")
        
        # 4. معلومات المدرسة
        st.subheader("🏫 معلومات المدرسة")
        school_name = st.text_input("اسم المدرسة", placeholder="مدرسة قطر النموذجية")
        
        # 5. التوقيعات
        st.subheader("✍️ التوقيعات")
        coordinator_name = st.text_input("منسق/ة المشاريع")
        academic_deputy = st.text_input("النائب الأكاديمي")
        admin_deputy = st.text_input("النائب الإداري")
        principal_name = st.text_input("مدير/ة المدرسة")
        
        st.markdown("---")
        
        # زر التشغيل
        run_analysis = st.button(
            "▶️ تشغيل التحليل",
            use_container_width=True,
            type="primary",
            disabled=not uploaded_files
        )
    
    # المحتوى الرئيسي
    if not uploaded_files:
        st.info("📤 من الشريط الجانبي ارفع ملفات Excel للبدء في التحليل")
        
    elif run_analysis:
        sheets_to_use = st.session_state.selected_sheets
        
        if not sheets_to_use:
            tmp = []
            for file in uploaded_files:
                try:
                    file.seek(0)
                    xls = pd.ExcelFile(file)
                    for sheet in xls.sheet_names:
                        tmp.append((file, sheet))
                except Exception as e:
                    st.error(f"❌ خطأ في قراءة الملف: {e}")
            sheets_to_use = tmp
        
        if not sheets_to_use:
            st.warning("⚠️ لم يتم العثور على أوراق داخل الملفات المرفوعة.")
        else:
            # شريط التقدم
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            with st.spinner("⏳ جاري التحليل..."):
                rows = []
                total_sheets = len(sheets_to_use)
                
                for idx, (file, sheet) in enumerate(sheets_to_use):
                    # تحديث شريط التقدم
                    progress = (idx + 1) / total_sheets
                    progress_bar.progress(progress)
                    status_text.text(f"📊 تحليل الورقة {idx + 1} من {total_sheets}: {sheet}")
                    
                    # إعادة تعيين مؤشر الملف
                    file.seek(0)
                    
                    # التحليل
                    sheet_results = analyze_excel_file(file, sheet, due_start, due_end)
                    rows.extend(sheet_results)
                
                progress_bar.empty()
                status_text.empty()
                
                if rows:
                    df = pd.DataFrame(rows)
                    st.session_state.analysis_results = df
                    st.session_state.pivot_table = create_pivot_table(df)
                    
                    # إحصائيات
                    subjects_count = df['subject'].nunique() if 'subject' in df.columns else 0
                    students_count = len(st.session_state.pivot_table)
                    
                    st.success(
                        f"✅ تم تحليل {students_count} طالب عبر {subjects_count} مادة بنجاح!"
                    )
                    
                    # حفظ الإحصائيات
                    st.session_state.analysis_stats = {
                        'students': students_count,
                        'subjects': subjects_count,
                        'total_assessments': df['total_count'].sum() if 'total_count' in df.columns else 0,
                        'completed': df['completed_count'].sum() if 'completed_count' in df.columns else 0,
                    }
                else:
                    st.warning(
                        "⚠️ لم يتم استخراج بيانات من الأوراق المحددة. "
                        "تأكد من تنسيق الجداول وتواريخ الاستحقاق."
                    )
    
    # عرض النتائج
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results
    
    if pivot is not None and not pivot.empty and df is not None:
        # ملخص النتائج
        st.subheader("📈 ملخص النتائج")
        
        c1, c2, c3, c4, c5 = st.columns(5)
        
        with c1:
            st.metric("👥 إجمالي الطلاب", len(pivot))
        
        with c2:
            subjects = df['subject'].nunique() if 'subject' in df.columns else 0
            st.metric("📚 عدد المواد", subjects)
        
        with c3:
            avg = 0.0
            if 'المتوسط' in pivot.columns:
                mean_val = pivot['المتوسط'].mean()
                avg = float(mean_val) if pd.notna(mean_val) else 0.0
            st.metric("📊 متوسط الإنجاز", f"{avg:.1f}%")
        
        with c4:
            platinum_count = 0
            if 'الفئة' in pivot.columns:
                platinum_count = int((pivot['الفئة'] == 'بلاتيني 🥇').sum())
            st.metric("🥇 فئة بلاتيني", platinum_count)
        
        with c5:
            zero = 0
            if 'المتوسط' in pivot.columns:
                zero = int((pivot['المتوسط'] == 0).sum())
            st.metric("⚠️ بدون إنجاز", zero)
        
        st.divider()
        
        # الجدول التفصيلي
        st.subheader("📋 جدول النتائج التفصيلي")
        
        # خيارات الفلترة
        col_filter1, col_filter2 = st.columns(2)
        
        with col_filter1:
            if 'الصف' in pivot.columns:
                levels = ['الكل'] + sorted(pivot['الصف'].dropna().unique().tolist())
                selected_level = st.selectbox("فلتر حسب الصف", levels)
            else:
                selected_level = 'الكل'
        
        with col_filter2:
            if 'الفئة' in pivot.columns:
                categories = ['الكل'] + CATEGORY_ORDER
                selected_category = st.selectbox("فلتر حسب الفئة", categories)
            else:
                selected_category = 'الكل'
        
        # تطبيق الفلترة
        filtered_pivot = pivot.copy()
        
        if selected_level != 'الكل':
            filtered_pivot = filtered_pivot[filtered_pivot['الصف'] == selected_level]
        
        if selected_category != 'الكل':
            filtered_pivot = filtered_pivot[filtered_pivot['الفئة'] == selected_category]
        
        # عرض الجدول
        st.dataframe(
            filtered_pivot,
            use_container_width=True,
            height=420
        )
        
        # زر تحميل CSV
        csv = filtered_pivot.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 تحميل الجدول (CSV)",
            csv,
            f"ingaz_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "text/csv",
            key='download-csv'
        )
        
        st.divider()
        
        # الرسوم البيانية
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
            
            if not agg_df.empty:
                st.plotly_chart(
                    chart_stacked_by_subject(agg_df, mode=mode),
                    use_container_width=True
                )
            else:
                st.info("لا توجد بيانات كافية لعرض الرسم البياني")
        
        except Exception as e:
            st.error(f"خطأ في الرسم: {e}")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.divider()
        
        # التقارير الفردية
        st.subheader("📑 التقارير الفردية (PDF)")
        
        students = []
        if 'الطالب' in pivot.columns:
            students = sorted(pivot['الطالب'].dropna().astype(str).unique().tolist())
        
        if students:
            csel, crec = st.columns([2, 3])
            
            with csel:
                sel = st.selectbox("اختر الطالب", students, index=0)
                
                row = pivot[pivot['الطالب'] == sel].head(1)
                g = str(row['الصف'].iloc[0]) if not row.empty and 'الصف' in row.columns else ''
                s = str(row['الشعبة'].iloc[0]) if not row.empty and 'الشعبة' in row.columns else ''
            
            with crec:
                reco = st.text_area(
                    "توصية منسق المشاريع",
                    value="",
                    height=120,
                    placeholder="اكتب التوصيات هنا..."
                )
            
            # بيانات الطالب
            sdata = pd.DataFrame()
            if 'student_name' in df.columns:
                sdata = df[df['student_name'].str.strip().eq(sel.strip())].copy()
            
            if not sdata.empty:
                table = sdata[['subject', 'total_count', 'completed_count']].rename(
                    columns={
                        'subject': 'المادة',
                        'total_count': 'إجمالي',
                        'completed_count': 'منجز'
                    }
                )
                table['متبقي'] = (table['إجمالي'] - table['منجز']).clip(lower=0).astype(int)
                
                avg_stu = 0.0
                if 'solve_pct' in sdata.columns:
                    avg_stu = float(sdata['solve_pct'].mean())
                    if pd.isna(avg_stu):
                        avg_stu = 0.0
                
                st.markdown("### معاينة سريعة")
                st.dataframe(table, use_container_width=True, height=260)
                
                # إنشاء PDF
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
                
                if pdf_one:
                    st.download_button(
                        "📥 تحميل تقرير الطالب (PDF)",
                        pdf_one,
                        file_name=f"student_report_{sel}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            
            st.markdown("---")
            
            # تصدير جميع التقارير
            st.subheader("📦 تصدير جميع التقارير (ZIP)")
            
            same_reco = st.checkbox("استخدم نفس التوصية لكل الطلاب", value=True)
            
            if st.button("إنشاء ملف ZIP لكل التقارير", type="primary"):
                with st.spinner("جاري إنشاء حزمة التقارير..."):
                    try:
                        buf = io.BytesIO()
                        
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                            for stu in students:
                                r = pivot[pivot['الطالب'] == stu].head(1)
                                g = str(r['الصف'].iloc[0]) if not r.empty and 'الصف' in r.columns else ''
                                s = str(r['الشعبة'].iloc[0]) if not r.empty and 'الشعبة' in r.columns else ''
                                
                                sd = pd.DataFrame()
                                if 'student_name' in df.columns:
                                    sd = df[df['student_name'].str.strip().eq(stu.strip())].copy()
                                
                                if not sd.empty:
                                    t = sd[['subject', 'total_count', 'completed_count']].rename(
                                        columns={
                                            'subject': 'المادة',
                                            'total_count': 'إجمالي',
                                            'completed_count': 'منجز'
                                        }
                                    )
                                    t['متبقي'] = (t['إجمالي'] - t['منجز']).clip(lower=0).astype(int)
                                    
                                    av = 0.0
                                    if 'solve_pct' in sd.columns:
                                        av = float(sd['solve_pct'].mean())
                                        if pd.isna(av):
                                            av = 0.0
                                    
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
                                    
                                    if pdfb:
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
                        
                        st.success(f"✅ تم إنشاء {len(students)} تقرير بنجاح!")
                    
                    except Exception as e:
                        st.error(f"❌ خطأ في إنشاء الحزمة: {e}")
    
    # الفوتر
    st.markdown(f"""
    <div class="footer">
    <div class="line"></div>
    <div class="school">مدرسة عثمان بن عفان النموذجية للبنين</div>
    <div class="rights">© {datetime.now().year} جميع الحقوق محفوظة</div>
    <div class="contact">للتواصل: <a href="mailto:S.mahgoub0101@education.qa">S.mahgoub0101@education.qa</a></div>
    <div class="credit">تطوير وتصميم: قسم التحول الرقمي | النسخة المحسّنة 2.0</div>
    </div>
    """, unsafe_allow_html=True)

# ============== تشغيل التطبيق ==============

if __name__ == "__main__":
    main()

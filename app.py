# -*- coding: utf-8 -*-
"""
تطبيق إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم
النسخة المحسّنة 2.1
"""

import os
import io
import re
import zipfile
import logging
import unicodedata
import warnings
import tempfile
from datetime import datetime, date
from typing import Tuple, Optional, List, Dict, Any
from functools import wraps
import time

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from fpdf import FPDF

# ================== دعم العربية ==================
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    AR_OK = True
except ImportError:
    AR_OK = False
    warnings.warn(
        "⚠️ arabic_reshaper غير مثبت — للتثبيت: pip install arabic-reshaper python-bidi"
    )

def rtl(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    if AR_OK:
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text

# ================== ثوابت وهوية ==================
QATAR_MAROON = (138, 21, 56)
QATAR_GOLD = (201, 166, 70)

CATEGORY_COLORS = {
    'بلاتيني 🥇': '#E5E4E2',
    'ذهبي 🥈': '#C9A646',
    'فضي 🥉': '#C0C0C0',
    'برونزي': '#CD7F32',
    'بحاجة لتحسين': '#FF9800',
    'لا يستفيد': '#8A1538'
}
CATEGORY_ORDER = ['بلاتيني 🥇', 'ذهبي 🥈', 'فضي 🥉', 'برونزي', 'بحاجة لتحسين', 'لا يستفيد']

# ================== لوجينغ ==================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("ingaz-app")

def log_performance(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        t0 = time.time()
        name = func.__name__
        logger.info(f"🔄 بدء {name}")
        try:
            res = func(*args, **kwargs)
            logger.info(f"✅ {name} تم في {time.time()-t0:.2f}s")
            return res
        except Exception as e:
            logger.error(f"❌ {name} فشل: {e}")
            raise
    return wrapper

def safe_execute(default_return=None, error_message="حدث خطأ"):
    def deco(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger.error(f"{error_message} في {func.__name__}: {e}")
                if st:
                    st.error(f"{error_message}: {e}")
                return default_return
        return wrapper
    return deco

# ================== أدوات نص/أرقام ==================
def _normalize_arabic_digits(s: str) -> str:
    if not isinstance(s, str):
        return str(s) if s is not None else ""
    return s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

def _strip_invisible_and_diacritics(s: str) -> str:
    if not isinstance(s, str):
        return ""
    invis = ['\u200e','\u200f','\u202a','\u202b','\u202c','\u202d','\u202e',
             '\u2066','\u2067','\u2068','\u2069','\u200b','\u200c','\u200d',
             '\ufeff','\xa0','\u0640']
    for ch in invis:
        s = s.replace(ch, '')
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = ' '.join(s.split())
    return s.strip()

def _norm_month_key(s: str) -> str:
    s = _strip_invisible_and_diacritics(s)
    s = _normalize_arabic_digits(s).lower().strip()
    s = s.replace("أ","ا").replace("إ","ا").replace("آ","ا").replace("ة","ه").replace("ـ","")
    return s

_AR_MONTHS = { _norm_month_key(k): v for k, v in {
    "يناير":1,"فبراير":2,"مارس":3,"ابريل":4,"أبريل":4,"مايو":5,"يونيو":6,"يوليو":7,"اغسطس":8,"أغسطس":8,"سبتمبر":9,"اكتوبر":10,"أكتوبر":10,"نوفمبر":11,"ديسمبر":12,
    "كانون الثاني":1,"شباط":2,"اذار":3,"نيسان":4,"ايار":5,"حزيران":6,"تموز":7,"اب":8,"ايلول":9,"تشرين الاول":10,"تشرين الثاني":11,"كانون الاول":12,
    "جانفي":1,"فيفري":2,"افريل":4,"ماي":5,"جوان":6,"جويلية":7,"اوت":8,"دجنبر":12
}.items() }

_EN_MONTHS = {k: v for k, v in {
    "jan":1,"january":1,"feb":2,"february":2,"mar":3,"march":3,"apr":4,"april":4,"may":5,
    "jun":6,"june":6,"jul":7,"july":7,"aug":8,"august":8,"sep":9,"sept":9,"september":9,
    "oct":10,"october":10,"nov":11,"november":11,"dec":12,"december":12
}.items() }

# ================== تواريخ ==================
def parse_due_date_cell(cell, default_year: int = None) -> Optional[date]:
    if default_year is None:
        default_year = date.today().year
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    if isinstance(cell, (pd.Timestamp, datetime)):
        try:
            return cell.date() if hasattr(cell, 'date') else cell
        except Exception:
            return None
    if isinstance(cell, (int, float)) and not pd.isna(cell):
        try:
            if 1 <= cell <= 100000:
                base = pd.to_datetime("1899-12-30")
                result = base + pd.to_timedelta(float(cell), unit="D")
                if 1900 <= result.year <= 2200:
                    return result.date()
        except Exception:
            pass
    try:
        s = str(cell).strip()
        if not s or s.lower() in ['nan','none','nat','null']:
            return None
        s_clean = _strip_invisible_and_diacritics(_normalize_arabic_digits(s))
        # EN: "Oct 2" / "2 Oct"
        m1 = re.search(r"([a-zA-Z]{3,})\s*[-/،,\s]*\s*(\d{1,2})", s_clean, re.IGNORECASE)
        if m1:
            m = _EN_MONTHS.get(m1.group(1).lower()); d = int(m1.group(2))
            if m:
                try: return date(default_year, m, d)
                except: return date(default_year, m, min(d, 28))
        m2 = re.search(r"(\d{1,2})\s*[-/،,\s]*\s*([a-zA-Z]{3,})", s_clean, re.IGNORECASE)
        if m2:
            d = int(m2.group(1)); m = _EN_MONTHS.get(m2.group(2).lower())
            if m:
                try: return date(default_year, m, d)
                except: return date(default_year, m, min(d, 28))
        # AR: "19 أكتوبر" / "أكتوبر 19"
        m3 = re.search(r"(\d{1,2})\s*[-/،,\s]*\s*([^\d\s]+)", s_clean)
        if m3:
            d = int(m3.group(1)); m = _AR_MONTHS.get(_norm_month_key(m3.group(2)))
            if m:
                try: return date(default_year, m, d)
                except: return date(default_year, m, min(d, 28))
        m4 = re.search(r"([^\d\s]+)\s*[-/،,\s]*\s*(\d{1,2})", s_clean)
        if m4:
            m = _AR_MONTHS.get(_norm_month_key(m4.group(1))); d = int(m4.group(2))
            if m:
                try: return date(default_year, m, d)
                except: return date(default_year, m, min(d, 28))
        parsed = pd.to_datetime(s_clean, dayfirst=True, errors="coerce")
        if pd.notna(parsed):
            d = parsed.date()
            if parsed.year < 1900: d = d.replace(year=default_year)
            return d
    except Exception:
        pass
    return None

def in_range(d: Optional[date], start: Optional[date], end: Optional[date]) -> bool:
    if not (start and end):
        return True
    if d is None:
        return False
    if start > end:
        start, end = end, start
    return start <= d <= end

# ================== تحقق بنية ==================
def validate_excel_structure(df: pd.DataFrame, sheet_name: str) -> Tuple[bool, str]:
    if df is None or df.empty:
        return False, "الملف فارغ"
    if df.shape[0] < 4:
        return False, f"عدد الصفوف قليل جداً ({df.shape[0]} صف)"
    if df.shape[1] < 8:
        return False, f"عدد الأعمدة قليل جداً ({df.shape[1]} عمود)"
    if len(df.iloc[4:, 0].dropna()) == 0:
        return False, "لا توجد أسماء طلاب في العمود الأول"
    return True, ""

# ================== واجهة ==================
def setup_app():
    APP_TITLE = "إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم"
    st.set_page_config(page_title=APP_TITLE, page_icon="https://i.imgur.com/XLef7tS.png",
                       layout="wide", initial_sidebar_state="expanded")
    defaults = {"analysis_results": None, "pivot_table": None, "font_info": None,
                "logo_path": None, "selected_sheets": [], "analysis_stats": {}}
    for k,v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v
    if st.session_state.font_info is None:
        st.session_state.font_info = prepare_default_font()
    apply_custom_styles(); render_header(APP_TITLE)

def apply_custom_styles():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
    *{font-family:'Cairo','Segoe UI',-apple-system,sans-serif;direction:rtl}
    .main,body,.stApp{background:#fff;direction:rtl}
    section[data-testid="stSidebar"]{right:0!important;left:auto!important}
    .main .block-container{padding-right:5rem!important;padding-left:1rem!important}
    .header-container{background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
      padding:44px 36px;color:#fff;text-align:center;margin-bottom:18px;
      border-bottom:4px solid #C9A646;box-shadow:0 6px 20px rgba(138,21,56,.25);position:relative}
    .header-container::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;
      background:linear-gradient(90deg,#C9A646 0%,#E8D4A0 50%,#C9A646 100%)}
    .header-container h1{margin:0 0 6px 0;font-size:32px;font-weight:800}
    .chart-container{background:#fff;border:2px solid #E5E7EB;border-right:5px solid #8A1538;
      border-radius:12px;padding:16px;margin:12px 0;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .chart-title{font-size:20px;font-weight:800;color:#8A1538;text-align:center;margin-bottom:10px}
    .stProgress > div > div{background:#8A1538!important}
    </style>
    """, unsafe_allow_html=True)

def render_header(title: str):
    st.markdown(f"""
    <div class='header-container'>
      <h1>{title}</h1>
      <p class='subtitle'>لوحة مهنية لقياس التقدم وتحليل النتائج - النسخة 2.1</p>
      <p class='description'>المنطق: '-' غير مستحق | M متبقي | القيمة = منجز</p>
    </div>
    """, unsafe_allow_html=True)

# ================== ملفات/خط ==================
@safe_execute(default_return=("", None), error_message="خطأ في إعداد الخط")
def prepare_default_font() -> Tuple[str, Optional[str]]:
    font_name = "ARFont"
    candidates = [
        "/usr/share/fonts/truetype/cairo/Cairo-Regular.ttf",
        "/usr/share/fonts/truetype/amiri/Amiri-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "C:\\Windows\\Fonts\\Cairo-Regular.ttf",
        "C:\\Windows\\Fonts\\Amiri-Regular.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
    ]
    for p in candidates:
        if os.path.exists(p): return font_name, p
    return "", None

@safe_execute(default_return=None, error_message="خطأ في معالجة الشعار")
def prepare_logo_file(logo_file) -> Optional[str]:
    if logo_file is None: return None
    ext = os.path.splitext(logo_file.name)[1].lower()
    if ext not in [".png",".jpg",".jpeg"]:
        st.warning("⚠️ يرجى رفع شعار PNG/JPG"); return None
    logo_file.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(logo_file.read())
        return tmp.name

# ================== تحليل Excel ==================
def parse_sheet_name(sheet_name: str) -> Tuple[str, str, str]:
    try:
        parts = sheet_name.strip().split()
        if len(parts) < 3: return sheet_name.strip(), "", ""
        section = parts[-1]
        level = _normalize_arabic_digits(parts[-2])
        subject = " ".join(parts[:-2])
        if not str(level).isdigit():
            level = _normalize_arabic_digits(parts[-1])
            subject = " ".join(parts[:-1]); section = ""
        return subject, level, section
    except Exception as e:
        logger.warning(f"parse_sheet_name: {e}")
        return sheet_name, "", ""

@log_performance
def analyze_excel_file(file, sheet_name: str,
                       due_start: Optional[date]=None, due_end: Optional[date]=None) -> List[Dict[str, Any]]:
    try:
        data = file.getvalue() if hasattr(file, "getvalue") else file.read()
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, header=None)

        ok, msg = validate_excel_structure(df, sheet_name)
        if not ok:
            st.error(f"❌ '{sheet_name}': {msg}"); return []

        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        filter_active = (due_start is not None and due_end is not None)
        if filter_active and due_start > due_end:
            due_start, due_end = due_end, due_start

        assessment_columns, skipped, no_date_cols = [], [], 0
        for c in range(7, df.shape[1]):
            title = df.iloc[0, c] if c < df.shape[1] else None
            t = "" if pd.isna(title) else str(title).strip()
            if (not t) or t.lower().startswith("unnamed"):
                skipped.append(f"عمود {c+1} - عنوان غير صالح"); continue
            due_dt = None
            if filter_active:
                due_cell = df.iloc[2, c] if 2 < df.shape[0] and c < df.shape[1] else None
                due_dt = parse_due_date_cell(due_cell, default_year=date.today().year)
                if due_dt is None:
                    no_date_cols += 1
                elif not in_range(due_dt, due_start, due_end):
                    skipped.append(f"'{t}' - خارج النطاق ({due_dt})"); continue
            has_any = False
            for r in range(4, min(len(df), 50)):
                if r >= df.shape[0] or c >= df.shape[1]: break
                if pd.notna(df.iloc[r, c]): has_any = True; break
            if not has_any:
                skipped.append(f"'{t}' - عمود فارغ تماماً"); continue
            assessment_columns.append({'index': c, 'title': t, 'due_date': due_dt, 'has_date': due_dt is not None})

        if not assessment_columns:
            st.warning(f"⚠️ '{sheet_name}': لا أعمدة تقييم صالحة")
            if skipped:
                with st.expander(f"📋 الأعمدة المتجاهلة ({len(skipped)})"):
                    for r in skipped[:15]: st.text("• "+r)
            return []

        cols_with_dates = sum(1 for c in assessment_columns if c['has_date'])
        msg = f"✅ '{sheet_name}': {len(assessment_columns)} عمود"
        if filter_active:
            msg += f" ({cols_with_dates} ضمن النطاق" + (f"، {no_date_cols} بدون تاريخ" if no_date_cols>0 else "") + ")"
        st.success(msg)

        student_data, NOT_DUE = {}, {'-','—','–','','NAN','NONE','_'}
        students_count = rows_processed = 0

        for r in range(4, len(df)):
            student = df.iloc[r, 0]
            if pd.isna(student) or str(student).strip() == "": continue
            name = " ".join(str(student).strip().split()); rows_processed += 1
            if name not in student_data:
                student_data[name] = {'total':0,'done':0,'pending':[]}; students_count += 1
            for col in assessment_columns:
                idx = col['index']; title = col['title']
                if idx >= df.shape[1]: continue
                raw = df.iloc[r, idx]
                s = "" if pd.isna(raw) else _strip_invisible_and_diacritics(str(raw)).strip()
                sup = s.upper()
                if sup in NOT_DUE: continue
                if sup == 'M':
                    student_data[name]['total'] += 1
                    if title not in student_data[name]['pending']: student_data[name]['pending'].append(title)
                else:
                    student_data[name]['total'] += 1; student_data[name]['done'] += 1

        results = []
        for name, data in student_data.items():
            total, done, pending = data['total'], data['done'], data['pending']
            if total == 0: continue
            pct = (done/total*100) if total>0 else 0.0
            results.append({
                "student_name": name,
                "subject": subject,
                "level": str(level_from_name).strip(),
                "section": str(section_from_name).strip(),
                "solve_pct": round(pct,1),
                "completed_count": int(done),
                "total_count": int(total),
                "pending_titles": ", ".join(pending) if pending else "-",
                "sheet_name": sheet_name
            })

        if results:
            st.info(f"📊 تمت معالجة {rows_processed} صف ودُمجت إلى {students_count} طالب")
        else:
            st.warning(f"⚠️ '{sheet_name}': لا طلاب بتقييمات مستحقة")
        return results

    except Exception as e:
        st.error(f"❌ خطأ في تحليل '{sheet_name}': {e}")
        import traceback; 
        with st.expander("🔍 تفاصيل الخطأ التقنية"): st.code(traceback.format_exc())
        return []

# ================== تصنيف وفيفوت ==================
def categorize_vectorized(series: pd.Series) -> pd.Series:
    conds = [series>=90, (series>=80)&(series<90), (series>=70)&(series<80),
             (series>=60)&(series<70), (series>0)&(series<60), series==0]
    choices = ['بلاتيني 🥇','ذهبي 🥈','فضي 🥉','برونزي','بحاجة لتحسين','لا يستفيد']
    return pd.Series(np.select(conds, choices, default='لا يستفيد'), index=series.index)

def _canonicalize_level_section(dfc: pd.DataFrame) -> pd.DataFrame:
    """يوحد الصف/الشعبة لكل طالب على الأكثر شيوعًا، لمنع صف مزدوج لنفس الطالب"""
    dfc['level'] = dfc['level'].astype(str).map(_normalize_arabic_digits).str.strip()
    dfc['section'] = dfc['section'].astype(str).apply(_strip_invisible_and_diacritics).str.strip()
    def most_common(series: pd.Series):
        s = series.replace('', np.nan).dropna()
        return s.mode().iloc[0] if not s.mode().empty else ''
    canon = dfc.groupby('student_name').agg(_level=('level', most_common),
                                            _section=('section', most_common)).reset_index()
    dfc = dfc.merge(canon, on='student_name', how='left')
    dfc['level'] = np.where(dfc['level'].eq('')|dfc['level'].isna(), dfc['_level'], dfc['level'])
    dfc['section'] = np.where(dfc['section'].eq('')|dfc['section'].isna(), dfc['_section'], dfc['section'])
    return dfc.drop(columns=['_level','_section'])

@st.cache_data(show_spinner=False)
@log_performance
def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if df is None or df.empty:
            return pd.DataFrame()
        logger.info(f"🔄 معالجة {len(df)} سجل")
        dfc = df.drop_duplicates(subset=['student_name','level','section','subject'], keep='last')
        dfc = _canonicalize_level_section(dfc)

        unique_students = dfc[['student_name','level','section']].drop_duplicates()
        unique_students = unique_students.sort_values(['level','section','student_name']).reset_index(drop=True)
        st.info(f"👥 عدد الطلاب الفريدين: {len(unique_students)}")

        result = unique_students.copy()
        subjects = sorted(dfc['subject'].dropna().unique())
        st.info(f"📚 المواد ({len(subjects)}): {', '.join(subjects)}")

        for subject in subjects:
            sd = dfc[dfc['subject']==subject].copy()
            sd = sd.drop_duplicates(subset=['student_name','level','section'], keep='last')
            for col in ['total_count','completed_count','solve_pct']:
                if col in sd.columns: sd[col] = pd.to_numeric(sd[col], errors='coerce').fillna(0)
            cols = sd[['student_name','level','section','total_count','completed_count','solve_pct']].rename(columns={
                'total_count': f'{subject} - إجمالي','completed_count': f'{subject} - منجز','solve_pct': f'{subject} - النسبة'
            })
            result = result.merge(cols, on=['student_name','level','section'], how='left')
            pend = sd[['student_name','level','section','pending_titles']].rename(columns={'pending_titles': f'{subject} - متبقي'})
            result = result.merge(pend, on=['student_name','level','section'], how='left')

        pct_cols = [c for c in result.columns if 'النسبة' in c]
        if pct_cols:
            def calc_avg(row):
                vals = row[pct_cols].dropna()
                return float(vals.mean()) if len(vals)>0 else 0.0
            result['المتوسط'] = result.apply(calc_avg, axis=1).round(1)
            result['الفئة'] = categorize_vectorized(result['المتوسط'])

        result = result.rename(columns={'student_name':'الطالب','level':'الصف','section':'الشعبة'})

        for c in result.columns:
            if ('إجمالي' in c) or ('منجز' in c): result[c] = result[c].fillna(0).astype(int)
            elif ('النسبة' in c) or (c=='المتوسط'): result[c] = result[c].fillna(0).round(1)
            elif 'متبقي' in c: result[c] = result[c].fillna('-')

        result = result.drop_duplicates(subset=['الطالب','الصف','الشعبة'], keep='first').reset_index(drop=True)
        logger.info(f"✅ Pivot: {len(result)} × {len(result.columns)}")
        return result
    except Exception as e:
        logger.error(f"pivot error: {e}")
        import traceback; 
        with st.expander("🔍 تفاصيل الخطأ"): st.code(traceback.format_exc())
        return pd.DataFrame()

# ============== تحليلات إضافية ==============

def subject_completion_summary(df: pd.DataFrame, section: Optional[str]=None) -> pd.DataFrame:
    """نسبة الحل العامة لكل مادة، مع عدّ الطلاب"""
    if df is None or df.empty: return pd.DataFrame()
    data = df.copy()
    if section:
        # فلترة على الشعبة
        data = data[data['section'].astype(str).str.strip()==str(section).strip()]
    g = data.groupby('subject', dropna=True)
    out = g.agg(
        متوسط_النسبة=('solve_pct', lambda s: round(float(np.nanmean(s)) if len(s)>0 else 0.0, 1)),
        عدد_الطلاب=('student_name','nunique')
    ).reset_index().rename(columns={'subject':'المادة'})
    return out.sort_values('متوسط_النسبة', ascending=False)

def section_completion_summary(df: pd.DataFrame) -> pd.DataFrame:
    """متوسط الإنجاز لكل شعبة (جميع المواد)"""
    if df is None or df.empty: return pd.DataFrame()
    # نحسب متوسط solve_pct للطالب أولاً ثم نأخذ متوسط الشعب
    per_student = df.groupby(['student_name','section'], as_index=False).agg(النسبة=('solve_pct', lambda s: float(np.nanmean(s)) if len(s)>0 else 0.0))
    out = per_student.groupby('section', as_index=False).agg(متوسط_الشعبة=('النسبة', lambda s: round(float(np.nanmean(s)),1)))
    out = out.rename(columns={'section':'الشعبة'}).sort_values('متوسط_الشعبة', ascending=False)
    return out

# ============== رسوم بيانية جديدة ==============

def chart_subject_completion_table_bar(df_subject: pd.DataFrame) -> go.Figure:
    """شريط أفقي لِمتوسط إنجاز كل مادة"""
    if df_subject.empty: return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=df_subject['المادة'],
        x=df_subject['متوسط_النسبة'],
        orientation='h',
        text=[f"{v:.1f}%" for v in df_subject['متوسط_النسبة']],
        textposition='inside',
        marker=dict(color='#8A1538', line=dict(color='white', width=1))
    ))
    fig.update_layout(
        title=dict(text="متوسط نسبة الحل لكل مادة", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        xaxis=dict(range=[0,100], title="النسبة (%)", gridcolor='#E5E7EB'),
        yaxis=dict(autorange='reversed', title="المادة"),
        plot_bgcolor='white', paper_bgcolor='white', font=dict(family='Cairo'),
        height=max(380, len(df_subject)*36)
    )
    return fig

def chart_section_avg_bar(df_section: pd.DataFrame) -> go.Figure:
    """متوسط الإنجاز على مستوى الشعب"""
    if df_section.empty: return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_section['الشعبة'],
        y=df_section['متوسط_الشعبة'],
        text=[f"{v:.1f}%" for v in df_section['متوسط_الشعبة']],
        textposition='outside',
        marker=dict(color='#C9A646', line=dict(color='white', width=1))
    ))
    fig.update_layout(
        title=dict(text="متوسط الإنجاز حسب الشعبة", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        yaxis=dict(range=[0,100], title="النسبة (%)", gridcolor='#E5E7EB'),
        xaxis=dict(title="الشعبة"),
        plot_bgcolor='white', paper_bgcolor='white', font=dict(family='Cairo'),
        height=420
    )
    return fig

# ============== PDF ==============
@safe_execute(default_return=b"", error_message="خطأ في إنشاء PDF")
def make_student_pdf_fpdf(
    school_name: str, student_name: str, grade: str, section: str,
    table_df: pd.DataFrame, overall_avg: float, reco_text: str,
    coordinator_name: str, academic_deputy: str, admin_deputy: str, principal_name: str,
    font_info: Tuple[str, Optional[str]], logo_path: Optional[str] = None,
) -> bytes:
    font_name, font_path = font_info
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)  # ✅ مهم لمنع خطأ المساحة
    pdf.add_page()

    if font_path and os.path.exists(font_path):
        try:
            pdf.add_font(font_name, "", font_path, uni=True)
        except Exception as e:
            logger.warning(f"خط: {e}"); font_name = ""

    def set_font(size=12, color=(0,0,0)):
        if font_name: pdf.set_font(font_name, size=size)
        else: pdf.set_font("Helvetica", size=size)
        pdf.set_text_color(*color)

    pdf.set_fill_color(*QATAR_MAROON); pdf.rect(0, 0, 210, 20, style="F")
    if logo_path and os.path.exists(logo_path):
        try: pdf.image(logo_path, x=185, y=2.5, w=20)
        except Exception as e: logger.warning(f"شعار: {e}")

    set_font(14, (255,255,255)); pdf.set_xy(10,7)
    pdf.cell(0, 8, rtl("إنجاز - تقرير أداء الطالب"), align="R")

    set_font(18, QATAR_MAROON); pdf.set_y(28)
    pdf.cell(0, 10, rtl("تقرير أداء الطالب - نظام قطر للتعليم"), ln=1, align="R")
    pdf.set_draw_color(*QATAR_GOLD); pdf.set_line_width(0.6); pdf.line(30,38,200,38)

    set_font(12); pdf.ln(6)
    pdf.cell(0,8, rtl(f"اسم المدرسة: {school_name or '—'}"), ln=1, align="R")
    pdf.cell(0,8, rtl(f"اسم الطالب: {student_name}"), ln=1, align="R")
    pdf.cell(0,8, rtl(f"الصف: {grade or '—'}     الشعبة: {section or '—'}"), ln=1, align="R")
    pdf.ln(2)

    # جدول — عرض <= 190مم
    headers = [rtl("المادة"), rtl("عدد التقييمات الإجمالي"), rtl("عدد التقييمات المنجزة"), rtl("عدد التقييمات المتبقية")]
    widths = [76, 38, 38, 38]  # 190مم
    pdf.set_fill_color(*QATAR_MAROON); set_font(12, (255,255,255))
    pdf.set_y(pdf.get_y()+4)
    for w,h in zip(widths, headers): pdf.cell(w, 9, h, border=0, align="C", fill=True)
    pdf.ln(9)

    set_font(11); total_total = total_solved = 0
    for _, r in table_df.iterrows():
        sub = rtl(str(r['المادة'])); tot = int(r['إجمالي']); solv = int(r['منجز']); rem = int(max(tot-solv,0))
        total_total += tot; total_solved += solv
        pdf.set_fill_color(247,247,247)
        pdf.cell(widths[0], 8, sub, 0, 0, "C", True)
        pdf.cell(widths[1], 8, str(tot), 0, 0, "C", True)
        pdf.cell(widths[2], 8, str(solv), 0, 0, "C", True)
        pdf.cell(widths[3], 8, str(rem), 0, 1, "C", True)

    pdf.ln(3); set_font(12, QATAR_MAROON); pdf.cell(0,8, rtl("الإحصائيات"), ln=1, align="R")
    set_font(12); remaining = max(total_total-total_solved, 0)
    pdf.cell(0,8, rtl(f"منجز: {total_solved}    متبقي: {remaining}    نسبة حل التقييمات: {overall_avg:.1f}%"), ln=1, align="R")

    pdf.ln(2); set_font(12, QATAR_MAROON); pdf.cell(0,8, rtl("توصية منسق المشاريع:"), ln=1, align="R")
    set_font(11); 
    for line in (reco_text or "—").splitlines():
        pdf.multi_cell(0, 7, rtl(line), align="R")

    pdf.ln(2); set_font(12, QATAR_MAROON); pdf.cell(0,8, rtl("روابط مهمة:"), ln=1, align="R")
    set_font(11)
    pdf.cell(0,7, rtl("رابط نظام قطر: https://portal.education.qa"), ln=1, align="R")
    pdf.cell(0,7, rtl("استعادة كلمة المرور: https://password.education.qa"), ln=1, align="R")
    pdf.cell(0,7, rtl("قناة قطر للتعليم: https://edu.tv.qa"), ln=1, align="R")

    pdf.ln(4); set_font(12, QATAR_MAROON); pdf.cell(0,8, rtl("التوقيعات"), ln=1, align="R")
    set_font(11); pdf.set_draw_color(*QATAR_GOLD)
    boxes = [("منسق المشاريع", coordinator_name), ("النائب الأكاديمي", academic_deputy),
             ("النائب الإداري", admin_deputy), ("مدير المدرسة", principal_name)]
    x_left, x_right = 10, 110; y0 = pdf.get_y()+2; w, h = 90, 18
    for i, (title, name) in enumerate(boxes):
        row, col = i//2, i%2
        x = x_right if col==0 else x_left; yb = y0 + row*(h+6)
        pdf.rect(x, yb, w, h)
        set_font(11)
        pdf.set_xy(x, yb+3); pdf.cell(w-4, 6, rtl(f"{title} / {name or '—'}"), align="R")
        pdf.set_xy(x, yb+10); pdf.cell(w-4, 6, rtl("التوقيع: __________________    التاريخ: __________"), align="R")

    out = pdf.output(dest="S")
    return out if isinstance(out,(bytes,bytearray)) else str(out).encode("latin-1","ignore")

# ============== التطبيق ==============
def main():
    setup_app()

    # === Sidebar ===
    with st.sidebar:
        st.image("https://i.imgur.com/XLef7tS.png", width=110)
        st.markdown("---")
        st.header("⚙️ الإعدادات")

        uploaded_files = st.file_uploader("اختر ملفات Excel", type=["xlsx","xls"], accept_multiple_files=True)
        selected_sheets, all_sheets, sheet_file_map = [], [], {}
        if uploaded_files:
            for i, f in enumerate(uploaded_files):
                try:
                    f.seek(0); xls = pd.ExcelFile(f)
                    for s in xls.sheet_names:
                        label = f"[ملف {i+1}] {s}"
                        all_sheets.append(label); sheet_file_map[label]=(f, s)
                except Exception as e:
                    st.error(f"قراءة الملف فشلت: {e}")
            if all_sheets:
                st.info(f"📋 {len(all_sheets)} ورقة / {len(uploaded_files)} ملف")
                select_all = st.checkbox("✔️ اختر الجميع", value=True, key="select_all_sheets")
                chosen = all_sheets if select_all else st.multiselect("اختر الأوراق للتحليل", all_sheets, default=all_sheets[:1])
                selected_sheets = [sheet_file_map[c] for c in chosen]
        st.session_state.selected_sheets = selected_sheets

        st.subheader("⏳ فلترة الأعمدة حسب تاريخ الاستحقاق")
        enable_date_filter = st.checkbox("تفعيل فلتر التاريخ", value=False,
            help="يقرأ التاريخ من H3 لكل عمود. الأعمدة خارج النطاق تُستبعد.")
        if enable_date_filter:
            default_start = date.today().replace(day=1); default_end = date.today()
            st.info("ℹ️ تحليل الأعمدة ضمن النطاق فقط")
            rng = st.date_input("اختر المدى", value=(default_start, default_end), format="YYYY-MM-DD")
            if isinstance(rng,(list,tuple)) and len(rng)>=2:
                due_start, due_end = rng[0], rng[1]
            else:
                due_start, due_end = None, None
        else:
            due_start, due_end = None, None
            st.success("✅ المنطق: '-' غير مستحق | M متبقي | القيمة = منجز")

        st.subheader("🖼️ شعار المدرسة (اختياري)")
        logo_file = st.file_uploader("ارفع شعار PNG/JPG", type=["png","jpg","jpeg"], key="logo_file")
        st.session_state.logo_path = prepare_logo_file(logo_file)

        st.markdown("---")
        st.subheader("🏫 معلومات المدرسة")
        school_name = st.text_input("اسم المدرسة", placeholder="مدرسة قطر النموذجية")

        st.subheader("✍️ التوقيعات")
        coordinator_name = st.text_input("منسق/ة المشاريع")
        academic_deputy  = st.text_input("النائب الأكاديمي")
        admin_deputy     = st.text_input("النائب الإداري")
        principal_name   = st.text_input("مدير/ة المدرسة")

        st.markdown("---")
        run_analysis = st.button("▶️ تشغيل التحليل", use_container_width=True, type="primary", disabled=not uploaded_files)

    # === Main content ===
    if not uploaded_files:
        st.info("📤 ارفع ملفات Excel من الشريط الجانبي للبدء")

    elif run_analysis:
        sheets_to_use = st.session_state.selected_sheets
        if not sheets_to_use:
            tmp = []
            for f in uploaded_files:
                try:
                    f.seek(0); xls = pd.ExcelFile(f)
                    for s in xls.sheet_names: tmp.append((f,s))
                except Exception as e:
                    st.error(f"قراءة الملف فشلت: {e}")
            sheets_to_use = tmp

        if not sheets_to_use:
            st.warning("⚠️ لم يتم العثور على أوراق داخل الملفات.")
        else:
            prog = st.progress(0); status = st.empty()
            with st.spinner("⏳ جاري التحليل..."):
                rows = []; total = len(sheets_to_use)
                for i,(f,s) in enumerate(sheets_to_use):
                    prog.progress((i+1)/total); status.text(f"📊 تحليل الورقة {i+1}/{total}: {s}")
                    f.seek(0); rows.extend(analyze_excel_file(f, s, due_start, due_end))
                prog.empty(); status.empty()

                if rows:
                    df = pd.DataFrame(rows)
                    st.session_state.analysis_results = df
                    st.session_state.pivot_table = create_pivot_table(df)
                    subjects_count = df['subject'].nunique() if 'subject' in df.columns else 0
                    students_count = len(st.session_state.pivot_table)
                    st.success(f"✅ تم تحليل {students_count} طالب عبر {subjects_count} مادة")
                    st.session_state.analysis_stats = {
                        'students': students_count, 'subjects': subjects_count,
                        'total_assessments': df['total_count'].sum() if 'total_count' in df.columns else 0,
                        'completed': df['completed_count'].sum() if 'completed_count' in df.columns else 0,
                    }
                else:
                    st.warning("⚠️ لم يتم استخراج بيانات. تأكد من التنسيق.")

    # === عرض النتائج ===
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results

    if pivot is not None and not pivot.empty and df is not None:
        # ملخص
        st.subheader("📈 ملخص النتائج")
        c1,c2,c3,c4,c5 = st.columns(5)
        with c1: st.metric("👥 إجمالي الطلاب", len(pivot))
        with c2:
            subjects = df['subject'].nunique() if 'subject' in df.columns else 0
            st.metric("📚 عدد المواد", subjects)
        with c3:
            avg = float(pivot['المتوسط'].mean()) if 'المتوسط' in pivot.columns else 0.0
            st.metric("📊 متوسط الإنجاز", f"{avg:.1f}%")
        with c4:
            pcount = int((pivot['الفئة']=='بلاتيني 🥇').sum()) if 'الفئة' in pivot.columns else 0
            st.metric("🥇 فئة بلاتيني", pcount)
        with c5:
            zero = int((pivot['المتوسط']==0).sum()) if 'المتوسط' in pivot.columns else 0
            st.metric("⚠️ بدون إنجاز", zero)

        st.divider()

        # ================== فلاتر إضافية (الصف + الشعبة + الفئة) ==================
        st.subheader("📋 جدول النتائج التفصيلي")
        colf1, colf2, colf3 = st.columns(3)
        with colf1:
            levels = ['الكل'] + sorted(pivot['الصف'].dropna().unique().tolist()) if 'الصف' in pivot.columns else ['الكل']
            selected_level = st.selectbox("فلتر حسب الصف", levels)
        with colf2:
            sections = ['الكل'] + sorted(pivot['الشعبة'].dropna().unique().tolist()) if 'الشعبة' in pivot.columns else ['الكل']
            selected_section = st.selectbox("فلتر حسب الشعبة", sections)
        with colf3:
            categories = ['الكل'] + CATEGORY_ORDER if 'الفئة' in pivot.columns else ['الكل']
            selected_category = st.selectbox("فلتر حسب الفئة", categories)

        filtered_pivot = pivot.copy()
        if selected_level != 'الكل':
            filtered_pivot = filtered_pivot[filtered_pivot['الصف']==selected_level]
        if selected_section != 'الكل':
            filtered_pivot = filtered_pivot[filtered_pivot['الشعبة']==selected_section]
        if selected_category != 'الكل':
            filtered_pivot = filtered_pivot[filtered_pivot['الفئة']==selected_category]

        st.dataframe(filtered_pivot, use_container_width=True, height=420)
        csv = filtered_pivot.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 تحميل الجدول (CSV)", csv,
                           f"ingaz_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                           "text/csv", key='download-csv')

        st.divider()

        # ================== رسوم موجودة ==================
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">🍩 التوزيع العام للفئات</h2>', unsafe_allow_html=True)
        counts = filtered_pivot['الفئة'].value_counts().reindex(CATEGORY_ORDER, fill_value=0) if 'الفئة' in filtered_pivot.columns else pd.Series(dtype=int)
        if not counts.empty:
            fig_donut = go.Figure([go.Pie(labels=counts.index, values=counts.values, hole=0.55,
                marker=dict(colors=[CATEGORY_COLORS[k] for k in counts.index]),
                textinfo='label+value', textfont=dict(size=13, family='Cairo'),
                hovertemplate="%{label}: %{value} طالب<extra></extra>")])
            fig_donut.update_layout(title=dict(text="توزيع عام للفئات", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
                                    showlegend=False, height=400, paper_bgcolor='white', plot_bgcolor='white')
            st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.info("لا توجد بيانات لعرض التوزيع.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">🎯 مؤشر متوسط الإنجاز</h2>', unsafe_allow_html=True)
        avg2 = float(filtered_pivot['المتوسط'].mean()) if 'المتوسط' in filtered_pivot.columns and not filtered_pivot.empty else 0.0
        fig_g = go.Figure(go.Indicator(
            mode="gauge+number", value=avg2, number={'suffix': "%", 'font': {'family':'Cairo','size':40}},
            gauge={'axis':{'range':[0,100],'tickfont':{'family':'Cairo'}}, 'bar':{'color':'#8A1538'},
                   'steps':[{'range':[0,60],'color':'#ffebee'},{'range':[60,70],'color':'#fff3e0'},
                            {'range':[70,80],'color':'#f1f8e9'},{'range':[80,90],'color':'#e8f5e9'},
                            {'range':[90,100],'color':'#e0f7fa'}],
                   'threshold':{'line':{'color':CATEGORY_COLORS['ذهبي 🥈'],'width':4},'thickness':0.75,'value':80}}))
        fig_g.update_layout(title=dict(text="متوسط الإنجاز العام", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
                            paper_bgcolor='white', plot_bgcolor='white', height=350, font=dict(family='Cairo'))
        st.plotly_chart(fig_g, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ================== تحليل نسبة الحل العامة لكل مادة + فلترة الشعبة ==================
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">📚 نسبة الحل العامة لكل مادة</h2>', unsafe_allow_html=True)
        # نفلتر DataFrame الخام بنفس فلاتر الصف/الشعبة
        df_filtered = df.copy()
        if selected_level != 'الكل':
            df_filtered = df_filtered[df_filtered['level'].astype(str)==str(selected_level)]
        if selected_section != 'الكل':
            df_filtered = df_filtered[df_filtered['section'].astype(str)==str(selected_section)]
        sub_tbl = subject_completion_summary(df_filtered, section=None)  # section أخذناه من الفلتر أعلاه
        if not sub_tbl.empty:
            st.dataframe(sub_tbl, use_container_width=True, height=300)
            st.plotly_chart(chart_subject_completion_table_bar(sub_tbl), use_container_width=True)
        else:
            st.info("لا توجد بيانات كافية لحساب متوسطات المواد.")
        st.markdown('</div>', unsafe_allow_html=True)

        # ================== رسم بياني على مستوى الشعب ==================
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">🏷️ متوسط الإنجاز على مستوى الشعب</h2>', unsafe_allow_html=True)
        df_for_sections = df.copy()
        if selected_level != 'الكل':
            df_for_sections = df_for_sections[df_for_sections['level'].astype(str)==str(selected_level)]
        sec_tbl = section_completion_summary(df_for_sections)
        if not sec_tbl.empty:
            st.dataframe(sec_tbl, use_container_width=True, height=260)
            st.plotly_chart(chart_section_avg_bar(sec_tbl), use_container_width=True)
        else:
            st.info("لا توجد بيانات كافية لعرض متوسطات الشعب.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.divider()

        # ================== التقارير الفردية ==================
        st.subheader("📑 التقارير الفردية (PDF)")
        students = sorted(filtered_pivot['الطالب'].dropna().astype(str).unique().tolist()) if 'الطالب' in filtered_pivot.columns else []
        if students:
            csel, crec = st.columns([2,3])
            with csel:
                sel = st.selectbox("اختر الطالب", students, index=0)
                row = filtered_pivot[filtered_pivot['الطالب']==sel].head(1)
                g = str(row['الصف'].iloc[0]) if not row.empty else ''
                s = str(row['الشعبة'].iloc[0]) if not row.empty else ''
                student_category = row['الفئة'].iloc[0] if not row.empty and 'الفئة' in row.columns else ''
            with crec:
                from_text = {
                    'بلاتيني 🥇': f"أداء متميز جدًا. استمر {sel} بهذا النسق.",
                    'ذهبي 🥈': f"أداء قوي وقابل للتحسن نحو البلاتيني.",
                    'فضي 🥉': f"أداء جيد ويحتاج مزيد التثبيت.",
                    'برونزي': f"نوصي بزيادة الالتزام والمتابعة.",
                    'بحاجة لتحسين': f"يلزم خطة متابعة مكثفة.",
                    'لا يستفيد': f"ابدأ فورًا بتفعيل النظام؛ نحن هنا للمساعدة."
                }
                auto = from_text.get(student_category, "نوصي بالمتابعة المستمرة.")
                reco = st.text_area("توصية منسق المشاريع", value=auto, height=180)

            sdata = df[(df['student_name'].str.strip()==sel.strip())].copy()
            if not sdata.empty:
                table = sdata[['subject','total_count','completed_count']].rename(columns={'subject':'المادة','total_count':'إجمالي','completed_count':'منجز'})
                table['متبقي'] = (table['إجمالي'] - table['منجز']).clip(lower=0).astype(int)
                avg_stu = float(sdata['solve_pct'].mean()) if 'solve_pct' in sdata.columns else 0.0
                if pd.isna(avg_stu): avg_stu = 0.0
                st.markdown("### معاينة سريعة"); st.dataframe(table, use_container_width=True, height=260)
                pdf_one = make_student_pdf_fpdf(
                    school_name=school_name or "", student_name=sel, grade=g, section=s,
                    table_df=table[['المادة','إجمالي','منجز','متبقي']], overall_avg=avg_stu, reco_text=reco,
                    coordinator_name=coordinator_name or "", academic_deputy=academic_deputy or "",
                    admin_deputy=admin_deputy or "", principal_name=principal_name or "",
                    font_info=st.session_state.font_info, logo_path=st.session_state.logo_path
                )
                if pdf_one:
                    st.download_button("📥 تحميل تقرير الطالب (PDF)", pdf_one,
                        file_name=f"student_report_{sel}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf", use_container_width=True)

            st.markdown("---")
            st.subheader("📦 تصدير جميع التقارير (ZIP)")
            use_auto_all = st.checkbox("✨ توصيات تلقائية لكل طالب", value=True)
            same_reco = st.checkbox("استخدم نفس التوصية للجميع", value=True) if not use_auto_all else False

            if st.button("إنشاء ملف ZIP لكل التقارير", type="primary"):
                with st.spinner("جاري إنشاء الحزمة..."):
                    try:
                        buf = io.BytesIO()
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                            for stu in students:
                                r = filtered_pivot[filtered_pivot['الطالب']==stu].head(1)
                                g = str(r['الصف'].iloc[0]) if not r.empty else ''
                                s = str(r['الشعبة'].iloc[0]) if not r.empty else ''
                                sd = df[df['student_name'].str.strip()==stu.strip()].copy()
                                if sd.empty: continue
                                t = sd[['subject','total_count','completed_count']].rename(columns={'subject':'المادة','total_count':'إجمالي','completed_count':'منجز'})
                                t['متبقي'] = (t['إجمالي'] - t['منجز']).clip(lower=0).astype(int)
                                av = float(sd['solve_pct'].mean()) if 'solve_pct' in sd.columns else 0.0
                                if pd.isna(av): av = 0.0
                                if use_auto_all:
                                    cat = r['الفئة'].iloc[0] if not r.empty and 'الفئة' in r.columns else ''
                                    auto_map = {'بلاتيني 🥇':"أداء متميز جدًا.",
                                                'ذهبي 🥈':"أداء قوي وقابل للتحسن.",
                                                'فضي 🥉':"أداء جيد.",
                                                'برونزي':"نوصي بزيادة الالتزام.",
                                                'بحاجة لتحسين':"خطة متابعة مكثفة.",
                                                'لا يستفيد':"ابدأ بتفعيل النظام."}
                                    rtext = auto_map.get(cat, "نوصي بالمتابعة المستمرة.")
                                elif same_reco:
                                    rtext = reco
                                else:
                                    rtext = ""
                                pdfb = make_student_pdf_fpdf(
                                    school_name=school_name or "", student_name=stu, grade=g, section=s,
                                    table_df=t[['المادة','إجمالي','منجز','متبقي']], overall_avg=av, reco_text=rtext,
                                    coordinator_name=coordinator_name or "", academic_deputy=academic_deputy or "",
                                    admin_deputy=admin_deputy or "", principal_name=principal_name or "",
                                    font_info=st.session_state.font_info, logo_path=st.session_state.logo_path
                                )
                                if pdfb:
                                    safe = re.sub(r"[^\w\-]+", "_", str(stu))
                                    z.writestr(f"{safe}.pdf", pdfb)
                        buf.seek(0)
                        st.download_button("⬇️ تحميل الحزمة (ZIP)", buf.getvalue(),
                                           file_name=f"student_reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                           mime="application/zip", use_container_width=True)
                        st.success(f"✅ تم إنشاء {len(students)} تقرير")
                    except Exception as e:
                        st.error(f"❌ خطأ في إنشاء الحزمة: {e}")

    # فوتر
    st.markdown(f"""
    <div style="margin-top:22px;background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);color:#fff;border-radius:10px;padding:12px 10px;text-align:center;">
      <div style="font-weight:800;font-size:15px;margin:2px 0 4px">مدرسة عثمان بن عفان النموذجية للبنين</div>
      <div style="font-weight:700;font-size:12px;margin:0 0 4px;opacity:.95">© {datetime.now().year} جميع الحقوق محفوظة</div>
    </div>
    """, unsafe_allow_html=True)

# ============== تشغيل ==============
if __name__ == "__main__":
    main()

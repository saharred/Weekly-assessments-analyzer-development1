# -*- coding: utf-8 -*-
"""
إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم
النسخة 2.4
- هوية عنابية/ذهبية جمالية
- تحليل مواد وشُعب ورسوم
- تقارير PDF/ZIP
- رفع بيانات المعلمات والربط بالشُعب
- مراسلة المعلمات بضغطة (mailto) مع توصية وأسماء الطلاب "لا يستفيد"
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
from urllib.parse import quote  # لترميز نص البريد

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
    warnings.warn("⚠️ arabic_reshaper غير مثبت — للتثبيت: pip install arabic-reshaper python-bidi")

def rtl(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    if AR_OK:
        try: return get_display(arabic_reshaper.reshape(text))
        except Exception: return text
    return text

# ================== ثوابت وهوية ==================
QATAR_MAROON = (138, 21, 56)       # RGB
QATAR_GOLD   = (201, 166, 70)

CATEGORY_COLORS = {
    'بلاتيني 🥇': '#E5E4E2',
    'ذهبي 🥈'   : '#C9A646',
    'فضي 🥉'    : '#C0C0C0',
    'برونزي'    : '#CD7F32',
    'بحاجة لتحسين': '#FF9800',
    'لا يستفيد' : '#8A1538'
}
CATEGORY_ORDER = ['بلاتيني 🥇', 'ذهبي 🥈', 'فضي 🥉', 'برونزي', 'بحاجة لتحسين', 'لا يستفيد']

# ================== لوجينغ ==================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger("ingaz-app")

def log_performance(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        t0 = time.time(); name = func.__name__
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
            try: return func(*args, **kwargs)
            except Exception as e:
                logger.error(f"{error_message} في {func.__name__}: {e}")
                if st: st.error(f"{error_message}: {e}")
                return default_return
        return wrapper
    return deco

# ================== أدوات نص/أرقام ==================
def _normalize_arabic_digits(s: str) -> str:
    if not isinstance(s, str): return str(s) if s is not None else ""
    return s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

def _strip_invisible_and_diacritics(s: str) -> str:
    if not isinstance(s, str): return ""
    invis = ['\u200e','\u200f','\u202a','\u202b','\u202c','\u202d','\u202e',
             '\u2066','\u2067','\u2068','\u2069','\u200b','\u200c','\u200d',
             '\ufeff','\xa0','\u0640']
    for ch in invis: s = s.replace(ch, '')
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = ' '.join(s.split())
    return s.strip()

def _norm_month_key(s: str) -> str:
    s = _strip_invisible_and_diacritics(s)
    s = _normalize_arabic_digits(s).lower().strip()
    s = s.replace("أ","ا").replace("إ","ا").replace("آ","ا").replace("ة","ه").replace("ـ","")
    return s

_AR_MONTHS_BASE = {
    "يناير":1,"فبراير":2,"مارس":3,"ابريل":4,"أبريل":4,"مايو":5,"يونيو":6,"يوليو":7,"اغسطس":8,"أغسطس":8,"سبتمبر":9,"اكتوبر":10,"أكتوبر":10,"نوفمبر":11,"ديسمبر":12,
    "كانون الثاني":1,"شباط":2,"اذار":3,"نيسان":4,"ايار":5,"حزيران":6,"تموز":7,"اب":8,"ايلول":9,"تشرين الاول":10,"تشرين الثاني":11,"كانون الاول":12,
    "جانفي":1,"فيفري":2,"افريل":4,"ماي":5,"جوان":6,"جويلية":7,"اوت":8,"دجنبر":12
}
_AR_MONTHS = { _norm_month_key(k): v for k, v in _AR_MONTHS_BASE.items() }

_EN_MONTHS = {
    "jan":1,"january":1,"feb":2,"february":2,"mar":3,"march":3,"apr":4,"april":4,"may":5,
    "jun":6,"june":6,"jul":7,"july":7,"aug":8,"august":8,"sep":9,"sept":9,"september":9,
    "oct":10,"october":10,"nov":11,"november":11,"dec":12,"december":12
}

# ================== تواريخ ==================
def parse_due_date_cell(cell, default_year: int = None) -> Optional[date]:
    if default_year is None: default_year = date.today().year
    if cell is None or (isinstance(cell, float) and pd.isna(cell)): return None
    if isinstance(cell, (pd.Timestamp, datetime)):
        try: return cell.date() if hasattr(cell, 'date') else cell
        except Exception: return None
    if isinstance(cell, (int, float)) and not pd.isna(cell):
        try:
            if 1 <= cell <= 100000:
                base = pd.to_datetime("1899-12-30")
                result = base + pd.to_timedelta(float(cell), unit="D")
                if 1900 <= result.year <= 2200: return result.date()
        except Exception: pass
    try:
        s = str(cell).strip()
        if not s or s.lower() in ['nan','none','nat','null']: return None
        s_clean = _strip_invisible_and_diacritics(_normalize_arabic_digits(s))
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
    except Exception: pass
    return None

def in_range(d: Optional[date], start: Optional[date], end: Optional[date]) -> bool:
    if not (start and end): return True
    if d is None: return False
    if start > end: start, end = end, start
    return start <= d <= end

# ================== تحقق بنية ==================
def validate_excel_structure(df: pd.DataFrame, sheet_name: str) -> Tuple[bool, str]:
    if df is None or df.empty: return False, "الملف فارغ"
    if df.shape[0] < 4: return False, f"عدد الصفوف قليل جداً ({df.shape[0]} صف)"
    if df.shape[1] < 8: return False, f"عدد الأعمدة قليل جداً ({df.shape[1]} عمود)"
    if len(df.iloc[4:, 0].dropna()) == 0: return False, "لا توجد أسماء طلاب في العمود الأول"
    return True, ""

# ================== واجهة ==================
def setup_app():
    APP_TITLE = "إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم"
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="https://i.imgur.com/XLef7tS.png",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    defaults = {
        "analysis_results": None,
        "pivot_table": None,
        "font_info": None,
        "logo_path": None,
        "selected_sheets": [],
        "analysis_stats": {},
        "teachers_df": None,  # ✅ سنخزن بيانات المعلمات هنا
    }
    for k, v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v
    if st.session_state.font_info is None:
        st.session_state.font_info = prepare_default_font()
    apply_custom_styles()
    render_header(APP_TITLE)

def apply_custom_styles():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
    *{font-family:'Cairo','Segoe UI',-apple-system,sans-serif;direction:rtl}
    .main,body,.stApp{background:#fff;direction:rtl}
    section[data-testid="stSidebar"]{right:0!important;left:auto!important}
    .main .block-container{padding-right:4.5rem!important;padding-left:1rem!important}

    /* Header */
    .header-container{
      background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
      padding:36px 28px;color:#fff;text-align:center;margin-bottom:18px;
      border-bottom:4px solid #C9A646;box-shadow:0 6px 20px rgba(138,21,56,.25);position:relative}
    .header-container::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;
      background:linear-gradient(90deg,#C9A646 0%,#E8D4A0 50%,#C9A646 100%)}
    .header-container h1{margin:0 0 6px 0;font-size:30px;font-weight:800;letter-spacing:.2px}
    .header-sub{font-size:15px;font-weight:700;color:#EADAA9;margin-top:4px}

    /* Cards/Charts */
    .chart-container{background:#fff;border:2px solid #E9E9EE;border-right:6px solid #8A1538;
      border-radius:14px;padding:16px;margin:12px 0;box-shadow:0 2px 10px rgba(0,0,0,.07)}
    .chart-title{font-size:20px;font-weight:800;color:#8A1538;text-align:center;margin-bottom:10px}

    /* Sidebar */
    [data-testid="stSidebar"]{
      background:linear-gradient(180deg,#8A1538 0%,#6B1029 100%)!important;
      border-left:2px solid #C9A646;box-shadow:-4px 0 16px rgba(0,0,0,.15)}
    [data-testid="stSidebar"] *{color:#fff!important}
    [data-testid="stSidebar"] input,[data-testid="stSidebar"] textarea,[data-testid="stSidebar"] select{
      color:#000!important;background:#fff!important;text-align:right}
    .stProgress > div > div{background:#8A1538!important}
    </style>
    """, unsafe_allow_html=True)

def render_header(title: str):
    st.markdown(f"""
    <div class='header-container'>
      <h1>{title}</h1>
      <div class='header-sub'>نحو تنمية رقمية مستدامة</div>
      <div style="font-size:12px;opacity:.95;margin-top:6px;">المنطق: '-' غير مستحق | M متبقي | القيمة = منجز</div>
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
        if os.path.exists(p):
            return font_name, p
    return "", None

@safe_execute(default_return=None, error_message="خطأ في معالجة الشعار")
def prepare_logo_file(logo_file) -> Optional[str]:
    if logo_file is None: return None
    ext = os.path.splitext(logo_file.name)[1].lower()
    if ext not in [".png", ".jpg", ".jpeg"]:
        st.warning("⚠️ يرجى رفع شعار PNG/JPG"); return None
    logo_file.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(logo_file.read()); return tmp.name

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
                       due_start: Optional[date] = None, due_end: Optional[date] = None) -> List[Dict[str, Any]]:
    try:
        data = file.getvalue() if hasattr(file, "getvalue") else file.read()
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, header=None)
        ok, msg = validate_excel_structure(df, sheet_name)
        if not ok:
            st.error(f"❌ '{sheet_name}': {msg}")
            return []
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        filter_active = (due_start is not None and due_end is not None)
        if filter_active and due_start > due_end: due_start, due_end = due_end, due_start

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
                if due_dt is None: no_date_cols += 1
                elif not in_range(due_dt, due_start, due_end):
                    skipped.append(f"'{t}' - خارج النطاق ({due_dt})"); continue
            has_any = any(pd.notna(df.iloc[r, c]) for r in range(4, min(len(df), 50)) if r < df.shape[0])
            if not has_any:
                skipped.append(f"'{t}' - عمود فارغ تماماً"); continue
            assessment_columns.append({'index': c, 'title': t, 'due_date': due_dt, 'has_date': due_dt is not None})

        if not assessment_columns:
            st.warning(f"⚠️ '{sheet_name}': لا أعمدة تقييم صالحة")
            return []

        student_data, NOT_DUE = {}, {'-','—','–','','NAN','NONE','_'}
        rows_processed = 0
        for r in range(4, len(df)):
            student = df.iloc[r, 0]
            if pd.isna(student) or str(student).strip() == "": continue
            name = " ".join(str(student).strip().split())
            rows_processed += 1
            if name not in student_data: student_data[name] = {'total': 0, 'done': 0, 'pending': []}
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
            pct = (done / total * 100) if total > 0 else 0.0
            results.append({
                "student_name": name, "subject": subject,
                "level": str(level_from_name).strip(), "section": str(section_from_name).strip(),
                "solve_pct": round(pct, 1), "completed_count": int(done), "total_count": int(total),
                "pending_titles": ", ".join(pending) if pending else "-", "sheet_name": sheet_name
            })
        if results:
            st.info(f"📊 تمت معالجة {rows_processed} صف")
        else:
            st.warning(f"⚠️ '{sheet_name}': لا طلاب بتقييمات مستحقة")
        return results

    except Exception as e:
        st.error(f"❌ خطأ في تحليل '{sheet_name}': {e}")
        import traceback
        with st.expander("🔍 تفاصيل الخطأ التقنية"):
            st.code(traceback.format_exc())
        return []

# ================== تصنيف وفيفوت ==================
def categorize_vectorized(series: pd.Series) -> pd.Series:
    conds = [
        series >= 90,
        (series >= 80) & (series < 90),
        (series >= 70) & (series < 80),
        (series >= 60) & (series < 70),
        (series > 0) & (series < 60),
        series == 0
    ]
    choices = ['بلاتيني 🥇','ذهبي 🥈','فضي 🥉','برونزي','بحاجة لتحسين','لا يستفيد']
    return pd.Series(np.select(conds, choices, default='لا يستفيد'), index=series.index)

def _canonicalize_level_section(dfc: pd.DataFrame) -> pd.DataFrame:
    dfc['level']   = dfc['level'].astype(str).map(_normalize_arabic_digits).str.strip()
    dfc['section'] = dfc['section'].astype(str).apply(_strip_invisible_and_diacritics).str.strip()
    def most_common(s: pd.Series):
        s = s.replace('', np.nan).dropna()
        return s.mode().iloc[0] if not s.mode().empty else ''
    canon = dfc.groupby('student_name').agg(_level=('level', most_common),
                                            _section=('section', most_common)).reset_index()
    dfc = dfc.merge(canon, on='student_name', how='left')
    dfc['level']   = np.where(dfc['level'].eq('')  | dfc['level'].isna(),   dfc['_level'],   dfc['level'])
    dfc['section'] = np.where(dfc['section'].eq('')| dfc['section'].isna(), dfc['_section'], dfc['section'])
    return dfc.drop(columns=['_level','_section'])

@st.cache_data(show_spinner=False)
@log_performance
def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if df is None or df.empty: return pd.DataFrame()
        dfc = df.drop_duplicates(subset=['student_name','level','section','subject'], keep='last')
        dfc = _canonicalize_level_section(dfc)
        unique_students = dfc[['student_name','level','section']].drop_duplicates().sort_values(
            ['level','section','student_name']).reset_index(drop=True)
        result = unique_students.copy()
        subjects = sorted(dfc['subject'].dropna().unique())
        for subject in subjects:
            sd = dfc[dfc['subject']==subject].copy()
            sd = sd.drop_duplicates(subset=['student_name','level','section'], keep='last')
            for col in ['total_count','completed_count','solve_pct']:
                if col in sd.columns: sd[col] = pd.to_numeric(sd[col], errors='coerce').fillna(0)
            cols = sd[['student_name','level','section','total_count','completed_count','solve_pct']].rename(columns={
                'total_count': f'{subject} - إجمالي',
                'completed_count': f'{subject} - منجز',
                'solve_pct': f'{subject} - النسبة'
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
            result['الفئة']    = categorize_vectorized(result['المتوسط'])
        result = result.rename(columns={'student_name':'الطالب','level':'الصف','section':'الشعبة'})
        for c in result.columns:
            if ('إجمالي' in c) or ('منجز' in c): result[c] = result[c].fillna(0).astype(int)
            elif ('النسبة' in c) or (c=='المتوسط'): result[c] = result[c].fillna(0).round(1)
            elif 'متبقي' in c: result[c] = result[c].fillna('-')
        result = result.drop_duplicates(subset=['الطالب','الصف','الشعبة'], keep='first').reset_index(drop=True)
        return result
    except Exception as e:
        logger.error(f"pivot error: {e}")
        import traceback
        with st.expander("🔍 تفاصيل الخطأ"):
            st.code(traceback.format_exc())
        return pd.DataFrame()

# ============== تحليلات إضافية ==============
def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if 'solve_pct' in out.columns: out = out.rename(columns={'solve_pct':'percent'})
    if 'student_name' in out.columns: out = out.rename(columns={'student_name':'student'})
    if 'percent' in out.columns: out['category'] = categorize_vectorized(out['percent'])
    return out

def aggregate_by_subject(df_norm: pd.DataFrame) -> pd.DataFrame:
    if df_norm.empty: return pd.DataFrame()
    rows = []
    for s in df_norm['subject'].dropna().unique():
        sub = df_norm[df_norm['subject'] == s]; n = len(sub)
        for cat in CATEGORY_ORDER:
            cnt = (sub['category'] == cat).sum() if 'category' in sub.columns else 0
            pct = (cnt / n * 100) if n > 0 else 0.0
            rows.append({'subject': s, 'category': cat, 'count': int(cnt), 'percent_share': round(pct, 1)})
    agg = pd.DataFrame(rows)
    if agg.empty: return agg
    order = agg.groupby('subject')['percent_share'].sum().sort_values(ascending=False).index.tolist()
    agg['subject'] = pd.Categorical(agg['subject'], categories=order, ordered=True)
    return agg.sort_values('subject')

def subject_completion_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    out = df.groupby('subject', dropna=True).agg(
        متوسط_النسبة=('solve_pct', lambda s: round(float(np.nanmean(s)) if len(s)>0 else 0.0, 1)),
        عدد_الطلاب=('student_name','nunique')
    ).reset_index().rename(columns={'subject':'المادة'})
    return out.sort_values('متوسط_النسبة', ascending=False)

def section_completion_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    per_student = df.groupby(['student_name','section'], as_index=False)\
                    .agg(النسبة=('solve_pct', lambda s: float(np.nanmean(s)) if len(s)>0 else 0.0))
    out = per_student.groupby('section', as_index=False)\
                     .agg(متوسط_الشعبة=('النسبة', lambda s: round(float(np.nanmean(s)),1)))
    out = out.rename(columns={'section':'الشعبة'}).sort_values('متوسط_الشعبة', ascending=False)
    return out

# ============== رسوم بيانية ==============
def chart_stacked_by_subject(agg_df: pd.DataFrame, mode='percent') -> go.Figure:
    fig = go.Figure(); colors = [CATEGORY_COLORS[c] for c in CATEGORY_ORDER]
    for i, cat in enumerate(CATEGORY_ORDER):
        d = agg_df[agg_df['category'] == cat]
        vals = d['percent_share'] if mode == 'percent' else d['count']
        text = [(f"{v:.1f}%" if mode == 'percent' else str(int(v))) if v > 0 else "" for v in vals]
        hover = "<b>%{y}</b><br>الفئة: " + cat + "<br>" + ("النسبة: %{x:.1f}%<extra></extra>" if mode=='percent' else "العدد: %{x}<extra></extra>")
        fig.add_trace(go.Bar(
            name=cat, x=vals, y=d['subject'], orientation='h',
            marker=dict(color=colors[i], line=dict(color='white', width=1)),
            text=text, textposition='inside', textfont=dict(size=11, family='Cairo'),
            hovertemplate=hover
        ))
    fig.update_layout(
        title=dict(text="توزيع الفئات حسب المادة", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        xaxis=dict(title="النسبة المئوية (%)" if mode=='percent' else "عدد الطلاب", gridcolor='#E5E7EB', range=[0,100] if mode=='percent' else None),
        yaxis=dict(title="المادة", autorange='reversed'),
        barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(family='Cairo'),
        height=480
    )
    return fig

def chart_subject_completion_bar(tbl: pd.DataFrame) -> go.Figure:
    if tbl.empty: return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=tbl['المادة'], x=tbl['متوسط_النسبة'], orientation='h',
        text=[f"{v:.1f}%" for v in tbl['متوسط_النسبة']], textposition='inside',
        marker=dict(color='#8A1538', line=dict(color='white', width=1))
    ))
    fig.update_layout(
        title=dict(text="متوسط نسبة الحل لكل مادة", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        xaxis=dict(range=[0,100], title="النسبة (%)", gridcolor='#E5E7EB'),
        yaxis=dict(autorange='reversed', title="المادة"),
        plot_bgcolor='white', paper_bgcolor='white', font=dict(family='Cairo'),
        height=max(380, len(tbl)*36)
    )
    return fig

def chart_section_avg_bar(df_section: pd.DataFrame) -> go.Figure:
    if df_section.empty: return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_section['الشعبة'], y=df_section['متوسط_الشعبة'],
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

# ============== تقارير وصفية/كمية ==============
def _strengths_weaknesses(avg: float) -> Tuple[str, str]:
    if avg >= 85: return "التزام مرتفع – إنجاز مستقر", "ضرورة تثبيت المستوى والمتابعة"
    if avg >= 70: return "تفاعل جيد – إمكانات قوية", "تذبذب في بعض المهام"
    if avg >= 50: return "تحسن ملحوظ عند المتابعة", "ضعف الالتزام بالمواعيد"
    return "وجود نماذج متميزة فردية", "انخفاض عام في الإنجاز"

def recommendations_pack() -> List[str]:
    return [
        "تصميم مسابقات تحفيزية قصيرة داخل الحصة.",
        "فتح نظام قطر داخل الصف ومتابعة الحل المباشر.",
        "التواصل مع ولي الأمر للطلاب منخفضي الإنجاز.",
        "تخصيص حصص دعم للمفاهيم الحرجة.",
    ]

def report_per_subject(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return pd.DataFrame()
    g = df.groupby('subject', as_index=False).agg(
        متوسط=('solve_pct', lambda s: round(float(np.nanmean(s)),1)),
        طلاب=('student_name','nunique')
    )
    g['قوة'], g['ضعف'] = zip(*g['متوسط'].map(_strengths_weaknesses))
    g['توصيات'] = " • " + " | ".join(recommendations_pack())
    return g.rename(columns={'subject':'المادة'})

def report_per_section_subject(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return pd.DataFrame()
    g = df.groupby(['section','subject'], as_index=False).agg(
        متوسط=('solve_pct', lambda s: round(float(np.nanmean(s)),1)),
        طلاب=('student_name','nunique')
    ).rename(columns={'section':'الشعبة','subject':'المادة'})
    g['قوة'], g['ضعف'] = zip(*g['متوسط'].map(_strengths_weaknesses))
    g['توصيات'] = " • " + " | ".join(recommendations_pack())
    return g.sort_values(['الشعبة','متوسط'], ascending=[True, False]).reset_index(drop=True)

# ============== توصية تلقائية موحّدة ==============
def get_auto_reco(category: str, student_name: str = "") -> str:
    mapping = {
        'بلاتيني 🥇': "أداء متميز جدًا. استمر بهذا النسق.",
        'ذهبي 🥈': "أداء قوي وقابل للتحسن نحو البلاتيني.",
        'فضي 🥉': "أداء جيد ويحتاج مزيد التثبيت.",
        'برونزي': "نوصي بزيادة الالتزام والمتابعة.",
        'بحاجة لتحسين': "يلزم خطة متابعة مكثفة.",
        'لا يستفيد': "ابدأ فورًا بتفعيل النظام؛ نحن هنا للمساعدة."
    }
    return mapping.get(category, "نوصي بالمتابعة المستمرة.")

# ============== PDF فردي ==============
@safe_execute(default_return=b"", error_message="خطأ في إنشاء PDF")
def make_student_pdf_fpdf(
    school_name: str, student_name: str, grade: str, section: str,
    table_df: pd.DataFrame, overall_avg: float, reco_text: str,
    coordinator_name: str, academic_deputy: str, admin_deputy: str, principal_name: str,
    font_info: Tuple[str, Optional[str]], logo_path: Optional[str] = None,
) -> bytes:
    font_name, font_path = font_info
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    if font_path and os.path.exists(font_path):
        try: pdf.add_font(font_name, "", font_path, uni=True)
        except Exception as e: logger.warning(f"خط: {e}"); font_name = ""

    def set_font(size=12, color=(0,0,0)):
        if font_name: pdf.set_font(font_name, size=size)
        else: pdf.set_font("Helvetica", size=size)
        pdf.set_text_color(*color)

    pdf.set_fill_color(*QATAR_MAROON); pdf.rect(0, 0, 210, 22, style="F")
    if logo_path and os.path.exists(logo_path):
        try: pdf.image(logo_path, x=185, y=2.5, w=20)
        except Exception as e: logger.warning(f"شعار: {e}")

    set_font(14, (255,255,255)); pdf.set_xy(10,6)
    pdf.cell(0, 7, rtl("إنجاز - تحليل التقييمات الأسبوعية على نظام قطر للتعليم"), align="R")
    set_font(11, (234, 218, 169)); pdf.set_xy(10, 13)
    pdf.cell(0, 7, rtl("نحو تنمية رقمية مستدامة"), align="R")

    set_font(12); pdf.set_y(28)
    pdf.cell(0,8, rtl(f"اسم المدرسة: {school_name or '—'}"), ln=1, align="R")
    pdf.cell(0,8, rtl(f"اسم الطالب: {student_name}"), ln=1, align="R")
    pdf.cell(0,8, rtl(f"الصف: {grade or '—'}     الشعبة: {section or '—'}"), ln=1, align="R"); pdf.ln(2)

    headers = [rtl("المادة"), rtl("عدد التقييمات الإجمالي"), rtl("عدد التقييمات المنجزة"), rtl("عدد التقييمات المتبقية")]
    widths  = [76, 38, 38, 38]
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
    set_font(11)
    for line in (reco_text or "—").splitlines(): pdf.multi_cell(0, 7, rtl(line), align="R")

    pdf.set_y(-18); set_font(9, (90,90,90))
    pdf.cell(0, 8, rtl("إعداد الرسوم البيانية في الصور حسب توقعات مبدئية | تطوير وتنفيذ: سحر عثمان"), 0, 0, 'C')

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
            help="يُقرأ التاريخ من H3 لكل عمود. الأعمدة خارج النطاق تُستبعد.")
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
        # 📚 بيانات المعلمات
        st.subheader("📚 بيانات المعلمات (الشُعب والبريد)")
        teacher_file = st.file_uploader("ارفع ملف المعلمات (Excel/CSV)",
                                        type=["xlsx","xls","csv"], key="teacher_file")
        teachers_df = None
        if teacher_file:
            try:
                if teacher_file.name.endswith(".csv"):
                    teachers_df = pd.read_csv(teacher_file)
                else:
                    teachers_df = pd.read_excel(teacher_file)
                # توحيد أسماء الأعمدة الشائعة
                cols_map = {c.strip(): c for c in teachers_df.columns}
                def find_col(options):
                    for k in cols_map:
                        base = _strip_invisible_and_diacritics(k)
                        if base in options: return cols_map[k]
                    return None
                sec_col = find_col(["الشعبة","شعبة","Section","section"])
                name_col= find_col(["اسم المعلمة","المعلمة","Teacher","teacher","اسم المعلم","المعلم"])
                mail_col= find_col(["البريد الإلكتروني","البريد الالكتروني","email","Email","البريد"])
                if not (sec_col and name_col and mail_col):
                    st.error("❌ يجب أن يحتوي الملف على أعمدة: الشعبة، اسم المعلمة، البريد الإلكتروني")
                else:
                    teachers_df = teachers_df[[sec_col, name_col, mail_col]].rename(
                        columns={sec_col:"الشعبة", name_col:"اسم المعلمة", mail_col:"البريد الإلكتروني"}
                    )
                    teachers_df["الشعبة"] = teachers_df["الشعبة"].astype(str).apply(_strip_invisible_and_diacritics).map(_normalize_arabic_digits).str.strip()
                    st.session_state.teachers_df = teachers_df.copy()
                    st.success("✅ تم تحميل بيانات المعلمات")
                    st.dataframe(teachers_df, use_container_width=True, height=200)
            except Exception as e:
                st.error(f"❌ خطأ في قراءة ملف المعلمات: {e}")

        run_analysis = st.button("▶️ تشغيل التحليل", use_container_width=True, type="primary", disabled=not uploaded_files)

    # === Main content ===
    if not uploaded_files:
        st.info("📤 ارفع ملفات Excel من الشريط الجانبي للبدء")

    elif run_analysis:
        sheets_to_use = st.session_state.selected_sheets or []
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

        # فلاتر جدول
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
        if selected_level  != 'الكل': filtered_pivot = filtered_pivot[filtered_pivot['الصف']==selected_level]
        if selected_section!= 'الكل': filtered_pivot = filtered_pivot[filtered_pivot['الشعبة']==selected_section]
        if selected_category!= 'الكل': filtered_pivot = filtered_pivot[filtered_pivot['الفئة']==selected_category]

        st.dataframe(filtered_pivot, use_container_width=True, height=440)
        csv = filtered_pivot.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 تحميل الجدول (CSV)", csv,
                           f"ingaz_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                           "text/csv", key='download-csv')

        st.divider()

        # 🍩 توزيع الفئات العام + مؤشر
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">🍩 التوزيع العام للفئات</h2>', unsafe_allow_html=True)
        if 'الفئة' in filtered_pivot.columns and not filtered_pivot.empty:
            counts = filtered_pivot['الفئة'].value_counts().reindex(CATEGORY_ORDER, fill_value=0)
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
                   'threshold':{'line':{'color':'#C9A646','width':4},'thickness':0.75,'value':80}}))
        fig_g.update_layout(title=dict(text="متوسط الإنجاز العام", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
                            paper_bgcolor='white', plot_bgcolor='white', height=350, font=dict(family='Cairo'))
        st.plotly_chart(fig_g, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ✅ توزيع الفئات حسب المادة
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">📊 توزيع الفئات حسب المادة الدراسية</h2>', unsafe_allow_html=True)
        try:
            normalized = normalize_dataframe(df.copy())
            if selected_level  != 'الكل': normalized = normalized[normalized['level'].astype(str)==str(selected_level)]
            if selected_section!= 'الكل': normalized = normalized[normalized['section'].astype(str)==str(selected_section)]
            agg_df = aggregate_by_subject(normalized)
            if not agg_df.empty:
                st.plotly_chart(chart_stacked_by_subject(agg_df, mode='percent'), use_container_width=True)
            else:
                st.info("لا توجد بيانات كافية لعرض الرسم.")
        except Exception as e:
            st.error(f"خطأ في الرسم: {e}")
        st.markdown('</div>', unsafe_allow_html=True)

        # 📚 نسبة الحل العامة لكل مادة
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">📚 نسبة الحل العامة لكل مادة</h2>', unsafe_allow_html=True)
        df_filtered = df.copy()
        if selected_level  != 'الكل': df_filtered = df_filtered[df_filtered['level'].astype(str)==str(selected_level)]
        if selected_section!= 'الكل': df_filtered = df_filtered[df_filtered['section'].astype(str)==str(selected_section)]
        sub_tbl = subject_completion_summary(df_filtered)
        if not sub_tbl.empty:
            st.dataframe(sub_tbl, use_container_width=True, height=300)
            st.plotly_chart(chart_subject_completion_bar(sub_tbl), use_container_width=True)
        else:
            st.info("لا توجد بيانات كافية لحساب متوسطات المواد.")
        st.markdown('</div>', unsafe_allow_html=True)

        # 🏷️ متوسط الإنجاز على مستوى الشعب
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">🏷️ متوسط الإنجاز على مستوى الشعب</h2>', unsafe_allow_html=True)
        df_for_sections = df.copy()
        if selected_level  != 'الكل': df_for_sections = df_for_sections[df_for_sections['level'].astype(str)==str(selected_level)]
        sec_tbl = section_completion_summary(df_for_sections)
        if not sec_tbl.empty:
            st.dataframe(sec_tbl, use_container_width=True, height=260)
            st.plotly_chart(chart_section_avg_bar(sec_tbl), use_container_width=True)
        else:
            st.info("لا توجد بيانات كافية لعرض متوسطات الشعب.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.divider()

        # ================== التقارير الوصفية/الكمية ==================
        st.subheader("📝 تقارير وصفية وكمية")
        colr1, colr2 = st.columns(2)
        with colr1:
            st.markdown("#### تقرير مُكي/وصفي لكل مادة")
            rep_sub = report_per_subject(df_filtered)
            if not rep_sub.empty:
                st.dataframe(rep_sub, use_container_width=True, height=360)
                st.download_button(
                    "📥 تصدير تقرير المواد (CSV)",
                    rep_sub.to_csv(index=False).encode('utf-8-sig'),
                    file_name=f"report_subjects_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("لا توجد بيانات لتقرير المواد.")
        with colr2:
            st.markdown("#### تقرير كمي/وصفي لكل شعبة × مادة")
            df_secsub = df.copy()
            if selected_level  != 'الكل': df_secsub = df_secsub[df_secsub['level'].astype(str)==str(selected_level)]
            rep_secsub = report_per_section_subject(df_secsub)
            if selected_section != 'الكل':
                rep_secsub = rep_secsub[rep_secsub['الشعبة'].astype(str)==str(selected_section)]
            if not rep_secsub.empty:
                st.dataframe(rep_secsub, use_container_width=True, height=360)
                st.download_button(
                    "📥 تصدير تقرير الشُعب×المواد (CSV)",
                    rep_secsub.to_csv(index=False).encode('utf-8-sig'),
                    file_name=f"report_sections_subjects_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("لا توجد بيانات لتقرير الشُعب × المواد.")

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
                reco = st.text_area("توصية منسق المشاريع", value=get_auto_reco(student_category, sel), height=180)

            sdata = df[(df['student_name'].str.strip()==sel.strip())].copy()
            if not sdata.empty:
                table = sdata[['subject','total_count','completed_count']].rename(
                    columns={'subject':'المادة','total_count':'إجمالي','completed_count':'منجز'}
                )
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
                                t = sd[['subject','total_count','completed_count']].rename(
                                    columns={'subject':'المادة','total_count':'إجمالي','completed_count':'منجز'}
                                )
                                t['متبقي'] = (t['إجمالي'] - t['منجز']).clip(lower=0).astype(int)
                                av = float(sd['solve_pct'].mean()) if 'solve_pct' in sd.columns else 0.0
                                if pd.isna(av): av = 0.0
                                if use_auto_all:
                                    cat = r['الفئة'].iloc[0] if not r.empty and 'الفئة' in r.columns else ''
                                    rtext = get_auto_reco(cat, stu)
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

        st.divider()

        # ================== 💌 مراسلة المعلمات بضغطة ==================
        st.subheader("💌 مراسلة المعلمات للطلاب «لا يستفيد»")
        teachers_df = st.session_state.get("teachers_df")
        if teachers_df is None or teachers_df.empty:
            st.info("📥 ارفعي ملف المعلمات (الشعبة، اسم المعلمة، البريد الإلكتروني) من الشريط الجانبي لتمكين هذه الميزة.")
        else:
            # جدول الطلاب لا يستفيد بحسب الفلاتر الحالية
            non_benefit = filtered_pivot.copy()
            non_benefit = non_benefit[non_benefit['الفئة'] == 'لا يستفيد']
            if non_benefit.empty:
                st.success("لا يوجد طلاب ضمن فئة «لا يستفيد» في نطاق الفلاتر الحالية 👌")
            else:
                # ربط بالشُعب → المعلمات
                tmp = non_benefit[['الطالب','الصف','الشعبة']].copy()
                tmp['الشعبة'] = tmp['الشعبة'].astype(str).apply(_strip_invisible_and_diacritics).map(_normalize_arabic_digits).str.strip()
                tdf = teachers_df.copy()
                tdf['الشعبة'] = tdf['الشعبة'].astype(str).apply(_strip_invisible_and_diacritics).map(_normalize_arabic_digits).str.strip()
                joined = tmp.merge(tdf, on='الشعبة', how='left')

                # توصية عامة قابلة للتعديل لكل المعلمات
                default_reco = "يرجى فتح نظام قطر داخل الصف والمتابعة الفورية، والتواصل مع ولي الأمر، وتخصيص خطة دعم قصيرة المدى."
                global_reco = st.text_area("✍️ نص التوصية العامة (يمكن التعديل قبل الإرسال):", value=default_reco, height=100)

                st.markdown("#### قائمة المعلمات والشُعب")
                # تجميع حسب المعلمة/الشعبة
                groups = joined.groupby(['اسم المعلمة','البريد الإلكتروني','الشعبة'])
                for (t_name, t_email, sect), gdf in groups:
                    students_list = sorted(gdf['الطالب'].dropna().astype(str).unique().tolist())
                    count = len(students_list)
                    if not t_email or pd.isna(t_email):
                        st.warning(f"⚠️ الشعبة {sect}: لا يوجد بريد للمعلمة {t_name or '—'}")
                        continue

                    # موضوع ونص الرسالة
                    subject = f"متابعة طلاب الشعبة {sect} - فئة لا يستفيد"
                    body_lines = [
                        f"السلام عليكم أستاذة {t_name or ''},",
                        "",
                        "أرفق لحضرتك قائمة الطلاب الذين لم يستفيدوا بعد من نظام قطر للتعليم:",
                        *[f"- {s}" for s in students_list],
                        "",
                        "التوصية:",
                        global_reco.strip(),
                        "",
                        f"عدد الطلاب: {count}",
                        "",
                        "شاكرين تعاونك."
                    ]
                    body = "\n".join(body_lines)

                    # رابط mailto (مُشفَّر)
                    mailto_link = f"mailto:{t_email}?subject={quote(subject)}&body={quote(body)}"

                    # عرض البطاقة
                    st.markdown(
                        f"""
                        <div class="chart-container" style="padding:14px">
                          <div style="display:flex;justify-content:space-between;gap:8px;align-items:center;">
                            <div>
                              <div style="font-weight:800;color:#8A1538;">الشعبة: {sect}</div>
                              <div>المعلمة: <b>{t_name or '—'}</b></div>
                              <div>عدد الطلاب: <b>{count}</b></div>
                            </div>
                            <div>
                              <a href="{mailto_link}" style="background:#8A1538;color:white;padding:10px 14px;border-radius:10px;text-decoration:none;font-weight:700">📧 أرسلي البريد الآن</a>
                            </div>
                          </div>
                          <div style="margin-top:8px;font-size:13px">
                            <b>أسماء الطلاب:</b> {', '.join(students_list)}
                          </div>
                          <div style="margin-top:8px;font-size:13px">
                            <b>التوصية الحالية:</b> {global_reco}
                          </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

    # Footer
    st.markdown(f"""
    <div style="margin-top:22px;background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
                color:#fff;border-radius:12px;padding:12px 10px;text-align:center;">
      <div style="font-weight:800;font-size:15px;margin:2px 0 4px">مدرسة عثمان بن عفان النموذجية للبنين</div>
      <div style="font-weight:700;font-size:12px;margin:0 0 4px;opacity:.95">© {datetime.now().year} جميع الحقوق محفوظة</div>
      <div style="font-weight:600;font-size:12px;opacity:.95">إعداد الرسوم البيانية في الصور حسب توقعات مبدئية | تطوير وتنفيذ: سحر عثمان</div>
    </div>
    """, unsafe_allow_html=True)

# ============== تشغيل ==============
if __name__ == "__main__":
    main()

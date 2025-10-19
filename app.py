import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import zipfile
import io
import base64
from datetime import datetime
from typing import Tuple, Dict, List
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="محلل التقييمات الأسبوعية",
    page_icon="https://i.imgur.com/XLef7tS.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700&display=swap');
    
    * { 
        font-family: 'Cairo', 'Segoe UI', -apple-system, sans-serif;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
    }
    
    /* Main Background */
    .main { background: #FFFFFF; }
    body { background: #FFFFFF; }
    .stApp { background: #FFFFFF; }
    
    /* Header Container - Enhanced Branding */
    .header-container {
        background: linear-gradient(135deg, #8A1538 0%, #6B1029 100%);
        padding: 56px 48px;
        border-radius: 0;
        color: #FFFFFF;
        text-align: center;
        margin-bottom: 40px;
        box-shadow: 0 6px 20px rgba(138, 21, 56, 0.25);
        border-bottom: 4px solid #C9A646;
        position: relative;
    }
    
    .header-container::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: linear-gradient(90deg, #C9A646 0%, #E8D4A0 50%, #C9A646 100%);
    }
    
    .header-container h1 { 
        margin: 0 0 20px 0;
        font-size: 40px;
        font-weight: 700;
        line-height: 1.25;
        color: #FFFFFF !important;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        letter-spacing: -0.01em;
    }
    
    .header-container .subtitle { 
        font-size: 18px;
        font-weight: 600;
        opacity: 1;
        margin: 0 0 16px 0;
        color: #FFFFFF !important;
        letter-spacing: 0.01em;
    }
    
    .header-container .accent-line {
        font-size: 15px;
        color: #C9A646;
        font-weight: 600;
        margin: 0 0 14px 0;
        letter-spacing: 0.03em;
        text-transform: uppercase;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.2);
    }
    
    .header-container .description {
        font-size: 14px;
        opacity: 0.95;
        margin: 0;
        color: #FFFFFF !important;
        font-weight: 400;
        letter-spacing: 0.01em;
    }
    
    /* Sidebar - Solid Qatar Maroon */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #8A1538 0%, #6B1029 100%) !important;
        border-right: 2px solid #C9A646;
        box-shadow: 4px 0 16px rgba(0, 0, 0, 0.15);
    }
    
    [data-testid="stSidebar"] * { 
        color: white !important;
    }
    
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        font-weight: 600;
        color: #FFFFFF !important;
    }
    
    [data-testid="stSidebar"] hr {
        border-color: rgba(201, 166, 70, 0.3) !important;
        margin: 20px 0 !important;
    }
    
    /* Section Box - Clean & Minimal */
    .section-box {
        background: #F5F5F5;
        padding: 28px;
        border-radius: 12px;
        margin: 28px 0;
        border-right: 6px solid #8A1538;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    }
    
    /* Metric Box - Modern Card Style */
    .metric-box {
        background: #FFFFFF;
        border: 2px solid #E8E8E8;
        border-right: 5px solid #8A1538;
        padding: 24px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 3px 12px rgba(0, 0, 0, 0.08);
        transition: all 0.3s ease;
    }
    
    .metric-box:hover {
        border-right-color: #C9A646;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.12);
        transform: translateY(-3px);
    }
    
    .metric-value {
        font-size: 40px;
        font-weight: 700;
        color: #8A1538;
        line-height: 1.1;
        margin-bottom: 8px;
    }
    
    .metric-label {
        font-size: 13px;
        font-weight: 600;
        color: #4A4A4A;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    
    /* Buttons - Clean Qatar Maroon */
    .stButton > button {
        background: linear-gradient(135deg, #8A1538 0%, #6B1029 100%) !important;
        color: white !important;
        border: none !important;
        padding: 14px 28px !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(138, 21, 56, 0.25) !important;
        letter-spacing: 0.02em !important;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #6B1029 0%, #8A1538 100%) !important;
        box-shadow: 0 6px 20px rgba(138, 21, 56, 0.35) !important;
        transform: translateY(-2px) !important;
    }
    
    /* Divider */
    hr { 
        border-color: #E8E8E8 !important;
        margin: 32px 0 !important;
        border-width: 1px !important;
    }
    
    /* Typography */
    h1, h2, h3, h4, h5, h6 { 
        color: #FFFFFF;
        font-weight: 600;
    }
    
    /* Main content headers */
    .main h1, .main h2, .main h3, .main h4, .main h5, .main h6 {
        color: #8A1538;
    }
    
    h1 { font-size: 36px; line-height: 1.3; margin-bottom: 16px; font-weight: 700; }
    h2 { font-size: 28px; line-height: 1.35; margin-bottom: 20px; font-weight: 600; }
    h3 { font-size: 22px; line-height: 1.4; margin-bottom: 16px; font-weight: 600; }
    h4 { font-size: 18px; line-height: 1.5; margin-bottom: 12px; font-weight: 600; }
    
    p, div, span {
        font-size: 15px;
        line-height: 1.7;
        color: #2C2C2C;
    }
    
    /* Logo Containers */
    .logo-header-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 24px 48px;
        background: #F5F5F5;
        border-bottom: 3px solid #8A1538;
        margin-bottom: 0;
    }
    
    .logo-left {
        display: flex;
        align-items: center;
        padding: 12px;
    }
    
    .logo-right-group {
        display: flex;
        gap: 24px;
        align-items: center;
        padding: 12px;
    }
    
    .logo-sidebar-container {
        text-align: center;
        padding: 28px 24px;
        margin-bottom: 28px;
        border-bottom: 2px solid rgba(201, 166, 70, 0.3);
    }
    
    .logo-footer-container {
        text-align: center;
        padding: 24px;
        margin: 0 auto 20px;
    }
    
    /* Download Buttons */
    .stDownloadButton > button {
        background: #FFFFFF !important;
        color: #8A1538 !important;
        border: 2px solid #8A1538 !important;
        padding: 12px 26px !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.3s ease !important;
    }
    
    .stDownloadButton > button:hover {
        background: #8A1538 !important;
        color: #FFFFFF !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 12px rgba(138, 21, 56, 0.25) !important;
    }
    
    /* File Uploader Button - Browse Files */
    [data-testid="stFileUploader"] label {
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }
    
    [data-testid="stFileUploader"] button {
        background: rgba(255, 255, 255, 0.15) !important;
        color: #FFFFFF !important;
        border: 2px solid rgba(255, 255, 255, 0.3) !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    
    [data-testid="stFileUploader"] button:hover {
        background: rgba(255, 255, 255, 0.25) !important;
        border-color: #C9A646 !important;
    }
    
    [data-testid="stFileUploader"] section {
        border-color: rgba(255, 255, 255, 0.3) !important;
    }
    
    [data-testid="stFileUploader"] small {
        color: #FFFFFF !important;
        opacity: 0.9;
    }
    
    /* Dataframe Styling */
    [data-testid="stDataFrame"] {
        border: 2px solid #E8E8E8;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
    }
    
    /* Metrics Enhancement */
    [data-testid="stMetricValue"] {
        font-size: 40px !important;
        font-weight: 700 !important;
        color: #8A1538 !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 13px !important;
        font-weight: 600 !important;
        color: #4A4A4A !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
    }
    
    /* Success/Info Messages */
    .stSuccess {
        background-color: rgba(138, 21, 56, 0.1) !important;
        color: #8A1538 !important;
        border-right: 4px solid #8A1538 !important;
    }
    
    .stInfo {
        background-color: rgba(201, 166, 70, 0.1) !important;
        color: #4A4A4A !important;
        border-right: 4px solid #C9A646 !important;
    }
    
    /* Chart Container */
    .chart-container {
        background: white;
        border: 2px solid #E5E7EB;
        border-right: 5px solid #8A1538;
        border-radius: 12px;
        padding: 24px;
        margin: 20px 0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    }
    
    .chart-title {
        font-size: 24px;
        font-weight: 700;
        color: #8A1538;
        text-align: center;
        margin-bottom: 16px;
        font-family: 'Cairo', sans-serif;
    }
    
    .info-box {
        background: #F5F5F5;
        border-right: 4px solid #C9A646;
        padding: 16px;
        border-radius: 8px;
        margin: 16px 0;
        font-family: 'Cairo', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# CONSTANTS FOR CHARTS
# ============================================

CATEGORY_THRESHOLDS = {
    'بلاتيني 🥇': (90, 100),
    'ذهبي 🥈': (80, 89.99),
    'فضي 🥉': (70, 79.99),
    'برونزي': (60, 69.99),
    'بحاجة لتحسين': (0, 59.99)
}

CATEGORY_COLORS = {
    'بلاتيني 🥇': '#E5E4E2',
    'ذهبي 🥈': '#C9A646',
    'فضي 🥉': '#C0C0C0',
    'برونزي': '#CD7F32',
    'بحاجة لتحسين': '#8A1538'
}

CATEGORY_ORDER = ['بلاتيني 🥇', 'ذهبي 🥈', 'فضي 🥉', 'برونزي', 'بحاجة لتحسين']

SUFFIX_KEYWORDS = {
    'total': ['إجمالي', 'اجمالي', 'Total'],
    'solved': ['منجز', 'Solved', 'Completed'],
    'remaining': ['متبقي', 'متبقّي', 'Remaining'],
    'percent': ['النسبة', 'نسبة', 'Percent', '%']
}

# ============================================
# HELPER FUNCTIONS
# ============================================

def parse_sheet_name(sheet_name):
    """
    تحليل اسم الورقة لاستخراج المادة والمستوى والشعبة
    التنسيق المتوقع: "اسم المادة المستوى الشعبة"
    مثال: "التربية الاسلامية 01 1"
    """
    try:
        parts = sheet_name.strip().split()
        
        if len(parts) < 3:
            return sheet_name.strip(), "", ""
        
        section = parts[-1]
        level = parts[-2]
        subject_parts = parts[:-2]
        subject = " ".join(subject_parts)
        
        if not (level.isdigit() or (level.startswith('0') and len(level) <= 2)):
            subject = " ".join(parts[:-1])
            level = parts[-1]
            section = ""
        
        logger.info(f"تحليل الورقة: '{sheet_name}' → المادة: '{subject}', المستوى: '{level}', الشعبة: '{section}'")
        
        return subject, level, section
        
    except Exception as e:
        logger.error(f"خطأ في تحليل اسم الورقة '{sheet_name}': {str(e)}")
        return sheet_name, "", ""

@st.cache_data
def analyze_excel_file(file, sheet_name):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        
        due_dates = []
        try:
            for col_idx in range(7, min(df.shape[1], 20)):
                cell_value = df.iloc[1, col_idx]
                if pd.notna(cell_value):
                    try:
                        due_date = pd.to_datetime(cell_value)
                        if 2000 <= due_date.year <= 2100:
                            due_dates.append(due_date.date())
                    except (ValueError, TypeError):
                        continue
        except (IndexError, KeyError):
            pass
        
        level = level_from_name
        section = section_from_name
        
        assessment_columns = []
        for col_idx in range(7, df.shape[1]):
            title = df.iloc[0, col_idx] if col_idx < df.shape[1] else None
            
            if pd.isna(title):
                break
            
            all_dash = True
            for row_idx in range(4, min(len(df), 20)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    if cell_str not in ['-', '—', '']:
                        all_dash = False
                        break
            
            if not all_dash:
                assessment_columns.append({
                    'index': col_idx,
                    'title': str(title).strip() if pd.notna(title) else f"تقييم {len(assessment_columns) + 1}"
                })
        
        total_assessments = len(assessment_columns)
        
        if total_assessments == 0:
            st.warning(f"⚠️ لم يتم العثور على تقييمات في ورقة: {sheet_name}")
            return []
        
        assessment_titles = [col['title'] for col in assessment_columns]
        
        results = []
        
        try:
            for idx in range(4, len(df)):
                student_name = df.iloc[idx, 0]
                
                if pd.isna(student_name) or str(student_name).strip() == "":
                    continue
                
                student_name_clean = " ".join(str(student_name).strip().split())
                
                completed_count = 0
                pending_titles = []
                
                for i, col_info in enumerate(assessment_columns):
                    col_idx = col_info['index']
                    
                    if col_idx >= df.shape[1]:
                        pending_titles.append(col_info['title'])
                        continue
                    
                    cell_value = df.iloc[idx, col_idx]
                    
                    is_completed = False
                    
                    if pd.isna(cell_value):
                        is_completed = False
                    else:
                        cell_str = str(cell_value).strip().upper()
                        
                        if cell_str in ['M', 'I', 'AB', 'X', '-', '—', '', 'NAN', 'NONE']:
                            is_completed = False
                        else:
                            try:
                                num_value = float(cell_str.replace(',', '.'))
                                if num_value > 0:
                                    is_completed = True
                                else:
                                    is_completed = False
                            except (ValueError, TypeError):
                                if len(cell_str) > 0 and cell_str not in ['M', 'I', 'AB', 'X', '-', '—']:
                                    is_completed = True
                    
                    if is_completed:
                        completed_count += 1
                    else:
                        pending_titles.append(col_info['title'])
                
                solve_pct = (completed_count / total_assessments * 100) if total_assessments > 0 else 0.0
                
                results.append({
                    "student_name": student_name_clean,
                    "subject": subject,
                    "level": str(level).strip(),
                    "section": str(section).strip(),
                    "solve_pct": round(solve_pct, 1),
                    "completed_count": completed_count,
                    "total_count": total_assessments,
                    "pending_titles": ", ".join(pending_titles) if pending_titles else "-",
                    "due_dates": due_dates
                })
        except (IndexError, KeyError) as e:
            logger.error(f"خطأ في قراءة البيانات: {str(e)}")
        
        logger.info(f"✅ تم تحليل {len(results)} طالب من ورقة {sheet_name} - إجمالي {total_assessments} تقييم")
        return results
        
    except Exception as e:
        logger.error(f"❌ خطأ في analyze_excel_file: {str(e)}")
        st.error(f"❌ خطأ في تحليل الملف: {str(e)}")
        return []

@st.cache_data
def create_pivot_table(df):
    try:
        if df.empty:
            st.warning("⚠️ لا توجد بيانات للتحليل")
            return pd.DataFrame()
        
        logger.info(f"Columns in dataframe: {df.columns.tolist()}")
        
        df_clean = df.copy()
        
        required_cols = ['student_name', 'level', 'section', 'subject', 'total_count', 'completed_count', 'solve_pct']
        missing_cols = [col for col in required_cols if col not in df_clean.columns]
        if missing_cols:
            st.error(f"❌ أعمدة مفقودة: {', '.join(missing_cols)}")
            return pd.DataFrame()
        
        df_clean = df_clean.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'], keep='first')
        
        unique_students = df_clean[['student_name', 'level', 'section']].drop_duplicates()
        unique_students = unique_students.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
        result = unique_students.copy()
        
        subjects = sorted(df_clean['subject'].unique())
        
        for subject in subjects:
            subject_data = df_clean[df_clean['subject'] == subject].copy()
            
            subject_df = subject_data[['student_name', 'level', 'section', 'total_count', 'completed_count', 'solve_pct']].copy()
            
            subject_df['total_count'] = subject_df['total_count'].fillna(0)
            subject_df['completed_count'] = subject_df['completed_count'].fillna(0)
            subject_df['solve_pct'] = subject_df['solve_pct'].fillna(0)
            
            subject_df = subject_df.rename(columns={
                'total_count': f"{subject} - إجمالي",
                'completed_count': f"{subject} - منجز",
                'solve_pct': f"{subject} - النسبة"
            })
            
            subject_df = subject_df.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')
            
            result = result.merge(subject_df, on=['student_name', 'level', 'section'], how='left')
            
            pending_data = subject_data[['student_name', 'level', 'section', 'pending_titles']].copy()
            pending_data = pending_data.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')
            pending_data = pending_data.rename(columns={'pending_titles': f"{subject} - متبقي"})
            result = result.merge(pending_data, on=['student_name', 'level', 'section'], how='left')
        
        pct_cols = [col for col in result.columns if 'النسبة' in col]
        if pct_cols:
            result['المتوسط'] = result[pct_cols].mean(axis=1, skipna=True)
            result['المتوسط'] = result['المتوسط'].fillna(0)
            
            def categorize(pct):
                if pd.isna(pct) or pct == 0:
                    return "لا يستفيد 🚫"
                elif pct >= 90:
                    return "بلاتينية 🥇"
                elif pct >= 80:
                    return "ذهبي 🥈"
                elif pct >= 70:
                    return "فضي 🥉"
                elif pct >= 60:
                    return "برونزي"
                else:
                    return "يحتاج تحسين"
            
            result['الفئة'] = result['المتوسط'].apply(categorize)
        
        result = result.rename(columns={
            'student_name': 'الطالب',
            'level': 'الصف',
            'section': 'الشعبة'
        })
        
        for col in result.columns:
            if 'إجمالي' in col or 'منجز' in col:
                result[col] = result[col].fillna(0).astype(int)
            elif 'النسبة' in col or col == 'المتوسط':
                result[col] = result[col].fillna(0).round(1)
            elif 'متبقي' in col:
                result[col] = result[col].fillna('-')
        
        result = result.drop_duplicates(subset=['الطالب', 'الصف', 'الشعبة'], keep='first')
        
        return result.reset_index(drop=True)
        
    except Exception as e:
        logger.error(f"خطأ في create_pivot_table: {str(e)}")
        st.error(f"❌ خطأ في معالجة البيانات: {str(e)}")
        return pd.DataFrame()

# ============================================
# CHART FUNCTIONS
# ============================================

def assign_category(percent: float) -> str:
    """تعيين الفئة بناءً على النسبة المئوية"""
    if pd.isna(percent):
        return 'بحاجة لتحسين'
    
    for category, (min_val, max_val) in CATEGORY_THRESHOLDS.items():
        if min_val <= percent <= max_val:
            return category
    
    return 'بحاجة لتحسين'

def is_wide_format(df: pd.DataFrame) -> bool:
    """كشف ما إذا كان DataFrame بصيغة واسعة"""
    subject_pattern_count = sum(1 for col in df.columns if ' - ' in str(col))
    return subject_pattern_count > len(df.columns) * 0.2

def extract_subject_from_column(col_name: str) -> Tuple[str, str]:
    """استخراج اسم المادة ونوع الحقل من اسم العمود"""
    if ' - ' not in str(col_name):
        return None, None
    
    parts = str(col_name).split(' - ')
    if len(parts) != 2:
        return None, None
    
    subject = parts[0].strip()
    suffix = parts[1].strip()
    
    for field_type, keywords in SUFFIX_KEYWORDS.items():
        if any(keyword in suffix for keyword in keywords):
            return subject, field_type
    
    return subject, 'unknown'

def wide_to_long(df: pd.DataFrame) -> pd.DataFrame:
    """تحويل الصيغة الواسعة إلى طويلة"""
    student_cols = []
    for col in ['الطالب', 'student', 'Student', 'الاسم', 'name']:
        if col in df.columns:
            student_cols.append(col)
            break
    
    for col in ['الصف', 'grade', 'Grade', 'المستوى', 'level']:
        if col in df.columns:
            student_cols.append(col)
            break
    
    for col in ['الشعبة', 'section', 'Section', 'الفصل']:
        if col in df.columns:
            student_cols.append(col)
            break
    
    subject_data = {}
    for col in df.columns:
        if col in student_cols:
            continue
        
        subject, field_type = extract_subject_from_column(col)
        if subject and field_type != 'unknown':
            if subject not in subject_data:
                subject_data[subject] = {}
            subject_data[subject][field_type] = col
    
    long_rows = []
    
    for idx, row in df.iterrows():
        student_info = {col: row[col] for col in student_cols}
        
        for subject, fields in subject_data.items():
            record = student_info.copy()
            record['subject'] = subject
            
            total = row.get(fields.get('total'), 0)
            solved = row.get(fields.get('solved'), 0)
            percent = row.get(fields.get('percent'), None)
            
            total = 0 if pd.isna(total) else float(total)
            solved = 0 if pd.isna(solved) else float(solved)
            
            if pd.isna(percent) or percent is None:
                percent = (solved / total * 100) if total > 0 else 0.0
            else:
                percent = float(percent)
            
            record['total'] = int(total)
            record['solved'] = int(solved)
            record['percent'] = round(percent, 1)
            record['category'] = assign_category(percent)
            
            long_rows.append(record)
    
    return pd.DataFrame(long_rows)

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """تطبيع DataFrame للصيغة القياسية"""
    df = df.copy()
    
    if is_wide_format(df):
        df = wide_to_long(df)
    
    column_mapping = {
        'الطالب': 'student',
        'Student': 'student',
        'الاسم': 'student',
        'الصف': 'grade',
        'Grade': 'grade',
        'المستوى': 'grade',
        'الشعبة': 'section',
        'Section': 'section',
        'المادة': 'subject',
        'Subject': 'subject',
        'إجمالي': 'total',
        'Total': 'total',
        'منجز': 'solved',
        'Solved': 'solved',
        'النسبة': 'percent',
        'Percent': 'percent',
        'الفئة': 'category',
        'Category': 'category'
    }
    
    df = df.rename(columns=column_mapping)
    
    required_cols = ['subject', 'percent']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")
    
    if 'percent' not in df.columns and 'total' in df.columns and 'solved' in df.columns:
        df['percent'] = df.apply(
            lambda row: (row['solved'] / row['total'] * 100) if row['total'] > 0 else 0.0,
            axis=1
        )
    
    if 'category' not in df.columns:
        df['category'] = df['percent'].apply(assign_category)
    
    df['percent'] = df['percent'].fillna(0)
    df['category'] = df['category'].fillna('بحاجة لتحسين')
    
    return df

def aggregate_by_subject(df: pd.DataFrame) -> pd.DataFrame:
    """تجميع البيانات حسب المادة والفئة"""
    agg_data = []
    
    for subject in df['subject'].unique():
        subject_df = df[df['subject'] == subject]
        total_students = len(subject_df)
        avg_completion = subject_df['percent'].mean()
        
        for category in CATEGORY_ORDER:
            count = len(subject_df[subject_df['category'] == category])
            percent_share = (count / total_students * 100) if total_students > 0 else 0
            
            agg_data.append({
                'subject': subject,
                'category': category,
                'count': count,
                'percent_share': round(percent_share, 1),
                'avg_completion': round(avg_completion, 1)
            })
    
    agg_df = pd.DataFrame(agg_data)
    
    subject_order = (
        agg_df.groupby('subject')['avg_completion']
        .first()
        .sort_values(ascending=False)
        .index.tolist()
    )
    
    agg_df['subject'] = pd.Categorical(agg_df['subject'], categories=subject_order, ordered=True)
    agg_df = agg_df.sort_values('subject')
    
    return agg_df

def create_stacked_bar_chart(agg_df: pd.DataFrame, mode: str = 'percent') -> go.Figure:
    """إنشاء رسم بياني شريطي مكدس أفقي"""
    subjects = agg_df['subject'].unique()
    
    fig = go.Figure()
    
    for category in CATEGORY_ORDER:
        category_data = agg_df[agg_df['category'] == category]
        
        if mode == 'percent':
            values = category_data['percent_share'].tolist()
            text = [f"{v:.1f}%" if v > 0 else "" for v in values]
            hovertemplate = (
                "<b>%{y}</b><br>"
                "الفئة: " + category + "<br>"
                "العدد: %{customdata[0]}<br>"
                "النسبة: %{x:.1f}%<br>"
                "<extra></extra>"
            )
        else:
            values = category_data['count'].tolist()
            text = [str(int(v)) if v > 0 else "" for v in values]
            hovertemplate = (
                "<b>%{y}</b><br>"
                "الفئة: " + category + "<br>"
                "العدد: %{x}<br>"
                "النسبة: %{customdata[0]:.1f}%<br>"
                "<extra></extra>"
            )
        
        fig.add_trace(go.Bar(
            name=category,
            x=values,
            y=category_data['subject'].tolist(),
            orientation='h',
            marker=dict(
                color=CATEGORY_COLORS[category],
                line=dict(color='white', width=1)
            ),
            text=text,
            textposition='inside',
            textfont=dict(size=11, color='black', family='Cairo'),
            hovertemplate=hovertemplate,
            customdata=np.column_stack((
                category_data['count'].tolist(),
                category_data['percent_share'].tolist()
            ))
        ))
    
    title = "توزيع الفئات حسب المادة" if mode == 'percent' else "عدد الطلاب حسب الفئة والمادة"
    xaxis_title = "النسبة المئوية (%)" if mode == 'percent' else "عدد الطلاب"
    
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=20, family='Cairo', color='#8A1538'),
            x=0.5,
            xanchor='center'
        ),
        xaxis=dict(
            title=dict(
                text=xaxis_title,
                font=dict(size=14, family='Cairo', color='#111827')
            ),
            tickfont=dict(size=12, family='Cairo'),
            gridcolor='#E5E7EB',
            range=[0, 100] if mode == 'percent' else None
        ),
        yaxis=dict(
            title=dict(
                text="المادة",
                font=dict(size=14, family='Cairo', color='#111827')
            ),
            tickfont=dict(size=12, family='Cairo'),
            autorange='reversed'
        ),
        barmode='stack',
        height=max(400, len(subjects) * 60),
        margin=dict(l=200, r=50, t=80, b=50),
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Cairo'),
        legend=dict(
            title=dict(text="الفئة", font=dict(size=14, family='Cairo')),
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='center',
            x=0.5,
            font=dict(size=12, family='Cairo')
        ),
        hovermode='closest'
    )
    
    return fig

def render_subject_category_chart(df: pd.DataFrame) -> Tuple[go.Figure, pd.DataFrame]:
    """عرض رسم توزيع الفئات حسب المادة"""
    
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">📊 توزيع الفئات حسب المادة الدراسية</h2>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
        <strong>📌 معايير التصنيف:</strong><br>
        🥇 <strong>بلاتيني:</strong> 90% فأكثر | 
        🥈 <strong>ذهبي:</strong> 80-89% | 
        🥉 <strong>فضي:</strong> 70-79% | 
        🟤 <strong>برونزي:</strong> 60-69% | 
        🔴 <strong>بحاجة لتحسين:</strong> أقل من 60%
    </div>
    """, unsafe_allow_html=True)
    
    try:
        normalized_df = normalize_dataframe(df)
        
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        filtered_df = normalized_df.copy()
        
        with col_filter1:
            if 'grade' in normalized_df.columns:
                grades = ['الكل'] + sorted(normalized_df['grade'].dropna().unique().tolist())
                selected_grade = st.selectbox('🎓 الصف', grades, key='grade_filter')
                if selected_grade != 'الكل':
                    filtered_df = filtered_df[filtered_df['grade'] == selected_grade]
        
        with col_filter2:
            if 'section' in normalized_df.columns:
                sections = ['الكل'] + sorted(normalized_df['section'].dropna().unique().tolist())
                selected_section = st.selectbox('📚 الشعبة', sections, key='section_filter')
                if selected_section != 'الكل':
                    filtered_df = filtered_df[filtered_df['section'] == selected_section]
        
        with col_filter3:
            chart_mode = st.radio(
                'نوع العرض',
                ['النسبة المئوية (%)', 'العدد المطلق'],
                horizontal=True,
                key='chart_mode'
            )
            mode = 'percent' if chart_mode == 'النسبة المئوية (%)' else 'count'
        
        agg_df = aggregate_by_subject(filtered_df)
        
        fig = create_stacked_bar_chart(agg_df, mode=mode)
        
        st.plotly_chart(fig, use_container_width=True, key='category_chart')
        
        col_download1, col_download2 = st.columns(2)
        
        with col_download1:
            csv = agg_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 تحميل البيانات (CSV)",
                data=csv,
                file_name=f"subject_categories_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col_download2:
            try:
                png_bytes = fig.to_image(format="png", width=1200, height=800, scale=2)
                st.download_button(
                    label="📥 تحميل الرسم البياني (PNG)",
                    data=png_bytes,
                    file_name=f"chart_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.png",
                    mime="image/png",
                    use_container_width=True
                )
            except Exception as e:
                st.info("💡 لتحميل الصورة، قم بتثبيت: pip install kaleido")
        
        with st.expander("📋 عرض البيانات المُجمّعة"):
            display_df = agg_df.pivot(
                index='subject',
                columns='category',
                values='count'
            ).fillna(0).astype(int)
            
            display_df['المجموع'] = display_df.sum(axis=1)
            
            avg_completion = agg_df.groupby('subject')['avg_completion'].first()
            display_df['متوسط الإنجاز (%)'] = avg_completion
            
            display_df.index.name = 'المادة'
            
            st.dataframe(display_df, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        return fig, agg_df
        
    except Exception as e:
        st.error(f"❌ خطأ في معالجة البيانات: {str(e)}")
        st.exception(e)
        return None, None

# ============================================
# MAIN APPLICATION
# ============================================

st.markdown("")
st.markdown("")

col1, col2, col3 = st.columns([1, 1.5, 1])
with col1:
    st.markdown("<div class='logo-left'>", unsafe_allow_html=True)
    st.image("https://i.imgur.com/QfVfT9X.jpeg", width=120)
    st.markdown("</div>", unsafe_allow_html=True)

with col3:
    st.markdown("<div class='logo-right-group'>", unsafe_allow_html=True)
    st.image("https://i.imgur.com/jFzu8As.jpeg", width=120)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("""
<div class='header-container'>
    <div style='display: flex; align-items: center; justify-content: center; gap: 16px; margin-bottom: 20px;'>
        <svg width="48" height="48" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
            <rect x="4" y="4" width="40" height="40" rx="4" fill="#C9A646" opacity="0.15"/>
            <path d="M12 32V24M18 32V20M24 32V16M30 32V22M36 32V18" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
            <circle cx="12" cy="24" r="2.5" fill="#C9A646"/>
            <circle cx="18" cy="20" r="2.5" fill="#C9A646"/>
            <circle cx="24" cy="16" r="2.5" fill="#C9A646"/>
            <circle cx="30" cy="22" r="2.5" fill="#C9A646"/>
            <circle cx="36" cy="18" r="2.5" fill="#C9A646"/>
            <path d="M12 24L18 20L24 16L30 22L36 18" stroke="#C9A646" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <h1 style='margin: 0; font-size: 40px; font-weight: 700; line-height: 1.25; color: #FFFFFF; text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); letter-spacing: -0.01em;'>
            نظام قطر للتعليم - محلل التقييمات الأسبوعية
        </h1>
    </div>
    <p class='subtitle'>وزارة التربية والتعليم والتعليم العالي</p>
    <p class='accent-line'>ضمان تنمية رقمية مستدامة</p>
    <p class='description'>نظام تحليل شامل وموثوق لنتائج الطلاب</p>
</div>
""", unsafe_allow_html=True)

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

with st.sidebar:
    st.markdown("<div class='logo-sidebar-container'>", unsafe_allow_html=True)
    st.image("https://i.imgur.com/XLef7tS.png", width=110)
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("---")
    st.header("⚙️ الإعدادات")
    
    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader("اختر ملفات Excel", type=["xlsx", "xls"], accept_multiple_files=True)
    
    selected_sheets = []
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        all_sheets = []
        sheet_file_map = {}
        for file_idx, file in enumerate(uploaded_files):
            try:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    sheet_display = f"[ملف {file_idx+1}] {sheet}"
                    all_sheets.append(sheet_display)
                    sheet_file_map[sheet_display] = (file, sheet)
            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف: {e}")
        
        if all_sheets:
            st.info(f"📋 وجدت {len(all_sheets)} مادة من {len(uploaded_files)} ملفات")
            select_all = st.checkbox("✔️ اختر الجميع", value=True)
            if select_all:
                selected_sheets_display = all_sheets
            else:
                selected_sheets_display = st.multiselect("اختر المواد للتحليل", all_sheets, default=[])
            selected_sheets = [(sheet_file_map[s][0], sheet_file_map[s][1]) for s in selected_sheets_display]
    else:
        st.info("📤 ارفع ملفات Excel للبدء")
    
    st.markdown("---")
    st.subheader("🏫 معلومات المدرسة")
    school_name = st.text_input("اسم المدرسة", placeholder="مدرسة قطر النموذجية")
    
    st.subheader("✍️ التوقيعات")
    coordinator_name = st.text_input("منسق المشاريع")
    academic_deputy = st.text_input("النائب الأكاديمي")
    admin_deputy = st.text_input("النائب الإداري")
    principal_name = st.text_input("مدير المدرسة")
    
    st.markdown("---")
    run_analysis = st.button("▶️ تشغيل التحليل", use_container_width=True, type="primary", disabled=not (uploaded_files and selected_sheets))
    
    st.markdown("---")
    st.markdown("<div class='logo-footer-container'>", unsafe_allow_html=True)
    st.image("https://i.imgur.com/XLef7tS.png", width=120)
    st.markdown("</div>", unsafe_allow_html=True)

if not uploaded_files:
    st.info("📤 الرجاء رفع ملفات Excel من الشريط الجانبي للبدء في التحليل")
elif run_analysis:
    with st.spinner("⏳ جاري التحليل، الرجاء الانتظار..."):
        all_results = []
        for file, sheet in selected_sheets:
            results = analyze_excel_file(file, sheet)
            all_results.extend(results)
        
        if all_results:
            df = pd.DataFrame(all_results)
            st.session_state.analysis_results = df
            pivot = create_pivot_table(df)
            st.session_state.pivot_table = pivot
            st.success(f"✅ تم تحليل {len(pivot)} طالب من {len(set(df['subject']))} مادة بنجاح")

if st.session_state.pivot_table is not None:
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results
    
    st.markdown("### 📈 ملخص النتائج")
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("👥 إجمالي الطلاب", len(pivot))
    with col2:
        st.metric("📚 عدد المواد", df['subject'].nunique())
    with col3:
        avg = pivot['المتوسط'].mean() if 'المتوسط' in pivot.columns else 0
        st.metric("📊 متوسط الإنجاز", f"{avg:.1f}%")
    with col4:
        platinum = len(pivot[pivot['الفئة'].str.contains('بلاتينية', na=False)])
        st.metric("🥇 فئة بلاتينية", platinum)
    with col5:
        zero = len(pivot[pivot['المتوسط'] == 0])
        st.metric("⚠️ بدون إنجاز", zero)
    
    st.divider()
    
    st.subheader("📋 جدول النتائج التفصيلي")
    st.dataframe(pivot, use_container_width=True, height=400)
    
    st.divider()
    
    st.subheader("💾 تحميل النتائج")
    col1, col2 = st.columns(2)
    
    with col1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot.to_excel(writer, index=False, sheet_name='النتائج')
        st.download_button(
            "📥 تحميل Excel",
            output.getvalue(),
            f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        csv_data = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 تحميل CSV",
            csv_data,
            f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    st.divider()
    
    render_subject_category_chart(df)
    
    st.divider()
    
    st.markdown("""
    <div style='margin-top: 80px; padding: 0; background: transparent;'>
        <div style='width: 100%; height: 4px; background: linear-gradient(90deg, transparent 0%, #C9A646 20%, #E8D4A0 50%, #C9A646 80%, transparent 100%); margin-bottom: 40px;'></div>
        
        <div style='text-align: center; padding: 48px 32px; background: linear-gradient(135deg, #8A1538 0%, #6B1029 100%); border-radius: 12px; box-shadow: 0 8px 24px rgba(138, 21, 56, 0.25); position: relative; overflow: hidden;'>
            
            <div style='position: absolute; top: 0; left: 0; right: 0; height: 4px; background: linear-gradient(90deg, #C9A646 0%, #E8D4A0 50%, #C9A646 100%);'></div>
            
            <div style='margin-bottom: 24px;'>
                <img src='https://i.imgur.com/XLef7tS.png' style='width: 100px; height: auto; opacity: 0.95;' alt='Ministry Logo'>
            </div>
            
            <p style='color: #FFFFFF; font-weight: 700; font-size: 16px; margin-bottom: 20px; letter-spacing: 0.03em; line-height: 1.6;'>
                © 2025 وزارة التربية والتعليم والتعليم العالي
            </p>
            <p style='color: #FFFFFF; font-weight: 700; font-size: 16px; margin-bottom: 8px; letter-spacing: 0.02em;'>
                جميع الحقوق محفوظة
            </p>
            
            <div style='width: 80px; height: 3px; background: #C9A646; margin: 24px auto; border-radius: 2px;'></div>
            
            <p style='color: #FFFFFF; font-weight: 700; font-size: 17px; margin-bottom: 16px; letter-spacing: 0.01em;'>
                مدرسة عثمان بن عفان النموذجية للبنين
            </p>
            
            <p style='color: #FFFFFF; font-weight: 600; font-size: 15px; margin-bottom: 16px; opacity: 0.95;'>
                منسقة المشاريع الإلكترونية / سحر عثمان
            </p>
            
            <p style='color: #F5F5F5; font-size: 14px; margin: 0; opacity: 0.9;'>
                📧 للتواصل: <a href='mailto:S.mahgoub0101@education.qa' style='color: #C9A646; font-weight: 600; text-decoration: none; transition: opacity 0.3s; border-bottom: 1px solid #C9A646;'>S.mahgoub0101@education.qa</a>
            </p>
            
            <p style='color: #F5F5F5; font-size: 12px; margin-top: 24px; opacity: 0.8; letter-spacing: 0.02em;'>
                تطوير وتصميم: قسم التحول الرقمي
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)

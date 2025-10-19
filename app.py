import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
import base64
from datetime import datetime

# Page config
st.set_page_config(
    page_title="محلل التقييمات الأسبوعية",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced CSS with improved color palette and typography
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    
    * {
        font-family: 'Cairo', 'Arial', sans-serif;
    }
    
    :root {
        --primary: #8B3A3A;
        --primary-dark: #6B2A2A;
        --primary-light: #D4A574;
        --secondary: #F5E6D3;
        --accent: #C41E3A;
        --success: #27AE60;
        --warning: #F39C12;
        --danger: #E74C3C;
        --text-dark: #2C3E50;
        --text-light: #95A5A6;
        --bg-light: #F8F9FA;
    }
    
    /* Main layout */
    .main {
        background: linear-gradient(135deg, #f8f9fa 0%, #f0ebe5 100%);
    }
    
    /* Header styling */
    .header-container {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%);
        padding: 40px;
        border-radius: 20px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 30px rgba(139, 58, 58, 0.25);
        border: 2px solid #D4A574;
    }
    
    .header-container h1 {
        margin: 0;
        font-size: 36px;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    
    .header-container p {
        margin: 12px 0 0 0;
        font-size: 16px;
        opacity: 0.95;
        font-weight: 400;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #8B3A3A 0%, #A0483D 100%);
        box-shadow: 2px 0 15px rgba(0, 0, 0, 0.1);
    }
    
    [data-testid="stSidebar"] * {
        color: white;
    }
    
    /* Section boxes */
    .section-box {
        background: white;
        padding: 25px;
        border-radius: 15px;
        margin: 20px 0;
        border-left: 5px solid #8B3A3A;
        box-shadow: 0 4px 15px rgba(139, 58, 58, 0.1);
        transition: all 0.3s ease;
    }
    
    .section-box:hover {
        box-shadow: 0 8px 25px rgba(139, 58, 58, 0.15);
        transform: translateY(-2px);
    }
    
    /* Metric boxes */
    .metric-box {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        text-align: center;
        font-weight: 600;
        box-shadow: 0 6px 20px rgba(139, 58, 58, 0.2);
        border: 2px solid #D4A574;
        transition: all 0.3s ease;
    }
    
    .metric-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 30px rgba(139, 58, 58, 0.3);
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%) !important;
        color: white !important;
        border: 2px solid #D4A574 !important;
        padding: 12px 28px !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(139, 58, 58, 0.2) !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 8px 20px rgba(139, 58, 58, 0.35) !important;
        border-color: white !important;
    }
    
    .stButton > button:active {
        transform: translateY(-1px) !important;
    }
    
    /* Input fields */
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select,
    .stMultiSelect > div > div > div {
        border: 2px solid #D4A574 !important;
        border-radius: 10px !important;
        font-family: 'Cairo', 'Arial', sans-serif !important;
    }
    
    .stTextInput > div > div > input:focus,
    .stSelectbox > div > div > select:focus {
        border-color: #8B3A3A !important;
        box-shadow: 0 0 10px rgba(139, 58, 58, 0.3) !important;
    }
    
    /* Dividers */
    hr {
        border-color: #8B3A3A !important;
        margin: 25px 0 !important;
    }
    
    /* Info and status boxes */
    .stSuccess {
        background-color: #E8F8F5 !important;
        border-left: 5px solid #27AE60 !important;
        border-radius: 10px !important;
    }
    
    .stWarning {
        background-color: #FEF5E7 !important;
        border-left: 5px solid #F39C12 !important;
        border-radius: 10px !important;
    }
    
    .stError {
        background-color: #FADBD8 !important;
        border-left: 5px solid #E74C3C !important;
        border-radius: 10px !important;
    }
    
    .stInfo {
        background-color: #F4ECF7 !important;
        border-left: 5px solid #8B3A3A !important;
        border-radius: 10px !important;
    }
    
    /* Card styling */
    .card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border-right: 4px solid #8B3A3A;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
        margin: 8px 0;
        transition: all 0.2s ease;
    }
    
    .card:hover {
        box-shadow: 0 4px 15px rgba(139, 58, 58, 0.15);
    }
    
    /* Tables */
    table {
        border-collapse: collapse;
        width: 100%;
    }
    
    thead {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%) !important;
        color: white !important;
    }
    
    tbody tr:nth-child(odd) {
        background-color: #F8F9FA !important;
    }
    
    tbody tr:hover {
        background-color: #F5E6D3 !important;
    }
    
    /* Text styling */
    h1, h2, h3, h4, h5, h6 {
        color: #8B3A3A;
        font-weight: 700;
    }
    
    .subheading {
        color: #A0483D;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Ministry Logo SVG
MINISTRY_LOGO_SVG = """
<svg width="120" height="120" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg">
    <defs>
        <linearGradient id="grad1" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" style="stop-color:#8B3A3A;stop-opacity:1" />
            <stop offset="100%" style="stop-color:#A0483D;stop-opacity:1" />
        </linearGradient>
    </defs>
    <circle cx="60" cy="60" r="55" fill="url(#grad1)" stroke="#D4A574" stroke-width="2"/>
    <circle cx="60" cy="60" r="50" fill="white"/>
    <text x="60" y="50" font-family="Arial, sans-serif" font-size="18" font-weight="bold" text-anchor="middle" fill="#8B3A3A">وزارة</text>
    <text x="60" y="72" font-family="Arial, sans-serif" font-size="14" text-anchor="middle" fill="#8B3A3A">التعليم</text>
    <circle cx="60" cy="60" r="48" fill="none" stroke="#8B3A3A" stroke-width="1.5" stroke-dasharray="8,4"/>
</svg>
"""

# ================== HELPER FUNCTIONS ==================

def parse_sheet_name(sheet_name):
    """Extract subject, level, and section from sheet name"""
    parts = sheet_name.strip().split()
    level = ""
    section = ""
    subject_parts = []
    
    for part in parts:
        if part.isdigit() or (part.startswith('0') and len(part) <= 2):
            if not level:
                level = part
            else:
                section = part
        else:
            subject_parts.append(part)
    
    subject = " ".join(subject_parts) if subject_parts else sheet_name
    return subject, level, section

def analyze_excel_file(file, sheet_name):
    """Analyze a single Excel sheet"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        
        # استخراج جميع التواريخ من الصف الثاني (row 1) - الأعمدة I, J, K (indices 8, 9, 10)
        due_dates = []
        try:
            for col_idx in [8, 9, 10]:  # I, J, K columns
                if col_idx < df.shape[1]:
                    cell_value = df.iloc[1, col_idx]
                    if pd.notna(cell_value):
                        try:
                            # محاولة تحويل القيمة إلى تاريخ
                            due_date = pd.to_datetime(cell_value)
                            # التحقق من أنها تاريخ حقيقي
                            if 2000 <= due_date.year <= 2100:
                                due_dates.append(due_date.date())
                        except:
                            pass
        except:
            pass
        
        if len(df) > 1:
            level_from_excel = str(df.iloc[1, 1]).strip() if pd.notna(df.iloc[1, 1]) else ""
            section_from_excel = str(df.iloc[1, 2]).strip() if pd.notna(df.iloc[1, 2]) else ""
            level = level_from_excel if level_from_excel and level_from_excel != 'nan' else level_from_name
            section = section_from_excel if section_from_excel and section_from_excel != 'nan' else section_from_name
        else:
            level = level_from_name
            section = section_from_name
        
        assessment_titles = []
        for col_idx in range(7, df.shape[1]):
            title = df.iloc[0, col_idx]
            if pd.notna(title):
                title_str = str(title).strip()
                if title_str and title_str not in ['-', '—', 'nan', '']:
                    assessment_titles.append(title_str)
        
        total_assessments = len(assessment_titles)
        results = []
        
        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue
            
            student_name_clean = " ".join(str(student_name).strip().split())
            m_count = 0
            pending_titles = []
            
            for i, col_idx in enumerate(range(7, df.shape[1])):
                if i < len(assessment_titles):
                    cell_value = df.iloc[idx, col_idx]
                    if pd.isna(cell_value):
                        m_count += 1
                        pending_titles.append(assessment_titles[i])
                    else:
                        cell_str = str(cell_value).strip().upper()
                        if cell_str in ['-', '—', 'NAN', '']:
                            m_count += 1
                            pending_titles.append(assessment_titles[i])
                        elif cell_str == 'M':
                            m_count += 1
                            pending_titles.append(assessment_titles[i])
            
            completed_count = total_assessments - m_count
            solve_pct = (completed_count / total_assessments * 100) if total_assessments > 0 else 0.0
            
            results.append({
                "student_name": student_name_clean,
                "subject": subject,
                "level": str(level).strip(),
                "section": str(section).strip(),
                "solve_pct": solve_pct,
                "completed_count": completed_count,
                "total_count": total_assessments,
                "pending_titles": ", ".join(pending_titles) if pending_titles else "",
                "due_dates": due_dates  # تخزين جميع التواريخ
            })
        
        return results
    except Exception as e:
        st.error(f"خطأ في تحليل الورقة {sheet_name}: {str(e)}")
        return []

def create_pivot_table(df):
    """Create pivot table - ONE ROW PER STUDENT - NO DUPLICATES"""
    df_clean = df.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'], keep='first')
    unique_students = df_clean.groupby(['student_name', 'level', 'section']).size().reset_index(name='count')
    unique_students = unique_students[['student_name', 'level', 'section']]
    unique_students = unique_students.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
    result = unique_students.copy()
    subjects = sorted(df_clean['subject'].unique())
    
    for subject in subjects:
        subject_df = df_clean[df_clean['subject'] == subject][['student_name', 'level', 'section', 
                                                                'total_count', 'completed_count', 
                                                                'pending_titles', 'solve_pct']].copy()
        subject_df = subject_df.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')
        subject_df = subject_df.rename(columns={
            'total_count': f"{subject} - إجمالي التقييمات",
            'completed_count': f"{subject} - المنجز",
            'pending_titles': f"{subject} - عناوين التقييمات المتبقية",
            'solve_pct': f"{subject} - نسبة الإنجاز %"
        })
        result = result.merge(subject_df, on=['student_name', 'level', 'section'], how='left')
    
    pct_cols = [col for col in result.columns if 'نسبة الإنجاز %' in col]
    if pct_cols:
        result['نسبة حل التقييمات في جميع المواد'] = result[pct_cols].mean(axis=1)
        
        def categorize(pct):
            if pd.isna(pct):
                return "-"
            elif pct == 0:
                return "لا يستفيد من النظام 🚫"
            elif pct >= 90:
                return "البلاتينية 🥇"
            elif pct >= 80:
                return "الذهبي 🥈"
            elif pct >= 70:
                return "الفضي 🥉"
            elif pct >= 60:
                return "البرونزي"
            else:
                return "يحتاج تحسين ⚠️"
        
        result['الفئة'] = result['نسبة حل التقييمات في جميع المواد'].apply(categorize)
    
    result = result.rename(columns={
        'student_name': 'اسم الطالب',
        'level': 'الصف',
        'section': 'الشعبة'
    })
    
    result = result.drop_duplicates(subset=['اسم الطالب', 'الصف', 'الشعبة'], keep='first')
    return result.reset_index(drop=True)

def generate_student_html_report(student_row, school_name="", coordinator="", academic="", admin="", principal="", logo_base64=""):
    """Generate individual student HTML report"""
    student_name = student_row['اسم الطالب']
    level = student_row['الصف']
    section = student_row['الشعبة']
    total_assessments = 0
    total_completed = 0
    subjects_data = []
    
    for col in student_row.index:
        if ' - إجمالي التقييمات' in col:
            subject = col.replace(' - إجمالي التقييمات', '')
            total_col = f"{subject} - إجمالي التقييمات"
            completed_col = f"{subject} - المنجز"
            pending_col = f"{subject} - عناوين التقييمات المتبقية"
            
            if pd.notna(student_row[total_col]):
                total = int(student_row[total_col])
                completed = int(student_row[completed_col]) if pd.notna(student_row[completed_col]) else 0
                pending_titles = str(student_row[pending_col]) if pd.notna(student_row[pending_col]) and str(student_row[pending_col]) != "" else "-"
                pct = (completed / total * 100) if total > 0 else 0
                total_assessments += total
                total_completed += completed
                subjects_data.append({'subject': subject, 'total': total, 'completed': completed, 'pending': pending_titles, 'pct': pct})
    
    subjects_html = ""
    for data in subjects_data:
        subjects_html += f"""
        <tr>
            <td style="text-align: right; padding: 12px;">{data['subject']}</td>
            <td style="text-align: center; padding: 12px;">{data['total']}</td>
            <td style="text-align: center; padding: 12px;">{data['completed']}</td>
            <td style="text-align: center; padding: 12px; color: #8B3A3A; font-weight: bold;">{data['pct']:.1f}%</td>
            <td style="text-align: right; padding: 12px; font-size: 12px;">{data['pending']}</td>
        </tr>"""
    
    solve_pct = (total_completed / total_assessments * 100) if total_assessments > 0 else 0
    remaining = total_assessments - total_completed
    
    if solve_pct == 0:
        recommendation = "الطالب لم يستفيد من النظام"
        category_color = "#9E9E9E"
        category_name = "لا يستفيد"
    elif solve_pct >= 90:
        recommendation = "أداء ممتاز"
        category_color = "#27AE60"
        category_name = "البلاتينية 🥇"
    elif solve_pct >= 80:
        recommendation = "أداء جيد جداً"
        category_color = "#8BC34A"
        category_name = "الذهبي 🥈"
    elif solve_pct >= 70:
        recommendation = "أداء جيد"
        category_color = "#F39C12"
        category_name = "الفضي 🥉"
    elif solve_pct >= 60:
        recommendation = "أداء مقبول"
        category_color = "#E67E22"
        category_name = "البرونزي"
    else:
        recommendation = "يحتاج اهتماماً أكثر"
        category_color = "#E74C3C"
        category_name = "يحتاج تحسين ⚠️"
    
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 100px; margin: 10px;">' if logo_base64 else ""
    school_section = f"<h2 style='color: #8B3A3A; margin: 10px 0; font-size: 18px;'>{school_name}</h2>" if school_name else ""
    
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_name}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            body {{font-family: 'Cairo', sans-serif; direction: rtl; padding: 20px; background: #f5f5f5;}}
            .container {{max-width: 800px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 5px 20px rgba(0,0,0,0.1); border-radius: 15px;}}
            .header {{display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid #8B3A3A; padding-bottom: 20px; margin-bottom: 30px;}}
            .header-left {{text-align: left;}} .header-right {{text-align: center; flex: 1;}}
            h1 {{color: #8B3A3A; margin: 10px 0; font-size: 24px; font-weight: 700;}}
            h3 {{color: #A0483D; font-size: 16px; font-weight: 600;}}
            .info-box {{background: #F5E6D3; padding: 20px; border-radius: 10px; margin: 20px 0; border-left: 5px solid #8B3A3A;}}
            table {{width: 100%; border-collapse: collapse; margin: 20px 0;}}
            th {{background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%); color: white; padding: 12px; text-align: center; border: 1px solid #8B3A3A; font-weight: 600;}}
            td {{padding: 12px; border: 1px solid #ddd; text-align: center;}}
            tr:nth-child(even) {{background: #f9f9f9;}}
            .stats {{display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin: 20px 0;}}
            .stat {{background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; font-weight: 600;}}
            .stat-value {{font-size: 24px; font-weight: bold; margin: 10px 0;}}
            .category {{background: {category_color}; color: white; padding: 12px; border-radius: 8px; display: inline-block; margin: 15px 0; font-weight: 600;}}
            .recommendation {{background: {category_color}; color: white; padding: 20px; border-radius: 10px; margin: 20px 0; text-align: center; font-weight: 600; font-size: 16px;}}
            .signature {{margin-top: 40px; border-top: 2px solid #8B3A3A; padding-top: 20px;}}
            .sig-line {{margin: 20px 0; font-size: 14px;}}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="header-left">{logo_html}</div>
                <div class="header-right">
                    {school_section}
                    <h1>تقرير أداء الطالب</h1>
                    <p style="color: #8B3A3A; font-weight: 600;">وزارة التربية والتعليم والتعليم العالي</p>
                </div>
            </div>
            
            <div class="info-box">
                <h3>معلومات الطالب</h3>
                <p><strong>الاسم:</strong> {student_name}</p>
                <p><strong>الصف:</strong> {level} | <strong>الشعبة:</strong> {section}</p>
                <div class="category">الفئة: {category_name}</div>
            </div>
            
            <table>
                <thead><tr><th>المادة</th><th>الإجمالي</th><th>المنجز</th><th>النسبة</th><th>المتبقي</th></tr></thead>
                <tbody>{subjects_html}</tbody>
            </table>
            
            <div class="stats">
                <div class="stat"><div style="font-size: 14px;">إجمالي</div><div class="stat-value">{total_assessments}</div></div>
                <div class="stat"><div style="font-size: 14px;">منجز</div><div class="stat-value">{total_completed}</div></div>
                <div class="stat"><div style="font-size: 14px;">متبقي</div><div class="stat-value">{remaining}</div></div>
                <div class="stat"><div style="font-size: 14px;">النسبة</div><div class="stat-value">{solve_pct:.1f}%</div></div>
            </div>
            
            <div class="recommendation">{recommendation}</div>
            
            <div class="signature">
                <div class="sig-line"><strong>منسق المشاريع:</strong> {coordinator if coordinator else "_____________"}</div>
                <div class="sig-line"><strong>النائب الأكاديمي:</strong> {academic if academic else "_____________"}</div>
                <div class="sig-line"><strong>النائب الإداري:</strong> {admin if admin else "_____________"}</div>
                <div class="sig-line"><strong>مدير المدرسة:</strong> {principal if principal else "_____________"}</div>
                <p style="text-align: center; color: #999; font-size: 12px; margin-top: 30px;">
                    تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}
                </p>
            </div>
        </div>
    </body>
    </html>"""
    return html

# ================== MAIN APP ==================

st.markdown(f"""
<div class="header-container">
    <div style="display: flex; justify-content: center; align-items: center; gap: 20px; margin-bottom: 15px;">
        {MINISTRY_LOGO_SVG}
    </div>
    <h1>📊 محلل التقييمات الأسبوعية</h1>
    <p style="font-size: 14px; margin: 10px 0;">وزارة التربية والتعليم والتعليم العالي</p>
    <p style="font-size: 13px; color: #D4A574; font-weight: 600; margin: 5px 0;">لضمان تنمية رقمية مستدامة</p>
    <p style="font-size: 12px; opacity: 0.9;">نظام تحليل شامل وموثوق لنتائج الطلاب</p>
</div>
""", unsafe_allow_html=True)

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

# Sidebar
with st.sidebar:
    st.markdown(f"<div style='text-align: center; margin: 20px 0;'>{MINISTRY_LOGO_SVG}</div>", unsafe_allow_html=True)
    st.markdown("---")
    st.header("⚙️ الإعدادات والتحليل")
    
    # File Upload
    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="📌 يدعم تحليل عدة ملفات في آن واحد"
    )
    
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        try:
            all_sheets = []
            sheet_file_map = {}
            for file_idx, file in enumerate(uploaded_files):
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    sheet_display = f"[ملف {file_idx+1}] {sheet}"
                    all_sheets.append(sheet_display)
                    sheet_file_map[sheet_display] = (file, sheet)
            
            if all_sheets:
                st.info(f"📊 وجدت {len(all_sheets)} مادة من {len(uploaded_files)} ملفات")
                
                # خيار اختيار الأوراق
                select_all = st.checkbox("✅ اختر جميع الأوراق", value=True)
                
                if select_all:
                    selected_sheets_display = all_sheets
                else:
                    selected_sheets_display = st.multiselect(
                        "🔍 فلتر الأوراق: اختر الأوراق المراد تحليلها",
                        all_sheets,
                        default=[]
                    )
                
                selected_sheets = [(sheet_file_map[s][0], sheet_file_map[s][1]) for s in selected_sheets_display]
            else:
                selected_sheets = []
        except Exception as e:
            st.error(f"خطأ: {e}")
            selected_sheets = []
    else:
        st.info("💡 ارفع ملفات Excel للبدء")
        selected_sheets = []
        select_all = False
    
    st.markdown("---")
    st.subheader("🏫 معلومات المدرسة")
    school_name = st.text_input("اسم المدرسة", placeholder="مثال: مدرسة قطر النموذجية")
    
    st.subheader("🖼️ شعار الوزارة")
    uploaded_logo = st.file_uploader("ارفع شعار", type=["png", "jpg", "jpeg"])
    logo_base64 = ""
    if uploaded_logo:
        logo_base64 = base64.b64encode(uploaded_logo.read()).decode()
        st.success("✅ تم رفع الشعار")
    
    st.markdown("---")
    st.subheader("📅 فلتر التاريخ")
    
    date_filter_type = st.radio("نوع الفلتر:", ["بدون فلتر", "من تاريخ إلى تاريخ", "من تاريخ إلى الآن"])
    
    from_date = None
    to_date = None
    
    st.caption("💡 سيتم قراءة التواريخ تلقائياً من الملفات (الأعمدة I, J, K - الصف الثاني)")
    
    if date_filter_type == "من تاريخ إلى تاريخ":
        col1, col2 = st.columns(2)
        with col1:
            from_date = st.date_input("من تاريخ", key="from_date")
        with col2:
            to_date = st.date_input("إلى تاريخ", key="to_date")
    elif date_filter_type == "من تاريخ إلى الآن":
        from_date = st.date_input("من تاريخ", key="from_date_now")
        to_date = pd.Timestamp.now().date()
    
    st.markdown("---")
    st.subheader("✍️ معلومات التوقيعات")
    
    coordinator_name = st.text_input("منسق المشاريع", placeholder="أدخل الاسم")
    academic_deputy = st.text_input("النائب الأكاديمي", placeholder="أدخل الاسم")
    admin_deputy = st.text_input("النائب الإداري", placeholder="أدخل الاسم")
    principal_name = st.text_input("مدير المدرسة", placeholder="أدخل الاسم")
    
    st.markdown("---")
    run_analysis = st.button(
        "🚀 تشغيل التحليل",
        use_container_width=True,
        type="primary",
        disabled=not (uploaded_files and selected_sheets)
    )

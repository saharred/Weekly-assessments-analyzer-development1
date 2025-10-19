import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
import base64
from datetime import datetime
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
    
    /* Header Container - Modern & Professional */
    .header-container {
        background: linear-gradient(to bottom, #8A1538 0%, #6B1029 100%);
        padding: 48px 40px;
        border-radius: 0;
        color: white;
        text-align: center;
        margin-bottom: 32px;
        box-shadow: 0 4px 12px rgba(138, 21, 56, 0.15);
        border-bottom: 3px solid #C5A572;
    }
    
    .header-container h1 { 
        margin: 0 0 12px 0;
        font-size: 32px;
        font-weight: 700;
        line-height: 1.3;
    }
    
    .header-container .subtitle { 
        font-size: 14px;
        font-weight: 500;
        opacity: 0.95;
        margin: 0 0 8px 0;
    }
    
    .header-container .accent-line {
        font-size: 12px;
        color: #C5A572;
        font-weight: 600;
        margin: 0 0 16px 0;
    }
    
    .header-container .description {
        font-size: 12px;
        opacity: 0.9;
        margin: 0;
    }
    
    /* Sidebar - Solid Qatar Maroon */
    [data-testid="stSidebar"] {
        background: #8A1538 !important;
        border-right: 1px solid #6B1029;
        box-shadow: 2px 0 8px rgba(0, 0, 0, 0.1);
    }
    
    [data-testid="stSidebar"] * { 
        color: white !important;
    }
    
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        font-weight: 600;
    }
    
    /* Section Box - Clean & Minimal */
    .section-box {
        background: #FAFAF9;
        padding: 24px;
        border-radius: 8px;
        margin: 24px 0;
        border-left: 4px solid #8A1538;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    }
    
    /* Metric Box - Modern Card Style */
    .metric-box {
        background: #FFFFFF;
        border: 2px solid #E8E8E6;
        border-left: 4px solid #8A1538;
        padding: 20px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
        transition: all 0.2s ease;
    }
    
    .metric-box:hover {
        border-left-color: #C5A572;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        transform: translateY(-2px);
    }
    
    .metric-value {
        font-size: 36px;
        font-weight: 700;
        color: #8A1538;
        line-height: 1.2;
        margin-bottom: 4px;
    }
    
    .metric-label {
        font-size: 12px;
        font-weight: 500;
        color: #5A5A5A;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Buttons - Clean Qatar Maroon */
    .stButton > button {
        background: #8A1538 !important;
        color: white !important;
        border: none !important;
        padding: 12px 24px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 4px rgba(138, 21, 56, 0.2) !important;
    }
    
    .stButton > button:hover {
        background: #6B1029 !important;
        box-shadow: 0 4px 8px rgba(138, 21, 56, 0.3) !important;
        transform: translateY(-1px) !important;
    }
    
    /* Divider */
    hr { 
        border-color: #E8E8E6 !important;
        margin: 24px 0 !important;
        border-width: 1px !important;
    }
    
    /* Typography */
    h1, h2, h3, h4, h5, h6 { 
        color: #FFFFFF;
        font-weight: 600;
    }
    
    h1 { font-size: 32px; line-height: 1.3; margin-bottom: 12px; }
    h2 { font-size: 24px; line-height: 1.4; margin-bottom: 16px; }
    h3 { font-size: 20px; line-height: 1.4; margin-bottom: 12px; }
    h4 { font-size: 16px; line-height: 1.5; margin-bottom: 8px; }
    
    p, div, span {
        font-size: 14px;
        line-height: 1.6;
        color: #2C2C2C;
    }
    
    /* Logo Containers */
    .logo-header-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 20px 40px;
        background: #FAFAF9;
        border-bottom: 2px solid #E8E8E6;
        margin-bottom: 0;
    }
    
    .logo-left {
        display: flex;
        align-items: center;
        padding: 10px;
    }
    
    .logo-right-group {
        display: flex;
        gap: 20px;
        align-items: center;
        padding: 10px;
    }
    
    .logo-sidebar-container {
        text-align: center;
        padding: 24px 20px;
        margin-bottom: 24px;
    }
    
    .logo-footer-container {
        text-align: center;
        padding: 20px;
        margin: 0 auto 16px;
    }
    
    /* Download Buttons */
    .stDownloadButton > button {
        background: #FFFFFF !important;
        color: #8A1538 !important;
        border: 2px solid #8A1538 !important;
        padding: 10px 22px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.2s ease !important;
    }
    
    .stDownloadButton > button:hover {
        background: #8A1538 !important;
        color: #FFFFFF !important;
        transform: translateY(-1px) !important;
    }
    
    /* File Uploader Button - Browse Files */
    [data-testid="stFileUploader"] label {
        color: #8A1538 !important;
        font-weight: 600 !important;
    }
    
    [data-testid="stFileUploader"] button {
        background: #FFFFFF !important;
        color: #8A1538 !important;
        border: 2px solid #8A1538 !important;
        font-weight: 600 !important;
    }
    
    [data-testid="stFileUploader"] button:hover {
        background: #8A1538 !important;
        color: #FFFFFF !important;
    }
    
    [data-testid="stFileUploader"] section {
        border-color: #8A1538 !important;
    }
    
    [data-testid="stFileUploader"] small {
        color: #FFFFFF !important;
    }
    
    /* Dataframe Styling */
    [data-testid="stDataFrame"] {
        border: 1px solid #E8E8E6;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    }
    
    /* Metrics Enhancement */
    [data-testid="stMetricValue"] {
        font-size: 36px !important;
        font-weight: 700 !important;
        color: #8A1538 !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 12px !important;
        font-weight: 500 !important;
        color: #5A5A5A !important;
        text-transform: uppercase !important;
    }
</style>
""", unsafe_allow_html=True)

def parse_sheet_name(sheet_name):
    try:
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
    except Exception as e:
        logger.error(f"خطأ: {str(e)}")
        return sheet_name, "", ""

@st.cache_data
def analyze_excel_file(file, sheet_name):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        
        due_dates = []
        try:
            for col_idx in [8, 9, 10]:
                if col_idx < df.shape[1]:
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
        
        assessment_titles = []
        try:
            for col_idx in range(7, df.shape[1]):
                title = df.iloc[0, col_idx]
                if pd.notna(title):
                    title_str = str(title).strip()
                    if title_str and title_str not in ['-', '—', 'nan', '']:
                        assessment_titles.append(title_str)
        except (IndexError, KeyError):
            pass
        
        total_assessments = len(assessment_titles)
        results = []
        
        try:
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
                    "due_dates": due_dates
                })
        except (IndexError, KeyError) as e:
            logger.error(f"خطأ: {str(e)}")
        
        return results
    except Exception as e:
        logger.error(f"خطأ: {str(e)}")
        st.error(f"خطأ: {str(e)}")
        return []

@st.cache_data
def create_pivot_table(df):
    try:
        df_clean = df.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'], keep='first')
        unique_students = df_clean.groupby(['student_name', 'level', 'section']).size().reset_index(name='count')
        unique_students = unique_students[['student_name', 'level', 'section']]
        unique_students = unique_students.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
        result = unique_students.copy()
        
        subjects = sorted(df_clean['subject'].unique())
        
        for subject in subjects:
            subject_df = df_clean[df_clean['subject'] == subject][['student_name', 'level', 'section', 'total_count', 'completed_count', 'pending_titles', 'solve_pct']].copy()
            subject_df = subject_df.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')
            subject_df = subject_df.rename(columns={
                'total_count': f"{subject} - إجمالي",
                'completed_count': f"{subject} - منجز",
                'pending_titles': f"{subject} - متبقي",
                'solve_pct': f"{subject} - النسبة"
            })
            result = result.merge(subject_df, on=['student_name', 'level', 'section'], how='left')
        
        pct_cols = [col for col in result.columns if 'النسبة' in col]
        if pct_cols:
            result['المتوسط'] = result[pct_cols].mean(axis=1)
            
            def categorize(pct):
                if pd.isna(pct):
                    return "-"
                elif pct == 0:
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
        
        result = result.rename(columns={'student_name': 'الطالب', 'level': 'الصف', 'section': 'الشعبة'})
        result = result.drop_duplicates(subset=['الطالب', 'الصف', 'الشعبة'], keep='first')
        return result.reset_index(drop=True)
    except Exception as e:
        logger.error(f"خطأ: {str(e)}")
        return pd.DataFrame()

# Header with Logos - Improved Layout
st.markdown("")
st.markdown("")

col1, col2, col3 = st.columns([1, 1.5, 1])
with col1:
    st.markdown("<div class='logo-left'>", unsafe_allow_html=True)
    st.image("https://i.imgur.com/QfVfT9X.jpeg", width=120)
    st.markdown("</div>", unsafe_allow_html=True)

# Main Header
st.markdown("""
<div class='header-container'>
    <h1>📊 محلل التقييمات الأسبوعية</h1>
    <p class='subtitle'>وزارة التربية والتعليم والتعليم العالي</p>
    <p class='accent-line'>ضمان تنمية رقمية مستدامة</p>
    <p class='description'>نظام تحليل شامل وموثوق لنتائج الطلاب</p>
</div>
""", unsafe_allow_html=True)

# Session State
if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

# Sidebar
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

# Main Content
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
    
    # Metrics Section
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
    
    # Data Table
    st.subheader("📋 جدول النتائج التفصيلي")
    st.dataframe(pivot, use_container_width=True, height=400)
    
    st.divider()
    
    # Download Section
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
    
    # Footer
    st.markdown("""
    <div style='text-align: center; margin-top: 48px; padding: 24px; color: #5A5A5A; font-size: 12px; border-top: 2px solid #E8E8E6;'>
        <p style='color: #8A1538; font-weight: 600; margin-bottom: 8px;'>
            © جميع الحقوق محفوظة لمدرسة عثمان بن عفان النموذجية للبنين
        </p>
        <p style='margin: 0;'>
            منسقة المشاريع الإلكترونية / سحر عثمان
        </p>
    </div>
    """, unsafe_allow_html=True)

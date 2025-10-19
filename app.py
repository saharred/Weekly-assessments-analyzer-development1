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
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    * { font-family: 'Cairo', 'Arial', sans-serif; }
    .main { background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%); }
    .header-container {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%);
        padding: 40px; border-radius: 20px; color: white; text-align: center;
        margin-bottom: 30px; box-shadow: 0 10px 30px rgba(139, 58, 58, 0.25);
        border: 2px solid #D4A574;
    }
    .header-container h1 { margin: 0; font-size: 36px; font-weight: 700; }
    .header-container p { margin: 12px 0 0 0; font-size: 16px; opacity: 0.95; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #8B3A3A 0%, #A0483D 100%);
        box-shadow: 2px 0 15px rgba(0, 0, 0, 0.1);
    }
    [data-testid="stSidebar"] * { color: white; }
    .section-box {
        background: white; padding: 25px; border-radius: 15px; margin: 20px 0;
        border-left: 5px solid #8B3A3A; box-shadow: 0 4px 15px rgba(139, 58, 58, 0.1);
    }
    .metric-box {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%);
        color: white; padding: 25px; border-radius: 12px; text-align: center;
        font-weight: 600; box-shadow: 0 6px 20px rgba(139, 58, 58, 0.2);
        border: 2px solid #D4A574;
    }
    .stButton > button {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%) !important;
        color: white !important; border: 2px solid #D4A574 !important;
        padding: 12px 28px !important; border-radius: 10px !important;
        font-weight: 600 !important; font-size: 14px !important;
    }
    hr { border-color: #8B3A3A !important; margin: 25px 0 !important; }
    h1, h2, h3, h4, h5, h6 { color: #8B3A3A; font-weight: 700; }
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

st.markdown("")
st.markdown("")
st.markdown("")

col1, col2, col3 = st.columns([1, 1.5, 1])
with col3:
    st.image("https://i.imgur.com/jFzu8As.jpeg", width=100)

st.markdown("<div class='header-container'><h1>📊 محلل التقييمات الأسبوعية</h1><p style='font-size: 14px; margin: 10px 0; font-weight: 600;'>وزارة التربية والتعليم والتعليم العالي</p><p style='font-size: 13px; color: #D4A574; font-weight: 600; margin: 5px 0;'>ضمان تنمية رقمية مستدامة</p><p style='font-size: 12px; opacity: 0.9;'>نظام تحليل شامل وموثوق لنتائج الطلاب</p></div>", unsafe_allow_html=True)

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

with st.sidebar:
    col_left, col_mid, col_right = st.columns([1, 1, 1])
    with col_left:
        st.image("https://i.imgur.com/1bX5dzp.jpeg", width=70)
    with col_right:
        st.image("https://i.imgur.com/QfVfT9X.jpeg", width=70)
    
    st.markdown("<div style='text-align: center; margin: 20px 0;'><img src='https://i.imgur.com/3ASAXDc.png' style='width: 120px; height: auto;'></div>", unsafe_allow_html=True)
    
    st.markdown("---")
    st.header("الإعدادات")
    
    st.subheader("تحميل الملفات")
    uploaded_files = st.file_uploader("اختر ملفات Excel", type=["xlsx", "xls"], accept_multiple_files=True)
    
    selected_sheets = []
    if uploaded_files:
        st.success(f"تم رفع {len(uploaded_files)} ملف")
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
                st.error(f"خطأ: {e}")
        
        if all_sheets:
            st.info(f"وجدت {len(all_sheets)} مادة من {len(uploaded_files)} ملفات")
            select_all = st.checkbox("اختر الجميع", value=True)
            if select_all:
                selected_sheets_display = all_sheets
            else:
                selected_sheets_display = st.multiselect("اختر المواد", all_sheets, default=[])
            selected_sheets = [(sheet_file_map[s][0], sheet_file_map[s][1]) for s in selected_sheets_display]
    else:
        st.info("ارفع ملفات Excel للبدء")
    
    st.markdown("---")
    st.subheader("معلومات المدرسة")
    school_name = st.text_input("اسم المدرسة", placeholder="مدرسة قطر النموذجية")
    
    st.subheader("التوقيعات")
    coordinator_name = st.text_input("منسق المشاريع")
    academic_deputy = st.text_input("النائب الأكاديمي")
    admin_deputy = st.text_input("النائب الإداري")
    principal_name = st.text_input("مدير المدرسة")
    
    st.markdown("---")
    run_analysis = st.button("تشغيل التحليل", use_container_width=True, type="primary", disabled=not (uploaded_files and selected_sheets))

if not uploaded_files:
    st.info("ارفع ملفات Excel من الشريط الجانبي")
elif run_analysis:
    with st.spinner("جاري التحليل..."):
        all_results = []
        for file, sheet in selected_sheets:
            results = analyze_excel_file(file, sheet)
            all_results.extend(results)
        
        if all_results:
            df = pd.DataFrame(all_results)
            st.session_state.analysis_results = df
            pivot = create_pivot_table(df)
            st.session_state.pivot_table = pivot
            st.success(f"تم تحليل {len(pivot)} طالب من {len(set(df['subject']))} مادة")

if st.session_state.pivot_table is not None:
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("الطلاب", len(pivot))
    with col2:
        st.metric("المواد", df['subject'].nunique())
    with col3:
        avg = pivot['المتوسط'].mean() if 'المتوسط' in pivot.columns else 0
        st.metric("المتوسط", f"{avg:.1f}%")
    with col4:
        platinum = len(pivot[pivot['الفئة'].str.contains('بلاتينية', na=False)])
        st.metric("بلاتينية", platinum)
    with col5:
        zero = len(pivot[pivot['المتوسط'] == 0])
        st.metric("بدون إنجاز", zero)
    
    st.divider()
    st.subheader("البيانات")
    st.dataframe(pivot, use_container_width=True)
    
    col1, col2 = st.columns(2)
    with col1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot.to_excel(writer, index=False, sheet_name='النتائج')
        st.download_button("تحميل Excel", output.getvalue(), f"results_{datetime.now().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    
    with col2:
        csv_data = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("تحميل CSV", csv_data, f"results_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv", use_container_width=True)

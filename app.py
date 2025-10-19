import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
import base64
from datetime import datetime

# Page config - Enhanced styling
st.set_page_config(
    page_title="محلل التقييمات الأسبوعية",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Styling - Maroon and White theme
st.markdown("""
<style>
    :root {
        --primary-color: #8B3A3A;
        --secondary-color: #FFFFFF;
        --accent-color: #D4A574;
    }
    
    /* Main container styling */
    .main {
        background: linear-gradient(135deg, #f8f9fa 0%, #f0f0f0 100%);
    }
    
    /* Header styling */
    .header-container {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%);
        padding: 30px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 8px 16px rgba(139, 58, 58, 0.2);
    }
    
    .header-container h1 {
        margin: 0;
        font-size: 32px;
        font-weight: bold;
    }
    
    .header-container p {
        margin: 10px 0 0 0;
        font-size: 16px;
        opacity: 0.95;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #8B3A3A 0%, #A0483D 100%);
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: white;
    }
    
    /* Section styling */
    .section-box {
        background: white;
        padding: 25px;
        border-radius: 12px;
        margin: 20px 0;
        border-left: 5px solid #8B3A3A;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
    
    .metric-box {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        font-weight: bold;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #8B3A3A 0%, #A0483D 100%) !important;
        color: white !important;
        border: none !important;
        padding: 12px 24px !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 16px rgba(139, 58, 58, 0.3) !important;
    }
    
    /* Input fields */
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select {
        border: 2px solid #8B3A3A !important;
        border-radius: 8px !important;
    }
    
    .stTextInput > div > div > input:focus,
    .stSelectbox > div > div > select:focus {
        border-color: #D4A574 !important;
        box-shadow: 0 0 8px rgba(139, 58, 58, 0.2) !important;
    }
    
    /* Divider */
    hr {
        border-color: #8B3A3A !important;
    }
    
    /* Info boxes */
    .stInfo, .stSuccess, .stWarning, .stError {
        border-radius: 10px !important;
    }
    
    .stSuccess {
        background-color: #E8F5E9 !important;
        border-left: 5px solid #4CAF50 !important;
    }
</style>
""", unsafe_allow_html=True)

# Ministry Logo as SVG (Qatar Education Ministry)
MINISTRY_LOGO_SVG = """
<svg width="100" height="100" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
    <circle cx="50" cy="50" r="48" fill="#8B3A3A" stroke="#D4A574" stroke-width="2"/>
    <circle cx="50" cy="50" r="42" fill="#FFFFFF"/>
    <text x="50" y="45" font-family="Arial" font-size="24" font-weight="bold" text-anchor="middle" fill="#8B3A3A">وزارة</text>
    <text x="50" y="65" font-family="Arial" font-size="16" text-anchor="middle" fill="#8B3A3A">التعليم</text>
    <circle cx="50" cy="50" r="35" fill="none" stroke="#8B3A3A" stroke-width="1.5" stroke-dasharray="5,5"/>
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
        
        if len(df) > 1:
            level_from_excel = str(df.iloc[1, 1]).strip() if pd.notna(df.iloc[1, 1]) else ""
            section_from_excel = str(df.iloc[1, 2]).strip() if pd.notna(df.iloc[1, 2]) else ""
            
            level = level_from_excel if level_from_excel and level_from_excel != 'nan' else level_from_name
            section = section_from_excel if section_from_excel and section_from_excel != 'nan' else section_from_name
        else:
            level = level_from_name
            section = section_from_name
        
        # جمع التقييمات الصحيحة (ليست فارغة وليست - و —)
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
            
            if total_assessments > 0:
                solve_pct = (completed_count / total_assessments) * 100
            else:
                solve_pct = 0.0
            
            results.append({
                "student_name": student_name_clean,
                "subject": subject,
                "level": str(level).strip(),
                "section": str(section).strip(),
                "solve_pct": solve_pct,
                "completed_count": completed_count,
                "total_count": total_assessments,
                "pending_titles": ", ".join(pending_titles) if pending_titles else ""
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
        
        result = result.merge(
            subject_df,
            on=['student_name', 'level', 'section'],
            how='left'
        )
    
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
    result = result.reset_index(drop=True)
    
    return result

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
                
                subjects_data.append({
                    'subject': subject,
                    'total': total,
                    'completed': completed,
                    'pending': pending_titles,
                    'pct': pct
                })
    
    subjects_html = ""
    for data in subjects_data:
        subjects_html += f"""
        <tr>
            <td style="text-align: right; padding: 12px; border: 1px solid #ddd;">{data['subject']}</td>
            <td style="text-align: center; padding: 12px; border: 1px solid #ddd;">{data['total']}</td>
            <td style="text-align: center; padding: 12px; border: 1px solid #ddd;">{data['completed']}</td>
            <td style="text-align: center; padding: 12px; border: 1px solid #ddd; color: #8B3A3A; font-weight: bold;">{data['pct']:.1f}%</td>
            <td style="text-align: right; padding: 12px; border: 1px solid #ddd; font-size: 12px;">{data['pending']}</td>
        </tr>
        """
    
    solve_pct = (total_completed / total_assessments * 100) if total_assessments > 0 else 0
    remaining = total_assessments - total_completed
    
    if solve_pct == 0:
        recommendation = "الطالب لم يستفيد من النظام - يرجى التواصل مع ولي الأمر فوراً 🚫"
        category_color = "#9E9E9E"
        category_name = "لا يستفيد من النظام"
    elif solve_pct >= 90:
        recommendation = "أداء ممتاز! استمر في التميز 🌟"
        category_color = "#4CAF50"
        category_name = "البلاتينية 🥇"
    elif solve_pct >= 80:
        recommendation = "أداء جيد جداً، حافظ على مستواك 👍"
        category_color = "#8BC34A"
        category_name = "الذهبي 🥈"
    elif solve_pct >= 70:
        recommendation = "أداء جيد، يمكنك التحسن أكثر ✓"
        category_color = "#FFC107"
        category_name = "الفضي 🥉"
    elif solve_pct >= 60:
        recommendation = "أداء مقبول، تحتاج لمزيد من الجهد ⚠️"
        category_color = "#FF9800"
        category_name = "البرونزي"
    else:
        recommendation = "يرجى الاهتمام أكثر بالتقييمات ومراجعة المواد"
        category_color = "#F44336"
        category_name = "يحتاج تحسين ⚠️"
    
    logo_html = ""
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 80px; margin-bottom: 10px;" />'
    
    school_section = f"<h2 style='text-align: center; color: #8B3A3A; margin: 5px 0; font-size: 20px;'>{school_name}</h2>" if school_name else ""
    
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_name}</title>
        <style>
            @page {{ size: A4; margin: 15mm; }}
            body {{ font-family: 'Arial', 'Helvetica', sans-serif; direction: rtl; padding: 20px; background: #f5f5f5; }}
            .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid #8B3A3A; padding-bottom: 20px; margin-bottom: 30px; }}
            .header-right {{ text-align: center; flex: 1; }}
            .header-left {{ text-align: left; }}
            h1 {{ color: #8B3A3A; margin: 10px 0; font-size: 24px; font-weight: bold; }}
            h2 {{ color: #8B3A3A; margin: 5px 0; font-size: 20px; }}
            h3 {{ color: #A0483D; font-size: 16px; }}
            .student-info {{ background: #FFF8F0; padding: 20px; border-radius: 8px; margin-bottom: 25px; border-left: 5px solid #8B3A3A; }}
            .student-info p {{ margin: 8px 0; font-size: 15px; }}
            table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
            th {{ background: #8B3A3A; color: white; padding: 12px; text-align: center; border: 1px solid #8B3A3A; font-size: 13px; }}
            td {{ padding: 12px; border: 1px solid #ddd; font-size: 14px; }}
            tr:nth-child(even) {{ background-color: #fafafa; }}
            .stats-section {{ background: #FFF8F0; padding: 20px; border-radius: 8px; margin: 25px 0; border-left: 5px solid #8B3A3A; }}
            .stats-section h3 {{ color: #8B3A3A; margin-top: 0; }}
            .stats-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-top: 15px; }}
            .stat-box {{ background: white; padding: 15px; border-radius: 8px; text-align: center; box-shadow: 0 2px 5px rgba(139, 58, 58, 0.1); border-top: 3px solid #8B3A3A; }}
            .stat-value {{ font-size: 24px; font-weight: bold; color: #8B3A3A; }}
            .stat-label {{ font-size: 12px; color: #666; margin-top: 5px; }}
            .recommendation {{ background: {category_color}; color: white; padding: 20px; border-radius: 8px; margin: 25px 0; text-align: center; font-size: 16px; font-weight: bold; }}
            .category-badge {{ background: #8B3A3A; color: white; padding: 10px 15px; border-radius: 5px; display: inline-block; margin: 10px 0; font-weight: bold; }}
            .signatures {{ margin-top: 40px; border-top: 2px solid #8B3A3A; padding-top: 20px; }}
            .signature-line {{ margin: 15px 0; font-size: 14px; }}
            @media print {{
                body {{ background: white; padding: 0; }}
                .container {{ box-shadow: none; max-width: 100%; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="header-left">
                    {logo_html}
                </div>
                <div class="header-right">
                    {school_section}
                    <h1>📊 تقرير أداء الطالب</h1>
                    <p style="color: #8B3A3A; font-size: 13px; font-weight: bold;">وزارة التربية والتعليم والتعليم العالي</p>
                </div>
            </div>
            
            <div class="student-info">
                <h3>📋 معلومات الطالب</h3>
                <p><strong>اسم الطالب:</strong> {student_name}</p>
                <p><strong>الصف:</strong> {level} &nbsp;&nbsp;&nbsp; <strong>الشعبة:</strong> {section}</p>
                <div class="category-badge">الفئة: {category_name}</div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>المادة</th>
                        <th>إجمالي التقييمات</th>
                        <th>المنجز</th>
                        <th>نسبة الإنجاز %</th>
                        <th>التقييمات المتبقية</th>
                    </tr>
                </thead>
                <tbody>
                    {subjects_html}
                </tbody>
            </table>
            
            <div class="stats-section">
                <h3>📊 الإحصائيات الإجمالية</h3>
                <div class="stats-grid">
                    <div class="stat-box">
                        <div class="stat-label">إجمالي التقييمات</div>
                        <div class="stat-value">{total_assessments}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">المنجز</div>
                        <div class="stat-value">{total_completed}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">المتبقي</div>
                        <div class="stat-value">{remaining}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">نسبة الإنجاز</div>
                        <div class="stat-value">{solve_pct:.1f}%</div>
                    </div>
                </div>
            </div>
            
            <div class="recommendation">
                💡 {recommendation}
            </div>
            
            <div class="signatures">
                <div class="signature-line"><strong>منسق المشاريع/</strong> {coordinator if coordinator else "_____________"}</div>
                <div class="signature-line">
                    <strong>النائب الأكاديمي/</strong> {academic if academic else "_____________"} &nbsp;&nbsp;&nbsp;
                    <strong>النائب الإداري/</strong> {admin if admin else "_____________"}
                </div>
                <div class="signature-line"><strong>مدير المدرسة/</strong> {principal if principal else "_____________"}</div>
                
                <p style="text-align: center; color: #999; margin-top: 30px; font-size: 11px;">
                    تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                </p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html

# ================== MAIN APP ==================

# Header
st.markdown(f"""
<div class="header-container">
    <h1>📊 محلل التقييمات الأسبوعية</h1>
    <p>نظام تحليل شامل لنتائج الطلاب | وزارة التربية والتعليم والتعليم العالي</p>
</div>
""", unsafe_allow_html=True)

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

# Sidebar - Settings
with st.sidebar:
    st.markdown("---")
    st.markdown(f"{MINISTRY_LOGO_SVG}", unsafe_allow_html=True)
    st.markdown("---")
    
    st.header("⚙️ الإعدادات")
    
    # File Upload Section
    st.subheader("📁 تحميل الملفات")
    
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel (يمكنك اختيار أكثر من ملف)",
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
                try:
                    xls = pd.ExcelFile(file)
                    sheets = xls.sheet_names
                    for sheet in sheets:
                        sheet_display = f"[ملف {file_idx+1}] {sheet}"
                        all_sheets.append(sheet_display)
                        sheet_file_map[sheet_display] = (file, sheet)
                except Exception as e:
                    st.warning(f"⚠️ خطأ في قراءة الملف")
            
            if all_sheets:
                st.info(f"📊 وجدت {len(all_sheets)} مادة من {len(uploaded_files)} ملفات")
                
                selected_sheets_display = st.multiselect(
                    "اختر المواد المراد تحليلها",
                    all_sheets,
                    default=all_sheets
                )
                
                selected_sheets = [(sheet_file_map[s][0], sheet_file_map[s][1]) for s in selected_sheets_display]
            else:
                st.error("❌ لم يتم العثور على أي مواد")
                selected_sheets = []
        except Exception as e:
            st.error(f"خطأ: {e}")
            selected_sheets = []
    else:
        st.info("💡 ارفع ملفات Excel لبدء التحليل")
        selected_sheets = []
    
    st.divider()
    
    # School and Signatures Settings
    st.subheader("🏫 معلومات المدرسة")
    
    school_name = st.text_input(
        "📛 اسم المدرسة",
        value="",
        placeholder="مثال: مدرسة قطر النموذجية"
    )
    
    st.subheader("🖼️ شعار الوزارة/المدرسة")
    uploaded_logo = st.file_uploader(
        "ارفع شعار (اختياري)",
        type=["png", "jpg", "jpeg"],
        help="سيظهر الشعار في رأس التقارير"
    )
    
    logo_base64 = ""
    if uploaded_logo:
        logo_bytes = uploaded_logo.read()
        logo_base64 = base64.b64encode(logo_bytes).decode()
        st.success("✅ تم رفع الشعار")
    
    st.divider()
    
    st.subheader("✍️ معلومات التوقيعات")
    
    coordinator_name = st.text_input(
        "👤 منسق المشاريع",
        value="",
        placeholder="أدخل اسم منسق المشاريع"
    )
    
    academic_deputy = st.text_input(
        "👨‍🏫 النائب الأكاديمي",
        value="",
        placeholder="أدخل اسم النائب الأكاديمي"
    )
    
    admin_deputy = st.text_input(
        "👨‍💼 النائب الإداري",
        value="",
        placeholder="أدخل اسم النائب الإداري"
    )
    
    principal_name = st.text_input(
        "🎓 مدير المدرسة",
        value="",
        placeholder="أدخل اسم مدير المدرسة"

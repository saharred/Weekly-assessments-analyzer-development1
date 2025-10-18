import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
from datetime import datetime

# Page config
st.set_page_config(
    page_title="محلل التقييمات الأسبوعية",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
        
        assessment_titles = []
        for col_idx in range(7, df.shape[1]):
            title = df.iloc[0, col_idx]
            if pd.notna(title) and str(title).strip():
                assessment_titles.append(str(title).strip())
        
        total_assessments = len(assessment_titles)
        
        results = []
        
        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]
            
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue
            
            student_name_clean = " ".join(str(student_name).strip().split())
            
            m_count = 0
            pending_titles = []
            
            for i, col_idx in enumerate(range(7, 7 + total_assessments)):
                if col_idx < df.shape[1]:
                    cell_value = df.iloc[idx, col_idx]
                    
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip().upper()
                        
                        if cell_str == 'M':
                            m_count += 1
                            if i < len(assessment_titles):
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
    
    # Remove any duplicate entries in raw data first
    df_clean = df.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'], keep='first')
    
    # Get unique students only
    unique_students = df_clean[['student_name', 'level', 'section']].drop_duplicates()
    unique_students = unique_students.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
    
    result = unique_students.copy()
    
    subjects = sorted(df_clean['subject'].unique())
    
    for subject in subjects:
        subject_df = df_clean[df_clean['subject'] == subject][['student_name', 'level', 'section', 
                                                                'total_count', 'completed_count', 
                                                                'pending_titles', 'solve_pct']].copy()
        
        # Ensure no duplicates per subject
        subject_df = subject_df.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')
        
        # Rename columns with subject prefix
        subject_df = subject_df.rename(columns={
            'total_count': f"{subject} - إجمالي التقييمات",
            'completed_count': f"{subject} - المنجز",
            'pending_titles': f"{subject} - عناوين التقييمات المتبقية",
            'solve_pct': f"{subject} - نسبة الإنجاز %"
        })
        
        # Merge
        result = result.merge(
            subject_df,
            on=['student_name', 'level', 'section'],
            how='left'
        )
    
    # Calculate overall percentage
    pct_cols = [col for col in result.columns if 'نسبة الإنجاز %' in col]
    if pct_cols:
        result['نسبة حل التقييمات في جميع المواد'] = result[pct_cols].mean(axis=1)
        
        # Add category based on overall percentage
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
    
    # Rename to Arabic
    result = result.rename(columns={
        'student_name': 'اسم الطالب',
        'level': 'الصف',
        'section': 'الشعبة'
    })
    
    # Final check - remove any remaining duplicates
    result = result.drop_duplicates(subset=['اسم الطالب', 'الصف', 'الشعبة'], keep='first')
    result = result.reset_index(drop=True)
    
    return result

def generate_student_html_report(student_row, school_name="", coordinator="", academic="", admin="", principal="", logo_base64=""):
    """Generate individual student HTML report with customizable signatures and logo"""
    
    student_name = student_row['اسم الطالب']
    level = student_row['الصف']
    section = student_row['الشعبة']
    
    total_assessments = 0
    total_completed = 0
    
    subjects_html = ""
    
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
                
                total_assessments += total
                total_completed += completed
                
                subjects_html += f"""
                <tr>
                    <td style="text-align: right; padding: 12px; border: 1px solid #ddd;">{subject}</td>
                    <td style="text-align: center; padding: 12px; border: 1px solid #ddd;">{total}</td>
                    <td style="text-align: center; padding: 12px; border: 1px solid #ddd;">{completed}</td>
                    <td style="text-align: right; padding: 12px; border: 1px solid #ddd;">{pending_titles}</td>
                </tr>
                """
    
    solve_pct = (total_completed / total_assessments * 100) if total_assessments > 0 else 0
    remaining = total_assessments - total_completed
    
    if solve_pct == 0:
        recommendation = "الطالب لم يستفيد من النظام - يرجى التواصل مع ولي الأمر فوراً 🚫"
        category_color = "#9E9E9E"
    elif solve_pct >= 90:
        recommendation = "أداء ممتاز! استمر في التميز 🌟"
        category_color = "#4CAF50"
    elif solve_pct >= 80:
        recommendation = "أداء جيد جداً، حافظ على مستواك 👍"
        category_color = "#8BC34A"
    elif solve_pct >= 70:
        recommendation = "أداء جيد، يمكنك التحسن أكثر ✓"
        category_color = "#FFC107"
    elif solve_pct >= 60:
        recommendation = "أداء مقبول، تحتاج لمزيد من الجهد ⚠️"
        category_color = "#FF9800"
    else:
        recommendation = "يرجى الاهتمام أكثر بالتقييمات ومراجعة المواد"
        category_color = "#F44336"
    
    # Logo section
    logo_html = ""
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 80px; margin-bottom: 10px;" />'
    
    # School name section
    school_section = f"<h2 style='text-align: center; color: #1976D2; margin: 5px 0;'>{school_name}</h2>" if school_name else ""
    
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_name}</title>
        <style>
            @page {{ size: A4; margin: 15mm; }}
            body {{ font-family: Arial, sans-serif; direction: rtl; padding: 20px; background: #f5f5f5; }}
            .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            .header {{ text-align: center; border-bottom: 3px solid #1976D2; padding-bottom: 20px; margin-bottom: 30px; }}
            h1 {{ color: #1976D2; margin: 10px 0; font-size: 24px; }}
            h2 {{ color: #1976D2; margin: 5px 0; font-size: 20px; }}
            .student-info {{ background: #E3F2FD; padding: 20px; border-radius: 8px; margin-bottom: 25px; }}
            .student-info h3 {{ margin-top: 0; color: #1565C0; }}
            table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
            th {{ background: #1976D2; color: white; padding: 12px; text-align: center; border: 1px solid #1565C0; font-size: 14px; }}
            td {{ padding: 12px; border: 1px solid #ddd; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .stats-section {{ background: #FFF3E0; padding: 20px; border-radius: 8px; margin: 25px 0; }}
            .stats-section h3 {{ color: #E65100; margin-top: 0; }}
            .stats-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-top: 15px; }}
            .stat-box {{ background: white; padding: 15px; border-radius: 8px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            .stat-value {{ font-size: 28px; font-weight: bold; color: {category_color}; }}
            .stat-label {{ font-size: 13px; color: #666; margin-top: 5px; }}
            .recommendation {{ background: {category_color}; color: white; padding: 20px; border-radius: 8px; margin: 25px 0; text-align: center; font-size: 16px; font-weight: bold; }}
            .signatures {{ margin-top: 40px; border-top: 2px solid #ddd; padding-top: 20px; }}
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
                {logo_html}
                {school_section}
                <h1>📊 تقرير أداء الطالب - نظام قطر للتعليم</h1>
            </div>
            
            <div class="student-info">
                <h3>معلومات الطالب</h3>
                <p><strong>اسم الطالب:</strong> {student_name}</p>
                <p><strong>الصف:</strong> {level} &nbsp;&nbsp;&nbsp; <strong>الشعبة:</strong> {section}</p>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>المادة</th>
                        <th>عدد التقييمات الإجمالي</th>
                        <th>عدد التقييمات المنجزة</th>
                        <th>عنوان التقييمات المتبقية</th>
                    </tr>
                </thead>
                <tbody>
                    {subjects_html}
                </tbody>
            </table>
            
            <div class="stats-section">
                <h3>الإحصائيات</h3>
                <div class="stats-grid">
                    <div class="stat-box">
                        <div class="stat-label">منجز</div>
                        <div class="stat-value">{total_completed}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">متبقي</div>
                        <div class="stat-value">{remaining}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">نسبة حل التقييمات</div>
                        <div class="stat-value">{solve_pct:.1f}%</div>
                    </div>
                </div>
            </div>
            
            <div class="recommendation">
                توصية منسق المشاريع: {recommendation}
            </div>
            
            <div class="signatures">
                <div class="signature-line"><strong>منسق المشاريع/</strong> {coordinator if coordinator else "_____________"}</div>
                <div class="signature-line">
                    <strong>النائب الأكاديمي/</strong> {academic if academic else "_____________"} &nbsp;&nbsp;&nbsp;
                    <strong>النائب الإداري/</strong> {admin if admin else "_____________"}
                </div>
                <div class="signature-line"><strong>مدير المدرسة/</strong> {principal if principal else "_____________"}</div>
                
                <p style="text-align: center; color: #999; margin-top: 30px; font-size: 12px;">
                    تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}
                </p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html
    
    student_name = student_row['اسم الطالب']
    level = student_row['الصف']
    section = student_row['الشعبة']
    
    total_assessments = 0
    total_completed = 0
    
    subjects_html = ""
    
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
                
                total_assessments += total
                total_completed += completed
                
                subjects_html += f"""
                <tr>
                    <td style="text-align: right; padding: 12px; border: 1px solid #ddd;">{subject}</td>
                    <td style="text-align: center; padding: 12px; border: 1px solid #ddd;">{total}</td>
                    <td style="text-align: center; padding: 12px; border: 1px solid #ddd;">{completed}</td>
                    <td style="text-align: right; padding: 12px; border: 1px solid #ddd;">{pending_titles}</td>
                </tr>
                """
    
    solve_pct = (total_completed / total_assessments * 100) if total_assessments > 0 else 0
    remaining = total_assessments - total_completed
    
    if solve_pct >= 90:
        recommendation = "أداء ممتاز! استمر في التميز 🌟"
        category_color = "#4CAF50"
    elif solve_pct >= 80:
        recommendation = "أداء جيد جداً، حافظ على مستواك 👍"
        category_color = "#8BC34A"
    elif solve_pct >= 70:
        recommendation = "أداء جيد، يمكنك التحسن أكثر ✓"
        category_color = "#FFC107"
    elif solve_pct >= 60:
        recommendation = "أداء مقبول، تحتاج لمزيد من الجهد ⚠️"
        category_color = "#FF9800"
    else:
        recommendation = "يرجى الاهتمام أكثر بالتقييمات ومراجعة المواد"
        category_color = "#F44336"
    
    # School name section
    school_section = f"<h2 style='text-align: center; color: #1976D2;'>{school_name}</h2>" if school_name else ""
    
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_name}</title>
        <style>
            body {{ font-family: Arial, sans-serif; direction: rtl; padding: 20px; background: #f5f5f5; }}
            .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            .header {{ text-align: center; border-bottom: 3px solid #1976D2; padding-bottom: 20px; margin-bottom: 30px; }}
            h1 {{ color: #1976D2; margin: 10px 0; }}
            h2 {{ color: #1976D2; margin: 5px 0; }}
            .student-info {{ background: #E3F2FD; padding: 20px; border-radius: 8px; margin-bottom: 25px; }}
            table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
            th {{ background: #1976D2; color: white; padding: 12px; text-align: center; border: 1px solid #1565C0; }}
            td {{ padding: 12px; border: 1px solid #ddd; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .stats-section {{ background: #FFF3E0; padding: 20px; border-radius: 8px; margin: 25px 0; }}
            .stat-value {{ font-size: 32px; font-weight: bold; color: {category_color}; }}
            .recommendation {{ background: {category_color}; color: white; padding: 20px; border-radius: 8px; margin: 25px 0; text-align: center; font-size: 18px; }}
            .signatures {{ margin-top: 40px; border-top: 2px solid #ddd; padding-top: 20px; }}
            .signature-line {{ margin: 15px 0; font-size: 15px; }}
            @media print {{
                body {{ background: white; padding: 0; }}
                .container {{ box-shadow: none; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                {school_section}
                <h1>📊 تقرير أداء الطالب - نظام قطر للتعليم</h1>
            </div>
            
            <div class="student-info">
                <h2>معلومات الطالب</h2>
                <p><strong>اسم الطالب:</strong> {student_name}</p>
                <p><strong>الصف:</strong> {level} &nbsp;&nbsp; <strong>الشعبة:</strong> {section}</p>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>المادة</th>
                        <th>عدد التقييمات الإجمالي</th>
                        <th>عدد التقييمات المنجزة</th>
                        <th>عنوان التقييمات المتبقية</th>
                    </tr>
                </thead>
                <tbody>
                    {subjects_html}
                </tbody>
            </table>
            
            <div class="stats-section">
                <h3>الإحصائيات</h3>
                <p><strong>منجز:</strong> <span class="stat-value">{total_completed}</span></p>
                <p><strong>متبقي:</strong> <span class="stat-value">{remaining}</span></p>
                <p><strong>نسبة حل التقييمات:</strong> <span class="stat-value">{solve_pct:.1f}%</span></p>
            </div>
            
            <div class="recommendation">
                توصية منسق المشاريع: {recommendation}
            </div>
            
            <div class="signatures">
                <div class="signature-line"><strong>منسق المشاريع/</strong> {coordinator if coordinator else "_____________"}</div>
                <div class="signature-line">
                    <strong>النائب الأكاديمي/</strong> {academic if academic else "_____________"} &nbsp;&nbsp;&nbsp;
                    <strong>النائب الإداري/</strong> {admin if admin else "_____________"}
                </div>
                <div class="signature-line"><strong>مدير المدرسة/</strong> {principal if principal else "_____________"}</div>
                
                <p style="text-align: center; color: #999; margin-top: 30px; font-size: 12px;">
                    تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}
                </p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html

# ================== MAIN APP ==================

st.title("📊 محلل التقييمات الأسبوعية")
st.markdown("---")

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

with st.sidebar:
    st.header("⚙️ الإعدادات")
    
    st.subheader("📁 تحميل الملفات")
    
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        
        try:
            xls = pd.ExcelFile(uploaded_files[0])
            sheets = xls.sheet_names
            
            selected_sheets = st.multiselect(
                "اختر الأوراق (المواد)",
                sheets,
                default=sheets
            )
        except Exception as e:
            st.error(f"خطأ: {e}")
            selected_sheets = []
    else:
        selected_sheets = []
    
    st.divider()
    
    # Charts Section
    st.subheader("📈 الرسوم البيانية")
    
    import matplotlib
    matplotlib.rcParams['axes.unicode_minus'] = False
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📊 متوسط الإنجاز حسب المادة**")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        subject_avg = df.groupby('subject')['solve_pct'].mean().sort_values(ascending=True)
        
        colors = plt.cm.viridis(range(len(subject_avg)))
        y_pos = range(len(subject_avg))
        bars = ax.barh(y_pos, subject_avg.values, color=colors, edgecolor='black', linewidth=1.5)
        
        for i, (bar, value) in enumerate(zip(bars, subject_avg.values)):
            ax.text(value + 2, i, f'{value:.1f}%', va='center', fontsize=11, fontweight='bold')
        
        ax.set_yticks(y_pos)
        ax.set_yticklabels([f"#{i+1}" for i in y_pos], fontsize=10)
        ax.set_xlabel("Average Completion Rate (%)", fontsize=12, fontweight='bold')
        ax.set_title("Performance by Subject", fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='x', alpha=0.3, linestyle='--')
        ax.set_xlim(0, 110)
        
        ax.axvspan(0, 60, alpha=0.1, color='red')
        ax.axvspan(60, 80, alpha=0.1, color='yellow')
        ax.axvspan(80, 100, alpha=0.1, color='green')
        
        plt.tight_layout()
        st.pyplot(fig)
        
        st.caption("**المواد:**")
        for i, subj in enumerate(subject_avg.index):
            st.caption(f"#{i+1}: {subj} ({subject_avg.values[i]:.1f}%)")
    
    with col2:
        st.markdown("**📈 توزيع النسب الإجمالية**")
        
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            fig, ax = plt.subplots(figsize=(10, 6))
            
            overall_scores = pivot['نسبة حل التقييمات في جميع المواد'].dropna()
            
            n, bins, patches = ax.hist(overall_scores, bins=20, edgecolor='black', linewidth=1.5)
            
            for i, patch in enumerate(patches):
                bin_center = (bins[i] + bins[i+1]) / 2
                if bin_center >= 80:
                    patch.set_facecolor('#4CAF50')
                elif bin_center >= 60:
                    patch.set_facecolor('#FFC107')
                else:
                    patch.set_facecolor('#F44336')
            
            mean_val = overall_scores.mean()
            ax.axvline(mean_val, color='blue', linestyle='--', linewidth=2.5, 
                      label=f'Average: {mean_val:.1f}%', zorder=10)
            
            ax.set_xlabel("Completion Rate (%)", fontsize=12, fontweight='bold')
            ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
            ax.set_title("Overall Performance Distribution", fontsize=14, fontweight='bold', pad=20)
            ax.legend(fontsize=11, loc='upper left')
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            
            plt.tight_layout()
            st.pyplot(fig)
    
    st.divider()
    
    # Subject Analysis
    st.subheader("📚 التحليل حسب المادة")
    
    subjects = sorted(df['subject'].unique())
    
    selected_subject = st.selectbox(
        "اختر المادة للتحليل التفصيلي:",
        subjects,
        key="subject_analysis"
    )
    
    if selected_subject:
        subject_df = df[df['subject'] == selected_subject]
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("👥 عدد الطلاب", len(subject_df))
        with col2:
            st.metric("📈 متوسط الإنجاز", f"{subject_df['solve_pct'].mean():.1f}%")
        with col3:
            st.metric("🏆 أعلى نسبة", f"{subject_df['solve_pct'].max():.1f}%")
        with col4:
            st.metric("⚠️ أقل نسبة", f"{subject_df['solve_pct'].min():.1f}%")
        
        st.markdown("#### 📊 توزيع الطلاب")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            excellent = len(subject_df[subject_df['solve_pct'] >= 90])
            st.metric("ممتاز (90%+)", excellent, 
                     delta=f"{excellent/len(subject_df)*100:.1f}%" if len(subject_df) > 0 else "0%")
        
        with col2:
            good = len(subject_df[(subject_df['solve_pct'] >= 70) & (subject_df['solve_pct'] < 90)])
            st.metric("جيد (70-89%)", good,
                     delta=f"{good/len(subject_df)*100:.1f}%" if len(subject_df) > 0 else "0%")
        
        with col3:
            weak = len(subject_df[subject_df['solve_pct'] < 70])
            st.metric("يحتاج دعم (<70%)", weak,
                     delta=f"{weak/len(subject_df)*100:.1f}%" if len(subject_df) > 0 else "0%",
                     delta_color="inverse")
        
        # Top and Bottom students
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### 🌟 أفضل 5 طلاب")
            top_5 = subject_df.nlargest(5, 'solve_pct')[['student_name', 'solve_pct', 'completed_count', 'total_count']]
            for idx, row in top_5.iterrows():
                st.text(f"• {row['student_name']}: {row['solve_pct']:.1f}% ({row['completed_count']}/{row['total_count']})")
        
        with col2:
            st.markdown("##### ⚠️ يحتاجون دعم (أقل 5)")
            bottom_5 = subject_df.nsmallest(5, 'solve_pct')[['student_name', 'solve_pct', 'completed_count', 'total_count']]
            for idx, row in bottom_5.iterrows():
                st.text(f"• {row['student_name']}: {row['solve_pct']:.1f}% ({row['completed_count']}/{row['total_count']})")
        
        # Chart for this subject
        st.markdown("##### 📊 رسم بياني للمادة")
        
        fig, ax = plt.subplots(figsize=(12, 5))
        
        categories = pd.cut(subject_df['solve_pct'], 
                           bins=[0, 50, 70, 80, 90, 100], 
                           labels=['<50%', '50-70%', '70-80%', '80-90%', '90-100%'])
        
        category_counts = categories.value_counts().sort_index()
        
        colors_cat = ['#F44336', '#FF9800', '#FFC107', '#8BC34A', '#4CAF50']
        bars = ax.bar(range(len(category_counts)), category_counts.values, 
                     color=colors_cat, edgecolor='black', linewidth=1.5)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.3,
                   f'{int(height)}',
                   ha='center', va='bottom', fontsize=12, fontweight='bold')
        
        ax.set_xticks(range(len(category_counts)))
        ax.set_xticklabels(category_counts.index, fontsize=11)
        ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
        ax.set_title(f"Performance Distribution - {selected_subject}", fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        
        plt.tight_layout()
        st.pyplot(fig)
    
    st.divider()
    
    # School and Signatures Settings
    st.subheader("🏫 معلومات المدرسة")
    
    school_name = st.text_input(
        "اسم المدرسة",
        value="",
        placeholder="مثال: مدرسة قطر النموذجية"
    )
    
    # Logo upload
    st.subheader("🖼️ شعار الوزارة/المدرسة")
    uploaded_logo = st.file_uploader(
        "ارفع شعار (اختياري)",
        type=["png", "jpg", "jpeg"],
        help="سيظهر الشعار في رأس التقارير"
    )
    
    # Convert logo to base64 if uploaded
    logo_base64 = ""
    if uploaded_logo:
        import base64
        logo_bytes = uploaded_logo.read()
        logo_base64 = base64.b64encode(logo_bytes).decode()
        st.success("✅ تم رفع الشعار")
    
    st.subheader("✍️ التوقيعات")
    
    coordinator_name = st.text_input(
        "منسق المشاريع",
        value="سحر عثمان",
        placeholder="اسم منسق المشاريع"
    )
    
    academic_deputy = st.text_input(
        "النائب الأكاديمي",
        value="مريم القضع",
        placeholder="اسم النائب الأكاديمي"
    )
    
    admin_deputy = st.text_input(
        "النائب الإداري",
        value="دلال الفهيدة",
        placeholder="اسم النائب الإداري"
    )
    
    principal_name = st.text_input(
        "مدير المدرسة",
        value="منيرة الهاجري",
        placeholder="اسم مدير المدرسة"
    )
    
    st.divider()
    
    run_analysis = st.button(
        "🚀 تشغيل التحليل",
        use_container_width=True,
        type="primary",
        disabled=not (uploaded_files and selected_sheets)
    )

if not uploaded_files:
    st.info("👈 الرجاء رفع ملفات Excel من الشريط الجانبي")

elif run_analysis:
    with st.spinner("جاري التحليل..."):
        try:
            all_results = []
            progress_bar = st.progress(0)
            total_sheets = len(uploaded_files) * len(selected_sheets)
            current = 0
            
            for file in uploaded_files:
                for sheet in selected_sheets:
                    results = analyze_excel_file(file, sheet)
                    all_results.extend(results)
                    current += 1
                    progress_bar.progress(current / total_sheets)
            
            if all_results:
                df = pd.DataFrame(all_results)
                st.session_state.analysis_results = df
                
                pivot = create_pivot_table(df)
                st.session_state.pivot_table = pivot
                
                st.success(f"✅ تم تحليل {len(pivot)} طالب فريد من {len(selected_sheets)} مادة!")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")
        
        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")

if st.session_state.pivot_table is not None:
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results
    
    st.markdown("## 📈 الإحصائيات العامة")
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("👥 عدد الطلاب", len(pivot))
    with col2:
        st.metric("📚 عدد المواد", df['subject'].nunique())
    with col3:
        avg = pivot['نسبة حل التقييمات في جميع المواد'].mean() if 'نسبة حل التقييمات في جميع المواد' in pivot.columns else 0
        st.metric("📈 متوسط النسبة", f"{avg:.1f}%")
    with col4:
        platinum = len(pivot[pivot['الفئة'].str.contains('البلاتينية', na=False)]) if 'الفئة' in pivot.columns else 0
        st.metric("🥇 البلاتينية", platinum)
    with col5:
        not_using = len(pivot[pivot['الفئة'].str.contains('لا يستفيد', na=False)]) if 'الفئة' in pivot.columns else 0
        needs_improvement = len(pivot[pivot['الفئة'].str.contains('يحتاج تحسين', na=False)]) if 'الفئة' in pivot.columns else 0
        st.metric("⚠️ يحتاج تحسين", needs_improvement)
    
    # Additional metrics row
    if not_using > 0:
        st.warning(f"🚫 **تنبيه:** {not_using} طالب لا يستفيد من النظام (نسبة الإنجاز 0%)")
    
    st.divider()
    
    st.subheader("📊 البيانات التفصيلية")
    
    display_pivot = pivot.copy()
    
    for col in display_pivot.columns:
        if 'نسبة' in col:
            display_pivot[col] = display_pivot[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    
    display_pivot = display_pivot.fillna("-")
    
    st.dataframe(display_pivot, use_container_width=True, height=500)
    
    st.markdown("### 📥 تنزيل النتائج")
    
    col1, col2 = st.columns(2)
    
    with col1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot.to_excel(writer, index=False, sheet_name='النتائج')
        
        st.download_button(
            "📊 تحميل Excel",
            output.getvalue(),
            f"results_{datetime.now().strftime('%Y%m%d')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        csv_data = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📄 تحميل CSV",
            csv_data,
            f"results_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    st.divider()
    
    st.subheader("📄 تقارير الطلاب الفردية")
    
    col1, col2 = st.columns(2)
    
    with col1:
        report_type = st.radio("نوع التقرير:", ["طالب واحد", "جميع الطلاب"])
    
    with col2:
        if report_type == "طالب واحد":
            selected_student = st.selectbox("اختر الطالب:", pivot['اسم الطالب'].tolist())
    
    if st.button("🔄 إنشاء التقارير", use_container_width=True):
        # Get settings from sidebar
        settings = {
            'school': school_name,
            'coordinator': coordinator_name,
            'academic': academic_deputy,
            'admin': admin_deputy,
            'principal': principal_name,
            'logo': logo_base64
        }
        
        if report_type == "طالب واحد":
            student_row = pivot[pivot['اسم الطالب'] == selected_student].iloc[0]
            html = generate_student_html_report(
                student_row,
                settings['school'],
                settings['coordinator'],
                settings['academic'],
                settings['admin'],
                settings['principal'],
                settings['logo']
            )
            
            st.download_button(
                f"📥 تحميل تقرير {selected_student}",
                html.encode('utf-8'),
                f"تقرير_{selected_student}.html",
                "text/html",
                use_container_width=True
            )
            
            st.success("✅ تم إنشاء التقرير!")
        else:
            with st.spinner(f"جاري إنشاء {len(pivot)} تقرير..."):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for _, row in pivot.iterrows():
                        html = generate_student_html_report(
                            row,
                            settings['school'],
                            settings['coordinator'],
                            settings['academic'],
                            settings['admin'],
                            settings['principal'],
                            settings['logo']
                        )
                        filename = f"تقرير_{row['اسم الطالب']}.html"
                        zf.writestr(filename, html.encode('utf-8'))
                
                st.download_button(
                    f"📦 تحميل جميع التقارير ({len(pivot)})",
                    zip_buffer.getvalue(),
                    f"reports_{datetime.now().strftime('%Y%m%d')}.zip",
                    "application/zip",
                    use_container_width=True
                )
                
                st.success(f"✅ تم إنشاء {len(pivot)} تقرير بنجاح!")

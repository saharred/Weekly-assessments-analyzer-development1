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
    """Create pivot table - ONE ROW PER STUDENT"""
    
    # Group by student to ensure uniqueness
    unique_students = df.groupby(['student_name', 'level', 'section']).first().reset_index()[['student_name', 'level', 'section']]
    
    result = unique_students.copy()
    
    subjects = sorted(df['subject'].unique())
    
    for subject in subjects:
        subject_df = df[df['subject'] == subject].copy()
        
        # Group by student and take first (should be only one per student per subject)
        subject_grouped = subject_df.groupby(['student_name', 'level', 'section']).first().reset_index()
        
        # Rename columns
        subject_grouped = subject_grouped.rename(columns={
            'total_count': f"{subject} - إجمالي التقييمات",
            'completed_count': f"{subject} - المنجز",
            'pending_titles': f"{subject} - عناوين التقييمات المتبقية",
            'solve_pct': f"{subject} - نسبة الإنجاز %"
        })
        
        # Merge
        result = result.merge(
            subject_grouped[['student_name', 'level', 'section', 
                           f"{subject} - إجمالي التقييمات",
                           f"{subject} - المنجز", 
                           f"{subject} - عناوين التقييمات المتبقية",
                           f"{subject} - نسبة الإنجاز %"]],
            on=['student_name', 'level', 'section'],
            how='left'
        )
    
    # Calculate overall
    pct_cols = [col for col in result.columns if 'نسبة الإنجاز %' in col]
    if pct_cols:
        result['نسبة حل التقييمات في جميع المواد'] = result[pct_cols].mean(axis=1)
    
    result = result.rename(columns={
        'student_name': 'اسم الطالب',
        'level': 'الصف',
        'section': 'الشعبة'
    })
    
    return result

def generate_student_html_report(student_row):
    """Generate individual student HTML report"""
    
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
    
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_name}</title>
        <style>
            body {{ font-family: Arial, sans-serif; direction: rtl; padding: 20px; }}
            .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 30px; }}
            .header {{ text-align: center; border-bottom: 3px solid #1976D2; padding-bottom: 20px; margin-bottom: 30px; }}
            h1 {{ color: #1976D2; }}
            .student-info {{ background: #E3F2FD; padding: 20px; border-radius: 8px; margin-bottom: 25px; }}
            table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
            th {{ background: #1976D2; color: white; padding: 12px; text-align: center; border: 1px solid #1565C0; }}
            td {{ padding: 12px; border: 1px solid #ddd; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .stats-section {{ background: #FFF3E0; padding: 20px; border-radius: 8px; margin: 25px 0; }}
            .stat-value {{ font-size: 32px; font-weight: bold; color: {category_color}; }}
            .recommendation {{ background: {category_color}; color: white; padding: 20px; border-radius: 8px; margin: 25px 0; text-align: center; font-size: 18px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
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
            
            <div style="margin-top: 40px; border-top: 2px solid #ddd; padding-top: 20px;">
                <p>منسق المشاريع/ سحر عثمان</p>
                <p>النائب الأكاديمي/ مريم القضع &nbsp;&nbsp; النائب الإداري/ دلال الفهيدة</p>
                <p>مدير المدرسة/ منيرة الهاجري</p>
                <p style="text-align: center; color: #999; margin-top: 20px;">تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}</p>
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
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("👥 عدد الطلاب", len(pivot))
    with col2:
        st.metric("📚 عدد المواد", df['subject'].nunique())
    with col3:
        avg = pivot['نسبة حل التقييمات في جميع المواد'].mean() if 'نسبة حل التقييمات في جميع المواد' in pivot.columns else 0
        st.metric("📈 متوسط النسبة", f"{avg:.1f}%")
    with col4:
        st.metric("📋 السجلات", len(df))
    
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
        if report_type == "طالب واحد":
            student_row = pivot[pivot['اسم الطالب'] == selected_student].iloc[0]
            html = generate_student_html_report(student_row)
            
            st.download_button(
                f"📥 تحميل تقرير {selected_student}",
                html.encode('utf-8'),
                f"تقرير_{selected_student}.html",
                "text/html",
                use_container_width=True
            )
        else:
            with st.spinner(f"جاري إنشاء {len(pivot)} تقرير..."):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for _, row in pivot.iterrows():
                        html = generate_student_html_report(row)
                        filename = f"تقرير_{row['اسم الطالب']}.html"
                        zf.writestr(filename, html.encode('utf-8'))
                
                st.download_button(
                    f"📦 تحميل جميع التقارير ({len(pivot)})",
                    zip_buffer.getvalue(),
                    f"reports_{datetime.now().strftime('%Y%m%d')}.zip",
                    "application/zip",
                    use_container_width=True
                )
                
                st.success(f"✅ تم إنشاء {len(pivot)} تقرير!")

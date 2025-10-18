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
    
    # Try to find level and section (numbers)
    level = ""
    section = ""
    subject_parts = []
    
    for part in parts:
        # Check if it's a number (level or section)
        if part.isdigit() or (part.startswith('0') and len(part) <= 2):
            if not level:
                level = part
            else:
                section = part
        else:
            subject_parts.append(part)
    
    subject = " ".join(subject_parts) if subject_parts else sheet_name
    
    # If no section found, try to get from Excel data
    return subject, level, section

def analyze_excel_file(file, sheet_name):
    """Analyze a single Excel sheet"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # Parse sheet name
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)
        
        # Try to get level and section from Excel (row 2, columns B and C)
        if len(df) > 1:
            level_from_excel = str(df.iloc[1, 1]).strip() if pd.notna(df.iloc[1, 1]) else ""
            section_from_excel = str(df.iloc[1, 2]).strip() if pd.notna(df.iloc[1, 2]) else ""
            
            # Use Excel data if available, otherwise use from sheet name
            level = level_from_excel if level_from_excel and level_from_excel != 'nan' else level_from_name
            section = section_from_excel if section_from_excel and section_from_excel != 'nan' else section_from_name
        else:
            level = level_from_name
            section = section_from_name
        
        # Get assessment titles from H1 onwards
        assessment_titles = []
        for col_idx in range(7, df.shape[1]):
            title = df.iloc[0, col_idx]
            if pd.notna(title) and str(title).strip():
                assessment_titles.append(str(title).strip())
        
        total_assessments = len(assessment_titles)
        
        results = []
        
        # Process each student starting from row 5
        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]
            
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue
            
            # Clean student name - remove extra spaces
            student_name_clean = " ".join(str(student_name).strip().split())
            
            # Count M (not submitted)
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
            
            # Calculate completed
            completed_count = total_assessments - m_count
            
            # Calculate percentage
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
    """Create pivot table with subjects as multiple columns"""
    # Get unique students
    students_base = df[['student_name', 'level', 'section']].drop_duplicates().reset_index(drop=True)
    
    # Start with base columns
    result = students_base.copy()
    
    # For each subject, add columns
    subjects = sorted(df['subject'].unique())
    
    for subject in subjects:
        subject_df = df[df['subject'] == subject].copy()
        
        # Create a unique key for merging
        subject_df['key'] = (
            subject_df['student_name'].astype(str) + '|' + 
            subject_df['level'].astype(str) + '|' + 
            subject_df['section'].astype(str)
        )
        result['key'] = (
            result['student_name'].astype(str) + '|' + 
            result['level'].astype(str) + '|' + 
            result['section'].astype(str)
        )
        
        # Prepare subject data with renamed columns
        subject_cols = subject_df[['key', 'total_count', 'completed_count', 'pending_titles', 'solve_pct']].copy()
        subject_cols.columns = [
            'key',
            f"{subject} - إجمالي التقييمات",
            f"{subject} - المنجز",
            f"{subject} - عناوين التقييمات المتبقية",
            f"{subject} - نسبة الإنجاز %"
        ]
        
        # Merge with result
        result = result.merge(subject_cols, on='key', how='left')
    
    # Remove the key column
    result = result.drop(columns=['key'])
    
    # Calculate overall average percentage
    pct_cols = [col for col in result.columns if 'نسبة الإنجاز %' in col]
    if pct_cols:
        result['نسبة حل التقييمات في جميع المواد'] = result[pct_cols].mean(axis=1)
    
    # Rename base columns to Arabic
    result = result.rename(columns={
        'student_name': 'اسم الطالب',
        'level': 'الصف',
        'section': 'الشعبة'
    })
    
    return result

# ================== MAIN APP ==================

st.title("📊 محلل التقييمات الأسبوعية")
st.markdown("---")

# Initialize session state
if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

# Sidebar
with st.sidebar:
    st.header("⚙️ الإعدادات")
    
    st.subheader("📁 تحميل الملفات")
    st.info("👇 اختر ملفات Excel للتحليل")
    
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="يمكنك رفع ملف واحد أو أكثر"
    )
    
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        
        try:
            xls = pd.ExcelFile(uploaded_files[0])
            sheets = xls.sheet_names
            
            st.subheader("📋 اختيار الأوراق")
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

# Main area
if not uploaded_files:
    st.info("👈 الرجاء رفع ملفات Excel من الشريط الجانبي")
    
    st.markdown("""
    ## 🎯 كيفية الاستخدام
    
    1. **ارفع ملفات Excel** من الشريط الجانبي
    2. **اختر المواد** المراد تحليلها
    3. **اضغط على "تشغيل التحليل"**
    """)

elif run_analysis:
    with st.spinner("جاري التحليل..."):
        try:
            all_results = []
            progress_bar = st.progress(0)
            total_sheets = len(uploaded_files) * len(selected_sheets)
            current = 0
            
            for file in uploaded_files:
                for sheet in selected_sheets:
                    st.text(f"📊 تحليل: {sheet}")
                    results = analyze_excel_file(file, sheet)
                    all_results.extend(results)
                    current += 1
                    progress_bar.progress(current / total_sheets)
            
            if all_results:
                df = pd.DataFrame(all_results)
                st.session_state.analysis_results = df
                
                pivot = create_pivot_table(df)
                st.session_state.pivot_table = pivot
                
                st.success(f"✅ تم تحليل {len(df)} سجل من {len(selected_sheets)} مادة!")
                
                # Debug info
                st.info(f"📊 تفاصيل التحليل: {len(pivot)} طالب فريد في الجدول النهائي")
                
                # Show duplicate detection
                unique_students = df.groupby('student_name').size()
                if unique_students.max() != len(selected_sheets):
                    st.warning(f"⚠️ تحذير: بعض الطلاب لديهم بيانات ناقصة في بعض المواد")
                    missing_data = unique_students[unique_students < len(selected_sheets)]
                    if len(missing_data) > 0:
                        st.text(f"الطلاب بالبيانات الناقصة ({len(missing_data)}):")
                        for name, count in missing_data.head(10).items():
                            st.text(f"  - {name}: موجود في {count} من {len(selected_sheets)} مادة")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")
        
        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")

# Display results
if st.session_state.pivot_table is not None:
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results
    
    # Summary
    st.markdown("## 📈 الإحصائيات العامة")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("👥 عدد الطلاب", len(pivot))
    with col2:
        st.metric("📚 عدد المواد", df['subject'].nunique())
    with col3:
        avg_overall = pivot['نسبة حل التقييمات في جميع المواد'].mean() if 'نسبة حل التقييمات في جميع المواد' in pivot.columns else 0
        st.metric("📈 متوسط النسبة", f"{avg_overall:.1f}%")
    with col4:
        st.metric("📋 إجمالي السجلات", len(df))
    
    st.divider()
    
    # Main table
    st.subheader("📊 البيانات التفصيلية")
    st.info("💡 كل مادة لها 4 أعمدة: الإجمالي - المنجز - عناوين المتبقية - نسبة الإنجاز")
    
    display_pivot = pivot.copy()
    
    for col in display_pivot.columns:
        if 'نسبة' in col:
            display_pivot[col] = display_pivot[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    
    display_pivot = display_pivot.fillna("-")
    
    st.dataframe(display_pivot, use_container_width=True, height=500)
    
    # Downloads
    st.markdown("### 📥 تنزيل النتائج")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot.to_excel(writer, index=False, sheet_name='النتائج')
        excel_data = output.getvalue()
        
        st.download_button(
            "📊 تحميل Excel",
            excel_data,
            f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        csv_data = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📄 تحميل CSV",
            csv_data,
            f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    st.success("✅ التحليل اكتمل بنجاح!")

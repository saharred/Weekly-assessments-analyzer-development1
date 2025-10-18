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
    
    if len(parts) >= 3:
        subject = " ".join(parts[:-2])
        level = parts[-2]
        section = parts[-1]
    elif len(parts) == 2:
        subject = parts[0]
        level = parts[1]
        section = ""
    else:
        subject = sheet_name
        level = ""
        section = ""
    
    return subject, level, section

def analyze_excel_file(file, sheet_name):
    """Analyze a single Excel sheet"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # Parse sheet name
        subject, level, section = parse_sheet_name(sheet_name)
        
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
                "student_name": str(student_name).strip(),
                "subject": subject,
                "level": level,
                "section": section,
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
    # Create a unique identifier for each student
    df['student_id'] = df['student_name'] + '_' + df['level'].astype(str) + '_' + df['section'].astype(str)
    
    # Get unique students
    students = df[['student_name', 'level', 'section', 'student_id']].drop_duplicates()
    
    # Create the base dataframe
    result = students[['student_name', 'level', 'section']].copy()
    
    # For each subject, add three columns
    for subject in sorted(df['subject'].unique()):
        subject_data = df[df['subject'] == subject].set_index('student_id')
        
        # Merge total assessments
        total_col = f"{subject} - إجمالي التقييمات"
        result = result.merge(
            subject_data[['total_count']].rename(columns={'total_count': total_col}),
            left_on=result['student_name'] + '_' + result['level'].astype(str) + '_' + result['section'].astype(str),
            right_index=True,
            how='left'
        )
        
        # Merge completed assessments
        completed_col = f"{subject} - المنجز"
        result = result.merge(
            subject_data[['completed_count']].rename(columns={'completed_count': completed_col}),
            left_on=result['student_name'] + '_' + result['level'].astype(str) + '_' + result['section'].astype(str),
            right_index=True,
            how='left'
        )
        
        # Merge pending titles
        pending_col = f"{subject} - عناوين التقييمات المتبقية"
        result = result.merge(
            subject_data[['pending_titles']].rename(columns={'pending_titles': pending_col}),
            left_on=result['student_name'] + '_' + result['level'].astype(str) + '_' + result['section'].astype(str),
            right_index=True,
            how='left'
        )
        
        # Merge percentage
        pct_col = f"{subject} - نسبة الإنجاز %"
        result = result.merge(
            subject_data[['solve_pct']].rename(columns={'solve_pct': pct_col}),
            left_on=result['student_name'] + '_' + result['level'].astype(str) + '_' + result['section'].astype(str),
            right_index=True,
            how='left'
        )
    
    # Remove the merge key column
    result = result.drop(columns=['key_0'], errors='ignore')
    
    # Calculate overall average
    pct_cols = [col for col in result.columns if col.endswith('نسبة الإنجاز %')]
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
    
    # File upload
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
        
        # Get sheets from first file
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
    
    # Quantitative and Qualitative Reports
    st.markdown("## 📑 التقارير الكمية والوصفية")
    
    tab1, tab2, tab3 = st.tabs(["📚 تقارير المواد", "🏫 تقارير الشعب", "📊 تقرير شامل"])
    
    with tab1:
        st.subheader("📚 التقارير الكمية والوصفية للمواد")
        
        subjects = df['subject'].unique()
        
        selected_subject = st.selectbox(
            "اختر المادة لعرض التقرير:",
            subjects,
            key="subject_report_select"
        )
        
        if selected_subject:
            subject_df = df[df['subject'] == selected_subject]
            
            # Generate subject report
            report = f"""
# 📊 التقرير الكمي والوصفي - {selected_subject}
{'='*80}

## 📈 الإحصائيات الكمية

### 1. معلومات عامة:
- **عدد الطلاب الكلي**: {len(subject_df)} طالب/طالبة
- **المستويات**: {', '.join(subject_df['level'].unique())}
- **الشعب**: {', '.join(subject_df['section'].unique())}
- **إجمالي التقييمات**: {subject_df['total_count'].iloc[0] if len(subject_df) > 0 else 0}

### 2. مؤشرات الأداء:
- **متوسط نسبة الإنجاز**: {subject_df['solve_pct'].mean():.2f}%
- **أعلى نسبة إنجاز**: {subject_df['solve_pct'].max():.2f}%
- **أقل نسبة إنجاز**: {subject_df['solve_pct'].min():.2f}%
- **الانحراف المعياري**: {subject_df['solve_pct'].std():.2f}%

### 3. توزيع الطلاب حسب الأداء:
- **ممتاز (90-100%)**: {len(subject_df[subject_df['solve_pct'] >= 90])} طالب ({len(subject_df[subject_df['solve_pct'] >= 90])/len(subject_df)*100:.1f}%)
- **جيد جداً (80-89%)**: {len(subject_df[(subject_df['solve_pct'] >= 80) & (subject_df['solve_pct'] < 90)])} طالب ({len(subject_df[(subject_df['solve_pct'] >= 80) & (subject_df['solve_pct'] < 90)])/len(subject_df)*100:.1f}%)
- **جيد (70-79%)**: {len(subject_df[(subject_df['solve_pct'] >= 70) & (subject_df['solve_pct'] < 80)])} طالب ({len(subject_df[(subject_df['solve_pct'] >= 70) & (subject_df['solve_pct'] < 80)])/len(subject_df)*100:.1f}%)
- **مقبول (60-69%)**: {len(subject_df[(subject_df['solve_pct'] >= 60) & (subject_df['solve_pct'] < 70)])} طالب ({len(subject_df[(subject_df['solve_pct'] >= 60) & (subject_df['solve_pct'] < 70)])/len(subject_df)*100:.1f}%)
- **ضعيف (أقل من 60%)**: {len(subject_df[subject_df['solve_pct'] < 60])} طالب ({len(subject_df[subject_df['solve_pct'] < 60])/len(subject_df)*100:.1f}%)

{'='*80}

## 📝 التحليل الوصفي

### 1. التقييم العام للمادة:
"""
            # Add qualitative assessment
            avg_pct = subject_df['solve_pct'].mean()
            if avg_pct >= 90:
                assessment = "**ممتاز** 🌟\nأداء المادة متميز جداً. معظم الطلاب يحققون نتائج ممتازة."
            elif avg_pct >= 80:
                assessment = "**جيد جداً** 👍\nأداء المادة جيد بشكل عام. هناك مجال للتحسين لبعض الطلاب."
            elif avg_pct >= 70:
                assessment = "**جيد** ✓\nأداء المادة مقبول. يحتاج بعض الطلاب إلى دعم إضافي."
            elif avg_pct >= 60:
                assessment = "**مقبول** ⚠️\nأداء المادة متوسط. يُنصح بتكثيف المتابعة والدعم."
            else:
                assessment = "**يحتاج تحسين** ❌\nأداء المادة ضعيف. يتطلب تدخلاً عاجلاً ودعماً مكثفاً."
            
            report += assessment + "\n\n"
            
            report += f"""
### 2. نقاط القوة:
"""
            excellent_students = subject_df[subject_df['solve_pct'] >= 90]
            if len(excellent_students) > 0:
                report += f"- {len(excellent_students)} طالب حققوا نسبة 90% فأكثر\n"
                report += f"- أفضل أداء: {excellent_students['student_name'].iloc[0]} ({excellent_students['solve_pct'].iloc[0]:.1f}%)\n"
            
            report += f"""
### 3. نقاط تحتاج إلى تحسين:
"""
            weak_students = subject_df[subject_df['solve_pct'] < 60]
            if len(weak_students) > 0:
                report += f"- {len(weak_students)} طالب يحتاجون إلى دعم عاجل (أقل من 60%)\n"
                report += f"- متوسط إنجازهم: {weak_students['solve_pct'].mean():.1f}%\n"
            
            report += f"""
### 4. التوصيات:
"""
            if avg_pct >= 80:
                report += "- الحفاظ على المستوى الحالي\n"
                report += "- تحفيز الطلاب المتميزين بأنشطة إثرائية\n"
                report += f"- متابعة الطلاب الـ {len(subject_df[subject_df['solve_pct'] < 80])} الذين أقل من 80%\n"
            else:
                report += "- تكثيف المتابعة للطلاب ضعيفي الأداء\n"
                report += "- تنظيم حصص دعم إضافية\n"
                report += "- مراجعة طرق التدريس والتقييم\n"
                report += "- تفعيل التواصل مع أولياء الأمور\n"
            
            report += f"""
{'='*80}

## 👥 قائمة الطلاب المتميزين (أعلى 5):
"""
            top_5 = subject_df.nlargest(5, 'solve_pct')[['student_name', 'solve_pct', 'completed_count', 'total_count']]
            for idx, row in top_5.iterrows():
                report += f"- {row['student_name']}: {row['solve_pct']:.1f}% ({row['completed_count']}/{row['total_count']})\n"
            
            report += f"""
## ⚠️ قائمة الطلاب الذين يحتاجون دعم (أقل 5):
"""
            bottom_5 = subject_df.nsmallest(5, 'solve_pct')[['student_name', 'solve_pct', 'completed_count', 'total_count']]
            for idx, row in bottom_5.iterrows():
                report += f"- {row['student_name']}: {row['solve_pct']:.1f}% ({row['completed_count']}/{row['total_count']})\n"
            
            report += f"""
{'='*80}
تاريخ إنشاء التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
            
            # Display report
            st.text_area(
                "التقرير الكامل:",
                report,
                height=600,
                key=f"subject_report_{selected_subject}"
            )
            
            # Download button
            st.download_button(
                f"📥 تحميل تقرير {selected_subject}",
                report.encode('utf-8'),
                f"تقرير_{selected_subject}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                "text/plain",
                use_container_width=True
            )
    
    with tab2:
        st.subheader("🏫 التقارير الكمية والوصفية للشعب")
        
        # Get unique level-section combinations
        df['level_section'] = df['level'].astype(str) + ' - ' + df['section'].astype(str)
        sections = df['level_section'].unique()
        
        selected_section = st.selectbox(
            "اختر المستوى والشعبة:",
            sections,
            key="section_report_select"
        )
        
        if selected_section:
            level, section = selected_section.split(' - ')
            section_df = df[(df['level'] == level) & (df['section'] == section)]
            
            # Generate section report
            section_report = f"""
# 📊 التقرير الكمي والوصفي - المستوى {level} - الشعبة {section}
{'='*80}

## 📈 الإحصائيات الكمية

### 1. معلومات عامة:
- **عدد الطلاب**: {len(section_df['student_name'].unique())} طالب/طالبة
- **عدد المواد**: {len(section_df['subject'].unique())} مادة
- **المواد**: {', '.join(section_df['subject'].unique())}

### 2. مؤشرات الأداء العامة:
- **متوسط نسبة الإنجاز**: {section_df['solve_pct'].mean():.2f}%
- **أعلى نسبة**: {section_df['solve_pct'].max():.2f}%
- **أقل نسبة**: {section_df['solve_pct'].min():.2f}%

### 3. أداء المواد في هذه الشعبة:
"""
            for subject in section_df['subject'].unique():
                subj_data = section_df[section_df['subject'] == subject]
                section_report += f"\n**{subject}**:\n"
                section_report += f"- عدد الطلاب: {len(subj_data)}\n"
                section_report += f"- متوسط الإنجاز: {subj_data['solve_pct'].mean():.1f}%\n"
                section_report += f"- طلاب ممتازون (90%+): {len(subj_data[subj_data['solve_pct'] >= 90])}\n"
                section_report += f"- طلاب يحتاجون دعم (<60%): {len(subj_data[subj_data['solve_pct'] < 60])}\n"
            
            section_report += f"""
{'='*80}

## 📝 التحليل الوصفي

### 1. التقييم العام للشعبة:
"""
            avg_pct = section_df['solve_pct'].mean()
            if avg_pct >= 85:
                assessment = "**شعبة متميزة** 🌟\nالأداء العام للشعبة ممتاز."
            elif avg_pct >= 75:
                assessment = "**شعبة جيدة** 👍\nالأداء العام جيد مع وجود مجال للتحسين."
            elif avg_pct >= 65:
                assessment = "**شعبة مقبولة** ⚠️\nتحتاج الشعبة إلى المزيد من الدعم."
            else:
                assessment = "**شعبة تحتاج تدخل** ❌\nتحتاج الشعبة إلى خطة تحسين شاملة."
            
            section_report += assessment + "\n\n"
            
            section_report += f"""
### 2. التوصيات:
- تنظيم جلسات دعم جماعية للمواد الضعيفة
- تكريم الطلاب المتميزين
- التواصل مع أولياء أمور الطلاب ضعيفي الأداء
- تفعيل التعلم التعاوني بين الطلاب

{'='*80}
تاريخ إنشاء التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
            
            st.text_area(
                "التقرير الكامل:",
                section_report,
                height=600,
                key=f"section_report_{selected_section}"
            )
            
            st.download_button(
                f"📥 تحميل تقرير الشعبة {level}-{section}",
                section_report.encode('utf-8'),
                f"تقرير_الشعبة_{level}_{section}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                "text/plain",
                use_container_width=True
            )
    
    with tab3:
        st.subheader("📊 التقرير الشامل")
        
        if st.button("🔄 إنشاء تقرير شامل لجميع المواد والشعب", use_container_width=True):
            with st.spinner("جاري إنشاء التقرير الشامل..."):
                
                comprehensive_report = f"""
# 📊 التقرير الشامل - جميع المواد والشعب
{'='*80}
تاريخ التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 📈 ملخص تنفيذي

### الإحصائيات العامة:
- **إجمالي الطلاب**: {len(df['student_name'].unique())}
- **عدد المواد**: {len(df['subject'].unique())}
- **عدد الشعب**: {len(df.groupby(['level', 'section']))}
- **متوسط الأداء العام**: {df['solve_pct'].mean():.2f}%

{'='*80}

## 📚 تقارير المواد:
"""
                
                # Generate reports for all subjects
                for subject in df['subject'].unique():
                    subject_df = df[df['subject'] == subject]
                    comprehensive_report += f"""
### {subject}
- عدد الطلاب: {len(subject_df)}
- متوسط الإنجاز: {subject_df['solve_pct'].mean():.2f}%
- ممتازون (90%+): {len(subject_df[subject_df['solve_pct'] >= 90])}
- يحتاجون دعم (<60%): {len(subject_df[subject_df['solve_pct'] < 60])}

"""
                
                comprehensive_report += f"""
{'='*80}

## 🏫 تقارير الشعب:
"""
                
                # Generate reports for all sections
                for (level, section), group in df.groupby(['level', 'section']):
                    comprehensive_report += f"""
### المستوى {level} - الشعبة {section}
- عدد الطلاب: {len(group['student_name'].unique())}
- متوسط الإنجاز: {group['solve_pct'].mean():.2f}%
- عدد المواد: {len(group['subject'].unique())}

"""
                
                comprehensive_report += f"""
{'='*80}
نهاية التقرير
"""
                
                st.text_area(
                    "التقرير الشامل:",
                    comprehensive_report,
                    height=600
                )
                
                st.download_button(
                    "📥 تحميل التقرير الشامل",
                    comprehensive_report.encode('utf-8'),
                    f"تقرير_شامل_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    "text/plain",
                    use_container_width=True
                )
    
    st.divider()
    
    # Run button
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
    
    ### 📋 متطلبات الملف
    - الصف 1: عناوين التقييمات (من H1 يميناً)
    - الصف 5 فما بعد: أسماء الطلاب والنتائج
    - اسم الورقة: المادة + المستوى + الشعبة
    - M = لم يتم التسليم
    
    ### 📊 المخرجات
    سيتم عرض جدول بحيث:
    - كل **طالب** في صف
    - كل **مادة** في عمود
    - نسبة إنجاز الطالب في كل مادة
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
                
                # Create pivot table
                pivot = create_pivot_table(df)
                st.session_state.pivot_table = pivot
                
                st.success(f"✅ تم تحليل {len(df)} سجل من {len(selected_sheets)} مادة!")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")
        
        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

# Display results
if st.session_state.pivot_table is not None:
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results
    
    # Summary metrics
    st.markdown("## 📈 الإحصائيات العامة")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("👥 عدد الطلاب", len(pivot))
    with col2:
        st.metric("📚 عدد المواد", df['subject'].nunique())
    with col3:
        avg_overall = pivot['نسبة حل التقييمات في جميع المواد'].mean() if 'نسبة حل التقييمات في جميع المواد' in pivot.columns else 0
        st.metric("📈 متوسط النسبة الإجمالية", f"{avg_overall:.1f}%")
    with col4:
        total_records = len(df)
        st.metric("📋 إجمالي السجلات", total_records)
    
    st.divider()
    
    # Main pivot table
    st.subheader("📊 البيانات التفصيلية")
    st.info("💡 كل مادة لها 4 أعمدة: الإجمالي - المنجز - عناوين المتبقية - نسبة الإنجاز")
    
    # Format the display
    display_pivot = pivot.copy()
    
    # Format percentage columns
    for col in display_pivot.columns:
        if 'نسبة' in col and col != 'نسبة حل التقييمات في جميع المواد':
            display_pivot[col] = display_pivot[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    
    # Format overall percentage
    if 'نسبة حل التقييمات في جميع المواد' in display_pivot.columns:
        display_pivot['نسبة حل التقييمات في جميع المواد'] = display_pivot['نسبة حل التقييمات في جميع المواد'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    
    # Fill NA values
    display_pivot = display_pivot.fillna("-")
    
    st.dataframe(display_pivot, use_container_width=True, height=500)
    
    # Download buttons
    st.markdown("### 📥 تنزيل النتائج")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Download pivot table as Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot.to_excel(writer, index=False, sheet_name='الجدول المحوري')
        excel_pivot = output.getvalue()
        
        st.download_button(
            "📊 تحميل الجدول المحوري (Excel)",
            excel_pivot,
            f"pivot_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        # Download raw data as Excel
        output_raw = io.BytesIO()
        with pd.ExcelWriter(output_raw, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='البيانات الخام')
        excel_raw = output_raw.getvalue()
        
        st.download_button(
            "📋 تحميل البيانات الخام (Excel)",
            excel_raw,
            f"raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col3:
        # Download CSV
        csv_pivot = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📄 تحميل CSV",
            csv_pivot,
            f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    st.divider()
    
    # Statistics by subject
    st.subheader("📊 إحصائيات حسب المادة")
    
    subject_stats = []
    subjects = df['subject'].unique()
    
    for subject in subjects:
        subject_df = df[df['subject'] == subject]
        
        stats = {
            'المادة': subject,
            'عدد الطلاب': len(subject_df),
            'متوسط النسبة': f"{subject_df['solve_pct'].mean():.1f}%",
            'أعلى نسبة': f"{subject_df['solve_pct'].max():.1f}%",
            'أقل نسبة': f"{subject_df['solve_pct'].min():.1f}%",
            'طلاب 100%': len(subject_df[subject_df['solve_pct'] == 100]),
            'طلاب أقل من 50%': len(subject_df[subject_df['solve_pct'] < 50])
        }
        subject_stats.append(stats)
    
    stats_df = pd.DataFrame(subject_stats)
    st.dataframe(stats_df, use_container_width=True, hide_index=True)
    
    st.divider()
    
    # Charts
    st.subheader("📈 الرسوم البيانية")
    
    # Configure matplotlib for better rendering
    import matplotlib
    matplotlib.rcParams['axes.unicode_minus'] = False
    plt.rcParams['font.size'] = 11
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📊 متوسط النسب حسب المادة**")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        subject_avg = df.groupby('subject')['solve_pct'].mean().sort_values(ascending=True)
        
        # Create horizontal bar chart
        y_pos = range(len(subject_avg))
        colors = plt.cm.viridis(range(len(subject_avg)))
        bars = ax.barh(y_pos, subject_avg.values, color=colors, edgecolor='black', linewidth=1.5)
        
        # Add value labels
        for i, (bar, value) in enumerate(zip(bars, subject_avg.values)):
            ax.text(value + 2, i, f'{value:.1f}%', va='center', fontsize=11, fontweight='bold')
        
        # Set labels with subject names (in Arabic)
        ax.set_yticks(y_pos)
        ax.set_yticklabels([f"  {subj}  " for subj in subject_avg.index], fontsize=10)
        
        ax.set_xlabel("Average Completion Rate (%)", fontsize=12, fontweight='bold')
        ax.set_title("Performance by Subject", fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='x', alpha=0.3, linestyle='--')
        ax.set_xlim(0, 110)
        
        # Add colored background for reference zones
        ax.axvspan(0, 60, alpha=0.1, color='red')
        ax.axvspan(60, 80, alpha=0.1, color='yellow')
        ax.axvspan(80, 100, alpha=0.1, color='green')
        
        plt.tight_layout()
        st.pyplot(fig)
        st.caption("🟢 80-100% ممتاز | 🟡 60-80% جيد | 🔴 0-60% يحتاج تحسين")
    
    with col2:
        st.markdown("**📈 توزيع النسب الإجمالية**")
        
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            fig, ax = plt.subplots(figsize=(10, 6))
            
            overall_scores = pivot['نسبة حل التقييمات في جميع المواد'].dropna()
            
            n, bins, patches = ax.hist(overall_scores, bins=20, edgecolor='black', linewidth=1.5)
            
            # Color gradient based on score ranges
            for i, patch in enumerate(patches):
                bin_center = (bins[i] + bins[i+1]) / 2
                if bin_center >= 80:
                    patch.set_facecolor('#4CAF50')  # Green
                elif bin_center >= 60:
                    patch.set_facecolor('#FFC107')  # Yellow
                else:
                    patch.set_facecolor('#F44336')  # Red
            
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
            st.caption("🟢 ممتاز | 🟡 جيد | 🔴 يحتاج تحسين")
    
    # Additional charts
    st.markdown("### 📊 تحليلات إضافية")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📋 توزيع الطلاب حسب الفئات**")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            overall_scores = pivot['نسبة حل التقييمات في جميع المواد'].dropna()
            
            # Categorize
            categories = pd.cut(overall_scores, 
                               bins=[0, 50, 70, 80, 90, 100], 
                               labels=['0-50%\nWeak', '50-70%\nAcceptable', '70-80%\nGood', '80-90%\nVery Good', '90-100%\nExcellent'])
            
            category_counts = categories.value_counts().sort_index()
            
            colors_cat = ['#F44336', '#FF9800', '#FFC107', '#8BC34A', '#4CAF50']
            bars = ax.bar(range(len(category_counts)), category_counts.values, 
                         color=colors_cat, edgecolor='black', linewidth=1.5)
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                       f'{int(height)}',
                       ha='center', va='bottom', fontsize=13, fontweight='bold')
            
            ax.set_xticks(range(len(category_counts)))
            ax.set_xticklabels(category_counts.index, fontsize=10, rotation=0)
            ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
            ax.set_title("Student Distribution by Performance Level", fontsize=14, fontweight='bold', pad=20)
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            ax.set_ylim(0, max(category_counts.values) * 1.15)
            
            plt.tight_layout()
            st.pyplot(fig)
    
    with col2:
        st.markdown("**👥 عدد الطلاب لكل مادة**")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        subject_counts = df.groupby('subject').size().sort_values(ascending=False)
        
        colors_subjects = plt.cm.Set3(range(len(subject_counts)))
        bars = ax.bar(range(len(subject_counts)), subject_counts.values, 
                     color=colors_subjects, edgecolor='black', linewidth=1.5)
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                   f'{int(height)}',
                   ha='center', va='bottom', fontsize=13, fontweight='bold')
        
        # Use numbers instead of subject names on x-axis, show legend
        ax.set_xticks(range(len(subject_counts)))
        ax.set_xticklabels([f"#{i+1}" for i in range(len(subject_counts))], fontsize=11)
        ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
        ax.set_title("Students per Subject", fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.set_ylim(0, max(subject_counts.values) * 1.15)
        
        plt.tight_layout()
        st.pyplot(fig)
        
        # Show subject names as list below
        st.caption("**Subject Names:**")
        for i, subj in enumerate(subject_counts.index):
            st.caption(f"#{i+1}: {subj} ({subject_counts.values[i]} students)")
    
    # Comparison chart
    st.markdown("### 📊 مقارنة الأداء بين المواد")
    
    fig, ax = plt.subplots(figsize=(14, 6))
    
    # Get data for each category per subject
    subjects = df['subject'].unique()
    categories = ['90-100%', '80-89%', '70-79%', '60-69%', '<60%']
    
    data_matrix = []
    for subject in subjects:
        subj_data = df[df['subject'] == subject]
        counts = [
            len(subj_data[subj_data['solve_pct'] >= 90]),
            len(subj_data[(subj_data['solve_pct'] >= 80) & (subj_data['solve_pct'] < 90)]),
            len(subj_data[(subj_data['solve_pct'] >= 70) & (subj_data['solve_pct'] < 80)]),
            len(subj_data[(subj_data['solve_pct'] >= 60) & (subj_data['solve_pct'] < 70)]),
            len(subj_data[subj_data['solve_pct'] < 60])
        ]
        data_matrix.append(counts)
    
    x = range(len(subjects))
    width = 0.15
    colors_bars = ['#4CAF50', '#8BC34A', '#FFC107', '#FF9800', '#F44336']
    
    for i, (category, color) in enumerate(zip(categories, colors_bars)):
        values = [data_matrix[j][i] for j in range(len(subjects))]
        ax.bar([p + width * i for p in x], values, width, label=category, 
               color=color, edgecolor='black', linewidth=1)
    
    ax.set_xlabel("Subject #", fontsize=12, fontweight='bold')
    ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
    ax.set_title("Performance Distribution by Subject", fontsize=14, fontweight='bold', pad=20)
    ax.set_xticks([p + width * 2 for p in x])
    ax.set_xticklabels([f"#{i+1}" for i in range(len(subjects))], fontsize=11)
    ax.legend(title="Performance", fontsize=10, title_fontsize=11)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    st.pyplot(fig)
    
    # Show subject mapping
    st.caption("**Subject Mapping:**")
    cols = st.columns(min(3, len(subjects)))
    for i, subj in enumerate(subjects):
        with cols[i % len(cols)]:
            st.caption(f"#{i+1}: {subj}")
    
    # Heatmap
    st.markdown("### 🔥 خريطة حرارية: أداء أفضل 20 طالب")
    
    # Get subject columns from pivot
    subject_cols_pct = [col for col in pivot.columns if 'نسبة الإنجاز %' in col and 'جميع المواد' not in col]
    
    if len(subject_cols_pct) > 0 and len(pivot) > 0:
        # Get top 20 students by overall average
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            top_20 = pivot.nlargest(20, 'نسبة حل التقييمات في جميع المواد')
        else:
            top_20 = pivot.head(20)
        
        fig, ax = plt.subplots(figsize=(14, max(10, len(top_20) * 0.4)))
        
        # Prepare data for heatmap
        heatmap_data = top_20[subject_cols_pct].values
        
        im = ax.imshow(heatmap_data, cmap='RdYlGn', aspect='auto', vmin=0, vmax=100)
        
        # Set ticks
        ax.set_xticks(range(len(subject_cols_pct)))
        # Use subject number instead of name
        subject_labels = [f"S{i+1}" for i in range(len(subject_cols_pct))]
        ax.set_xticklabels(subject_labels, rotation=0, ha='center', fontsize=10)
        
        ax.set_yticks(range(len(top_20)))
        student_labels = [f"{row['اسم الطالب'][:25]}" for _, row in top_20.iterrows()]
        ax.set_yticklabels(student_labels, fontsize=9)
        
        # Add colorbar
        cbar = plt.colorbar(im, ax=ax)
        cbar.set_label('Completion Rate (%)', fontsize=11, fontweight='bold')
        
        # Add values in cells
        for i in range(len(top_20)):
            for j in range(len(subject_cols_pct)):
                value = heatmap_data[i, j]
                if pd.notna(value) and value > 0:
                    text_color = 'white' if value < 50 else 'black'
                    ax.text(j, i, f'{value:.0f}', ha='center', va='center', 
                           color=text_color, fontsize=8, fontweight='bold')
        
        ax.set_title("Performance Heatmap - Top 20 Students", fontsize=14, fontweight='bold', pad=20)
        plt.tight_layout()
        st.pyplot(fig)
        
        # Show subject mapping for heatmap
        st.caption("**Subject Key:**")
        cols_key = st.columns(min(4, len(subject_cols_pct)))
        for i, col in enumerate(subject_cols_pct):
            subject_name = col.replace(' - نسبة الإنجاز %', '')
            with cols_key[i % len(cols_key)]:
                st.caption(f"S{i+1}: {subject_name}")
        
        st.info("💡 الألوان: 🟢 أخضر = ممتاز (80%+) | 🟡 أصفر = متوسط (50-80%) | 🔴 أحمر = ضعيف (<50%)")
    else:
        st.warning("⚠️ لا توجد بيانات كافية لعرض الخريطة الحرارية")
    
    st.divider()
    
    # Top and bottom performers
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏆 أفضل 10 طلاب")
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            top_10 = pivot.nlargest(10, 'نسبة حل التقييمات في جميع المواد')[['اسم الطالب', 'الصف', 'الشعبة', 'نسبة حل التقييمات في جميع المواد']]
            top_10['نسبة حل التقييمات في جميع المواد'] = top_10['نسبة حل التقييمات في جميع المواد'].apply(lambda x: f"{x:.1f}%")
            st.dataframe(top_10, hide_index=True, use_container_width=True)
    
    with col2:
        st.subheader("⚠️ أقل 10 طلاب (تحتاج دعم)")
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            bottom_10 = pivot.nsmallest(10, 'نسبة حل التقييمات في جميع المواد')[['اسم الطالب', 'الصف', 'الشعبة', 'نسبة حل التقييمات في جميع المواد']]
            bottom_10['نسبة حل التقييمات في جميع المواد'] = bottom_10['نسبة حل التقييمات في جميع المواد'].apply(lambda x: f"{x:.1f}%")
            st.dataframe(bottom_10, hide_index=True, use_container_width=True)

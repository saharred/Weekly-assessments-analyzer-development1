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
    """Create pivot table with subjects as columns"""
    # Create a unique identifier for each student
    df['student_id'] = df['student_name'] + '_' + df['level'] + '_' + df['section']
    
    # Create pivot table
    pivot = df.pivot_table(
        index=['student_name', 'level', 'section'],
        columns='subject',
        values='solve_pct',
        aggfunc='first'
    ).reset_index()
    
    # Rename columns
    pivot.columns.name = None
    
    # Add overall statistics
    subject_cols = [col for col in pivot.columns if col not in ['student_name', 'level', 'section']]
    if subject_cols:
        pivot['نسبة حل التقييمات في جميع المواد'] = pivot[subject_cols].mean(axis=1)
    
    # Reorder columns
    final_cols = ['student_name', 'level', 'section'] + subject_cols
    if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
        final_cols.append('نسبة حل التقييمات في جميع المواد')
    
    pivot = pivot[final_cols]
    
    # Rename to Arabic
    pivot = pivot.rename(columns={
        'student_name': 'اسم الطالب',
        'level': 'الصف',
        'section': 'الشعبة'
    })
    
    return pivot

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
    st.info("💡 كل مادة في عمود منفصل - النسب المئوية تمثل نسبة إنجاز الطالب في كل مادة")
    
    # Format percentages
    display_pivot = pivot.copy()
    
    # Format all numeric columns as percentages
    for col in display_pivot.columns:
        if col not in ['اسم الطالب', 'الصف', 'الشعبة']:
            if display_pivot[col].dtype in ['float64', 'int64']:
                display_pivot[col] = display_pivot[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    
    st.dataframe(display_pivot, use_container_width=True, height=500)
    
    # Download buttons
    col1, col2 = st.columns(2)
    
    with col1:
        # Download pivot table
        csv_pivot = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 تحميل الجدول المحوري (CSV)",
            csv_pivot,
            f"pivot_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    with col2:
        # Download raw data
        csv_raw = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 تحميل البيانات الخام (CSV)",
            csv_raw,
            f"raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
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
    
    # Configure matplotlib for Arabic
    plt.rcParams['font.family'] = 'DejaVu Sans'
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.text("متوسط النسب حسب المادة")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        subject_avg = df.groupby('subject')['solve_pct'].mean().sort_values(ascending=True)
        
        colors = plt.cm.viridis(range(len(subject_avg)))
        bars = ax.barh(range(len(subject_avg)), subject_avg.values, color=colors, edgecolor='black')
        
        # Add value labels
        for i, (bar, value) in enumerate(zip(bars, subject_avg.values)):
            ax.text(value + 1, i, f'{value:.1f}%', va='center', fontsize=10, fontweight='bold')
        
        ax.set_yticks(range(len(subject_avg)))
        ax.set_yticklabels(subject_avg.index, fontsize=11)
        ax.set_xlabel("Average Completion Rate (%)", fontsize=12, fontweight='bold')
        ax.set_title("Performance by Subject", fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='x', alpha=0.3, linestyle='--')
        ax.set_xlim(0, 110)
        
        plt.tight_layout()
        st.pyplot(fig)
    
    with col2:
        st.text("توزيع النسب الإجمالية")
        
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            fig, ax = plt.subplots(figsize=(10, 6))
            
            overall_scores = pivot['نسبة حل التقييمات في جميع المواد'].dropna()
            
            n, bins, patches = ax.hist(overall_scores, bins=20, color='#667eea', edgecolor='black', alpha=0.7)
            
            # Color gradient
            cm = plt.cm.RdYlGn
            for i, patch in enumerate(patches):
                patch.set_facecolor(cm(bins[i]/100))
            
            mean_val = overall_scores.mean()
            ax.axvline(mean_val, color='red', linestyle='--', linewidth=2.5, label=f'Mean: {mean_val:.1f}%')
            
            ax.set_xlabel("Completion Rate (%)", fontsize=12, fontweight='bold')
            ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
            ax.set_title("Overall Performance Distribution", fontsize=14, fontweight='bold', pad=20)
            ax.legend(fontsize=11)
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            
            plt.tight_layout()
            st.pyplot(fig)
    
    # Additional charts
    st.markdown("### 📊 تحليلات إضافية")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.text("توزيع الطلاب حسب الفئات")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            overall_scores = pivot['نسبة حل التقييمات في جميع المواد'].dropna()
            
            # Categorize
            categories = pd.cut(overall_scores, 
                               bins=[0, 50, 70, 80, 90, 100], 
                               labels=['< 50%', '50-70%', '70-80%', '80-90%', '90-100%'])
            
            category_counts = categories.value_counts().sort_index()
            
            colors_cat = ['#ff6b6b', '#ff9a56', '#ffd89b', '#a8edea', '#f093fb']
            bars = ax.bar(range(len(category_counts)), category_counts.values, color=colors_cat, edgecolor='black', linewidth=1.5)
            
            # Add value labels
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height,
                       f'{int(height)}',
                       ha='center', va='bottom', fontsize=12, fontweight='bold')
            
            ax.set_xticks(range(len(category_counts)))
            ax.set_xticklabels(category_counts.index, fontsize=11, rotation=0)
            ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
            ax.set_title("Student Distribution by Category", fontsize=14, fontweight='bold', pad=20)
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            
            plt.tight_layout()
            st.pyplot(fig)
    
    with col2:
        st.text("مقارنة المواد (عدد الطلاب)")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        subject_counts = df.groupby('subject').size().sort_values(ascending=False)
        
        colors_subjects = plt.cm.Set3(range(len(subject_counts)))
        bars = ax.bar(range(len(subject_counts)), subject_counts.values, color=colors_subjects, edgecolor='black', linewidth=1.5)
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom', fontsize=12, fontweight='bold')
        
        ax.set_xticks(range(len(subject_counts)))
        ax.set_xticklabels(subject_counts.index, fontsize=11, rotation=45, ha='right')
        ax.set_ylabel("Number of Students", fontsize=12, fontweight='bold')
        ax.set_title("Students per Subject", fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        
        plt.tight_layout()
        st.pyplot(fig)
    
    # Heatmap
    st.markdown("### 🔥 خريطة حرارية: أداء الطلاب حسب المادة")
    
    # Create heatmap data
    heatmap_data = pivot.set_index(['اسم الطالب', 'الصف', 'الشعبة'])
    subject_cols = [col for col in heatmap_data.columns if col != 'نسبة حل التقييمات في جميع المواد']
    
    if len(subject_cols) > 0 and len(heatmap_data) > 0:
        fig, ax = plt.subplots(figsize=(12, max(8, len(heatmap_data) * 0.3)))
        
        # Prepare data for heatmap
        heatmap_values = heatmap_data[subject_cols].head(20)  # Show top 20 students
        
        im = ax.imshow(heatmap_values, cmap='RdYlGn', aspect='auto', vmin=0, vmax=100)
        
        # Set ticks
        ax.set_xticks(range(len(subject_cols)))
        ax.set_xticklabels(subject_cols, rotation=45, ha='right', fontsize=10)
        
        ax.set_yticks(range(len(heatmap_values)))
        student_labels = [f"{idx[0][:20]}" for idx in heatmap_values.index]  # Truncate long names
        ax.set_yticklabels(student_labels, fontsize=9)
        
        # Add colorbar
        cbar = plt.colorbar(im, ax=ax)
        cbar.set_label('Completion Rate (%)', fontsize=11, fontweight='bold')
        
        # Add values in cells
        for i in range(len(heatmap_values)):
            for j in range(len(subject_cols)):
                value = heatmap_values.iloc[i, j]
                if pd.notna(value):
                    text_color = 'white' if value < 50 else 'black'
                    ax.text(j, i, f'{value:.0f}%', ha='center', va='center', 
                           color=text_color, fontsize=8, fontweight='bold')
        
        ax.set_title("Performance Heatmap (Top 20 Students)", fontsize=14, fontweight='bold', pad=20)
        plt.tight_layout()
        st.pyplot(fig)
        
        st.info("💡 الألوان: 🟢 أخضر = ممتاز | 🟡 أصفر = متوسط | 🔴 أحمر = يحتاج تحسين")
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

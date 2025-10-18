import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
from datetime import datetime, timedelta
from pathlib import Path
from src.analyzer import AssessmentAnalyzer, generate_html_report
from src.email_reports import SubjectReportGenerator, EmailSender

# Page config
st.set_page_config(
    page_title="Weekly Assessments Analyzer v3.7",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Title
st.title("📊 Weekly Assessments Analyzer v3.7")

# Initialize session state
if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = None

# Sidebar
with st.sidebar:
    st.header("⚙️ الإعدادات")
    
    # File upload
    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    
    if uploaded_files:
        st.session_state.uploaded_files = uploaded_files
        
        # Get sheets from first file for preview
        try:
            file_path = uploaded_files[0]
            xls = pd.ExcelFile(file_path)
            sheets = xls.sheet_names
            
            selected_sheets = st.multiselect(
                "اختر الأوراق",
                sheets,
                default=sheets if len(sheets) <= 3 else sheets[:3]
            )
        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {e}")
            selected_sheets = []
    else:
        selected_sheets = []
    
    # Analysis parameters
    st.subheader("🔧 معاملات التحليل")
    
    col1, col2 = st.columns(2)
    
    with col1:
        start_col = st.text_input(
            "عمود التقييمات",
            value="H",
            max_chars=2,
            help="عمود بداية التقييمات (H افتراضي)"
        ).upper()
        
        names_row = st.number_input(
            "صف أسماء الطلاب",
            value=5,
            min_value=1,
            help="صف أسماء الطلاب (5 افتراضي)"
        )
    
    with col2:
        names_col = st.text_input(
            "عمود أسماء الطلاب",
            value="A",
            max_chars=2,
            help="عمود أسماء الطلاب (A افتراضي)"
        ).upper()
        
        due_row = st.number_input(
            "صف تواريخ الاستحقاق",
            value=3,
            min_value=1,
            help="صف تواريخ الاستحقاق (3 افتراضي)"
        )
    
    # Date filter
    st.subheader("📅 تصفية التواريخ")
    enable_filter = st.checkbox("تفعيل تصفية النطاق الزمني", value=False)
    
    if enable_filter:
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input(
                "من",
                value=datetime(2025, 1, 1),
                help="تاريخ البداية"
            )
        with col2:
            end_date = st.date_input(
                "إلى",
                value=datetime(2025, 12, 31),
                help="تاريخ النهاية"
            )
        date_range = (start_date, end_date)
    else:
        date_range = None
    
    # Action button
    st.divider()
    run_analysis = st.button(
        "🚀 تشغيل التحليل الآن",
        use_container_width=True,
        type="primary"
    )

# Main content
if run_analysis and uploaded_files and selected_sheets:
    with st.spinner("جاري التحليل..."):
        try:
            analyzer = AssessmentAnalyzer(
                start_col_letter=start_col,
                names_row=names_row,
                names_col=names_col,
                due_row=due_row,
                date_range=date_range
            )
            
            results = []
            for uploaded_file in uploaded_files:
                file_results = analyzer.analyze_file(
                    uploaded_file,
                    selected_sheets
                )
                results.extend(file_results)
            
            if results:
                st.session_state.analysis_results = pd.DataFrame(results)
                st.success("✅ تم إكمال التحليل بنجاح!")
            else:
                st.warning("لم يتم العثور على بيانات للتحليل.")
        
        except Exception as e:
            st.error(f"❌ خطأ في التحليل: {str(e)}")

# Display results
if st.session_state.analysis_results is not None:
    df = st.session_state.analysis_results
    
    # Summary statistics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("👥 عدد الطلاب", len(df))
    with col2:
        avg_solve = df["solve_pct"].mean()
        st.metric("📈 متوسط النسبة", f"{avg_solve:.1f}%")
    with col3:
        platinum = len(df[df["category"] == "البلاتينية"])
        st.metric("🏆 البلاتينية", platinum)
    with col4:
        needs_improvement = len(df[df["category"] == "تحتاج إلى تحسين"])
        st.metric("⚠️ يحتاج تحسين", needs_improvement)
    
    st.divider()
    
    # Summary table
    st.subheader("📊 جدول الملخص")
    
    # Prepare display columns with Arabic headers
    display_df = df[[
        "student_name", "subject", "class", "section", "total_material_solved", 
        "total_assessments", "unsolved_assessment_count", "unsolved_titles", 
        "solve_pct", "category", "recommendation"
    ]].copy()
    
    # Rename columns to Arabic
    arabic_headers = {
        "student_name": "اسم الطالب",
        "subject": "المادة",
        "class": "المستوى",
        "section": "الشعبة",
        "total_material_solved": "تقييمات منجزة",
        "total_assessments": "تقييمات متبقية",
        "unsolved_assessment_count": "عدد غير منجزة",
        "unsolved_titles": "عناوين غير منجزة",
        "solve_pct": "نسبة الإنجاز %",
        "category": "الفئة",
        "recommendation": "التوصية"
    }
    
    display_df = display_df.rename(columns=arabic_headers)
    
    # Format solve_pct for display
    display_df["نسبة الإنجاز %"] = display_df["نسبة الإنجاز %"].apply(lambda x: f"{x:.2f}%")
    
    # Display with RTL support
    st.dataframe(display_df, use_container_width=True, height=400, hide_index=True)
    
    # Download CSV
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        label="📥 تحميل ملخص CSV",
        data=csv,
        file_name=f"assessment_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        use_container_width=True
    )
    
    st.divider()
    
    # Charts
    st.subheader("📈 الرسوم البيانية")
    
    col1, col2 = st.columns(2)
    
    # Category distribution
    with col1:
        st.text("توزيع الفئات")
        fig, ax = plt.subplots(figsize=(10, 6))
        category_counts = df["category"].value_counts()
        
        # Color mapping for Arabic categories
        color_map = {
            "البلاتينية": "#f093fb",
            "الذهبي": "#ffd89b",
            "الفضي": "#a8edea",
            "البرونزي": "#ff9a56",
            "تحتاج إلى تحسين": "#ff6b6b"
        }
        colors = [color_map.get(cat, "#999999") for cat in category_counts.index]
        
        category_counts.plot(kind="bar", ax=ax, color=colors, edgecolor="black", linewidth=1.5)
        ax.set_xlabel("الفئة", fontsize=12, fontweight="bold")
        ax.set_ylabel("عدد الطلاب", fontsize=12, fontweight="bold")
        ax.set_title("توزيع الفئات", fontsize=14, fontweight="bold")
        ax.tick_params(axis="x", rotation=45)
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig, use_container_width=True)
    
    # Solve percentage histogram
    with col2:
        st.text("توزيع نسبة الإنجاز")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.hist(df["solve_pct"], bins=15, color="#667eea", edgecolor="black", linewidth=1.5)
        ax.set_xlabel("نسبة الإنجاز (%)", fontsize=12, fontweight="bold")
        ax.set_ylabel("عدد الطلاب", fontsize=12, fontweight="bold")
        ax.set_title("توزيع نسبة الإنجاز", fontsize=14, fontweight="bold")
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig, use_container_width=True)
    
    # Top 10 remaining assessments
    st.text("أكثر 10 طلاب لديهم تقييمات متبقية")
    fig, ax = plt.subplots(figsize=(12, 6))
    top_remaining = df.nlargest(10, "total_assessments")[["student_name", "total_assessments"]].copy()
    top_remaining = top_remaining.sort_values("total_assessments")
    
    ax.barh(range(len(top_remaining)), top_remaining["total_assessments"].values, color="#FF9800", edgecolor="black", linewidth=1.5)
    ax.set_yticks(range(len(top_remaining)))
    ax.set_yticklabels(top_remaining["student_name"].values, fontsize=10)
    ax.set_xlabel("التقييمات المتبقية", fontsize=12, fontweight="bold")
    ax.set_title("أكثر الطلاب لديهم تقييمات متبقية", fontsize=14, fontweight="bold")
    ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)
    
    # Additional statistics
    st.subheader("📊 إحصائيات إضافية")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        subjects = df["subject"].nunique()
        st.metric("📚 عدد المواد", subjects)
    
    with col2:
        avg_assessments = df["total_material_solved"].mean()
        st.metric("✅ متوسط التقييمات المنجزة", f"{avg_assessments:.1f}")
    
    with col3:
        zero_solved = len(df[df["total_material_solved"] == 0])
        st.metric("⚠️ طلاب لم ينجزوا شيئاً", zero_solved)
    
    st.divider()
    
    # Advanced filters
    st.subheader("🔍 تصفية متقدمة")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        selected_subjects = st.multiselect(
            "اختر المواد",
            df["subject"].unique(),
            default=df["subject"].unique()
        )
    
    with col2:
        selected_categories = st.multiselect(
            "اختر الفئات",
            ["البلاتينية", "الذهبي", "الفضي", "البرونزي", "تحتاج إلى تحسين"],
            default=["البلاتينية", "الذهبي", "الفضي", "البرونزي", "تحتاج إلى تحسين"]
        )
    
    with col3:
        min_solve_pct = st.slider(
            "الحد الأدنى لنسبة الإنجاز",
            min_value=0.0,
            max_value=100.0,
            value=0.0,
            step=5.0
        )
    
    # Apply filters
    filtered_df = df[
        (df["subject"].isin(selected_subjects)) &
        (df["category"].isin(selected_categories)) &
        (df["solve_pct"] >= min_solve_pct)
    ]
    
    st.info(f"📋 تم العثور على {len(filtered_df)} من أصل {len(df)} طالب")
    
    st.divider()
    
    # Subject Analysis & Reports
    st.subheader("📑 التقارير الوصفية حسب المادة")
    
    # Group by subject
    subjects = df['subject'].unique()
    
    col1, col2 = st.columns(2)
    
    with col1:
        selected_subject_report = st.selectbox(
            "اختر المادة للتقرير الوصفي",
            subjects,
            key="subject_report"
        )
    
    with col2:
        report_type = st.radio(
            "نوع التقرير",
            ["عرض على الشاشة", "تحميل نصي", "إرسال بريد إلكتروني"],
            horizontal=True
        )
    
    if st.button("📊 إنشاء التقرير الوصفي", use_container_width=True):
        # Filter data for selected subject
        subject_data = df[df['subject'] == selected_subject_report]
        
        # Group by level and section
        grouped = subject_data.groupby(['class', 'section'])
        
        report_generator = SubjectReportGenerator()
        
        for (level, section), group_data in grouped:
            students_list = group_data.to_dict('records')
            
            # Generate report
            report = report_generator.generate_subject_report(
                selected_subject_report,
                str(level),
                str(section),
                students_list
            )
            
            # Identify inactive students
            inactive = group_data[group_data['solve_pct'] < 70].to_dict('records')
            critical = group_data[group_data['solve_pct'] < 50].to_dict('records')
            
            if report_type == "عرض على الشاشة":
                st.text_area(
                    f"تقرير {selected_subject_report} - {level}/{section}",
                    value=report,
                    height=400,
                    disabled=True
                )
            
            elif report_type == "تحميل نصي":
                st.download_button(
                    label=f"📥 تحميل تقرير {selected_subject_report}_{level}_{section}",
                    data=report.encode('utf-8'),
                    file_name=f"report_{selected_subject_report}_{level}_{section}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime="text/plain; charset=utf-8",
                    use_container_width=True
                )
            
            elif report_type == "إرسال بريد إلكتروني":
                st.warning("⚙️ إعدادات البريد الإلكتروني")
                
                email_col1, email_col2 = st.columns(2)
                
                with email_col1:
                    teacher_email = st.text_input(
                        "بريد المعلم الإلكتروني",
                        placeholder="teacher@example.com"
                    )
                    smtp_server = st.text_input(
                        "خادم SMTP",
                        value="smtp.gmail.com"
                    )
                
                with email_col2:
                    sender_email = st.text_input(
                        "البريد المرسل",
                        placeholder="your-email@gmail.com"
                    )
                    sender_password = st.text_input(
                        "كلمة المرور (أو App Password)",
                        type="password"
                    )
                
                if st.button("✉️ إرسال البريد", use_container_width=True):
                    if not (teacher_email and sender_email and sender_password):
                        st.error("❌ يرجى ملء جميع بيانات البريد")
                    else:
                        try:
                            email_sender = EmailSender(
                                smtp_server=smtp_server,
                                smtp_port=587,
                                sender_email=sender_email,
                                sender_password=sender_password
                            )
                            
                            success, message = email_sender.send_subject_report(
                                teacher_email=teacher_email,
                                subject=selected_subject_report,
                                level=str(level),
                                section=str(section),
                                report_content=report,
                                inactive_students=inactive,
                                critical_students=critical
                            )
                            
                            if success:
                                st.success(f"✅ {message}")
                                st.info(f"تم إرسال التقرير إلى {teacher_email}")
                            else:
                                st.error(f"❌ {message}")
                        
                        except Exception as e:
                            st.error(f"❌ خطأ: {str(e)}")
                            st.info("💡 تأكد من:\n• استخدام Gmail App Password (وليس كلمة المرور العادية)\n• تفعيل المصادقة الثنائية\n• السماح للتطبيقات الأقل أماناً")
    
    st.divider()
    
    # Summary by Subject
    st.subheader("📊 ملخص حسب المادة")
    
    subject_summary = []
    
    for subject in subjects:
        subject_df = df[df['subject'] == subject]
        
        summary_item = {
            "المادة": subject,
            "عدد الطلاب": len(subject_df),
            "متوسط النسبة %": f"{subject_df['solve_pct'].mean():.2f}",
            "الفئة الأولى": len(subject_df[subject_df['solve_pct'] >= 90]),
            "الفئة الثانية": len(subject_df[subject_df['solve_pct'] >= 70]),
            "غير فاعلين": len(subject_df[subject_df['solve_pct'] < 70]),
            "في الخطر": len(subject_df[subject_df['solve_pct'] < 50])
        }
        subject_summary.append(summary_item)
    
    summary_df = pd.DataFrame(subject_summary)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
    # Export subject summary
    st.download_button(
        label="📥 تحميل ملخص المواد (CSV)",
        data=summary_df.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"subject_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        use_container_width=True
    )
    
    st.divider()
    
    # HTML Reports Generation
    st.subheader("📄 تقارير الطلاب")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        report_type = st.radio(
            "نوع التقرير",
            ["جميع الطلاب", "الطلاب المصفاة", "فئة محددة"],
            horizontal=True
        )
    
    with col2:
        if report_type == "فئة محددة":
            selected_category = st.selectbox(
                "اختر الفئة",
                ["البلاتينية", "الذهبي", "الفضي", "البرونزي", "تحتاج إلى تحسين"]
            )
        else:
            selected_category = None
    
    # Determine which data to use for reports
    if report_type == "جميع الطلاب":
        report_data = df
    elif report_type == "الطلاب المصفاة":
        report_data = filtered_df
    else:  # فئة محددة
        report_data = df[df["category"] == selected_category]
    
    if st.button("🔄 إنشاء تقارير HTML للطلاب (قابلة للطباعة PDF)", use_container_width=True):
        if len(report_data) == 0:
            st.warning("⚠️ لا توجد بيانات للتقرير")
        else:
            with st.spinner(f"جاري إنشاء {len(report_data)} تقرير..."):
                try:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for _, row in report_data.iterrows():
                            html_content = generate_html_report(row)
                            filename = f"{row['subject']}_{row['student_name']}.html"
                            zf.writestr(filename, html_content.encode("utf-8"))
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        label=f"📦 تحميل {len(report_data)} تقرير",
                        data=zip_buffer.getvalue(),
                        file_name=f"student_reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    st.success(f"✅ تم إنشاء {len(report_data)} تقرير بنجاح!")
                
                except Exception as e:
                    st.error(f"❌ خطأ في إنشاء التقارير: {str(e)}")

else:
    if not uploaded_files:
        st.info("👈 الرجاء تحميل ملفات Excel من الشريط الجانبي")
    elif not selected_sheets:
        st.info("👈 الرجاء اختيار أوراق العمل")
    else:
        st.info("👈 انقر على 'تشغيل التحليل الآن' للبدء")
```

---

## 📄 **requirements.txt**
```
pandas>=2.2.2
openpyxl>=3.1.5
xlrd==2.0.1
matplotlib>=3.8.0
streamlit>=1.38.0
python-dotenv>=1.0.0

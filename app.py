import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import date, timedelta

# --- Configuration and Setup ---
st.set_page_config(
    page_title="أي إنجاز - محلل تقييمات الطلاب",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Arabic text for Streamlit components
ARABIC_TEXT = {
    "title": "أي إنجاز - محلل تقييمات الطلاب",
    "subtitle": "تحليل بيانات التقييمات الأسبوعية لطلاب قطر",
    "upload_file": "قم بتحميل ملف التقييمات الأسبوعية (Excel)",
    "file_format_note": "يرجى التأكد من أن الملف بصيغة Excel ويحتوي على أوراق عمل لكل شعبة.",
    "data_summary": "ملخص البيانات المحملة",
    "grade": "الصف",
    "section": "الشعبة",
    "student_count": "عدد الطلاب",
    "avg_achievement": "متوسط نسبة الإنجاز",
    "overall_achievement": "نسبة الإنجاز الكلية",
    "date_filter_title": "فلترة حسب تاريخ الاستحقاق",
    "start_date": "من تاريخ",
    "end_date": "إلى تاريخ",
    "select_subject": "اختر المادة لعرض التفاصيل",
    "student_categorization": "تصنيف الطلاب حسب نسبة الإنجاز",
    "category": "الفئة",
    "platinum": "بلاتيني (95% فأكثر)",
    "gold": "ذهبي (85% - 94.9%)",
    "silver": "فضي (75% - 84.9%)",
    "bronze": "برونزي (65% - 74.9%)",
    "needs_improvement": "يحتاج إلى تحسين (أقل من 65%)",
    "visualization_title": "توزيع نسبة الإنجاز",
    "recommendations_title": "توصيات للمعلمين",
    "recommendation_1": "ركز على الطلاب في فئة 'يحتاج إلى تحسين' من خلال خطط دعم فردية.",
    "recommendation_2": "استخدم بيانات الطلاب البلاتينيين والذهبيين كأمثلة للنجاح.",
    "recommendation_3": "تحقق من التقييمات التي لم يتم حلها لتحديد ما إذا كانت بسبب صعوبة المادة أو عدم الاستحقاق.",
    "email_alert_title": "تنبيهات البريد الإلكتروني للطلاب غير النشطين",
    "inactive_students_note": "الطلاب الذين لم يحلوا أي تقييمات مستحقة في الفترة المحددة.",
    "generate_email": "إنشاء نص تنبيه البريد الإلكتروني",
    "email_subject": "تنبيه: نشاط الطالب في التقييمات الأسبوعية",
    "email_body_template": "تحية طيبة،\n\nنود أن نلفت انتباهكم إلى أن الطالب/ة **{student_name}** من شعبة **{section}** لم يقم بحل أي من التقييمات المستحقة في الفترة المحددة.\n\nيرجى التواصل مع الطالب/ة وولي أمره/ا لتقديم الدعم اللازم.\n\nمع خالص التقدير،\nإدارة المدرسة",
    "top_sections_title": "المراكز الثلاثة الأولى في الإنجاز",
    "subject": "المادة",
    "rank": "الترتيب",
    "achievement_rate": "نسبة الإنجاز",
    "section_achievement_report": "تقرير إنجاز المادة والشعبة",
    "export_excel": "تصدير التقرير إلى Excel",
    "no_data_message": "يرجى تحميل ملف Excel للبدء بالتحليل.",
    "no_assessments_in_range": "لا توجد تقييمات مستحقة في نطاق التاريخ المحدد.",
    "overall_column": "Overall",
    "due_date_row": 1 # 0-indexed row for due dates (row 2 in Excel)
}

# --- Data Processing Functions ---

@st.cache_data
def process_excel_file(uploaded_file):
    """
    Reads the Excel file, processes each sheet, and returns a combined DataFrame
    and a summary DataFrame.
    """
    xls = pd.ExcelFile(uploaded_file)
    all_data = []
    summary_data = []

    for sheet_name in xls.sheet_names:
        # Extract Grade and Section from sheet name (e.g., "الصف ثالث1")
        # Assuming the format is "الصف [Grade][Section]"
        parts = sheet_name.split()
        if len(parts) >= 2:
            grade_section = parts[-1]
            grade = grade_section[:-1] # e.g., "ثالث"
            section = grade_section[-1] # e.g., "1"
        else:
            grade = "غير محدد"
            section = sheet_name

        try:
            # Read the sheet, skipping the first row (header) to get to the due dates
            df = xls.parse(sheet_name, header=None)

            # Due dates are in the second row (index 1)
            due_dates = df.iloc[ARABIC_TEXT["due_date_row"]].copy()
            
            # The actual data starts from the third row (index 2)
            df.columns = df.iloc[2]
            df = df[3:].reset_index(drop=True)
            
            # Clean column names (remove NaN and convert to string)
            df.columns = [str(col) for col in df.columns]

            # Find the 'Overall' column (index 5, 0-indexed)
            overall_col_name = df.columns[5]
            df[overall_col_name] = pd.to_numeric(df[overall_col_name], errors='coerce')

            # Add Grade and Section columns
            df[ARABIC_TEXT["grade"]] = grade
            df[ARABIC_TEXT["section"]] = section
            df["Sheet_Name"] = sheet_name
            
            # Prepare assessment columns and due dates
            assessment_cols = df.columns[6:]
            assessment_due_dates = due_dates.iloc[6:]
            
            # Store data for later use
            all_data.append(df)

            # Calculate summary
            student_count = len(df)
            avg_achievement = df[overall_col_name].mean() if student_count > 0 else 0
            
            summary_data.append({
                ARABIC_TEXT["grade"]: grade,
                ARABIC_TEXT["section"]: section,
                ARABIC_TEXT["student_count"]: student_count,
                ARABIC_TEXT["avg_achievement"]: f"{avg_achievement:.2f}%",
                ARABIC_TEXT["overall_achievement"]: df[overall_col_name].sum() / (student_count * 100) if student_count > 0 else 0,
                "Raw_Avg_Achievement": avg_achievement,
                "Assessment_Cols": assessment_cols.tolist(),
                "Assessment_Due_Dates": assessment_due_dates.to_dict()
            })

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة ورقة العمل '{sheet_name}': {e}")
            continue

    if not all_data:
        return None, None, None

    combined_df = pd.concat(all_data, ignore_index=True)
    summary_df = pd.DataFrame(summary_data)
    
    # Extract all unique due dates and assessment names
    all_due_dates = {}
    for item in summary_data:
        all_due_dates.update(item["Assessment_Due_Dates"])
    
    # Convert due dates to datetime objects
    for key, value in all_due_dates.items():
        try:
            # Attempt to parse as date, handle NaT/None
            if pd.notna(value):
                all_due_dates[key] = pd.to_datetime(value).date()
            else:
                all_due_dates[key] = None
        except:
            all_due_dates[key] = None # Fallback for unparseable dates

    return combined_df, summary_df, all_due_dates

def filter_data_by_date(df, all_due_dates, start_date, end_date):
    """
    Filters the combined DataFrame to only include assessments with due dates
    within the specified range.
    """
    if df is None:
        return None, None

    # Identify assessment columns that fall within the date range
    valid_assessment_cols = []
    for col_name, due_date in all_due_dates.items():
        if due_date and start_date <= due_date <= end_date:
            valid_assessment_cols.append(col_name)

    if not valid_assessment_cols:
        return None, None

    # The 'Overall' column is used for the main achievement rate and is not filtered by date
    # The filtering logic here is mainly for the detailed analysis and the 'Top 3' calculation
    
    # For the purpose of the 'Top 3' and 'Section Achievement Report', we need to calculate
    # a new achievement rate based *only* on the assessments in the date range.
    
    # Identify subject columns (assuming the first part of the assessment name is the subject)
    subject_cols = {}
    for col in valid_assessment_cols:
        # Assuming assessment name is like "Subject - Assessment Name"
        subject = col.split(' - ')[0].strip() if ' - ' in col else "مادة غير محددة"
        if subject not in subject_cols:
            subject_cols[subject] = []
        subject_cols[subject].append(col)

    # Calculate a new 'Filtered_Achievement' column
    # This is a simplified calculation: average of the percentages in the valid columns
    # Note: The user explicitly requested to keep the system's 'Overall' column for the main summary.
    # This new calculation is only for the new features (Top 3, Section Report).
    
    df_filtered = df.copy()
    
    # Convert valid assessment columns to numeric, coercing errors to NaN
    for col in valid_assessment_cols:
        df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce')
        
    # Calculate the average percentage across the valid assessments for each student
    df_filtered['Filtered_Achievement'] = df_filtered[valid_assessment_cols].mean(axis=1)
    
    # Calculate the achievement rate per subject and section
    section_achievement_report = []
    for subject, cols in subject_cols.items():
        # Calculate the average percentage across the subject's valid assessments for each student
        df_filtered[f'Subject_Achievement_{subject}'] = df_filtered[cols].mean(axis=1)
        
        # Group by section and calculate the average achievement for the subject
        subject_group = df_filtered.groupby([ARABIC_TEXT["grade"], ARABIC_TEXT["section"]])[f'Subject_Achievement_{subject}'].mean().reset_index()
        subject_group.rename(columns={f'Subject_Achievement_{subject}': ARABIC_TEXT["achievement_rate"]}, inplace=True)
        subject_group[ARABIC_TEXT["subject"]] = subject
        
        section_achievement_report.append(subject_group)

    if section_achievement_report:
        section_achievement_df = pd.concat(section_achievement_report, ignore_index=True)
        section_achievement_df[ARABIC_TEXT["achievement_rate"]] = section_achievement_df[ARABIC_TEXT["achievement_rate"]].fillna(0).round(2)
    else:
        section_achievement_df = pd.DataFrame()

    return df_filtered, section_achievement_df

def categorize_students(df):
    """Categorizes students based on their 'Overall' achievement percentage."""
    if df is None or ARABIC_TEXT["overall_column"] not in df.columns:
        return pd.DataFrame()

    def get_category(percentage):
        if percentage >= 95:
            return ARABIC_TEXT["platinum"]
        elif percentage >= 85:
            return ARABIC_TEXT["gold"]
        elif percentage >= 75:
            return ARABIC_TEXT["silver"]
        elif percentage >= 65:
            return ARABIC_TEXT["bronze"]
        else:
            return ARABIC_TEXT["needs_improvement"]

    df[ARABIC_TEXT["category"]] = df[ARABIC_TEXT["overall_column"]].apply(get_category)
    
    category_order = [
        ARABIC_TEXT["platinum"], ARABIC_TEXT["gold"], ARABIC_TEXT["silver"],
        ARABIC_TEXT["bronze"], ARABIC_TEXT["needs_improvement"]
    ]
    
    category_counts = df.groupby(ARABIC_TEXT["category"]).size().reset_index(name='Count')
    category_counts[ARABIC_TEXT["category"]] = pd.Categorical(
        category_counts[ARABIC_TEXT["category"]], categories=category_order, ordered=True
    )
    category_counts = category_counts.sort_values(ARABIC_TEXT["category"])
    
    return category_counts

def get_top_sections(section_achievement_df, top_n=3):
    """Calculates and returns the top N sections per subject."""
    if section_achievement_df.empty:
        return pd.DataFrame()

    top_sections = section_achievement_df.sort_values(
        by=[ARABIC_TEXT["subject"], ARABIC_TEXT["achievement_rate"]],
        ascending=[True, False]
    ).groupby(ARABIC_TEXT["subject"]).head(top_n).reset_index(drop=True)
    
    top_sections[ARABIC_TEXT["rank"]] = top_sections.groupby(ARABIC_TEXT["subject"]).cumcount() + 1
    
    # Format achievement rate as percentage string
    top_sections[ARABIC_TEXT["achievement_rate"]] = top_sections[ARABIC_TEXT["achievement_rate"]].apply(lambda x: f"{x:.2f}%")
    
    return top_sections[[ARABIC_TEXT["subject"], ARABIC_TEXT["rank"], ARABIC_TEXT["grade"], ARABIC_TEXT["section"], ARABIC_TEXT["achievement_rate"]]]

def to_excel(df):
    """Converts a DataFrame to an Excel file in memory."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    processed_data = output.getvalue()
    return processed_data

# --- Streamlit App Layout ---

st.title(ARABIC_TEXT["title"])
st.markdown(f"### {ARABIC_TEXT['subtitle']}")

uploaded_file = st.file_uploader(ARABIC_TEXT["upload_file"], type=["xlsx", "xls"])

if uploaded_file:
    combined_df, summary_df, all_due_dates = process_excel_file(uploaded_file)

    if combined_df is not None:
        
        # --- Sidebar for Date Filtering ---
        st.sidebar.header(ARABIC_TEXT["date_filter_title"])
        
        # Determine the min and max dates from the assessment due dates
        valid_dates = [d for d in all_due_dates.values() if d is not None]
        
        if valid_dates:
            min_date = min(valid_dates)
            max_date = max(valid_dates)
            
            # Set default range to the last 7 days or the full range if less than 7 days
            default_start_date = max(min_date, max_date - timedelta(days=6))
            
            start_date = st.sidebar.date_input(
                ARABIC_TEXT["start_date"],
                value=default_start_date,
                min_value=min_date,
                max_value=max_date
            )
            end_date = st.sidebar.date_input(
                ARABIC_TEXT["end_date"],
                value=max_date,
                min_value=min_date,
                max_value=max_date
            )
            
            if start_date > end_date:
                st.sidebar.error("تاريخ البداية يجب أن يكون قبل تاريخ النهاية.")
                st.stop()
                
            # Filter data based on the selected date range
            filtered_df, section_achievement_df = filter_data_by_date(combined_df, all_due_dates, start_date, end_date)
            
            if filtered_df is None:
                st.warning(ARABIC_TEXT["no_assessments_in_range"])
                st.stop()
                
        else:
            st.sidebar.info("لا توجد تواريخ استحقاق صالحة في الملف للفلترة.")
            filtered_df = combined_df
            section_achievement_df = pd.DataFrame() # Empty if no dates to filter by

        # --- Main Content ---
        
        # 1. Data Summary Table
        st.header(ARABIC_TEXT["data_summary"])
        # Display Grade and Section correctly
        summary_display_df = summary_df[[
            ARABIC_TEXT["grade"], ARABIC_TEXT["section"], ARABIC_TEXT["student_count"], ARABIC_TEXT["avg_achievement"]
        ]]
        st.dataframe(summary_display_df, hide_index=True, use_container_width=True)

        # 2. Student Categorization
        st.header(ARABIC_TEXT["student_categorization"])
        category_counts = categorize_students(combined_df)
        if not category_counts.empty:
            col1, col2 = st.columns([1, 2])
            with col1:
                st.dataframe(category_counts.rename(columns={'Count': 'عدد الطلاب'}), hide_index=True, use_container_width=True)
            with col2:
                fig_pie = px.pie(
                    category_counts,
                    values='Count',
                    names=ARABIC_TEXT["category"],
                    title=ARABIC_TEXT["visualization_title"],
                    color=ARABIC_TEXT["category"],
                    color_discrete_map={
                        ARABIC_TEXT["platinum"]: 'green',
                        ARABIC_TEXT["gold"]: 'gold',
                        ARABIC_TEXT["silver"]: 'silver',
                        ARABIC_TEXT["bronze"]: 'brown',
                        ARABIC_TEXT["needs_improvement"]: 'red'
                    }
                )
                st.plotly_chart(fig_pie, use_container_width=True)

        # 3. Top 3 Sections Ranking (New Feature)
        if not section_achievement_df.empty:
            st.header(ARABIC_TEXT["top_sections_title"])
            top_sections_df = get_top_sections(section_achievement_df)
            
            # Group by subject and display the top 3 for each
            for subject, group in top_sections_df.groupby(ARABIC_TEXT["subject"]):
                st.markdown(f"#### {ARABIC_TEXT['subject']}: {subject}")
                st.dataframe(
                    group[[ARABIC_TEXT["rank"], ARABIC_TEXT["grade"], ARABIC_TEXT["section"], ARABIC_TEXT["achievement_rate"]]],
                    hide_index=True,
                    use_container_width=True
                )

        # 4. Section Achievement Report (New Feature)
        if not section_achievement_df.empty:
            st.header(ARABIC_TEXT["section_achievement_report"])
            
            # Pivot the table for better display: Subject as index, Section as columns
            pivot_df = section_achievement_df.pivot_table(
                index=ARABIC_TEXT["subject"],
                columns=[ARABIC_TEXT["grade"], ARABIC_TEXT["section"]],
                values=ARABIC_TEXT["achievement_rate"]
            ).fillna(0).applymap(lambda x: f"{x:.2f}%")
            
            st.dataframe(pivot_df, use_container_width=True)
            
            # Excel Export Button
            excel_data = to_excel(section_achievement_df)
            st.download_button(
                label=ARABIC_TEXT["export_excel"],
                data=excel_data,
                file_name="section_achievement_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 5. Recommendations
        st.header(ARABIC_TEXT["recommendations_title"])
        st.markdown(f"- {ARABIC_TEXT['recommendation_1']}")
        st.markdown(f"- {ARABIC_TEXT['recommendation_2']}")
        st.markdown(f"- {ARABIC_TEXT['recommendation_3']}")

        # 6. Email Alert Generation
        st.header(ARABIC_TEXT["email_alert_title"])
        st.info(ARABIC_TEXT["inactive_students_note"])
        
        # Identify inactive students (simplistic: Overall < 1%)
        inactive_students = combined_df[combined_df[ARABIC_TEXT["overall_column"]] < 1]
        
        if not inactive_students.empty:
            st.dataframe(
                inactive_students[['Student Name', ARABIC_TEXT["grade"], ARABIC_TEXT["section"], ARABIC_TEXT["overall_column"]]].rename(
                    columns={'Student Name': 'اسم الطالب', ARABIC_TEXT["overall_column"]: ARABIC_TEXT["overall_achievement"]}
                ),
                hide_index=True,
                use_container_width=True
            )
            
            if st.button(ARABIC_TEXT["generate_email"]):
                email_list = []
                for index, row in inactive_students.iterrows():
                    email_body = ARABIC_TEXT["email_body_template"].format(
                        student_name=row['Student Name'],
                        section=f"{row[ARABIC_TEXT['grade']]}{row[ARABIC_TEXT['section']]}"
                    )
                    email_list.append(f"**{ARABIC_TEXT['email_subject']}**\n\n{email_body}")
                
                st.subheader("نصوص البريد الإلكتروني الجاهزة")
                for email in email_list:
                    st.code(email, language='markdown')
        else:
            st.success("لا يوجد طلاب غير نشطين (نسبة إنجازهم أقل من 1%) في البيانات المحملة.")

    else:
        st.error("لم يتم العثور على بيانات صالحة في الملف المحمل.")

else:
    st.info(ARABIC_TEXT["no_data_message"])
    st.markdown(f"**ملاحظة حول تنسيق الملف:** {ARABIC_TEXT['file_format_note']}")

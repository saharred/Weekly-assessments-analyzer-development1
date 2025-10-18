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
    # Format: "المادة المستوى الشعبة" or "المادة 01 6"
    parts = sheet_name.strip().split()
    
    if len(parts) >= 3:
        subject = " ".join(parts[:-2])  # كل شيء قبل آخر عنصرين
        level = parts[-2]  # قبل الأخير
        section = parts[-1]  # الأخير
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
    """Analyze a single Excel sheet with new structure"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # Parse sheet name
        subject, level, section = parse_sheet_name(sheet_name)
        
        # Get assessment titles from H1 onwards (row 0, starting from column 7)
        assessment_titles = []
        for col_idx in range(7, df.shape[1]):  # Starting from H (index 7)
            title = df.iloc[0, col_idx]
            if pd.notna(title) and str(title).strip():
                assessment_titles.append(str(title).strip())
        
        total_assessments = len(assessment_titles)
        
        # Get due dates from row 3 (index 2)
        due_dates = []
        for col_idx in range(7, 7 + total_assessments):
            due_date = df.iloc[2, col_idx]
            if pd.notna(due_date):
                due_dates.append(str(due_date))
            else:
                due_dates.append("")
        
        results = []
        
        # Process each student starting from row 5 (index 4)
        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]  # Column A
            
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue
            
            # Get overall percentage from column F (index 5)
            overall_pct_cell = df.iloc[idx, 5]
            
            # Parse percentage
            if pd.notna(overall_pct_cell):
                overall_str = str(overall_pct_cell).replace('%', '').strip()
                try:
                    overall_pct = float(overall_str)
                except:
                    overall_pct = 0.0
            else:
                overall_pct = 0.0
            
            # Count M (not submitted) assessments from H onwards (starting col 7)
            m_count = 0  # Count of "M" (لم يتم التسليم)
            pending_titles = []
            
            for i, col_idx in enumerate(range(7, 7 + total_assessments)):
                if col_idx < df.shape[1]:
                    cell_value = df.iloc[idx, col_idx]
                    
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip().upper()
                        
                        # Count only M as not submitted
                        if cell_str == 'M':
                            m_count += 1
                            if i < len(assessment_titles):
                                pending_titles.append(assessment_titles[i])
            
            # Calculate completed assessments
            completed_count = total_assessments - m_count
            pending_count = m_count
            
            # Calculate solve percentage
            if total_assessments > 0:
                solve_pct = (completed_count / total_assessments) * 100
            else:
                solve_pct = 0.0
            
            # Categorize student
            if solve_pct >= 90:
                category = "البلاتينية"
                recommendation = "أداء ممتاز! استمر في التميز 🌟"
            elif solve_pct >= 80:
                category = "الذهبي"
                recommendation = "أداء جيد جداً، حافظ على مستواك 🥇"
            elif solve_pct >= 70:
                category = "الفضي"
                recommendation = "أداء جيد، يمكنك التحسن أكثر 🥈"
            elif solve_pct >= 60:
                category = "البرونزي"
                recommendation = "أداء مقبول، تحتاج لمزيد من الجهد 🥉"
            else:
                category = "تحتاج إلى تحسين"
                recommendation = "يرجى الاهتمام أكثر بالتقييمات ⚠️"
            
            results.append({
                "student_name": str(student_name).strip(),
                "subject": subject,
                "level": level,
                "section": section,
                "total_assessments": total_assessments,
                "completed_assessments": completed_count,
                "pending_assessments": pending_count,
                "pending_titles": ", ".join(pending_titles) if pending_titles else "لا يوجد",
                "solve_pct": solve_pct,
                "overall_pct": overall_pct,
                "category": category,
                "recommendation": recommendation
            })
        
        return results
    
    except Exception as e:
        st.error(f"خطأ في تحليل الورقة {sheet_name}: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return []

def generate_html_report(student_data):
    """Generate HTML report for a student"""
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_data['student_name']}</title>
        <style>
            body {{
                font-family: 'Arial', sans-serif;
                direction: rtl;
                text-align: right;
                margin: 40px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: #333;
            }}
            .container {{
                max-width: 800px;
                margin: 0 auto;
                background: white;
                padding: 40px;
                border-radius: 20px;
                box-shadow: 0 10px 50px rgba(0,0,0,0.3);
            }}
            h1 {{
                color: #667eea;
                text-align: center;
                border-bottom: 3px solid #667eea;
                padding-bottom: 20px;
            }}
            .info-box {{
                background: #f8f9fa;
                padding: 20px;
                border-radius: 10px;
                margin: 20px 0;
            }}
            .metric {{
                display: inline-block;
                margin: 10px 20px;
                padding: 15px;
                background: white;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .category {{
                font-size: 24px;
                font-weight: bold;
                text-align: center;
                padding: 20px;
                border-radius: 10px;
                margin: 20px 0;
            }}
            .platinum {{ background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; }}
            .gold {{ background: linear-gradient(135deg, #ffd89b 0%, #19547b 100%); color: white; }}
            .silver {{ background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%); color: #333; }}
            .bronze {{ background: linear-gradient(135deg, #ff9a56 0%, #ff6a88 100%); color: white; }}
            .needs-improvement {{ background: linear-gradient(135deg, #ff6b6b 0%, #c92a2a 100%); color: white; }}
            @media print {{
                body {{ background: white; }}
                .container {{ box-shadow: none; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 تقرير التقييمات الأسبوعية</h1>
            
            <div class="info-box">
                <h2>👤 معلومات الطالب</h2>
                <p><strong>الاسم:</strong> {student_data['student_name']}</p>
                <p><strong>المادة:</strong> {student_data['subject']}</p>
                <p><strong>المستوى:</strong> {student_data['level']}</p>
                <p><strong>الشعبة:</strong> {student_data['section']}</p>
            </div>
            
            <div class="info-box">
                <h2>📈 الإحصائيات</h2>
                <div class="metric">
                    <strong>✅ منجز:</strong> {student_data['completed_assessments']}
                </div>
                <div class="metric">
                    <strong>📊 الإجمالي:</strong> {student_data['total_assessments']}
                </div>
                <div class="metric">
                    <strong>⏳ متبقي:</strong> {student_data['pending_assessments']}
                </div>
                <div class="metric">
                    <strong>📈 النسبة:</strong> {student_data['solve_pct']:.1f}%
                </div>
            </div>
            
            <div class="category {'platinum' if student_data['category'] == 'البلاتينية' else 'gold' if student_data['category'] == 'الذهبي' else 'silver' if student_data['category'] == 'الفضي' else 'bronze' if student_data['category'] == 'البرونزي' else 'needs-improvement'}">
                🏆 {student_data['category']}
            </div>
            
            <div class="info-box">
                <h2>💡 التوصية</h2>
                <p style="font-size: 18px;">{student_data['recommendation']}</p>
            </div>
            
            {f'''<div class="info-box">
                <h2>📝 التقييمات المتبقية</h2>
                <p>{student_data['pending_titles']}</p>
            </div>''' if student_data['pending_titles'] != 'لا يوجد' else ''}
            
            <p style="text-align: center; color: #999; margin-top: 40px;">
                تم الإنشاء بتاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M')}
            </p>
        </div>
    </body>
    </html>
    """
    return html

# ================== MAIN APP ==================

st.title("📊 محلل التقييمات الأسبوعية")
st.markdown("---")

# Initialize session state
if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

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
                "اختر الأوراق",
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
    2. **اختر الأوراق** المراد تحليلها
    3. **اضغط على "تشغيل التحليل"**
    
    ### 📋 متطلبات الملف
    - الصف 1: عناوين التقييمات (من H1 يميناً)
    - الصف 3: تواريخ الاستحقاق (من H3 يميناً)
    - الصف 5 فما بعد: أسماء الطلاب والنتائج
    - اسم الورقة: المادة + المستوى + الشعبة
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
                st.session_state.analysis_results = pd.DataFrame(all_results)
                st.success(f"✅ تم تحليل {len(all_results)} طالب من {len(selected_sheets)} ورقة!")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")
        
        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")
            st.exception(e)

# Display results
if st.session_state.analysis_results is not None:
    df = st.session_state.analysis_results
    
    # Summary metrics
    st.markdown("## 📈 الإحصائيات العامة")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("👥 الطلاب", len(df))
    with col2:
        st.metric("📚 المواد", df['subject'].nunique())
    with col3:
        st.metric("📈 متوسط الإنجاز", f"{df['solve_pct'].mean():.1f}%")
    with col4:
        st.metric("🏆 البلاتينية", len(df[df['category'] == 'البلاتينية']))
    
    st.divider()
    
    # Summary by subject
    st.subheader("📊 ملخص حسب المادة")
    
    subject_summary = df.groupby('subject').agg({
        'student_name': 'count',
        'solve_pct': 'mean',
        'completed_assessments': 'sum',
        'total_assessments': 'first'
    }).round(2)
    
    subject_summary.columns = ['عدد الطلاب', 'متوسط النسبة %', 'إجمالي المنجز', 'إجمالي التقييمات']
    st.dataframe(subject_summary, use_container_width=True)
    
    st.divider()
    
    # Data table
    st.subheader("📋 البيانات التفصيلية")
    
    display_df = df.copy()
    display_df['solve_pct'] = display_df['solve_pct'].apply(lambda x: f"{x:.1f}%")
    
    st.dataframe(display_df, use_container_width=True, height=400)
    
    # Download CSV
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        "📥 تحميل CSV",
        csv,
        f"analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        "text/csv"
    )
    
    st.divider()
    
    # Charts
    st.subheader("📈 الرسوم البيانية")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.text("توزيع الفئات")
        fig, ax = plt.subplots(figsize=(8, 5))
        
        category_counts = df['category'].value_counts()
        colors_map = {
            'البلاتينية': '#f093fb',
            'الذهبي': '#ffd89b',
            'الفضي': '#a8edea',
            'البرونزي': '#ff9a56',
            'تحتاج إلى تحسين': '#ff6b6b'
        }
        
        bar_colors = [colors_map.get(cat, '#999') for cat in category_counts.index]
        category_counts.plot(kind='bar', ax=ax, color=bar_colors, edgecolor='black')
        
        ax.set_xlabel("الفئة")
        ax.set_ylabel("العدد")
        ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
        plt.tight_layout()
        st.pyplot(fig)
    
    with col2:
        st.text("توزيع النسب")
        fig, ax = plt.subplots(figsize=(8, 5))
        
        ax.hist(df['solve_pct'], bins=15, color='#667eea', edgecolor='black')
        ax.set_xlabel("نسبة الإنجاز %")
        ax.set_ylabel("عدد الطلاب")
        ax.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig)
    
    st.divider()
    
    # Generate HTML reports
    st.subheader("📄 تقارير HTML")
    
    if st.button("🔄 إنشاء تقارير HTML", use_container_width=True):
        with st.spinner(f"جاري إنشاء {len(df)} تقرير..."):
            try:
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for _, row in df.iterrows():
                        html = generate_html_report(row)
                        filename = f"{row['subject']}_{row['level']}_{row['section']}_{row['student_name']}.html"
                        zf.writestr(filename, html.encode('utf-8'))
                
                zip_buffer.seek(0)
                
                st.download_button(
                    f"📦 تحميل {len(df)} تقرير",
                    zip_buffer.getvalue(),
                    f"reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    "application/zip"
                )
                
                st.success("✅ تم إنشاء التقارير!")
            
            except Exception as e:
                st.error(f"❌ خطأ: {str(e)}")

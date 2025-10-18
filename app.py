import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
from datetime import datetime

# Page config
st.set_page_config(
    page_title="Weekly Assessments Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================== HELPER FUNCTIONS ==================

def col_to_index(col_letter):
    """Convert column letter to zero-based index"""
    return sum((ord(c) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(col_letter.upper()))) - 1

def categorize_student(solve_pct):
    """Categorize student based on solve percentage"""
    if solve_pct >= 90:
        return "البلاتينية", "أداء ممتاز! استمر في التميز 🌟"
    elif solve_pct >= 80:
        return "الذهبي", "أداء جيد جداً، حافظ على مستواك 🥇"
    elif solve_pct >= 70:
        return "الفضي", "أداء جيد، يمكنك التحسن أكثر 🥈"
    elif solve_pct >= 60:
        return "البرونزي", "أداء مقبول، تحتاج لمزيد من الجهد 🥉"
    else:
        return "تحتاج إلى تحسين", "يرجى الاهتمام أكثر بالتقييمات ⚠️"

def analyze_excel_file(file, sheet_name, start_col, names_row, names_col, due_row):
    """Analyze a single Excel sheet"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # Extract metadata
        subject = df.iloc[0, 0] if pd.notna(df.iloc[0, 0]) else "غير محدد"
        level = df.iloc[1, 0] if pd.notna(df.iloc[1, 0]) else "غير محدد"
        section = df.iloc[1, 1] if pd.notna(df.iloc[1, 1]) else "غير محدد"
        
        # Get column indices
        start_col_idx = col_to_index(start_col)
        names_col_idx = col_to_index(names_col)
        
        results = []
        
        # Process each student
        for idx, row in df.iterrows():
            if idx < names_row - 1:
                continue
                
            student_name = row.iloc[names_col_idx]
            
            if pd.isna(student_name) or student_name == "" or str(student_name).strip() == "":
                continue
            
            # Count assessments
            total_solved = 0
            total_assessments = 0
            unsolved_titles = []
            
            for col_idx in range(start_col_idx, len(row)):
                assessment_title = df.iloc[due_row - 1, col_idx]
                
                if pd.notna(assessment_title) and str(assessment_title).strip() != "":
                    cell_value = row.iloc[col_idx]
                    
                    if pd.notna(cell_value):
                        if isinstance(cell_value, (int, float)):
                            if cell_value > 0:
                                total_solved += 1
                            else:
                                total_assessments += 1
                                unsolved_titles.append(str(assessment_title))
                        elif str(cell_value).strip().lower() in ['تم', 'done', 'x', '✓']:
                            total_solved += 1
                        else:
                            total_assessments += 1
                            unsolved_titles.append(str(assessment_title))
                    else:
                        total_assessments += 1
                        unsolved_titles.append(str(assessment_title))
            
            # Calculate percentage
            total = total_solved + total_assessments
            solve_pct = (total_solved / total * 100) if total > 0 else 0
            
            # Categorize
            category, recommendation = categorize_student(solve_pct)
            
            results.append({
                "student_name": str(student_name),
                "subject": str(subject),
                "class": str(level),
                "section": str(section),
                "total_material_solved": total_solved,
                "total_assessments": total_assessments,
                "unsolved_assessment_count": len(unsolved_titles),
                "unsolved_titles": ", ".join(unsolved_titles) if unsolved_titles else "لا يوجد",
                "solve_pct": solve_pct,
                "category": category,
                "recommendation": recommendation
            })
        
        return results
    
    except Exception as e:
        st.error(f"خطأ في تحليل الورقة {sheet_name}: {str(e)}")
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
                <p><strong>المستوى:</strong> {student_data['class']}</p>
                <p><strong>الشعبة:</strong> {student_data['section']}</p>
            </div>
            
            <div class="info-box">
                <h2>📈 الإحصائيات</h2>
                <div class="metric">
                    <strong>✅ منجز:</strong> {student_data['total_material_solved']}
                </div>
                <div class="metric">
                    <strong>⏳ متبقي:</strong> {student_data['total_assessments']}
                </div>
                <div class="metric">
                    <strong>📊 النسبة:</strong> {student_data['solve_pct']:.2f}%
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
                <h2>📝 التقييمات غير المنجزة</h2>
                <p>{student_data['unsolved_titles']}</p>
            </div>''' if student_data['unsolved_titles'] != 'لا يوجد' else ''}
            
            <p style="text-align: center; color: #999; margin-top: 40px;">
                تم الإنشاء بتاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M')}
            </p>
        </div>
    </body>
    </html>
    """
    return html

# ================== MAIN APP ==================

st.title("📊 Weekly Assessments Analyzer")
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
                default=sheets if len(sheets) <= 3 else sheets[:3]
            )
        except Exception as e:
            st.error(f"خطأ: {e}")
            selected_sheets = []
    else:
        selected_sheets = []
    
    st.divider()
    
    # Parameters
    st.subheader("🔧 معاملات التحليل")
    
    col1, col2 = st.columns(2)
    
    with col1:
        start_col = st.text_input("عمود التقييمات", value="H").upper()
        names_row = st.number_input("صف الأسماء", value=5, min_value=1)
    
    with col2:
        names_col = st.text_input("عمود الأسماء", value="A").upper()
        due_row = st.number_input("صف التواريخ", value=3, min_value=1)
    
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
    3. **اضبط المعاملات** إذا لزم الأمر
    4. **اضغط على "تشغيل التحليل"**
    
    ### 📋 متطلبات الملف
    - أسماء الطلاب في العمود A
    - التقييمات تبدأ من العمود H
    - تواريخ الاستحقاق في الصف 3
    """)

elif run_analysis:
    with st.spinner("جاري التحليل..."):
        try:
            all_results = []
            
            for file in uploaded_files:
                for sheet in selected_sheets:
                    results = analyze_excel_file(
                        file, sheet, start_col, names_row, names_col, due_row
                    )
                    all_results.extend(results)
            
            if all_results:
                st.session_state.analysis_results = pd.DataFrame(all_results)
                st.success(f"✅ تم تحليل {len(all_results)} طالب!")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")
        
        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")
            st.exception(e)

# Display results
if st.session_state.analysis_results is not None:
    df = st.session_state.analysis_results
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("👥 الطلاب", len(df))
    with col2:
        st.metric("📈 المتوسط", f"{df['solve_pct'].mean():.1f}%")
    with col3:
        st.metric("🏆 البلاتينية", len(df[df['category'] == 'البلاتينية']))
    with col4:
        st.metric("⚠️ يحتاج تحسين", len(df[df['category'] == 'تحتاج إلى تحسين']))
    
    st.divider()
    
    # Data table
    st.subheader("📊 البيانات")
    
    display_df = df.copy()
    display_df['solve_pct'] = display_df['solve_pct'].apply(lambda x: f"{x:.2f}%")
    
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
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("توزيع الفئات")
        fig, ax = plt.subplots(figsize=(8, 5))
        
        category_counts = df['category'].value_counts()
        colors = {'البلاتينية': '#f093fb', 'الذهبي': '#ffd89b', 
                  'الفضي': '#a8edea', 'البرونزي': '#ff9a56', 
                  'تحتاج إلى تحسين': '#ff6b6b'}
        
        bar_colors = [colors.get(cat, '#999') for cat in category_counts.index]
        category_counts.plot(kind='bar', ax=ax, color=bar_colors)
        
        ax.set_xlabel("الفئة")
        ax.set_ylabel("العدد")
        plt.xticks(rotation=45)
        st.pyplot(fig)
    
    with col2:
        st.subheader("توزيع النسب")
        fig, ax = plt.subplots(figsize=(8, 5))
        
        ax.hist(df['solve_pct'], bins=15, color='#667eea', edgecolor='black')
        ax.set_xlabel("نسبة الإنجاز %")
        ax.set_ylabel("عدد الطلاب")
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
                        filename = f"{row['subject']}_{row['student_name']}.html"
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

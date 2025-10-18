import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
from datetime import datetime
import re

# ========= إعداد الصفحة =========
st.set_page_config(
    page_title="محلل التقييمات الأسبوعية",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========= ثيم عنّابي/أبيض (CSS) =========
PRIMARY = "#8A1538"  # عنّابي قطر
PRIMARY_DARK = "#6b0f2b"
ACCENT = "#D9B3C2"
BG_SOFT = "#FBF9FA"
CARD_BG = "#FFFFFF"

st.markdown(
    f"""
    <style>
    html, body, [class*="css"] {{
      font-family: "Tajawal", "Cairo", "DejaVu Sans", Arial, sans-serif !important;
    }}
    /* رأس الصفحة والفواصل */
    h1, h2, h3, h4 {{
      color: {PRIMARY} !important;
    }}
    .block-container {{
      padding-top: 1.2rem;
      padding-bottom: 2rem;
      background: {BG_SOFT};
    }}
    /* أزرار */
    .stButton>button {{
      background: {PRIMARY};
      color: white;
      border-radius: 10px;
      border: 1px solid {PRIMARY_DARK};
    }}
    .stButton>button:hover {{
      background: {PRIMARY_DARK};
      color: #fff;
      border-color: {PRIMARY_DARK};
    }}
    /* الجداول */
    thead tr th {{
      background-color: {PRIMARY} !important;
      color: #fff !important;
      font-weight: 700 !important;
      border: 1px solid {PRIMARY_DARK} !important;
    }}
    tbody tr td {{
      border: 1px solid #eee !important;
    }}
    /* بطاقات */
    .card {{
      background: {CARD_BG};
      border: 1px solid #eee;
      border-radius: 14px;
      padding: 16px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }}
    /* شارة عنابية */
    .chip {{
      display:inline-block; 
      padding:6px 12px; 
      margin:4px 6px; 
      border-radius: 999px; 
      background:#fff; 
      color:{PRIMARY}; 
      border:1px solid {PRIMARY}; 
      font-weight:600; 
      font-size:13px;
    }}
    /* شريط نسبة الإنجاز */
    .badge-overall {{
      background:{PRIMARY};
      color:#fff;
      padding:8px 12px;
      border-radius:999px;
      display:inline-block;
      font-weight:700;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# ========= دعم العربية في الرسوم (اختياري) =========
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    def shape_ar(text: str) -> str:
        if text is None:
            return ""
        return get_display(arabic_reshaper.reshape(str(text)))
except Exception:
    def shape_ar(text: str) -> str:
        return str(text) if text is not None else ""

import matplotlib
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['axes.unicode_minus'] = False

# ================== دوال مساعدة ==================

def normalize_ar_name(s: str) -> str:
    """
    تطبيع أسماء الطلاب لمنع التكرار:
    - إزالة التطويل (ـ) والمدّ.
    - توحيد الألف (أ/إ/آ -> ا).
    - إزالة المسافات الزائدة والرموز غير المفيدة.
    """
    if s is None:
        return ""
    s = str(s)
    # إزالة التطويل
    s = s.replace("ـ", "")
    # توحيد الألف
    s = re.sub(r"[إأآٱ]", "ا", s)
    # إزالة التشكيل البسيط
    s = re.sub(r"[ًٌٍَُِّْـ]", "", s)
    # إزالة مسافات مضاعفة
    s = " ".join(s.split())
    return s.strip()

def parse_sheet_name(sheet_name: str):
    """استخراج المادة/الصف/الشعبة من اسم الورقة"""
    parts = sheet_name.strip().split()
    level, section = "", ""
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
    """
    تحليل ورقة واحدة — منطق العد كما هو:
    - نعد M فقط كـ (غير منجز)،
    - completed_count = total_assessments - m_count
    """
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

        # عناوين التقييمات من الصف 1 (index 0) بدءًا من العمود H (index 7)
        assessment_titles = []
        for col_idx in range(7, df.shape[1]):
            title = df.iloc[0, col_idx]
            if pd.notna(title) and str(title).strip():
                assessment_titles.append(str(title).strip())

        total_assessments = len(assessment_titles)
        results = []

        # الطلاب من الصف 5 (index 4)
        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue

            # تطبيع الاسم لمنع التكرار
            student_name_clean = normalize_ar_name(student_name)

            # --- منطقك الأصلي ---
            m_count = 0
            pending_titles = []

            for i, col_idx in enumerate(range(7, 7 + total_assessments)):
                if col_idx < df.shape[1]:
                    cell_value = df.iloc[idx, col_idx]
                    if pd.notna(cell_value):
                        if str(cell_value).strip().upper() == 'M':
                            m_count += 1
                            if i < len(assessment_titles):
                                pending_titles.append(assessment_titles[i])

            completed_count = total_assessments - m_count
            solve_pct = (completed_count / total_assessments) * 100 if total_assessments > 0 else 0.0

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

def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    """Pivot: صف واحد لكل طالب بلا تكرار (بعد تطبيع الاسم)"""
    # تأكيد التطبيع لو جاءت بيانات خارجية
    df = df.copy()
    df["student_name"] = df["student_name"].apply(normalize_ar_name)

    df_clean = df.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'], keep='first')

    unique_students = df_clean[['student_name', 'level', 'section']].drop_duplicates()
    unique_students = unique_students.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
    result = unique_students.copy()

    subjects = sorted(df_clean['subject'].unique())
    for subject in subjects:
        subject_df = df_clean[df_clean['subject'] == subject][[
            'student_name', 'level', 'section',
            'total_count', 'completed_count', 'pending_titles', 'solve_pct'
        ]].copy()

        subject_df = subject_df.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')

        subject_df = subject_df.rename(columns={
            'total_count': f"{subject} - إجمالي التقييمات",
            'completed_count': f"{subject} - المنجز",
            'pending_titles': f"{subject} - عناوين التقييمات المتبقية",
            'solve_pct': f"{subject} - نسبة الإنجاز %"
        })

        result = result.merge(subject_df, on=['student_name', 'level', 'section'], how='left')

    # النسبة العامة والفئة
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

    # إزالة أي أعمدة مكررة + دمج تكرار الصفوف (لو بقي فرق بسيط بالشعبة/الصف الفارغ)
    result = result.loc[:, ~result.columns.duplicated()]
    result['الصف'] = result['الصف'].fillna("").astype(str).str.strip()
    result['الشعبة'] = result['الشعبة'].fillna("").astype(str).str.strip()
    result = result.drop_duplicates(subset=['اسم الطالب', 'الصف', 'الشعبة'], keep='first').reset_index(drop=True)
    return result

def generate_student_html_report(student_row: pd.Series, school_name="", coordinator="", academic="", admin="", principal="", logo_base64="") -> str:
    """تقرير الطالب — شعار يمين، نسبة إنجاز، مواد كـ شارات عنابية."""
    student_name = student_row['اسم الطالب']
    level = student_row['الصف']
    section = student_row['الشعبة']

    total_assessments = 0
    total_completed = 0
    subjects_html = ""
    subject_list = []

    for col in student_row.index:
        if ' - إجمالي التقييمات' in col:
            subject = col.replace(' - إجمالي التقييمات', '')
            subject_list.append(subject)
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
                    <td style="text-align: right; padding: 10px; border: 1px solid #eee;">{subject}</td>
                    <td style="text-align: center; padding: 10px; border: 1px solid #eee;">{total}</td>
                    <td style="text-align: center; padding: 10px; border: 1px solid #eee;">{completed}</td>
                    <td style="text-align: right; padding: 10px; border: 1px solid #eee;">{pending_titles}</td>
                </tr>
                """

    # النسبة العامة
    if 'نسبة حل التقييمات في جميع المواد' in student_row.index and pd.notna(student_row['نسبة حل التقييمات في جميع المواد']):
        overall_pct = float(student_row['نسبة حل التقييمات في جميع المواد'])
    else:
        overall_pct = (total_completed / total_assessments * 100) if total_assessments > 0 else 0.0

    remaining = total_assessments - total_completed

    if overall_pct == 0:
        recommendation = "الطالب لم يستفد من النظام - يرجى التواصل مع ولي الأمر فورًا 🚫"
        category_color = "#9E9E9E"
    elif overall_pct >= 90:
        recommendation = "أداء ممتاز! استمر في التميز 🌟"
        category_color = PRIMARY
    elif overall_pct >= 80:
        recommendation = "أداء جيد جدًا، حافظ على مستواك 👍"
        category_color = "#A63D5C"
    elif overall_pct >= 70:
        recommendation = "أداء جيد، يمكنك التحسن أكثر ✓"
        category_color = "#C97286"
    elif overall_pct >= 60:
        recommendation = "أداء مقبول، تحتاج لمزيد من الجهد ⚠️"
        category_color = "#E09BAC"
    else:
        recommendation = "يرجى الاهتمام أكثر بالتقييمات ومراجعة المواد"
        category_color = "#F05C6B"

    logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 80px; margin-bottom: 10px;" />' if logo_base64 else ""
    school_section = f"<h2 style='color: {PRIMARY}; margin: 0;'>{school_name}</h2>" if school_name else ""

    header_html = f"""
        <div style="display:flex; flex-direction:row-reverse; align-items:center; justify-content:space-between; gap:10px;">
            <div style="min-width:100px; text-align:right;">{logo_html}</div>
            <div style="flex:1; text-align:right;">
                {school_section}
                <h1 style="color:{PRIMARY}; margin:5px 0 0 0; font-size:24px;">📊 تقرير أداء الطالب - نظام قطر للتعليم</h1>
            </div>
        </div>
    """

    # شارات المواد (عنابية)
    subjects_badge = ""
    if subject_list:
        chips = "".join([f"<span class='chip'>{subj}</span>" for subj in subject_list])
        subjects_badge = f"""
        <div style="background:#F2E8EC; border:1px solid {ACCENT}; padding:10px; border-radius:10px; margin: 10px 0;">
            <strong style="color:{PRIMARY};">المواد:</strong> {chips}
        </div>
        """

    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>تقرير {student_name}</title>
        <style>
            @page {{ size: A4; margin: 14mm; }}
            body {{ font-family: "Tajawal","Cairo","DejaVu Sans", Arial, sans-serif; direction: rtl; padding: 20px; background: {BG_SOFT}; }}
            .container {{ max-width: 840px; margin: 0 auto; background: {CARD_BG}; padding: 24px 28px; border: 1px solid #eee; border-radius: 14px; }}
            .header {{ border-bottom: 3px solid {PRIMARY}; padding-bottom: 16px; margin-bottom: 22px; }}
            .student-info {{ background: #F3F7FB; padding: 16px; border-radius: 10px; margin-bottom: 18px; }}
            .student-info h3 {{ margin: 0 0 10px 0; color: {PRIMARY}; }}
            table {{ width: 100%; border-collapse: collapse; margin: 14px 0; }}
            th {{ background: {PRIMARY}; color: white; padding: 10px; text-align: center; border: 1px solid {PRIMARY_DARK}; font-size: 14px; }}
            td {{ padding: 10px; border: 1px solid #eee; }}
            tr:nth-child(even) {{ background-color: #fafafa; }}
            .stats-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 10px; }}
            .stat-box {{ background: #fff; border: 1px solid #eee; padding: 14px; border-radius: 10px; text-align: center; }}
            .stat-value {{ font-size: 26px; font-weight: bold; color: {PRIMARY}; }}
            .stat-label {{ font-size: 13px; color: #666; margin-top: 6px; }}
            .overall-badge {{ background: {PRIMARY}; color: white; padding: 8px 12px; border-radius: 999px; font-weight: 700; display:inline-block; }}
            .recommendation {{ background: {category_color}; color: white; padding: 14px; border-radius: 10px; margin: 16px 0; text-align: center; font-size: 15px; font-weight: 700; }}
            .signatures {{ margin-top: 24px; border-top: 2px solid #eee; padding-top: 16px; }}
            .signature-line {{ margin: 10px 0; font-size: 14px; }}
            @media print {{
                body {{ background: white; padding: 0; }}
                .container {{ box-shadow: none; max-width: 100%; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">{header_html}</div>

            <div class="student-info">
                <h3>معلومات الطالب</h3>
                <p><strong>اسم الطالب:</strong> {student_name}</p>
                <p><strong>الصف:</strong> {level} &nbsp;&nbsp;&nbsp; <strong>الشعبة:</strong> {section}</p>
                <div class="overall-badge">نسبة الإنجاز الكلية: {overall_pct:.1f}%</div>
                {subjects_badge}
            </div>

            <table>
                <thead>
                    <tr>
                        <th>المادة</th>
                        <th>عدد التقييمات الإجمالي</th>
                        <th>عدد التقييمات المنجزة</th>
                        <th>عناوين التقييمات المتبقية</th>
                    </tr>
                </thead>
                <tbody>
                    {subjects_html}
                </tbody>
            </table>

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
                    <div class="stat-value">{overall_pct:.1f}%</div>
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

                <p style="text-align: center; color: #999; margin-top: 16px; font-size: 12px;">
                    تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}
                </p>
            </div>
        </div>
    </body>
    </html>
    """
    return html

# ================== الواجهة الرئيسية ==================

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

    st.subheader("🏫 معلومات المدرسة")
    school_name = st.text_input("اسم المدرسة", value="", placeholder="مثال: مدرسة قطر النموذجية")

    st.subheader("🖼️ شعار الوزارة/المدرسة")
    uploaded_logo = st.file_uploader("ارفع شعار (اختياري)", type=["png", "jpg", "jpeg"], help="سيظهر الشعار في رأس التقارير")
    logo_base64 = ""
    if uploaded_logo:
        import base64
        logo_bytes = uploaded_logo.read()
        logo_base64 = base64.b64encode(logo_bytes).decode()
        st.success("✅ تم رفع الشعار")

    st.subheader("✍️ التوقيعات")
    coordinator_name = st.text_input("منسق المشاريع", value="سحر عثمان")
    academic_deputy = st.text_input("النائب الأكاديمي", value="مريم القضع")
    admin_deputy = st.text_input("النائب الإداري", value="دلال الفهيدة")
    principal_name = st.text_input("مدير المدرسة", value="منيرة الهاجري")

    st.divider()
    run_analysis = st.button(
        "🚀 تشغيل التحليل",
        use_container_width=True,
        type="primary",
        disabled=not (uploaded_files and selected_sheets)
    )

# إرشاد أولي
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

                # إزالة أعمدة مكررة والتأكيد على عدم تكرار الأسماء
                pivot = pivot.loc[:, ~pivot.columns.duplicated()]
                st.session_state.pivot_table = pivot
                st.success(f"✅ تم تحليل {len(pivot)} طالب فريد من {len(selected_sheets)} مادة!")
            else:
                st.warning("⚠️ لم يتم العثور على بيانات")

        except Exception as e:
            st.error(f"❌ خطأ: {str(e)}")

# ================== النتائج والرسوم ==================
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
        needs_improvement = len(pivot[pivot['الفئة'].str.contains('يحتاج تحسين', na=False)]) if 'الفئة' in pivot.columns else 0
        st.metric("⚠️ يحتاج تحسين", needs_improvement)

    st.divider()

    st.subheader("📊 البيانات التفصيلية")
    display_pivot = pivot.copy().loc[:, ~pivot.columns.duplicated()]
    for col in display_pivot.columns:
        if 'نسبة' in col:
            display_pivot[col] = display_pivot[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    display_pivot = display_pivot.fillna("-")
    st.dataframe(display_pivot, use_container_width=True, height=520)

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

    # ===== الرسوم البيانية (عربي) =====
    st.subheader("📈 الرسوم البيانية")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📊 متوسط الإنجاز حسب المادة**")
        fig, ax = plt.subplots(figsize=(10, 6))
        subject_avg = df.groupby('subject')['solve_pct'].mean().sort_values(ascending=True)
        y_pos = range(len(subject_avg))
        bars = ax.barh(list(y_pos), subject_avg.values, edgecolor='black', linewidth=1.2)
        ax.set_yticks(list(y_pos))
        ax.set_yticklabels([shape_ar(s) for s in subject_avg.index], fontsize=11)
        for i, (bar, value) in enumerate(zip(bars, subject_avg.values)):
            ax.text(value + 1, i, f'{value:.1f}%', va='center', fontsize=11, fontweight='bold')
        ax.set_xlabel(shape_ar("نسبة الإنجاز (%)"), fontsize=12, fontweight='bold')
        ax.set_title(shape_ar("متوسط الإنجاز حسب المادة"), fontsize=14, fontweight='bold', pad=16)
        ax.grid(axis='x', alpha=0.25, linestyle='--')
        ax.set_xlim(0, max(100, (subject_avg.values.max() if len(subject_avg) else 100) + 10))
        plt.tight_layout()
        st.pyplot(fig)

    with col2:
        st.markdown("**📈 توزيع النسب الإجمالية**")
        if 'نسبة حل التقييمات في جميع المواد' in pivot.columns:
            fig, ax = plt.subplots(figsize=(10, 6))
            overall_scores = pivot['نسبة حل التقييمات في جميع المواد'].dropna()
            n, bins, patches = ax.hist(overall_scores, bins=20, edgecolor='black', linewidth=1.2)
            mean_val = overall_scores.mean() if len(overall_scores) else 0
            ax.axvline(mean_val, color=PRIMARY, linestyle='--', linewidth=2.0,
                       label=shape_ar(f'المتوسط: {mean_val:.1f}%'), zorder=10)
            ax.set_xlabel(shape_ar("نسبة الإنجاز (%)"), fontsize=12, fontweight='bold')
            ax.set_ylabel(shape_ar("عدد الطلاب"), fontsize=12, fontweight='bold')
            ax.set_title(shape_ar("توزيع الأداء العام"), fontsize=14, fontweight='bold', pad=16)
            ax.legend(fontsize=11, loc='upper left')
            ax.grid(axis='y', alpha=0.25, linestyle='--')
            plt.tight_layout()
            st.pyplot(fig)

    st.divider()

    # ===== رسم بياني للفئات =====
    st.subheader("🏆 توزيع الفئات")
    view_type = st.radio("اختر نوع الرسم:", ["دائري (Donut)", "أعمدة (Bar)"], horizontal=True)
    if 'الفئة' in pivot.columns:
        cat_order = [
            "البلاتينية 🥇",
            "الذهبي 🥈",
            "الفضي 🥉",
            "البرونزي",
            "يحتاج تحسين ⚠️",
            "لا يستفيد من النظام 🚫",
            "-"
        ]
        cat_counts = pivot['الفئة'].fillna("-").value_counts().reindex(cat_order, fill_value=0)
        total_students = int(cat_counts.sum())
        st.caption(f"إجمالي الطلاب المحسوبين: {total_students}")

        if view_type == "دائري (Donut)":
            fig, ax = plt.subplots(figsize=(8, 6))
            labels = [shape_ar(lbl) for lbl in cat_counts.index]
            values = cat_counts.values
            wedges, texts, autotexts = ax.pie(
                values,
                labels=labels,
                autopct=lambda p: f"{p:.1f}%\n({int(round(p*total_students/100))})" if p > 0 else "",
                startangle=90,
                pctdistance=0.78,
                wedgeprops=dict(linewidth=1.2, edgecolor='white'),
                textprops=dict(fontsize=11)
            )
            centre_circle = plt.Circle((0, 0), 0.55, fc='white')
            fig.gca().add_artist(centre_circle)
            ax.set_title(shape_ar("توزيع الطلاب حسب الفئات"), fontsize=14, fontweight='bold', pad=12)
            st.pyplot(fig)
        else:
            fig, ax = plt.subplots(figsize=(10, 6))
            y_pos = range(len(cat_counts))
            bars = ax.barh(list(y_pos), cat_counts.values, edgecolor='black', linewidth=1.2)
            ax.set_yticks(list(y_pos))
            ax.set_yticklabels([shape_ar(c) for c in cat_counts.index], fontsize=11)
            for i, (bar, val) in enumerate(zip(bars, cat_counts.values)):
                ax.text(bar.get_width() + max(1, total_students*0.01), i, f"{val}", va='center', fontsize=11, fontweight='bold')
            ax.set_xlabel(shape_ar("عدد الطلاب"), fontsize=12, fontweight='bold')
            ax.set_title(shape_ar("توزيع الطلاب حسب الفئات"), fontsize=14, fontweight='bold', pad=12)
            ax.grid(axis='x', alpha=0.25, linestyle='--')
            plt.tight_layout()
            st.pyplot(fig)
    else:
        st.info("لا يمكن رسم الفئات: عمود 'الفئة' غير موجود في الجدول المحوري.")

    st.divider()

    # ===== التحليل حسب المادة =====
    st.subheader("📚 التحليل حسب المادة")
    subjects = sorted(df['subject'].unique())
    display_subjects = [shape_ar(s) for s in subjects]
    subj_map = dict(zip(display_subjects, subjects))
    selected_subject_display = st.selectbox("اختر المادة للتحليل التفصيلي:", display_subjects, key="subject_analysis")
    selected_subject = subj_map[selected_subject_display]

    if selected_subject:
        subject_df = df[df['subject'] == selected_subject]
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("👥 عدد الطلاب", len(subject_df))
        with col2: st.metric("📈 متوسط الإنجاز", f"{subject_df['solve_pct'].mean():.1f}%")
        with col3: st.metric("🏆 أعلى نسبة", f"{subject_df['solve_pct'].max():.1f}%")
        with col4: st.metric("⚠️ أقل نسبة", f"{subject_df['solve_pct'].min():.1f}%")

        st.markdown("##### 📊 رسم بياني للمادة")
        fig, ax = plt.subplots(figsize=(12, 5))
        categories = pd.cut(
            subject_df['solve_pct'],
            bins=[0, 50, 70, 80, 90, 100],
            labels=[shape_ar('<50%'), shape_ar('50-70%'), shape_ar('70-80%'),
                    shape_ar('80-90%'), shape_ar('90-100%')]
        )
        category_counts = categories.value_counts().sort_index()
        bars = ax.bar(range(len(category_counts)), category_counts.values, edgecolor='black', linewidth=1.2)
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., h + 0.3, f'{int(h)}', ha='center', va='bottom', fontsize=12, fontweight='bold')
        ax.set_xticks(range(len(category_counts)))
        ax.set_xticklabels(category_counts.index, fontsize=11)
        ax.set_ylabel(shape_ar("عدد الطلاب"), fontsize=12, fontweight='bold')
        ax.set_title(shape_ar(f"توزيع الأداء - {selected_subject_display}"), fontsize=14, fontweight='bold', pad=16)
        ax.grid(axis='y', alpha=0.25, linestyle='--')
        plt.tight_layout()
        st.pyplot(fig)

    st.divider()

    # ===== تقارير الطلاب =====
    st.subheader("📄 تقارير الطلاب الفردية")
    col1, col2 = st.columns(2)
    with col1:
        report_type = st.radio("نوع التقرير:", ["طالب واحد", "جميع الطلاب"])
    with col2:
        selected_student = st.selectbox("اختر الطالب:", pivot['اسم الطالب'].tolist()) if report_type == "طالب واحد" else None

    if st.button("🔄 إنشاء التقارير", use_container_width=True):
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
                student_row, settings['school'], settings['coordinator'],
                settings['academic'], settings['admin'], settings['principal'], settings['logo']
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
                            row, settings['school'], settings['coordinator'],
                            settings['academic'], settings['admin'], settings['principal'], settings['logo']
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

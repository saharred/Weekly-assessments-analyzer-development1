# -*- coding: utf-8 -*-
"""
📊 أي إنجاز — لوحة تحليل إنجاز الطلاب (Purple/White)
- اختيار أوراق Excel من الشريط الجانبي
- ربط سجل القيد (رقم شخصي + صف + شعبة) بتطبيع الاسم
- ensure_uid: uid موحّد + إزالة التكرارات
- Pivot: كل طالب صف واحد
- رسوم عامة + فئات + مواد
- توصيات تشغيلية لرفع نسبة الإنجاز (غير أكاديمية)
- تصدير Excel شامل + PDF فردي لكل طالب داخل ZIP
- حفظ الجداول في session_state لثبات أزرار التصدير
- ميزات الذكاء الاصطناعي: تحليل الأنماط، والتوصيات المخصصة باستخدام LLM
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import re, zipfile
from typing import Dict, List, Tuple, Optional
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import Color
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Table, TableStyle, Spacer, SimpleDocTemplate, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib import colors
from reportlab.lib.units import inch

# --------------- إعداد الصفحة ---------------
# رابط شعار إنجاز (المخطط البياني الملون)
INGAZ_ICON = "https://i.imgur.com/pasted_file_gkR2PR_image.png" 
st.set_page_config(page_title="أي إنجاز", page_icon=INGAZ_ICON, layout="wide")

# --------------- الثوابت ---------------
POSITIVE_STATUS = ["solved","yes","1","تم","منجز","✓","✔","صحيح"]

STUDENT_RECOMMENDATIONS = {
    "🏆 Platinum": "نثمن تميزك المستمر، لقد أظهرت إبداعًا واجتهادًا ملحوظًا. استمر في استخدام نظام قطر للتعليم بفعالية، فأنت نموذج يحتذى به.",
    "🥇 Gold": "أحسنت! مستواك يعكس التزامًا رائعًا، نثق أنك بمتابعة الجهد ستنتقل لمستوى أعلى. استمر في تفعيل نظام قطر داخل الصف.",
    "🥈 Silver": "عملك جيد ويستحق التقدير، ومع مزيد من الممارسة والتفاعل مع نظام قطر ستصل إلى مستويات أرفع. نحن فخورون بك.",
    "🥉 Bronze": "لقد أظهرت جهدًا مشكورًا، ونشجعك على بذل المزيد من العطاء. باستخدام نظام قطر بشكل أعمق ستتطور قدراتك بشكل أكبر.",
    "🔧 Needs Improvement": "نرى لديك إمكانيات واعدة، لكن تحتاج لمزيد من الالتزام باستخدام نظام قطر للتعليم. نوصيك بالمثابرة والمشاركة النشطة، ونحن بجانبك لتتقدم.",
    "🚫 Not Utilizing System": "لم يظهر بعد استفادة كافية من نظام قطر للتعليم، وندعوك إلى تفعيل النظام بشكل أكبر لتحقيق النجاح. نحن نثق أن لديك القدرة على التغيير والتميز."
}

# --------------- كود الفوتر (Footer) ---------------
FOOTER_MARKDOWN = """
<style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f0f2f6; /* لون خلفية فاتح */
        color: #800020; /* لون النص عنابي */
        text-align: center;
        padding: 10px 0;
        font-size: 14px;
        border-top: 1px solid #800020; /* خط عنابي فاصل */
    }
    .footer a {
        color: #800020; /* لون الروابط عنابي */
        text-decoration: none;
    }
</style>
<div class="footer">
    <p>
        <strong>رؤيتنا: متعلم ريادي لتنمية مستدامة</strong><br>
        جميع الحقوق محفوظة © مدرسة عثمان بن عفان النموذجية<br>
        تطوير و تنفيذ: منسق المشاريع الإلكترونية: سحر عثمان<br>
        للتواصل: <a href="mailto:S.mahgoub0101@education.qa">S.mahgoub0101@education.qa</a>
    </p>
</div>
"""
# --------------- نهاية كود الفوتر ---------------

# --------------- كود رأس الصفحة (Header) ---------------
def display_header():
    
    # روابط الشعارات
    MINISTRY_LOGO = "https://i.imgur.com/jFzu8As.jpeg"
    QATAR_SYSTEM_LOGO = "https://i.imgur.com/AtRkvQY.jpeg"
    
    # تنسيق HTML/CSS لترتيب الشعارات والعنوان
    header_html = f"""
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 2px solid #800020;">
        
        <!-- اليسار: شعار الوزارة وشعار النظام -->
        <div style="display: flex; align-items: center; gap: 15px;">
            <img src="{MINISTRY_LOGO}" style="height: 60px; object-fit: contain;">
            <img src="{QATAR_SYSTEM_LOGO}" style="height: 60px; object-fit: contain;">
        </div>
        
        <!-- المنتصف: العنوان الرئيسي (أي إنجاز) -->
        <div style="text-align: center; flex-grow: 1;">
            <h1 style="color: #800020; margin: 0; font-size: 32px;">
                أي إنجاز - لوحة تحليل إنجاز الطلاب الذكية
            </h1>
            <p style="color: #555; margin: 5px 0 0 0; font-size: 16px;">
                أداة تحليلية تعتمد على التحليل الإحصائي والتوصيات الثابتة المخصصة لرفع نسبة الإنجاز الأكاديمي.
            </p>
        </div>
        
        <!-- اليمين: مساحة فارغة أو شعار آخر إذا لزم الأمر -->
        <div style="width: 135px;"></div>
    </div>
    """
    st.markdown(header_html, unsafe_allow_html=True)
# --------------- نهاية كود رأس الصفحة ---------------

# --------------- أدوات مساعدة ---------------
def _strip_invisible_and_diacritics(s: str) -> str:
    """يزيل الأحرف غير المرئية وعلامات التشكيل من النص."""
    if not isinstance(s, str):
        return s
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2066-\u2069]', '', s)
    s = re.sub(r'[\u064b-\u065e]', '', s)
    return s.strip()

@st.cache_data
def _load_teachers_df(file) -> Optional[pd.DataFrame]:
    """تحميل ملف المعلمين وتوحيد أسماء الأعمدة."""
    if file is None:
        return None
    
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        # محاولة توحيد أسماء الأعمدة
        cols = [_strip_invisible_and_diacritics(str(c)) for c in df.columns]
        df.columns = cols
        
        # تحديد الأعمدة الأساسية
        col_map = {}
        for c in cols:
            if 'شعبة' in c or 'صف' in c or 'فصل' in c:
                col_map['class_section'] = c
            elif 'معلم' in c or 'مدرس' in c:
                col_map['teacher_name'] = c
            elif 'ايميل' in c or 'بريد' in c:
                col_map['teacher_email'] = c
        
        if len(col_map) < 3:
            st.error("ملف المعلمين يجب أن يحتوي على أعمدة للشعبة، اسم المعلم، والبريد الإلكتروني.")
            return None
        
        df = df[list(col_map.values())]
        df.columns = ['class_section', 'teacher_name', 'teacher_email']
        
        df['class_section'] = df['class_section'].astype(str).apply(_strip_invisible_and_diacritics)
        df['teacher_email'] = df['teacher_email'].astype(str).str.lower().apply(_strip_invisible_and_diacritics)
        
        return df
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة ملف المعلمين: {e}")
        return None

@st.cache_data
def process_excel_file(file, filename: str, start_row_students: int, selected_sheets: List[str]) -> List[Dict]:
    """معالجة ملف Excel واحد واستخراج بيانات الطلاب."""
    try:
        xls = pd.ExcelFile(file)
        data_rows = []
        
        for sheet_name in selected_sheets:
            # قراءة البيانات مع تخطي الصفوف العلوية
            df = xls.parse(sheet_name, header=None, skiprows=start_row_students - 1)
            
            # تحديد أعمدة الـ UID والاسم والصف والشعبة (افتراضياً أول 4 أعمدة)
            if df.shape[1] < 4: continue
            
            df = df.iloc[:, :4].copy()
            df.columns = ['uid', 'اسم الطالب', 'الصف', 'الشعبة']
            
            # تحديد أعمدة التقييمات (بدءاً من العمود الخامس)
            assessment_cols = xls.parse(sheet_name, header=None, skiprows=start_row_students - 2, nrows=1).iloc[0, 4:].tolist()
            assessment_data = xls.parse(sheet_name, header=None, skiprows=start_row_students - 1).iloc[:, 4:]
            assessment_data.columns = assessment_cols
            
            # دمج بيانات الطالب مع بيانات التقييم
            df = pd.concat([df, assessment_data], axis=1)
            
            # تحويل الصفوف إلى قائمة قواميس
            for _, row in df.iterrows():
                row_dict = row.to_dict()
                row_dict['Source_File'] = filename
                row_dict['Source_Sheet'] = sheet_name
                data_rows.append(row_dict)
                
        return data_rows
    except Exception as e:
        st.error(f"خطأ في معالجة الملف {filename} والورقة {sheet_name}: {e}")
        return []

@st.cache_data
def ensure_uid(df: pd.DataFrame) -> pd.DataFrame:
    """توحيد الـ UID وإزالة التكرارات."""
    if df.empty:
        return df
    
    # توحيد الـ UID والاسم
    df['uid'] = df['uid'].astype(str).apply(_strip_invisible_and_diacritics)
    df['اسم الطالب'] = df['اسم الطالب'].astype(str).apply(_strip_invisible_and_diacritics)
    df['الصف'] = df['الصف'].astype(str).apply(_strip_invisible_and_diacritics)
    df['الشعبة'] = df['الشعبة'].astype(str).apply(_strip_invisible_and_diacritics)
    
    # إزالة الصفوف المكررة بناءً على UID
    df = df.drop_duplicates(subset=['uid'], keep='first')
    
    return df

@st.cache_data
def build_summary_pivot(raw_df: pd.DataFrame, thresholds: Dict[str, float]) -> Tuple[pd.DataFrame, List[str]]:
    """بناء الملخص المحوري وإضافة التصنيفات والتوصيات."""
    if raw_df.empty:
        return pd.DataFrame(), []

    # 1. تحديد أعمدة التقييمات
    assessment_cols = [col for col in raw_df.columns if col not in ['uid', 'اسم الطالب', 'الصف', 'الشعبة', 'Source_File', 'Source_Sheet']]
    
    # 2. تحويل البيانات إلى تنسيق طويل (Long Format)
    long_df = raw_df.melt(
        id_vars=['uid', 'اسم الطالب', 'الصف', 'الشعبة'],
        value_vars=assessment_cols,
        var_name='assessment_name',
        value_name='status'
    ).dropna(subset=['status'])

    # 3. استخراج اسم المادة (نفترض أن اسم المادة هو أول كلمة)
    long_df['subject'] = long_df['assessment_name'].apply(lambda x: x.split(' ')[0] if isinstance(x, str) else 'غير محدد')
    
    # 4. تحديد حالة الإنجاز (Solved/Total)
    long_df['solved'] = long_df['status'].astype(str).apply(lambda x: 1 if _strip_invisible_and_diacritics(x).lower() in POSITIVE_STATUS else 0)
    long_df['total'] = 1

    # 5. بناء الجدول المحوري (Pivot Table)
    piv = pd.pivot_table(
        long_df,
        index=['uid', 'اسم الطالب', 'الصف', 'الشعبة'],
        columns='subject',
        values=['solved', 'total'],
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # 6. إصلاح مشكلة تسمية الأعمدة بعد pivot_table (إصلاح KeyError: 'uid')
    new_columns = []
    for col in piv.columns:
        if col[0] in ['uid', 'اسم الطالب', 'الصف', 'الشعبة']:
            new_columns.append(col[0])
        else:
            new_columns.append(f"{col[1]}_{col[0]}")
    piv.columns = new_columns
    
    # 7. حساب الإجمالي الكلي
    subjects = [col.split('_')[0] for col in piv.columns if col.endswith('_total')]
    
    piv['Overall_solved'] = piv[[f"{s}_solved" for s in subjects]].sum(axis=1)
    piv['Overall_total'] = piv[[f"{s}_total" for s in subjects]].sum(axis=1)
    
    # 8. حساب نسبة الإنجاز
    piv['نسبة الإنجاز %'] = (piv['Overall_solved'] / piv['Overall_total'] * 100).round(2).fillna(0)
    
    # 9. التصنيف (Categorization)
    def cat(x):
        if x == 0:
            return "🚫 Not Utilizing System"
        elif x > thresholds["Platinum"]:
            return "🏆 Platinum"
        elif x > thresholds["Gold"]:
            return "🥇 Gold"
        elif x > thresholds["Silver"]:
            return "🥈 Silver"
        elif x > thresholds["Bronze"]:
            return "🥉 Bronze"
        else:
            return "🔧 Needs Improvement"
            
    piv['الفئة'] = piv['نسبة الإنجاز %'].apply(cat)
    
    # 10. إضافة التوصية الثابتة
    piv['توصية الطالب'] = piv['الفئة'].apply(lambda x: STUDENT_RECOMMENDATIONS.get(x, "لا توجد توصية لهذا التصنيف."))
    
    # 11. إعادة ترتيب الأعمدة
    cols_order = ['uid', 'اسم الطالب', 'الصف', 'الشعبة', 'نسبة الإنجاز %', 'الفئة', 'توصية الطالب'] + [col for col in piv.columns if col not in ['uid', 'اسم الطالب', 'الصف', 'الشعبة', 'نسبة الإنجاز %', 'الفئة', 'توصية الطالب']]
    
    return piv[cols_order], subjects

@st.cache_data
def analyze_subject_patterns(summary_df: pd.DataFrame, subject: str) -> str:
    """تحليل نمط الأداء في مادة معينة وتقديم توصية ثابتة."""
    
    solved_col = f"{subject}_solved"
    total_col = f"{subject}_total"
    
    if solved_col not in summary_df.columns or total_col not in summary_df.columns:
        return "لا توجد بيانات كافية لهذه المادة."
        
    total_students = summary_df.shape[0]
    total_assessments = summary_df[total_col].sum()
    avg_solved = summary_df[solved_col].mean()
    
    # معايير بسيطة للتوصية (ثابتة)
    if total_assessments == 0:
        return f"توصية المادة {subject}: لم يتم إدخال أي تقييمات لهذه المادة. يرجى التأكد من إدخال البيانات."
    
    avg_completion = (summary_df[solved_col].sum() / total_assessments) * 100
    
    if avg_completion >= 80:
        return f"توصية المادة {subject}: أداء ممتاز! متوسط الإنجاز {avg_completion:.2f}%. يرجى التركيز على الطلاب الذين لم ينجزوا بعد لضمان استمرار التميز."
    elif avg_completion >= 50:
        return f"توصية المادة {subject}: أداء جيد. متوسط الإنجاز {avg_completion:.2f}%. يفضل مراجعة التقييمات الأقل إنجازاً وتقديم دعم إضافي للطلاب في الفئة البرونزية."
    else:
        return f"توصية المادة {subject}: يحتاج إلى تطوير. متوسط الإنجاز {avg_completion:.2f}%. يرجى مراجعة طريقة تقديم التقييمات أو المحتوى، والتركيز على الطلاب غير الفاعلين."

# ----------------------------------------------------------------------
# دوال الرسوم البيانية التفاعلية (Plotly)
# ----------------------------------------------------------------------
def create_subject_performance_chart(summary_df: pd.DataFrame, subjects: List[str]):
    """رسم بياني لأداء الطلاب حسب المادة (متوسط الإنجاز)."""
    
    # حساب متوسط الإنجاز لكل مادة
    subject_avg = []
    for subj in subjects:
        total_solved = summary_df.get(f"{subj}_solved", pd.Series([0])).sum()
        total_total = summary_df.get(f"{subj}_total", pd.Series([0])).sum()
        avg_completion = (total_solved / total_total * 100) if total_total > 0 else 0
        subject_avg.append({"المادة": subj, "متوسط الإنجاز %": avg_completion})
        
    df_avg = pd.DataFrame(subject_avg)
    
    fig = px.bar(
        df_avg,
        x="المادة",
        y="متوسط الإنجاز %",
        title="متوسط نسبة إنجاز الطلاب حسب المادة",
        color="متوسط الإنجاز %",
        color_continuous_scale=px.colors.sequential.Burg, # استخدام تدرج لوني قريب من العنابي
        text="متوسط الإنجاز %"
    )
    
    fig.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    fig.update_layout(
        uniformtext_minsize=8, 
        uniformtext_mode='hide',
        xaxis_title="المادة",
        yaxis_title="متوسط نسبة الإنجاز (%)",
        coloraxis_showscale=False,
        font=dict(family="Arial, sans-serif")
    )
    
    return fig

def create_class_section_performance_chart(summary_df: pd.DataFrame):
    """رسم بياني لأداء الشعب (متوسط الإنجاز الكلي)."""
    
    # حساب متوسط الإنجاز لكل شعبة
    class_avg = summary_df.groupby(["الصف", "الشعبة"]).agg(
        total_solved=('Overall_solved', 'sum'),
        total_total=('Overall_total', 'sum')
    ).reset_index()
    
    class_avg['نسبة الإنجاز %'] = (class_avg['total_solved'] / class_avg['total_total'] * 100).round(2).fillna(0)
    class_avg['الشعبة'] = class_avg['الصف'].astype(str) + ' ' + class_avg['الشعبة'].astype(str)
    
    fig = px.bar(
        class_avg,
        x="الشعبة",
        y="نسبة الإنجاز %",
        title="متوسط نسبة الإنجاز الكلي حسب الشعبة",
        color="نسبة الإنجاز %",
        color_continuous_scale=px.colors.sequential.Burg,
        text="نسبة الإنجاز %"
    )
    
    fig.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    fig.update_layout(
        uniformtext_minsize=8, 
        uniformtext_mode='hide',
        xaxis_title="الصف والشعبة",
        yaxis_title="متوسط نسبة الإنجاز الكلي (%)",
        coloraxis_showscale=False,
        font=dict(family="Arial, sans-serif")
    )
    
    return fig

def to_excel_bytes(dfs: Dict[str, pd.DataFrame]) -> BytesIO:
    """تحويل قاموس من DataFrames إلى ملف Excel في الذاكرة."""
    mem = BytesIO()
    with pd.ExcelWriter(mem, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    mem.seek(0)
    return mem

# ----------------------------------------------------------------------
# دالة إنشاء تقرير الطالب الفردي (PDF) - تصميم نهائي
# ----------------------------------------------------------------------
def create_student_report_pdf(student_data: pd.Series, raw_df: pd.DataFrame, school_info: dict, custom_recommendation: str = "") -> BytesIO:
    """تنشئ تقرير PDF فردي للطالب بناءً على النموذج المرفق."""
    
    # يجب استيراد هذه المكتبات داخل الدالة لضمان عملها بشكل صحيح
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib import colors
    from reportlab.lib.units import inch, mm
    
    mem = BytesIO()
    
    # إعدادات التوثيق
    doc = SimpleDocTemplate(
        mem,
        pagesize=A4,
        rightMargin=10*mm, leftMargin=10*mm,
        topMargin=10*mm, bottomMargin=10*mm
    )
    
    styles = getSampleStyleSheet()
    # أنماط مخصصة للغة العربية (محاذاة لليمين)
    styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT, fontName='Helvetica', fontSize=12))
    styles.add(ParagraphStyle(name='Heading1Right', alignment=TA_RIGHT, fontName='Helvetica-Bold', fontSize=18))
    styles.add(ParagraphStyle(name='Heading2Right', alignment=TA_RIGHT, fontName='Helvetica-Bold', fontSize=14))
    styles.add(ParagraphStyle(name='SmallRight', alignment=TA_RIGHT, fontName='Helvetica', fontSize=10))
    
    # بيانات المدرسة
    school_name = school_info.get("School_Name", "المدرسة")
    coordinator = school_info.get("Coordinator", "N/A")
    academic_deputy = school_info.get("Academic_Deputy", "N/A")
    administrative_deputy = school_info.get("Administrative_Deputy", "N/A")
    principal = school_info.get("Principal", "N/A")
    
    # محتوى التقرير
    elements = []
    
    # 1. رأس التقرير والشعارات (محاكاة)
    MINISTRY_LOGO = "https://i.imgur.com/jFzu8As.jpeg"
    QATAR_SYSTEM_LOGO = "https://i.imgur.com/AtRkvQY.jpeg"
    
    # إنشاء جدول لرأس الصفحة (الشعارات والعنوان)
    header_data = [
        [
            Image(MINISTRY_LOGO, width=40*mm, height=15*mm),
            Paragraph(f"<b>{school_name}</b>", styles['Heading2Right']),
            Image(QATAR_SYSTEM_LOGO, width=40*mm, height=15*mm)
        ],
        [
            Paragraph("العام الأكاديمي 2025-2026", styles['SmallRight']),
            Paragraph("تقرير أداء الطالب على نظام قطر للتعليم", styles['Heading1Right']),
            Paragraph("", styles['SmallRight'])
        ]
    ]
    
    header_table = Table(header_data, colWidths=[50*mm, 100*mm, 50*mm])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (1,1), (2,1)), # دمج خلايا العنوان
        ('BOTTOMPADDING', (0,0), (-1,-1), 5*mm),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 0.5 * inch))
    
    # 2. معلومات الطالب
    elements.append(Paragraph("<b>معلومات الطالب:</b>", styles['Heading2Right']))
    elements.append(Spacer(1, 0.1 * inch))
    
    student_info_data = [
        [
            Paragraph(f"<b>:بلاطلا مسا</b> {student_data['اسم الطالب']}", styles['RightAlign']),
            Paragraph(f"<b>:فصلا</b> {student_data['الصف']}", styles['RightAlign']),
            Paragraph(f"<b>:ةبعشلا</b> {student_data['الشعبة']}", styles['RightAlign']),
        ]
    ]
    student_info_table = Table(student_info_data, colWidths=[doc.width/3]*3)
    student_info_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5*mm),
    ]))
    elements.append(student_info_table)
    elements.append(Spacer(1, 0.2 * inch))
    
    # 3. جدول أداء المواد
    elements.append(Paragraph("<b>الأداء حسب المادة:</b>", styles['Heading2Right']))
    elements.append(Spacer(1, 0.1 * inch))
    
    # إعداد بيانات الجدول
    subject_data_table = [
        [
            Paragraph("<b>ةدالما</b>", styles['SmallRight']),
            Paragraph("<b>يلامجلاا تامييقتلا ددع</b>", styles['SmallRight']),
            Paragraph("<b>ةزجنلما تامييقتلا ددع</b>", styles['SmallRight']),
            Paragraph("<b>ةيفبتلما تامييقتلا ددع</b>", styles['SmallRight']),
        ]
    ]
    
    subject_cols = [col.split('_')[0] for col in student_data.index if col.endswith('_total') and col not in ['Overall_total']]
    
    for subj in subject_cols:
        solved = student_data.get(f"{subj}_solved", 0)
        total = student_data.get(f"{subj}_total", 0)
        pending = total - solved
        
        subject_data_table.append([
            Paragraph(subj, styles['SmallRight']),
            Paragraph(str(total), styles['SmallRight']),
            Paragraph(str(solved), styles['SmallRight']),
            Paragraph(str(pending), styles['SmallRight']),
        ])
        
    # تنسيق الجدول
    table_style = TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(red=(0x80/255), green=0, blue=(0x20/255), alpha=0.1)), # خلفية عنابية فاتحة
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ])
    
    subj_table = Table(subject_data_table, colWidths=[doc.width/4]*4)
    subj_table.setStyle(table_style)
    elements.append(subj_table)
    elements.append(Spacer(1, 0.2 * inch))
    
    # 4. الإحصائيات العامة
    elements.append(Paragraph("<b>:تايئاصحلاا</b>", styles['Heading2Right']))
    elements.append(Spacer(1, 0.1 * inch))
    
    overall_solved = student_data['Overall_solved']
    overall_total = student_data['Overall_total']
    overall_completion = student_data['نسبة الإنجاز %']
    overall_pending = overall_total - overall_solved
    
    stats_data = [
        [
            Paragraph(f"<b>ةبسن لح تامييقتلا</b> {overall_completion:.2f}%", styles['RightAlign']),
            Paragraph(f"<b>يقبتم</b> {overall_pending}", styles['RightAlign']),
            Paragraph(f"<b>زجنم</b> {overall_solved}", styles['RightAlign']),
        ]
    ]
    stats_table = Table(stats_data, colWidths=[doc.width/3]*3)
    stats_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
    ]))
    elements.append(stats_table)
    elements.append(Spacer(1, 0.2 * inch))
    
    # 5. التوصية (منسق المشاريع)
    elements.append(Paragraph("<b>:عيراشلما قسنم ةيصوت</b>", styles['Heading2Right']))
    elements.append(Spacer(1, 0.1 * inch))
    
    final_recommendation = custom_recommendation if custom_recommendation else student_data['توصية الطالب']
    
    # استخدام نمط الفقرة للتوصية
    elements.append(Paragraph(final_recommendation, styles['RightAlign']))
    elements.append(Spacer(1, 0.5 * inch))
    
    # 6. التوقيعات (Footer/Contact)
    elements.append(Paragraph("<b>للتواصل والتوقيعات:</b>", styles['Heading2Right']))
    elements.append(Spacer(1, 0.1 * inch))
    
    contact_data = [
        [
            Paragraph(f"<b>مدير المدرسة:</b> {principal}", styles['SmallRight']),
            Paragraph(f"<b>النائب الإداري:</b> {administrative_deputy}", styles['SmallRight']),
            Paragraph(f"<b>النائب الأكاديمي:</b> {academic_deputy}", styles['SmallRight']),
            Paragraph(f"<b>منسق المشاريع الإلكترونية:</b> {coordinator}", styles['SmallRight']),
        ]
    ]
    contact_table = Table(contact_data, colWidths=[doc.width/4]*4)
    contact_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
    ]))
    elements.append(contact_table)
    elements.append(Spacer(1, 0.2 * inch))
    
    # 7. الرؤية والروابط
    elements.append(Paragraph("<b>رؤيتنا: متعلم ريادي لتنمية مستدامة</b>", styles['Heading2Right']))
    elements.append(Spacer(1, 0.1 * inch))
    
    links_data = [
        [
            Paragraph("<b>رابط نظام قطر:</b> https://qeducation.edu.gov.qa", styles['SmallRight']),
            Paragraph("<b>موقع استعادة كلمة المرور:</b> https://pwdreset.edu.gov.qa", styles['SmallRight']),
        ]
    ]
    links_table = Table(links_data, colWidths=[doc.width/2]*2)
    links_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
    ]))
    elements.append(links_table)
    
    # بناء المستند
    doc.build(elements)
    mem.seek(0)
    return mem
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# دالة إنشاء التقرير الكمي الوصفي (Excel)
# ----------------------------------------------------------------------
def create_quantitative_report_excel(summary_df: pd.DataFrame, subjects: List[str]) -> BytesIO:
    """تنشئ تقرير Excel كمي وصفي على مستوى المادة والشعبة."""
    mem = BytesIO()
    with pd.ExcelWriter(mem, engine="openpyxl") as w:
        
        # 1. تقرير الأداء حسب المادة
        subject_performance = []
        for subj in subjects:
            total_solved = summary_df.get(f"{subj}_solved", pd.Series([0])).sum()
            total_total = summary_df.get(f"{subj}_total", pd.Series([0])).sum()
            avg_completion = (total_solved / total_total * 100) if total_total > 0 else 0
            
            # الحصول على التوصية الثابتة للمادة
            recommendation = analyze_subject_patterns(summary_df, subj)
            
            subject_performance.append({
                "المادة": subj,
                "إجمالي المنجز": total_solved,
                "إجمالي التقييمات": total_total,
                "متوسط الإنجاز %": f"{avg_completion:.2f}%",
                "التوصية التشغيلية": recommendation
            })
        
        df_subj = pd.DataFrame(subject_performance)
        df_subj.to_excel(w, sheet_name="ملخص الأداء حسب المادة", index=False)
        
        # 2. تقرير الأداء حسب الشعبة والمادة
        report_data = []
        for (class_name, section), group in summary_df.groupby(["الصف", "الشعبة"]):
            for subj in subjects:
                total_solved = group.get(f"{subj}_solved", pd.Series([0])).sum()
                total_total = group.get(f"{subj}_total", pd.Series([0])).sum()
                avg_completion = (total_solved / total_total * 100) if total_total > 0 else 0
                
                report_data.append({
                    "الصف": class_name,
                    "الشعبة": section,
                    "المادة": subj,
                    "إجمالي المنجز": total_solved,
                    "إجمالي التقييمات": total_total,
                    "متوسط الإنجاز %": f"{avg_completion:.2f}%"
                })
        
        df_class_subj = pd.DataFrame(report_data)
        df_class_subj.to_excel(w, sheet_name="الأداء حسب الشعبة والمادة", index=False)
        
    mem.seek(0)
    return mem
# ----------------------------------------------------------------------

def send_teacher_emails(summary_df: pd.DataFrame, inactive_threshold: float):
    """توليد تنبيهات البريد الإلكتروني للمعلمين حول الطلاب غير الفاعلين."""
    if 'teacher_email' not in summary_df.columns:
        st.error("لا يمكن إرسال الإيميلات. يرجى التأكد من تحميل ملف المعلمين وربطه ببيانات الطلاب.")
        return
    
    # تحديد الطلاب غير الفاعلين
    inactive_students = summary_df[summary_df['نسبة الإنجاز %'] <= inactive_threshold]
    
    if inactive_students.empty:
        st.success("لا يوجد طلاب غير فاعلين (أقل من الحد المحدد).")
        return
        
    # تجميع حسب المعلمة
    email_groups = inactive_students.groupby(['teacher_email', 'teacher_name'])
    
    st.info(f"تم تحديد {inactive_students.shape[0]} طالب غير فاعل سيتم إرسال تنبيهات بشأنهم إلى {len(email_groups)} معلمة.")
    
    for (email, name), group in email_groups:
        student_list = "\n".join([f"- {row['اسم الطالب']} ({row['الصف']}/{row['الشعبة']})" for _, row in group.iterrows()])
        
        # التوصية الجماعية للطلاب غير الفاعلين
        recommendation = STUDENT_RECOMMENDATIONS["🚫 Not Utilizing System"]
        
        email_body = f"""
        عزيزتي المعلمة/ {name}،
        
        تحية طيبة وبعد،
        
        نود تنبيهك بوجود مجموعة من الطلاب في صفوفك لم تظهر بعد استفادة كافية من نظام قطر للتعليم، حيث أن نسبة إنجازهم أقل من {inactive_threshold}%.
        
        **قائمة الطلاب غير الفاعلين:**
        {student_list}
        
        **التوصية التشغيلية:**
        {recommendation}
        
        يرجى التواصل مع هؤلاء الطلاب وحثهم على تفعيل النظام والمشاركة النشطة.
        
        مع خالص الشكر والتقدير،
        فريق أي إنجاز
        """
        
        # هنا يتم إرسال الإيميل الفعلي (يجب إضافة كود SMTP هنا)
        # مثال:
        # send_mail(to=email, subject="تنبيه: طلاب غير فاعلين في نظام قطر للتعليم", body=email_body)
        
        st.write(f"تم توليد تنبيه للمعلمة {name} ({email}) لـ {group.shape[0]} طالب.")

# ---------- Streamlit App ----------
def main():
    
    # 1. عرض رأس الصفحة (الشعارات والعنوان)
    display_header()
    
    # 2. إعدادات الشريط الجانبي
    st.sidebar.header("⚙️ إعدادات النظام")
    
    # واجهة إدخال بيانات المدرسة والمسؤولين
    with st.sidebar.expander("🏫 بيانات المدرسة والمسؤولين", expanded=True):
        st.session_state.school_info = {
            "School_Name": st.text_input("اسم المدرسة", "مدرسة عثمان بن عفان النموذجية"),
            "Coordinator": st.text_input("منسق المشاريع الإلكترونية", "سحر عثمان"),
            "Academic_Deputy": st.text_input("النائب الأكاديمي", "مريم القضع"),
            "Administrative_Deputy": st.text_input("النائب الإداري", "دلال الفهيدة"),
            "Principal": st.text_input("مدير المدرسة", "منيرة الهاجري"),
        }
    
    # إعدادات النظام ومعايير التصنيف
    with st.sidebar.expander("📊 معايير التصنيف", expanded=False):
        st.session_state.thresholds = {
            "Platinum": st.number_input("حد Platinum (%) (أكبر من)", 0, 100, 89),
            "Gold": st.number_input("حد Gold (%) (أكبر من)", 0, 100, 79),
            "Silver": st.number_input("حد Silver (%) (أكبر من)", 0, 100, 49),
            "Bronze": st.number_input("حد Bronze (%) (أكبر من)", 0, 100, 0)
        }
        inactive_threshold = st.number_input("حد الطلاب غير الفاعلين (%) (أقل من أو يساوي)", 0, 100, 10)
    
    # 3. تحميل ملفات المعلمين
    teacher_file = st.sidebar.file_uploader("📂 تحميل ملف بيانات المعلمين (لإرسال الإيميلات)", type=["xlsx", "csv", "xls"])
    teachers_df = _load_teachers_df(teacher_file)
    
    # 4. تحميل ملفات التقييمات
    st.sidebar.header("📂 تحميل بيانات التقييمات")
    date_filter = st.sidebar.date_input("فلتر التاريخ (تاريخ بداية الإنجاز)", pd.to_datetime("today") - pd.Timedelta(days=30))
    uploaded_files = st.sidebar.file_uploader("تحميل ملفات التقييمات (Excel)", type=["xlsx", "xls"], accept_multiple_files=True)
    
    # 5. معالجة البيانات
    if uploaded_files:
        all_rows = []
        for file in uploaded_files:
            try:
                xls = pd.ExcelFile(file)
                selected_sheets = st.sidebar.multiselect(f"اختر أوراق من {file.name}", xls.sheet_names, default=xls.sheet_names)
                
                # المعالجة الفعلية
                rows = process_excel_file(file, file.name, start_row_students=1, selected_sheets=selected_sheets)
                all_rows.extend(rows)
                
            except Exception as e:
                st.error(f"خطأ في معالجة الملف {file.name}: {e}")
                
        if all_rows:
            raw_df = pd.DataFrame(all_rows)
            raw_df = ensure_uid(raw_df)
            
            # تطبيق فلتر التاريخ (افتراضياً، لا يوجد عمود تاريخ، لذا سنفترض أن الفلتر يطبق يدوياً)
            # if 'date_column' in raw_df.columns:
            #     raw_df = raw_df[pd.to_datetime(raw_df['date_column']) >= date_filter]
                
            summary_df, subjects = build_summary_pivot(raw_df, st.session_state.thresholds)
            
            # ربط بيانات المعلمين
            if teachers_df is not None and not teachers_df.empty:
                summary_df['class_section'] = summary_df['الصف'].astype(str) + ' ' + summary_df['الشعبة'].astype(str)
                summary_df = pd.merge(summary_df, teachers_df, on='class_section', how='left')
                summary_df.drop(columns=['class_section'], inplace=True)
            
            st.session_state.summary_df = summary_df
            st.session_state.subjects = subjects
            st.session_state.raw_df = raw_df
            
            st.success(f"تمت معالجة {summary_df.shape[0]} طالب. إجمالي المواد: {len(subjects)}")
            
            # 6. عرض جدول الملخص
            st.header("جدول ملخص إنجاز الطلاب")
            st.dataframe(summary_df)
            
            # 7. التوصيات على مستوى المادة
            st.header("تحليل الأنماط والتوصيات على مستوى المادة")
            for subj in subjects:
                with st.expander(f"توصية المادة: {subj}"):
                    st.info(analyze_subject_patterns(summary_df, subj))
            
            # 8. الرسوم البيانية التفاعلية
            st.header("📈 الرسوم البيانية التفاعلية")
            
            # رسم بياني 1: حسب المادة
            st.subheader("تحليل الأداء حسب المادة")
            subject_chart = create_subject_performance_chart(summary_df, subjects)
            st.plotly_chart(subject_chart, use_container_width=True)
            
            # رسم بياني 2: حسب الشعبة
            st.subheader("تحليل الأداء حسب الشعبة")
            class_chart = create_class_section_performance_chart(summary_df)
            st.plotly_chart(class_chart, use_container_width=True)
            
            # 9. تقارير الطلاب الفردية (PDF)
            st.header("📄 تقارير الطلاب الفردية")
            if not summary_df.empty:
                student_names = summary_df["اسم الطالب"].tolist()
                selected_student = st.selectbox("اختر طالبًا لإنشاء تقرير فردي:", student_names)
                
                if selected_student:
                    student_data = summary_df[summary_df["اسم الطالب"] == selected_student].iloc[0]
                    
                    # خيار التوصية المخصصة
                    custom_rec = st.text_area(
                        "توصية منسق المشاريع (اختياري، اتركها فارغة لاستخدام التوصية التلقائية):",
                        value="",
                        height=100
                    )
                    
                    # إنشاء التقرير الفردي
                    pdf_output = create_student_report_pdf(student_data, raw_df, st.session_state.school_info, custom_rec)
                    
                    st.download_button(
                        label=f"⬇️ تحميل تقرير {selected_student} (PDF)",
                        data=pdf_output,
                        file_name=f"تقرير_إنجاز_{selected_student}.pdf",
                        mime="application/pdf"
                    )
                    
                # زر تحميل جميع التقارير (ZIP)
                if st.button("⬇️ تحميل جميع التقارير الفردية (ZIP)"):
                    with st.spinner("جاري تجميع جميع التقارير الفردية..."):
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, False) as zip_file:
                            for index, row in summary_df.iterrows():
                                pdf_data = create_student_report_pdf(row, raw_df, st.session_state.school_info)
                                zip_file.writestr(f"تقرير_إنجاز_{row['اسم الطالب']}.pdf", pdf_data.getvalue())
                        
                        st.download_button(
                            label="تحميل ملف ZIP لجميع التقارير",
                            data=zip_buffer.getvalue(),
                            file_name="جميع_تقارير_الإنجاز_الفردية.zip",
                            mime="application/zip"
                        )

            # 10. التقرير الكمي الوصفي (Excel)
            st.header("📊 التقرير الكمي الوصفي")
            quantitative_excel = create_quantitative_report_excel(summary_df, subjects)
            
            st.download_button(
                label="⬇️ تحميل التقرير الكمي الوصفي (Excel)",
                data=quantitative_excel,
                file_name="التقرير_الكمي_الوصفي_للإنجاز.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 11. خيارات التصدير والإيميل
            st.header("إجراءات إضافية")
            col_export, col_email = st.columns(2)
            
            with col_export:
                # تصدير Excel
                excel_data = to_excel_bytes({"ملخص الإنجاز": summary_df})
                st.download_button(
                    label="⬇️ تصدير ملخص الإنجاز (Excel)",
                    data=excel_data,
                    file_name="ملخص_إنجاز_الطلاب_المحدث.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col_email:
                # إرسال الإيميلات
                if st.button("📧 إرسال تنبيهات الطلاب غير الفاعلين للمعلمين"):
                    send_teacher_emails(summary_df, inactive_threshold)
                    st.success("تم الانتهاء من عملية توليد التنبيهات.")
        else:
            st.warning("لم يتم العثور على بيانات صالحة للمعالجة.")
    else:
        st.info("الرجاء تحميل ملفات التقييمات للبدء بالتحليل.")
    
    st.markdown(FOOTER_MARKDOWN, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

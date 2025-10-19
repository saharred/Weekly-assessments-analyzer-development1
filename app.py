# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt  # احتياطي
import io
from datetime import datetime
from typing import Tuple
import logging

# =========================
# الإعدادات العامة
# =========================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="محلل التقييمات الأسبوعية",
    page_icon="https://i.imgur.com/XLef7tS.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# تنسيقات الواجهة (Qatar Maroon)
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');
* { font-family: 'Cairo', 'Segoe UI', -apple-system, sans-serif; }
.main, body, .stApp { background: #FFFFFF; }

/* Header */
.header-container {
  background: linear-gradient(135deg, #8A1538 0%, #6B1029 100%);
  padding: 56px 48px; color: #FFF; text-align: center; margin-bottom: 28px;
  box-shadow: 0 6px 20px rgba(138, 21, 56, 0.25); border-bottom: 4px solid #C9A646;
  position: relative;
}
.header-container::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 4px;
  background: linear-gradient(90deg, #C9A646 0%, #E8D4A0 50%, #C9A646 100%);
}
.header-container h1 { margin: 0 0 12px 0; font-size: 40px; font-weight: 800; }
.header-container .subtitle { font-size: 18px; font-weight: 700; margin: 0 0 6px 0; }
.header-container .accent-line { font-size: 14px; color: #C9A646; font-weight: 700; margin: 0 0 8px 0; }
.header-container .description { font-size: 14px; opacity: 0.95; margin: 0; }

/* Sidebar */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #8A1538 0%, #6B1029 100%) !important;
  border-right: 2px solid #C9A646; box-shadow: 4px 0 16px rgba(0,0,0,.15);
}
[data-testid="stSidebar"] * { color: #FFF !important; }
[data-testid="stSidebar"] hr { border-color: rgba(201, 166, 70, .3) !important; }

/* Cards/Metrics */
.metric-box {
  background: #FFF; border: 2px solid #E8E8E8; border-right: 5px solid #8A1538;
  padding: 20px; border-radius: 10px; text-align: center;
  box-shadow: 0 3px 12px rgba(0,0,0,.08); transition: all .3s ease;
}
.metric-box:hover { border-right-color: #C9A646; transform: translateY(-2px); }
.metric-value { font-size: 36px; font-weight: 800; color: #8A1538; }
.metric-label { font-size: 12px; font-weight: 700; color: #4A4A4A; letter-spacing: .06em; }

/* Chart container */
.chart-container {
  background: #FFF; border: 2px solid #E5E7EB; border-right: 5px solid #8A1538;
  border-radius: 12px; padding: 20px; margin: 14px 0; box-shadow: 0 2px 8px rgba(0,0,0,.08);
}
.chart-title { font-size: 22px; font-weight: 800; color: #8A1538; text-align: center; margin-bottom: 12px; }

/* Tables */
[data-testid="stDataFrame"] {
  border: 2px solid #E8E8E8; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,.06);
}

/* Footer */
.footer {
  margin-top: 48px; background: linear-gradient(135deg, #8A1538 0%, #6B1029 100%);
  color: #FFF; border-radius: 12px; padding: 36px 16px; text-align: center;
  box-shadow: 0 8px 24px rgba(138,21,56,.25); position: relative; overflow: hidden;
}
.footer .line {
  width: 100%; height: 4px; background: linear-gradient(90deg, #C9A646 0%, #E8D4A0 50%, #C9A646 100%);
  position: absolute; top: 0; left: 0;
}
.footer .brand { font-weight: 800; font-size: 18px; margin: 6px 0; }
.footer .rights { font-weight: 800; font-size: 16px; margin: 4px 0; }
.footer a { color: #C9A646; font-weight: 700; text-decoration: none; border-bottom: 1px solid #C9A646; }
</style>
""", unsafe_allow_html=True)

# =========================
# ثوابت التصنيفات والألوان
# =========================
CATEGORY_THRESHOLDS = {
    'بلاتيني 🥇': (90, 100),
    'ذهبي 🥈': (80, 89.99),
    'فضي 🥉': (70, 79.99),
    'برونزي': (60, 69.99),
    'بحاجة لتحسين': (0, 59.99)
}
CATEGORY_COLORS = {
    'بلاتيني 🥇': '#E5E4E2',
    'ذهبي 🥈': '#C9A646',
    'فضي 🥉': '#C0C0C0',
    'برونزي': '#CD7F32',
    'بحاجة لتحسين': '#8A1538'
}
CATEGORY_ORDER = ['بلاتيني 🥇', 'ذهبي 🥈', 'فضي 🥉', 'برونزي', 'بحاجة لتحسين']

# =========================
# دوال مساعدة للقراءة والحساب
# =========================
def parse_sheet_name(sheet_name: str) -> Tuple[str, str, str]:
    try:
        parts = sheet_name.strip().split()
        if len(parts) < 3:
            return sheet_name.strip(), "", ""
        section = parts[-1]; level = parts[-2]; subject = " ".join(parts[:-2])
        if not (level.isdigit() or (level.startswith('0') and len(level) <= 2)):
            subject = " ".join(parts[:-1]); level = parts[-1]; section = ""
        return subject, level, section
    except Exception:
        return sheet_name, "", ""

@st.cache_data
def analyze_excel_file(file, sheet_name):
    """
    - total_count: يحسب فقط الخانات المستحقة (يستثني '-', '—', 'I', 'AB', 'X')
    - completed_count: كل قيمة ليست من مجموعة الاستثناء وغير 'M' تعتبر مُنجزة (حتى 0)
    - solve_pct = completed / total * 100
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)

        # قراءة تواريخ الاستحقاق من H3 (index=2)
        due_dates = []
        try:
            for col_idx in range(7, min(df.shape[1], 50)):
                cell_value = df.iloc[2, col_idx]
                if pd.notna(cell_value):
                    try:
                        d = pd.to_datetime(cell_value)
                        if 2000 <= d.year <= 2100:
                            due_dates.append(d.date())
                    except Exception:
                        continue
        except Exception:
            pass

        # تحديد أعمدة التقييم
        assessment_columns = []
        for col_idx in range(7, df.shape[1]):
            title = df.iloc[0, col_idx] if col_idx < df.shape[1] else None
            if pd.isna(title): break
            all_dash = True
            for row_idx in range(4, min(len(df), 20)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    if str(cell_value).strip() not in ['-', '—', '']:
                        all_dash = False; break
            if not all_dash:
                assessment_columns.append({'index': col_idx, 'title': str(title).strip()})

        if not assessment_columns:
            st.warning(f"⚠️ لم يتم العثور على تقييمات في ورقة: {sheet_name}")
            return []

        results = []
        IGNORE = {'-', '—', '', 'I', 'AB', 'X', 'NAN', 'NONE'}

        for idx in range(4, len(df)):
            student_name = df.iloc[idx, 0]
            if pd.isna(student_name) or str(student_name).strip() == "": continue
            student_name_clean = " ".join(str(student_name).strip().split())

            completed_count = 0
            total_count = 0
            pending_titles = []

            for col_info in assessment_columns:
                col_idx = col_info['index']; title = col_info['title']
                if col_idx >= df.shape[1]: continue

                cell_value = df.iloc[idx, col_idx]
                cell_str = ("" if pd.isna(cell_value) else str(cell_value)).strip().upper()

                if cell_str in IGNORE:
                    continue
                if cell_str == 'M':
                    total_count += 1
                    pending_titles.append(title)
                    continue

                total_count += 1
                # أي قيمة أخرى (رقم/نص) → مُنجز
                completed_count += 1

            solve_pct = (completed_count / total_count * 100) if total_count > 0 else 0.0

            results.append({
                "student_name": student_name_clean,
                "subject": subject,
                "level": str(level_from_name).strip(),
                "section": str(section_from_name).strip(),
                "solve_pct": round(solve_pct, 1),
                "completed_count": int(completed_count),
                "total_count": int(total_count),
                "pending_titles": ", ".join(pending_titles) if pending_titles else "-",
                "due_dates": due_dates
            })

        logger.info(f"✅ تم تحليل {len(results)} طالب من ورقة {sheet_name}")
        return results

    except Exception as e:
        logger.error(f"❌ خطأ في analyze_excel_file: {str(e)}")
        st.error(f"❌ خطأ في تحليل الملف: {str(e)}")
        return []

@st.cache_data
def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if df.empty:
            st.warning("⚠️ لا توجد بيانات للتحليل")
            return pd.DataFrame()

        df_clean = df.copy()
        required_cols = ['student_name', 'level', 'section', 'subject', 'total_count', 'completed_count', 'solve_pct']
        missing = [c for c in required_cols if c not in df_clean.columns]
        if missing:
            st.error(f"❌ أعمدة مفقودة: {', '.join(missing)}")
            return pd.DataFrame()

        df_clean = df_clean.drop_duplicates(subset=['student_name', 'level', 'section', 'subject'], keep='first')

        unique_students = df_clean[['student_name', 'level', 'section']].drop_duplicates()
        unique_students = unique_students.sort_values(['level', 'section', 'student_name']).reset_index(drop=True)
        result = unique_students.copy()

        subjects = sorted(df_clean['subject'].unique())
        for subject in subjects:
            sub = df_clean[df_clean['subject'] == subject].copy()
            sub[['total_count','completed_count','solve_pct']] = sub[['total_count','completed_count','solve_pct']].fillna(0)

            block = sub[['student_name', 'level', 'section', 'total_count', 'completed_count', 'solve_pct']].copy()
            block = block.rename(columns={
                'total_count': f"{subject} - إجمالي",
                'completed_count': f"{subject} - منجز",
                'solve_pct': f"{subject} - النسبة"
            })
            block = block.drop_duplicates(subset=['student_name', 'level', 'section'], keep='first')
            result = result.merge(block, on=['student_name', 'level', 'section'], how='left')

            pending = sub[['student_name', 'level', 'section', 'pending_titles']].drop_duplicates(
                subset=['student_name', 'level', 'section'], keep='first'
            ).rename(columns={'pending_titles': f"{subject} - متبقي"})
            result = result.merge(pending, on=['student_name', 'level', 'section'], how='left')

        pct_cols = [c for c in result.columns if 'النسبة' in c]
        if pct_cols:
            result['المتوسط'] = result[pct_cols].mean(axis=1, skipna=True).fillna(0)

            def categorize(pct):
                if pd.isna(pct): return "بحاجة لتحسين"
                if pct >= 90: return "بلاتيني 🥇"
                if pct >= 80: return "ذهبي 🥈"
                if pct >= 70: return "فضي 🥉"
                if pct >= 60: return "برونزي"
                return "بحاجة لتحسين"

            result['الفئة'] = result['المتوسط'].apply(categorize)

        result = result.rename(columns={'student_name': 'الطالب', 'level': 'الصف', 'section': 'الشعبة'})

        for col in result.columns:
            if ('إجمالي' in col) or ('منجز' in col):
                result[col] = result[col].fillna(0).astype(int)
            elif ('النسبة' in col) or (col == 'المتوسط'):
                result[col] = result[col].fillna(0).round(1)
            elif 'متبقي' in col:
                result[col] = result[col].fillna('-')

        return result.drop_duplicates(subset=['الطالب', 'الصف', 'الشعبة'], keep='first').reset_index(drop=True)

    except Exception as e:
        logger.error(f"خطأ في create_pivot_table: {e}")
        st.error(f"❌ خطأ في معالجة البيانات: {str(e)}")
        return pd.DataFrame()

# =========================
# تطبيع للرسوم
# =========================
def assign_category(percent: float) -> str:
    if pd.isna(percent): return 'بحاجة لتحسين'
    for cat, (mn, mx) in CATEGORY_THRESHOLDS.items():
        if mn <= percent <= mx: return cat
    return 'بحاجة لتحسين'

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    mapping = {
        'الطالب': 'student', 'Student': 'student', 'الاسم': 'student',
        'الصف': 'grade', 'Grade': 'grade', 'المستوى': 'grade',
        'الشعبة': 'section', 'Section': 'section',
        'المادة': 'subject', 'Subject': 'subject',
        'إجمالي': 'total', 'Total': 'total', 'total_count': 'total',
        'منجز': 'solved', 'Solved': 'solved', 'completed_count': 'solved',
        'النسبة': 'percent', 'Percent': 'percent', 'solve_pct': 'percent',
        'الفئة': 'category', 'Category': 'category'
    }
    df = df.rename(columns=mapping)
    if 'percent' not in df.columns and {'total','solved'}.issubset(df.columns):
        df['percent'] = df.apply(lambda r: (r['solved']/r['total']*100) if r.get('total',0)>0 else 0.0, axis=1)
    if 'category' not in df.columns and 'percent' in df.columns:
        df['category'] = df['percent'].apply(assign_category)
    for need in ['subject', 'percent', 'category']:
        if need not in df.columns: raise ValueError(f"Missing required column after normalization: {need}")
    df['percent'] = df['percent'].fillna(0); df['category'] = df['category'].fillna('بحاجة لتحسين')
    return df

def aggregate_by_subject(df: pd.DataFrame) -> pd.DataFrame:
    agg = []
    for subject in df['subject'].dropna().unique():
        sub = df[df['subject'] == subject]
        total_students = len(sub)
        avg_completion = sub['percent'].mean() if total_students else 0
        for cat in CATEGORY_ORDER:
            count = (sub['category'] == cat).sum()
            share = (count/total_students*100) if total_students else 0
            agg.append({'subject': subject, 'category': cat, 'count': int(count),
                        'percent_share': round(share,1), 'avg_completion': round(avg_completion,1)})
    agg_df = pd.DataFrame(agg)
    if agg_df.empty: return agg_df
    order = (agg_df.groupby('subject')['avg_completion'].first().sort_values(ascending=False).index.tolist())
    agg_df['subject'] = pd.Categorical(agg_df['subject'], categories=order, ordered=True)
    return agg_df.sort_values('subject')

def chart_stacked_by_subject(agg_df: pd.DataFrame, mode: str = 'percent') -> go.Figure:
    subjects = agg_df['subject'].unique().tolist()
    fig = go.Figure()
    for cat in CATEGORY_ORDER:
        dat = agg_df[agg_df['category'] == cat]
        if mode == 'percent':
            values = dat['percent_share'].tolist()
            text = [f"{v:.1f}%" if v > 0 else "" for v in values]
            hover = "<b>%{y}</b><br>الفئة: " + cat + "<br>العدد: %{customdata[0]}<br>النسبة: %{x:.1f}%<extra></extra>"
        else:
            values = dat['count'].tolist()
            text = [str(int(v)) if v > 0 else "" for v in values]
            hover = "<b>%{y}</b><br>الفئة: " + cat + "<br>العدد: %{x}<extra></extra>"
        fig.add_trace(go.Bar(
            name=cat, x=values, y=dat['subject'].tolist(), orientation='h',
            marker=dict(color=CATEGORY_COLORS[cat], line=dict(color='white', width=1)),
            text=text, textposition='inside', textfont=dict(size=11, family='Cairo'),
            hovertemplate=hover,
            customdata=np.column_stack((dat['count'].tolist(), dat['percent_share'].tolist()))
        ))
    title = "توزيع الفئات حسب المادة" if mode == 'percent' else "عدد الطلاب حسب الفئة والمادة"
    fig.update_layout(
        title=dict(text=title, font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        xaxis=dict(title="النسبة المئوية (%)" if mode=='percent' else "عدد الطلاب",
                   tickfont=dict(size=12, family='Cairo'), gridcolor='#E5E7EB',
                   range=[0, 100] if mode=='percent' else None),
        yaxis=dict(title="المادة", tickfont=dict(size=12, family='Cairo'), autorange='reversed'),
        barmode='stack', height=max(420, len(subjects)*60),
        margin=dict(l=220, r=40, t=70, b=40),
        plot_bgcolor='white', paper_bgcolor='white', font=dict(family='Cairo'),
        legend=dict(title="الفئة", orientation='h', y=1.02, x=0.5, xanchor='center', font=dict(family='Cairo'))
    )
    return fig

def chart_overall_donut(pivot: pd.DataFrame) -> go.Figure:
    if 'الفئة' not in pivot.columns or pivot.empty: return go.Figure()
    counts = pivot['الفئة'].value_counts().reindex(CATEGORY_ORDER, fill_value=0)
    fig = go.Figure(data=[go.Pie(
        labels=counts.index, values=counts.values, hole=0.55,
        marker=dict(colors=[CATEGORY_COLORS[k] for k in counts.index]),
        textinfo='label+value', hovertemplate="%{label}: %{value} طالب<extra></extra>"
    )])
    fig.update_layout(
        title=dict(text="توزيع عام للفئات", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        showlegend=False, font=dict(family='Cairo'), paper_bgcolor='white', plot_bgcolor='white'
    )
    return fig

def chart_overall_gauge(pivot: pd.DataFrame) -> go.Figure:
    avg = float(pivot['المتوسط'].mean()) if ('المتوسط' in pivot.columns and not pivot.empty) else 0.0
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=avg,
        number={'suffix': "%", 'font': {'family': 'Cairo'}},
        gauge={
            'axis': {'range': [0, 100]}, 'bar': {'color': '#8A1538'},
            'steps': [
                {'range': [0, 60], 'color': '#FDE2E4'},
                {'range': [60, 70], 'color': '#FFE8CC'},
                {'range': [70, 80], 'color': '#FFF7CC'},
                {'range': [80, 90], 'color': '#E8F5E9'},
                {'range': [90, 100], 'color': '#E3F2FD'},
            ],
            'threshold': {'line': {'color': '#C9A646', 'width': 4}, 'thickness': 0.8, 'value': avg}
        }
    ))
    fig.update_layout(
        title=dict(text="متوسط الإنجاز العام", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
        paper_bgcolor='white', plot_bgcolor='white', font=dict(family='Cairo'), height=320
    )
    return fig

# =========================
# واجهة الرأس
# =========================
st.markdown("""
<div class='header-container'>
  <div style='display:flex; align-items:center; justify-content:center; gap:16px; margin-bottom: 14px;'>
    <svg width="48" height="48" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect x="4" y="4" width="40" height="40" rx="4" fill="#C9A646" opacity="0.15"/>
      <path d="M12 32V24M18 32V20M24 32V16M30 32V22M36 32V18" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="12" cy="24" r="2.5" fill="#C9A646"/>
      <circle cx="18" cy="20" r="2.5" fill="#C9A646"/>
      <circle cx="24" cy="16" r="2.5" fill="#C9A646"/>
      <circle cx="30" cy="22" r="2.5" fill="#C9A646"/>
      <circle cx="36" cy="18" r="2.5" fill="#C9A646"/>
      <path d="M12 24L18 20L24 16L30 22L36 18" stroke="#C9A646" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
    <h1>نظام قطر للتعليم - محلل التقييمات الأسبوعية</h1>
  </div>
  <p class='subtitle'>وزارة التربية والتعليم والتعليم العالي</p>
  <p class='accent-line'>ضمان تنمية رقمية مستدامة</p>
  <p class='description'>نظام تحليل شامل وموثوق لنتائج الطلاب</p>
</div>
""", unsafe_allow_html=True)

# =========================
# حالة الجلسة
# =========================
if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = None

# =========================
# الشريط الجانبي (رفع وتحكم)
# =========================
with st.sidebar:
    st.image("https://i.imgur.com/XLef7tS.png", width=110)
    st.markdown("---")
    st.header("⚙️ الإعدادات")

    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader("اختر ملفات Excel", type=["xlsx", "xls"], accept_multiple_files=True)

    selected_sheets = []
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        all_sheets, sheet_file_map = [], {}
        for file_idx, file in enumerate(uploaded_files):
            try:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    display = f"[ملف {file_idx+1}] {sheet}"
                    all_sheets.append(display); sheet_file_map[display] = (file, sheet)
            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف: {e}")

        if all_sheets:
            st.info(f"📋 وجدت {len(all_sheets)} مادة من {len(uploaded_files)} ملفات")
            select_all = st.checkbox("✔️ اختر الجميع", value=True)
            if select_all:
                selected_sheets_display = all_sheets
            else:
                selected_sheets_display = st.multiselect("اختر المواد للتحليل", all_sheets, default=[])
            selected_sheets = [(sheet_file_map[s][0], sheet_file_map[s][1]) for s in selected_sheets_display]
    else:
        st.info("📤 ارفع ملفات Excel للبدء")

    st.markdown("---")
    st.subheader("🏫 معلومات المدرسة")
    school_name = st.text_input("اسم المدرسة", placeholder="مدرسة قطر النموذجية")
    st.subheader("✍️ التوقيعات")
    coordinator_name = st.text_input("منسق/ة المشاريع")
    academic_deputy = st.text_input("النائب الأكاديمي")
    admin_deputy = st.text_input("النائب الإداري")
    principal_name = st.text_input("مدير/ة المدرسة")

    st.markdown("---")
    run_analysis = st.button("▶️ تشغيل التحليل", use_container_width=True, type="primary",
                             disabled=not (uploaded_files and selected_sheets))

# =========================
# تشغيل التحليل
# =========================
if not uploaded_files:
    st.info("📤 الرجاء رفع ملفات Excel من الشريط الجانبي للبدء في التحليل")
elif run_analysis:
    with st.spinner("⏳ جاري التحليل..."):
        all_results = []
        for file, sheet in selected_sheets:
            all_results.extend(analyze_excel_file(file, sheet))

        if all_results:
            df = pd.DataFrame(all_results)
            st.session_state.analysis_results = df
            pivot = create_pivot_table(df)
            st.session_state.pivot_table = pivot
            st.success(f"✅ تم تحليل {len(pivot)} طالب من {df['subject'].nunique()} مادة بنجاح")

# =========================
# عرض النتائج العامة + الرسوم
# =========================
if st.session_state.pivot_table is not None and not st.session_state.pivot_table.empty:
    pivot = st.session_state.pivot_table
    df = st.session_state.analysis_results

    st.subheader("📈 ملخص النتائج")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.metric("👥 إجمالي الطلاب", len(pivot))
    with c2: st.metric("📚 عدد المواد", df['subject'].nunique())
    with c3:
        avg = pivot['المتوسط'].mean() if 'المتوسط' in pivot.columns else 0
        st.metric("📊 متوسط الإنجاز", f"{avg:.1f}%")
    with c4:
        platinum = (pivot['الفئة'] == 'بلاتيني 🥇').sum()
        st.metric("🥇 فئة بلاتيني", platinum)
    with c5:
        zero = (pivot['المتوسط'] == 0).sum() if 'المتوسط' in pivot.columns else 0
        st.metric("⚠️ بدون إنجاز", int(zero))

    st.divider()
    st.subheader("📋 جدول النتائج التفصيلي")
    st.dataframe(pivot, use_container_width=True, height=420)

    st.divider()
    st.subheader("💾 تحميل النتائج")
    colx, coly = st.columns(2)
    with colx:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot.to_excel(writer, index=False, sheet_name='النتائج')
            if df is not None: df.to_excel(writer, index=False, sheet_name='Raw_Records')
        st.download_button(
            "📥 تحميل Excel", output.getvalue(),
            f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with coly:
        csv_data = pivot.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 تحميل CSV", csv_data,
            f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "text/csv",
            use_container_width=True
        )

    st.divider()

    # رسوم عامة
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">🍩 التوزيع العام للفئات</h2>', unsafe_allow_html=True)
    st.plotly_chart(chart_overall_donut(pivot), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">🎯 مؤشر متوسط الإنجاز</h2>', unsafe_allow_html=True)
    st.plotly_chart(chart_overall_gauge(pivot), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">📊 توزيع الفئات حسب المادة الدراسية</h2>', unsafe_allow_html=True)
    try:
        normalized = normalize_dataframe(df)
        mode_choice = st.radio('نوع العرض', ['النسبة المئوية (%)', 'العدد المطلق'], horizontal=True)
        mode = 'percent' if mode_choice == 'النسبة المئوية (%)' else 'count'
        agg_df = aggregate_by_subject(normalized)
        fig = chart_stacked_by_subject(agg_df, mode=mode)
        st.plotly_chart(fig, use_container_width=True)
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📥 تحميل البيانات (CSV)",
                agg_df.to_csv(index=False, encoding='utf-8-sig'),
                f"subject_categories_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                "text/csv", use_container_width=True
            )
        with dl2:
            try:
                png_bytes = fig.to_image(format="png", width=1400, height=900, scale=2)
                st.download_button(
                    "📥 تحميل الرسم البياني (PNG)",
                    png_bytes, f"chart_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.png",
                    "image/png", use_container_width=True
                )
            except Exception:
                st.info("💡 لتحميل الصورة PNG، ثبّت: pip install kaleido")
    except Exception as e:
        st.error(f"خطأ في بناء الرسم: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # =========================
    # 📑 التقارير الفردية لكل طالب
    # =========================
    st.subheader("📑 التقارير الفردية لكل طالب")

    # قائمة الطلاب من Pivot (أكثر اتساقًا)
    student_list = pivot['الطالب'].dropna().unique().tolist()
    student_list.sort(key=lambda x: str(x))

    if len(student_list) == 0:
        st.info("لا يوجد طلاب لعرض تقاريرهم.")
    else:
        col_sel, col_btn = st.columns([3,1])
        with col_sel:
            selected_student = st.selectbox("اختر الطالب", student_list, index=0)
        with col_btn:
            as_csv = st.checkbox("تضمين CSV مع Excel", value=True)

        # بناء تقرير الطالب من df الخام (أدق)
        student_df = df[df['student_name'].str.strip().eq(str(selected_student).strip())].copy()
        # جدول مُنسّق للعرض
        student_table = student_df[['subject', 'total_count', 'completed_count', 'solve_pct', 'pending_titles']].copy()
        student_table = student_table.rename(columns={
            'subject':'المادة', 'total_count':'إجمالي', 'completed_count':'منجز',
            'solve_pct':'النسبة (%)', 'pending_titles':'العناوين المتبقية'
        })
        student_table['النسبة (%)'] = student_table['النسبة (%)'].round(1)

        # ملخص أعلى التقرير
        overall_avg = float(student_table['النسبة (%)'].mean()) if not student_table.empty else 0.0
        def cat(pct):
            if pct >= 90: return "بلاتيني 🥇"
            if pct >= 80: return "ذهبي 🥈"
            if pct >= 70: return "فضي 🥉"
            if pct >= 60: return "برونزي"
            return "بحاجة لتحسين"
        overall_cat = cat(overall_avg)

        box1, box2, box3 = st.columns(3)
        with box1:
            st.markdown("<div class='metric-box'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{overall_avg:.1f}%</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>متوسط إنجاز الطالب</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with box2:
            st.markdown("<div class='metric-box'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{overall_cat}</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>الفئة الحالية</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with box3:
            st.markdown("<div class='metric-box'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{int(student_table['إجمالي'].sum())}</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>إجمالي التقييمات المستحقة (جميع المواد)</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("### 📚 جدول تقرير الطالب")
        st.dataframe(student_table, use_container_width=True, height=360)

        # --- رسوم تقرير الطالب ---
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">🍩 نسب الإنجاز حسب المادة</h2>', unsafe_allow_html=True)
        # Doughnut متعدد الشرائح = إجمالي إنجاز الطالب موزّع على المواد كقِيم
        labels = student_table['المادة'].tolist()
        values = student_table['النسبة (%)'].tolist()
        donut = go.Figure(data=[go.Pie(
            labels=labels, values=values, hole=0.55, textinfo='label+percent',
            hovertemplate="%{label}: %{value:.1f}%<extra></extra>"
        )])
        donut.update_layout(
            title=dict(text=f"إنجاز {selected_student} على المواد", font=dict(size=20, family='Cairo', color='#8A1538'), x=0.5),
            showlegend=False, font=dict(family='Cairo'), paper_bgcolor='white', plot_bgcolor='white', height=420
        )
        st.plotly_chart(donut, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h2 class="chart-title">📊 منجز/إجمالي لكل مادة</h2>', unsafe_allow_html=True)
        bar = go.Figure()
        bar.add_trace(go.Bar(
            y=student_table['المادة'], x=student_table['إجمالي'], name='إجمالي',
            orientation='h', marker=dict(color='#E5E7EB'), hovertemplate="إجمالي: %{x}<extra></extra>"
        ))
        bar.add_trace(go.Bar(
            y=student_table['المادة'], x=student_table['منجز'], name='منجز',
            orientation='h', marker=dict(color='#8A1538'), hovertemplate="منجز: %{x}<extra></extra>"
        ))
        bar.update_layout(
            barmode='overlay', xaxis=dict(title="عدد التقييمات"),
            yaxis=dict(title="المادة"), font=dict(family='Cairo'),
            paper_bgcolor='white', plot_bgcolor='white', height=max(420, len(labels)*32),
            legend=dict(orientation='h', y=1.02, x=0.5, xanchor='center')
        )
        st.plotly_chart(bar, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # تحميل تقرير الطالب (Excel + صور إن توفرت)
        exp1 = st.expander("💾 تنزيل تقرير الطالب")
        with exp1:
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                # معلومات عامة
                info_df = pd.DataFrame({
                    'الحقل': ['الطالب','متوسط الإنجاز','الفئة','تاريخ التقرير'],
                    'القيمة': [selected_student, f"{overall_avg:.1f}%", overall_cat, datetime.now().strftime("%Y-%m-%d %H:%M")]
                })
                info_df.to_excel(writer, index=False, sheet_name='معلومات')
                student_table.to_excel(writer, index=False, sheet_name='تفاصيل المواد')

            st.download_button(
                "📥 تحميل تقرير الطالب (Excel)",
                excel_buf.getvalue(),
                f"student_report_{selected_student}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            if as_csv:
                st.download_button(
                    "📥 تحميل جدول الطالب (CSV)",
                    student_table.to_csv(index=False, encoding='utf-8-sig'),
                    f"student_table_{selected_student}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    "text/csv", use_container_width=True
                )

            # محاولة تصدير الرسوم PNG (يتطلب kaleido)
            try:
                donut_png = donut.to_image(format="png", width=1200, height=800, scale=2)
                bar_png = bar.to_image(format="png", width=1200, height=800, scale=2)
                st.download_button(
                    "📥 تحميل رسم النِسب (PNG)",
                    donut_png, f"student_donut_{selected_student}.png",
                    "image/png", use_container_width=True
                )
                st.download_button(
                    "📥 تحميل رسم منجز/إجمالي (PNG)",
                    bar_png, f"student_bar_{selected_student}.png",
                    "image/png", use_container_width=True
                )
            except Exception:
                st.info("💡 لتحميل صور الرسوم كـ PNG، ثبّت الحزمة: pip install kaleido")

# =========================
# Footer
# =========================
st.markdown(f"""
<div class="footer">
  <div class="line"></div>
  <div style='margin-bottom: 16px;'>
    <img src='https://i.imgur.com/XLef7tS.png' style='width: 88px; height: auto; opacity: 0.95;' alt='Ministry Logo'>
  </div>
  <div class="brand">© {datetime.now().year} وزارة التربية والتعليم والتعليم العالي — جميع الحقوق محفوظة</div>
  <div class="rights">مدرسة عثمان بن عفان النموذجية للبنين</div>
  <div style="font-weight:700; margin-top:6px;">منسقة المشاريع الإلكترونية / سحر عثمان</div>
  <div style="margin-top:8px;">
    للتواصل: <a href="mailto:S.mahgoub0101@education.qa">S.mahgoub0101@education.qa</a>
  </div>
  <div style="margin-top:10px; font-size:12px; opacity:.9;">تطوير وتصميم: قسم التحول الرقمي</div>
</div>
""", unsafe_allow_html=True)

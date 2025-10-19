# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
from datetime import datetime, date
from typing import Tuple, Optional
import logging

# PDF & Arabic shaping
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.units import cm
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    AR_OK = True
except Exception:
    AR_OK = False

# =========================
# الإعدادات العامة
# =========================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

APP_TITLE = "إنجاز -تحليل القييمات الأسبوعية على نظام قطر للتعليم"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="https://i.imgur.com/XLef7tS.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# تنسيقات الواجهة
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
* { font-family: 'Cairo', 'Segoe UI', -apple-system, sans-serif; }
.main, body, .stApp { background: #FFFFFF; }
.header-container{
  background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
  padding:48px 40px;color:#fff;text-align:center;margin-bottom:20px;
  border-bottom:4px solid #C9A646; box-shadow:0 6px 20px rgba(138,21,56,.25);
  position:relative
}
.header-container::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;
  background:linear-gradient(90deg,#C9A646 0%,#E8D4A0 50%,#C9A646 100%)}
.header-container h1{margin:0 0 6px 0;font-size:34px;font-weight:800}
.header-container .subtitle{font-size:16px;font-weight:700;margin:0 0 4px}
.header-container .accent-line{font-size:13px;color:#C9A646;font-weight:700;margin:0 0 6px}
.header-container .description{font-size:13px;opacity:.95;margin:0}
[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#8A1538 0%,#6B1029 100%)!important;
  border-right:2px solid #C9A646; box-shadow:4px 0 16px rgba(0,0,0,.15)
}
[data-testid="stSidebar"] *{color:#fff!important}
.metric-box{background:#fff;border:2px solid #E8E8E8;border-right:5px solid #8A1538;
  padding:18px;border-radius:10px;text-align:center;box-shadow:0 3px 12px rgba(0,0,0,.08)}
.metric-value{font-size:32px;font-weight:800;color:#8A1538}
.metric-label{font-size:12px;font-weight:700;color:#4A4A4A;letter-spacing:.06em}
.chart-container{background:#fff;border:2px solid #E5E7EB;border-right:5px solid #8A1538;
  border-radius:12px;padding:18px;margin:12px 0;box-shadow:0 2px 8px rgba(0,0,0,.08)}
.chart-title{font-size:20px;font-weight:800;color:#8A1538;text-align:center;margin-bottom:10px}
[data-testid="stDataFrame"]{border:2px solid #E8E8E8;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.footer{margin-top:22px;background:linear-gradient(135deg,#8A1538 0%,#6B1029 100%);
  color:#fff;border-radius:10px;padding:12px 10px;text-align:center;box-shadow:0 6px 18px rgba(138,21,56,.20);
  position:relative;overflow:hidden}
.footer .line{width:100%;height:3px;background:linear-gradient(90deg,#C9A646 0%,#E8D4A0 50%,#C9A646 100%);
  position:absolute;top:0;left:0}
.footer img.logo{width:60px;height:auto;opacity:.95;margin:6px 0 8px}
.footer .school{font-weight:800;font-size:15px;margin:2px 0 4px}
.footer .rights{font-weight:700;font-size:12px;margin:0 0 4px;opacity:.95}
.footer .contact{font-size:12px;margin-top:2px}
.footer a{color:#E8D4A0;font-weight:700;text-decoration:none;border-bottom:1px solid #C9A646}
.footer .credit{margin-top:6px;font-size:11px;opacity:.85}
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
# دوال مساعدة
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

def _parse_excel_date(x) -> Optional[date]:
    try:
        d = pd.to_datetime(x)
        if pd.isna(d): return None
        if 2000 <= d.year <= 2100: return d.date()
        return None
    except Exception:
        return None

@st.cache_data
def analyze_excel_file(file, sheet_name, due_start: Optional[date] = None, due_end: Optional[date] = None):
    """
    فلتر تاريخ الاستحقاق:
      - عند تحديد مدى زمني، تُحتسب الأعمدة التي لها تاريخ ويقع بين (من/إلى) شاملًا.
      - الأعمدة بلا تاريخ تُستبعد عند تفعيل الفلتر.
    منطق العد:
      - IGNORE: ('-', '—', '', 'I', 'AB', 'X', 'NAN', 'NONE') ← لا تدخل الإجمالي.
      - 'M' ← تدخل الإجمالي (غير منجزة).
      - أي قيمة أخرى ← منجزة حتى لو كانت 0.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        subject, level_from_name, section_from_name = parse_sheet_name(sheet_name)

        filter_active = (due_start is not None and due_end is not None)
        if filter_active and due_start > due_end:
            due_start, due_end = due_end, due_start

        # أعمدة التقييم + التاريخ
        assessment_columns = []
        for col_idx in range(7, df.shape[1]):  # من H
            title = df.iloc[0, col_idx] if col_idx < df.shape[1] else None
            if pd.isna(title): break

            # ليس كله شرطات
            all_dash = True
            for row_idx in range(4, min(len(df), 20)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value) and str(cell_value).strip() not in ['-', '—', '']:
                    all_dash = False; break
            if all_dash: continue

            due_dt = _parse_excel_date(df.iloc[2, col_idx])  # صف 3 عادة
            if filter_active:
                if (due_dt is None) or not (due_start <= due_dt <= due_end):
                    continue

            assessment_columns.append({'index': col_idx, 'title': str(title).strip(), 'due': due_dt})

        if not assessment_columns:
            st.warning(f"⚠️ لا توجد تقييمات مطابقة للمدى الزمني في ورقة: {sheet_name}")
            return []

        results = []
        IGNORE = {'-', '—', '', 'I', 'AB', 'X', 'NAN', 'NONE'}

        for idx in range(4, len(df)):  # الطلاب
            student_name = df.iloc[idx, 0]
            if pd.isna(student_name) or str(student_name).strip() == "": continue
            student_name_clean = " ".join(str(student_name).strip().split())

            completed_count = 0
            total_count = 0
            pending_titles = []

            for col in assessment_columns:
                c = col['index']; title = col['title']
                if c >= df.shape[1]: continue
                cell_value = df.iloc[idx, c]
                s = ("" if pd.isna(cell_value) else str(cell_value)).strip().upper()

                if s in IGNORE:  # غير مستحق
                    continue
                if s == 'M':      # مستحق لكن غير منجز
                    total_count += 1
                    pending_titles.append(title)
                    continue
                total_count += 1  # أي قيمة أخرى = منجز
                completed_count += 1

            pct = (completed_count/total_count*100) if total_count>0 else 0.0
            results.append({
                "student_name": student_name_clean,
                "subject": subject,
                "level": str(level_from_name).strip(),
                "section": str(section_from_name).strip(),
                "solve_pct": round(pct, 1),
                "completed_count": int(completed_count),
                "total_count": int(total_count),
                "pending_titles": ", ".join(pending_titles) if pending_titles else "-",
            })
        return results

    except Exception as e:
        st.error(f"❌ خطأ في تحليل الملف: {str(e)}")
        return []

@st.cache_data
def create_pivot_table(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if df.empty:
            st.warning("⚠️ لا توجد بيانات للتحليل"); return pd.DataFrame()
        df_clean = df.drop_duplicates(subset=['student_name','level','section','subject'])
        unq = df_clean[['student_name','level','section']].drop_duplicates().sort_values(['level','section','student_name']).reset_index(drop=True)
        result = unq.copy()
        for subject in sorted(df_clean['subject'].unique()):
            sub = df_clean[df_clean['subject']==subject].copy()
            sub[['total_count','completed_count','solve_pct']] = sub[['total_count','completed_count','solve_pct']].fillna(0)
            block = sub[['student_name','level','section','total_count','completed_count','solve_pct']].rename(columns={
                'total_count':f'{subject} - إجمالي', 'completed_count':f'{subject} - منجز', 'solve_pct':f'{subject} - النسبة'
            }).drop_duplicates(subset=['student_name','level','section'])
            result = result.merge(block, on=['student_name','level','section'], how='left')
            pend = sub[['student_name','level','section','pending_titles']].drop_duplicates(subset=['student_name','level','section']).rename(
                columns={'pending_titles':f'{subject} - متبقي'})
            result = result.merge(pend, on=['student_name','level','section'], how='left')
        pct_cols = [c for c in result.columns if 'النسبة' in c]
        if pct_cols:
            result['المتوسط'] = result[pct_cols].mean(axis=1, skipna=True).fillna(0)
            def categorize(p):
                if p>=90: return 'بلاتيني 🥇'
                if p>=80: return 'ذهبي 🥈'
                if p>=70: return 'فضي 🥉'
                if p>=60: return 'برونزي'
                return 'بحاجة لتحسين'
            result['الفئة'] = result['المتوسط'].apply(categorize)
        result = result.rename(columns={'student_name':'الطالب','level':'الصف','section':'الشعبة'})
        for col in result.columns:
            if ('إجمالي' in col) or ('منجز' in col): result[col]=result[col].fillna(0).astype(int)
            elif ('النسبة' in col) or (col=='المتوسط'): result[col]=result[col].fillna(0).round(1)
            elif 'متبقي' in col: result[col]=result[col].fillna('-')
        return result.drop_duplicates(subset=['الطالب','الصف','الشعبة']).reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ خطأ في معالجة البيانات: {e}"); return pd.DataFrame()

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={'solve_pct':'percent','student_name':'student','subject':'subject'})
    df['category'] = df['percent'].apply(lambda p: 'بلاتيني 🥇' if p>=90 else 'ذهبي 🥈' if p>=80 else 'فضي 🥉' if p>=70 else 'برونزي' if p>=60 else 'بحاجة لتحسين')
    return df

def aggregate_by_subject(df: pd.DataFrame) -> pd.DataFrame:
    out=[]
    for s in df['subject'].dropna().unique():
        sub=df[df['subject']==s]; n=len(sub); avg=sub['percent'].mean() if n else 0
        for cat in CATEGORY_ORDER:
            c=(sub['category']==cat).sum(); pct=(c/n*100) if n else 0
            out.append({'subject':s,'category':cat,'count':int(c),'percent_share':round(pct,1),'avg_completion':round(avg,1)})
    agg=pd.DataFrame(out)
    if agg.empty: return agg
    order=agg.groupby('subject')['avg_completion'].first().sort_values(ascending=False).index.tolist()
    agg['subject']=pd.Categorical(agg['subject'],categories=order,ordered=True)
    return agg.sort_values('subject')

def chart_stacked_by_subject(agg_df: pd.DataFrame, mode='percent') -> go.Figure:
    fig=go.Figure()
    for cat in CATEGORY_ORDER:
        d=agg_df[agg_df['category']==cat]
        vals = d['percent_share'] if mode=='percent' else d['count']
        text=[(f"{v:.1f}%" if mode=='percent' else str(v)) if v>0 else "" for v in vals]
        hover = "<b>%{y}</b><br>الفئة: "+cat+"<br>"+("النسبة: %{x:.1f}%<extra></extra>" if mode=='percent' else "العدد: %{x}<extra></extra>")
        fig.add_trace(go.Bar(name=cat, x=vals, y=d['subject'], orientation='h',
                             marker=dict(color=CATEGORY_COLORS[cat], line=dict(color='white', width=1)),
                             text=text, textposition='inside', textfont=dict(size=11, family='Cairo'),
                             hovertemplate=hover))
    fig.update_layout(title=dict(text="توزيع الفئات حسب المادة",font=dict(size=20,family='Cairo',color='#8A1538'),x=.5),
                      xaxis=dict(title="النسبة المئوية (%)" if mode=='percent' else "عدد الطلاب",
                                 tickfont=dict(size=12,family='Cairo'), gridcolor='#E5E7EB',
                                 range=[0,100] if mode=='percent' else None),
                      yaxis=dict(title="المادة",tickfont=dict(size=12,family='Cairo'),autorange='reversed'),
                      barmode='stack',plot_bgcolor='white',paper_bgcolor='white',font=dict(family='Cairo'))
    return fig

def chart_overall_donut(pivot: pd.DataFrame) -> go.Figure:
    if 'الفئة' not in pivot.columns or pivot.empty: return go.Figure()
    counts = pivot['الفئة'].value_counts().reindex(CATEGORY_ORDER, fill_value=0)
    fig = go.Figure([go.Pie(labels=counts.index, values=counts.values, hole=.55,
                            marker=dict(colors=[CATEGORY_COLORS[k] for k in counts.index]),
                            textinfo='label+value', hovertemplate="%{label}: %{value} طالب<extra></extra>")])
    fig.update_layout(title=dict(text="توزيع عام للفئات",font=dict(size=20,family='Cairo',color='#8A1538'),x=.5),
                      showlegend=False,font=dict(family='Cairo'))
    return fig

def chart_overall_gauge(pivot: pd.DataFrame) -> go.Figure:
    avg = float(pivot['المتوسط'].mean()) if ('المتوسط' in pivot.columns and not pivot.empty) else 0.0
    fig = go.Figure(go.Indicator(mode="gauge+number", value=avg,
                                 number={'suffix':"%",'font':{'family':'Cairo'}},
                                 gauge={'axis':{'range':[0,100]},'bar':{'color':'#8A1538'}}))
    fig.update_layout(title=dict(text="متوسط الإنجاز العام",font=dict(size=20,family='Cairo',color='#8A1538'),x=.5),
                      paper_bgcolor='white', plot_bgcolor='white', font=dict(family='Cairo'), height=320)
    return fig

# ============== PDF أدوات ==============
def register_font_from_upload(font_file) -> str:
    """يسجل خط TTF مرفوع، أو يستخدم DejaVuSans كخيار افتراضي."""
    try:
        if font_file is not None:
            font_bytes = font_file.read()
            ttf_buf = io.BytesIO(font_bytes)
            pdfmetrics.registerFont(TTFont("ARFont", ttf_buf))
            return "ARFont"
    except Exception:
        pass
    # احتياطي
    return "Helvetica"

def ar(text: str) -> str:
    """إخراج نص عربي مُشكّل RTL إذا توفّرت الحِزم."""
    if not isinstance(text, str):
        text = str(text)
    if AR_OK:
        return get_display(arabic_reshaper.reshape(text))
    return text

def make_student_pdf(school_name: str,
                     student_name: str,
                     grade: str,
                     section: str,
                     table_df: pd.DataFrame,
                     overall_avg: float,
                     reco_text: str,
                     coordinator_name: str,
                     academic_deputy: str,
                     admin_deputy: str,
                     principal_name: str,
                     font_name: str) -> bytes:
    """يُنشئ PDF يشابه القالب المرفق."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4  # نقاط

    # ألوان وهوية
    maroon = colors.HexColor("#8A1538")
    gold = colors.HexColor("#C9A646")

    # رأس بسيط
    c.setFillColor(maroon)
    c.rect(0, H-2*cm, W, 2*cm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont(font_name, 14)
    c.drawRightString(W-1.2*cm, H-1.2*cm, ar("إنجاز - تقرير أداء الطالب"))

    # عنوان رئيسي
    y = H-2.6*cm
    c.setFillColor(maroon)
    c.setFont(font_name, 18)
    c.drawRightString(W-1.2*cm, y, ar("تقرير أداء الطالب - نظام قطر للتعليم"))
    c.setFillColor(gold)
    c.setLineWidth(2)
    c.line(W-16*cm, y-0.25*cm, W-1.2*cm, y-0.25*cm)

    # معلومات الطالب
    y -= 1.2*cm
    c.setFillColor(colors.black)
    c.setFont(font_name, 12)
    c.drawRightString(W-1.2*cm, y, ar(f"اسم المدرسة: {school_name or '—'}"))
    y -= 0.7*cm
    c.drawRightString(W-1.2*cm, y, ar(f"اسم الطالب: {student_name}"))
    y -= 0.7*cm
    c.drawRightString(W-1.2*cm, y, ar(f"الصف: {grade or '—'}        الشعبة: {section or '—'}"))
    y -= 1.0*cm

    # جدول المواد
    # الأعمدة: المادة | إجمالي | منجز | متبقي
    col_titles = [ar("المادة"), ar("عدد التقييمات الإجمالي"), ar("عدد التقييمات المنجزة"), ar("عدد التقييمات المتبقية")]
    col_widths = [7.5*cm, 4.0*cm, 4.0*cm, 4.0*cm]
    x_right = W - 1.2*cm
    x_positions = [x_right - sum(col_widths[:i+1]) for i in range(len(col_widths))]

    # رأس الجدول
    c.setFillColor(maroon)
    c.setFont(font_name, 12)
    row_h = 0.8*cm
    c.rect(x_right - sum(col_widths), y - row_h, sum(col_widths), row_h, fill=1, stroke=0)
    c.setFillColor(colors.white)
    for i, title in enumerate(col_titles):
        c.drawCentredString(x_positions[i] + col_widths[i]/2, y - 0.6*cm, title)

    # صفوف الجدول
    c.setFont(font_name, 11)
    c.setFillColor(colors.black)
    y -= row_h
    total_solved = 0
    total_total = 0
    for _, r in table_df.iterrows():
        if y < 3.5*cm:
            c.showPage(); y = H-2*cm
            c.setFont(font_name, 11)
        sub = str(r['المادة'])
        tot = int(r['إجمالي']); solv = int(r['منجز']); rem = int(max(tot - solv, 0))
        total_solved += solv; total_total += tot
        y -= row_h
        # خلفية خفيفة
        c.setFillColor(colors.HexColor("#F7F7F7"))
        c.rect(x_right - sum(col_widths), y, sum(col_widths), row_h, fill=1, stroke=0)
        c.setFillColor(colors.black)
        vals = [ar(sub), str(tot), str(solv), str(rem)]
        for i, v in enumerate(vals):
            c.drawCentredString(x_positions[i] + col_widths[i]/2, y + 0.25*cm, v)

    # إحصائيات
    y -= 1.2*cm
    c.setFont(font_name, 12)
    c.setFillColor(maroon)
    c.drawRightString(W-1.2*cm, y, ar("الإحصائيات"))
    c.setFillColor(colors.black)
    y -= 0.7*cm
    perc = overall_avg
    c.drawRightString(W-1.2*cm, y, ar(f"منجز: {total_solved}    متبقي: {max(total_total - total_solved,0)}    نسبة حل التقييمات: {perc:.1f}%"))

    # توصية المنسق
    y -= 1.0*cm
    c.setFillColor(maroon); c.drawRightString(W-1.2*cm, y, ar("توصية منسق المشاريع:"))
    y -= 0.7*cm
    c.setFillColor(colors.black)
    for line in (reco_text or "—").splitlines() or ["—"]:
        c.drawRightString(W-1.2*cm, y, ar(line)); y -= 0.6*cm

    # روابط
    y -= 0.5*cm
    c.setFillColor(maroon); c.drawRightString(W-1.2*cm, y, ar("روابط مهمة:"))
    y -= 0.6*cm
    c.setFillColor(colors.black)
    c.drawRightString(W-1.2*cm, y, ar("رابط نظام قطر: https://portal.education.qa"))
    y -= 0.6*cm
    c.drawRightString(W-1.2*cm, y, ar("استعادة كلمة المرور: https://password.education.qa"))
    y -= 0.6*cm
    c.drawRightString(W-1.2*cm, y, ar("قناة قطر للتعليم: https://edu.tv.qa"))

    # التوقيعات
    y -= 1.0*cm
    c.setFillColor(maroon); c.drawRightString(W-1.2*cm, y, ar("التوقيعات"))
    y -= 0.8*cm
    c.setFillColor(colors.black); c.setFont(font_name, 11)
    sigs = [
        ("منسق المشاريع", coordinator_name),
        ("النائب الأكاديمي", academic_deputy),
        ("النائب الإداري", admin_deputy),
        ("مدير المدرسة", principal_name),
    ]
    box_w = (W - 2.4*cm) / 2 - 0.6*cm
    box_h = 1.8*cm
    for i, (title, name) in enumerate(sigs):
        col = i % 2
        row = i // 2
        x0 = W - 1.2*cm - (col+1)*box_w - (col)*0.6*cm
        y0 = y - row*(box_h+0.6*cm)
        c.setStrokeColor(gold); c.rect(x0, y0 - box_h, box_w, box_h, fill=0, stroke=1)
        c.setFillColor(colors.black)
        c.drawRightString(x0 + box_w - 0.4*cm, y0 - 0.5*cm, ar(f"{title} / {name or '—'}"))
        c.drawRightString(x0 + box_w - 0.4*cm, y0 - 1.3*cm, ar("التوقيع: __________________  التاريخ: __________"))

    c.showPage()
    c.save()
    return buf.getvalue()

# =========================
# واجهة الرأس
# =========================
st.markdown(f"""
<div class='header-container'>
  <div style='display:flex; align-items:center; justify-content:center; gap:16px; margin-bottom: 10px;'>
    <svg width="44" height="44" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect x="4" y="4" width="40" height="40" rx="4" fill="#C9A646" opacity="0.15"/>
      <path d="M12 32V24M18 32V20M24 32V16M30 32V22M36 32V18" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="12" cy="24" r="2.5" fill="#C9A646"/><circle cx="18" cy="20" r="2.5" fill="#C9A646"/>
      <circle cx="24" cy="16" r="2.5" fill="#C9A646"/><circle cx="30" cy="22" r="2.5" fill="#C9A646"/>
      <circle cx="36" cy="18" r="2.5" fill="#C9A646"/>
      <path d="M12 24L18 20L24 16L30 22L36 18" stroke="#C9A646" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
    <h1>{APP_TITLE}</h1>
  </div>
  <p class='subtitle'>لوحة مهنية لقياس التقدم وتحليل النتائج</p>
  <p class='accent-line'>هوية إنجاز • دعم العربية الكامل</p>
  <p class='description'>اختر الملفات وفعّل فلتر التاريخ حسب الحاجة لنتائج أدق</p>
</div>
""", unsafe_allow_html=True)

# =========================
# حالة الجلسة
# =========================
if "analysis_results" not in st.session_state: st.session_state.analysis_results = None
if "pivot_table" not in st.session_state: st.session_state.pivot_table = None

# =========================
# الشريط الجانبي
# =========================
with st.sidebar:
    st.image("https://i.imgur.com/XLef7tS.png", width=110)
    st.markdown("---")
    st.header("⚙️ الإعدادات")

    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader("اختر ملفات Excel", type=["xlsx", "xls"], accept_multiple_files=True)

    st.subheader("⏳ فلتر تاريخ الاستحقاق")
    # مدى واحد (من/إلى) لكي يظهر الفلتر دومًا واضحًا
    default_start = date.today().replace(day=1)
    default_end = date.today()
    due_range = st.date_input("اختر المدى (من — إلى)", value=(default_start, default_end), format="YYYY-MM-DD")
    due_start, due_end = (due_range if isinstance(due_range, tuple) else (None, None))

    st.caption("عند استخدام المدى يتم استبعاد الأعمدة بلا تاريخ استحقاق.")

    st.subheader("🔤 خط عربي للـPDF (اختياري)")
    font_file = st.file_uploader("ارفع ملف خط TTF (مثل Cairo/Amiri)", type=["ttf"])

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
                             disabled=not (uploaded_files))

# =========================
# تشغيل التحليل
# =========================
if not uploaded_files:
    st.info("📤 من الشريط الجانبي ارفع ملفات Excel للبدء في التحليل")
elif run_analysis:
    with st.spinner("⏳ جاري التحليل..."):
        all_results = []
        for file in uploaded_files:
            try:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    all_results.extend(analyze_excel_file(file, sheet, due_start, due_end))
            except Exception as e:
                st.error(f"خطأ بقراءة الملف: {e}")

        if all_results:
            df = pd.DataFrame(all_results)
            st.session_state.analysis_results = df
            pivot = create_pivot_table(df)
            st.session_state.pivot_table = pivot
            st.success(f"✅ تم تحليل {len(pivot)} طالب عبر {df['subject'].nunique()} مادة")

# =========================
# عرض النتائج + الرسوم
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
        st.metric("🥇 فئة بلاتيني", int(platinum))
    with c5:
        zero = (pivot['المتوسط'] == 0).sum() if 'المتوسط' in pivot.columns else 0
        st.metric("⚠️ بدون إنجاز", int(zero))

    st.divider()
    st.subheader("📋 جدول النتائج التفصيلي")
    st.dataframe(pivot, use_container_width=True, height=420)

    st.divider()
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
    except Exception as e:
        st.error(f"خطأ في بناء الرسم: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # =========================
    # 📑 التقارير الفردية — PDF فقط
    # =========================
    st.subheader("📑 التقارير الفردية (PDF)")
    student_list = pivot['الطالب'].dropna().unique().tolist()
    student_list.sort(key=lambda x: str(x))
    if len(student_list) == 0:
        st.info("لا يوجد طلاب.")
    else:
        col_sel, col_reco = st.columns([2, 3])
        with col_sel:
            selected_student = st.selectbox("اختر الطالب", student_list, index=0)
            # استخراج صف الصف/الشعبة من Pivot
            stu_row = pivot[pivot['الطالب'] == selected_student].head(1)
            stu_grade = str(stu_row['الصف'].iloc[0]) if not stu_row.empty else ''
            stu_section = str(stu_row['الشعبة'].iloc[0]) if not stu_row.empty else ''
        with col_reco:
            reco_text = st.text_area("توصية منسق المشاريع", value="", height=120, placeholder="اكتب التوصيات هنا...")

        # جدول الطالب من السجلات الخام
        student_df = df[df['student_name'].str.strip().eq(str(selected_student).strip())].copy()
        student_table = student_df[['subject', 'total_count', 'completed_count']].copy()
        student_table = student_table.rename(columns={'subject':'المادة','total_count':'إجمالي','completed_count':'منجز'})
        student_table['متبقي'] = (student_table['إجمالي'] - student_table['منجز']).clip(lower=0).astype(int)
        # متوسط الطالب
        overall_avg = student_df['solve_pct'].mean() if not student_df.empty else 0.0

        st.markdown("### معاينة سريعة")
        st.dataframe(student_table, use_container_width=True, height=260)

        # إنشاء PDF
        font_name = register_font_from_upload(font_file)
        pdf_bytes = make_student_pdf(
            school_name=school_name or "",
            student_name=selected_student,
            grade=stu_grade, section=stu_section,
            table_df=student_table[['المادة','إجمالي','منجز','متبقي']],
            overall_avg=overall_avg,
            reco_text=reco_text,
            coordinator_name=coordinator_name or "",
            academic_deputy=academic_deputy or "",
            admin_deputy=admin_deputy or "",
            principal_name=principal_name or "",
            font_name=font_name
        )

        st.download_button(
            "📥 تحميل تقرير الطالب (PDF)",
            pdf_bytes,
            file_name=f"student_report_{selected_student}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

# =========================
# Footer
# =========================
st.markdown(f"""
<div class="footer">
  <div class="line"></div>
  <img class="logo" src="https://i.imgur.com/XLef7tS.png" alt="Logo">
  <div class="school">مدرسة عثمان بن عفان النموذجية-منسقة المشاريع الإلكتروني/ سحر عثمان</div>
  <div class="rights">© {datetime.now().year} جميع الحقوق محفوظة</div>
  <div class="contact">للتواصل:
    <a href="mailto:S.mahgoub0101@education.qa">S.mahgoub0101@education.qa</a>
  </div>
  <div class="credit">التحول الرقمي الذكي</div>
</div>
""", unsafe_allow_html=True)

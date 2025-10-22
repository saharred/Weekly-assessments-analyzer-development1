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
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import openai # لاستخدام ميزات الذكاء الاصطناعي

# تهيئة عميل OpenAI (يستخدم المتغيرات البيئية تلقائيًا)
try:
    client = openai.OpenAI()
except Exception as e:
    # st.error(f"فشل تهيئة عميل الذكاء الاصطناعي: {e}. قد لا تعمل ميزات الذكاء الاصطناعي.")
    client = None

# --------------- إعداد الصفحة ---------------
st.set_page_config(page_title="أي إنجاز", page_icon="📊", layout="wide")

POSITIVE_STATUS = ["solved","yes","1","تم","منجز","✓","✔","✅"]
NEGATIVE_STATUS = ["no","0","غير منجز","لم يحل","خطأ"]

# --------------- التوصيات الثابتة الجديدة (على مستوى الطالب) ---------------
STUDENT_RECOMMENDATIONS = {
    "🏆 Platinum": "نثمن تميزك المستمر، لقد أظهرت إبداعًا واجتهادًا ملحوظًا. استمر في استخدام نظام قطر للتعليم بفعالية، فأنت نموذج يحتذى به.",
    "🥇 Gold": "أحسنت! مستواك يعكس التزامًا رائعًا، نثق أنك بمتابعة الجهد ستنتقل لمستوى أعلى. استمر في تفعيل نظام قطر داخل الصف.",
    "🥈 Silver": "عملك جيد ويستحق التقدير، ومع مزيد من الممارسة والتفاعل مع نظام قطر ستصل إلى مستويات أرفع. نحن فخورون بك.",
    "🥉 Bronze": "لقد أظهرت جهدًا مشكورًا، ونشجعك على بذل المزيد من العطاء. باستخدام نظام قطر بشكل أعمق ستتطور قدراتك بشكل أكبر.",
    "🔧 Needs Improvement": "نرى لديك إمكانيات واعدة، لكن تحتاج لمزيد من الالتزام باستخدام نظام قطر للتعليم. نوصيك بالمثابرة والمشاركة النشطة، ونحن بجانبك لتتقدم.",
    "🚫 Not Utilizing System": "لم يظهر بعد استفادة كافية من نظام قطر للتعليم، وندعوك إلى تفعيل النظام بشكل أكبر لتحقيق النجاح. نحن نثق أن لديك القدرة على التغيير والتميز."
}

# --------------- أدوات مساعدة ---------------
def _strip_invisible_and_diacritics(s: str) -> str:
    # إزالة الأحرف غير المرئية والتشكيل
    s = re.sub(r"[\u200b-\u200f\u202a-\u202e\u064b-\u0652\u0640]", "", s)
    return s

def _normalize_arabic_digits(s: str) -> str:
    # تحويل الأرقام العربية إلى هندية (لاتينية) لتوحيد الصفوف والشعب
    s = str(s).replace("٠", "0").replace("١", "1").replace("٢", "2").replace("٣", "3")
    s = s.replace("٤", "4").replace("٥", "5").replace("٦", "6").replace("٧", "7")
    s = s.replace("٨", "8").replace("٩", "9")
    return s

def arabic_cleanup(s: str) -> str:
    if pd.isna(s): return ""
    return re.sub(r"\s+"," ",str(s).strip())

def normalize_name(s: str) -> str:
    s = arabic_cleanup(s)
    s = _strip_invisible_and_diacritics(s)
    s = s.replace("أ","ا").replace("إ","ا").replace("آ","ا")
    return s

def normalize_status(val) -> int:
    if pd.isna(val): return 0
    try:
        return 1 if float(str(val)) > 0 else 0
    except:
        return 1 if str(val).lower() in POSITIVE_STATUS else 0

def parse_sheet_subject(sheet_name: str) -> Tuple[str,str,str]:
    name = arabic_cleanup(sheet_name); grade=""; section=""
    # محاولة استخراج (المادة) (الصف) (الشعبة)
    m = re.search(r"(.+?)\s+(\d{1,2})\s+([A-Za-z0-9]+)\s*$", name)
    if m:
        # توحيد تنسيق الصف ليكون رقمين (مثل 07)
        grade_num = int(m.group(2))
        grade = f"المستوى {grade_num:02d}"
        return arabic_cleanup(m.group(1)), grade, m.group(3)
    return name, grade, section

def detect_header_row(df: pd.DataFrame, default_header_row: int) -> int:
    hr = default_header_row if default_header_row < len(df) else 0
    if df.iloc[hr].notna().sum() <= 1 and len(df)>0:
        if df.iloc[0].notna().sum() > 1: hr = 0
    return hr

def process_excel_file(file_obj, file_name, start_row_students: int,
                       selected_sheets: Optional[List[str]]=None) -> List[Dict]:
    rows = []
    try:
        xls = pd.ExcelFile(file_obj)
    except Exception as e:
        st.error(f"تعذّر فتح {file_name}: {e}")
        return rows
    sheets = xls.sheet_names if not selected_sheets else [s for s in xls.sheet_names if s in selected_sheets]
    for sh in sheets:
        try:
            raw = pd.read_excel(xls, sheet_name=sh, header=None)
            if raw.empty:
                st.warning(f"الشيت '{sh}' فارغ في {file_name}"); continue
            hr = detect_header_row(raw, max(0, start_row_students-3))
            header = [arabic_cleanup(x) for x in raw.iloc[hr].tolist()]
            eval_cols = {idx:title for idx,title in enumerate(header) if idx>0 and title}
            subject, grade, section = parse_sheet_subject(sh)
            first_row = max(start_row_students, hr+1)
            for r in range(first_row, len(raw)):
                row = raw.iloc[r]
                student_name = arabic_cleanup(row[0]) if 0 in raw.columns else ""
                if len(student_name) < 2: continue
                for c_idx, eval_name in eval_cols.items():
                    if c_idx < len(row):
                        solved = normalize_status(row[c_idx])
                        rows.append({
                            "student_name": student_name,
                            "student_name_norm": normalize_name(student_name),
                            "student_id": "", # يُفترض أن يكون في العمود 1 إذا وجد
                            "subject": subject,
                            "evaluation": eval_name,
                            "solved": int(solved),
                            "class": grade,
                            "section": section,
                            "teacher_email": "" # سيتم ملؤه لاحقًا
                        })
        except Exception as e:
            st.warning(f"تعذّر قراءة '{sh}' في {file_name}: {e}")
    return rows

def _load_teachers_df(file) -> Optional[pd.DataFrame]:
    """يرفع ملف المعلمات ويعيد DataFrame موحّد الأعمدة: الشعبة، اسم المعلمة، البريد الإلكتروني
       يدعم CSV وXLSX."""
    if file is None:
        return None

    name = file.name.lower()

    # 1) قراءة الملف حسب الامتداد
    try:
        if name.endswith(".csv"):
            tdf = pd.read_csv(file)
        elif name.endswith(".xlsx"):
            tdf = pd.read_excel(file, engine="openpyxl")
        elif name.endswith(".xls"):
            # نحاول xlrd .. وإن لم توجد نطلب من المستخدم التحويل
            try:
                import xlrd  # noqa: F401
                tdf = pd.read_excel(file, engine="xlrd")
            except Exception:
                st.error(
                    "هذا الملف بصيغة .xls ويتطلب حزمة xlrd غير متوفرة.\n"
                    "فضلاً احفظي الملف كـ **.xlsx** أو **.csv** ثم ارفعيه مجددًا."
                )
                return None
        else:
            st.error("صيغة الملف غير مدعومة. ارفعي ملفًا بصيغة **CSV** أو **XLSX**.")
            return None
    except Exception as e:
        st.error(f"❌ قراءة ملف المعلمات فشلت: {e}")
        return None

    # 2) توحيد أسماء الأعمدة (مرن مع الاختلافات)
    def _norm_header(x: str) -> str:
        x = _strip_invisible_and_diacritics(str(x)).lower()
        x = x.replace("أ","ا").replace("إ","ا").replace("آ","ا").replace("ة","ه")
        # إزالة كل الرموز غير الحروف/الأرقام لتسهيل المطابقة
        return re.sub(r"[^0-9a-z\u0600-\u06FF]+", "", x)

    cols_map = {c: _norm_header(c) for c in tdf.columns}

    def find_col(possible_keys: List[str]) -> Optional[str]:
        norm_keys = [_norm_header(k) for k in possible_keys]
        for original, normed in cols_map.items():
            if normed in norm_keys:
                return original
        return None

    # تم تعديل مفاتيح البحث لتكون أكثر شمولاً
    sec_col  = find_col(["الشعبة", "شعبة", "section", "القسم", "صف", "الصف", "الصف والشعبة"])
    name_col = find_col(["اسم المعلمة", "المعلمة", "Teacher", "teacher",
                         "اسم المعلم", "المعلم", "اسم المعلمه", "المعلم المسؤول"])
    mail_col = find_col(["البريد الإلكتروني", "البريد الالكتروني", "البريد",
                         "email", "e-mail", "ايميل", "الايميل", "بريد المعلم"])

    if not (sec_col and name_col and mail_col):
        st.error("❌ يجب أن يحتوي الملف على أعمدة: **الشعبة**، **اسم المعلمة**، **البريد الإلكتروني**.")
        with st.expander("الأعمدة التي تم اكتشافها"):
            st.write(list(tdf.columns))
        return None

    tdf = tdf[[sec_col, name_col, mail_col]].rename(columns={
        sec_col:  "الشعبة",
        name_col: "اسم المعلمة",
        mail_col: "البريد الإلكتروني"
    })

    # 3) تنظيف القيم
    tdf["الشعبة"] = (tdf["الشعبة"].astype(str)
                      .apply(_strip_invisible_and_diacritics)
                      .map(_normalize_arabic_digits)
                      .str.strip())
    tdf["اسم المعلمة"]      = tdf["اسم المعلمة"].astype(str).str.strip()
    tdf["البريد الإلكتروني"] = tdf["البريد الإلكتروني"].astype(str).str.strip()

    return tdf

# --------- أهم رقعة: ضمان uid دائمًا ---------
def ensure_uid(df: pd.DataFrame) -> pd.DataFrame:
    """يضمن وجود uid، وتوحيد الصف/الشعبة، وإزالة التكرار الأساسي."""
    if "student_name_norm" not in df.columns:
        df["student_name_norm"] = df.get("student_name", "").astype(str).apply(normalize_name)

    df["student_id"] = df.get("student_id", "").fillna("").astype(str).str.strip()
    uid = df["student_id"].copy()
    mask = (uid == "")
    uid[mask] = df.loc[mask, "student_name_norm"].astype(str)
    df["uid"] = uid

    df["class"]   = df.get("class","").fillna("").astype(str).str.strip()
    df["section"] = df.get("section","").fillna("").astype(str).str.strip()
    # توحيد تنسيق الصف ليكون رقمين (مثل 07)
    df["class"] = df["class"].str.replace(r"المستوى\s*(\d)", r"المستوى 0\1", regex=True)

    # إزالة تكرار نفس (uid, subject, evaluation)
    if {"subject","evaluation","solved"}.issubset(df.columns):
        df = (df.sort_values(["uid","subject","evaluation","solved"],
                             ascending=[True,True,True,False])
                .drop_duplicates(subset=["uid","subject","evaluation"], keep="first"))
    return df

# --------- Pivot/ملخص يعتمد uid ---------
def build_summary_pivot(df: pd.DataFrame, teachers_df: Optional[pd.DataFrame], thresholds: Dict[str,int]):
    if df.empty:
        return pd.DataFrame(), []
    
    # 1. تأكيد uid
    if "uid" not in df.columns:
        df = ensure_uid(df.copy())

    # 2. إجماليات لكل (uid/مادة)
    grp = (df.groupby(["uid","subject"], dropna=False)
             .agg(solved=("solved","sum"), total=("solved","count"))
             .reset_index())
    subjects = sorted(grp["subject"].dropna().unique().tolist())

    # 3. Pivot Table (إصلاح مشكلة KeyError: 'uid')
    piv = grp.pivot_table(index=["uid"], columns="subject",
                          values=["solved","total"], fill_value=0, aggfunc="sum").reset_index()
    
    # إصلاح تسمية الأعمدة بعد pivot_table
    new_columns = []
    for met, subj in piv.columns:
        if subj == "":
            new_columns.append("uid")
        else:
            new_columns.append(f"{subj}_{met}")
    piv.columns = new_columns

    # 4. دمج البيانات الوصفية (اسم الطالب، الصف، الشعبة)
    meta = (df.sort_values(["uid"])
              .groupby("uid", as_index=False)
              .agg(student_name=("student_name","first"),
                   student_id=("student_id","first"),
                   classx=("class","first"),
                   sectionx=("section","first")))

    piv = meta.merge(piv, on="uid", how="left")

    # 5. حساب الإنجاز العام والتصنيف
    solved_cols = [c for c in piv.columns if str(c).endswith("_solved")]
    total_cols  = [c for c in piv.columns if str(c).endswith("_total")]
    piv["Overall_Solved"] = piv[solved_cols].sum(axis=1) if solved_cols else 0
    piv["Overall_Total"]  = piv[total_cols].sum(axis=1)  if total_cols else 0
    piv["Overall_Completion"] = (piv["Overall_Solved"]/piv["Overall_Total"]*100).fillna(0).round(2)

    # *****************************************************************
    # تعديل دالة التصنيف لتتوافق مع المعايير الجديدة (أكبر من)
    # *****************************************************************
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
            return "🔧 Needs Improvement" # نسبة إنجاز > 0 ولكنها لم تصل لعتبة البرونزي
    piv["Category"] = piv["Overall_Completion"].apply(cat)
    # *****************************************************************

    # 6. دمج بيانات المعلمين (إذا توفرت)
    if teachers_df is not None:
        # إنشاء عمود للمطابقة (الصف والشعبة)
        piv["merge_key"] = piv["classx"] + " " + piv["sectionx"]
        teachers_df["merge_key"] = teachers_df["الشعبة"]
        
        # دمج بيانات المعلمين
        piv = piv.merge(teachers_df[["merge_key", "اسم المعلمة", "البريد الإلكتروني"]], 
                        left_on="merge_key", right_on="merge_key", how="left").drop(columns=["merge_key"])
        
        # إضافة عمود بريد المعلم لكل مادة (افتراضًا أن المعلم هو نفسه لكل المواد في الشعبة)
        piv["teacher_email"] = piv["البريد الإلكتروني"].fillna("")
        piv["teacher_name"] = piv["اسم المعلمة"].fillna("")
    else:
        piv["teacher_email"] = ""
        piv["teacher_name"] = ""

    # 7. التوصيات التشغيلية (الثابتة)
    piv["Student_Recommendation"] = piv["Category"].apply(lambda x: STUDENT_RECOMMENDATIONS.get(x, STUDENT_RECOMMENDATIONS["🔧 Needs Improvement"]))

    # 8. إعادة تسمية الأعمدة النهائية
    out = piv.rename(columns={
        "student_name":"اسم الطالب", "student_id":"الرقم الشخصي",
        "classx":"الصف", "الشعبة":"الشعبة",
        "Overall_Total":"إجمالي التقييمات",
        "Overall_Solved":"المنجز",
        "Overall_Completion":"نسبة الإنجاز %",
        "Category":"الفئة",
        "teacher_name": "اسم المعلمة",
        "teacher_email": "بريد المعلمة",
        "Student_Recommendation": "توصية الطالب"
    })

    # 9. pending لكل مادة + تنظيف أسماء الأعمدة
    for subj in subjects:
        t=f"{subj}_total"; s=f"{subj}_solved"
        if t in out.columns and s in out.columns:
            out[f"{subj}_pending"] = (out[t]-out[s]).clip(lower=0)
    out.columns = out.columns.astype(str).str.strip()

    base = ["اسم الطالب","الرقم الشخصي","الصف","الشعبة","اسم المعلمة","بريد المعلمة",
            "إجمالي التقييمات","المنجز","نسبة الإنجاز %","الفئة", "توصية الطالب"]
    others = [c for c in out.columns if c not in base and c != "uid"]
    out = out[base + others]
    return out, subjects

# --------- ميزة الذكاء الاصطناعي: التوصيات المخصصة (تم إزالتها واستبدالها بالثابتة) ---------
# @st.cache_data(show_spinner="جاري توليد التوصيات الذكية...")
# def generate_ai_recommendation(row: pd.Series, subjects: List[str]) -> str:
#     ...

# --------- دالة إرسال البريد الإلكتروني ---------
def send_teacher_emails(summary_df: pd.DataFrame, inactive_threshold: float = 10.0):
    if not st.session_state.get("smtp_configured", False):
        st.warning("لم يتم إعداد بيانات خادم SMTP. لا يمكن إرسال الإيميلات.")
        return False
    
    # تحديد الطلاب غير الفاعلين (أقل من نسبة إنجاز معينة)
    inactive_students = summary_df[summary_df["نسبة الإنجاز %"] < inactive_threshold]
    
    if inactive_students.empty:
        st.info("لا يوجد طلاب غير فاعلين لإرسال تنبيهات بشأنهم.")
        return True

    # التجميع حسب المعلم والبريد الإلكتروني
    teacher_groups = inactive_students.groupby(["بريد المعلمة", "اسم المعلمة"])
    
    for (email, teacher_name), group in teacher_groups:
        if not email or email == "nan":
            st.warning(f"تجاهل إرسال تنبيهات لـ {teacher_name}: لا يوجد بريد إلكتروني مسجل.")
            continue
        
        # بناء محتوى الإيميل
        student_list = "\n".join([f"- {row['اسم الطالب']} (نسبة الإنجاز: {row['نسبة الإنجاز %']:.1f}%)" 
                                  for index, row in group.iterrows()])
        
        subject = f"تنبيه: طلاب غير فاعلين في الشعبة {group['الشعبة'].iloc[0]}"
        body = f"""
        عزيزتي المعلمة {teacher_name}،
        
        تحية طيبة وبعد،
        
        نود تنبيهك بوجود مجموعة من الطلاب في شعبتك لم تتجاوز نسبة إنجازهم {inactive_threshold:.1f}% في التقييمات الأخيرة.
        
        **قائمة الطلاب غير الفاعلين:**
        {student_list}
        
        **التوصية التشغيلية العامة:**
        {analyze_teacher_group(group)}
        
        الرجاء التواصل مع هؤلاء الطلاب ومتابعتهم لرفع نسبة إنجازهم.
        
        شكراً لجهودك،
        نظام أي إنجاز الآلي
        """
        
        # هنا يتم استدعاء دالة الإرسال الفعلية (يجب أن يتم تنفيذها خارج هذا الكود)
        # مثال:
        # try:
        #     _send_email(to=email, subject=subject, body=body)
        #     st.success(f"تم إرسال تنبيه إلى {teacher_name} ({email}).")
        # except Exception as e:
        #     st.error(f"فشل إرسال الإيميل إلى {teacher_name}: {e}")
        
        st.info(f"تم توليد تنبيه لـ {teacher_name} ({email}) بخصوص {len(group)} طالب غير فاعل.")
        
    return True

# --------- ميزة الذكاء الاصطناعي: توصية على مستوى المعلم (المجموعة) ---------
@st.cache_data(show_spinner="جاري توليد توصية للمعلم...")
def analyze_teacher_group(group_df: pd.DataFrame) -> str:
    # تم إزالة الاعتماد على LLM واستبدالها بمنطق بسيط
    avg_completion = group_df["نسبة الإنجاز %"].mean()
    
    # استخدام التوصيات الثابتة على مستوى المادة (المرفقة في pasted_content_3.txt)
    if avg_completion >= 90:
        return "أظهر طلاب الصف التزامًا عاليًا بإنجاز التقييمات الأسبوعية على نظام قطر للتعليم بنسبة تفوق 90%. نوصي بالاستمرار على هذا النهج مع تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، وتوظيف النظام في رقمنة استراتيجية الصفوف المقلوبة لزيادة عمق الفهم."
    elif avg_completion >= 75:
        return "حقق الصف مستوى جيد جدًا في حل التقييمات الأسبوعية. نقترح تعزيز هذا الأداء عبر تذكير الطلاب في نهاية كل حصة بإنجاز التقييمات، مع توظيف نظام قطر للتعليم لدعم الصفوف المقلوبة وتطوير مهارات التفكير العليا."
    elif avg_completion >= 60:
        return "بلغ الصف نسبة إنجاز متوسطة في التقييمات الأسبوعية (60–75%). نوصي بتكثيف تذكير الطلاب بإنجاز التقييمات في نهاية كل حصة، وتفعيل دور النظام في الصفوف المقلوبة لتشجيع الطلاب على التفاعل المستمر."
    elif avg_completion >= 40:
        return "نسبة الحل ما زالت تحتاج إلى تحسين، إذ لم تتجاوز 60%. نوصي بالتركيز على تذكير الطلاب يوميًا في نهاية كل حصة بأهمية حل التقييمات، مع دمج استراتيجيات التعلم النشط ورقمنة الصفوف المقلوبة باستخدام نظام قطر للتعليم."
    elif avg_completion > 0:
        return "نسبة الإنجاز في التقييمات الأسبوعية ضعيفة على مستوى الصف. نوصي بتكثيف الجهود عبر التذكير المستمر بنهاية الحصص بإنجاز التقييمات، وتبسيط التمارين داخل نظام قطر للتعليم ضمن استراتيجية الصفوف المقلوبة لزيادة التفاعل."
    else:
        return "لم ينجز الصف أي تقييم أسبوعي في هذه المادة. نوصي بإطلاق خطة عاجلة تشمل: تذكير الطلاب بنهاية كل حصة بأهمية إنجاز التقييمات، والتواصل مع أولياء الأمور، مع اعتماد نظام قطر للتعليم كمنصة رئيسية لرقمنة استراتيجية الصفوف المقلوبة."

# --------- ميزة الذكاء الاصطناعي: تحليل الأنماط على مستوى المادة (تم استبدالها بالثابتة) ---------
@st.cache_data(show_spinner="جاري تحليل أنماط الإنجاز للمادة...")
def analyze_subject_patterns(summary_df: pd.DataFrame, subject: str) -> str:
    # تحديد متوسط نسبة الإنجاز للمادة
    total_solved = summary_df[f"{subject}_solved"].sum()
    total_total = summary_df[f"{subject}_total"].sum()
    avg_completion = (total_solved / total_total * 100) if total_total > 0 else 0
    
    # استخدام التوصيات الثابتة على مستوى المادة (المرفقة في pasted_content_3.txt)
    if avg_completion >= 90:
        return "المادة حققت نسبة إنجاز مرتفعة جدًا في التقييمات الأسبوعية على نظام قطر للتعليم. يُوصى بدعم استدامة هذا المستوى عبر توثيق أفضل الممارسات وتعميمها بين الصفوف، مع الحرص على التواصل المستمر مع أولياء الأمور لتعزيز الشراكة التربوية. كما يُوصى بـ تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، ورقمنة استراتيجية الصفوف المقلوبة بتوظيف نظام قطر للتعليم."
    elif avg_completion >= 75:
        return "المادة أظهرت نسبة إنجاز جيدة جدًا مع فرصة للارتقاء إلى مستوى الامتياز. يُوصى بزيادة التحفيز والمتابعة، والتواصل مع أولياء الأمور لدعم انتظام الطلاب، مع التأكيد على تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، ورقمنة استراتيجية الصفوف المقلوبة بتوظيف نظام قطر للتعليم."
    elif avg_completion >= 60:
        return "متوسط الإنجاز في المادة يعكس تفاعلًا مقبولًا مع التقييمات الأسبوعية، لكنه بحاجة إلى دفع إضافي. يُوصى بتعزيز المتابعة والتواصل مع أولياء الأمور لرفع مستوى الالتزام، مع الاستمرار في تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، ورقمنة استراتيجية الصفوف المقلوبة بتوظيف نظام قطر للتعليم."
    elif avg_completion >= 40:
        return "نسبة الإنجاز متوسطة منخفضة وتحتاج إلى رفع. يُوصى بزيادة المتابعة من القسم والتواصل مع أولياء الأمور لتحفيز الطلاب على الالتزام، مع التشديد على تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، ورقمنة استراتيجية الصفوف المقلوبة بتوظيف نظام قطر للتعليم."
    elif avg_completion > 0:
        return "المادة أظهرت ضعفًا في إنجاز التقييمات الأسبوعية. يُوصى بتدخل مباشر من القسم مع تفعيل التواصل مع أولياء الأمور بشكل منتظم لتعزيز التزام الطلاب، مع التركيز على تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، ورقمنة استراتيجية الصفوف المقلوبة بتوظيف نظام قطر للتعليم."
    else:
        return "لم يتم تسجيل أي إنجاز في التقييمات الأسبوعية لهذه المادة. يُوصى بمتابعة عاجلة من القسم، مع تكثيف التواصل مع أولياء الأمور لتوضيح أهمية الالتزام بالنظام، والتركيز على تذكير الطلاب دائمًا بحل التقييمات بنهاية كل حصة، ورقمنة استراتيجية الصفوف المقلوبة بتوظيف نظام قطر للتعليم."


def to_excel_bytes(sheets: Dict[str,pd.DataFrame]) -> bytes:
    mem = BytesIO()
    with pd.ExcelWriter(mem, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31] or "Sheet1", index=False)
    mem.seek(0); return mem.getvalue()

# ---------- Streamlit App (الواجهة الرئيسية) ----------
def main():
    st.title("📊 أي إنجاز - لوحة تحليل إنجاز الطلاب الذكية")
    st.markdown("أداة تحليلية تعتمد على التحليل الإحصائي والتوصيات الثابتة المخصصة لرفع نسبة الإنجاز الأكاديمي.")

    # 1. إعدادات النظام
    with st.sidebar.expander("⚙️ إعدادات النظام ومعايير التصنيف"):
        st.session_state.smtp_configured = st.checkbox("تم إعداد خادم SMTP (لإرسال الإيميلات)", value=False)
        inactive_threshold = st.slider("نسبة الإنجاز لاعتبار الطالب 'غير فاعل' (%)", 0, 50, 10)
        
        # *****************************************************************
        # تعديل القيم الافتراضية لعتبات التصنيف
        # *****************************************************************
        st.session_state.thresholds = {
            "Platinum": st.number_input("حد Platinum (أكبر من)", 0, 100, 89),
            "Gold": st.number_input("حد Gold (أكبر من)", 0, 100, 79),
            "Silver": st.number_input("حد Silver (أكبر من)", 0, 100, 49),
            "Bronze": st.number_input("حد Bronze (أكبر من)", 0, 100, 0) # البرونزي اكبر من صفر
        }
        # *****************************************************************

    # 2. تحميل ملفات المعلمين (لربط الطالب بالمعلم)
    teacher_file = st.sidebar.file_uploader("📂 تحميل ملف بيانات المعلمين (لإرسال الإيميلات)", type=["xlsx", "csv", "xls"])
    if teacher_file:
        teachers_df = _load_teachers_df(teacher_file)
        st.session_state.teachers_df = teachers_df
        if teachers_df is not None:
            st.sidebar.success(f"تم تحميل {len(teachers_df)} سجل معلم.")
            with st.sidebar.expander("معاينة بيانات المعلمين"):
                st.dataframe(teachers_df, use_container_width=True)
    else:
        st.session_state.teachers_df = None

    # 3. تحميل ملفات التقييمات
    uploaded_files = st.sidebar.file_uploader("📂 تحميل ملفات التقييمات (Excel)", type=["xlsx", "xls"], accept_multiple_files=True)
    
    if uploaded_files:
        raw_rows = []
        for file in uploaded_files:
            try:
                xls = pd.ExcelFile(file)
                selected_sheets = st.sidebar.multiselect(f"اختر أوراق من {file.name}", xls.sheet_names, default=xls.sheet_names)
                rows = process_excel_file(file, file.name, start_row_students=1, selected_sheets=selected_sheets)
                raw_rows.extend(rows)
            except Exception as e:
                st.error(f"فشل معالجة الملف {file.name}: {e}")
        
        if raw_rows:
            raw_df = pd.DataFrame(raw_rows)
            st.session_state.raw_df = raw_df
            st.sidebar.success(f"تم دمج {len(raw_df)} سجل بنجاح.")
            
            # 4. بناء الملخص المحوري
            with st.spinner("جاري تحليل البيانات وبناء التوصيات..."):
                summary_df, subjects = build_summary_pivot(
                    raw_df, 
                    st.session_state.teachers_df, 
                    st.session_state.get("thresholds", {"Platinum": 89, "Gold": 79, "Silver": 49, "Bronze": 0})
                )
            
            st.session_state.summary_df = summary_df
            st.session_state.subjects = subjects
            
            st.header("نتائج التحليل والإنجاز العام")
            st.dataframe(summary_df, use_container_width=True)
            
            # 5. الرسوم البيانية (مثال)
            if not summary_df.empty:
                col1, col2 = st.columns(2)
                with col1:
                    fig = px.histogram(summary_df, x="الفئة", color="الفئة", 
                                       title="توزيع الطلاب حسب فئة الإنجاز",
                                       category_orders={"الفئة": ["🏆 Platinum", "🥇 Gold", "🥈 Silver", "🥉 Bronze", "🔧 Needs Improvement", "🚫 Not Utilizing System"]})
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    fig2 = px.box(summary_df, y="نسبة الإنجاز %", color="الشعبة",
                                  title="توزيع نسب الإنجاز حسب الشعبة")
                    st.plotly_chart(fig2, use_container_width=True)

            # 6. التوصيات على مستوى المادة (الثابتة)
            st.header("تحليل الأنماط والتوصيات على مستوى المادة")
            for subj in subjects:
                with st.expander(f"✨ تحليل مادة: {subj}"):
                    st.markdown(f"**التوصية التشغيلية للمادة:**")
                    st.info(analyze_subject_patterns(summary_df, subj))
            
            # 7. خيارات التصدير والإيميل
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

if __name__ == "__main__":
    main()

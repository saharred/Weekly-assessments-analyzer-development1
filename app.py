# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Any
from datetime import datetime, date

def _parse_sheet(df: pd.DataFrame, sheet_name: str, today: date) -> pd.DataFrame:
    """
    توقعات البنية:
      - الصف 0: عناوين الأعمدة (H1.. = أسماء التقييمات)
      - الصف 2: مواعيد الاستحقاق (H3..)
    سنحاول أن نكون متسامحين لو تغيّر الموضع، لكن الافتراضي كما فوق.
    """
    # إذا جاءت برؤوس تلقائية: نقرأ بدون header ثم نعين الصف0 رؤوسًا
    # ضمان عدم فقدان صف 2 (مواعيد الاستحقاق)
    if df.columns.dtype != object:
        df.columns = [str(c) for c in df.columns]

    # لو جاءت بheader=0 من المصدر، نعيد قراءتها كبدون رؤوس:
    if not str(df.columns[0]).startswith("Unnamed"):
        # نعيد بناء بحيث الصف 0 رؤوس
        headers = df.iloc[0].fillna("").astype(str).tolist()
        df = df.iloc[1:].reset_index(drop=True)
        df.columns = headers
        due_row = 2 - 1  # لأننا نزعنا صف البداية
    else:
        # رؤوسنا ليست مناسبة؛ نعين يدويًا
        headers = df.iloc[0].fillna("").astype(str).tolist()
        df.columns = headers
        df = df.iloc[1:].reset_index(drop=True)
        due_row = 2  # الصف الثالث الأصلي

    # اسم الطالب: أول عمود غير فارغ يسمى غالبًا "اسم الطالب" أو ما شابه
    possible_name_cols = [c for c in df.columns if "اسم" in str(c) or "Student" in str(c)]
    name_col = possible_name_cols[0] if possible_name_cols else df.columns[0]

    # استخراج صف الاستحقاق كمُعجم {عمود: تاريخ}
    full_df = df.copy()
    try:
        due_series = full_df.iloc[due_row].to_dict()
    except Exception:
        due_series = {}

    # تحديد نقطة بداية الأعمدة التقييمية: من H وما بعده أو بعد "Overall"
    cols = list(full_df.columns)
    def col_idx(c): 
        try: return cols.index(c)
        except: return -1

    # موقع Overall (إن وجد)
    overall_idx = max([col_idx(c) for c in cols if str(c).strip().lower()=="overall"] + [-1])

    # نحدّد الأعمدة من H وما بعده: نحاول إيجاد H عبر الترتيب
    # لو ما نقدر، نستخدم كل الأعمدة بعد overall_idx
    start_idx = max(overall_idx + 1, 7)  # 7 ≈ H (صفرية)
    assess_cols = cols[start_idx:]

    # استبعاد الأعمدة غير التقييمية الشائعة
    ignore_tokens = {"overall","Unnamed","ملاحظات","Notes"}
    assess_cols = [c for c in assess_cols if all(tok.lower() not in str(c).lower() for tok in ignore_tokens)]

    # هيكلة صفوف الطلاب فقط (استبعاد صف الاستحقاق وما بعده إذا تداخل)
    students_df = full_df.copy()
    # احذف صف الاستحقاق إن كان متداخلًا داخل البيانات:
    students_df = students_df.drop(students_df.index[[due_row]] , errors="ignore")

    out_rows = []
    # استخراج subject/class من اسم الشيت (نمط: "03/1 Arabic" أو "Arabic 03/1")
    subj = sheet_name
    cls  = ""
    # محاولة بسيطة للفصل
    if " " in sheet_name:
        parts = sheet_name.split(" ")
        # ابحث عن 03/1 كصف/شعبة
        cand = [p for p in parts if "/" in p]
        if cand:
            cls = cand[0]
            subj = sheet_name.replace(cls,"").strip()

    for _, row in students_df.iterrows():
        student = str(row.get(name_col, "")).strip()
        if not student:
            continue

        # حصر التقييمات التي تاريخها <= اليوم
        total_due = 0
        completed = 0
        not_submitted = 0

        for c in assess_cols:
            val = row.get(c, np.nan)
            # due date من صف due_series
            due_raw = due_series.get(c, "")
            due_date = None
            if isinstance(due_raw, (datetime,)):
                due_date = due_raw.date()
            else:
                # حاول تفكيك كنص
                try:
                    due_date = pd.to_datetime(str(due_raw), dayfirst=True, errors="coerce")
                    if pd.notna(due_date):
                        due_date = due_date.date()
                except Exception:
                    due_date = None

            if due_date and due_date <= today:
                total_due += 1
                s = str(val).strip().upper()
                if s == "M":
                    not_submitted += 1
                elif s in {"I","AB","X",""}:
                    # نتجاهلها في العد المُنجز لكن ما تعتبر تسليم
                    pass
                else:
                    # لو رقم أو أي قيمة تحسب كتسليم
                    try:
                        _ = float(str(val).replace(",",""))
                        completed += 1
                    except Exception:
                        # نصوص أخرى تُعامل كتسليم (حسب توصيفك: كل غير M يُعد تسليم)
                        completed += 1

        completion = (100.0 * completed / max(total_due, 1)) if total_due>0 else None

        out_rows.append({
            "student_name": student,
            "subject": subj,
            "class": cls,
            "completed": int(completed),
            "total_due": int(total_due),
            "not_submitted": int(not_submitted),
            "completion_rate": round(completion if completion is not None else 0.0, 2) if total_due>0 else 0.0,
            "has_due": bool(total_due>0)
        })

    return pd.DataFrame(out_rows)


def load_week_xlsx_files(uploaded_files) -> dict:
    """
    يعيد قاموس { sheet_name: DataFrame_processed } لجميع الشيتات في كل ملف.
    """
    out = {}
    for f in uploaded_files:
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            raw = pd.read_excel(f, sheet_name=sheet, header=None)
            try:
                df = _parse_sheet(raw, sheet_name=sheet, today=datetime.today().date())
                if not df.empty:
                    out[sheet] = df
            except Exception:
                # نتجاوز الشيتات التي لا يمكن قراءتها
                continue
    return out

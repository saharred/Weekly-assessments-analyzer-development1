import pandas as pd
from datetime import datetime, date
from typing import List, Dict, Optional, Tuple, Union
import streamlit as st
import re

# Category thresholds and recommendations
CATEGORY_CONFIG = {
    "البلاتينية": {
        "threshold": 90,
        "recommendation": "أشكرك يا بطل على تميزك"
    },
    "الذهبي": {
        "threshold": 80,
        "recommendation": "أداء جيد جدًا، حل جميع التقييمات حتى تصعد إلى الفئة البلاتينية، أنت قدّها"
    },
    "الفضي": {
        "threshold": 70,
        "recommendation": "أداء جيد، حل جميع التقييمات حتى تصعد إلى الفئة البلاتينية"
    },
    "البرونزي": {
        "threshold": 60,
        "recommendation": "لديك الكثير من التقييمات لم يتم حلها لكن ما زال هناك وقت للوصول إلى الفئة البلاتينية"
    },
    "تحتاج إلى تحسين": {
        "threshold": 0,
        "recommendation": "اجتهد أكثر، هناك فرصة للوصول إلى الفئة البلاتينية"
    }
}

ZERO_SOLVED_MESSAGE = "لم يتم حل التقييمات الأسبوعية، حاول وستجد الرحلة ممتعة"

# Thresholds for performance analysis
PERFORMANCE_THRESHOLD = 70  # Students below 70% are considered inactive
CRITICAL_THRESHOLD = 50    # Students below 50% are critical


class AssessmentAnalyzer:
    def __init__(
        self,
        start_col_letter: str = "H",
        names_row: int = 5,
        names_col: str = "A",
        due_row: int = 3,
        # يقبل تاريخين من نوع date أو datetime
        date_range: Optional[Tuple[Union[date, datetime], Union[date, datetime]]] = None
    ):
        """
        Initialize assessment analyzer
        
        Args:
            start_col_letter: Column letter for assessment names (default H)
            names_row: Row number for student names (default 5)
            names_col: Column letter for student names (default A)
            due_row: Row number for due dates (default 3)
            date_range: Optional date range filter (start_date, end_date)
        """
        self.start_col_letter = start_col_letter.upper()
        self.names_row = names_row - 1  # Convert to 0-indexed (first student row)
        self.names_col = self._col_letter_to_index(names_col.upper())
        self.due_row = due_row - 1  # Convert to 0-indexed (due date row)
        self.date_range = date_range
    
    def _col_letter_to_index(self, col_letter: str) -> int:
        """Convert column letter (A, B, ..., Z, AA, AB, ...) to 0-indexed integer."""
        col_letter = col_letter.upper()
        col_idx = 0
        for char in col_letter:
            col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
        return col_idx - 1
    
    def _index_to_col_letter(self, col_idx: int) -> str:
        """Convert 0-indexed integer to column letter."""
        col_idx += 1
        col_letter = ""
        while col_idx > 0:
            col_idx -= 1
            col_letter = chr(col_idx % 26 + ord('A')) + col_letter
            col_idx //= 26
        return col_letter
    
    def _normalize_arabic_digits(self, text: str) -> str:
        """Convert Arabic-Indic digits (٠-٩) to ASCII digits (0-9)."""
        if not isinstance(text, str):
            return str(text)
        return text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

    def _parse_date(self, date_obj) -> Optional[date]:
        """Parse various date formats (Arabic/English/Excel serial) to date."""
        # Empty
        try:
            if pd.isna(date_obj):
                return None
        except Exception:
            pass

        # Pandas/Datetime
        if isinstance(date_obj, pd.Timestamp):
            return date_obj.date()
        if isinstance(date_obj, datetime):
            return date_obj.date()

        # Excel serial number
        if isinstance(date_obj, (int, float)) and not pd.isna(date_obj):
            try:
                base = pd.to_datetime("1899-12-30")
                dt = base + pd.to_timedelta(float(date_obj), unit="D")
                return dt.date()
            except Exception:
                pass

        # String parsing
        if isinstance(date_obj, str):
            s = date_obj.strip()
            if not s:
                return None
            s = self._normalize_arabic_digits(s)

            # Arabic months
            arabic_months = {
                "يناير": 1, "كانون الثاني": 1, "جانفي": 1,
                "فبراير": 2, "شباط": 2, "فيفري": 2,
                "مارس": 3, "اذار": 3, "آذار": 3,
                "ابريل": 4, "أبريل": 4, "نيسان": 4, "افريل": 4,
                "مايو": 5, "ماي": 5, "ايار": 5, "أيار": 5,
                "يونيو": 6, "يونيه": 6, "حزيران": 6, "جوان": 6,
                "يوليو": 7, "يوليه": 7, "تموز": 7, "جويلية": 7,
                "اغسطس": 8, "أغسطس": 8, "اب": 8, "آب": 8, "اوت": 8,
                "سبتمبر": 9, "ايلول": 9, "أيلول": 9, "سيبتمبر": 9,
                "اكتوبر": 10, "أكتوبر": 10, "تشرين الاول": 10, "تشرين الأول": 10,
                "نوفمبر": 11, "تشرين الثاني": 11, "نونبر": 11,
                "ديسمبر": 12, "كانون الاول": 12, "كانون الأول": 12, "دجنبر": 12,
            }

            def normalize_hamza(text: str) -> str:
                return text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ـ", "")

            m = re.search(r"(\d{1,2})\s*[-/\s]*\s*([^\d\s]+)", s)
            if m:
                try:
                    day = int(m.group(1))
                    month_name = m.group(2).strip()
                    month = arabic_months.get(month_name)
                    if month is None:
                        nm = normalize_hamza(month_name)
                        for k, v in arabic_months.items():
                            if normalize_hamza(k) == nm:
                                month = v
                                break
                    if month:
                        # Use current year when year is not provided
                        y = date.today().year
                        # Clamp day to 28 to be safe
                        safe_day = min(day, 28)
                        return date(y, month, safe_day)
                except Exception:
                    pass

            # Fallback to pandas
            try:
                parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
                if pd.notna(parsed):
                    return parsed.date()
            except Exception:
                pass

        return None
    
    def _is_ignored_value(self, value) -> bool:
        """Check if value should be ignored (I, AB, X, dashes, or empty)."""
        if pd.isna(value):
            return True
        str_value = str(value).strip().upper()
        return str_value in ["I", "AB", "X", "", "-", "—", "–", "NAN", "NONE"]
    
    def _is_missing_value(self, value) -> bool:
        """Check if value is 'M' (missing submission)."""
        if pd.isna(value):
            return False
        return str(value).strip().upper() == "M"
    
    def _get_category(self, solve_pct: float) -> str:
        """Determine category based on solve_pct."""
        for category, config in sorted(
            CATEGORY_CONFIG.items(),
            key=lambda x: x[1]["threshold"],
            reverse=True
        ):
            if solve_pct >= config["threshold"]:
                return category
        return "تحتاج إلى تحسين"
    
    def _get_recommendation(self, category: str, total: int, solved: int) -> str:
        """Get recommendation text based on category."""
        # Special case: no assessments solved but total > 0
        if total > 0 and solved == 0:
            return ZERO_SOLVED_MESSAGE
        
        return CATEGORY_CONFIG.get(category, {}).get("recommendation", "")
    
    def _parse_sheet_name(self, sheet_name: str) -> Tuple[str, str, str]:
        """
        Parse sheet name format: 'المادة المستوى الشعبة'
        Example: 'التربية الاسلامية 01 6' → ('التربية الاسلامية', '01', '6')
        """
        parts = sheet_name.strip().split()
        
        if len(parts) >= 3:
            # Last part is section (شعبة), second to last is level (مستوى)
            section = parts[-1]
            level = parts[-2]
            # Rest is subject name
            subject = ' '.join(parts[:-2])
            return subject, level, section
        elif len(parts) == 2:
            subject = parts[0]
            level = parts[1]
            section = ""
            return subject, level, section
        else:
            return sheet_name, "", ""
    
    def analyze_sheet(
        self,
        df: pd.DataFrame,
        sheet_name: str
    ) -> List[Dict]:
        """Analyze a single sheet and return list of student records."""
        results = []
        
        # Parse sheet name to get subject, level, and section
        subject, level, section = self._parse_sheet_name(sheet_name)
        
        # Find assessment columns (from H1 rightward)
        start_col_idx = self._col_letter_to_index(self.start_col_letter)
        assessment_columns = []
        
        # Row 1 (index 0) for assessment names
        headers_row_idx = 0
        if headers_row_idx < len(df):
            for col_idx in range(start_col_idx, len(df.columns)):
                header = df.iloc[headers_row_idx, col_idx]
                if pd.isna(header):
                    continue
                header_str = str(header).strip()

                # Ignore empty headers, OVERALL, or headers containing dashes
                if not header_str:
                    continue
                if header_str.upper() == "OVERALL":
                    continue
                if any(ch in header_str for ch in ["-", "—", "–"]):
                    continue

                # Get due date from due_row
                due_date: Optional[date] = None
                if self.due_row < len(df):
                    due_date_raw = df.iloc[self.due_row, col_idx]
                    due_date = self._parse_date(due_date_raw)

                # Date range filter (accept date or datetime in input)
                if self.date_range:
                    if due_date is None:
                        continue
                    start_date_raw, end_date_raw = self.date_range
                    start_date = start_date_raw.date() if isinstance(start_date_raw, datetime) else start_date_raw
                    end_date = end_date_raw.date() if isinstance(end_date_raw, datetime) else end_date_raw
                    if start_date and end_date and start_date > end_date:
                        start_date, end_date = end_date, start_date
                    if start_date and end_date and not (start_date <= due_date <= end_date):
                        continue

                # Skip columns that are fully empty/dashes for all students
                is_all_dashes = True
                for row_idx in range(self.names_row, len(df)):
                    cell_val = df.iloc[row_idx, col_idx]
                    if not self._is_ignored_value(cell_val):
                        is_all_dashes = False
                        break
                if is_all_dashes:
                    continue

                assessment_columns.append({
                    "col_idx": col_idx,
                    "name": header_str,
                    "due_date": due_date,
                })
        
        if not assessment_columns:
            st.warning(f"لم أجد أسماء تقييمات في H1 يميناً في ورقة '{sheet_name}'.")
            return results
        
        # Process each student (starting from row 5, index 4)
        for student_row_idx in range(self.names_row, len(df)):
            student_name = df.iloc[student_row_idx, self.names_col]
            
            if pd.isna(student_name) or str(student_name).strip() == "":
                continue
            
            student_name = str(student_name).strip()
            
            # Skip if it looks like a header or total row
            if student_name.upper() in ["الطالب", "الطالبة", "المجموع", "TOTAL"]:
                continue
            
            # Calculate metrics
            total_assessments = 0
            solved_assessments = 0
            remaining = 0
            unsolved_titles = []
            
            for assessment in assessment_columns:
                col_idx = assessment["col_idx"]
                value = df.iloc[student_row_idx, col_idx]
                
                if self._is_ignored_value(value):
                    continue
                
                total_assessments += 1
                
                if self._is_missing_value(value):
                    remaining += 1
                    unsolved_titles.append(assessment["name"])
                else:
                    solved_assessments += 1
            
            # Skip students with no assessments
            if total_assessments == 0:
                continue
            
            # Calculate solve percentage
            solve_pct = (solved_assessments / total_assessments * 100) if total_assessments > 0 else 0
            
            # Get category and recommendation
            category = self._get_category(solve_pct)
            recommendation = self._get_recommendation(category, total_assessments, solved_assessments)
            
            results.append({
                "student_name": student_name,
                "class": level,
                "section": section,
                "subject": subject,
                "total_material_solved": solved_assessments,
                "total_assessments": total_assessments,
                "remaining": remaining,
                "unsolved_assessment_count": len(unsolved_titles),
                "unsolved_titles": ", ".join(unsolved_titles) if unsolved_titles else "-",
                "solve_pct": round(solve_pct, 2),
                "category": category,
                "recommendation": recommendation
            })
        
        return results
    
    def analyze_file(
        self,
        file_obj,
        sheets: List[str]
    ) -> List[Dict]:
        """Analyze an uploaded file for specified sheets."""
        results = []
        
        try:
            # Determine engine based on file extension
            file_name = file_obj.name if hasattr(file_obj, 'name') else str(file_obj)
            engine = "xlrd" if file_name.endswith(".xls") else None
            
            # Read Excel file
            xls = pd.ExcelFile(file_obj, engine=engine)
            
            for sheet_name in sheets:
                if sheet_name not in xls.sheet_names:
                    continue
                
                # Read sheet without headers
                df = pd.read_excel(
                    file_obj,
                    sheet_name=sheet_name,
                    header=None,
                    engine=engine
                )
                
                sheet_results = self.analyze_sheet(df, sheet_name)
                results.extend(sheet_results)
        
        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {str(e)}")
        
        return results


def generate_html_report(student_row: pd.Series) -> str:
    """Generate an RTL HTML report for a single student."""
    
    # Determine category color
    category_colors = {
        "البلاتينية": "#f093fb",
        "الذهبي": "#ffd89b",
        "الفضي": "#a8edea",
        "البرونزي": "#ff9a56",
        "تحتاج إلى تحسين": "#ff6b6b"
    }
    
    category_color = category_colors.get(student_row['category'], "#667eea")
    
    html = f"""<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تقرير الطالب - {student_row['student_name']}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif, 'Arial';
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            direction: rtl;
            text-align: right;
        }}
        
        .container {{
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .content {{
            padding: 40px;
        }}
        
        .student-info {{
            background: #f5f5f5;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            border-right: 4px solid #667eea;
        }}
        
        .info-row {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 15px;
        }}
        
        .info-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        
        .info-label {{
            font-weight: bold;
            color: #333;
            margin-left: 10px;
        }}
        
        .info-value {{
            color: #666;
            font-size: 1.1em;
        }}
        
        .category-box {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            margin: 30px 0;
        }}
        
        .category-box h2 {{
            font-size: 2em;
            margin-bottom: 10px;
        }}
        
        .category-box p {{
            font-size: 1.1em;
            opacity: 0.95;
        }}
        
        .recommendation {{
            background: #e8f5e9;
            border-right: 4px solid #4caf50;
            padding: 20px;
            border-radius: 4px;
            margin: 20px 0;
            direction: rtl;
            text-align: right;
        }}
        
        .recommendation p {{
            color: #2e7d32;
            font-size: 1.1em;
            line-height: 1.6;
        }}
        
        .metrics {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            margin: 30px 0;
        }}
        
        .metric {{
            background: #f9f9f9;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            border-top: 3px solid #667eea;
        }}
        
        .metric-value {{
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
        }}
        
        .metric-label {{
            font-size: 0.9em;
            color: #666;
            margin-top: 5px;
        }}
        
        .unsolved-section {{
            margin: 30px 0;
        }}
        
        .unsolved-section h3 {{
            color: #333;
            margin-bottom: 15px;
            font-size: 1.3em;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }}
        
        .unsolved-list {{
            background: #fff3e0;
            padding: 15px;
            border-radius: 8px;
        }}
        
        .unsolved-list p {{
            color: #e65100;
            line-height: 1.6;
        }}
        
        .footer {{
            background: #f5f5f5;
            padding: 20px;
            text-align: center;
            color: #999;
            font-size: 0.9em;
            margin-top: 30px;
            border-top: 1px solid #ddd;
        }}
        
        @media print {{
            body {{
                background: white;
            }}
            .container {{
                box-shadow: none;
                max-width: 100%;
            }}
            .header {{
                page-break-after: avoid;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>تقرير الطالب</h1>
            <p>Weekly Assessment Report</p>
        </div>
        
        <div class="content">
            <div class="student-info">
                <div class="info-row">
                    <div class="info-item">
                        <span class="info-value">{student_row['student_name']}</span>
                        <span class="info-label">اسم الطالب:</span>
                    </div>
                    <div class="info-item">
                        <span class="info-value">{student_row['subject']}</span>
                        <span class="info-label">المادة:</span>
                    </div>
                </div>
                <div class="info-row">
                    <div class="info-item">
                        <span class="info-value">{student_row['class']}</span>
                        <span class="info-label">المستوى:</span>
                    </div>
                    <div class="info-item">
                        <span class="info-value">{student_row['section']}</span>
                        <span class="info-label">الشعبة:</span>
                    </div>
                </div>
            </div>
            
            <div class="metrics">
                <div class="metric">
                    <div class="metric-value">{int(student_row['total_material_solved'])}</div>
                    <div class="metric-label">تقييمات منجزة</div>
                </div>
                <div class="metric">
                    <div class="metric-value">{int(student_row.get('remaining', student_row.get('unsolved_assessment_count', 0)))}</div>
                    <div class="metric-label">تقييمات متبقية</div>
                </div>
                <div class="metric">
                    <div class="metric-value">{int(student_row.get('total_assessments', 0))}</div>
                    <div class="metric-label">إجمالي التقييمات</div>
                </div>
                <div class="metric">
                    <div class="metric-value">{student_row['solve_pct']:.1f}%</div>
                    <div class="metric-label">نسبة الإنجاز</div>
                </div>
            </div>
            
            <div class="category-box" style="background: linear-gradient(135deg, {category_color} 0%, {category_color}dd 100%);">
                <h2>{student_row['category']}</h2>
                <p>الفئة</p>
            </div>
            
            <div class="recommendation">
                <p>💡 {student_row['recommendation']}</p>
            </div>
            
            <div class="unsolved-section">
                <h3>التقييمات غير المنجزة</h3>
                <div class="unsolved-list">
                    <p>{student_row['unsolved_titles']}</p>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p>تم إنشاء التقرير بواسطة Weekly Assessments Analyzer v3.7</p>
            <p>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>"""
    
    return html

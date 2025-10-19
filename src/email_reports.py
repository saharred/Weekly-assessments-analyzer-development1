import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing import List, Dict, Optional, Tuple

# Constants for performance analysis
PERFORMANCE_THRESHOLD = 70  # Students below 70% are inactive
CRITICAL_THRESHOLD = 50    # Students below 50% are critical


class SubjectReportGenerator:
    """Generate descriptive reports for each subject/class/section"""
    
    def __init__(self):
        self.now = datetime.now()
    
    def generate_subject_report(
        self,
        subject: str,
        level: str,
        section: str,
        students_data: List[Dict]
    ) -> str:
        """Generate a descriptive report for a subject"""
        
        df = pd.DataFrame(students_data)
        
        # Calculate statistics
        total_students = len(df)
        avg_solve_pct = df['solve_pct'].mean()
        
        # Categorize students
        high_performers = df[df['solve_pct'] >= 90]
        good_performers = df[(df['solve_pct'] >= 70) & (df['solve_pct'] < 90)]
        inactive_students = df[(df['solve_pct'] >= CRITICAL_THRESHOLD) & (df['solve_pct'] < PERFORMANCE_THRESHOLD)]
        critical_students = df[df['solve_pct'] < CRITICAL_THRESHOLD]
        
        # Generate report
        report = f"""
╔══════════════════════════════════════════════════════════════╗
║           تقرير تحليل التقييمات الأسبوعية                   ║
║        WEEKLY ASSESSMENT ANALYSIS REPORT                      ║
╚══════════════════════════════════════════════════════════════╝

📋 معلومات القسم:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
المادة:          {subject}
المستوى:         {level}
الشعبة:          {section}
تاريخ التقرير:   {self.now.strftime('%Y-%m-%d %H:%M')}

📊 الإحصائيات العامة:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
عدد الطلاب الكلي:       {total_students} طالب/طالبة
متوسط نسبة الإنجاز:     {avg_solve_pct:.2f}%
عدد الطلاب المتميزين:   {len(high_performers)} (≥ 90%)
عدد الطلاب الجيدين:     {len(good_performers)} (70% - 89%)
عدد الطلاب غير الفاعلين: {len(inactive_students)} (50% - 69%)
عدد الطلاب في الخطر:    {len(critical_students)} (< 50%)

🎯 الأداء التحليلي:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✨ الطلاب المتميزون ({len(high_performers)}):
"""
        
        if len(high_performers) > 0:
            report += self._format_student_list(high_performers, "⭐")
        else:
            report += "   لا يوجد طلاب متميزون حالياً\n"
        
        report += f"""
✅ الطلاب الجيدون ({len(good_performers)}):
"""
        if len(good_performers) > 0:
            report += self._format_student_list(good_performers, "✓")
        else:
            report += "   لا يوجد طلاب بأداء جيد\n"
        
        report += f"""
⚠️ الطلاب غير الفاعلين - يحتاجون متابعة ({len(inactive_students)}):
"""
        if len(inactive_students) > 0:
            report += self._format_student_list(inactive_students, "⚠")
            report += self._generate_inactive_actions(inactive_students)
        else:
            report += "   لا يوجد طلاب في هذه الفئة (جيد!)\n"
        
        report += f"""
🔴 الطلاب في الخطر الشديد - متابعة فورية ({len(critical_students)}):
"""
        if len(critical_students) > 0:
            report += self._format_student_list(critical_students, "🔴")
            report += self._generate_critical_actions(critical_students)
        else:
            report += "   لا يوجد طلاب في وضع حرج (ممتاز!)\n"
        
        report += f"""
📝 التوصيات:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        report += self._generate_recommendations(df, high_performers, good_performers, inactive_students, critical_students)
        
        report += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
تم إنشاء التقرير بواسطة: Weekly Assessments Analyzer v3.7
التاريخ: {self.now.strftime('%Y-%m-%d %H:%M:%S')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        
        return report
    
    def _format_student_list(self, students_df: pd.DataFrame, icon: str) -> str:
        """Format student list with statistics"""
        result = ""
        for idx, (_, student) in enumerate(students_df.iterrows(), 1):
            remaining_val = int(student.get('remaining', student.get('unsolved_assessment_count', 0)))
            total_val = int(student.get('total_assessments', student.get('total', 0)))
            solved_val = int(student.get('total_material_solved', student.get('solved', 0)))
            result += f"""   {icon} {idx}. {student['student_name']}
      النسبة: {student['solve_pct']:.2f}% | منجز: {solved_val} | متبقي: {remaining_val} | إجمالي: {total_val}
"""
        return result
    
    def _generate_inactive_actions(self, students_df: pd.DataFrame) -> str:
        """Generate action items for inactive students"""
        actions = "\n   الإجراءات المقترحة:\n"
        actions += "   • التواصل مع الطالب/الطالبة للتذكير\n"
        actions += "   • تقديم دعم إضافي في التقييمات\n"
        actions += "   • متابعة أسباب التأخر\n"
        actions += "   • التشاور مع ولي الأمر إذا لزم\n"
        return actions
    
    def _generate_critical_actions(self, students_df: pd.DataFrame) -> str:
        """Generate action items for critical students"""
        actions = "\n   الإجراءات المقترحة (فورية):\n"
        actions += "   • اتصال فوري مع الطالب/الطالبة وولي الأمر\n"
        actions += "   • جلسة تقوية فورية\n"
        actions += "   • تحديد أسباب الضعف\n"
        actions += "   • خطة دعم شاملة\n"
        actions += "   • متابعة يومية\n"
        return actions
    
    def _generate_recommendations(
        self,
        all_students: pd.DataFrame,
        high: pd.DataFrame,
        good: pd.DataFrame,
        inactive: pd.DataFrame,
        critical: pd.DataFrame
    ) -> str:
        """Generate general recommendations"""
        recommendations = ""
        
        # Overall assessment
        total = len(all_students)
        positive_percent = ((len(high) + len(good)) / total * 100) if total > 0 else 0
        
        recommendations += f"1. الأداء العام للفصل: {positive_percent:.1f}% أداء إيجابي\n"
        
        if len(critical) > 0:
            recommendations += f"\n2. ⚠️ تنبيه: يوجد {len(critical)} طالب/ة في وضع حرج\n"
            recommendations += "   يجب إجراء متابعة فورية ومكثفة\n"
        
        if len(inactive) > 0:
            recommendations += f"\n3. متابعة: يوجد {len(inactive)} طالب/ة بحاجة إلى تحفيز\n"
            recommendations += "   يفضل جلسات دعم تعليمي\n"
        
        if positive_percent >= 80:
            recommendations += "\n4. ✅ الأداء العام ممتاز، استمر على هذا النهج\n"
        elif positive_percent >= 60:
            recommendations += "\n4. 📈 الأداء جيد، هناك مجال للتحسن\n"
        else:
            recommendations += "\n4. 🔴 الأداء يحتاج تحسين فوري\n"
        
        recommendations += "\n5. الخطوات القادمة:\n"
        recommendations += "   • متابعة دورية أسبوعية\n"
        recommendations += "   • جلسات تعزيز للطلاب المتميزين\n"
        recommendations += "   • برامج دعم للطلاب الضعفاء\n"
        recommendations += "   • تواصل منتظم مع أولياء الأمور\n"
        
        return recommendations


class EmailSender:
    """Send emails with reports"""
    
    def __init__(self, smtp_server: str, smtp_port: int, sender_email: str, sender_password: str):
        """
        Initialize email sender
        
        Args:
            smtp_server: SMTP server address
            smtp_port: SMTP port (usually 587 for TLS)
            sender_email: Sender email address
            sender_password: Sender password or app password
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
    
    def send_subject_report(
        self,
        teacher_email: str,
        subject: str,
        level: str,
        section: str,
        report_content: str,
        inactive_students: List[Dict],
        critical_students: List[Dict]
    ) -> tuple:
        """Send subject report to teacher"""
        
        try:
            # Create message
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f"تقرير التقييمات الأسبوعية - {subject} ({level}/{section})"
            msg['From'] = self.sender_email
            msg['To'] = teacher_email
            
            # Create HTML version of report
            html_report = self._convert_to_html(
                report_content,
                subject,
                level,
                section,
                inactive_students,
                critical_students
            )
            
            # Attach parts
            part1 = MIMEText(report_content, 'plain', 'utf-8')
            part2 = MIMEText(html_report, 'html', 'utf-8')
            msg.attach(part1)
            msg.attach(part2)
            
            # Send email
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
            
            return True, "تم إرسال التقرير بنجاح"
        
        except Exception as e:
            return False, f"خطأ في الإرسال: {str(e)}"
    
    def _convert_to_html(
        self,
        text_report: str,
        subject: str,
        level: str,
        section: str,
        inactive_students: List[Dict],
        critical_students: List[Dict]
    ) -> str:
        """Convert text report to HTML"""
        
        inactive_html = self._format_students_html(inactive_students, "warning")
        critical_html = self._format_students_html(critical_students, "danger")
        
        html = f"""
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تقرير التقييمات</title>
    <style>
        body {{
            font-family: 'Segoe UI', Arial, sans-serif;
            background: #f5f5f5;
            padding: 20px;
            direction: rtl;
            text-align: right;
        }}
        .container {{
            max-width: 900px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0;
            font-size: 24px;
        }}
        .header p {{
            margin: 5px 0 0 0;
            opacity: 0.9;
        }}
        .info-box {{
            background: #f9f9f9;
            padding: 15px;
            border-right: 4px solid #667eea;
            margin-bottom: 20px;
            border-radius: 4px;
        }}
        .info-box p {{
            margin: 5px 0;
            color: #333;
        }}
        .info-box strong {{
            color: #667eea;
        }}
        .section {{
            margin: 20px 0;
        }}
        .section h2 {{
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
            margin-bottom: 15px;
        }}
        .student-list {{
            background: #f9f9f9;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 15px;
        }}
        .student-item {{
            padding: 10px;
            margin-bottom: 10px;
            background: white;
            border-right: 3px solid #667eea;
            border-radius: 4px;
        }}
        .warning {{
            border-right-color: #ff9800;
        }}
        .danger {{
            border-right-color: #f44336;
        }}
        .student-name {{
            font-weight: bold;
            font-size: 16px;
            color: #333;
        }}
        .student-stats {{
            font-size: 13px;
            color: #666;
            margin-top: 5px;
        }}
        .badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 12px;
            margin-left: 5px;
        }}
        .badge-warning {{
            background: #fff3cd;
            color: #856404;
        }}
        .badge-danger {{
            background: #f8d7da;
            color: #721c24;
        }}
        .footer {{
            text-align: center;
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #999;
            font-size: 12px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>تقرير التقييمات الأسبوعية</h1>
            <p>{subject} | المستوى {level} | الشعبة {section}</p>
        </div>
        
        <div class="info-box">
            <p><strong>المادة:</strong> {subject}</p>
            <p><strong>المستوى:</strong> {level}</p>
            <p><strong>الشعبة:</strong> {section}</p>
            <p><strong>التاريخ:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        </div>
        
        {critical_html if critical_students else ""}
        {inactive_html if inactive_students else ""}
        
        <div class="section">
            <h2>📝 الملاحظات:</h2>
            <pre style="background: #f9f9f9; padding: 15px; border-radius: 4px;">{text_report}</pre>
        </div>
        
        <div class="footer">
            <p>تم إنشاء التقرير بواسطة Weekly Assessments Analyzer v3.7</p>
            <p>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>
"""
        return html
    
    def _format_students_html(self, students: List[Dict], style: str) -> str:
        """Format students as HTML"""
        if not students:
            return ""
        
        title = "🔴 الطلاب في الخطر الشديد" if style == "danger" else "⚠️ الطلاب غير الفاعلين"
        
        html = f"""
        <div class="section">
            <h2>{title}</h2>
            <div class="student-list">
"""
        
        for student in students:
            badge_class = "danger" if style == "danger" else "warning"
            badge_text = "خطر" if style == "danger" else "تحذير"
            
            html += f"""
                <div class="student-item {style}">
                    <div class="student-name">
                        {student['student_name']}
                        <span class="badge badge-{badge_class}">{badge_text}</span>
                    </div>
                    <div class="student-stats">
                        النسبة: {student['solve_pct']:.2f}% | منجز: {int(student.get('total_material_solved', 0))} | متبقي: {int(student.get('remaining', student.get('unsolved_assessment_count', 0)))} | إجمالي: {int(student.get('total_assessments', 0))}
                    </div>
                </div>
"""
        
        html += """
            </div>
        </div>
"""
        return html

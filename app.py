import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from typing import Tuple, Dict, List
import io
import base64

# ============================================
# CONSTANTS & CONFIGURATION
# ============================================

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

# Keyword mapping for Arabic suffixes
SUFFIX_KEYWORDS = {
    'total': ['إجمالي', 'اجمالي', 'Total'],
    'solved': ['منجز', 'Solved', 'Completed'],
    'remaining': ['متبقي', 'متبقّي', 'Remaining'],
    'percent': ['النسبة', 'نسبة', 'Percent', '%']
}


# ============================================
# HELPER FUNCTIONS
# ============================================

def assign_category(percent: float) -> str:
    """
    Assign a category based on completion percentage.
    
    Args:
        percent: Completion percentage (0-100)
        
    Returns:
        Category name in Arabic
    """
    if pd.isna(percent):
        return 'بحاجة لتحسين'
    
    for category, (min_val, max_val) in CATEGORY_THRESHOLDS.items():
        if min_val <= percent <= max_val:
            return category
    
    return 'بحاجة لتحسين'


def is_wide_format(df: pd.DataFrame) -> bool:
    """
    Detect if the DataFrame is in WIDE format (one row per student with subject columns).
    
    Wide format has columns like: "التربية الإسلامية - إجمالي", "الرياضيات - منجز"
    Long format has columns like: subject, student, solved, total
    """
    # Check for subject-suffix pattern in column names
    subject_pattern_count = sum(1 for col in df.columns if ' - ' in str(col))
    
    # If more than 20% of columns have the " - " pattern, consider it wide
    return subject_pattern_count > len(df.columns) * 0.2


def extract_subject_from_column(col_name: str) -> Tuple[str, str]:
    """
    Extract subject name and field type from column name.
    
    Example: "التربية الإسلامية - إجمالي" → ("التربية الإسلامية", "total")
    
    Returns:
        (subject_name, field_type)
    """
    if ' - ' not in str(col_name):
        return None, None
    
    parts = str(col_name).split(' - ')
    if len(parts) != 2:
        return None, None
    
    subject = parts[0].strip()
    suffix = parts[1].strip()
    
    # Map suffix to field type
    for field_type, keywords in SUFFIX_KEYWORDS.items():
        if any(keyword in suffix for keyword in keywords):
            return subject, field_type
    
    return subject, 'unknown'


def wide_to_long(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert WIDE format to LONG (normalized) format.
    
    Input columns: "الطالب", "الصف", "الشعبة", "<Subject> - إجمالي", "<Subject> - منجز", etc.
    Output columns: student, grade, section, subject, total, solved, remaining, percent
    """
    # Identify student info columns (non-subject columns)
    student_cols = []
    for col in ['الطالب', 'student', 'Student', 'الاسم', 'name']:
        if col in df.columns:
            student_cols.append(col)
            break
    
    for col in ['الصف', 'grade', 'Grade', 'المستوى', 'level']:
        if col in df.columns:
            student_cols.append(col)
            break
    
    for col in ['الشعبة', 'section', 'Section', 'الفصل']:
        if col in df.columns:
            student_cols.append(col)
            break
    
    # Extract subject columns
    subject_data = {}
    for col in df.columns:
        if col in student_cols:
            continue
        
        subject, field_type = extract_subject_from_column(col)
        if subject and field_type != 'unknown':
            if subject not in subject_data:
                subject_data[subject] = {}
            subject_data[subject][field_type] = col
    
    # Build long format
    long_rows = []
    
    for idx, row in df.iterrows():
        student_info = {col: row[col] for col in student_cols}
        
        for subject, fields in subject_data.items():
            record = student_info.copy()
            record['subject'] = subject
            
            # Extract values
            total = row.get(fields.get('total'), 0)
            solved = row.get(fields.get('solved'), 0)
            percent = row.get(fields.get('percent'), None)
            
            # Handle None/NaN
            total = 0 if pd.isna(total) else float(total)
            solved = 0 if pd.isna(solved) else float(solved)
            
            # Calculate percent if missing
            if pd.isna(percent) or percent is None:
                percent = (solved / total * 100) if total > 0 else 0.0
            else:
                percent = float(percent)
            
            record['total'] = int(total)
            record['solved'] = int(solved)
            record['percent'] = round(percent, 1)
            record['category'] = assign_category(percent)
            
            long_rows.append(record)
    
    return pd.DataFrame(long_rows)


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize input DataFrame to standard format.
    
    Handles both WIDE and LONG formats and returns:
    [student, grade, section, subject, total, solved, percent, category]
    """
    df = df.copy()
    
    # Detect format
    if is_wide_format(df):
        df = wide_to_long(df)
    
    # Standardize column names
    column_mapping = {
        'الطالب': 'student',
        'Student': 'student',
        'الاسم': 'student',
        'الصف': 'grade',
        'Grade': 'grade',
        'المستوى': 'grade',
        'الشعبة': 'section',
        'Section': 'section',
        'المادة': 'subject',
        'Subject': 'subject',
        'إجمالي': 'total',
        'Total': 'total',
        'منجز': 'solved',
        'Solved': 'solved',
        'النسبة': 'percent',
        'Percent': 'percent',
        'الفئة': 'category',
        'Category': 'category'
    }
    
    df = df.rename(columns=column_mapping)
    
    # Ensure required columns exist
    required_cols = ['subject', 'percent']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")
    
    # Calculate percent if missing total/solved
    if 'percent' not in df.columns and 'total' in df.columns and 'solved' in df.columns:
        df['percent'] = df.apply(
            lambda row: (row['solved'] / row['total'] * 100) if row['total'] > 0 else 0.0,
            axis=1
        )
    
    # Assign categories
    if 'category' not in df.columns:
        df['category'] = df['percent'].apply(assign_category)
    
    # Fill NaN
    df['percent'] = df['percent'].fillna(0)
    df['category'] = df['category'].fillna('بحاجة لتحسين')
    
    return df


def aggregate_by_subject(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate data by subject and category.
    
    Returns:
        DataFrame with: subject, category, count, percent_of_subject, avg_completion
    """
    # Group by subject and category
    agg_data = []
    
    for subject in df['subject'].unique():
        subject_df = df[df['subject'] == subject]
        total_students = len(subject_df)
        avg_completion = subject_df['percent'].mean()
        
        for category in CATEGORY_ORDER:
            count = len(subject_df[subject_df['category'] == category])
            percent_share = (count / total_students * 100) if total_students > 0 else 0
            
            agg_data.append({
                'subject': subject,
                'category': category,
                'count': count,
                'percent_share': round(percent_share, 1),
                'avg_completion': round(avg_completion, 1)
            })
    
    agg_df = pd.DataFrame(agg_data)
    
    # Sort by average completion descending
    subject_order = (
        agg_df.groupby('subject')['avg_completion']
        .first()
        .sort_values(ascending=False)
        .index.tolist()
    )
    
    agg_df['subject'] = pd.Categorical(agg_df['subject'], categories=subject_order, ordered=True)
    agg_df = agg_df.sort_values('subject')
    
    return agg_df


def create_stacked_bar_chart(
    agg_df: pd.DataFrame,
    mode: str = 'percent'
) -> go.Figure:
    """
    Create a horizontal stacked bar chart.
    
    Args:
        agg_df: Aggregated DataFrame
        mode: 'percent' for % share or 'count' for absolute counts
        
    Returns:
        Plotly Figure object
    """
    # Prepare data
    subjects = agg_df['subject'].unique()
    
    fig = go.Figure()
    
    for category in CATEGORY_ORDER:
        category_data = agg_df[agg_df['category'] == category]
        
        if mode == 'percent':
            values = category_data['percent_share'].tolist()
            text = [f"{v:.1f}%" if v > 0 else "" for v in values]
            hovertemplate = (
                "<b>%{y}</b><br>"
                "الفئة: " + category + "<br>"
                "العدد: %{customdata[0]}<br>"
                "النسبة: %{x:.1f}%<br>"
                "<extra></extra>"
            )
        else:
            values = category_data['count'].tolist()
            text = [str(int(v)) if v > 0 else "" for v in values]
            hovertemplate = (
                "<b>%{y}</b><br>"
                "الفئة: " + category + "<br>"
                "العدد: %{x}<br>"
                "النسبة: %{customdata[0]:.1f}%<br>"
                "<extra></extra>"
            )
        
        fig.add_trace(go.Bar(
            name=category,
            x=values,
            y=category_data['subject'].tolist(),
            orientation='h',
            marker=dict(
                color=CATEGORY_COLORS[category],
                line=dict(color='white', width=1)
            ),
            text=text,
            textposition='inside',
            textfont=dict(size=11, color='black', family='Cairo'),
            hovertemplate=hovertemplate,
            customdata=np.column_stack((
                category_data['count'].tolist(),
                category_data['percent_share'].tolist()
            ))
        ))
    
    # Update layout
    title = "توزيع الفئات حسب المادة" if mode == 'percent' else "عدد الطلاب حسب الفئة والمادة"
    xaxis_title = "النسبة المئوية (%)" if mode == 'percent' else "عدد الطلاب"
    
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=20, family='Cairo', color='#8A1538', weight=700),
            x=0.5,
            xanchor='center'
        ),
        xaxis=dict(
            title=xaxis_title,
            titlefont=dict(size=14, family='Cairo', color='#111827'),
            tickfont=dict(size=12, family='Cairo'),
            gridcolor='#E5E7EB',
            range=[0, 100] if mode == 'percent' else None
        ),
        yaxis=dict(
            title="المادة",
            titlefont=dict(size=14, family='Cairo', color='#111827'),
            tickfont=dict(size=12, family='Cairo'),
            autorange='reversed'
        ),
        barmode='stack',
        height=max(400, len(subjects) * 60),
        margin=dict(l=200, r=50, t=80, b=50),
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Cairo'),
        legend=dict(
            title=dict(text="الفئة", font=dict(size=14, family='Cairo')),
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='center',
            x=0.5,
            font=dict(size=12, family='Cairo')
        ),
        hovermode='closest'
    )
    
    return fig


def export_chart_as_png(fig: go.Figure) -> bytes:
    """Export Plotly figure as PNG bytes."""
    return fig.to_image(format="png", width=1200, height=800, scale=2)


# ============================================
# MAIN FUNCTION
# ============================================

def render_subject_category_chart(df: pd.DataFrame) -> Tuple[go.Figure, pd.DataFrame]:
    """
    Render subject-level category distribution chart with full Streamlit UI.
    
    Args:
        df: Input DataFrame (wide or long format)
        
    Returns:
        (figure, aggregated_dataframe)
    """
    
    # Add custom CSS
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700&display=swap');
        
        .chart-container {
            background: white;
            border: 2px solid #E5E7EB;
            border-right: 5px solid #8A1538;
            border-radius: 12px;
            padding: 24px;
            margin: 20px 0;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        }
        
        .chart-title {
            font-size: 24px;
            font-weight: 700;
            color: #8A1538;
            text-align: center;
            margin-bottom: 16px;
            font-family: 'Cairo', sans-serif;
        }
        
        .info-box {
            background: #F5F5F5;
            border-right: 4px solid #C9A646;
            padding: 16px;
            border-radius: 8px;
            margin: 16px 0;
            font-family: 'Cairo', sans-serif;
        }
        
        .category-legend {
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            justify-content: center;
            margin: 16px 0;
            font-family: 'Cairo', sans-serif;
        }
        
        .category-item {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            background: white;
            border: 1px solid #E5E7EB;
            border-radius: 6px;
        }
        
        .category-color {
            width: 20px;
            height: 20px;
            border-radius: 4px;
            border: 1px solid #ccc;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Section header
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="chart-title">📊 توزيع الفئات حسب المادة الدراسية</h2>', unsafe_allow_html=True)
    
    # Info box with thresholds
    st.markdown("""
    <div class="info-box">
        <strong>📌 معايير التصنيف:</strong><br>
        🥇 <strong>بلاتيني:</strong> 90% فأكثر | 
        🥈 <strong>ذهبي:</strong> 80-89% | 
        🥉 <strong>فضي:</strong> 70-79% | 
        🟤 <strong>برونزي:</strong> 60-69% | 
        🔴 <strong>بحاجة لتحسين:</strong> أقل من 60%
    </div>
    """, unsafe_allow_html=True)
    
    try:
        # Normalize data
        normalized_df = normalize_dataframe(df)
        
        # Filters (if grade/section columns exist)
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        filtered_df = normalized_df.copy()
        
        with col_filter1:
            if 'grade' in normalized_df.columns:
                grades = ['الكل'] + sorted(normalized_df['grade'].dropna().unique().tolist())
                selected_grade = st.selectbox('🎓 الصف', grades, key='grade_filter')
                if selected_grade != 'الكل':
                    filtered_df = filtered_df[filtered_df['grade'] == selected_grade]
        
        with col_filter2:
            if 'section' in normalized_df.columns:
                sections = ['الكل'] + sorted(normalized_df['section'].dropna().unique().tolist())
                selected_section = st.selectbox('📚 الشعبة', sections, key='section_filter')
                if selected_section != 'الكل':
                    filtered_df = filtered_df[filtered_df['section'] == selected_section]
        
        with col_filter3:
            chart_mode = st.radio(
                'نوع العرض',
                ['النسبة المئوية (%)', 'العدد المطلق'],
                horizontal=True,
                key='chart_mode'
            )
            mode = 'percent' if chart_mode == 'النسبة المئوية (%)' else 'count'
        
        # Aggregate data
        agg_df = aggregate_by_subject(filtered_df)
        
        # Create chart
        fig = create_stacked_bar_chart(agg_df, mode=mode)
        
        # Display chart
        st.plotly_chart(fig, use_container_width=True, key='category_chart')
        
        # Download buttons
        col_download1, col_download2 = st.columns(2)
        
        with col_download1:
            # CSV download
            csv = agg_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 تحميل البيانات (CSV)",
                data=csv,
                file_name=f"subject_categories_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col_download2:
            # PNG download
            try:
                png_bytes = export_chart_as_png(fig)
                st.download_button(
                    label="📥 تحميل الرسم البياني (PNG)",
                    data=png_bytes,
                    file_name=f"chart_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.png",
                    mime="image/png",
                    use_container_width=True
                )
            except Exception as e:
                st.info("💡 لتحميل الصورة، قم بتثبيت: pip install kaleido")
        
        # Expandable data table
        with st.expander("📋 عرض البيانات المُجمّعة"):
            # Prepare display DataFrame
            display_df = agg_df.pivot(
                index='subject',
                columns='category',
                values='count'
            ).fillna(0).astype(int)
            
            # Add totals and average
            display_df['المجموع'] = display_df.sum(axis=1)
            
            # Add average completion
            avg_completion = agg_df.groupby('subject')['avg_completion'].first()
            display_df['متوسط الإنجاز (%)'] = avg_completion
            
            # Rename index
            display_df.index.name = 'المادة'
            
            st.dataframe(
                display_df.style.background_gradient(
                    subset=['متوسط الإنجاز (%)'],
                    cmap='RdYlGn',
                    vmin=0,
                    vmax=100
                ),
                use_container_width=True
            )
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        return fig, agg_df
        
    except Exception as e:
        st.error(f"❌ خطأ في معالجة البيانات: {str(e)}")
        st.exception(e)
        return None, None


# ============================================
# UNIT TESTS
# ============================================

def test_wide_to_long_parser():
    """Test the wide→long parser with synthetic Arabic column names."""
    
    # Create synthetic WIDE data
    test_data = {
        'الطالب': ['أحمد', 'فاطمة', 'محمد'],
        'الصف': ['01', '01', '02'],
        'الشعبة': ['1', '1', '2'],
        'التربية الإسلامية - إجمالي': [5, 5, 5],
        'التربية الإسلامية - منجز': [5, 4, 2],
        'التربية الإسلامية - النسبة': [100, 80, 40],
        'الرياضيات - إجمالي': [10, 10, 10],
        'الرياضيات - منجز': [8, 9, 5],
        'الرياضيات - النسبة': [80, 90, 50]
    }
    
    df_wide = pd.DataFrame(test_data)
    
    print("🧪 Testing WIDE → LONG parser...")
    print("\nOriginal WIDE format:")
    print(df_wide.head())
    
    # Convert to long
    df_long = wide_to_long(df_wide)
    
    print("\nConverted LONG format:")
    print(df_long.head(6))
    
    # Assertions
    assert 'subject' in df_long.columns, "Missing 'subject' column"
    assert 'percent' in df_long.columns, "Missing 'percent' column"
    assert 'category' in df_long.columns, "Missing 'category' column"
    assert len(df_long) == 6, f"Expected 6 rows (3 students × 2 subjects), got {len(df_long)}"
    
    print("\n✅ All tests passed!")


# ============================================
# EXAMPLE USAGE
# ============================================

def main():
    st.set_page_config(
        page_title="تحليل الفئات - نظام قطر للتعليم",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 نظام تحليل الفئات الدراسية")
    st.markdown("---")
    
    # Create sample data
    sample_data = {
        'الطالب': ['أحمد محمد', 'فاطمة علي', 'محمد حسن', 'سارة أحمد', 'علي خالد'] * 3,
        'الصف': ['01'] * 15,
        'الشعبة': ['1'] * 15,
        'التربية الإسلامية - إجمالي': [5] * 15,
        'التربية الإسلامية - منجز': [5, 4, 4, 3, 2, 5, 5, 4, 3, 5, 4, 3, 2, 1, 5],
        'الرياضيات - إجمالي': [10] * 15,
        'الرياضيات - منجز': [10, 9, 8, 7, 5, 9, 8, 7, 6, 10, 9, 8, 7, 6, 5],
        'اللغة العربية - إجمالي': [8] * 15,
        'اللغة العربية - منجز': [8, 7, 6, 5, 4, 7, 6, 5, 8, 7, 6, 5, 4, 3, 8]
    }
    
    df = pd.DataFrame(sample_data)
    
    # Render the chart
    fig, agg_df = render_subject_category_chart(df)
    
    # Additional statistics
    if agg_df is not None:
        st.markdown("---")
        st.subheader("📈 إحصائيات إضافية")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_subjects = agg_df['subject'].nunique()
            st.metric("عدد المواد", total_subjects)
        
        with col2:
            platinum_total = agg_df[agg_df['category'] == 'بلاتيني 🥇']['count'].sum()
            st.metric("إجمالي البلاتيني", platinum_total)
        
        with col3:
            avg_all = agg_df.groupby('subject')['avg_completion'].first().mean()
            st.metric("المتوسط العام", f"{avg_all:.1f}%")
        
        with col4:
            needs_improvement = agg_df[agg_df['category'] == 'بحاجة لتحسين']['count'].sum()
            st.metric("بحاجة لتحسين", needs_improvement)


if __name__ == "__main__":
    # Run tests
    print("=" * 60)
    test_wide_to_long_parser()
    print("=" * 60)
    
    # Run main app
    main()

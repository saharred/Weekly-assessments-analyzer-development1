import streamlit as st
from streamlit_drawable_canvas import st_canvas
from datetime import datetime
import pandas as pd
from PIL import Image
import io
import base64

def setup_form_theme_and_signatures():
    """
    Creates a fully-styled signatures section with proper form input styling.
    
    Fixes the white-on-white text issue by explicitly overriding input colors.
    Returns a dictionary with captured signature data.
    """
    
    # Inject CSS to fix form styling and create signature section
    st.markdown("""
    <style>
        /* ============================================
           CRITICAL FIX: Form Input Text Color Override
           ============================================ */
        
        /* Force dark text in ALL input fields (overrides inherited white) */
        input:not([type="checkbox"]):not([type="radio"]),
        textarea,
        select,
        [data-baseweb="select"] > div,
        [data-baseweb="input"] > div > input,
        [data-baseweb="textarea"] > textarea {
            color: #111827 !important;
            background-color: #FFFFFF !important;
            border: 1.5px solid #E5E7EB !important;
            border-radius: 6px !important;
            padding: 10px 12px !important;
            font-size: 14px !important;
            font-family: 'Cairo', sans-serif !important;
            direction: rtl !important;
            text-align: right !important;
        }
        
        /* Placeholder text - visible but subtle */
        input::placeholder,
        textarea::placeholder {
            color: #6B7280 !important;
            opacity: 1 !important;
        }
        
        /* Focus state - Qatar maroon accent */
        input:focus,
        textarea:focus,
        select:focus {
            outline: 1.5px solid #8A1538 !important;
            border-color: #8A1538 !important;
            box-shadow: 0 0 0 2px rgba(138, 21, 56, 0.1) !important;
        }
        
        /* Date picker specific styling */
        input[type="date"] {
            color: #111827 !important;
            background-color: #FFFFFF !important;
        }
        
        /* Select dropdown styling */
        [data-baseweb="select"] [role="button"] {
            background-color: #FFFFFF !important;
            color: #111827 !important;
            border: 1.5px solid #E5E7EB !important;
        }
        
        /* Disabled state */
        input:disabled,
        textarea:disabled,
        select:disabled {
            background-color: #F5F5F5 !important;
            color: #6B7280 !important;
            cursor: not-allowed !important;
        }
        
        /* ============================================
           DARK MODE SUPPORT (Media Query)
           ============================================ */
        
        @media (prefers-color-scheme: dark) {
            input:not([type="checkbox"]):not([type="radio"]),
            textarea,
            select {
                color: #F9FAFB !important;
                background-color: #1F2937 !important;
                border-color: #374151 !important;
            }
            
            input::placeholder,
            textarea::placeholder {
                color: #9CA3AF !important;
            }
        }
        
        /* ============================================
           SIGNATURES SECTION STYLING
           ============================================ */
        
        .signatures-container {
            background: #FFFFFF;
            border: 2px solid #E5E7EB;
            border-right: 5px solid #8A1538;
            border-radius: 12px;
            padding: 32px;
            margin: 32px 0;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            direction: rtl;
        }
        
        .signatures-title {
            font-size: 24px;
            font-weight: 700;
            color: #8A1538;
            margin-bottom: 24px;
            text-align: right;
            font-family: 'Cairo', sans-serif;
            border-bottom: 2px solid #C9A646;
            padding-bottom: 12px;
        }
        
        .signature-field-label {
            font-size: 15px;
            font-weight: 600;
            color: #111827;
            margin-bottom: 8px;
            text-align: right;
            font-family: 'Cairo', sans-serif;
            display: block;
        }
        
        .signature-canvas-container {
            background: #FFFFFF;
            border: 2px dashed #C9A646;
            border-radius: 8px;
            padding: 16px;
            margin: 16px 0;
            text-align: center;
        }
        
        .signature-canvas-label {
            font-size: 14px;
            font-weight: 600;
            color: #6B7280;
            margin-bottom: 12px;
            display: block;
            font-family: 'Cairo', sans-serif;
        }
        
        /* Button styling for signature controls */
        .stButton > button {
            background: linear-gradient(135deg, #8A1538 0%, #6B1029 100%) !important;
            color: #FFFFFF !important;
            border: none !important;
            padding: 10px 24px !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            font-family: 'Cairo', sans-serif !important;
            transition: all 0.3s ease !important;
            cursor: pointer !important;
        }
        
        .stButton > button:hover {
            background: linear-gradient(135deg, #6B1029 0%, #8A1538 100%) !important;
            box-shadow: 0 4px 12px rgba(138, 21, 56, 0.3) !important;
            transform: translateY(-2px) !important;
        }
        
        /* Secondary button (clear) */
        .stButton.secondary > button {
            background: #FFFFFF !important;
            color: #8A1538 !important;
            border: 2px solid #8A1538 !important;
        }
        
        .stButton.secondary > button:hover {
            background: #8A1538 !important;
            color: #FFFFFF !important;
        }
        
        /* ============================================
           ACCESSIBILITY ENHANCEMENTS
           ============================================ */
        
        /* Ensure minimum contrast ratios (WCAG AA) */
        label {
            color: #111827 !important;
            font-weight: 500 !important;
        }
        
        /* Focus indicators for keyboard navigation */
        *:focus-visible {
            outline: 2px solid #8A1538 !important;
            outline-offset: 2px !important;
        }
        
        /* High contrast mode support */
        @media (prefers-contrast: high) {
            input, textarea, select {
                border-width: 2px !important;
            }
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Create the signatures section
    st.markdown('<div class="signatures-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="signatures-title">📝 خانة التوقيعات</h2>', unsafe_allow_html=True)
    
    # Create three columns for the form fields
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<label class="signature-field-label">اسم الموقع *</label>', unsafe_allow_html=True)
        signatory_name = st.text_input(
            "اسم الموقع",
            key="sig_name",
            placeholder="أدخل الاسم الكامل",
            label_visibility="collapsed"
        )
    
    with col2:
        st.markdown('<label class="signature-field-label">المنصب / الوظيفة *</label>', unsafe_allow_html=True)
        signatory_role = st.text_input(
            "المنصب",
            key="sig_role",
            placeholder="مثال: مدير المدرسة",
            label_visibility="collapsed"
        )
    
    with col3:
        st.markdown('<label class="signature-field-label">التاريخ *</label>', unsafe_allow_html=True)
        signature_date = st.date_input(
            "التاريخ",
            value=datetime.now(),
            key="sig_date",
            label_visibility="collapsed"
        )
    
    st.markdown("---")
    
    # Signature Canvas Section
    st.markdown('<div class="signature-canvas-container">', unsafe_allow_html=True)
    st.markdown('<label class="signature-canvas-label">✍️ منطقة التوقيع الرقمي</label>', unsafe_allow_html=True)
    
    # Create signature pad using streamlit-drawable-canvas
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)",  # Transparent fill
        stroke_width=3,
        stroke_color="#111827",  # Dark stroke
        background_color="#FFFFFF",  # White canvas
        background_image=None,
        update_streamlit=True,
        height=180,
        width=520,
        drawing_mode="freedraw",
        point_display_radius=0,
        key="signature_canvas",
    )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Control buttons
    col_clear, col_download = st.columns(2)
    
    with col_clear:
        if st.button("🗑️ مسح التوقيع", key="clear_sig", use_container_width=True):
            st.rerun()
    
    with col_download:
        if canvas_result.image_data is not None:
            # Convert canvas to PNG
            img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            byte_im = buf.getvalue()
            
            # Create download button
            st.download_button(
                label="⬇️ تنزيل التوقيع (PNG)",
                data=byte_im,
                file_name=f"signature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png",
                key="download_sig",
                use_container_width=True
            )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Prepare return data
    signature_data = {
        "name": signatory_name,
        "role": signatory_role,
        "date": signature_date.strftime("%Y-%m-%d") if signature_date else None,
        "signature_image": canvas_result.image_data if canvas_result.image_data is not None else None,
        "has_signature": canvas_result.image_data is not None
    }
    
    return signature_data


# ============================================
# EXAMPLE USAGE
# ============================================

def main():
    st.set_page_config(
        page_title="نظام التوقيعات الإلكترونية",
        page_icon="✍️",
        layout="wide"
    )
    
    # Add Google Fonts
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700&display=swap');
        * {
            font-family: 'Cairo', sans-serif;
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.title("📋 نظام التوقيعات الإلكترونية")
    st.markdown("---")
    
    # Call the signatures function
    signature_data = setup_form_theme_and_signatures()
    
    # Display captured data (for testing/debugging)
    st.markdown("---")
    st.subheader("📊 البيانات المُدخلة")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("الاسم", signature_data["name"] if signature_data["name"] else "لم يُدخل")
    
    with col2:
        st.metric("المنصب", signature_data["role"] if signature_data["role"] else "لم يُدخل")
    
    with col3:
        st.metric("التاريخ", signature_data["date"] if signature_data["date"] else "لم يُدخل")
    
    with col4:
        st.metric("التوقيع", "✅ موجود" if signature_data["has_signature"] else "❌ غير موجود")
    
    # Show raw data
    with st.expander("🔍 عرض البيانات الخام (للمطورين)"):
        st.json({
            "name": signature_data["name"],
            "role": signature_data["role"],
            "date": signature_data["date"],
            "has_signature": signature_data["has_signature"]
        })


if __name__ == "__main__":
    main()

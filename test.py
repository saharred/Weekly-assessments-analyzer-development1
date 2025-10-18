import streamlit as st

st.title("اختبار رفع الملفات")

uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])

if uploaded_file:
    st.success(f"تم رفع الملف: {uploaded_file.name}")
else:
    st.warning("لم يتم رفع أي ملف")

# Sidebar
with st.sidebar:
    st.markdown(f"<div style='text-align: center; margin: 20px 0;'>{MINISTRY_LOGO_SVG}</div>", unsafe_allow_html=True)
    st.markdown("---")
    st.header("⚙️ الإعدادات والتحليل")
    
    # File Upload
    st.subheader("📁 تحميل الملفات")
    uploaded_files = st.file_uploader(
        "اختر ملفات Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="📌 يدعم تحليل عدة ملفات في آن واحد"
    )
    
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        try:
            all_sheets = []
            sheet_file_map = {}
            for file_idx, file in enumerate(uploaded_files):
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    sheet_display = f"[ملف {file_idx+1}] {sheet}"
                    all_sheets.append(sheet_display)
                    sheet_file_map[sheet_display] = (file, sheet)
            
            if all_sheets:
                st.info(f"📊 وجدت {len(all_sheets)} مادة من {len(uploaded_files)} ملفات")
                
                # خيار اختيار الأوراق
                select_all = st.checkbox("✅ اختر جميع الأوراق", value=True)
                
                if select_all:
                    selected_sheets_display = all_sheets
                else:
                    selected_sheets_display = st.multiselect(
                        "🔍 فلتر الأوراق: اختر الأوراق المراد تحليلها",
                        all_sheets,
                        default=[]
                    )
                
                selected_sheets = [(sheet_file_map[s][0], sheet_file_map[s][1]) for s in selected_sheets_display]
            else:
                selected_sheets = []
        except Exception as e:
            st.error(f"خطأ: {e}")
            selected_sheets = []
    else:
        st.info("💡 ارفع ملفات Excel للبدء")
        selected_sheets = []
        select_all = False
    
    st.markdown("---")
    st.subheader("🏫 معلومات المدرسة")
    school_name = st.text_input("📛 اسم المدرسة", placeholder="مثال: مدرسة قطر النموذجية")
    
    st.subheader("🖼️ شعار الوزارة/المدرسة")
    uploaded_logo = st.file_uploader("ارفع شعار", type=["png", "jpg", "jpeg"], help="سيظهر في التقارير الفردية")
    logo_base64 = ""
    if uploaded_logo:
        logo_base64 = base64.b64encode(uploaded_logo.read()).decode()
        st.success("✅ تم رفع الشعار")
    
    st.markdown("---")
    st.subheader("📅 فلتر التاريخ")
    
    date_filter_type = st.radio("🔍 نوع الفلتر:", ["بدون فلتر", "من تاريخ إلى تاريخ", "من تاريخ إلى الآن"])
    
    from_date = None
    to_date = None
    
    st.caption("💡 سيتم قراءة التواريخ تلقائياً من الملفات (الأعمدة I, J, K - الصف الثاني)")
    
    if date_filter_type == "من تاريخ إلى تاريخ":
        col1, col2 = st.columns(2)
        with col1:
            from_date = st.date_input("من تاريخ", key="from_date")
        with col2:
            to_date = st.date_input("إلى تاريخ", key="to_date")
    elif date_filter_type == "من تاريخ إلى الآن":
        from_date = st.date_input("من تاريخ", key="from_date_now")
        to_date = pd.Timestamp.now().date()
    
    st.markdown("---")
    st.subheader("✍️ معلومات التوقيعات")
    
    coordinator_name = st.text_input("👤 منسق المشاريع", placeholder="أدخل الاسم")
    academic_deputy = st.text_input("👨‍🏫 النائب الأكاديمي", placeholder="أدخل الاسم")
    admin_deputy = st.text_input("👨‍💼 النائب الإداري", placeholder="أدخل الاسم")
    principal_name = st.text_input("🎓 مدير المدرسة", placeholder="أدخل الاسم")
    
    st.markdown("---")
    
    # التحقق من صحة البيانات
    if st.checkbox("🔒 التحقق من صحة البيانات"):
        st.info("""
        ✓ الملفات المرفوعة: {}
        ✓ الأوراق المختارة: {}
        ✓ النطاق الزمني: {}
        """.format(len(uploaded_files) if uploaded_files else 0, len(selected_sheets), date_filter_type))
    
    run_analysis = st.button(
        "🚀 تشغيل التحليل والفلترة",
        use_container_width=True,
        type="primary",
        disabled=not (uploaded_files and selected_sheets)
    )

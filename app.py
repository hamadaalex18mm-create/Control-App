import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="برنامج توزيع كراسات الإجابة", layout="wide")

# ==========================================
# الحل النووي: قلب المنصة بالكامل لتدعم العربية (RTL)
# ==========================================
st.markdown(
    """
    <style>
    /* شقلبة التطبيق كله ليكون من اليمين لليسار */
    .stApp {
        direction: rtl !important;
    }
    
    /* إجبار كل النصوص والفقرات على محاذاة اليمين */
    p, div, span, label, h2, h3, h4, h5, h6 {
        text-align: right !important;
    }
    
    /* استثناء العنوان الرئيسي ليبقى متوسطاً في منتصف الصفحة */
    h1 {
        text-align: center !important;
        width: 100% !important;
    }

    /* إصلاح اتجاه الجداول */
    [data-testid="stDataFrame"] {
        direction: rtl !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)
# ==========================================

# --- تنسيق الهيدر (الشعارات والعنوان) ---
# بسبب شقلبة التطبيق، العمود الأول المكتوب في الكود سيظهر على "اليمين" في الشاشة
col_faculty_logo, col_title, col_unit_logo = st.columns([1, 3, 1])

with col_faculty_logo:
    # 1. شعار الكلية (سيظهر يميناً)
    if os.path.exists("logo_faculty.jpg"):
        st.image("logo_faculty.jpg", use_container_width=True)
    elif os.path.exists("logo_faculty.png"):
        st.image("logo_faculty.png", use_container_width=True)

with col_title:
    # العنوان في النص
    st.markdown(
        """
        <div style='margin-top: -30px; margin-bottom: 20px;'>
            <h1 style='color: #1E3A8A; font-family: Arial, sans-serif;'>توزيع كراسات الإجابة (الشجرة)</h1>
        </div>
        """,
        unsafe_allow_html=True
    )

with col_unit_logo:
    # 2. شعار الوحدة (سيظهر يساراً)
    if os.path.exists("logo_unit.jpg"):
        st.image("logo_unit.jpg", use_container_width=True)
    elif os.path.exists("logo_unit.png"):
        st.image("logo_unit.png", use_container_width=True)

st.markdown("---")

# --- الحروف الأبجدية ---
ARABIC_LETTERS = "أ ب ج د هـ و ز ح ط ي ك ل م ن س ع ف ص ق ر ش ت ث خ ذ ض ظ غ".split(" ")
ENGLISH_LETTERS = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z".split(" ")

if 'base_df' not in st.session_state:
    st.session_state.base_df = None

# --- مرحلة رفع الملف ---
if st.session_state.base_df is None:
    st.markdown(
        "<div style='background-color: #e8f4f8; padding: 15px; border-radius: 5px; color: #004d40; border: 1px solid #b6e3f4; margin-bottom: 15px; font-size: 16px;'>"
        "ℹ️ <strong>ملاحظة:</strong> يرجى رفع ملف الإكسيل الأساسي (يحتوي على: رقم اللجنة، مكان اللجنة)"
        "</div>", 
        unsafe_allow_html=True
    )
    uploaded_file = st.file_uploader("", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
            
            if "رقم اللجنة" in df.columns and "مكان اللجنة" in df.columns:
                st.session_state.base_df = df[['رقم اللجنة', 'مكان اللجنة']].copy()
                st.session_state.base_df['عدد الحضور'] = 0
                st.rerun()
            else:
                st.error("الملف يجب أن يحتوي على عمودين باسم: 'رقم اللجنة' و 'مكان اللجنة'")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

# --- مرحلة واجهة العمل ---
if st.session_state.base_df is not None:
    # العمود الأول (col_settings) سيظهر على اليمين، والعمود الثاني (col_data) على اليسار
    col_settings, col_data = st.columns([1, 4])
    
    with col_settings:
        st.markdown("<h3>الإعدادات</h3>", unsafe_allow_html=True)
        lang = st.radio("اختر لغة الشجرة:", ["عربي", "إنجليزي"])
        st.markdown("---")
        # الزرار هياخد العرض كله عشان شكله يبقى متناسق يمين
        if st.button("رفع ملف جديد", use_container_width=True):
            st.session_state.base_df = None
            st.rerun()

    with col_data:
        st.markdown("<h3>إدخال أعداد الحضور</h3>", unsafe_allow_html=True)
        
        # الترتيب الطبيعي للأعمدة، وبما إن التطبيق يمين-ليسار، رقم اللجنة هيكون أول واحد يمين
        input_display_cols = ['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور']
        df_for_editor = st.session_state.base_df[input_display_cols]
        
        input_table_height = (len(st.session_state.base_df) + 1) * 38
        
        edited_df = st.data_editor(
            df_for_editor,
            disabled=["رقم اللجنة", "مكان اللجنة"],
            hide_index=True,
            use_container_width=True,
            height=input_table_height
        )
        
        calc_button = st.button("حساب الشجرة وتوليد النتيجة", type="primary")

        if calc_button:
            with st.spinner('جاري الحساب...'):
                letters = ARABIC_LETTERS if lang == "عربي" else ENGLISH_LETTERS
                tree_results = []
                current_letter_idx = 0
                current_paper = 1
                
                for index, row in edited_df.iterrows():
                    try: attendance = int(row['عدد الحضور'])
                    except: attendance = 0
                        
                    if attendance <= 0:
                        tree_results.append("")
                        continue
                    
                    remaining = attendance
                    committee_tree = []
                    while remaining > 0:
                        if current_letter_idx >= len(letters): break
                        available_in_current = 101 - current_paper
                        letter = letters[current_letter_idx]
                        if remaining <= available_in_current:
                            end_paper = current_paper + remaining - 1
                            if lang == "عربي":
                                committee_tree.append(f"{letter} {end_paper}-{current_paper}")
                            else:
                                committee_tree.append(f"{letter} {current_paper}-{end_paper}")
                            current_paper = end_paper + 1
                            remaining = 0
                            if current_paper > 100:
                                current_letter_idx += 1
                                current_paper = 1
                        else:
                            if lang == "عربي":
                                committee_tree.append(f"{letter} 100-{current_paper}")
                            else:
                                committee_tree.append(f"{letter} {current_paper}-100")
                            remaining -= available_in_current
                            current_letter_idx += 1
                            current_paper = 1
                    tree_results.append(", ".join(committee_tree))
                
                result_df = edited_df.copy()
                col_name = 'الشجرة (عربي)' if lang == "عربي" else 'الشجرة (انجليزي)'
                result_df[col_name] = tree_results
                
                total_attendance = pd.to_numeric(result_df['عدد الحضور'], errors='coerce').fillna(0).sum()
                total_row = pd.DataFrame({'رقم اللجنة': ['الإجمالي'], 'مكان اللجنة': [''], 'عدد الحضور': [int(total_attendance)], col_name: ['']})
                result_df_with_total = pd.concat([result_df, total_row], ignore_index=True)
                
                # ترتيب الأعمدة للمخرجات
                display_cols = ['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور', col_name]
                result_df_display = result_df_with_total[display_cols]
                
                st.markdown(f"<div style='color: green; font-weight: bold; margin-bottom: 10px;'>تم الحساب بنجاح! إجمالي عدد الحضور: {int(total_attendance)} طالب</div>", unsafe_allow_html=True)
                
                output_table_height = (len(result_df_display) + 1) * 38
                st.dataframe(result_df_display, hide_index=True, use_container_width=True, height=output_table_height)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    excel_cols = ['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور', col_name]
                    result_df_with_total[excel_cols].to_excel(writer, index=False, sheet_name='الشجرة')
                    writer.sheets['الشجرة'].sheet_view.rightToLeft = True
                
                st.download_button(
                    label="📥 تحميل النتيجة في ملف إكسيل",
                    data=output.getvalue(),
                    file_name="توزيع_كراسات_الإجابة.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.markdown(
                    "<p style='color: gray; margin-top: 10px; font-size: 14px;'>"
                    "💡 نصيحة: للطباعة كـ PDF، قم بفتح ملف الإكسيل الذي تم تحميله، ثم اضغط (Ctrl+P) واختر Save as PDF"
                    "</p>", 
                    unsafe_allow_html=True
                )

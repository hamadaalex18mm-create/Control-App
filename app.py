import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="برنامج توزيع كراسات الإجابة", layout="wide")

# كود CSS لضبط المحاذاة لليمين وتوسيط العنوان
st.markdown(
    """
    <style>
    /* محاذاة كل النصوص لليمين */
    .stMarkdown, .stText, p, label { text-align: right !important; direction: rtl !important; }
    
    /* ضبط الراديو بوتون ليكون على اليمين تماماً */
    .stRadio { direction: rtl !important; }
    div[role="radiogroup"] { align-items: flex-start !important; width: 100%; }
    
    /* إجبار الأزرار العادية (مثل زر رفع ملف) على التوجه لليمين */
    .stButton > button {
        margin-left: auto !important;
        margin-right: 0 !important;
        display: block !important;
    }
    
    /* ضبط الخلايا داخل الجداول لتكون نصوصها يمين */
    [data-testid="stDataFrame"] div[role="gridcell"], 
    [data-testid="stDataFrame"] div[role="columnheader"] {
        text-align: right !important;
        justify-content: flex-end !important;
    }
    
    /* زرار التحميل */
    .stDownloadButton { display: flex; justify-content: flex-end; }

    /* حماية العنوان الرئيسي ليبقى في المنتصف */
    h1 {
        text-align: center !important;
        width: 100% !important;
        direction: rtl !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- تنسيق الهيدر (الشعارات والعنوان) ---
# تقسيم الشاشة لـ 3 أجزاء: يسار (1)، منتصف (3)، يمين (1)
col_left, col_space, col_right = st.columns([1, 3, 1])

with col_right:
    # 1. شعار الكلية (أعلى اليمين)
    if os.path.exists("logo_faculty.jpg"):
        st.image("logo_faculty.jpg", use_container_width=True)
    elif os.path.exists("logo_faculty.png"):
        st.image("logo_faculty.png", use_container_width=True)

with col_space:
    # العنوان في النص (داخل العمود الأوسط)
    st.markdown(
        """
        <div style='display: flex; justify-content: center; align-items: center; height: 100%; margin-top: 20px;'>
            <h1 style='color: #1E3A8A; font-family: Arial, sans-serif; margin: 0;'>توزيع كراسات الإجابة (الشجرة)</h1>
        </div>
        """,
        unsafe_allow_html=True
    )

with col_left:
    # 2. شعار الوحدة (أعلى اليسار)
    if os.path.exists("logo_unit.jpg"):
        st.image("logo_unit.jpg", use_container_width=True)
    elif os.path.exists("logo_unit.png"):
        st.image("logo_unit.png", use_container_width=True)

st.markdown("---")

# --- الحروف الأبجدية (مؤمنة ضد القطع) ---
ARABIC_LETTERS = "أ ب ج د هـ و ز ح ط ي ك ل م ن س ع ف ص ق ر ش ت ث خ ذ ض ظ غ".split(" ")
ENGLISH_LETTERS = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z".split(" ")

if 'base_df' not in st.session_state:
    st.session_state.base_df = None

# --- مرحلة رفع الملف ---
if st.session_state.base_df is None:
    st.markdown(
        "<div dir='rtl' style='text-align: right; background-color: #e8f4f8; padding: 15px; border-radius: 5px; color: #004d40; border: 1px solid #b6e3f4; margin-bottom: 15px; font-size: 16px;'>"
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
    col_data, col_settings = st.columns([4, 1])
    
    with col_settings:
        st.markdown("<h3 dir='rtl' style='text-align: right;'>الإعدادات</h3>", unsafe_allow_html=True)
        lang = st.radio("اختر لغة الشجرة:", ["عربي", "إنجليزي"])
        st.markdown("---")
        if st.button("رفع ملف جديد"):
            st.session_state.base_df = None
            st.rerun()

    with col_data:
        st.markdown("<h3 dir='rtl' style='text-align: right;'>إدخال أعداد الحضور</h3>", unsafe_allow_html=True)
        
        # الترتيب ده بيخلي "رقم اللجنة" أقصى اليمين في جدول المدخلات
        input_display_cols = ['عدد الحضور', 'مكان اللجنة', 'رقم اللجنة']
        df_for_editor = st.session_state.base_df[input_display_cols]
        
        input_table_height = (len(st.session_state.base_df) + 1) * 38
        
        edited_df = st.data_editor(
            df_for_editor,
            disabled=["رقم اللجنة", "مكان اللجنة"],
            hide_index=True,
            use_container_width=True,
            height=input_table_height
        )
        
        st.markdown("<div dir='rtl' style='text-align: right;'>", unsafe_allow_html=True)
        calc_button = st.button("حساب الشجرة وتوليد النتيجة", type="primary")
        st.markdown("</div>", unsafe_allow_html=True)

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

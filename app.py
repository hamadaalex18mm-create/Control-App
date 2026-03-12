import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="برنامج توزيع كراسات الإجابة", layout="wide")

# كود CSS احترافي لضبط اتجاه الجدول وكل المكونات لليمين
st.markdown(
    """
    <style>
    /* ضبط اتجاه الصفحة بالكامل لليمين */
    .main { direction: rtl; }
    div[data-testid="stBlock"] { direction: rtl; }
    
    /* محاذاة نصوص الجدول لليمين */
    [data-testid="stTableSummary"] { direction: rtl; text-align: right; }
    th { text-align: right !important; }
    td { text-align: right !important; }
    
    /* جعل العناوين والنصوص تتبع الاتجاه العربي */
    h1, h2, h3, p, span { text-align: right; direction: rtl; }
    
    /* ضبط اتجاه جدول البيانات */
    .stDataFrame { direction: rtl; }
    </style>
    """,
    unsafe_allow_html=True
)

# --- تنسيق الهيدر (الشعارات والعنوان) ---
col_left, col_space, col_right = st.columns([1, 3, 1])

if os.path.exists("logo_unit.jpg"):
    col_left.image("logo_unit.jpg", use_container_width=True)
elif os.path.exists("logo_unit.png"):
    col_left.image("logo_unit.png", use_container_width=True)

if os.path.exists("logo_faculty.jpg"):
    col_right.image("logo_faculty.jpg", use_container_width=True)
elif os.path.exists("logo_faculty.png"):
    col_right.image("logo_faculty.png", use_container_width=True)

st.markdown(
    """
    <div style='text-align: center; margin-top: -50px; margin-bottom: 30px;'>
        <h1 style='color: #1E3A8A; font-family: Arial, sans-serif;'>برنامج توزيع كراسات الإجابة (الشجرة)</h1>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("---")

# --- الحروف الأبجدية ---
ARABIC_LETTERS = ["أ", "ب", "ج", "د", "هـ", "و", "ز", "ح", "ط", "ي", "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق", "ر", "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ"]
ENGLISH_LETTERS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

if 'base_df' not in st.session_state:
    st.session_state.base_df = None

# 1. رفع الملف
if st.session_state.base_df is None:
    st.markdown("<h3 style='text-align: right;'>يرجى رفع ملف الإكسيل الأساسي</h3>", unsafe_allow_html=True)
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
                st.error("تأكد من وجود أعمدة باسم 'رقم اللجنة' و 'مكان اللجنة'")
        except:
            st.error("خطأ في قراءة الملف")

# 2. واجهة العمل
if st.session_state.base_df is not None:
    m_col1, m_col2 = st.columns([1, 4])
    
    with m_col1:
        st.markdown("<h3 style='text-align: right;'>الإعدادات</h3>", unsafe_allow_html=True)
        lang = st.radio("لغة الشجرة:", ["عربي", "إنجليزي"])
        if st.button("رفع ملف جديد"):
            st.session_state.base_df = None
            st.rerun()

    with m_col2:
        st.markdown("<h3 style='text-align: right;'>إدخال أعداد الحضور</h3>", unsafe_allow_html=True)
        
        # ترتيب الأعمدة في الإدخال: رقم اللجنة يمين
        input_cols = ['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور']
        input_height = (len(st.session_state.base_df) + 1) * 38
        
        edited_df = st.data_editor(
            st.session_state.base_df[input_cols],
            disabled=["رقم اللجنة", "مكان اللجنة"],
            hide_index=True,
            use_container_width=True,
            height=input_height
        )
        
        if st.button("حساب الشجرة", type="primary"):
            letters = ARABIC_LETTERS if lang == "عربي" else ENGLISH_LETTERS
            tree_results = []
            curr_letter_idx = 0
            curr_paper = 1
            
            for index, row in edited_df.iterrows():
                try: attendance = int(row['عدد الحضور'])
                except: attendance = 0
                
                if attendance <= 0:
                    tree_results.append("")
                    continue
                
                remaining = attendance
                committee_tree = []
                while remaining > 0:
                    if curr_letter_idx >= len(letters): break
                    space = 101 - curr_paper
                    letter = letters[curr_letter_idx]
                    
                    if remaining <= space:
                        end = curr_paper + remaining - 1
                        if lang == "عربي": committee_tree.append(f"{letter} {end}-{curr_paper}")
                        else: committee_tree.append(f"{letter} {curr_paper}-{end}")
                        curr_paper = end + 1
                        remaining = 0
                        if curr_paper > 100: curr_letter_idx += 1; curr_paper = 1
                    else:
                        if lang == "عربي": committee_tree.append(f"{letter} 100-{curr_paper}")
                        else: committee_tree.append(f"{letter} {curr_paper}-100")
                        remaining -= space
                        curr_letter_idx += 1
                        curr_paper = 1
                tree_results.append(", ".join(committee_tree))
            
            # تجهيز النتيجة
            result_df = edited_df.copy()
            tree_col = 'الشجرة (عربي)' if lang == "عربي" else 'الشجرة (إنجليزي)'
            result_df[tree_col] = tree_results
            
            # حساب الإجمالي
            total = pd.to_numeric(result_df['عدد الحضور'], errors='coerce').sum()
            total_row = pd.DataFrame({'رقم اللجنة': ['الإجمالي'], 'مكان اللجنة': [''], 'عدد الحضور': [int(total)], tree_col: ['']})
            final_df = pd.concat([result_df, total_row], ignore_index=True)
            
            # الترتيب في العرض: رقم اللجنة أول واحد من اليمين
            display_order = ['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور', tree_col]
            
            st.markdown(f"<p dir='rtl' style='color: green; font-weight: bold;'>إجمالي الحضور: {int(total)}</p>", unsafe_allow_html=True)
            
            # عرض الجدول مع ستايل المحاذاة لليمين
            st.dataframe(final_df[display_order].style.set_properties(**{'text-align': 'right'}), hide_index=True, use_container_width=True, height=(len(final_df)+1)*38)
            
            # تحميل إكسيل

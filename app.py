import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="توزيع كراسات الإجابة", layout="wide")

# ==========================================
# كود الـ CSS لإجبار كل عناصر التطبيق لليمين
# ==========================================
st.markdown(
    """
    <style>
    .stMarkdown, .stText { text-align: right !important; direction: rtl !important; }
    .stAlert { direction: rtl !important; text-align: right !important; }
    .stRadio { direction: rtl !important; text-align: right !important; }
    [data-testid="stDataFrame"] div[role="gridcell"] {
        justify-content: flex-end !important;
        text-align: right !important;
    }
    [data-testid="stDataFrame"] div[role="columnheader"] {
        justify-content: flex-end !important;
    }
    .stButton > button, .stDownloadButton > button {
        margin-left: auto !important;
        margin-right: 0 !important;
        direction: rtl !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)
# ==========================================

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
        <h1 style='color: #1E3A8A; font-family: Arial, sans-serif;'>توزيع كراسات الإجابة (الشجرة)</h1>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("---")

# --- باقي منطق البرنامج مقسم لسطور قصيرة ---
ARABIC_LETTERS = [
    "أ", "ب", "ج", "د", "هـ", "و", "ز", "ح", "ط", "ي", "ك", "ل", "م", 
    "ن", "س", "ع", "ف", "ص", "ق", "ر", "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ"
]

ENGLISH_LETTERS = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", 
    "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
]

if 'base_df' not in st.session_state:
    st.session_state.base_df = None

if st.session_state.base_df is None:
    st.info("يرجى رفع ملف الإكسيل الأساسي (يحتوي على: رقم اللجنة، مكان اللجنة)")
    uploaded_file = st.file_uploader("اختر ملف الإكسيل", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
            
            if "رقم اللجنة" in df.columns and "مكان اللجنة" in df.columns:
                st.session_state.base_df = df[['رقم اللجنة', 'مكان اللجنة']].copy()
                st.session_state.base_df['عدد الح

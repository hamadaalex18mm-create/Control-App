import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="برنامج توزيع كراسات الإجابة", layout="wide")

# ==========================================
# كود الـ CSS النهائي لإجبار كل عناصر التطبيق لليمين والعنوان في النص
# ==========================================
st.markdown(
    """
    <style>
    /* محاذاة كل النصوص لليمين (باستثناء العنوان الرئيسي) */
    .stMarkdown, .stText, p, label, .stDataFrame td, .stDataFrame th, .stRadio { text-align: right !important; direction: rtl !important; }
    
    /* ضبط الراديو بوتون (اختيار اللغة) ليكون على اليمين تماماً */
    div[role="radiogroup"] { align-items: flex-start; direction: rtl; }
    
    /* ضبط الخلايا داخل الجداول لتكون نصوصها يمين */
    [data-testid="stDataFrame"] div[role="gridcell"], 
    [data-testid="stDataFrame"] div[role="columnheader"] {
        text-align: right !important;
        justify-content: flex-end !important;
    }
    
    /* زرار التحميل */
    .stDownloadButton { display: flex; justify-content: flex-end; }

    /* التعديل السحري للعنوان الرئيسي: إجبار h1 على التوسط في منتصف الصفحة */
    h1 {
        text-align: center !important;
        width: 100% !important;
        direction: rtl !important; /* للحفاظ على تنسيق العربي داخل التوسيط */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- تنسيق الهيدر (الشعارات والعنوان) ---
# العمود الأول (يسار)، والعمود التالت (يمين) للشعارات، والعنوان في النص (col_space)
col_left, col_space, col_right = st.columns([1, 3, 1])

# --- تعديل: تبديل مكان الصورتين ---

# 1. شعار الكلية (أعلى اليمين)
if os.path.exists("logo_faculty.jpg"):
    col_right.image("logo_faculty.jpg", use_container_width=True)
elif os.path.exists("logo_faculty.png"):
    col_right.image("logo_faculty.png", use_container_width=True)

# العنوان في النص (تم برمجته في CSS ليكون في منتصف الصفحة)
st.markdown(
    """
    <div style='text-align: center; margin-top: -50px; margin-bottom: 30px;'>
        <h1 style='color: #1E3A8A; font-family: Arial, sans-serif;'>برنامج توزيع كراسات الإجابة (الشجرة)</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# 2. شعار الوحدة (أعلى اليسار)
if os.path.exists("logo_unit.jpg"):
    col_left.image("logo_unit.jpg", use_container_width=True)
elif os.path.exists("logo_unit.png"):
    col_left.image("logo_unit.png", use_container_width=True)

st.markdown("---")

# --- باقي منطق البرنامج ---
ARABIC_LETTERS = ["أ", "

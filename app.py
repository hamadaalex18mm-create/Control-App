import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="برنامج توزيع كراسات الإجابة", layout="wide")

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

# --- باقي منطق البرنامج ---
ARABIC_LETTERS = ["أ", "ب", "ج", "د", "هـ", "و", "ز", "ح", "ط", "ي", "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق", "ر", "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ"]
ENGLISH_LETTERS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

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
                st.session_state.base_df['عدد الحضور'] = 0
                st.success("تم رفع الملف بنجاح! جاري التجهيز...")
                st.rerun()
            else:
                st.error("الملف يجب أن يحتوي على عمودين باسم: 'رقم اللجنة' و 'مكان اللجنة'")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

if st.session_state.base_df is not None:
    m_col1, m_col2 = st.columns([1, 4])
    
    with m_col1:
        st.markdown("<h3 dir='rtl' style='text-align: right;'>الإعدادات</h3>", unsafe_allow_html=True)
        lang = st.radio("اختر لغة الشجرة:", ["عربي", "إنجليزي"])
        st.markdown("---")
        if st.button("تفريغ البيانات لرفع ملف جديد"):
            st.session_state.base_df = None
            st.rerun()

    with m_col2:
        st.markdown("<h3 dir='rtl' style='text-align: right;'>إدخال أعداد الحضور</h3>", unsafe_allow_html=True)
        st.markdown("<p dir='rtl' style='text-align: right; color: gray;'>نصيحة: للطباعة كـ PDF، قم بتحميل ملف الإكسيل وافتحه، ثم اضغط (Ctrl+P) واختر Save as PDF</p>", unsafe_allow_html=True)
        
        # جدول الإدخال
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
                    try:
                        attendance = int(row['عدد الحضور'])
                    except (ValueError, TypeError):
                        attendance = 0
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
                
                # ترتيب الأعمدة الطبيعي للإكسيل والـ HTML
                natural_order = ['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور', col_name]
                final_df = result_df[natural_order]
                final_df_with_total = pd.concat([final_df, total_row[natural_order]], ignore_index=True)
                
                st.markdown(f"<div dir='rtl' style='text-align: right; color: green; font-weight: bold; margin-bottom: 15px; margin-top: 20px;'>تم الحساب بنجاح! إجمالي عدد الحضور: {int(total_attendance)} طالب</div>", unsafe_allow_html=True)
                
                # ==========================================
                # بناء جدول HTML مخصص للتغلب على مشاكل Streamlit
                # ==========================================
                html_table = final_df_with_total.to_html(index=False, classes="custom-rtl-table")
                custom_css = """
                <style>
                    .custom-rtl-table {
                        width: 100%;
                        border-collapse: collapse;
                        direction: rtl; /* دي اللي بتعكس الجدول وتخلي رقم اللجنة على اليمين */
                        text-align: right;
                        font-family: Arial, sans-serif;
                        margin-bottom: 20px;
                        background-color: white;
                    }
                    .custom-rtl-table th, .custom-rtl-table td {
                        border: 1px solid #e2e8f0;
                        padding: 12px;
                        text-align: right !important; /* إجبار المحاذاة لليمين */
                    }
                    .custom-rtl-table th {
                        background-color: #f8fafc;
                        color: #1e293b;
                        font-weight: bold;
                        border-bottom: 2px solid #cbd5e1;
                    }
                    .custom-rtl-table tr:hover {
                        background-color: #f1f5f9;
                    }
                </style>
                """
                st.markdown(custom_css + html_table, unsafe_allow_html=True)
                # ==========================================
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df_with_total.to_excel(writer, index=False, sheet_name='الشجرة')
                    worksheet = writer.sheets['الشجرة']
                    worksheet.sheet_view.rightToLeft = True
                
                st.download_button(label="📥 تحميل النتيجة في ملف إكسيل", data=output.getvalue(), file_name="توزيع_كراسات_الإجابة.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="برنامج توزيع كراسات الإجابة", layout="wide")

# --- تنسيق الهيدر (الشعارات والعنوان) ---
# ترتيب الأعمدة في Streamlit بيكون من اليسار لليمين
# العمود الأول (يسار) = شعار الوحدة ، العمود الثالث (يمين) = شعار الكلية
col_left, col_space, col_right = st.columns([1, 3, 1])

# شعار الوحدة (أعلى اليسار) 
if os.path.exists("logo_unit.jpg"):
    col_left.image("logo_unit.jpg", use_container_width=True)
elif os.path.exists("logo_unit.png"):
    col_left.image("logo_unit.png", use_container_width=True)

# شعار الكلية (أعلى اليمين) 
if os.path.exists("logo_faculty.jpg"):
    col_right.image("logo_faculty.jpg", use_container_width=True)
elif os.path.exists("logo_faculty.png"):
    col_right.image("logo_faculty.png", use_container_width=True)

# العنوان في المنتصف
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
        st.markdown("### الإعدادات")
        lang = st.radio("اختر لغة الشجرة:", ["عربي", "إنجليزي"])
        st.markdown("---")
        if st.button("تفريغ البيانات لرفع ملف جديد"):
            st.session_state.base_df = None
            st.rerun()

    with m_col2:
        st.markdown("### إدخال أعداد الحضور")
        st.caption("نصيحة: للطباعة كـ PDF، قم بتحميل ملف الإكسيل وافتحه واضغط (Ctrl+P) ثم اختر Save as PDF")
        
        edited_df = st.data_editor(
            st.session_state.base_df,
            disabled=["رقم اللجنة", "مكان اللجنة"],
            hide_index=True,
            use_container_width=True
        )
        
        if st.button("حساب الشجرة وتوليد النتيجة", type="primary"):
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
                        if current_letter_idx >= len(letters):
                            st.warning("تم تجاوز عدد الحروف المتاحة في الأبجدية!")
                            break

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
                    
                    tree_results.append(" - ".join(committee_tree))
                
                result_df = edited_df.copy()
                col_name = 'الشجرة (عربي)' if lang == "عربي" else 'الشجرة (انجليزي)'
                result_df.insert(0, col_name, tree_results)
                
                st.success("تم الحساب بنجاح!")
                st.dataframe(result_df, hide_index=True, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='الشجرة')
                
                st.download_button(
                    label="📥 تحميل النتيجة في ملف إكسيل",
                    data=output.getvalue(),
                    file_name="توزيع_كراسات_الإجابة.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
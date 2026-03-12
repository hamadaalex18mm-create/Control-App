import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="توزيع كراسات الإجابة", layout="wide")

# ==========================================
# ستايل الواجهة (Streamlit UI)
# ==========================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
    * { font-family: 'Tajawal', sans-serif !important; }
    .stApp, .stMarkdown, .stText, p, label, input, .stSelectbox, .stDateInput { 
        text-align: right !important; direction: rtl !important; font-size: 16px !important;
    }
    h1 { text-align: center !important; color: #1E3A8A !important; font-weight: 700 !important; }
    h3 { 
        color: #004d40 !important; font-weight: 700 !important; padding-top: 15px; 
        border-bottom: 2px solid #b6e3f4; padding-bottom: 10px; margin-bottom: 20px; 
        text-align: right !important; direction: rtl !important;
    }
    .stButton > button { font-size: 18px !important; font-weight: bold !important; width: 100% !important; }
    </style>
    """,
    unsafe_allow_html=True
)

# --- الهيدر (الشعارات والعنوان) ---
col_left, col_space, col_right = st.columns([1, 3, 1])
with col_left:
    if os.path.exists("logo_unit.png"): st.image("logo_unit.png", use_container_width=True)
with col_space:
    st.markdown("<h1 style='margin-top: 20px;'>نظام توزيع كراسات الإجابة</h1>", unsafe_allow_html=True)
with col_right:
    if os.path.exists("logo_faculty.png"): st.image("logo_faculty.png", use_container_width=True)

st.markdown("---")

ARABIC_LETTERS = "أ ب ج د هـ و ز ح ط ي ك ل م ن س ع ف ص ق ر ش ت ث خ ذ ض ظ غ".split(" ")
ENGLISH_LETTERS = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z".split(" ")

if 'base_df' not in st.session_state: st.session_state.base_df = None

# --- رفع الملف ---
if st.session_state.base_df is None:
    uploaded_file = st.file_uploader("ارفع ملف الإكسيل الأساسي (رقم اللجنة، مكان اللجنة)", type=['xlsx'])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        if "رقم اللجنة" in df.columns and "مكان اللجنة" in df.columns:
            st.session_state.base_df = df[['رقم اللجنة', 'مكان اللجنة']].copy()
            st.session_state.base_df['عدد الحضور'] = 0
            st.rerun()

# --- واجهة الإدخال ---
if st.session_state.base_df is not None:
    st.markdown("<h3>البيانات الأساسية</h3>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        control_name = st.text_input("اسم الكنترول:", value="كنترول الفرقة الأولى")
        exam_date = st.text_input("تاريخ الامتحان:", value="الخميس 12-03-2026")
    with c2:
        course_name = st.text_input("اسم المقرر:", value="محاسبة دولية")
        lang = st.selectbox("لغة الشجرة:", ["عربي", "إنجليزي"])
        
    st.markdown("<h3>أعداد الحضور</h3>", unsafe_allow_html=True)
    edited_df = st.data_editor(
        st.session_state.base_df[['عدد الحضور', 'مكان اللجنة', 'رقم اللجنة']],
        disabled=["رقم اللجنة", "مكان اللجنة"],
        hide_index=True, use_container_width=True
    )
    
    if st.button("توليد ملف الإكسيل المنسق", type="primary"):
        # الحسابات
        letters = ARABIC_LETTERS if lang == "عربي" else ENGLISH_LETTERS
        tree_results = []
        curr_let_idx, curr_paper = 0, 1
        
        for _, row in edited_df.iterrows():
            attendance = int(row['عدد الحضور'])
            if attendance <= 0:
                tree_results.append(""); continue
            
            rem = attendance
            comm_tree = []
            while rem > 0 and curr_let_idx < len(letters):
                avail = 101 - curr_paper
                let = letters[curr_let_idx]
                take = min(rem, avail)
                end = curr_paper + take - 1
                
                # تنسيق النص بناءً على اللغة
                text = f"{let} {end}-{curr_paper}" if lang == "عربي" else f"{let} {curr_paper}-{end}"
                comm_tree.append(text)
                
                rem -= take
                curr_paper = end + 1
                if curr_paper > 100:
                    curr_let_idx += 1; curr_paper = 1
            tree_results.append(", ".join(comm_tree))

        # تحضير البيانات النهائية
        final_df = edited_df.copy()
        col_tree = "توزيع الشجرة"
        final_df[col_tree] = tree_results
        final_df = final_df[['رقم اللجنة', 'مكان اللجنة', 'عدد الحضور', col_tree]]
        
        # --- إنشاء ملف الإكسيل المنسق ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=5)
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            worksheet.sheet_view.rightToLeft = True
            
            from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
            
            # تنسيق الهيدر العلوي (البيانات الأساسية)
            header_font = Font(name='Tajawal', size=14, bold=True)
            info_style = Alignment(horizontal='right', vertical='center')
            
            worksheet['D1'] = f"الكنترول: {control_name}"
            worksheet['D1'].font = header_font
            worksheet['D1'].alignment = info_style
            
            worksheet['A1'] = f"المقرر: {course_name}"
            worksheet['A1'].font = header_font
            worksheet['A1'].alignment = Alignment(horizontal='left')
            
            worksheet['D2'] = f"التاريخ: {exam_date}"
            worksheet['D2'].font = header_font
            worksheet['D2'].alignment = info_style

            # تنسيق رأس الجدول
            header_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            for cell in worksheet[6]: # الصف رقم 6 هو رأس الجدول
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.fill = header_fill
                cell.border = thin_border
            
            # تنسيق محتوى الجدول وتوسيطه
            for row in worksheet.iter_rows(min_row=7, max_row=worksheet.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border

            # ضبط عرض الأعمدة تلقائياً
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 25
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 50

        st.success("تم تجهيز الملف بتنسيق الطباعة!")
        st.download_button(
            "📥 تحميل ملف الإكسيل المنسق للطباعة",
            output.getvalue(),
            "كشف_توزيع_الشجرة.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

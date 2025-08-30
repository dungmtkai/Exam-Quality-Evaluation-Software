import streamlit as st
import pandas as pd
from processor import calculate_question_stats
from io import BytesIO

from processor import evaluate_exam_difficulty_mix   # <- bạn import hàm đánh giá đã viết

st.set_page_config(page_title="Tính độ khó câu hỏi", layout="wide")

st.title("📊 Công cụ tính độ khó câu hỏi từ file Excel")

# Upload file Excel
uploaded_file = st.file_uploader("📁 Tải lên file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df_input = pd.read_excel(uploaded_file)
        # Chuyển đổi object về string để tránh lỗi serialization
        for col in df_input.columns:
            if df_input[col].dtype == 'object':
                df_input[col] = df_input[col].astype(str)
        st.dataframe(df_input, use_container_width=True)
        st.success("✅ File đã được tải lên thành công!")

        # Tính toán độ khó từng câu
        result_df = calculate_question_stats(df_input)

        st.subheader("📋 Kết quả tính độ khó từng câu:")
        st.dataframe(result_df, use_container_width=True)

        # ---- 🔹 ĐÁNH GIÁ ĐỀ THI (thêm mới) ----
        st.subheader("📊 Đánh giá tổng quan đề thi:")

        summary_df, conclusion, disc_info = evaluate_exam_difficulty_mix(
            result_df,
            tolerance=0.05,
            check_discrimination=True  # có thể bật/tắt
        )

        st.write("### 🔎 Cơ cấu độ khó so với mục tiêu")
        st.dataframe(summary_df, use_container_width=True)

        st.markdown(f"### ✅ Kết luận: **{conclusion}**")

        if disc_info:
            st.write("### 📐 Thống kê độ phân biệt")
            st.json(disc_info)

        # ---- Xuất file Excel ----
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Độ khó')
            processed_data = output.getvalue()
            return processed_data

        excel_data = convert_df_to_excel(result_df)
        st.download_button(
            label="⬇️ Tải kết quả về (.xlsx)",
            data=excel_data,
            file_name="do_kho_cau_hoi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Đã xảy ra lỗi: {e}")
else:
    st.info("📌 Vui lòng tải lên file Excel có chứa các cột Câu hỏi để bắt đầu.")

import streamlit as st
import pandas as pd
from processor import calculate_difficulty_from_df
from io import BytesIO

st.set_page_config(page_title="Tính độ khó câu hỏi", layout="wide")

st.title("📊 Công cụ tính độ khó câu hỏi từ file Excel")

# Upload file Excel
uploaded_file = st.file_uploader("📁 Tải lên file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df_input = pd.read_excel(uploaded_file)
        st.dataframe(df_input, use_container_width=True)
        st.success("✅ File đã được tải lên thành công!")

        # Tính toán độ khó
        result_df = calculate_difficulty_from_df(df_input)

        st.subheader("📋 Kết quả tính độ khó:")
        st.dataframe(result_df, use_container_width=True)

        # Tải về kết quả
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
    st.info("📌 Vui lòng tải lên file Excel có chứa các cột Câu 1 đến Câu 40 để bắt đầu.")

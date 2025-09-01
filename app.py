import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import các hàm xử lý từ các file riêng biệt
from processor_multiple_choice import calculate_question_stats
from processor_essay import calculate_essay_stats
from processor_common import evaluate_exam_difficulty_mix

st.set_page_config(page_title="Tính độ khó câu hỏi", layout="wide")

st.title("📊 Công cụ tính độ khó câu hỏi từ file Excel")

# Chọn hình thức đề thi
exam_type = st.selectbox(
    "📝 Chọn hình thức đề thi:",
    ["Trắc nghiệm", "Tự luận"],
    help="Chọn loại đề thi để áp dụng phương pháp tính phù hợp"
)

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
        
        # Hiển thị hình thức đề thi đã chọn
        st.info(f"📝 Hình thức đề thi: **{exam_type}**")

        # Xử lý theo hình thức đề thi
        if exam_type == "Trắc nghiệm":
            # Tính toán độ khó từng câu cho trắc nghiệm
            result_df = calculate_question_stats(df_input)

            st.subheader("📋 Kết quả tính độ khó từng câu (Trắc nghiệm):")
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

            # ---- Xuất file Word ----
            def convert_to_word(result_df, summary_df, conclusion, disc_info):
                doc = Document()
                
                # Tiêu đề
                title = doc.add_heading('BÁO CÁO ĐÁNH GIÁ ĐỘ KHÓ ĐỀ THI TRẮC NGHIỆM', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Kết quả từng câu
                doc.add_heading('1. Kết quả tính độ khó từng câu', level=1)
                
                # Thêm bảng kết quả
                table = doc.add_table(rows=1, cols=len(result_df.columns))
                table.style = 'Table Grid'
                
                # Header
                header_cells = table.rows[0].cells
                for i, col in enumerate(result_df.columns):
                    header_cells[i].text = str(col)
                
                # Data
                for _, row in result_df.iterrows():
                    row_cells = table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
                
                doc.add_page_break()
                
                # Đánh giá tổng quan
                doc.add_heading('2. Đánh giá tổng quan đề thi', level=1)
                
                # Cơ cấu độ khó
                doc.add_heading('2.1. Cơ cấu độ khó so với mục tiêu', level=2)
                
                # Thêm bảng summary
                summary_table = doc.add_table(rows=1, cols=len(summary_df.columns))
                summary_table.style = 'Table Grid'
                
                # Header
                header_cells = summary_table.rows[0].cells
                for i, col in enumerate(summary_df.columns):
                    header_cells[i].text = str(col)
                
                # Data
                for _, row in summary_df.iterrows():
                    row_cells = summary_table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
                
                doc.add_paragraph()
                
                # Thống kê độ phân biệt
                if disc_info:
                    doc.add_heading('2.2. Thống kê độ phân biệt', level=2)
                    for key, value in disc_info.items():
                        doc.add_paragraph(f'{key}: {value}')
                
                doc.add_paragraph()
                
                # Kết luận
                doc.add_heading('3. Kết luận', level=1)
                conclusion_para = doc.add_paragraph(conclusion)
                conclusion_para.runs[0].bold = True
                
                # Lưu vào BytesIO
                output = BytesIO()
                doc.save(output)
                output.seek(0)
                return output.getvalue()
            
            word_data = convert_to_word(result_df, summary_df, conclusion, disc_info)
            st.download_button(
                label="⬇️ Tải báo cáo Word (.docx)",
                data=word_data,
                file_name="bao_cao_do_kho_trac_nghiem.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        else:  # Tự luận
            st.subheader("📋 Xử lý đề thi tự luận")
            
            # Đọc sheet 2 nếu có (chứa điểm tối đa)
            max_scores_df = None
            try:
                # Đọc tất cả sheets
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                
                if len(sheet_names) >= 2:
                    # Đọc sheet 2 (index 1)
                    max_scores_df = pd.read_excel(uploaded_file, sheet_name=sheet_names[1])
                    st.info(f"📊 Đã tìm thấy sheet điểm tối đa: {sheet_names[1]}")
                    
                    # Hiển thị điểm tối đa
                    with st.expander("📋 Xem điểm tối đa từng câu"):
                        st.dataframe(max_scores_df, use_container_width=True)
                else:
                    st.warning("⚠️ Không tìm thấy sheet thứ 2 chứa điểm tối đa. Sẽ sử dụng điểm cao nhất thực tế.")
            except Exception as e:
                st.warning(f"⚠️ Không thể đọc sheet 2: {e}")
            
            # Tính toán độ khó từng câu cho tự luận
            result_df = calculate_essay_stats(df_input, max_scores_df)
            
            st.subheader("📋 Kết quả tính độ khó từng câu (Tự luận):")
            st.dataframe(result_df, use_container_width=True)
            
            # Giải thích cách tính
            with st.expander("📚 Giải thích cách tính cho tự luận"):
                st.markdown("""
                **Độ khó (P)**: 
                - Công thức: `P = (Điểm TB của tất cả SV / Điểm tối đa) × 100`
                - Điểm tối đa lấy từ sheet 2 hoặc điểm cao nhất thực tế
                
                **Độ phân biệt (D)**:
                - Công thức: `D = (Điểm TB nhóm cao - Điểm TB nhóm thấp) / Điểm tối đa`
                - D ≥ 0.4: Rất tốt
                - 0.3 ≤ D < 0.4: Tốt  
                - 0.2 ≤ D < 0.3: Trung bình
                - D < 0.2: Kém
                """)
            
            # ---- 🔹 ĐÁNH GIÁ ĐỀ THI ----
            st.subheader("📊 Đánh giá tổng quan đề thi:")
            
            summary_df, conclusion, disc_info = evaluate_exam_difficulty_mix(
                result_df,
                tolerance=0.05,
                check_discrimination=True
            )
            
            st.write("### 🔎 Cơ cấu độ khó so với mục tiêu")
            st.dataframe(summary_df, use_container_width=True)
            
            st.markdown(f"### ✅ Kết luận: **{conclusion}**")
            
            if disc_info:
                st.write("### 📐 Thống kê độ phân biệt")
                st.json(disc_info)
            
            # ---- Xuất file Word ----
            def convert_to_word(result_df, summary_df, conclusion, disc_info, max_scores_df=None):
                doc = Document()
                
                # Tiêu đề
                title = doc.add_heading('BÁO CÁO ĐÁNH GIÁ ĐỘ KHÓ ĐỀ THI TỰ LUẬN', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Kết quả từng câu
                doc.add_heading('1. Kết quả tính độ khó từng câu', level=1)
                
                # Thêm bảng kết quả
                table = doc.add_table(rows=1, cols=len(result_df.columns))
                table.style = 'Table Grid'
                
                # Header
                header_cells = table.rows[0].cells
                for i, col in enumerate(result_df.columns):
                    header_cells[i].text = str(col)
                
                # Data
                for _, row in result_df.iterrows():
                    row_cells = table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
                
                # Điểm tối đa nếu có
                if max_scores_df is not None:
                    doc.add_paragraph()
                    doc.add_heading('1.1. Điểm tối đa từng câu', level=2)
                    max_table = doc.add_table(rows=1, cols=len(max_scores_df.columns))
                    max_table.style = 'Table Grid'
                    
                    # Header
                    header_cells = max_table.rows[0].cells
                    for i, col in enumerate(max_scores_df.columns):
                        header_cells[i].text = str(col)
                    
                    # Data
                    for _, row in max_scores_df.iterrows():
                        row_cells = max_table.add_row().cells
                        for i, val in enumerate(row):
                            row_cells[i].text = str(val)
                
                doc.add_page_break()
                
                # Đánh giá tổng quan
                doc.add_heading('2. Đánh giá tổng quan đề thi', level=1)
                
                # Cơ cấu độ khó
                doc.add_heading('2.1. Cơ cấu độ khó so với mục tiêu', level=2)
                
                # Thêm bảng summary
                summary_table = doc.add_table(rows=1, cols=len(summary_df.columns))
                summary_table.style = 'Table Grid'
                
                # Header
                header_cells = summary_table.rows[0].cells
                for i, col in enumerate(summary_df.columns):
                    header_cells[i].text = str(col)
                
                # Data
                for _, row in summary_df.iterrows():
                    row_cells = summary_table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
                
                doc.add_paragraph()
                
                # Thống kê độ phân biệt
                if disc_info:
                    doc.add_heading('2.2. Thống kê độ phân biệt', level=2)
                    for key, value in disc_info.items():
                        doc.add_paragraph(f'{key}: {value}')
                
                doc.add_paragraph()
                
                # Giải thích cách tính
                doc.add_heading('2.3. Giải thích cách tính', level=2)
                doc.add_paragraph('Độ khó (P):')
                doc.add_paragraph('• Công thức: P = (Điểm TB của tất cả SV / Điểm tối đa) × 100', style='List Bullet')
                doc.add_paragraph('• Điểm tối đa lấy từ sheet 2 hoặc điểm cao nhất thực tế', style='List Bullet')
                
                doc.add_paragraph('Độ phân biệt (D):')
                doc.add_paragraph('• Công thức: D = (Điểm TB nhóm cao - Điểm TB nhóm thấp) / Điểm tối đa', style='List Bullet')
                doc.add_paragraph('• D ≥ 0.4: Rất tốt', style='List Bullet')
                doc.add_paragraph('• 0.3 ≤ D < 0.4: Tốt', style='List Bullet')
                doc.add_paragraph('• 0.2 ≤ D < 0.3: Trung bình', style='List Bullet')
                doc.add_paragraph('• D < 0.2: Kém', style='List Bullet')
                
                doc.add_paragraph()
                
                # Kết luận
                doc.add_heading('3. Kết luận', level=1)
                conclusion_para = doc.add_paragraph(conclusion)
                conclusion_para.runs[0].bold = True
                
                # Lưu vào BytesIO
                output = BytesIO()
                doc.save(output)
                output.seek(0)
                return output.getvalue()
            
            word_data = convert_to_word(result_df, summary_df, conclusion, disc_info, max_scores_df)
            st.download_button(
                label="⬇️ Tải báo cáo Word (.docx)",
                data=word_data,
                file_name="bao_cao_do_kho_tu_luan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"❌ Đã xảy ra lỗi: {e}")
else:
    st.info("📌 Vui lòng tải lên file Excel có chứa các cột Câu hỏi để bắt đầu.")

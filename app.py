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
    ["Trắc nghiệm", "Tự luận", "Hỗn hợp"],
    help="Chọn loại đề thi để áp dụng phương pháp tính phù hợp"
)

# Upload file Excel
uploaded_file = st.file_uploader("📁 Tải lên file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Hiển thị thông tin file và loại đề thi
        st.success("✅ File đã được tải lên thành công!")
        st.info(f"📝 Hình thức đề thi: **{exam_type}**")
        
        # Hiển thị dữ liệu theo loại đề thi
        if exam_type == "Trắc nghiệm":
            # Đọc và hiển thị dữ liệu trắc nghiệm
            df_input = pd.read_excel(uploaded_file)
            # Chuyển đổi object về string để tránh lỗi serialization
            for col in df_input.columns:
                if df_input[col].dtype == 'object':
                    df_input[col] = df_input[col].astype(str)
            
            st.subheader("📊 Dữ liệu đã tải lên:")
            st.dataframe(df_input, use_container_width=True)
            
        elif exam_type == "Tự luận":
            # Đọc tất cả sheets cho tự luận
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            df_input = pd.read_excel(uploaded_file, sheet_name=0)
            # Chuyển đổi object về string để tránh lỗi serialization
            for col in df_input.columns:
                if df_input[col].dtype == 'object':
                    df_input[col] = df_input[col].astype(str)
            
            st.subheader("📊 Dữ liệu đã tải lên:")
            if len(sheet_names) >= 2:
                # Hiển thị cả 2 sheet nếu có
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Sheet 1: Điểm sinh viên** ({sheet_names[0]})")
                    st.dataframe(df_input, use_container_width=True, height=300)
                    
                with col2:
                    df_max = pd.read_excel(uploaded_file, sheet_name=1)
                    st.write(f"**Sheet 2: Điểm tối đa** ({sheet_names[1]})")
                    st.dataframe(df_max, use_container_width=True, height=300)
            else:
                # Chỉ có 1 sheet
                st.dataframe(df_input, use_container_width=True)
                st.warning("⚠️ Không có sheet điểm tối đa. Sẽ sử dụng điểm cao nhất thực tế.")
                
        elif exam_type == "Hỗn hợp":
            # Đọc và hiển thị dữ liệu hỗn hợp
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            if len(sheet_names) < 2:
                st.error("❌ File Excel phải có ít nhất 2 sheet: (1) Trắc nghiệm, (2) Tự luận")
            else:
                st.subheader("📊 Dữ liệu đã tải lên:")
                
                # Hiển thị 2 sheet chính
                col1, col2 = st.columns(2)
                
                df_mcq = pd.read_excel(uploaded_file, sheet_name=0)
                df_essay = pd.read_excel(uploaded_file, sheet_name=1)
                
                with col1:
                    st.write(f"**Sheet 1: Trắc nghiệm** ({sheet_names[0]})")
                    st.dataframe(df_mcq, use_container_width=True, height=300)
                
                with col2:
                    st.write(f"**Sheet 2: Tự luận** ({sheet_names[1]})")
                    st.dataframe(df_essay, use_container_width=True, height=300)
                
                # Sheet 3 nếu có
                if len(sheet_names) >= 3:
                    with st.expander("📋 Sheet 3: Điểm tối đa (nếu có)"):
                        df_max = pd.read_excel(uploaded_file, sheet_name=2)
                        st.info(f"Sheet name: {sheet_names[2]}")
                        st.dataframe(df_max, use_container_width=True)
                else:
                    st.warning("⚠️ Không có sheet điểm tối đa. Sẽ sử dụng điểm cao nhất thực tế cho tự luận.")

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
        elif exam_type == "Hỗn hợp":
            # Phần hiển thị đã được xử lý ở trên, bây giờ chỉ cần xử lý
            try:
                if len(sheet_names) >= 2:
                    # Lấy df_max nếu có
                    df_max = None
                    if len(sheet_names) >= 3:
                        df_max = pd.read_excel(uploaded_file, sheet_name=2)
                        
                    # Tính toán
                    from mixed_exam_evaluation import calculate_mix_stats

                    all_results = calculate_mix_stats(df_mcq, df_essay, df_max)

                st.subheader("📋 Kết quả chi tiết từng câu hỏi (Hỗn hợp):")
                st.dataframe(all_results, use_container_width=True)

                # Đánh giá tổng quan sử dụng evaluate_exam_difficulty_mix
                st.subheader("📊 Đánh giá tổng quan đề hỗn hợp:")

                # Sử dụng hàm evaluate_exam_difficulty_mix cho consistency
                summary_df, conclusion, disc_info = evaluate_exam_difficulty_mix(
                    all_results,
                    tolerance=0.05,
                    check_discrimination=True
                )

                st.write("### 🔎 Cơ cấu độ khó so với mục tiêu")
                st.dataframe(summary_df, use_container_width=True)

                st.markdown(f"### ✅ Kết luận: **{conclusion}**")

                if disc_info:
                    st.write("### 📐 Thống kê độ phân biệt")
                    st.json(disc_info)
                
                # Hiển thị thống kê riêng cho từng loại
                with st.expander("📊 Thống kê chi tiết theo loại câu hỏi"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write("**Trắc nghiệm:**")
                        mc_rows = all_results[all_results['Loại câu'] == 'Trắc nghiệm']
                        if not mc_rows.empty and 'Độ khó (P)' in mc_rows.columns:
                            st.write(f"- Số câu: {len(mc_rows)}")
                            st.write(f"- Độ khó TB: {mc_rows['Độ khó (P)'].mean():.2f}")
                            if 'Độ phân biệt (D)' in mc_rows.columns:
                                st.write(f"- Độ phân biệt TB: {mc_rows['Độ phân biệt (D)'].mean():.3f}")
                    
                    with col2:
                        st.write("**Tự luận:**")
                        essay_rows = all_results[all_results['Loại câu'] == 'Tự luận']
                        if not essay_rows.empty and 'Độ khó (P)' in essay_rows.columns:
                            st.write(f"- Số câu: {len(essay_rows)}")
                            st.write(f"- Độ khó TB: {essay_rows['Độ khó (P)'].mean():.2f}")
                            if 'Độ phân biệt (D)' in essay_rows.columns:
                                st.write(f"- Độ phân biệt TB: {essay_rows['Độ phân biệt (D)'].mean():.3f}")


                # ---- Xuất file Word ----
                def convert_to_word(all_results, summary_df, conclusion, disc_info):
                    doc = Document()

                    # Tiêu đề
                    title = doc.add_heading('BÁO CÁO ĐÁNH GIÁ ĐỀ HỖN HỢP', 0)
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # Kết quả từng câu
                    doc.add_heading('1. Kết quả từng câu hỏi', level=1)
                    table = doc.add_table(rows=1, cols=len(all_results.columns))
                    table.style = 'Table Grid'
                    header_cells = table.rows[0].cells
                    for i, col in enumerate(all_results.columns):
                        header_cells[i].text = str(col)
                    for _, row in all_results.iterrows():
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
                    
                    # Thống kê theo loại câu
                    doc.add_heading('2.3. Thống kê theo loại câu hỏi', level=2)
                    
                    mc_rows = all_results[all_results['Loại câu'] == 'Trắc nghiệm']
                    essay_rows = all_results[all_results['Loại câu'] == 'Tự luận']
                    
                    doc.add_paragraph('Trắc nghiệm:')
                    if not mc_rows.empty and 'Độ khó (P)' in mc_rows.columns:
                        doc.add_paragraph(f'• Số câu: {len(mc_rows)}', style='List Bullet')
                        doc.add_paragraph(f'• Độ khó TB: {mc_rows["Độ khó (P)"].mean():.2f}', style='List Bullet')
                        if 'Độ phân biệt (D)' in mc_rows.columns:
                            doc.add_paragraph(f'• Độ phân biệt TB: {mc_rows["Độ phân biệt (D)"].mean():.3f}', style='List Bullet')
                    
                    doc.add_paragraph('Tự luận:')
                    if not essay_rows.empty and 'Độ khó (P)' in essay_rows.columns:
                        doc.add_paragraph(f'• Số câu: {len(essay_rows)}', style='List Bullet')
                        doc.add_paragraph(f'• Độ khó TB: {essay_rows["Độ khó (P)"].mean():.2f}', style='List Bullet')
                        if 'Độ phân biệt (D)' in essay_rows.columns:
                            doc.add_paragraph(f'• Độ phân biệt TB: {essay_rows["Độ phân biệt (D)"].mean():.3f}', style='List Bullet')
                    
                    doc.add_paragraph()
                    
                    # Kết luận
                    doc.add_heading('3. Kết luận', level=1)
                    conclusion_para = doc.add_paragraph(conclusion)
                    conclusion_para.runs[0].bold = True

                    # Lưu file
                    output = BytesIO()
                    doc.save(output)
                    output.seek(0)
                    return output.getvalue()


                word_data = convert_to_word(all_results, summary_df, conclusion, disc_info)
                st.download_button(
                        label="⬇️ Tải báo cáo Word (.docx)",
                        data=word_data,
                        file_name="bao_cao_de_hon_hop.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý đề hỗn hợp: {e}")

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
                doc.add_paragraph('• Công thức: D = (Điểm TB nhóm cao - Điểm TB nhóm thấp) / Điểm tối đa',
                                  style='List Bullet')
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

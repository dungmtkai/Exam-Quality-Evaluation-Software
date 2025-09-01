import pandas as pd
from processor_essay import calculate_essay_stats
from processor_multiple_choice import calculate_question_stats


def calculate_mix_stats(df_mc, df_e, df_max_score):
    """
    Tính toán thống kê kết hợp cho bài thi có cả trắc nghiệm và tự luận
    
    Parameters:
    -----------
    df_mc : pd.DataFrame
        DataFrame chứa dữ liệu điểm trắc nghiệm
    df_e : pd.DataFrame
        DataFrame chứa dữ liệu điểm tự luận
    df_max_score : pd.DataFrame
        DataFrame chứa điểm tối đa cho từng câu tự luận
        
    Returns:
    --------
    pd.DataFrame
        DataFrame chứa thống kê kết hợp với thông tin loại câu hỏi
    """
    
    # Tính thống kê cho câu trắc nghiệm
    stats_mc = calculate_question_stats(df_mc)
    # Thêm cột loại câu hỏi
    stats_mc['Loại câu'] = 'Trắc nghiệm'
    
    # Chuẩn hóa tên cột độ phân biệt cho trắc nghiệm
    if 'Độ phân biệt' in stats_mc.columns:
        stats_mc = stats_mc.rename(columns={'Độ phân biệt': 'Độ phân biệt (D)'})
    
    # Tính thống kê cho câu tự luận  
    stats_essay = calculate_essay_stats(df_e, df_max_score)
    # Thêm cột loại câu hỏi
    stats_essay['Loại câu'] = 'Tự luận'
    
    # Đổi tên cột Câu hỏi thành Câu để thống nhất
    if 'Câu hỏi' in stats_mc.columns:
        stats_mc = stats_mc.rename(columns={'Câu hỏi': 'Câu'})
    if 'Câu hỏi' in stats_essay.columns:
        stats_essay = stats_essay.rename(columns={'Câu hỏi': 'Câu'})
    
    # Thêm prefix để phân biệt câu trắc nghiệm và tự luận
    stats_mc['Câu'] = 'TN_' + stats_mc['Câu'].astype(str)
    stats_essay['Câu'] = 'TL_' + stats_essay['Câu'].astype(str)
    
    # Kết hợp 2 bảng thống kê
    stats_combined = pd.concat([stats_mc, stats_essay], ignore_index=True)
    
    # Sắp xếp các cột theo thứ tự mong muốn
    column_order = ['Câu', 'Loại câu', 'Độ khó (P)', 
                    'Mức độ', 'Độ phân biệt (D)', 'Mức độ phân biệt']
    
    # Chỉ giữ lại các cột có trong dữ liệu
    available_columns = [col for col in column_order if col in stats_combined.columns]
    stats_combined = stats_combined[available_columns]
    
    # Thêm thống kê tổng quan
    summary_stats = {
        'Tổng số câu': len(stats_combined),
        'Số câu trắc nghiệm': len(stats_mc),
        'Số câu tự luận': len(stats_essay),
        'Độ khó TB trắc nghiệm': stats_mc['Độ khó (P)'].mean() if 'Độ khó (P)' in stats_mc.columns else None,
        'Độ khó TB tự luận': stats_essay['Độ khó (P)'].mean() if 'Độ khó (P)' in stats_essay.columns else None,
        'Độ phân biệt TB trắc nghiệm': stats_mc['Độ phân biệt (D)'].mean() if 'Độ phân biệt (D)' in stats_mc.columns else None,
        'Độ phân biệt TB tự luận': stats_essay['Độ phân biệt (D)'].mean() if 'Độ phân biệt (D)' in stats_essay.columns else None
    }
    
    # In thống kê tổng quan
    print("\n=== THỐNG KÊ TỔNG QUAN BÀI THI KẾT HỢP ===")
    for key, value in summary_stats.items():
        if value is not None:
            if isinstance(value, float):
                print(f"{key}: {value:.3f}")
            else:
                print(f"{key}: {value}")
    
    print("\n=== CHI TIẾT TỪNG CÂU HỎI ===")
    
    return stats_combined
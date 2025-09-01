import pandas as pd
import numpy as np
import math
from processor_common import classify_difficulty

def calculate_essay_stats(df: pd.DataFrame, max_scores_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Tính toán độ khó và độ phân biệt cho câu hỏi tự luận
    
    Parameters:
    -----------
    df : pd.DataFrame
        DataFrame chứa dữ liệu điểm của sinh viên (sheet 1)
    max_scores_df : pd.DataFrame
        DataFrame chứa điểm tối đa của từng câu (sheet 2)
        
    Returns:
    --------
    pd.DataFrame
        DataFrame chứa kết quả phân tích độ khó và độ phân biệt
    """
    
    # Kiểm tra STT
    if 'STT' not in df.columns:
        df = df.copy()
        df['STT'] = range(1, len(df) + 1)
    
    # Lọc hợp lệ
    df_clean = df[df['STT'].apply(lambda x: str(x).isdigit())].copy()
    df_clean.reset_index(drop=True, inplace=True)
    
    # Các cột câu hỏi
    question_cols = [col for col in df_clean.columns if col.startswith("Câu")]
    
    # Lấy điểm tối đa từ sheet 2
    max_scores = {}
    if max_scores_df is not None:
        # Giả sử sheet 2 có cấu trúc: hàng 1 là tên câu, hàng 2 là điểm tối đa
        for col in question_cols:
            if col in max_scores_df.columns:
                # Lấy giá trị ở hàng đầu tiên (index 0) của cột tương ứng
                max_scores[col] = float(max_scores_df[col].iloc[0]) if not max_scores_df[col].empty else None
    
    # Tính tổng điểm mỗi SV
    df_clean["Tổng điểm"] = df_clean[question_cols].sum(axis=1)
    
    # Xếp hạng
    df_clean["Thứ hạng"] = df_clean["Tổng điểm"].rank(ascending=False, method="dense").astype(int)
    
    # Chia nhóm cao / thấp (27% mỗi nhóm)
    n = len(df_clean)
    group_size = math.floor(n * 0.27)
    df_clean = df_clean.sort_values(by="Thứ hạng").reset_index(drop=True)
    df_clean.loc[:group_size-1, "Nhóm"] = "Cao"
    df_clean.loc[n-group_size:, "Nhóm"] = "Thấp"
    df_clean["Nhóm"].fillna("Trung bình", inplace=True)
    
    # Nhóm cao & thấp
    high_group = df_clean[df_clean["Nhóm"] == "Cao"]
    low_group = df_clean[df_clean["Nhóm"] == "Thấp"]
    
    results = []
    total_students = df_clean.shape[0]
    
    for i, col in enumerate(question_cols):
        # Tính điểm trung bình
        mean_score = df_clean[col].mean()
        actual_max_score = df_clean[col].max()
        min_score = df_clean[col].min()
        std_score = df_clean[col].std()
        
        # Lấy điểm tối đa từ sheet 2 hoặc từ dữ liệu thực tế
        max_possible_score = max_scores.get(col, actual_max_score) if max_scores else actual_max_score
        
        # Độ khó mới: P = (Điểm TB / Điểm tối đa) × 100
        if max_possible_score > 0:
            P = round((mean_score / max_possible_score) * 100, 2)
        else:
            P = 0
        
        # Điểm TB nhóm cao và thấp
        mean_high = high_group[col].mean()
        mean_low = low_group[col].mean()
        
        # Độ phân biệt mới: D = (TB nhóm cao - TB nhóm thấp) / Điểm tối đa
        if max_possible_score > 0:
            D = round((mean_high - mean_low) / max_possible_score, 2)
        else:
            D = 0
        
        # Phân loại độ phân biệt theo tiêu chí mới
        if D >= 0.4:
            D_level = "Rất tốt"
        elif 0.3 <= D < 0.4:
            D_level = "Tốt"
        elif 0.2 <= D < 0.3:
            D_level = "Trung bình"
        elif D >= 0:
            D_level = "Kém"
        else:
            D_level = "Không đạt"
        
        results.append({
            "STT": i + 1,
            "Câu hỏi": col,
            "Tổng số SV": total_students,
            "Điểm TB": round(mean_score, 2),
            "Điểm tối đa": max_possible_score,
            "Điểm cao nhất (thực tế)": actual_max_score,
            "Điểm thấp nhất": min_score,
            "Độ lệch chuẩn": round(std_score, 2),
            "Độ khó (P)": P,
            "Mức độ": classify_difficulty(P),
            "Điểm TB - Nhóm cao": round(mean_high, 2),
            "Điểm TB - Nhóm thấp": round(mean_low, 2),
            "Độ phân biệt (D)": D,
            "Mức độ phân biệt": D_level
        })
    
    return pd.DataFrame(results)
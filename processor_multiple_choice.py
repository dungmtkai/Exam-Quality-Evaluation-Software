import pandas as pd
import math
from processor_common import classify_difficulty, classify_discrimination

def calculate_question_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tính toán độ khó và độ phân biệt cho câu hỏi trắc nghiệm
    
    Parameters:
    -----------
    df : pd.DataFrame
        DataFrame chứa dữ liệu điểm của sinh viên
        
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
    g = min(len(high_group), len(low_group))  # Số SV mỗi nhóm

    results = []
    total_students = df_clean.shape[0]

    for i, col in enumerate(question_cols):
        # Độ khó P: % sinh viên trả lời đúng
        num_correct = (df_clean[col] > 0).sum()
        P = round((num_correct / total_students) * 100, 2)

        # Số SV đúng trong nhóm cao và thấp
        gc = (high_group[col] > 0).sum()
        gt = (low_group[col] > 0).sum()

        # Độ phân biệt D
        D = round((gc - gt) / g, 2) if g > 0 else None
        D_level = classify_discrimination(D) if D is not None else "Không xác định"

        results.append({
            "STT": i + 1,
            "Câu hỏi": col,
            "Tổng số SV": total_students,
            "Số SV trả lời đúng": num_correct,
            "Độ khó (P)": P,
            "Mức độ": classify_difficulty(P),
            "Số SV đúng - Nhóm cao": gc,
            "Số SV đúng - Nhóm thấp": gt,
            "Độ phân biệt": D,
            "Mức độ phân biệt": D_level
        })

    return pd.DataFrame(results)
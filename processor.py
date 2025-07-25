import pandas as pd
from math import floor
from utils import classify_difficulty

def calculate_difficulty_from_df(df: pd.DataFrame) -> pd.DataFrame:
    # Loại bỏ dòng chứa chữ (ví dụ "Điểm TB") ở cột STT
    df_clean = df[df['STT'].apply(lambda x: str(x).isdigit())].copy()
    df_clean.reset_index(drop=True, inplace=True)

    question_cols = [col for col in df_clean.columns if col.startswith("Câu")]
    total_students = df_clean.shape[0]

    # Tính tổng điểm mỗi sinh viên
    df_clean["Tổng điểm"] = df_clean[question_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # Sắp xếp sinh viên theo điểm
    df_sorted = df_clean.sort_values(by="Tổng điểm", ascending=False).reset_index(drop=True)
    group_size = floor(total_students * 0.27)

    # Nhóm cao và thấp điểm
    top_group = df_sorted.iloc[:group_size]
    bottom_group = df_sorted.iloc[-group_size:]

    results = []
    for i, col in enumerate(question_cols):
        # Tính số SV trả lời đúng trong toàn bộ
        num_correct = (df_clean[col] > 0).sum()
        P = round((num_correct / total_students) * 100, 2)

        # Đếm số SV đúng trong nhóm cao và thấp
        gc = (top_group[col] > 0).sum()
        gt = (bottom_group[col] > 0).sum()
        D = round((gc - gt) / group_size, 2) if group_size > 0 else None

        results.append({
            "STT": i + 1,
            "Câu hỏi": col,
            "Số SV đúng": num_correct,
            "Tổng SV": total_students,
            "Độ khó (%)": P,
            "Mức độ": classify_difficulty(P),
            "Số SV đúng (nhóm cao)": gc,
            "Số SV đúng (nhóm thấp)": gt,
            "Độ phân biệt": D
        })

    return pd.DataFrame(results)

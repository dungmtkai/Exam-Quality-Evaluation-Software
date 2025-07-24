import pandas as pd
from utils import classify_difficulty

def calculate_difficulty_from_df(df: pd.DataFrame) -> pd.DataFrame:
    # Loại bỏ các dòng rỗng hoặc chứa chữ (như 'Điểm TB') ở cột STT
    df_clean = df[df['STT'].apply(lambda x: str(x).isdigit())].copy()
    df_clean.reset_index(drop=True, inplace=True)

    # Xác định cột câu hỏi
    question_cols = [col for col in df_clean.columns if col.startswith("Câu")]

    total_students = df_clean.shape[0]
    results = []
    for i, col in enumerate(question_cols):
        num_correct = (df_clean[col] > 0).sum()
        P = round((num_correct / total_students) * 100, 2)
        results.append({
            "STT": i + 1,
            "Câu hỏi": col,
            "Số SV đúng": num_correct,
            "Tổng SV": total_students,
            "Độ khó (%)": P,
            "Mức độ": classify_difficulty(P)
        })

    return pd.DataFrame(results)

import pandas as pd
import math

def classify_difficulty(P: float) -> str:
    if P >= 80:
        return "Dễ"
    elif 60 <= P < 80:
        return "Trung bình"
    elif 40 <= P < 60:
        return "Khó"
    else:
        return "Rất khó"

def classify_discrimination(D: float) -> str:
    if D >= 0.4:
        return "Rất tốt"
    elif 0.3 <= D < 0.4:
        return "Tốt"
    elif 0.2 <= D < 0.3:
        return "Chấp nhận được"
    elif 0 <= D < 0.2:
        return "Kém"
    else:
        return "Không đạt / âm"

def calculate_question_stats(df: pd.DataFrame) -> pd.DataFrame:
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

    # Chia nhóm cao / thấp
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
        num_correct = (df_clean[col] > 0).sum()
        P = round((num_correct / total_students) * 100, 2)

        gc = (high_group[col] > 0).sum()
        gt = (low_group[col] > 0).sum()

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



def calculate_essay_stats(df: pd.DataFrame) -> pd.DataFrame:
    import numpy as np
    
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
    
    # Chia nhóm cao / thấp
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
        max_score = df_clean[col].max()
        min_score = df_clean[col].min()
        
        # Độ khó cho tự luận: điểm TB / (điểm max - điểm min)
        # Nếu max = min (tất cả SV cùng điểm), độ khó = điểm TB / max
        if max_score > min_score:
            P = round((mean_score / (max_score - min_score)) * 100, 2)
        elif max_score > 0:
            P = round((mean_score / max_score) * 100, 2)
        else:
            P = 0
        
        # Độ phân biệt: Hệ số tương quan giữa điểm câu hỏi và tổng điểm (loại trừ câu đó)
        # Tính tổng điểm không bao gồm câu hiện tại
        total_without_current = df_clean["Tổng điểm"] - df_clean[col]
        
        # Tính hệ số tương quan Pearson
        if total_without_current.std() > 0 and df_clean[col].std() > 0:
            correlation = np.corrcoef(df_clean[col], total_without_current)[0, 1]
            D = round(correlation, 2)
        else:
            D = 0
        
        # Phân loại độ phân biệt cho tự luận (dựa trên hệ số tương quan)
        if D >= 0.4:
            D_level = "Rất tốt"
        elif 0.3 <= D < 0.4:
            D_level = "Tốt"
        elif 0.2 <= D < 0.3:
            D_level = "Chấp nhận được"
        elif 0 <= D < 0.2:
            D_level = "Kém"
        else:
            D_level = "Không đạt / âm"
        
        # Điểm TB nhóm cao và thấp
        mean_high = high_group[col].mean()
        mean_low = low_group[col].mean()
        
        results.append({
            "STT": i + 1,
            "Câu hỏi": col,
            "Tổng số SV": total_students,
            "Điểm TB": round(mean_score, 2),
            "Điểm cao nhất": max_score,
            "Điểm thấp nhất": min_score,
            "Độ khó (P)": P,
            "Mức độ": classify_difficulty(P),
            "Điểm TB - Nhóm cao": round(mean_high, 2),
            "Điểm TB - Nhóm thấp": round(mean_low, 2),
            "Độ phân biệt (r)": D,
            "Mức độ phân biệt": D_level
        })
    
    return pd.DataFrame(results)


import pandas as pd

def evaluate_exam_difficulty_mix(
    stats_df: pd.DataFrame,
    target_mix = {"Dễ": 0.50, "Trung bình": 0.30, "Khó": 0.20},
    tolerance: float = 0.05,
    check_discrimination: bool = False,
    min_good_D_share: float = 0.60,   # ít nhất 60% câu có D >= 0.2
    max_negative_D_share: float = 0.10 # không quá 10% câu có D < 0
):
    df = stats_df.copy()

    # Chuẩn hoá nhóm độ khó: gộp "Rất khó" vào "Khó"
    df["Mức độ (chuẩn)"] = df["Mức độ"].replace({"Rất khó": "Khó"})

    # Đếm và tính tỷ lệ thực tế
    total_items = len(df)
    counts = df["Mức độ (chuẩn)"].value_counts().reindex(["Dễ","Trung bình","Khó"], fill_value=0)
    percents = (counts / total_items).round(4)

    # Tạo bảng so sánh
    summary = pd.DataFrame({
        "Nhóm": ["Dễ","Trung bình","Khó"],
        "Số câu": counts.values,
        "Tỷ lệ thực tế": percents.values,
        "Tỷ lệ mục tiêu": [target_mix["Dễ"], target_mix["Trung bình"], target_mix["Khó"]],
    })
    summary["Khoảng chấp nhận"] = summary["Tỷ lệ mục tiêu"].apply(
        lambda x: f"[{round(x - tolerance, 3)}, {round(x + tolerance, 3)}]"
    )
    summary["Đạt nhóm?"] = (
        (summary["Tỷ lệ thực tế"] >= (summary["Tỷ lệ mục tiêu"] - tolerance)) &
        (summary["Tỷ lệ thực tế"] <= (summary["Tỷ lệ mục tiêu"] + tolerance))
    )

    # Kết luận theo cơ cấu độ khó
    pass_mix = summary["Đạt nhóm?"].all()

    # (Tùy chọn) kiểm tra chất lượng phân biệt
    disc_result = None
    pass_disc = True
    if check_discrimination and "Độ phân biệt" in df.columns:
        good_D_share = (df["Độ phân biệt"] >= 0.2).mean()
        negative_D_share = (df["Độ phân biệt"] < 0).mean()
        pass_disc = (good_D_share >= min_good_D_share) and (negative_D_share <= max_negative_D_share)
        disc_result = {
            "Tỷ lệ D >= 0.2": round(good_D_share, 4),
            "Tỷ lệ D < 0": round(negative_D_share, 4),
            "Đạt tiêu chí phân biệt?": pass_disc
        }

    # Kết luận cuối cùng
    if check_discrimination and disc_result is not None:
        overall_pass = pass_mix and pass_disc
    else:
        overall_pass = pass_mix

    conclusion = "Đạt chuẩn cơ cấu độ khó" if overall_pass else "Không đạt chuẩn cơ cấu độ khó"

    return summary, conclusion, disc_result


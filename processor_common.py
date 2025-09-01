import pandas as pd

def classify_difficulty(P: float) -> str:
    """Phân loại mức độ khó dựa trên chỉ số P"""
    if P >= 80:
        return "Dễ"
    elif 60 <= P < 80:
        return "Trung bình"
    elif 40 <= P < 60:
        return "Khó"
    else:
        return "Rất khó"

def classify_discrimination(D: float) -> str:
    """Phân loại mức độ phân biệt dựa trên chỉ số D"""
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

def evaluate_exam_difficulty_mix(
    stats_df: pd.DataFrame,
    target_mix = {"Dễ": 0.50, "Trung bình": 0.30, "Khó": 0.20},
    tolerance: float = 0.05,
    check_discrimination: bool = False,
    min_good_D_share: float = 0.60,   # ít nhất 60% câu có D >= 0.2
    max_negative_D_share: float = 0.10 # không quá 10% câu có D < 0
):
    """Đánh giá cơ cấu độ khó của đề thi"""
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
    
    # Tìm cột độ phân biệt (có thể là "Độ phân biệt" hoặc "Độ phân biệt (D)" hoặc "Độ phân biệt (r)")
    disc_col = None
    for col in df.columns:
        if "Độ phân biệt" in col:
            disc_col = col
            break
    
    if check_discrimination and disc_col:
        good_D_share = (df[disc_col] >= 0.2).mean()
        negative_D_share = (df[disc_col] < 0).mean()
        pass_disc = (good_D_share >= min_good_D_share) and (negative_D_share <= max_negative_D_share)
        disc_result = {
            "Tỷ lệ D >= 0.2": round(good_D_share, 4),
            "Tỷ lệ D < 0": round(negative_D_share, 4),
            "Đạt tiêu chí phân biệt?": pass_disc
        }

    # Kết luận chi tiết
    conclusions = []
    
    # Kết luận về cơ cấu độ khó
    if pass_mix:
        conclusions.append("✅ Đạt chuẩn cơ cấu độ khó")
    else:
        conclusions.append("❌ Không đạt chuẩn cơ cấu độ khó")
    
    # Kết luận về độ phân biệt (nếu có kiểm tra)
    if check_discrimination and disc_result is not None:
        if pass_disc:
            conclusions.append("✅ Đạt tiêu chí độ phân biệt")
        else:
            conclusions.append("❌ Không đạt tiêu chí độ phân biệt")
    
    # Kết luận tổng thể
    if check_discrimination and disc_result is not None:
        overall_pass = pass_mix and pass_disc
        if overall_pass:
            final_conclusion = "Đạt chuẩn tổng thể"
        else:
            final_conclusion = " | ".join(conclusions)
    else:
        final_conclusion = conclusions[0].replace("✅ ", "").replace("❌ ", "")

    return summary, final_conclusion, disc_result
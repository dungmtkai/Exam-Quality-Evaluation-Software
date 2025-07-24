def classify_difficulty(p: float) -> str:
    if p >= 80:
        return "Dễ"
    elif 60 <= p < 80:
        return "Trung bình"
    elif 20 <= p < 60:
        return "Khó"
    else:
        return "Rất khó"

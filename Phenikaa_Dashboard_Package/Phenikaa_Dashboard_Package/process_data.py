import pandas as pd
import json

file_path = "/Users/dotran/Library/CloudStorage/OneDrive-Personal/Phenikaa/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM_Độ.xlsx"

# Read Append1 (Sessions)
df_sessions = pd.read_excel(file_path, sheet_name="Append1")
df_sessions = df_sessions.dropna(subset=["ĐƠN VỊ"])
sessions = df_sessions.to_dict(orient="records")

# Read Dữ liệu học viên (Students)
df_students = pd.read_excel(file_path, sheet_name="Dữ liệu học viên")

# Process Đạt/Không đạt
def safe_float(s):
    try:
        s = str(s).strip()
        if '/' in s:
            parts = s.split('/')
            return (float(parts[0]) / float(parts[1])) * 10
        val = float(s)
        if val > 30000:
            date_val = pd.to_datetime('1899-12-30') + pd.to_timedelta(val, 'D')
            return (date_val.day / date_val.month) * 10
        return val * 10 if val <= 1 else val if val <= 10 else val / 10
    except:
        return None

def is_passed(score):
    val = safe_float(score)
    if val is None:
        return False
    return val >= 8

df_students['Điểm_Quy_Đổi'] = df_students['Điểm số'].apply(safe_float)
df_students['Trạng thái'] = df_students['Điểm số'].apply(lambda x: 'Đạt' if is_passed(x) else ('Không đạt' if not pd.isna(x) and str(x).strip() else ''))

# Clean student NaN
df_students = df_students.fillna("")
students = df_students.to_dict(orient="records")

output = {
    "sessions": sessions,
    "students": students
}

import math
# fix NaNs in sessions
for s in output["sessions"]:
    for k, v in s.items():
        if isinstance(v, float) and math.isnan(v):
            s[k] = ""

with open("/Users/dotran/Library/CloudStorage/OneDrive-Personal/Phenikaa/netlify-dashboard/data.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False)

print("Data exported successfully.")

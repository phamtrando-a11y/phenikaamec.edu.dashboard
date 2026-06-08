import pandas as pd
f2 = "/Users/dotran/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM (Độ)_unlocked.xlsx"
xl2 = pd.ExcelFile(f2)
sheet_s = "Sinh Hoạt chuyên môn 2026" if "Sinh Hoạt chuyên môn 2026" in xl2.sheet_names else "SHCM - Độ" if "SHCM - Độ" in xl2.sheet_names else "Bản sao của Sinh Hoạt chuyên mô"
df = pd.read_excel(xl2, sheet_name=sheet_s)
print(df[['THỜI GIAN', 'Ngày cụ thể']].head(10))

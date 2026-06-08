import pandas as pd
import numpy as np

def read_file(file_path):
    xl = pd.ExcelFile(file_path)
    sheet_s = "Sinh Hoạt chuyên môn 2026" if "Sinh Hoạt chuyên môn 2026" in xl.sheet_names else "SHCM - Độ" if "SHCM - Độ" in xl.sheet_names else "Bản sao của Sinh Hoạt chuyên mô"
    df_sessions = pd.read_excel(xl, sheet_name=sheet_s)
    column_mapping = {
        'TÊN ĐƠN VỊ': 'ĐƠN VỊ', 'Tình trạng': 'Tình trạng thực hiện',
        'Số học viên tham gia': 'Số học viên tham gia', 'Số Học viên đạt': 'Số Học viên đạt',
        'NỘI DUNG/CHƯƠNG TRÌNH': 'NỘI DUNG/CHƯƠNG TRÌNH', 'NGƯỜI PHỤ TRÁCH NỘI DUNG': 'NGƯỜI PHỤ TRÁCH NỘI DUNG',
        'Điểm danh\n13. Sinh hoạt chuyên môn': 'Gắn link Điểm danh và post test',
        'Bảng kiểm đào tạo nội bộ': 'Gắn link bảng kiểm đào tạo nội bộ',
        'Báo cáo': 'Gắn link bài báo cáo', 'Link video buổi đào tạo': 'Gắn link video buổi đào tạo'
    }
    df_sessions = df_sessions.rename(columns=column_mapping)
    if 'Ngày' not in df_sessions.columns: df_sessions['Ngày'] = np.nan
    if 'Tháng' not in df_sessions.columns: df_sessions['Tháng'] = np.nan
    if 'Năm' not in df_sessions.columns: df_sessions['Năm'] = np.nan
    
    df_students_full = pd.DataFrame()
    if "Dữ liệu học viên 2026" in xl.sheet_names:
        df_students_full = pd.read_excel(xl, sheet_name="Dữ liệu học viên 2026")
    elif "2026" in xl.sheet_names:
        df_2026_raw = pd.read_excel(xl, sheet_name="2026")
        df_students_full = pd.DataFrame()
        df_students_full['Họ và tên'] = df_2026_raw['Họ và tên'] if 'Họ và tên' in df_2026_raw.columns else df_2026_raw['Gốc'] if 'Gốc' in df_2026_raw.columns else np.nan
        df_students_full['Gốc'] = df_2026_raw['Gốc'] if 'Gốc' in df_2026_raw.columns else df_2026_raw['Họ và tên']
        df_students_full['Mã nhân sự'] = df_2026_raw['Mã nhân sự'] if 'Mã nhân sự' in df_2026_raw.columns else np.nan
        df_students_full['Điểm số'] = df_2026_raw['Điểm số'] if 'Điểm số' in df_2026_raw.columns else np.nan
        df_students_full['Đơn vị đào tạo'] = df_2026_raw['Đơn vị đào tạo'] if 'Đơn vị đào tạo' in df_2026_raw.columns else np.nan
        df_students_full['Tháng thực hiện'] = df_2026_raw['Tháng'] if 'Tháng' in df_2026_raw.columns else np.nan
        df_students_full['Ngày đào tạo'] = df_2026_raw['Ngày'] if 'Ngày' in df_2026_raw.columns else np.nan
        df_students_full['Năm'] = df_2026_raw['Năm'] if 'Năm' in df_2026_raw.columns else 2026
        df_students_full['Tên chương trình'] = df_2026_raw['Tên chương trình'] if 'Tên chương trình' in df_2026_raw.columns else np.nan
        
    df_students_scraped = pd.DataFrame()
    if "Dữ liệu học viên" in xl.sheet_names:
        df_students_scraped = pd.read_excel(xl, sheet_name="Dữ liệu học viên")
        
    return df_sessions, df_students_full, df_students_scraped

f1 = "/Users/dotran/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM.xlsx"
f2 = "/Users/dotran/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM (Độ)_unlocked.xlsx"

s1, sf1, ss1 = read_file(f1)
s2, sf2, ss2 = read_file(f2)

s1 = s1[s1['Tháng'] != 6]
s2 = s2[s2['Tháng'] == 6]
s_final = pd.concat([s1, s2], ignore_index=True)

month_col_1 = 'Tháng thực hiện' if 'Tháng thực hiện' in sf1.columns else 'Tháng'
sf1 = sf1[sf1[month_col_1] != 6] if month_col_1 in sf1.columns else sf1

month_col_2 = 'Tháng thực hiện' if 'Tháng thực hiện' in sf2.columns else 'Tháng'
sf2 = sf2[sf2[month_col_2] == 6] if month_col_2 in sf2.columns else sf2

sf_final = pd.concat([sf1, sf2], ignore_index=True)

print("s_final len:", len(s_final))
print("sf_final len:", len(sf_final))

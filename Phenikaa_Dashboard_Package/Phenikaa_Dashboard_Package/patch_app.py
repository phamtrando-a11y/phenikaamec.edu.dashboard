import re

with open('app.py', 'r') as f:
    content = f.read()

# Modify load_data definition
content = content.replace("def load_data():", "def load_data(creds_json=None):")

# Modify read_excel_sheets to only return df_sessions
old_read = """        def read_excel_sheets(xl):
            sheet_s = "Sinh Hoạt chuyên môn 2026" if "Sinh Hoạt chuyên môn 2026" in xl.sheet_names else "SHCM - Độ" if "SHCM - Độ" in xl.sheet_names else "Bản sao của Sinh Hoạt chuyên mô"
            df_sessions = pd.read_excel(xl, sheet_name=sheet_s)
            column_mapping = {
                'TÊN ĐƠN VỊ': 'ĐƠN VỊ', 'Tình trạng': 'Tình trạng thực hiện',
                'Số học viên tham gia': 'Số học viên tham gia', 'Số Học viên đạt': 'Số Học viên đạt',
                'NỘI DUNG/CHƯƠNG TRÌNH': 'NỘI DUNG/CHƯƠNG TRÌNH', 'NGƯỜI PHỤ TRÁCH NỘI DUNG': 'NGƯỜI PHỤ TRÁCH NỘI DUNG',
                'Điểm danh\\n13. Sinh hoạt chuyên môn': 'Gắn link Điểm danh và post test',
                'Bảng kiểm đào tạo nội bộ': 'Gắn link bảng kiểm đào tạo nội bộ',
                'Báo cáo': 'Gắn link bài báo cáo', 'Link video buổi đào tạo': 'Gắn link video buổi đào tạo',
                'TÌNH TRẠNG THỰC HIỆN': 'Tình trạng thực hiện',
                'SỐ HỌC VIÊN THAM GIA': 'Số học viên tham gia', 'SỐ HỌC VIÊN ĐẠT ĐIỂM TRÊN 80%': 'Số Học viên đạt',
                'GẮN LINK POST TEST': 'Gắn link Điểm danh và post test',
                'GẮN LINK BẢNG KIỂM': 'Gắn link bảng kiểm đào tạo nội bộ',
                'GẮN LINK BÀI BÁO CÁO': 'Gắn link bài báo cáo',
                'GẮN LINK VIDEO': 'Gắn link video buổi đào tạo'
            }
            df_sessions = df_sessions.rename(columns=column_mapping)
            
            if 'Ngày cụ thể' in df_sessions.columns:
                dt = pd.to_datetime(df_sessions['Ngày cụ thể'], errors='coerce')
                if 'Ngày' not in df_sessions.columns: df_sessions['Ngày'] = dt.dt.day
                if 'Tháng' not in df_sessions.columns: df_sessions['Tháng'] = dt.dt.month
                if 'Năm' not in df_sessions.columns: df_sessions['Năm'] = dt.dt.year
                
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
                
            return df_sessions, df_students_full, df_students_scraped"""

new_read = """        def read_excel_sheets(xl):
            sheet_s = "Sinh Hoạt chuyên môn 2026" if "Sinh Hoạt chuyên môn 2026" in xl.sheet_names else "SHCM - Độ" if "SHCM - Độ" in xl.sheet_names else "Bản sao của Sinh Hoạt chuyên mô"
            df_sessions = pd.read_excel(xl, sheet_name=sheet_s)
            column_mapping = {
                'TÊN ĐƠN VỊ': 'ĐƠN VỊ', 'Tình trạng': 'Tình trạng thực hiện',
                'Số học viên tham gia': 'Số học viên tham gia', 'Số Học viên đạt': 'Số Học viên đạt',
                'NỘI DUNG/CHƯƠNG TRÌNH': 'NỘI DUNG/CHƯƠNG TRÌNH', 'NGƯỜI PHỤ TRÁCH NỘI DUNG': 'NGƯỜI PHỤ TRÁCH NỘI DUNG',
                'Điểm danh\\n13. Sinh hoạt chuyên môn': 'Gắn link Điểm danh và post test',
                'Bảng kiểm đào tạo nội bộ': 'Gắn link bảng kiểm đào tạo nội bộ',
                'Báo cáo': 'Gắn link bài báo cáo', 'Link video buổi đào tạo': 'Gắn link video buổi đào tạo',
                'TÌNH TRẠNG THỰC HIỆN': 'Tình trạng thực hiện',
                'SỐ HỌC VIÊN THAM GIA': 'Số học viên tham gia', 'SỐ HỌC VIÊN ĐẠT ĐIỂM TRÊN 80%': 'Số Học viên đạt',
                'GẮN LINK POST TEST': 'Gắn link Điểm danh và post test',
                'GẮN LINK BẢNG KIỂM': 'Gắn link bảng kiểm đào tạo nội bộ',
                'GẮN LINK BÀI BÁO CÁO': 'Gắn link bài báo cáo',
                'GẮN LINK VIDEO': 'Gắn link video buổi đào tạo'
            }
            df_sessions = df_sessions.rename(columns=column_mapping)
            
            if 'Ngày cụ thể' in df_sessions.columns:
                dt = pd.to_datetime(df_sessions['Ngày cụ thể'], errors='coerce')
                if 'Ngày' not in df_sessions.columns: df_sessions['Ngày'] = dt.dt.day
                if 'Tháng' not in df_sessions.columns: df_sessions['Tháng'] = dt.dt.month
                if 'Năm' not in df_sessions.columns: df_sessions['Năm'] = dt.dt.year
                
            if 'Ngày' not in df_sessions.columns: df_sessions['Ngày'] = np.nan
            if 'Tháng' not in df_sessions.columns: df_sessions['Tháng'] = np.nan
            if 'Năm' not in df_sessions.columns: df_sessions['Năm'] = np.nan
                
            return df_sessions"""

content = content.replace(old_read, new_read)

# Modify the unpacking and concatenation
old_unpack = """        xl1 = pd.ExcelFile(xl_file)
        s1, sf1, ss1 = read_excel_sheets(xl1)
        
        file2_path = "/Users/trando/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM (Độ)_unlocked.xlsx"
        if os.path.exists(file2_path):
            xl2_file = file2_path
        else:
            file2_id = "12FpcGGKs4LdijihdYjQnS2qdf3RQW4-W" # User provided the exact same ID, so it downloads the merged file
            url2 = f"https://docs.google.com/spreadsheets/d/{file2_id}/export?format=xlsx"
            try:
                response2 = requests.get(url2, timeout=20)
                xl2_file = io.BytesIO(response2.content)
            except:
                xl2_file = None
                
        if xl2_file is not None:
            xl2 = pd.ExcelFile(xl2_file)
            s2, sf2, ss2 = read_excel_sheets(xl2)
            
            s1 = s1[s1['Tháng'] != 6]
            s2 = s2[s2['Tháng'] == 6]
            df_sessions = pd.concat([s1, s2], ignore_index=True)
            
            month_col_1 = 'Tháng thực hiện' if 'Tháng thực hiện' in sf1.columns else 'Tháng'
            if month_col_1 in sf1.columns: sf1 = sf1[sf1[month_col_1] != 6]
            
            month_col_2 = 'Tháng thực hiện' if 'Tháng thực hiện' in sf2.columns else 'Tháng'
            if month_col_2 in sf2.columns: sf2 = sf2[sf2[month_col_2] == 6]
            
            df_students_full = pd.concat([sf1, sf2], ignore_index=True)
            df_students_scraped = pd.concat([ss1, ss2], ignore_index=True)
        else:
            df_sessions, df_students_full, df_students_scraped = s1, sf1, ss1"""

new_unpack = """        xl1 = pd.ExcelFile(xl_file)
        s1 = read_excel_sheets(xl1)
        
        file2_path = "/Users/dotran/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM (Độ)_unlocked.xlsx"
        if os.path.exists(file2_path):
            xl2_file = file2_path
        else:
            file2_id = "12FpcGGKs4LdijihdYjQnS2qdf3RQW4-W"
            url2 = f"https://docs.google.com/spreadsheets/d/{file2_id}/export?format=xlsx"
            try:
                response2 = requests.get(url2, timeout=20)
                xl2_file = io.BytesIO(response2.content)
            except:
                xl2_file = None
                
        if xl2_file is not None:
            xl2 = pd.ExcelFile(xl2_file)
            s2 = read_excel_sheets(xl2)
            s1 = s1[s1['Tháng'] != 6]
            s2 = s2[s2['Tháng'] == 6]
            df_sessions = pd.concat([s1, s2], ignore_index=True)
        else:
            df_sessions = s1

        # Đọc dữ liệu học viên từ Data_HocVien.xlsx
        df_students_full = pd.DataFrame()
        if creds_json:
            try:
                from googleapiclient.discovery import build
                from google.oauth2 import service_account
                import googleapiclient.http
                
                creds_dict = json.loads(creds_json)
                creds = service_account.Credentials.from_service_account_info(
                    creds_dict, scopes=["https://www.googleapis.com/auth/drive.readonly"]
                )
                drive_service = build('drive', 'v3', credentials=creds)
                
                reg_folder_id = "1PqDmjmYzvy30CKG-iqdVSd2ibOBRDGO_"
                query = f"name='Data_HocVien.xlsx' and '{reg_folder_id}' in parents and trashed=false"
                results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
                files = results.get('files', [])
                
                if files:
                    file_id = files[0]['id']
                    request = drive_service.files().get_media(fileId=file_id)
                    fh = io.BytesIO()
                    downloader = googleapiclient.http.MediaIoBaseDownload(fh, request)
                    done = False
                    while done is False:
                        status, done = downloader.next_chunk()
                    fh.seek(0)
                    df_students_full = pd.read_excel(fh)
            except Exception as e:
                print(f"Lỗi khi tải Data_HocVien.xlsx: {e}")
        else:
            # Fallback local testing if no creds
            local_hv = "/Users/dotran/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/Data_HocVien.xlsx"
            if os.path.exists(local_hv):
                df_students_full = pd.read_excel(local_hv)
                
        df_students_scraped = df_students_full.copy()"""

content = content.replace(old_unpack, new_unpack)

# Also update the caller of load_data
content = content.replace("df_sessions, df_students = load_data()", """creds_json = st.secrets.get("google_credentials") if "google_credentials" in st.secrets else None
    df_sessions, df_students = load_data(creds_json)""")

with open('app.py', 'w') as f:
    f.write(content)

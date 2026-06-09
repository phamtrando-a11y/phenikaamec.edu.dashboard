import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import datetime
import io
import requests
import json
import base64
import os
try:
    from github import Github
except ImportError:
    Github = None
import dataframe_image as dfi
import uuid

st.set_page_config(page_title="Dashboard Báo Cáo CME", layout="wide")



@st.dialog("Ảnh báo cáo (Chuột phải -> Sao chép hình ảnh)", width="large")
def show_image_modal(df_styler):
    with st.spinner("Đang tạo ảnh, vui lòng chờ..."):
        filename = f"temp_table_{uuid.uuid4().hex[:8]}.png"
        try:
            dfi.export(df_styler, filename, table_conversion='matplotlib', max_rows=-1)
            st.image(filename, use_container_width=True)
        except Exception as e:
            st.error(f"Lỗi tạo ảnh: {e}")
        finally:
            if os.path.exists(filename):
                try:
                    os.remove(filename)
                except:
                    pass

@st.cache_data(ttl=10)
def load_data(creds_json=None):
    file_path = "/Users/trando/Library/CloudStorage/GoogleDrive-daotao.bvdhphenikaa@gmail.com/Drive của tôi/1_Đào tạo/13. Sinh hoạt chuyên môn/250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM.xlsx"
    
    # Kiểm tra xem đang chạy local hay online
    if os.path.exists(file_path):
        # Đọc trực tiếp từ local
        xl_file = file_path
    else:
        # Nếu chạy online (Streamlit Cloud), tải từ link Google Drive của bạn
        file_id = "12FpcGGKs4LdijihdYjQnS2qdf3RQW4-W"
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        try:
            response = requests.get(url, timeout=20)
            xl_file = io.BytesIO(response.content)
        except Exception as e:
            st.error(f"Không thể tải file Excel từ Google Drive: {e}")
            st.stop()
            
    # Đọc file Excel
    try:
        def read_excel_sheets(xl):
            sheet_s = "Sinh Hoạt chuyên môn 2026" if "Sinh Hoạt chuyên môn 2026" in xl.sheet_names else "SHCM - Độ" if "SHCM - Độ" in xl.sheet_names else "Bản sao của Sinh Hoạt chuyên mô"
            df_sessions = pd.read_excel(xl, sheet_name=sheet_s)
            column_mapping = {
                'TÊN ĐƠN VỊ': 'ĐƠN VỊ', 'Tình trạng': 'Tình trạng thực hiện',
                'Số học viên tham gia': 'Số học viên tham gia', 'Số Học viên đạt': 'Số Học viên đạt',
                'NỘI DUNG/CHƯƠNG TRÌNH': 'NỘI DUNG/CHƯƠNG TRÌNH', 'NGƯỜI PHỤ TRÁCH NỘI DUNG': 'NGƯỜI PHỤ TRÁCH NỘI DUNG',
                'Điểm danh\n13. Sinh hoạt chuyên môn': 'Gắn link Điểm danh và post test',
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
                
            return df_sessions

        xl1 = pd.ExcelFile(xl_file)
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
                
        df_students_scraped = df_students_full.copy()
            
    except Exception as e:
        st.error(f"Lỗi khi xử lý dữ liệu Excel: {e}")
        st.stop()
        
    df_sessions = df_sessions.dropna(subset=["ĐƠN VỊ"])
    
    df_sessions['Tình trạng thực hiện'] = df_sessions['Tình trạng thực hiện'].fillna('Chưa rõ').astype(str).str.strip()
    def normalize_status(s):
        s_lower = s.lower()
        if 'không thực hiện' in s_lower:
            return 'Không thực hiện'
        elif 'thực hiện' in s_lower or 'có' in s_lower:
            return 'Thực hiện'
        return 'Chưa cập nhật'
    df_sessions['Tình trạng'] = df_sessions['Tình trạng thực hiện'].apply(normalize_status)
    
    if 'Số học viên tham gia' in df_sessions.columns:
        df_sessions['Số học viên tham gia'] = pd.to_numeric(df_sessions['Số học viên tham gia'], errors='coerce').fillna(0)
    if 'Số Học viên đạt' in df_sessions.columns:
        df_sessions['Số Học viên đạt'] = pd.to_numeric(df_sessions['Số Học viên đạt'], errors='coerce').fillna(0)
        
    if not df_students_full.empty:
        if not df_students_scraped.empty:
            df_students = pd.concat([df_students_full, df_students_scraped], ignore_index=True)
        else:
            df_students = df_students_full
        # Loại bỏ trùng lặp dựa trên mã định danh và tên chương trình
        df_students['Họ Tên Đầy Đủ Temp'] = df_students['Họ và tên'].fillna(df_students['Gốc'])
        df_students['Mã_Định_Danh_Temp'] = df_students['Mã nhân sự'].fillna(df_students['Họ Tên Đầy Đủ Temp'])
        df_students = df_students.drop_duplicates(subset=['Mã_Định_Danh_Temp', 'Tên chương trình'])
        df_students = df_students.drop(columns=['Họ Tên Đầy Đủ Temp', 'Mã_Định_Danh_Temp'])
    else:
        df_students = df_students_scraped

    def make_date(row, d_col, m_col, y_col):
        try:
            d = int(float(str(row.get(d_col, np.nan)).strip()))
            m = int(float(str(row.get(m_col, np.nan)).strip()))
            y = int(float(str(row.get(y_col, np.nan)).strip()))
            return pd.Timestamp(year=y, month=m, day=d)
        except:
            return pd.NaT

    df_sessions['FullDate'] = df_sessions.apply(lambda r: make_date(r, 'Ngày', 'Tháng', 'Năm'), axis=1)
    
    if not df_students.empty:
        if 'Năm' not in df_students.columns: df_students['Năm'] = 2026
        df_students['FullDate'] = df_students.apply(lambda r: make_date(r, 'Ngày đào tạo', 'Tháng thực hiện', 'Năm'), axis=1)

    
    def is_passed(score):
        if pd.isna(score): return None
        if isinstance(score, datetime.datetime) or isinstance(score, pd.Timestamp):
            try: return (score.day / score.month) >= 0.8
            except: return False
            
        s = str(score).strip()
        if '-' in s and ':' in s:
            try:
                dt = pd.to_datetime(s)
                return (dt.day / dt.month) >= 0.8
            except: pass
            
        if s.endswith('%'):
            try: return float(s[:-1]) >= 80
            except: return False
            
        if '/' in s:
            try:
                parts = s.split('/')
                return (float(parts[0]) / float(parts[1])) >= 0.8
            except:
                return False
        try:
            val = float(s)
            if val > 30000:
                date_val = pd.to_datetime('1899-12-30') + pd.to_timedelta(val, 'D')
                val = (date_val.day / date_val.month) * 10
            if val <= 10: return val >= 8
            elif val <= 100: return val >= 80
            else: return False
        except:
            return False
        
    df_students['Trạng thái'] = df_students['Điểm số'].apply(lambda x: 'Đạt' if is_passed(x) else ('Không đạt' if not pd.isna(x) else ''))
    
    if not df_sessions.empty:
        for idx, row in df_sessions.iterrows():
            has_link = False
            link_col = 'Gắn link Điểm danh và post test'
            if link_col in row and pd.notna(row[link_col]) and str(row[link_col]).strip() != '':
                has_link = True
                
            session_name = row.get('NỘI DUNG/CHƯƠNG TRÌNH')
            if pd.notna(session_name):
                session_name_str = str(session_name).strip().lower()
                if not df_students.empty:
                    students_of_session = df_students[df_students['Tên chương trình'].astype(str).str.strip().str.lower() == session_name_str]
                    count_students = len(students_of_session)
                    passed_count = len(students_of_session[students_of_session['Trạng thái'] == 'Đạt'])
                else:
                    count_students = 0
                    passed_count = 0
                    
                df_sessions.at[idx, 'Số học viên tham gia'] = count_students
                df_sessions.at[idx, 'Số Học viên đạt'] = passed_count
    
    if 'Họ và tên' in df_students.columns and 'Gốc' in df_students.columns:
        df_students['Họ Tên Đầy Đủ'] = df_students['Họ và tên'].fillna(df_students['Gốc'])
    else:
        df_students['Họ Tên Đầy Đủ'] = df_students['Họ và tên'] if 'Họ và tên' in df_students.columns else "Không rõ"
        
    df_students['Mã_Định_Danh'] = df_students['Mã nhân sự'].fillna(df_students['Họ Tên Đầy Đủ'])
    
    def safe_float(s):
        try:
            if isinstance(s, datetime.datetime) or isinstance(s, pd.Timestamp):
                return (s.day / s.month) * 10
                
            s_str = str(s).strip()
            if '-' in s_str and ':' in s_str:
                dt = pd.to_datetime(s_str)
                return (dt.day / dt.month) * 10
                
            if s_str.endswith('%'):
                return float(s_str[:-1]) / 10 if float(s_str[:-1]) > 10 else float(s_str[:-1])
                
            if '/' in s_str:
                parts = s_str.split('/')
                return (float(parts[0]) / float(parts[1])) * 10
                
            val = float(s_str)
            if val > 30000:
                date_val = pd.to_datetime('1899-12-30') + pd.to_timedelta(val, 'D')
                return (date_val.day / date_val.month) * 10
            return val * 10 if val <= 1 else val if val <= 10 else val / 10
        except:
            return None
            
    df_students['Điểm_Quy_Đổi'] = df_students['Điểm số'].apply(safe_float)
    
    return df_sessions, df_students

try:
    creds_json = st.secrets.get("google_credentials") if "google_credentials" in st.secrets else None
    df_sessions, df_students = load_data(creds_json)
except Exception as e:
    st.error(f"Lỗi khi đọc file: {e}")
    st.stop()

# --- SIDEBAR MENU ---
st.sidebar.title("📊 Phenikaa CME")

if "google_credentials" in st.secrets:
    if st.sidebar.button("🔄 Lấy dữ liệu học viên mới nhất"):
        with st.spinner("Đang kết nối Google Drive và bóc tách dữ liệu... Vui lòng chờ 30s-1p..."):
            try:
                import google_drive_scraper
                main_excel_id = "12FpcGGKs4LdijihdYjQnS2qdf3RQW4-W"
                reg_folder_id = "1PqDmjmYzvy30CKG-iqdVSd2ibOBRDGO_"
                creds_json = st.secrets["google_credentials"]
                success = google_drive_scraper.run_scraper(creds_json, reg_folder_id, main_excel_id)
                if success:
                    st.cache_data.clear()
                    st.sidebar.success("Cập nhật thành công! Đang tải lại...")
                    st.rerun()
            except Exception as e:
                st.sidebar.error(f"Lỗi: {e}")

st.sidebar.divider()
page = st.sidebar.radio("Chọn chức năng:", ["Tổng quan", "Hoạt động khoa phòng", "Hoạt động đào tạo", "Đánh giá chỉ tiêu"])

st.title(page)
st.caption("Trực quan hóa dữ liệu điểm danh và cấp CME")

# --- GLOBAL FILTERS ---
temp_df = df_sessions.copy()
selected_years, selected_months, selected_units = [], [], []

if page in ["Tổng quan", "Hoạt động khoa phòng", "Đánh giá chỉ tiêu"]:
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        years = sorted(temp_df["Năm"].dropna().unique().tolist())
        selected_years = st.multiselect("Năm:", years, format_func=lambda x: str(int(x)))
        if selected_years: temp_df = temp_df[temp_df["Năm"].isin(selected_years)]
    
    with col_f2:
        months = sorted(temp_df["Tháng"].dropna().unique().tolist())
        selected_months = st.multiselect("Tháng:", months, format_func=lambda x: str(int(x)))
        if selected_months: temp_df = temp_df[temp_df["Tháng"].isin(selected_months)]
        
    with col_f3:
        units = sorted(temp_df["ĐƠN VỊ"].dropna().unique().tolist())
        selected_units = st.multiselect("🏥 Đơn Vị:", units)
        
    with col_f4:
        date_range = st.date_input("Khoảng thời gian (Ngày):", value=[])
        
elif page == "Hoạt động đào tạo":
    temp_st = df_students.copy()
    col_st1, col_st2 = st.columns(2)
    with col_st1:
        units = sorted(temp_st["Đơn vị đào tạo"].dropna().unique().tolist())
        selected_units = st.multiselect("🏥 Đơn Vị:", units)
    
    with col_st2:
        date_range = st.date_input("Khoảng thời gian (Ngày):", value=[])

# --- FILTER DATA ---
filtered_sessions = df_sessions.copy()
filtered_students = df_students.copy()

if selected_years:
    filtered_sessions = filtered_sessions[filtered_sessions["Năm"].isin(selected_years)]
if selected_months:
    filtered_sessions = filtered_sessions[filtered_sessions["Tháng"].isin(selected_months)]
    if not filtered_students.empty:
        filtered_students = filtered_students[filtered_students["Tháng thực hiện"].isin(selected_months)]
if selected_units:
    filtered_sessions = filtered_sessions[filtered_sessions["ĐƠN VỊ"].isin(selected_units)]
    if not filtered_students.empty:
        filtered_students = filtered_students[filtered_students["Đơn vị đào tạo"].isin(selected_units)]

# Apply Date Filter
if 'date_range' in locals() and date_range:
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
        filtered_sessions = filtered_sessions[(filtered_sessions['FullDate'].isna()) | ((filtered_sessions['FullDate'] >= start_date) & (filtered_sessions['FullDate'] <= end_date))]
        if not filtered_students.empty:
            filtered_students = filtered_students[(filtered_students['FullDate'].isna()) | ((filtered_students['FullDate'] >= start_date) & (filtered_students['FullDate'] <= end_date))]
    elif isinstance(date_range, tuple) and len(date_range) == 1:
        start_date = pd.to_datetime(date_range[0])
        filtered_sessions = filtered_sessions[(filtered_sessions['FullDate'].isna()) | (filtered_sessions['FullDate'] == start_date)]
        if not filtered_students.empty:
            filtered_students = filtered_students[(filtered_students['FullDate'].isna()) | (filtered_students['FullDate'] == start_date)]

time_info_str = ""
if 'date_range' in locals() and date_range:
    if isinstance(date_range, tuple) and len(date_range) == 2:
        time_info_str = f"Từ ngày {date_range[0].strftime('%d/%m/%Y')} đến {date_range[1].strftime('%d/%m/%Y')}"
    elif isinstance(date_range, tuple) and len(date_range) == 1:
        time_info_str = f"Ngày {date_range[0].strftime('%d/%m/%Y')}"
    else:
        time_info_str = f"Ngày {date_range.strftime('%d/%m/%Y')}"
elif selected_months:
    time_info_str = "Tháng " + ", ".join(map(str, sorted(selected_months)))
    if selected_years:
        time_info_str += f" năm {', '.join(map(str, sorted(selected_years)))}"
else:
    time_info_str = "Tất cả thời gian"

# =======================================================
# PAGE RENDERING
# =======================================================

if page == "Tổng quan":
    total_programs = len(filtered_sessions)
    total_participants = filtered_sessions["Số học viên tham gia"].sum()
    total_passed = filtered_sessions["Số Học viên đạt"].sum()

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    col1.metric("TỔNG SỐ HOẠT ĐỘNG", total_programs)
    col2.metric("TỔNG SỐ HỌC VIÊN THAM GIA", f"{total_participants:,.0f}")
    col3.metric("TỔNG SỐ HỌC VIÊN ĐẠT", f"{total_passed:,.0f}")

    st.markdown("---")
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.subheader("Tỷ lệ Đạt / Không đạt")
        if total_participants > 0:
            khong_dat = max(0, total_participants - total_passed)
            fig1 = px.pie(names=["Đạt", "Không đạt/Chưa rõ"], values=[total_passed, khong_dat],
                          color_discrete_sequence=["#2ecc71", "#e74c3c"], hole=0.4)
            fig1.update_traces(textinfo='percent', hovertemplate='<b>%{label}</b><br>%{value}<extra></extra>')
            fig1.update_layout(margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu học viên tham gia.")

    with col_chart2:
        st.subheader("Top 10 chương trình có nhiều học viên")
        if not filtered_sessions.empty:
            df_nd = filtered_sessions.groupby("NỘI DUNG/CHƯƠNG TRÌNH")[["Số học viên tham gia", "Số Học viên đạt"]].sum().reset_index()
            df_nd = df_nd.sort_values(by="Số học viên tham gia", ascending=True).tail(10)
            
            def truncate_name(name): return name if len(str(name)) <= 40 else str(name)[:37] + '...'
            df_nd["Tên Ngắn"] = df_nd["NỘI DUNG/CHƯƠNG TRÌNH"].apply(truncate_name)
            
            df_nd_melt = df_nd.melt(id_vars=["Tên Ngắn", "NỘI DUNG/CHƯƠNG TRÌNH"], value_vars=["Số học viên tham gia", "Số Học viên đạt"],
                                    var_name="Phân loại", value_name="Số lượng")
            fig2 = px.bar(df_nd_melt, y="NỘI DUNG/CHƯƠNG TRÌNH", x="Số lượng", color="Phân loại", orientation='h', barmode="group",
                          color_discrete_map={"Số học viên tham gia": "#3498db", "Số Học viên đạt": "#2ecc71"},
                          hover_data={"Tên Ngắn": False})
            fig2.update_layout(yaxis_title=None, xaxis_title=None, margin=dict(l=10, r=20, t=10, b=40),
                               legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig2.update_yaxes(tickmode='array', tickvals=df_nd["NỘI DUNG/CHƯƠNG TRÌNH"], ticktext=df_nd["Tên Ngắn"])
            st.plotly_chart(fig2, use_container_width=True)

elif page == "Hoạt động khoa phòng":
    st.subheader("📋 Danh sách nội dung đã tổ chức")
    if not filtered_sessions.empty:
        df_dept = filtered_sessions.copy()
        
        def format_time(row):
            d, m, y = row.get("Ngày"), row.get("Tháng"), row.get("Năm")
            try:
                if pd.notna(d) and pd.notna(m) and pd.notna(y):
                    return f"{int(float(str(d)))}/{int(float(str(m)))}/{int(float(str(y)))}"
            except:
                pass
            if pd.notna(d) and pd.notna(m) and pd.notna(y):
                return f"{d}/{m}/{y}"
            return ""
            
        df_dept["Thời gian"] = df_dept.apply(format_time, axis=1)
        
        cols_to_show = {
            "NỘI DUNG/CHƯƠNG TRÌNH": "Tên nội dung",
            "Thời gian": "Thời gian",
            "NGƯỜI PHỤ TRÁCH NỘI DUNG": "Người phụ trách",
            "Số học viên tham gia": "Tham gia",
            "Số Học viên đạt": "Đạt",
            "Gắn link Điểm danh và post test": "Điểm danh",
            "Gắn link bảng kiểm đào tạo nội bộ": "Bảng kiểm",
            "Gắn link bài báo cáo": "Báo cáo",
            "Gắn link video buổi đào tạo": "Video"
        }
        
        # Bổ sung các cột link nếu thiếu
        for c in list(cols_to_show.keys())[5:]:
            if c not in df_dept.columns:
                df_dept[c] = False
                
        df_dept = df_dept[list(cols_to_show.keys())].rename(columns=cols_to_show)
        
        # Tích hợp Ô tìm kiếm/lọc
        col_search1, col_search2 = st.columns([3, 1])
        with col_search1:
            search_query = st.text_input("🔍 Tìm kiếm nhanh (nhập Tên nội dung, Người phụ trách, hoặc Thời gian):", "")
        with col_search2:
            sort_col = st.selectbox("Sắp xếp theo:", ["Mặc định"] + list(cols_to_show.values()))
            
        if search_query:
            search_query_lower = search_query.lower()
            df_dept = df_dept[
                df_dept["Tên nội dung"].astype(str).str.lower().str.contains(search_query_lower, na=False, regex=False) |
                df_dept["Người phụ trách"].astype(str).str.lower().str.contains(search_query_lower, na=False, regex=False) |
                df_dept["Thời gian"].astype(str).str.lower().str.contains(search_query_lower, na=False, regex=False)
            ]
            
        if sort_col != "Mặc định":
            df_dept = df_dept.sort_values(by=sort_col, ascending=True)
            
        # Tính toán tổng số lượng của dữ liệu sau khi lọc để hiển thị cố định ở dưới bảng
        total_acts = len(df_dept)
        total_participants = int(pd.to_numeric(df_dept["Tham gia"], errors='coerce').sum())
        total_passed = int(pd.to_numeric(df_dept["Đạt"], errors='coerce').sum())
        
        df_dept["Tham gia"] = pd.to_numeric(df_dept["Tham gia"], errors='coerce').apply(lambda x: f"{int(x)}" if pd.notna(x) else "")
        df_dept["Đạt"] = pd.to_numeric(df_dept["Đạt"], errors='coerce').apply(lambda x: f"{int(x)}" if pd.notna(x) else "")
        
        def style_link_col(val, c, row):
            if pd.isna(val) or str(val).strip() in ['', 'False', 'nan']: return 'Không'
            val_str = str(val).lower()
            if 'chưa' in val_str or 'không' in val_str or 'none' in val_str: return 'Không'
            
            if c == "Điểm danh":
                tham_gia = str(row['Tham gia']).strip()
                if tham_gia == '' or tham_gia == '0':
                    return 'Lỗi (Trống Data)'
            return 'Có'
            
        def color_yes_no(val):
            if val == 'Có': return 'background-color: #d4edda; color: #155724;'
            if val == 'Không': return 'background-color: #f8d7da; color: #721c24;'
            if 'Lỗi' in str(val): return 'background-color: #fff3cd; color: #856404; font-weight: bold;'
            return ''
            
        df_dept_display = df_dept.copy()
        df_dept_display.index = range(1, len(df_dept_display) + 1)
        for c in ["Điểm danh", "Bảng kiểm", "Báo cáo", "Video"]:
            # Need to pass original row 'Tham gia' value
            df_dept_display[c] = df_dept.apply(lambda r: style_link_col(r[c], c, r), axis=1).values
            
        dept_styler = df_dept_display.style.map(color_yes_no, subset=["Điểm danh", "Bảng kiểm", "Báo cáo", "Video"])
            
        col_btn1, col_btn2 = st.columns([8, 2])
        with col_btn2:
            btn_dept = st.button("📸 Xem ảnh báo cáo", key="btn_dept", use_container_width=True)
        if btn_dept:
            styler_img = dept_styler.set_caption(f"Báo cáo Hoạt động Khoa phòng - {time_info_str}").set_table_styles([{
                'selector': 'caption',
                'props': [('font-size', '18px'), ('font-weight', 'bold'), ('color', 'black'), ('text-align', 'center'), ('margin-bottom', '10px')]
            }])
            show_image_modal(styler_img)
        st.dataframe(dept_styler, use_container_width=True, hide_index=True)
        
        # Hiển thị dòng tổng cộng cố định ở dưới bảng bằng markdown để không bị xáo trộn khi sort cột
        st.markdown(
            f"<div style='background-color:#f1f2f6; padding:10px; border-radius:5px; font-weight:bold; font-size:15px; color:#2f3542; margin-top:5px;'>"
            f"📊 TỔNG CỘNG: {total_acts} hoạt động &nbsp;&nbsp;|&nbsp;&nbsp; 👥 Tham gia: {total_participants:,.0f} &nbsp;&nbsp;|&nbsp;&nbsp; Đạt: {total_passed:,.0f}"
            f"</div>", 
            unsafe_allow_html=True
        )
    else:
        st.info("Không có dữ liệu.")

elif page == "Hoạt động đào tạo":
    st.subheader("👩‍🎓 Danh sách học viên")
    if not filtered_students.empty:
        # Group by employee to count sessions
        df_grp = filtered_students.groupby('Mã_Định_Danh').agg(
            Họ_và_Tên=('Họ Tên Đầy Đủ', 'first'),
            Mã_NV=('Mã nhân sự', 'first'),
            Đơn_vị=('Đơn vị đào tạo', 'first'),
            Số_buổi_tham_gia=('Mã_Định_Danh', 'count'),
            Số_buổi_Đạt=('Trạng thái', lambda x: (x == 'Đạt').sum())
        ).sort_values('Số_buổi_tham_gia', ascending=False).reset_index()
        
        # Remove empty Mã NV (NaN -> "")
        df_grp['Mã_NV'] = df_grp['Mã_NV'].fillna("")
        df_grp['Đơn_vị'] = df_grp['Đơn_vị'].fillna("")
        
        display_df = df_grp[['Họ_và_Tên', 'Mã_NV', 'Đơn_vị', 'Số_buổi_tham_gia', 'Số_buổi_Đạt']].copy()
        
        # Tích hợp Ô tìm kiếm học viên
        search_student = st.text_input("🔍 Tìm kiếm học viên (nhập Họ tên, Mã NV, hoặc Đơn vị):", "")
        if search_student:
            search_st_lower = search_student.lower()
            df_grp = df_grp[
                df_grp["Họ_và_Tên"].astype(str).str.lower().str.contains(search_st_lower, na=False, regex=False) |
                df_grp["Mã_NV"].astype(str).str.lower().str.contains(search_st_lower, na=False, regex=False) |
                df_grp["Đơn_vị"].astype(str).str.lower().str.contains(search_st_lower, na=False, regex=False)
            ].reset_index(drop=True)
            display_df = df_grp[['Họ_và_Tên', 'Mã_NV', 'Đơn_vị', 'Số_buổi_tham_gia', 'Số_buổi_Đạt']]
            
        total_students = len(display_df)
        total_sessions_all = display_df['Số_buổi_tham_gia'].sum()
        total_sessions_passed = display_df['Số_buổi_Đạt'].sum()
        display_df.index = range(1, len(display_df) + 1)
        
        # Interactive table (không đưa dòng TỔNG CỘNG vào dataframe hiển thị để tránh lỗi chọn hàng và sort)
        col_btn1, col_btn2 = st.columns([8, 2])
        with col_btn2:
            btn_main = st.button("📸 Xem ảnh báo cáo", key="btn_main", use_container_width=True)
        if btn_main:
            styler = display_df.style.set_caption(f"Danh sách buổi sinh hoạt chuyên môn - {time_info_str}").set_table_styles([{
                'selector': 'caption',
                'props': [('font-size', '18px'), ('font-weight', 'bold'), ('color', 'black'), ('text-align', 'center'), ('margin-bottom', '10px')]
            }])
            show_image_modal(styler)
        event = st.dataframe(display_df, use_container_width=True, hide_index=True, 
                             selection_mode="single-row", on_select="rerun")
        
        # Hiển thị dòng tổng cộng cố định ở dưới bảng
        st.markdown(
            f"<div style='background-color:#f1f2f6; padding:10px; border-radius:5px; font-weight:bold; font-size:15px; color:#2f3542; margin-top:5px; margin-bottom:15px;'>"
            f"📊 TỔNG CỘNG: {total_students} học viên &nbsp;&nbsp;|&nbsp;&nbsp; 📚 Tổng số lượt học: {total_sessions_all} &nbsp;&nbsp;|&nbsp;&nbsp; Đạt: {total_sessions_passed}"
            f"</div>", 
            unsafe_allow_html=True
        )
        
        if event and len(event.selection.rows) > 0:
            selected_idx = event.selection.rows[0]
            if selected_idx < len(df_grp):  # Bảo đảm index hợp lệ
                selected_id = df_grp.iloc[selected_idx]['Mã_Định_Danh']
                selected_name = df_grp.iloc[selected_idx]['Họ_và_Tên']
                
                st.markdown("---")
                st.subheader(f"📄 Chi tiết đào tạo: {selected_name}")
                
                detail_df = filtered_students[filtered_students['Mã_Định_Danh'] == selected_id].copy()
                
                def format_time_st(row):
                    d, m = row.get("Ngày đào tạo"), row.get("Tháng thực hiện")
                    try:
                        if pd.notna(d) and pd.notna(m):
                            if hasattr(d, 'day'):
                                d_val = d.day
                            else:
                                d_val = int(float(str(d).strip()))
                            
                            if hasattr(m, 'month'):
                                m_val = m.month
                            else:
                                m_val = int(float(str(m).strip()))
                            
                            return f"{d_val}/{m_val}"
                    except:
                        pass
                    if pd.notna(d) and pd.notna(m):
                        d_str = str(d).split('.')[0].strip()
                        m_str = str(m).split('.')[0].strip()
                        return f"{d_str}/{m_str}"
                    return ""
                detail_df["Ngày"] = detail_df.apply(format_time_st, axis=1)
                
                detail_df["Điểm_Quy_Đổi"] = detail_df["Điểm_Quy_Đổi"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "")
                
                detail_display = detail_df[['Tên chương trình', 'Ngày', 'Điểm_Quy_Đổi', 'Trạng thái']].rename(columns={
                    'Điểm_Quy_Đổi': 'Điểm số'
                })
                
                def color_status(val):
                    if val == 'Đạt': return 'background-color: #d4edda; color: #155724; font-weight: bold;'
                    if val == 'Không đạt': return 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                    return ''
                
                # Thêm dòng tổng cộng cho bảng chi tiết của học viên này
                total_detail_row = pd.DataFrame([{
                    'Tên chương trình': f'TỔNG CỘNG: {len(detail_display)} buổi học',
                    'Ngày': '',
                    'Điểm số': '',
                    'Trạng thái': f'Đạt { (detail_display["Trạng thái"] == "Đạt").sum() } / {len(detail_display)}'
                }])
                detail_display_with_total = pd.concat([detail_display, total_detail_row], ignore_index=True)
                detail_display_with_total.index = range(1, len(detail_display_with_total) + 1)
                    
                col_btn1, col_btn2 = st.columns([8, 2])
                with col_btn2:
                    btn_detail = st.button("📸 Xem ảnh báo cáo", key="btn_detail", use_container_width=True)
                if btn_detail:
                    styler = detail_display_with_total.style.map(color_status, subset=['Trạng thái']).set_caption(f"Chi tiết học viên: {selected_name} - {time_info_str}").set_table_styles([{
                        'selector': 'caption',
                        'props': [('font-size', '16px'), ('font-weight', 'bold'), ('color', 'black'), ('text-align', 'center'), ('margin-bottom', '10px')]
                    }])
                    show_image_modal(styler)
                st.dataframe(detail_display_with_total.style.map(color_status, subset=['Trạng thái']), use_container_width=True, hide_index=True)
    else:
        st.info("Không có dữ liệu học viên.")

elif page == "Đánh giá chỉ tiêu":
    st.subheader("🎯 Bảng đánh giá chỉ tiêu theo Đơn vị")
    if not filtered_sessions.empty:
        # Số tháng được chọn
        num_months = len(selected_months) if selected_months else 1
        
        col_ct1, col_ct2 = st.columns([1, 4])
        with col_ct1:
            chi_tieu_co_ban = st.number_input("Chỉ tiêu (buổi):", min_value=1, max_value=100, value=4 * num_months, step=1)
            
        st.caption("Chỉ tiêu mặc định là 1 buổi/tuần (4 buổi/tháng). Nếu bạn chọn nhiều tháng, chỉ tiêu sẽ tự động nhân lên.")
        
        # Nhóm theo đơn vị và tình trạng
        df_eval = filtered_sessions.groupby(['ĐƠN VỊ', 'Tình trạng']).size().unstack(fill_value=0)
        
        # Đảm bảo có đủ cột Thực hiện và Không thực hiện
        if 'Thực hiện' not in df_eval.columns: df_eval['Thực hiện'] = 0
        if 'Không thực hiện' not in df_eval.columns: df_eval['Không thực hiện'] = 0
            
        # Tính tổng cộng và chỉ tiêu
        df_eval['Tổng cộng'] = df_eval['Thực hiện'] + df_eval['Không thực hiện']
        if 'Chưa cập nhật' in df_eval.columns:
            df_eval['Tổng cộng'] += df_eval['Chưa cập nhật']
            
        df_eval['Chỉ tiêu'] = chi_tieu_co_ban
        df_eval['Đạt chỉ tiêu'] = df_eval.apply(lambda row: 'Đạt' if row['Thực hiện'] >= row['Chỉ tiêu'] else 'Không đạt', axis=1)
        
        df_eval = df_eval.reset_index()
        df_eval = df_eval.rename(columns={'ĐƠN VỊ': 'TÊN ĐƠN VỊ'})
        
        # Chọn và sắp xếp lại cột
        cols = ['TÊN ĐƠN VỊ', 'Không thực hiện', 'Thực hiện', 'Tổng cộng', 'Chỉ tiêu', 'Đạt chỉ tiêu']
        df_eval = df_eval[cols]
        
        # Tính dòng Tổng cộng cuối cùng
        total_khong_thuc_hien = df_eval['Không thực hiện'].sum()
        total_thuc_hien = df_eval['Thực hiện'].sum()
        total_tong_cong = df_eval['Tổng cộng'].sum()
        
        # Thêm dòng tổng cộng
        total_row = pd.DataFrame([{
            'TÊN ĐƠN VỊ': 'Tổng cộng',
            'Không thực hiện': total_khong_thuc_hien,
            'Thực hiện': total_thuc_hien,
            'Tổng cộng': total_tong_cong,
            'Chỉ tiêu': '',
            'Đạt chỉ tiêu': ''
        }])
        
        df_eval_display = pd.concat([df_eval, total_row], ignore_index=True)
        df_eval_display.index = range(1, len(df_eval_display) + 1)
        
        # Hàm hỗ trợ GitHub
        OVERRIDE_FILE = "evaluation_overrides.json"
        
        def get_github_repo():
            try:
                if Github is not None and getattr(st, "secrets", None) is not None:
                    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
                        g = Github(st.secrets["GITHUB_TOKEN"])
                        return g.get_repo(st.secrets["GITHUB_REPO"])
            except:
                pass
            return None

        def load_overrides():
            if not os.path.exists(OVERRIDE_FILE):
                repo = get_github_repo()
                if repo:
                    try:
                        contents = repo.get_contents(OVERRIDE_FILE)
                        data = json.loads(base64.b64decode(contents.content).decode('utf-8'))
                        with open(OVERRIDE_FILE, "w", encoding="utf-8") as f:
                            json.dump(data, f, ensure_ascii=False, indent=2)
                        return data
                    except:
                        pass
                return {}
            try:
                with open(OVERRIDE_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}

        def save_overrides(data):
            with open(OVERRIDE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            repo = get_github_repo()
            if repo:
                try:
                    content_str = json.dumps(data, ensure_ascii=False, indent=2)
                    try:
                        contents = repo.get_contents(OVERRIDE_FILE)
                        repo.update_file(contents.path, "Update evaluation overrides", content_str, contents.sha)
                    except:
                        repo.create_file(OVERRIDE_FILE, "Create evaluation overrides", content_str)
                except Exception as e:
                    st.warning(f"Không thể đồng bộ GitHub: {e}")

        # Tải cấu hình đã lưu
        overrides = load_overrides()
        
        # Xác định key_prefix cho cấu hình lưu
        if 'date_range' in locals() and date_range:
            key_prefix = f"date_{date_range}"
        elif selected_months:
            key_prefix = f"months_{'_'.join(map(str, sorted(selected_months)))}"
        else:
            key_prefix = "All"
            
        current_overrides = overrides.get(key_prefix, {})
        
        # Áp dụng overrides vào df_eval_display
        df_eval_display['Ghi chú'] = ""
        for idx, row in df_eval_display.iterrows():
            unit_name = row['TÊN ĐƠN VỊ']
            if unit_name in current_overrides:
                if 'Đạt chỉ tiêu' in current_overrides[unit_name]:
                    df_eval_display.at[idx, 'Đạt chỉ tiêu'] = current_overrides[unit_name]['Đạt chỉ tiêu']
                if 'Ghi chú' in current_overrides[unit_name]:
                    df_eval_display.at[idx, 'Ghi chú'] = current_overrides[unit_name]['Ghi chú']

        # Xóa số 0 hiển thị thành ô trống cho giống bảng Excel
        for c in ['Không thực hiện', 'Thực hiện', 'Tổng cộng', 'Chỉ tiêu']:
            df_eval_display[c] = df_eval_display[c].astype(str)
            df_eval_display[c] = df_eval_display[c].apply(lambda x: "" if x == "0" or x == "0.0" else x)
            
        # Format style to highlight "Không đạt"
        def color_dat(val):
            if val == 'Đạt':
                return 'background-color: #d4edda; color: #155724;'
            elif val == 'Không đạt':
                return 'background-color: #f8d7da; color: #721c24;'
            return ''
            
        col_btn1, col_btn2 = st.columns([8, 2])
        with col_btn2:
            btn_eval = st.button("📸 Xem ảnh báo cáo", key="btn_eval", use_container_width=True)
            
        eval_styler = df_eval_display.style.map(color_dat, subset=['Đạt chỉ tiêu'])
        if btn_eval:
            styler_img = eval_styler.set_caption(f"Bảng Đánh giá chỉ tiêu - {time_info_str}").set_table_styles([{
                'selector': 'caption',
                'props': [('font-size', '18px'), ('font-weight', 'bold'), ('color', 'black'), ('text-align', 'center'), ('margin-bottom', '10px')]
            }])
            show_image_modal(styler_img)
            
        edited_df = st.data_editor(
            eval_styler, 
            use_container_width=True, 
            hide_index=True, 
            height=800,
            column_config={
                "Đạt chỉ tiêu": st.column_config.SelectboxColumn(
                    "Đạt chỉ tiêu",
                    options=["Đạt", "Không đạt"],
                    required=True
                )
            },
            disabled=['TÊN ĐƠN VỊ', 'Không thực hiện', 'Thực hiện', 'Tổng cộng', 'Chỉ tiêu'],
            key="eval_editor"
        )
        
        # Check for changes and save
        has_changes = False
        for idx, row in edited_df.iterrows():
            unit_name = row['TÊN ĐƠN VỊ']
            if unit_name == 'Tổng cộng': continue
            orig_row = df_eval_display.loc[idx]
            if row['Đạt chỉ tiêu'] != orig_row['Đạt chỉ tiêu'] or row['Ghi chú'] != orig_row['Ghi chú']:
                has_changes = True
                if unit_name not in current_overrides: current_overrides[unit_name] = {}
                current_overrides[unit_name]['Đạt chỉ tiêu'] = row['Đạt chỉ tiêu']
                current_overrides[unit_name]['Ghi chú'] = row['Ghi chú']
                
        if has_changes:
            overrides[key_prefix] = current_overrides
            try:
                save_overrides(overrides)
                st.success("Đã lưu các thay đổi Ghi chú và Đạt chỉ tiêu!")
                st.rerun()
            except Exception as e:
                st.error(f"Không thể lưu thay đổi: {e}")
    else:
        st.info("Không có dữ liệu.")

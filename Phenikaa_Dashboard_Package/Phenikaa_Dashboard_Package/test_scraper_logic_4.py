import pandas as pd
import io
import requests

def normalize_text_std(text):
    if pd.isna(text): return ""
    return str(text).lower().strip()

def find_header_and_data(df):
    best_row_idx = 0
    max_matches = 0
    keywords = {'dấu thời gian', 'timestamp', 'điểm số', 'score', 'họ tên', 'họ và tên', 'tên', 
                'mã nhân viên', 'mã nhân sự', 'mã nv', 'mã ns', 'email', 'địa chỉ email'}
    
    for idx, row in df.iterrows():
        row_str = [str(val).lower() for val in row.values]
        matches = sum(1 for kw in keywords if any(kw in rv for rv in row_str))
        if matches > max_matches:
            max_matches = matches
            best_row_idx = idx
            
    if max_matches >= 2:
        columns = df.iloc[best_row_idx].values
        cleaned_columns = []
        for c in columns:
            c_str = str(c).strip()
            if not c_str or c_str.lower() == 'nan':
                cleaned_columns.append(f"Unnamed_{len(cleaned_columns)}")
            else:
                cleaned_columns.append(c_str)
        df_clean = df.iloc[best_row_idx+1:].copy()
        df_clean.columns = cleaned_columns
        return df_clean
    return df

def map_columns(df):
    cols = df.columns
    mapped = {'name_col': None, 'id_col': None, 'score_col': None, 'email_col': None}
    
    for col in cols:
        col_norm = normalize_text_std(col)
        if not mapped['name_col']:
            if col_norm in ('họ tên', 'họ và tên', 'họ tên đầy đủ', 'họ và tên đầy đủ', 'tên', 'name', 'nhân sự', 'nguyễn văn a', 'họ tên?'):
                mapped['name_col'] = col
            elif 'họ' in col_norm and 'tên' in col_norm:
                mapped['name_col'] = col
        if not mapped['id_col']:
            if col_norm in ('mã nhân sự', 'mã nhân viên', 'mã nv', 'mã ns', 'msnv', 'msns', 'id', 'mã số nhân viên'):
                mapped['id_col'] = col
            elif 'mã' in col_norm and ('nhân viên' in col_norm or 'nhân sự' in col_norm or 'nv' in col_norm or 'ns' in col_norm):
                mapped['id_col'] = col
        if not mapped['score_col']:
            if col_norm in ('điểm số', 'điểm', 'score', 'kết quả'):
                mapped['score_col'] = col
            elif 'điểm' in col_norm or 'score' in col_norm:
                mapped['score_col'] = col
        if not mapped['email_col']:
            if col_norm in ('email', 'địa chỉ email', 'email address'):
                mapped['email_col'] = col
    return mapped

url = "https://1drv.ms/x/c/e8b9546bb527d83d/IQCopwlj8hElQ7LzIUB0hnS7AeSbGjx5yeY0XiOeq5n04yk?e=3kodqM&download=1"
resp = requests.get(url)
print("Status Code:", resp.status_code)
df_dl = pd.read_excel(io.BytesIO(resp.content))

df_data = find_header_and_data(df_dl)
print("Headers:", df_data.columns.tolist()[:5])

col_map = map_columns(df_data)
print("Col map:", col_map)

records = []
for _, row in df_data.iterrows():
    name = row.get(col_map['name_col']) if col_map['name_col'] else None
    if pd.isna(name): continue
    records.append(name)
    
print(f"Found {len(records)} students")

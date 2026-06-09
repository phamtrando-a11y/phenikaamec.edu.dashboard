import pandas as pd
import io
import requests

from google_drive_scraper import find_header_and_data, map_columns, get_download_url

url = "https://1drv.ms/x/c/e8b9546bb527d83d/IQCopwlj8hElQ7LzIUB0hnS7AeSbGjx5yeY0XiOeq5n04yk?e=3kodqM"
dl_url, _ = get_download_url(url)
print("DL URL:", dl_url)

resp = requests.get(dl_url)
print("Status Code:", resp.status_code)
df_dl = pd.read_excel(io.BytesIO(resp.content))

df_data = find_header_and_data(df_dl)
col_map = map_columns(df_data)
print("Col map:", col_map)

records = []
for _, row in df_data.iterrows():
    name = row.get(col_map['name_col']) if col_map['name_col'] else None
    if pd.isna(name): continue
    records.append(name)
    
print(f"Found {len(records)} students")

import pandas as pd
import io
import requests

from google_drive_scraper import find_header_and_data, map_columns

url = "https://1drv.ms/x/c/e8b9546bb527d83d/IQCopwlj8hElQ7LzIUB0hnS7AeSbGjx5yeY0XiOeq5n04yk?e=3kodqM&download=1"
resp = requests.get(url)
print("Status Code:", resp.status_code)
df_dl = pd.read_excel(io.BytesIO(resp.content))

df_data, headers, header_idx = find_header_and_data(df_dl)
print("Header idx:", header_idx)
print("Headers:", headers[:5])

col_map = map_columns(headers)
print("Col map:", col_map)

records = []
for _, row in df_data.iterrows():
    name = row.get(col_map['name_col']) if col_map['name_col'] else None
    if pd.isna(name): continue
    records.append(name)
    
print(f"Found {len(records)} students")

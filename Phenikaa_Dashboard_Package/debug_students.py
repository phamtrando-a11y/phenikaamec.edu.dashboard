import pandas as pd
import json
import requests
import io
import os
import sys

from app import load_data, is_passed

df_sessions, df_students_full, df_students_scraped = load_data()

print("Scraped student columns:", df_students_scraped.columns.tolist() if not df_students_scraped.empty else "Empty")
if not df_students_scraped.empty:
    print("Unique programs in scraped data:")
    for prog in df_students_scraped['Tên chương trình'].unique():
        if "cổ tay" in str(prog).lower() or "enzyme" in str(prog).lower() or "dịch tễ" in str(prog).lower() or "nhóm máu" in str(prog).lower() or "q.clear" in str(prog).lower():
            print(f" - '{prog}'")

print("\nAll programs in sessions matching criteria:")
for s in df_sessions['NỘI DUNG/CHƯƠNG TRÌNH'].dropna().unique():
    s_low = str(s).lower()
    if "cổ tay" in s_low or "enzyme" in s_low or "dịch tễ" in s_low or "nhóm máu" in s_low or "q.clear" in s_low:
        print(f" -> Found in sessions: '{s}'")

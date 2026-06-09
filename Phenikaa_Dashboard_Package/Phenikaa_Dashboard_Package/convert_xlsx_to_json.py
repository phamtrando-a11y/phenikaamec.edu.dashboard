import os
import json
import pandas as pd
import numpy as np

def convert():
    dir_path = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(dir_path, "250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM_Độ.xlsx")
    json_path = os.path.join(dir_path, "netlify-dashboard", "data.json")

    print(f"Reading Excel file from {excel_path}...")
    df = pd.read_excel(excel_path, sheet_name="Append1")
    
    print("Cleaning data...")
    # Drop rows where 'ĐƠN VỊ' is null
    df = df.dropna(subset=["ĐƠN VỊ"])
    
    # Replace NaN with None so it becomes null in JSON
    df = df.replace({np.nan: None})
    
    # Convert dataframe to dict list
    records = df.to_dict(orient="records")
    
    print(f"Writing {len(records)} records to JSON at {json_path}...")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
        
    print("Successfully updated data.json!")

if __name__ == "__main__":
    convert()

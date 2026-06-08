import os
import pandas as pd

folder = "/Users/dotran/Library/CloudStorage/OneDrive-Personal/Phenikaa/SHCM các phòng ban"
search_str = "Dinh dưỡng cho người bệnh có bệnh lý đường tiêu hóa (Câu trả lời).xlsx"

print(f"Đang tìm kiếm '{search_str}' trong các file Excel...")
found = False

for root, dirs, files in os.walk(folder):
    for f in files:
        if f.endswith('.xlsx') and not f.startswith('~'):
            path = os.path.join(root, f)
            try:
                xls = pd.ExcelFile(path)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    # Convert to string and search
                    mask = df.astype(str).apply(lambda x: x.str.contains(search_str, case=False, na=False))
                    if mask.any().any():
                        print(f"\n✅ ĐÃ TÌM THẤY TẠI:")
                        print(f"📂 File: {path}")
                        print(f"📄 Sheet: {sheet}")
                        found = True
            except Exception as e:
                # Bỏ qua các lỗi đọc file (VD: file bị khóa hoặc hỏng)
                pass

if not found:
    print(f"\n❌ Không tìm thấy nội dung này trong bất kỳ file Excel nào.")
print("\nHoàn tất!")

import re

with open('google_drive_scraper.py', 'r') as f:
    content = f.read()

# 1. Replace Bước 1 to use the conversion trick
old_step1 = """    print("--- BƯỚC 1: ĐỌC FILE EXCEL CHÍNH TỪ GOOGLE DRIVE ---")
    request = drive_service.files().get_media(fileId=main_excel_id)
    main_file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(main_file_stream, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    main_file_stream.seek(0)"""

new_step1 = """    print("--- BƯỚC 1: ĐỌC FILE EXCEL CHÍNH TỪ GOOGLE DRIVE (KÈM CONVERT ĐỂ LẤY SMART CHIP) ---")
    try:
        # 1. Tạo bản sao tạm thời dưới định dạng Google Sheets để ép Google Drive render Smart Chips
        file_metadata = {
            'name': 'Temp_Dashboard_Sheet_For_Scraping',
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        temp_sheet = drive_service.files().copy(fileId=main_excel_id, body=file_metadata).execute()
        temp_id = temp_sheet['id']
        
        # 2. Export file Google Sheets tạm thời này ngược lại thành .xlsx
        # Quá trình này sẽ ép Google Sheets chuyển đổi Smart Chips thành Hyperlink tiêu chuẩn của Excel
        request = drive_service.files().export_media(
            fileId=temp_id, 
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        main_file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(main_file_stream, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            
        # 3. Xóa file tạm
        try:
            drive_service.files().delete(fileId=temp_id).execute()
        except:
            pass
            
        main_file_stream.seek(0)
    except Exception as e:
        print(f"Lỗi khi convert qua Google Sheets, fallback về tải gốc: {e}")
        request = drive_service.files().get_media(fileId=main_excel_id)
        main_file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(main_file_stream, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        main_file_stream.seek(0)"""

content = content.replace(old_step1, new_step1)

# 2. Replace requests.get to use User-Agent to avoid OneDrive 403 blocks
old_req = "resp = requests.get(dl_url, timeout=20)"
new_req = """headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
                                resp = requests.get(dl_url, headers=headers, timeout=20)"""

content = content.replace(old_req, new_req)

with open('google_drive_scraper.py', 'w') as f:
    f.write(content)

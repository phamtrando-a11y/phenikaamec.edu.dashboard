import json
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import toml

with open('.streamlit/secrets.toml', 'r') as f:
    secrets = toml.load(f)
creds_json = secrets['GOOGLE_CREDENTIALS']
creds_dict = json.loads(creds_json)
creds = service_account.Credentials.from_service_account_info(
    creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
)
drive_service = build('drive', 'v3', credentials=creds)

main_excel_id = "10hYk1D02E0p93OIQxR0R5H_1n_lI2XwN"
request = drive_service.files().get_media(fileId=main_excel_id)
main_file_stream = io.BytesIO()
downloader = MediaIoBaseDownload(main_file_stream, request)
done = False
while done is False:
    status, done = downloader.next_chunk()
main_file_stream.seek(0)

with open('live_excel_real.xlsx', 'wb') as f:
    f.write(main_file_stream.read())
print("Downloaded live excel! Size:", len(main_file_stream.getvalue()))

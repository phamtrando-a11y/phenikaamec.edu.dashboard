import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
import toml

# We need to test if Sheets API can extract from main_excel_id
with open('.streamlit/secrets.toml', 'r') as f:
    secrets = toml.load(f)
creds_dict = json.loads(secrets['GOOGLE_CREDENTIALS'])
creds = service_account.Credentials.from_service_account_info(
    creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
)
sheets_service = build('sheets', 'v4', credentials=creds)

main_excel_id = "10hYk1D02E0p93OIQxR0R5H_1n_lI2XwN"

try:
    result = sheets_service.spreadsheets().get(
        spreadsheetId=main_excel_id,
        ranges=['SHCM - Độ!T58:T62'],
        fields="sheets.data.rowData.values.hyperlink"
    ).execute()
    print("Sheets API Success:")
    print(json.dumps(result, indent=2))
except Exception as e:
    print("Sheets API Error:")
    print(e)

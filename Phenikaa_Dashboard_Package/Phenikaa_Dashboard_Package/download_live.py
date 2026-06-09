import requests

main_excel_id = "10hYk1D02E0p93OIQxR0R5H_1n_lI2XwN"
url = f"https://docs.google.com/spreadsheets/d/{main_excel_id}/export?format=xlsx"

resp = requests.get(url)
print("Downloaded size:", len(resp.content))

with open('live_excel.xlsx', 'wb') as f:
    f.write(resp.content)

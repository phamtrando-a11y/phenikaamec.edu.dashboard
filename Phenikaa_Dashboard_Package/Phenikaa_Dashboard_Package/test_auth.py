import requests

url = "https://docs.google.com/spreadsheets/d/12FpcGGKs4LdijihdYjQnS2qdf3RQW4-W/export?format=xlsx"
r = requests.get(url)
print(r.status_code)
print(r.headers.get('Content-Type'))

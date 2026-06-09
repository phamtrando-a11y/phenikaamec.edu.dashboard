import requests

url = "https://1drv.ms/x/c/e8b9546bb527d83d/IQCopwlj8hElQ7LzIUB0hnS7AeSbGjx5yeY0XiOeq5n04yk?e=3kodqM"
dl_url = url + "&download=1"

r = requests.get(dl_url, allow_redirects=True)
print("Status Code:", r.status_code)
print("Content-Type:", r.headers.get('Content-Type'))
print("Content size:", len(r.content))

with open("test_onedrive.xlsx", "wb") as f:
    f.write(r.content)

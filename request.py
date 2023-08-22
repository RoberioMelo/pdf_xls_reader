import requests

url = "https://api.pipe.run/v1/persons"

headers = {
    "accept": "application/json",
    "token": "06c3dcc312b890652c2ec5540a79ad7a"
}
params = {
    "with": "id,account_id,custom_fields"
}
response = requests.get(url, headers=headers, params=params)

print(response.text)
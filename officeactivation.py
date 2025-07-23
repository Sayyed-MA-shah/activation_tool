import requests

response = requests.post(
    "http://localhost:5000/get-key",
    headers={"Authorization": "Bearer secure123"},
    json={"type": "windows"}
)

print("Status Code:", response.status_code)
print("Response:", response.text)

import requests

url = "http://127.0.0.1:5000/batch_lookup"
data = {
    "items": [
        {
            "row": 1,
            "phrase": "condenser water pump roof",
            "current_match": "",
            "current_id": ""
        },
        {
            "row": 2,
            "phrase": "return fan 1",
            "current_match": "",
            "current_id": ""
        }
    ]
}

response = requests.post(url, json=data)
print(response.json())

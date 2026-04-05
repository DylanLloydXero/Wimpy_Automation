import requests
import json
import socket

# Before trying localhost:8000, wait, it might be running already by the bat file
payload = {
    "employees": [
        {
            "name": "SAYLIN",
            "reg_hours": 10.0,
            "sun_hours": 0.0,
            "hol_hours": 0.0,
            "leave_days": 0.0,
            "sick_days": 0.0,
            "bonus": "",
            "till_short": 0.0,
            "tips": 0.0,
            "clothing": 0.0,
            "rate": 30.33,
            "cell_number": None,
            "role": "WAITER"
        }
    ]
}

try:
    res = requests.post("http://127.0.0.1:8000/api/generate", json=payload)
    print(res.status_code)
    print(res.text)
except Exception as e:
    print(e)

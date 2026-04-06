import requests
import json
import os

BASE_URL = "http://127.0.0.1:8000"

def test_login():
    print("Testing Login...")
    payload = {"name": "Manager", "id_last4": "1234"}
    res = requests.post(f"{BASE_URL}/api/login", json=payload)
    print(f"Status: {res.status_code}, Response: {res.json()}")
    assert res.json().get("success") == True

def test_submit_request():
    print("\nTesting Submit Request...")
    # Test a weekday request (should be APPROVED)
    payload = {
        "employee": "Charnick Botha",
        "off_days": ["2026-05-20"], # A Wednesday
        "type": "OFF_DAY"
    }
    res = requests.post(f"{BASE_URL}/api/submit_request", json=payload)
    print(f"Weekday Request - Status: {res.status_code}, Response: {res.json()}")
    
    # Test a weekend request (should be PENDING_APPROVAL)
    payload = {
        "employee": "Charnick Botha",
        "off_days": ["2026-05-23"], # A Saturday
        "type": "OFF_DAY"
    }
    res = requests.post(f"{BASE_URL}/api/submit_request", json=payload)
    print(f"Weekend Request - Status: {res.status_code}, Response: {res.json()}")
    assert res.json().get("status") == "PENDING_APPROVAL"

def test_approve_request():
    print("\nTesting Approve Request...")
    # Get requests to find the pending one
    res = requests.get(f"{BASE_URL}/api/requests")
    pending = [r for r in res.json() if r["status"] == "PENDING_APPROVAL"]
    if not pending:
        print("No pending requests to approve.")
        return
    
    req_id = pending[0]["id"]
    payload = {"id": req_id, "action": "APPROVE"}
    res = requests.post(f"{BASE_URL}/api/approve_request", json=payload)
    print(f"Approve Request - Status: {res.status_code}, Response: {res.json()}")
    assert res.json().get("success") == True

def test_generate_snippets():
    print("\nTesting Generate Snippets...")
    payload = {
        "week_date": "2026-04-14",
        "rows": [
            {"Name": "Charnick Botha", "Wed": "07:00-17:30", "Thu": "OFF", "Fri": "07:00-17:30", "Sat": "07:00-15:00", "Sun": "OFF", "Mon": "07:00-17:30", "Tue": "07:00-17:30"}
        ]
    }
    res = requests.post(f"{BASE_URL}/api/ai/generate_snippets", json=payload)
    print(f"Generate Snippets - Status: {res.status_code}, Response: {res.json()}")
    assert "snippets" in res.json()

if __name__ == "__main__":
    try:
        test_login()
        test_submit_request()
        test_approve_request()
        test_generate_snippets()
        print("\nAll new feature tests passed!")
    except Exception as e:
        print(f"\nTests failed: {e}")

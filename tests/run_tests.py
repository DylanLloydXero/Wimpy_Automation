"""
run_tests.py
End-to-end test suite for Wimpy De Ville Manager.
Run with the server ALREADY RUNNING: python run_tests.py
Or run standalone (it starts/stops the server itself).
"""

import sys
import os
import json
import time
import threading
import urllib.request
import urllib.error

BASE = "http://127.0.0.1:8000"
PASS = 0
FAIL = 0

# [*] Helpers [*]

def req(method, path, data=None, files=None):
    url = BASE + path
    if data is not None:
        body = json.dumps(data).encode()
        r = urllib.request.Request(url, data=body, method=method,
                                    headers={"Content-Type": "application/json"})
    else:
        r = urllib.request.Request(url, method=method)
    try:
        with urllib.request.urlopen(r, timeout=10) as resp:
            return json.loads(resp.read()), resp.status
    except urllib.error.HTTPError as e:
        return json.loads(e.read()), e.code
    except Exception as ex:
        return {"error": str(ex)}, 0


def ok(name, result):
    global PASS
    PASS += 1
    print(f"  [*]  {name}")


def fail(name, reason):
    global FAIL
    FAIL += 1
    print(f"  [*]  {name}  [*]  {reason}")


def test(name, condition, reason=""):
    if condition:
        ok(name, None)
    else:
        fail(name, reason)


# [*] Tests [*]

def test_server_alive():
    print("\n[*] Server Health [*]")
    body, status = req("GET", "/")
    test("GET / returns 200", status == 200, f"got {status}")


def test_staff_api():
    print("\n[*] Staff Manager [*]")

    # List staff
    body, status = req("GET", "/api/staff")
    test("GET /api/staff returns 200", status == 200, f"got {status}")
    test("Staff list has at least 1 employee", isinstance(body.get("staff"), list) and len(body["staff"]) > 0,
         f"got: {body}")

    # Add a test staff member
    payload = {"staff": {
        "name": "_Test Employee",
        "role": "TESTER",
        "rate": 99.99,
        "cell_number": "0800000000",
        "id_number": "0000000000000",
        "start_date": "2025-01-01",
        "leave_credit": 5.0
    }}
    body, status = req("POST", "/api/staff/update", payload)
    test("POST /api/staff/update (add new)", status == 200 and body.get("success"), f"{body}")

    # Verify added
    body, status = req("GET", "/api/staff")
    names = [s["name"] for s in body.get("staff", [])]
    test("New employee appears in list", "_Test Employee" in names, f"names: {names}")

    # Delete test staff member
    body, status = req("POST", "/api/staff/delete", {"name": "_Test Employee"})
    test("POST /api/staff/delete", status == 200 and body.get("success"), f"{body}")

    # Verify deleted
    body, status = req("GET", "/api/staff")
    names = [s["name"] for s in body.get("staff", [])]
    test("Employee removed from list", "_Test Employee" not in names, f"still in names: {names}")


def test_archives_api():
    print("\n[*] Archives [*]")
    body, status = req("GET", "/api/archives")
    test("GET /api/archives returns 200", status == 200, f"got {status}")
    test("Archives has 'payslips' key", "payslips" in body, f"keys: {list(body.keys())}")
    test("Archives has 'rosters' key", "rosters" in body, f"keys: {list(body.keys())}")
    test("Archives has at least 1 roster", len(body.get("rosters", [])) > 0,
         f"rosters: {body.get('rosters')}")
    test("Archives has at least 1 payslip run", len(body.get("payslips", [])) > 0,
         f"payslips: {body.get('payslips')}")


def test_roster_save():
    print("\n[*] Roster Save [*]")
    from datetime import date, timedelta
    today = date.today()
    days_since_tue = (today.weekday() - 1) % 7
    this_tue = today - timedelta(days=days_since_tue)
    week_date = this_tue.isoformat()

    rows = [{"NAME": name, week_date: "07:00-15:30 (8.5)", "TOTAL": "8.5"}
            for name in ["Test A", "Test B"]]

    body, status = req("POST", "/api/save_roster", {"week_date": week_date, "rows": rows})
    test("POST /api/save_roster returns success", status == 200 and body.get("success"), f"{body}")

    # Check it appeared in archives
    body, status = req("GET", "/api/archives")
    match = f"Roster_{week_date}.xlsx"
    test("Saved roster appears in archives", match in body.get("rosters", []), f"rosters: {body.get('rosters')}")


def test_set_active_roster():
    print("\n[*] Active Roster Linking [*]")
    from datetime import date, timedelta
    today = date.today()
    days_since_tue = (today.weekday() - 1) % 7
    this_tue = today - timedelta(days=days_since_tue)
    week_date = this_tue.isoformat()

    body, status = req("POST", "/api/set_active_roster", {"week_date": week_date})
    test("POST /api/set_active_roster (existing week)", status == 200 and body.get("success"), f"{body}")

    fake_date = "2000-01-01"
    body, status = req("POST", "/api/set_active_roster", {"week_date": fake_date})
    test("POST /api/set_active_roster (non-existent date) returns error", not body.get("success"), f"{body}")


def test_payroll_process():
    print("\n[*] Payroll Processing [*]")
    # Use the test roster and clock-in files
    if not os.path.exists("test_docs/Test roster.xlsx") or not os.path.exists("test_docs/Test clock in.xlsx"):
        print("  [*]  Skipping: test files not found in test_docs/")
        return

    import urllib.request
    import io

    boundary = "----TestBoundary"

    def encode_file(field, path):
        with open(path, "rb") as f:
            data = f.read()
        filename = os.path.basename(path)
        return (
            f"--{boundary}\r\n"
            f"Content-Disposition: form-data; name=\"{field}\"; filename=\"{filename}\"\r\n"
            f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        ).encode() + data + b"\r\n"

    body_bytes = encode_file("roster", "test_docs/Test roster.xlsx")
    body_bytes += encode_file("clockin", "test_docs/Test clock in.xlsx")
    body_bytes += f"--{boundary}--\r\n".encode()

    r = urllib.request.Request(
        BASE + "/api/process",
        data=body_bytes,
        method="POST",
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )
    try:
        with urllib.request.urlopen(r, timeout=20) as resp:
            result = json.loads(resp.read())
            emps = result.get("employees", [])
            test("POST /api/process returns employees list", len(emps) > 0,
                 f"got: {result}")
            if emps:
                test("Employees have required fields (name, rate, reg_hours)",
                     all("name" in e and "rate" in e and "reg_hours" in e for e in emps),
                     f"sample: {emps[0]}")
    except Exception as ex:
        fail("POST /api/process", str(ex))


# [*] Main [*]

def wait_for_server(max_tries=20):
    for _ in range(max_tries):
        try:
            urllib.request.urlopen(BASE + "/", timeout=0.5)
            return True
        except:
            time.sleep(0.3)
    return False


def start_server_for_tests():
    import uvicorn
    from api import app as fastapi_app
    t = threading.Thread(target=lambda: uvicorn.run(fastapi_app, host="127.0.0.1", port=8000, log_level="error"), daemon=True)
    t.start()
    return t


if __name__ == "__main__":
    server_already_running = wait_for_server(max_tries=3)

    if not server_already_running:
        print("[*] Starting server for tests...")
        start_server_for_tests()
        if not wait_for_server(max_tries=30):
            print("[*] Server failed to start. Aborting.")
            sys.exit(1)
    else:
        print("[*] Using already-running server.")

    test_server_alive()
    test_staff_api()
    test_archives_api()
    test_roster_save()
    test_set_active_roster()
    test_payroll_process()

    print(f"\n{'[*]'*54}")
    total = PASS + FAIL
    print(f"  Results: {PASS}/{total} tests passed  {'[*]' if FAIL == 0 else '[*]'}")
    if FAIL > 0:
        print(f"  {FAIL} test(s) FAILED [*] check output above for details.")
    print()
    sys.exit(0 if FAIL == 0 else 1)

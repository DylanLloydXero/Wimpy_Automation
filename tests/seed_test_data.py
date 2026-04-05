"""
seed_test_data.py
Generates realistic fake Archives and a full test runner.
Run this once: python seed_test_data.py
"""

import os
import shutil
import random
import zipfile
import pandas as pd
from datetime import datetime, timedelta, date

# ─── Staff list (from actual imported staff) ───────────────────────────────
STAFF = [
    "Emmerencia Hahn", "Vuyokazi Mathole", "Charnick Botha",
    "Calvin Botha", "Jade Williams", "Leana Lizaan Martin",
    "Viola Hardneck", "Sindiswa Nomanya", "Matumela Mokoatle/TUMI",
    "Noxolo Guvuza/NOXI", "Eduan DeKOCK", "Kaylen Lakey",
    "Tamlin Botha", "Cameron Engelbrecht", "Lorne Africa", "Tanya Baard"
]

# ─── Preset shifts (Wed-Tue week) ─────────────────────────────────────────
SHIFTS = [
    ("07:00", "13:30"),
    ("07:00", "15:30"),
    ("07:00", "17:30"),
    ("09:00", "15:30"),
    ("09:00", "17:30"),
    ("10:00", "15:30"),
    ("10:00", "17:30"),
    ("11:00", "17:30"),
    None,   # OFF
]

def calc_hours(start, end):
    if not start or not end:
        return 0.0
    sh, sm = map(int, start.split(":"))
    eh, em = map(int, end.split(":"))
    diff = (eh + em/60) - (sh + sm/60)
    if diff < 0:
        diff += 24
    return round(max(0, diff), 1)

def random_shift():
    s = random.choice(SHIFTS)
    if s is None:
        return "OFF", 0.0
    label = f"{s[0]}-{s[1]}"
    return label, calc_hours(s[0], s[1])


def make_roster(week_end: date) -> pd.DataFrame:
    """Generate a random roster DataFrame for a Wed-Tue week."""
    rows = []
    # Week columns: Wed to Tue (7 days, ending on the given Tuesday)
    dates = [(week_end - timedelta(days=6-i)) for i in range(7)]
    col_names = [d.strftime("%Y-%m-%d") for d in dates]

    for name in STAFF:
        row = {"NAME": name}
        total = 0.0
        for col in col_names:
            shift, hrs = random_shift()
            row[col] = shift
            total += hrs
        row["TOTAL"] = round(total, 1)
        rows.append(row)
    return pd.DataFrame(rows)


def make_eft_summary(week_end: date) -> pd.DataFrame:
    """Generate a fake EFT summary."""
    rows = []
    for name in STAFF:
        hours = round(random.uniform(35, 55), 1)
        rate = 30.33
        gross = round(hours * rate, 2)
        uif = round(gross * 0.01, 2)
        net = round(gross - uif - random.uniform(0, 200), 2)
        rows.append({"Employee": name, "Hours": hours, "Gross (R)": gross, "UIF (R)": uif, "Net Pay (R)": max(0, net)})
    return pd.DataFrame(rows)


def seed_rosters(n=4):
    """Create n past roster files in Archives/Rosters/"""
    os.makedirs("Archives/Rosters", exist_ok=True)
    today = date.today()
    # Find the most recent Tuesday
    days_since_tue = (today.weekday() - 1) % 7
    latest_tue = today - timedelta(days=days_since_tue)

    created = []
    for i in range(n):
        week_end = latest_tue - timedelta(weeks=i)
        filename = f"Archives/Rosters/Roster_{week_end.isoformat()}.xlsx"
        if not os.path.exists(filename):
            df = make_roster(week_end)
            df.to_excel(filename, index=False)
            created.append(filename)
            print(f"  [OK] Created roster: {filename}")
        else:
            print(f"  [--] Already exists: {filename}")
    return created


def seed_payslips(n=3):
    """Create n past payslip archive folders in Archives/Payslips/"""
    os.makedirs("Archives/Payslips", exist_ok=True)
    today = date.today()
    days_since_tue = (today.weekday() - 1) % 7
    latest_tue = today - timedelta(days=days_since_tue)

    created = []
    for i in range(n):
        week_end = latest_tue - timedelta(weeks=i + 1)  # Past weeks only
        folder_name = f"Payroll_{week_end.isoformat()}_120000"
        folder_path = f"Archives/Payslips/{folder_name}"
        if os.path.exists(folder_path):
            print(f"  [--] Already exists: {folder_path}")
            continue
        os.makedirs(folder_path)

        # Create dummy EFT xlsx
        eft_df = make_eft_summary(week_end)
        eft_path = f"{folder_path}/EFT_Summary_{week_end.isoformat()}.xlsx"
        eft_df.to_excel(eft_path, index=False)

        # Create a dummy ZIP of "payslips"
        zip_path = f"{folder_path}/Payslips_{week_end.isoformat()}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zf:
            for name in STAFF[:5]:  # Just a few as placeholders
                safe = name.replace("/", "_").replace(" ", "_")
                zf.writestr(f"{safe}_payslip.pdf", f"Payslip for {name} - week ending {week_end}")

        created.append(folder_path)
        print(f"  [OK] Created payslip archive: {folder_path}")
    return created


def seed_test_docs():
    """Copy a meaningful roster into test_docs/ for uploading."""
    os.makedirs("test_docs", exist_ok=True)
    today = date.today()
    days_since_tue = (today.weekday() - 1) % 7
    week_end = today - timedelta(days=days_since_tue)

    path = "test_docs/Test roster (current week).xlsx"
    df = make_roster(week_end)
    df.to_excel(path, index=False)
    print(f"  [OK] Created current week test roster: {path}")


if __name__ == "__main__":
    print("\n[*] Seeding test data into Archives and test_docs...\n")
    seed_rosters(n=4)
    print()
    seed_payslips(n=3)
    print()
    seed_test_docs()
    print("\n[DONE] Check Archives/ and test_docs/ folders.\n")

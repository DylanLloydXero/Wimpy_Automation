import pandas as pd
from datetime import date, timedelta, datetime
import random
import os

# Read active staff list
master_path = "Staff_Details_Template.xlsx"
staff_df = pd.read_excel(master_path)
staff_names = staff_df["Name"].dropna().tolist()

# Define week (Wed to Tue)
today = date.today()
days_since_tue = (today.weekday() - 1) % 7
week_end = today - timedelta(days=days_since_tue)
dates = [(week_end - timedelta(days=6-i)) for i in range(7)]

# 1. Generate MOCK Roster
roster_rows = []
shifts = [
    ("07:00", "15:30", 8.5),
    ("09:00", "17:30", 8.5),
    ("10:00", "17:30", 7.5),
    ("OFF", "OFF", 0.0)
]

# We will store schedules to generate matching clock-ins
schedules = {}

for name in staff_names:
    row = {"NAME": name}
    total_hours = 0.0
    sched = []
    
    for d in dates:
        shift_choice = random.choice(shifts)
        if shift_choice[0] == "OFF":
            row[d.strftime("%Y-%m-%d")] = "OFF"
            sched.append(None)
        else:
            row[d.strftime("%Y-%m-%d")] = f"{shift_choice[0]}-{shift_choice[1]} ({shift_choice[2]})"
            total_hours += shift_choice[2]
            sched.append((d, shift_choice[0], shift_choice[1]))
            
    row["TOTAL"] = total_hours
    roster_rows.append(row)
    schedules[name] = sched

roster_df = pd.DataFrame(roster_rows)
roster_path = "test_docs/MOCK_Roster.xlsx"
roster_df.to_excel(roster_path, index=False)

# 2. Generate MOCK Clock-In Ankerdata Report
clock_in_rows = []

# Ankerdata format requires some formatting
# ['Company', 'Store', 'Server Number', 'Server Name', 'Time Logged In', 'Time Logged Out', 'Time Worked', 'Reason Out']

for name, sched in schedules.items():
    server_num = random.randint(100, 999)
    for s in sched:
        if s is not None:
            day_date, start_str, end_str = s
            # Create datetime objects
            time_in = datetime.strptime(f"{day_date} {start_str}:00", "%Y-%m-%d %H:%M:%S")
            time_out = datetime.strptime(f"{day_date} {end_str}:00", "%Y-%m-%d %H:%M:%S")
            
            # Simulate slight variations in clock-in (few minutes early/late)
            time_in += timedelta(minutes=random.randint(-15, 5))
            time_out += timedelta(minutes=random.randint(-5, 15))
            
            # Format times for Ankerdata standard
            clock_in_rows.append({
                "Company": "Wimpy De Ville",
                "Store": "100269 - Wimpy",
                "Server Number": server_num,
                "Server Name": name.upper(),
                "Time Logged In": time_in.strftime("%d/%m/%Y %H:%M:%S"),
                "Time Logged Out": time_out.strftime("%d/%m/%Y %H:%M:%S"),
                "Time Worked": "", # Script ignores this and calculates it
                "Reason Out": ""
            })

# Add empty rows at top to mimic Ankerdata report exactly
# Ankerdata report has two blank rows
clockin_df = pd.DataFrame(clock_in_rows)
blank_rows = pd.DataFrame([[None]*8, [None]*8], columns=clockin_df.columns)
clockin_df = pd.concat([blank_rows, clockin_df], ignore_index=True)

clockin_path = "test_docs/MOCK_ClockIn.xlsx"
clockin_df.to_excel(clockin_path, index=False)

print(f"Generated {roster_path}")
print(f"Generated {clockin_path}")
print("Mock files are ready for upload!")

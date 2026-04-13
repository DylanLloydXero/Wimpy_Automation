import pandas as pd

excel_path = "test_docs/WIMPY STAFF LIST 2026.xlsx"
master_path = "Staff_Details_Template.xlsx"

# Read skipping the first two title rows, ensuring cell and id are treated as strings
df_in = pd.read_excel(excel_path, skiprows=2, dtype={'ID NR.': str, 'CELL': str})

expected_cols = ["Name", "ID Number", "Rate", "Start Date", "Leave Credit", "Cell Number", "Role"]
rows = []

current_role = "WAITER"

for _, row in df_in.iterrows():
    name_raw = str(row.get('Name', '')).strip()
    
    if name_raw == "nan" or not name_raw:
        continue
        
    # Check if it's a role header (All caps, no spaces)
    if len(name_raw) > 2 and name_raw.isupper() and " " not in name_raw:
        current_role = name_raw.replace("S", "") if name_raw.endswith("S") else name_raw # WAITERS -> WAITER
        continue
        
    id_num = str(row.get('ID NR.', '')).replace("nan", "")
    # Pad to 13 digits if it lost the leading zero
    if len(id_num) == 12:
        id_num = "0" + id_num
        
    cell = str(row.get('CELL', '')).replace("nan", "")
    if "geen" in cell.lower():
        cell = ""
    else:
        cell = cell.replace(" ", "")

    leave_val = 0.0
    lv = str(row.get('LEAVE AVAILABLE', '')).replace("nan", "")
    if lv:
        try: leave_val = float(lv)
        except: pass
        
    start_date = ""
    sd = row.get('STARTING DATE', None)
    if pd.notna(sd):
        try: start_date = sd.strftime("%Y-%m-%d")
        except: start_date = str(sd).split(" ")[0]

    # Default wage
    rate = 30.33
    if "SUPERVISOR" in current_role or "OWNER" in current_role:
        rate = 40.00 # Placeholder for managers
        
    rows.append({
        "Name": name_raw,
        "ID Number": id_num,
        "Rate": rate,
        "Start Date": start_date,
        "Leave Credit": leave_val,
        "Cell Number": cell,
        "Role": current_role
    })

df_out = pd.DataFrame(rows, columns=expected_cols)

# Overwrite or create the database
df_out.to_excel(master_path, index=False)
print(f"Successfully imported {len(df_out)} staff members (with start dates and leave credits) into {master_path}!")

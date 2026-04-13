import pdfplumber
import pandas as pd
import os

pdf_path = "test_docs/WIMPY STAFF LIST 2026.pdf"
master_path = "Staff_Details_Template.xlsx"

expected_cols = ["Name", "ID Number", "Rate", "Start Date", "Leave Credit", "Cell Number", "Role"]
rows = []

current_role = "WAITER"

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        tables = page.extract_tables()
        if tables:
            for t in tables:
                for row in t:
                    if not row or not row[0]:
                        continue
                        
                    name_raw = str(row[0]).strip()
                    if name_raw in ["Name", "Staff list April 2026"]:
                        continue
                        
                    # Check if it's a role header
                    if len(name_raw) > 2 and name_raw.isupper() and " " not in name_raw:
                        current_role = name_raw.replace("S", "") if name_raw.endswith("S") else name_raw # WAITERS -> WAITER
                        continue
                        
                    id_num = str(row[1]) if len(row) > 1 and row[1] else ""
                    cell = str(row[2]) if len(row) > 2 and row[2] else ""
                    
                    if "geen" in cell.lower():
                        cell = ""
                    else:
                        cell = cell.replace(" ", "")
                    if id_num.startswith("*"):
                        id_num = id_num[1:]
                        
                    # Default wage and leave
                    rate = 30.33
                    if "SUPERVISOR" in current_role or "OWNER" in current_role:
                        rate = 40.00 # Placeholder for managers
                        
                    rows.append({
                        "Name": name_raw,
                        "ID Number": id_num,
                        "Rate": rate,
                        "Start Date": "",
                        "Leave Credit": 0.0,
                        "Cell Number": cell,
                        "Role": current_role
                    })

df = pd.DataFrame(rows, columns=expected_cols)

# Overwrite or create the database
df.to_excel(master_path, index=False)
print(f"Successfully imported {len(df)} staff members into {master_path}!")

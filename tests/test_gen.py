import asyncio
import json
import os
import shutil
import pandas as pd
from typing import List

# We need to test the logic from api.py generate_pdfs.
# We will just copy the function logic to debug here.

def debug_generate(employees):
    links = []
    leave_deductions = {}
    from payslip_generator import generate_payslip

    for emp in employees:
        try:
            # Ensure values are numeric to avoid "stuck" errors
            e_rate = float(emp.get("rate", 0.0) or 0.0)
            e_reg = float(emp.get("reg_hours", 0.0) or 0.0)
            e_sun = float(emp.get("sun_hours", 0.0) or 0.0)
            e_hol = float(emp.get("hol_hours", 0.0) or 0.0)
            e_leave_days = float(emp.get("leave_days", 0.0) or 0.0)
            e_sick_days = float(emp.get("sick_days", 0.0) or 0.0)
            e_bonus = float(emp.get("bonus", 0.0) or 0.0)
            e_short = float(emp.get("till_short", 0.0) or 0.0)
            e_tips = float(emp.get("tips", 0.0) or 0.0)
            e_clothing = float(emp.get("clothing", 0.0) or 0.0)

            clean_pay = {
                "name": str(emp.get("name", "Unknown")),
                "rate": e_rate,
                "reg_hours": e_reg,
                "sun_hours": e_sun,
                "hol_hours": e_hol,
                "leave_days": e_leave_days,
                "sick_days": e_sick_days,
                "bonus": e_bonus,
                "till_short": e_short,
                "tips": e_tips,
                "clothing": e_clothing,
                "id_number": str(emp.get("id_number", "")),
                "start_date": str(emp.get("start_date", "")),
                "leave_credit": str(emp.get("leave_credit", "")),
                "cell_number": str(emp.get("cell_number", "")),
                "role": str(emp.get("role", "WAITER"))
            }

            filepath = generate_payslip(clean_pay)
            urlpath = filepath.replace('\\', '/')
            
            leave_pay = clean_pay["leave_days"] * 7 * clean_pay["rate"]
            sick_pay = clean_pay["sick_days"] * 7 * clean_pay["rate"]
            
            gross = (clean_pay["reg_hours"] * clean_pay["rate"]) + \
                    (clean_pay["sun_hours"] * clean_pay["rate"] * 1.5) + \
                    (clean_pay["hol_hours"] * clean_pay["rate"] * 2) + \
                    leave_pay + sick_pay + clean_pay["bonus"] - clean_pay["till_short"]
            
            uif = gross * 0.01
            total_deduct = clean_pay["tips"] + uif + clean_pay["clothing"]
            net = gross - total_deduct
            
            message = f"Hi {clean_pay['name']}, here is your payslip for the week. Total Nett Salary: R {net:.2f}."
            
            phone = clean_pay['cell_number']
            if phone == 'None': phone = ''
            phone = ''.join(filter(str.isdigit, str(phone)))

            links.append({
                "name": clean_pay["name"],
                "url": f"/{urlpath}",
                "phone": phone,
                "wa_message": message,
                "net": net
            })
        except Exception as emp_e:
            print(f"Error processing employee {emp.get('name')}: {emp_e}")
            import traceback
            traceback.print_exc()
            continue
    print(links)

if __name__ == "__main__":
    payload = [
        {
            "name": "SAYLIN",
            "reg_hours": 10.0,
            "sun_hours": 0.0,
            "hol_hours": 0.0,
            "leave_days": 0.0,
            "sick_days": 0.0,
            "bonus": "",
            "till_short": 0.0,
            "tips": "",
            "clothing": 0.0,
            "rate": 30.33,
            "cell_number": None,
            "role": "WAITER"
        }
    ]
    debug_generate(payload)

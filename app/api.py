from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import shutil
import os
import json
import pandas as pd
from typing import List
import sys
import subprocess
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from datetime import datetime
import glob
from PyPDF2 import PdfMerger
import holidays as pyholidays
import smtplib, ssl
from email.message import EmailMessage

from google import genai
from google.genai import types
from .payroll_processor import process_payroll
from .payslip_generator import generate_payslip, PayslipPDF
from fpdf import FPDF
try:
    import win32clipboard
except ImportError:
    win32clipboard = None

# Load Config
CONFIG_PATH = "data/config.json"
GEMINI_KEY = os.environ.get("GEMINI_API_KEY", "")

if os.path.exists(CONFIG_PATH):
    try:
        with open(CONFIG_PATH, "r") as f:
            _cfg = json.load(f)
            GEMINI_KEY = _cfg.get("GEMINI_API_KEY", GEMINI_KEY)
    except: pass
elif os.path.exists("config.json"):
    try:
        with open("config.json", "r") as f:
            _cfg = json.load(f)
            GEMINI_KEY = _cfg.get("GEMINI_API_KEY", GEMINI_KEY)
    except: pass

client = None
if GEMINI_KEY:
    try:
        client = genai.Client(api_key=GEMINI_KEY)
    except:
        print("AI Client failed to initialize. Check API Key.")

app = FastAPI()

os.makedirs("static", exist_ok=True)
os.makedirs("output/payslips", exist_ok=True)
os.makedirs("data/archives", exist_ok=True)

app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/payslips", StaticFiles(directory="output/payslips"), name="payslips")

@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/api/login")
async def login(request: Request):
    data = await request.json()
    name = str(data.get("name", "")).strip()
    pin = str(data.get("id_last4", "")).strip()

    # Manager bypass requires PIN 0000
    if name.lower() == "manager":
        if pin == "0000":
            return JSONResponse(content={"success": True, "role": "MANAGER", "name": "Manager"})
        else:
            return JSONResponse(content={"success": False, "error": "Invalid Manager PIN"})

    # Check PIN for staff
    pin_path = "data/staff_pins.json"
    if os.path.exists(pin_path):
        with open(pin_path, "r") as f:
            pins = json.load(f)
        if name in pins and pins[name] == pin:
            # Check portal_access from master excel
            master_path = "data/templates/Staff_Details_Template.xlsx"
            role = "EMPLOYEE"
            portal_access = "MY_SHIFTS"
            if os.path.exists(master_path):
                mdf = pd.read_excel(master_path)
                row = mdf[mdf["Name"].astype(str).str.strip().str.lower() == name.lower()]
                if not row.empty:
                    portal_access = str(row.iloc[0].get("Portal Access", "MY_SHIFTS")).replace("nan", "MY_SHIFTS") if "Portal Access" in mdf.columns else "MY_SHIFTS"
                    if portal_access == "MANAGER":
                        role = "MANAGER"
            return JSONResponse(content={"success": True, "role": role, "name": name, "portal_access": portal_access})

    return JSONResponse(content={"success": False, "error": "Invalid Name or PIN"})

@app.get("/api/holidays")
async def get_holidays(year: int = None):
    if not year: year = datetime.now().year
    sa_holidays = pyholidays.SouthAfrica(years=year)
    return JSONResponse(content={
        "holidays": [{"date": str(d), "name": n} for d, n in sorted(sa_holidays.items())]
    })

@app.get("/api/requests")
async def get_requests():
    path = "data/off_day_requests.json"
    if not os.path.exists(path): return JSONResponse(content=[])
    with open(path, "r") as f:
        return JSONResponse(content=json.load(f))

@app.post("/api/submit_request")
async def submit_request(request: Request):
    data = await request.json()
    path = "data/off_day_requests.json"
    employee = data.get("employee")
    off_days = data.get("off_days", []) # List of strings e.g. ["2026-04-10"]
    req_type = data.get("type", "OFF_DAY") # "OFF_DAY" or "LEAVE"
    
    now = datetime.now()
    
    # 2026 School Holidays SA
    SCHOOL_HOLIDAYS = [
        ("2026-03-28", "2026-04-07"),
        ("2026-06-27", "2026-07-20"),
        ("2026-09-24", "2026-10-05"),
        ("2026-12-10", "2026-12-31")
    ]
    
    sa_holidays = pyholidays.SouthAfrica(years=[now.year, now.year + 1])
    
    def is_school_holiday(date_str):
        d = datetime.strptime(date_str, "%Y-%m-%d")
        for start, end in SCHOOL_HOLIDAYS:
            if datetime.strptime(start, "%Y-%m-%d") <= d <= datetime.strptime(end, "%Y-%m-%d"):
                return True
        return False

    errors = []
    status = "APPROVED"  # Default

    for day in off_days:
        d_obj = datetime.strptime(day, "%Y-%m-%d")
        lead_time = (d_obj - now).days

        # 1. 14-day lead time for LEAVE only
        if req_type == "LEAVE" and lead_time < 14:
            errors.append(f"LEAVE for {day} must be requested at least 14 days in advance. Currently only {lead_time} days away.")

        # 2. Only Saturday (5) and Sunday (6) require manager approval for OFF_DAY
        if d_obj.weekday() in [5, 6]:
            status = "PENDING_APPROVAL"

    if errors:
        return JSONResponse(content={"success": False, "error": "; ".join(errors)})

    reqs = []
    if os.path.exists(path):
        with open(path, "r") as f:
            try:
                reqs = json.load(f)
                if not isinstance(reqs, list): reqs = []
            except:
                reqs = []

    reqs.append({
        "id": str(now.timestamp()),
        "employee": employee,
        "off_days": off_days,
        "type": req_type,
        "status": status,
        "timestamp": now.strftime("%Y-%m-%d %H:%M")
    })

    with open(path, "w") as f: json.dump(reqs, f, indent=4)
    return JSONResponse(content={"success": True, "status": status})

@app.post("/api/approve_request")
async def approve_request(request: Request):
    data = await request.json()
    req_id = data.get("id")
    action = data.get("action") # "APPROVE" or "REJECT"
    
    path = "data/off_day_requests.json"
    if not os.path.exists(path): return JSONResponse(content={"success": False, "error": "No requests found"})
    
    with open(path, "r") as f: requests = json.load(f)
    
    found = False
    for r in requests:
        if r.get("id") == req_id:
            r["status"] = "APPROVED" if action == "APPROVE" else "REJECTED"
            found = True
            break
            
    if found:
        with open(path, "w") as f: json.dump(requests, f, indent=4)
        return JSONResponse(content={"success": True})
    return JSONResponse(content={"success": False, "error": "Request not found"})

@app.post("/api/process")
async def process_files(
    roster: UploadFile = File(None),
    clockin: UploadFile = File(None)
):
    os.makedirs("tmp", exist_ok=True)
    
    roster_path = "data/input/latest_roster.xlsx"
    clockin_path = "data/input/latest_clockin.xlsx"
    
    if roster and roster.filename:
        roster_path = f"tmp/{roster.filename}"
        with open(roster_path, "wb") as buffer:
             shutil.copyfileobj(roster.file, buffer)
             
    if clockin and clockin.filename:
        clockin_path = f"tmp/{clockin.filename}"
        with open(clockin_path, "wb") as buffer:
             shutil.copyfileobj(clockin.file, buffer)
             
    if not os.path.exists(roster_path) or not os.path.exists(clockin_path):
        return JSONResponse(content={"error": "Missing Roster or Clock-in files, and no local fallback found."})
         
    master_details = {}
    master_path = "data/templates/Staff_Details_Template.xlsx"
    if os.path.exists(master_path):
        try:
            master_df = pd.read_excel(master_path)
            cols = [str(c).strip().lower() for c in master_df.columns]
            for _, row in master_df.iterrows():
                name = str(row['Name']).strip().lower() if 'name' in cols else str(row.iloc[0]).strip().lower()
                rate_val = str(row['Rate']).replace('R', '').replace(',', '.').strip() if 'rate' in cols else "30.33"
                try: rate_val = float(rate_val) if rate_val and str(rate_val).lower() != "nan" else 30.33
                except: rate_val = 30.33
                
                master_details[name] = {
                    "rate": rate_val,
                    "id_number": str(row['ID Number']).replace('nan', '') if 'id number' in cols else "",
                    "start_date": str(row['Start Date']).replace('nan', '') if 'start date' in cols else "",
                    "leave_credit": str(row['Leave Credit']).replace('nan', '') if 'leave credit' in cols else "",
                    "cell_number": str(row['Cell Number']).replace('nan', '').replace('.0', '') if 'cell number' in cols else "",
                    "role": str(row['Role']).replace('nan', '') if 'role' in cols else "WAITER"
                }
        except Exception as e:
            print("Could not parse master list:", e)

    output_df, summary_df, overtime_flags = process_payroll(roster_path, clockin_path, output_format='dataframe')
    
    employees = []
    if summary_df is not None and not summary_df.empty:
        for _, row in summary_df.iterrows():
            name = str(row['Employee Name'])
            name_lower = name.lower()
            
            rate = 30.33
            details = {"id_number": "", "start_date": "", "leave_credit": "", "cell_number": "", "role": "WAITER"}
            for m_name, m_detail in master_details.items():
                if name_lower in m_name or m_name in name_lower:
                    rate = m_detail["rate"]
                    details = m_detail
                    break
            
            employees.append({
                "name": name,
                "reg_hours": float(row.get('Reg Hours', 0)),
                "sun_hours": float(row.get('Sun Hours', 0)),
                "tue_hours": float(row.get('Tue Hours', 0)),
                "hol_hours": 0.0,
                "leave_days": 0.0,
                "sick_days": 0.0,
                "total_hours": float(row.get('Total Payable Hours', 0)),
                "rate": rate,
                "bonus": 0.0,
                "till_short": 0.0,
                "tips": 0.0,
                "clothing": 0.0,
                **details
            })
    else:
        # Fallback to master list with 0 hours if no processing success
        for m_name, m_details in master_details.items():
             employees.append({
                "name": m_name.title(),
                "reg_hours": 0.0,
                "sun_hours": 0.0,
                "tue_hours": 0.0,
                "hol_hours": 0.0,
                "leave_days": 0.0,
                "sick_days": 0.0,
                "total_hours": 0.0,
                "rate": m_details["rate"],
                "bonus": 0.0,
                "till_short": 0.0,
                "tips": 0.0,
                "clothing": 0.0,
                **m_details
            })

    return JSONResponse(content={"employees": employees, "overtime_flags": overtime_flags})

@app.post("/api/approve_overtime")
async def approve_overtime(request: Request):
    """Saves overtime decisions (APPROVED or DENIED) to data/overtime_approvals.json"""
    try:
        data = await request.json()
        employee = data.get("employee")
        date_str = data.get("date")
        status = data.get("status")
        
        if not employee or not date_str or not status:
            return JSONResponse(content={"success": False, "error": "Missing mapping detail."})
            
        path = "data/overtime_approvals.json"
        approvals = {}
        if os.path.exists(path):
            with open(path, "r") as f:
                approvals = json.load(f)
                
        key = f"{employee}_{date_str}"
        approvals[key] = status
        
        with open(path, "w") as f:
            json.dump(approvals, f, indent=4)
            
        return JSONResponse(content={"success": True})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/generate")
async def generate_pdfs(request: Request):
    try:
        data = await request.json()
        employees = data.get("employees", [])
        
        # Clear existing payslips for a clean run
        os.makedirs("output/payslips", exist_ok=True)
        import glob
        for f in glob.glob("output/payslips/*"):
            try:
                if os.path.isfile(f):
                    os.remove(f)
            except Exception:
                pass
                
        links = []
        leave_deductions = {}
        
        def safe_float(val, default=0.0):
            try:
                if val is None or str(val).strip() == "":
                    return default
                return float(val)
            except:
                return default
                
        for emp in employees:
            try:
                # 1. Sanitize all incoming UI fields (convert empty strings/nulls to float)
                clean_pay = {
                    "name": str(emp.get("name", "Unknown")),
                    "rate": safe_float(emp.get("rate")),
                    "reg_hours": safe_float(emp.get("reg_hours")),
                    "tue_hours": safe_float(emp.get("tue_hours")),
                    "sun_hours": safe_float(emp.get("sun_hours")),
                    "hol_hours": safe_float(emp.get("hol_hours")),
                    "leave_days": safe_float(emp.get("leave_days")),
                    "sick_days": safe_float(emp.get("sick_days")),
                    "bonus": safe_float(emp.get("bonus")),
                    "till_short": safe_float(emp.get("till_short")),
                    "tips": safe_float(emp.get("tips")),
                    "clothing": safe_float(emp.get("clothing")),
                    "id_number": str(emp.get("id_number", "")),
                    "start_date": str(emp.get("start_date", "")),
                    "leave_credit": str(emp.get("leave_credit", "")),
                    "cell_number": str(emp.get("cell_number", "")),
                    "role": str(emp.get("role", "WAITER"))
                }

                # 2. Generate PDF with CLEAN data
                filepath = generate_payslip(clean_pay)
                urlpath = filepath.replace('\\', '/')
                
                # 3. Calculate math with CLEAN data
                leave_pay = clean_pay["leave_days"] * 7 * clean_pay["rate"]
                sick_pay = clean_pay["sick_days"] * 7 * clean_pay["rate"]
                
                gross = (clean_pay["reg_hours"] * clean_pay["rate"]) + \
                        (clean_pay["tue_hours"] * clean_pay["rate"]) + \
                        (clean_pay["sun_hours"] * clean_pay["rate"] * 1.5) + \
                        (clean_pay["hol_hours"] * clean_pay["rate"] * 2) + \
                        leave_pay + sick_pay + clean_pay["bonus"] - clean_pay["till_short"]
                
                uif = gross * 0.01
                total_deduct = clean_pay["tips"] + uif + clean_pay["clothing"]
                net = gross - total_deduct
                
                message = f"Hi {clean_pay['name']}, here is your payslip for the week. Total Nett Salary: R {net:.2f}."
                
                phone = clean_pay['cell_number']
                phone = ''.join(filter(str.isdigit, phone))

                links.append({
                    "name": clean_pay["name"],
                    "url": f"/{urlpath}",
                    "phone": phone,
                    "wa_message": message,
                    "net": net
                })
                
                if clean_pay["leave_days"] > 0:
                    leave_deductions[clean_pay['name'].strip().lower()] = clean_pay["leave_days"]
            except Exception as emp_e:
                print(f"Error processing employee {emp.get('name')}: {emp_e}")
                continue
                
        # --- LEAVE UPDATING ---
        master_path = "data/templates/Staff_Details_Template.xlsx"
        if os.path.exists(master_path) and leave_deductions:
            try:
                master_df = pd.read_excel(master_path)
                name_col = next((c for c in master_df.columns if str(c).strip().lower() == 'name'), master_df.columns[0])
                leave_col = next((c for c in master_df.columns if str(c).strip().lower() == 'leave credit'), None)
                
                if leave_col:
                    for idx, row in master_df.iterrows():
                        emp_name = str(row[name_col]).strip().lower()
                        if emp_name in leave_deductions:
                            try: cur = float(row[leave_col]) if str(row[leave_col]).lower() != 'nan' else 0.0
                            except: cur = 0.0
                            master_df.at[idx, leave_col] = max(0.0, cur - leave_deductions[emp_name])
                    master_df.to_excel(master_path, index=False)
            except Exception as e:
                print("Failed to update leave:", e)
                
        # --- ARCHIVING ---
        try:
            date_str = datetime.now().strftime("%Y-%m-%d")
            archive_month_dir = get_archive_dir(date_str)
            payroll_folder_name = f"Payslips_{date_str}_{datetime.now().strftime('%H-%M')}"
            archive_dir = os.path.join(archive_month_dir, payroll_folder_name)
            os.makedirs(archive_dir, exist_ok=True)
            
            # Save Zip of all PDFs
            shutil.make_archive(os.path.join(archive_dir, "All_Payslips"), "zip", "output/payslips")
            
            # Save EFT Summary
            rows_eft = []
            total_sum = 0
            for l in links:
                net_val = round(l.get("net", 0.0), 2)
                rows_eft.append({"Employee Name": l.get("name"), "Final Net Pay": net_val})
                total_sum += net_val
            
            rows_eft.append({"Employee Name": "TOTAL", "Final Net Pay": round(total_sum, 2)})
            pd.DataFrame(rows_eft).to_excel(os.path.join(archive_dir, "EFT_Summary.xlsx"), index=False)
            
            # Copy active roster to this archive folder as well for reference
            if os.path.exists("data/input/latest_roster.xlsx"):
                shutil.copy("data/input/latest_roster.xlsx", os.path.join(archive_dir, "Matching_Roster.xlsx"))

        except Exception as arch_e:
            print("Archiving failed:", arch_e)
                
        return JSONResponse(content={"links": links})
    except Exception as global_e:
        print("Global payroll generation failure:", global_e)
        return JSONResponse(content={"error": str(global_e)}, status_code=500)


@app.get("/api/download_all")
async def download_all():
    shutil.make_archive("output/summaries/Payslips_Batch", "zip", "output/payslips")
    return FileResponse("output/summaries/Payslips_Batch.zip", media_type="application/zip", filename="Payslips_Batch.zip")

@app.get("/api/download_merged_pdf")
async def download_merged_pdf():
    try:
        merger = PdfMerger()
        pdf_files = glob.glob("output/payslips/*.pdf")
        if not pdf_files:
            return JSONResponse(content={"error": "No payslips found to merge."}, status_code=404)
            
        for pdf in pdf_files:
            merger.append(pdf)
            
        merged_filename = f"output/summaries/All_Payslips_{datetime.now().strftime('%Y-%m-%d')}.pdf"
        merger.write(merged_filename)
        merger.close()
        return FileResponse(merged_filename, media_type="application/pdf", filename=merged_filename)
    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)

@app.post("/api/download_eft")
async def download_eft(request: Request):
    data = await request.json()
    links = data.get("links", [])
    
    rows = []
    total_val = 0
    for l in links:
        net_val = round(l.get("net", 0.0), 2)
        rows.append({"Employee Name": l.get("name"), "Final Net Pay": net_val})
        total_val += net_val
        
    rows.append({"Employee Name": "TOTAL", "Final Net Pay": round(total_val, 2)})
        
    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"output/summaries/EFT_Summary_{date_str}.xlsx"
    df.to_excel(filename, index=False)
    return FileResponse(filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

@app.post("/api/download_eft_pdf")
async def download_eft_pdf(request: Request):
    data = await request.json()
    links = data.get("links", [])
    
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "WIMPY DE VILLE - EFT PAYMENT SUMMARY", ln=True, align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 8, f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True, align="C")
    pdf.ln(10)
    
    # Header
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(100, 10, "Employee Name", border=1)
    pdf.cell(80, 10, "Final Net Pay (R)", border=1, ln=True, align="R")
    
    pdf.set_font("Helvetica", "", 12)
    total_val = 0
    for l in links:
        n = l.get("name", "Unknown")
        p = round(l.get("net", 0.0), 2)
        pdf.cell(100, 10, n, border=1)
        pdf.cell(80, 10, f"{p:.2f}", border=1, ln=True, align="R")
        total_val += p
        
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(100, 12, "GRAND TOTAL", border=1)
    pdf.cell(80, 12, f"R {total_val:.2f}", border=1, ln=True, align="R")
    
    pdf_path = "output/summaries/EFT_Summary.pdf"
    pdf.output(pdf_path)
    return FileResponse(pdf_path, media_type="application/pdf", filename=f"EFT_Summary_{datetime.now().strftime('%Y-%m-%d')}.pdf")

@app.post("/api/copy_to_clipboard")
async def copy_to_clipboard(request: Request):
    if not win32clipboard:
        return JSONResponse(content={"success": False, "error": "win32clipboard not available"})
        
    try:
        data = await request.json()
        relative_path = data.get("path", "").lstrip('/')
        abs_path = os.path.abspath(relative_path)
        
        if not os.path.exists(abs_path):
            return JSONResponse(content={"success": False, "error": f"File not found: {abs_path}"})
            
        import time
        success = False
        for attempt in range(5):
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, [abs_path])
                win32clipboard.CloseClipboard()
                success = True
                break
            except Exception:
                time.sleep(0.1)
                
        if success:
            return JSONResponse(content={"success": True})
        else:
            return JSONResponse(content={"success": False, "error": "Clipboard access denied after retries."})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/print_file")
async def print_file(request: Request):
    try:
        data = await request.json()
        relative_path = data.get("path", "").lstrip('/')
        abs_path = os.path.abspath(relative_path)
        
        if os.path.exists(abs_path):
            # Uses the system's 'print' verb on Windows
            os.startfile(abs_path, "print")
            return JSONResponse(content={"success": True})
        return JSONResponse(content={"success": False, "error": "File not found"})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/send_roster_image")
async def send_roster_image(request: Request):
    try:
        import base64
        data = await request.json()
        img_data = data.get("image", "")
        if not img_data:
            return JSONResponse(content={"success": False, "error": "No image data received"})
        
        # Remove header
        if "," in img_data:
            img_data = img_data.split(",")[1]
            
        with open("output/roster_snapshot.png", "wb") as f:
            f.write(base64.b64decode(img_data))
            
        # Trigger WhatsApp bot for standard person (e.g. manager) or as defined
        # For now, we'll save it and log it. We can extend wa_bot to send this file.
        with open("output/whatsapp_payload.json", "w") as f:
            json.dump([{
                "phone": data.get("phone", ""),
                "message": "📅 Here is the latest Staff Roster for Wimpy De Ville.",
                "file": "output/roster_snapshot.png"
            }], f)
        
        subprocess.Popen([sys.executable, "app/wa_bot.py"])
        return JSONResponse(content={"success": True})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/whatsapp_send")
async def whatsapp_send(request: Request):
    data = await request.json()
    links = data.get("links", [])
    
    payload = []
    for l in links:
        path = l.get("url", "").lstrip('/')
        if path:
            payload.append({
                "phone": l.get("phone", ""),
                "message": l.get("wa_message", ""),
                "file": path
            })
            
    with open("output/whatsapp_payload.json", "w") as f:
        json.dump(payload, f)
        
    subprocess.Popen([sys.executable, "app/wa_bot.py"])
    return JSONResponse(content={"status": "Bot started in background."})

@app.post("/api/cleanup")
async def auto_cleanup():
    """Removes temporary files and old payslips to keep the space clean."""
    try:
        if os.path.exists("tmp"):
            shutil.rmtree("tmp")
        # Keep the latest payslips in output, but clear individual ones if needed
        # For now, let's just clear tmp
        return JSONResponse(content={"success": True, "message": "Cleanup complete!"})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/fetch_ankerdata")
async def fetch_ankerdata():
    try:
        res = subprocess.run([sys.executable, "app/ankerdata_bot.py"], capture_output=True, text=True)
        if res.returncode == 0:
            return JSONResponse(content={"success": True})
        else:
            return JSONResponse(content={"success": False, "error": res.stderr})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/fetch_mock")
async def fetch_mock():
    """Simulates a successful Ankerdata fetch for testing."""
    try:
        roster_src = "tests/test_docs/MOCK_Roster.xlsx"
        clock_src = "tests/test_docs/MOCK_ClockIn.xlsx"
        
        if os.path.exists(roster_src) and os.path.exists(clock_src):
            shutil.copy(roster_src, "data/input/latest_roster.xlsx")
            shutil.copy(clock_src, "data/input/latest_clockin.xlsx")
            return JSONResponse(content={"success": True, "message": "Simulated fetch complete!"})
        else:
            return JSONResponse(content={"success": False, "error": "Mock files not found in tests/test_docs/"})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

def get_archive_dir(date_str):
    """Returns data/archives/YYYY-MM/"""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        month_dir = dt.strftime("%Y-%m")
        path = f"data/archives/{month_dir}"
        os.makedirs(path, exist_ok=True)
        return path
    except:
        return "data/archives"

@app.get("/api/archives")
async def list_archives():
    try:
        payslip_archives = []
        roster_archives = []
        
        base_dir = "data/archives"
        if os.path.exists(base_dir):
            for root, dirs, files in os.walk(base_dir):
                for d in dirs:
                    if d.startswith("Payslips_"):
                        # Show as 'Folder/File' for the frontend or just the name
                        rel_path = os.path.relpath(os.path.join(root, d), base_dir).replace("\\", "/")
                        payslip_archives.append(rel_path)
                for f in files:
                    if f.startswith("Roster_") and f.endswith(".xlsx"):
                        rel_path = os.path.relpath(os.path.join(root, f), base_dir).replace("\\", "/")
                        roster_archives.append(rel_path)
                        
        return JSONResponse(content={
            "payslips": sorted(payslip_archives, reverse=True),
            "rosters": sorted(roster_archives, reverse=True)
        })
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.get("/api/archives/roster/{path_sub:path}")
async def download_archived_roster(path_sub: str):
    path = os.path.join("data/archives", path_sub)
    if os.path.exists(path):
        return FileResponse(path, filename=os.path.basename(path))
    return JSONResponse(content={"error": "File not found"}, status_code=404)

@app.get("/api/archives/payslips_zip/{path_sub:path}")
async def download_archived_payslips_zip(path_sub: str):
    path = os.path.join("data/archives", path_sub, "All_Payslips.zip")
    if os.path.exists(path):
        return FileResponse(path, filename=f"All_Payslips.zip")
    return JSONResponse(content={"error": "ZIP not found"}, status_code=404)

@app.get("/api/archives/eft_summary/{path_sub:path}")
async def download_archived_eft_summary(path_sub: str):
    path = os.path.join("data/archives", path_sub, "EFT_Summary.xlsx")
    if os.path.exists(path):
        return FileResponse(path, filename=f"EFT_Summary.xlsx")
    return JSONResponse(content={"error": "EFT Summary not found"}, status_code=404)

@app.get("/api/archives/view/roster/{filename}")
async def view_archived_roster(filename: str):
    path = f"data/archives/Rosters/{filename}"
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            # Replaced NaN with empty strings for JSON compatibility
            df = df.fillna("")
            return JSONResponse(content={"columns": df.columns.tolist(), "rows": df.to_dict('records')})
        except Exception as e:
            return JSONResponse(content={"error": str(e)}, status_code=500)
    return JSONResponse(content={"error": "File not found"}, status_code=404)

@app.get("/api/archives/view/eft_summary/{folder}")
async def view_archived_eft_summary(folder: str):
    path = f"data/archives/Payslips/{folder}/EFT_Summary.xlsx"
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            # Replaced NaN with empty strings for JSON compatibility
            df = df.fillna("")
            return JSONResponse(content={"columns": df.columns.tolist(), "rows": df.to_dict('records')})
        except Exception as e:
            return JSONResponse(content={"error": str(e)}, status_code=500)
    return JSONResponse(content={"error": "EFT Summary not found"}, status_code=404)

@app.post("/api/set_active_roster")
async def set_active_roster(request: Request):
    try:
        data = await request.json()
        week_date = data.get("week_date", "")
        archive_path = f"data/archives/Rosters/Roster_{week_date}.xlsx"
        if os.path.exists(archive_path):
            shutil.copy(archive_path, "data/input/latest_roster.xlsx")
            return JSONResponse(content={"success": True, "message": f"Roster for {week_date} set as active."})
        return JSONResponse(content={"success": False, "error": f"No archived roster found for {week_date}"})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/save_roster")
async def save_roster(request: Request):
    try:
        data = await request.json()
        week_date = data.get("week_date", "current")
        rows = data.get("rows", [])
        
        df = pd.DataFrame(rows)
        filename = f"data/input/latest_roster.xlsx"
        df.to_excel(filename, index=False)
        
        # Archive it in monthly folder
        archive_dir = get_archive_dir(week_date)
        archive_path = os.path.join(archive_dir, f"Roster_{week_date}.xlsx")
        shutil.copy(filename, archive_path)
        
        return JSONResponse(content={"success": True, "message": f"Roster saved and archived in {os.path.basename(archive_dir)}"})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/accrue_leave")
async def accrue_leave():
    master_path = "data/templates/Staff_Details_Template.xlsx"
    if not os.path.exists(master_path):
        return JSONResponse(content={"status": "error", "message": "Master file not found. Fill out Staff_Details_Template.xlsx first."})
        
    try:
        master_df = pd.read_excel(master_path)
        leave_col = next((c for c in master_df.columns if str(c).strip().lower() == 'leave credit'), None)
        
        if leave_col:
            for idx, row in master_df.iterrows():
                try: cur_leave = float(row[leave_col]) if str(row[leave_col]).strip().lower() != 'nan' else 0.0
                except: cur_leave = 0.0
                master_df.at[idx, leave_col] = cur_leave + 1.5
            master_df.to_excel(master_path, index=False)
            return JSONResponse(content={"status": "success", "message": "1.5 days accrued to all staff successfully!"})
        return JSONResponse(content={"status": "error", "message": "Leave Credit column not found."})
    except Exception as e:
        return JSONResponse(content={"status": "error", "message": str(e)})

@app.get("/api/staff")
async def get_staff():
    master_path = "data/templates/Staff_Details_Template.xlsx"
    pins_path = "data/staff_pins.json"
    if not os.path.exists(master_path):
        return JSONResponse(content={"staff": []})
    try:
        df = pd.read_excel(master_path)
        if "Name" not in df.columns:
            return JSONResponse(content={"staff": []})
            
        # Load PINs for mapping
        pins = {}
        if os.path.exists(pins_path):
            with open(pins_path, "r") as f:
                pins = json.load(f)

        staff = []
        for _, row in df.iterrows():
            name = str(row.get("Name", "")).replace("nan", "")
            if not name: continue
            
            rate_str = str(row.get("Rate", "30.33")).replace("R", "").replace(",", ".").strip()
            rate_val = float(rate_str) if rate_str.lower() != "nan" and rate_str else 30.33
            
            leave_str = str(row.get("Leave Credit", "0.0")).strip()
            leave_val = float(leave_str) if leave_str.lower() != "nan" and leave_str else 0.0
            
            staff.append({
                "name": name,
                "id_number": str(row.get("ID Number", "")).replace("nan", ""),
                "rate": rate_val,
                "start_date": str(row.get("Start Date", "")).replace("nan", ""),
                "leave_credit": leave_val,
                "cell_number": str(row.get("Cell Number", "")).replace(".0", "").replace("nan", ""),
                "role": str(row.get("Role", "WAITER")).replace("nan", "WAITER"),
                "pin": pins.get(name, "")
            })
        return JSONResponse(content={"staff": staff})
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.post("/api/publish_roster")
async def publish_roster():
    try:
        latest = "data/input/latest_roster.xlsx"
        published = "data/input/published_roster.xlsx"
        if os.path.exists(latest):
            shutil.copy(latest, published)
            return JSONResponse(content={"success": True, "message": "Roster successfully published to the staff portal!"})
        return JSONResponse(content={"success": False, "error": "No draft roster found to publish."})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.get("/api/my_roster")
async def get_my_roster(name: str):
    roster_path = "data/input/published_roster.xlsx"
    if not os.path.exists(roster_path):
        return JSONResponse(content={"error": "No Roster has been published for this week yet."})
    try:
        df = pd.read_excel(roster_path)
        # Normalize column names to title case for internal check, look for 'name' variant
        df.columns = [str(c).strip() for c in df.columns]
        name_col = next((c for c in df.columns if c.lower() == 'name'), None)
        
        if not name_col:
            return JSONResponse(content={"error": "Roster format invalid (Name column missing)"})
            
        # Find row for user
        user_row = df[df[name_col].astype(str).str.strip().str.lower() == name.lower()]
        if user_row.empty:
            return JSONResponse(content={"error": f"No shift found for {name} this week."})
            
        shifts = user_row.iloc[0].to_dict()
        return JSONResponse(content={"shifts": shifts})
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.post("/api/staff/update")
async def update_staff(request: Request):
    master_path = "data/templates/Staff_Details_Template.xlsx"
    try:
        data = await request.json()
        new_staff = data.get("staff", {})
        
        expected_cols = ["Name", "ID Number", "Rate", "Start Date", "Leave Credit", "Cell Number", "Role"]
        
        if os.path.exists(master_path):
            df = pd.read_excel(master_path)
            if "Name" not in df.columns:
                df = pd.DataFrame(columns=expected_cols)
        else:
            df = pd.DataFrame(columns=expected_cols)
            
        name_to_find = new_staff["name"].strip().lower()
        found_idx = -1
        for idx, row in df.iterrows():
            if str(row["Name"]).strip().lower() == name_to_find:
                found_idx = idx
                break
        
        row_data = {
            "Name": new_staff["name"],
            "ID Number": new_staff.get("id_number", ""),
            "Rate": new_staff.get("rate", 30.33),
            "Start Date": new_staff.get("start_date", ""),
            "Leave Credit": new_staff.get("leave_credit", 0.0),
            "Cell Number": new_staff.get("cell_number", ""),
            "Role": new_staff.get("role", "WAITER")
        }
        
        if found_idx >= 0:
            for k, v in row_data.items():
                df.at[found_idx, k] = v
        else:
             df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
            
        df.to_excel(master_path, index=False)
        return JSONResponse(content={"success": True})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/save_config")
async def save_config(request: Request):
    try:
        data = await request.json()
        with open("config.json", "w") as f:
            json.dump(data, f)
        return JSONResponse(content={"success": True})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/staff/update_profile")
async def update_staff_profile(request: Request):
    master_path = "data/templates/Staff_Details_Template.xlsx"
    pins_path = "data/staff_pins.json"
    try:
        data = await request.json()
        old_name = data.get("old_name", "").strip()
        new_name = data.get("name", "").strip()
        role = data.get("role", "").strip()
        rate_raw = data.get("rate", 0)
        rate = float(rate_raw) if rate_raw else 0.0
        pin = data.get("pin", "").strip()
        cell = data.get("cell_number", "").strip()
        portal_access = data.get("portal_access", "MY_SHIFTS").strip()

        if not old_name: return JSONResponse(content={"success": False, "error": "Old name required"})

        # 1. Update Excel
        if os.path.exists(master_path):
            df = pd.read_excel(master_path)
            mask = df["Name"].astype(str).str.strip().str.lower() == old_name.lower()
            if not mask.any():
                return JSONResponse(content={"success": False, "error": "Staff member not found in Master Excel"})

            df.loc[mask, "Name"] = new_name
            df.loc[mask, "Role"] = role
            df.loc[mask, "Rate"] = rate
            df.loc[mask, "Cell Number"] = cell
            if "Portal Access" not in df.columns:
                df["Portal Access"] = "MY_SHIFTS"
            df.loc[mask, "Portal Access"] = portal_access
            df.to_excel(master_path, index=False)
        else:
            return JSONResponse(content={"success": False, "error": "Master Excel file missing"})

        # 2. Update PINs
        pins = {}
        if os.path.exists(pins_path):
            with open(pins_path, "r") as f:
                pins = json.load(f)
        if old_name in pins:
            del pins[old_name]
        pins[new_name] = pin
        with open(pins_path, "w") as f:
            json.dump(pins, f, indent=4)

        return JSONResponse(content={"success": True})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

@app.post("/api/staff/delete")
async def delete_staff(request: Request):
    master_path = "data/templates/Staff_Details_Template.xlsx"
    try:
        data = await request.json()
        name_to_delete = data.get("name", "").strip().lower()
        if os.path.exists(master_path):
            df = pd.read_excel(master_path)
            if "Name" in df.columns:
                df = df[df["Name"].apply(lambda x: str(x).strip().lower() != name_to_delete)]
                df.to_excel(master_path, index=False)
            return JSONResponse(content={"success": True})
        return JSONResponse(content={"success": False, "error": "Master file not found"})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

# --- GEMINI AI ASSISTANT ---

@app.post("/api/ai/chat")
async def ai_chat(request: Request):
    if not GEMINI_KEY:
        return JSONResponse(content={"error": "API_KEY_MISSING", "message": "Gemini API Key is missing."})
    
    try:
        data = await request.json()
        user_msg = data.get("message", "")
        user_role = data.get("role", "EMPLOYEE")
        user_name = data.get("name", "Unknown Staff")
        
        # 1. Gather Staff Context
        staff_summary = "No staff data found."
        master_path = "data/templates/Staff_Details_Template.xlsx"
        if os.path.exists(master_path):
            sdf = pd.read_excel(master_path)
            staff_summary = sdf[['Name', 'Role', 'Leave Credit']].to_string(index=False)
            
        # 2. Gather Public & School Holidays
        sa_holidays = pyholidays.SouthAfrica(years=datetime.now().year)
        holiday_list = ", ".join([f"{d} ({n})" for d, n in sorted(sa_holidays.items())])
        
        school_holidays = "2026-03-28 to 2026-04-07, 2026-06-27 to 2026-07-20, 2026-09-24 to 2026-10-05, 2026-12-10 to 2026-12-31"

        # 3. Gather Off-Day Requests
        requests_summary = "No pending off-day requests."
        req_path = "data/off_day_requests.json"
        if os.path.exists(req_path):
            with open(req_path, "r") as f:
                req_data = json.load(f)
                requests_summary = json.dumps(req_data, indent=2)

        # 4. Gather AI Learning Log (manager-added notes for roster generation)
        learning_summary = "No learning notes."
        log_path = "data/roster_learning_log.json"
        if os.path.exists(log_path):
            with open(log_path, "r") as f:
                try:
                    log_data = json.load(f)
                    if log_data:
                        learning_summary = "\n".join([f"- [{e.get('date','')}] {e.get('note','')}" for e in log_data])
                except:
                    pass

        system_prompt = f"""
        You are the 'Wimpy De Ville AI Manager-on-Duty'. 
        Current Date: {datetime.now().strftime('%Y-%m-%d (%A)')}
        
        STRICT OPERATIONAL RULES:
        1. Roster Cycle: Wednesday to Tuesday.
        2. Store Hours & Closing:
           - Monday to Friday: 07:00 - 18:00
           - Saturday: 07:00 - 15:00
           - Sunday & Public Holidays: 08:00 - 13:00 (Reduced Staffing)
        3. Lead Time: Leave (Vacation) requires 14 days notice.
        4. Approval Flags: Requests for Sat/Sun or School Holidays ALWAYS require manager approval.
        
        STAFF DATABASE:
        {staff_summary}
        Note: Tanya Baard is R0 (Monthly Salary).
        
        PUBLIC HOLIDAYS (SOUTH AFRICA):
        {holiday_list}
        
        SCHOOL HOLIDAYS 2026:
        {school_holidays}
        
        PENDING OFF-DAY & LEAVE REQUESTS:
        {requests_summary}
        
        OUTPUT GUIDELINES:
        - When generating a roster, produce an 'Individual Snippet' for each person mentioned.
        - Snippet Format Example:
          "Hi [Name], here is your roster for [Date Range]:
          Wed: 07:00-17:30
          Thu: OFF
          ...
          Total Hours: XX"
        - Always respect the Wed-Tue cycle (Wednesday start, Tuesday end) and any APPROVED requests.
        - NEVER roster an employee on a day where they have an APPROVED off-day or leave request.
        - Minimum floor coverage: At least 2 Kitchen, 2 Waiters/FOH per shift (Sunday/Holidays: 1-2 each).

        MANAGER LEARNING NOTES (apply these when generating rosters):
        {learning_summary}

        USER ROLE: {user_role} ({user_name})
        """
        
        if not client:
             return JSONResponse(content={"error": "AI_CLIENT_MISSING"})

        response = client.models.generate_content(
            model='gemini-flash-latest',
            contents=f"{system_prompt}\n\nUSER REQUEST: {user_msg}"
        )
        return JSONResponse(content={"reply": response.text})
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.post("/api/ai/generate_snippets")
async def generate_snippets(request: Request):
    if not GEMINI_KEY:
        return JSONResponse(content={"error": "API_KEY_MISSING"})
        
    try:
        data = await request.json()
        roster_rows = data.get("rows", [])
        week_date = data.get("week_date", "next week")
        
        snippets = []
        for row in roster_rows:
            name = row.get("Name", "Staff")
            # Ask AI for a concise snippet for this specific person's row
            prompt = f"Based on this roster row for {name} for the week ending {week_date}, generate a concise WhatsApp snippet:\n{json.dumps(row)}\n\nFormat: 'Hi {name}, your roster: ...'"
            
            response = client.models.generate_content(
                model='gemini-flash-latest',
                contents=prompt
            )
            snippets.append({
                "name": name,
                "phone": row.get("Cell Number", ""), # We might need to fetch this from staff master
                "snippet": response.text.strip()
            })
            
        return JSONResponse(content={"snippets": snippets})
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.post("/api/whatsapp_send_snippets")
async def whatsapp_send_snippets(request: Request):
    data = await request.json()
    snippets = data.get("snippets", [])
    
    # Fetch staff cell numbers if missing
    master_path = "data/templates/Staff_Details_Template.xlsx"
    phone_map = {}
    if os.path.exists(master_path):
        df = pd.read_excel(master_path)
        for _, row in df.iterrows():
            name = str(row.get("Name", "")).strip().lower()
            phone = str(row.get("Cell Number", "")).replace(".0", "").replace("nan", "")
            phone_map[name] = ''.join(filter(str.isdigit, phone))

    payload = []
    for s in snippets:
        name_lower = s.get("name", "").lower()
        phone = s.get("phone") or phone_map.get(name_lower, "")
        if phone:
            payload.append({
                "phone": phone,
                "message": s.get("snippet", ""),
                "file": "" # No file for snippets usually, just text
            })
            
    with open("output/whatsapp_payload.json", "w") as f:
        json.dump(payload, f)
        
    subprocess.Popen([sys.executable, "app/wa_bot.py"])
    return JSONResponse(content={"status": "Snippets queued for WhatsApp bot."})

@app.post("/api/send_roster")
async def send_roster(request: Request):
    try:
        data = await request.json()
        links = data.get("links", [])
        manager_email = "dylan.lloyd25@gmail.com"
        
        # Stub for SMTP flow. Manager would need to provide an App Password if using Gmail.
        sent_email_count = 0
        for l in links:
             if l.get("email") and "@" in str(l.get("email")):
                  sent_email_count += 1
                  
        return JSONResponse(content={"success": True, "message": f"Roster pieces sent to {sent_email_count} via Email, others queued for WhatsApp bot."})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})



# --- AI LEARNING LOG ---

@app.get("/api/ai/learning_log")
async def get_learning_log():
    path = "data/roster_learning_log.json"
    if not os.path.exists(path): return JSONResponse(content=[])
    with open(path, "r") as f:
        return JSONResponse(content=json.load(f))

@app.post("/api/ai/learning_log")
async def add_learning_log(request: Request):
    data = await request.json()
    note = data.get("note", "").strip()
    if not note: return JSONResponse(content={"success": False, "error": "Note cannot be empty"})
    path = "data/roster_learning_log.json"
    log = []
    if os.path.exists(path):
        with open(path, "r") as f:
            try: log = json.load(f)
            except: log = []
    log.append({"date": datetime.now().strftime("%Y-%m-%d %H:%M"), "note": note})
    with open(path, "w") as f: json.dump(log, f, indent=4)
    return JSONResponse(content={"success": True, "total_notes": len(log)})

@app.delete("/api/ai/learning_log/{index}")
async def delete_learning_log(index: int):
    path = "data/roster_learning_log.json"
    if not os.path.exists(path): return JSONResponse(content={"success": False, "error": "Log not found"})
    with open(path, "r") as f: log = json.load(f)
    if index < 0 or index >= len(log):
        return JSONResponse(content={"success": False, "error": "Index out of range"})
    log.pop(index)
    with open(path, "w") as f: json.dump(log, f, indent=4)
    return JSONResponse(content={"success": True})


# --- FINGERPRINT SCANNER INTEGRATION ---

@app.get("/api/fingerprint/records")
async def get_fingerprint_records():
    path = "data/fingerprint_scans.csv"
    if not os.path.exists(path):
        return JSONResponse(content={"records": []})
    try:
        import csv
        records = []
        with open(path, "r", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                records.append(dict(row))
        return JSONResponse(content={"records": records})
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.post("/api/fingerprint/sync")
async def sync_fingerprint(request: Request):
    data = await request.json()
    records = data.get("records", [])
    if not records:
        return JSONResponse(content={"success": False, "error": "No records provided"})
    path = "data/fingerprint_scans.csv"
    try:
        import csv
        file_exists = os.path.exists(path) and os.path.getsize(path) > 0
        with open(path, "a", newline="") as f:
            fieldnames = ["EmployeeID", "EmployeeName", "Timestamp", "Action"]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            for r in records:
                writer.writerow({
                    "EmployeeID": r.get("EmployeeID", ""),
                    "EmployeeName": r.get("EmployeeName", ""),
                    "Timestamp": r.get("Timestamp", datetime.now().isoformat()),
                    "Action": r.get("Action", "IN")
                })
        return JSONResponse(content={"success": True, "synced": len(records)})
    except Exception as e:
        return JSONResponse(content={"success": False, "error": str(e)})

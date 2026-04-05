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

import google.generativeai as genai
from .payroll_processor import process_payroll
from .payslip_generator import generate_payslip, PayslipPDF
from fpdf import FPDF
try:
    import win32clipboard
except ImportError:
    win32clipboard = None

# --- GEMINI SETUP ---
GEMINI_KEY = None
if os.path.exists("config.json"):
    try:
        with open("config.json", "r") as f:
            cfg = json.load(f)
            GEMINI_KEY = cfg.get("GEMINI_API_KEY")
    except: pass

if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)

app = FastAPI()

os.makedirs("static", exist_ok=True)
os.makedirs("output/payslips", exist_ok=True)
os.makedirs("data/archives/Rosters", exist_ok=True)
os.makedirs("data/archives/Payslips", exist_ok=True)

app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/payslips", StaticFiles(directory="output/payslips"), name="payslips")

@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

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

    output_df, summary_df = process_payroll(roster_path, clockin_path, output_format='dataframe')
    
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

    return JSONResponse(content={"employees": employees})

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
            from datetime import datetime
            week_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
            archive_dir = f"data/archives/Payslips/Payroll_{week_str}"
            os.makedirs(archive_dir, exist_ok=True)
            shutil.make_archive(f"{archive_dir}/All_Payslips", "zip", "output/payslips")
            
            rows_eft = []
            total_sum = 0
            for l in links:
                net_val = round(l.get("net", 0.0), 2)
                rows_eft.append({"Employee Name": l.get("name"), "Final Net Pay": net_val})
                total_sum += net_val
            
            rows_eft.append({"Employee Name": "TOTAL", "Final Net Pay": round(total_sum, 2)})
            pd.DataFrame(rows_eft).to_excel(f"{archive_dir}/EFT_Summary.xlsx", index=False)
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

@app.get("/api/archives")
async def list_archives():
    try:
        payslip_archives = []
        p_dir = "data/archives/Payslips"
        if os.path.exists(p_dir):
            for d in os.listdir(p_dir):
                if os.path.isdir(os.path.join(p_dir, d)):
                    payslip_archives.append(d)
        
        roster_archives = []
        r_dir = "data/archives/Rosters"
        if os.path.exists(r_dir):
            for f in os.listdir(r_dir):
                if f.endswith(".xlsx"):
                    roster_archives.append(f)
                    
        return JSONResponse(content={
            "payslips": sorted(payslip_archives, reverse=True),
            "rosters": sorted(roster_archives, reverse=True)
        })
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

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
        
        # Create Excel consistent with Test roster.xlsx
        # Format: NAME, MON, TUE, WED, THU, FRI, SAT, SUN, TOTAL
        df = pd.DataFrame(rows)
        
        filename = f"data/input/latest_roster.xlsx"
        df.to_excel(filename, index=False)
        
        # Archive it
        archive_path = f"data/archives/Rosters/Roster_{week_date}.xlsx"
        shutil.copy(filename, archive_path)
        
        return JSONResponse(content={"success": True, "message": f"Roster saved and archived as {week_date}"})
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
    if not os.path.exists(master_path):
        return JSONResponse(content={"staff": []})
    try:
        df = pd.read_excel(master_path)
        if "Name" not in df.columns:
            return JSONResponse(content={"staff": []})
            
        staff = []
        for _, row in df.iterrows():
            rate_str = str(row.get("Rate", "30.33")).replace("R", "").replace(",", ".").strip()
            rate_val = float(rate_str) if rate_str.lower() != "nan" and rate_str else 30.33
            
            leave_str = str(row.get("Leave Credit", "0.0")).strip()
            leave_val = float(leave_str) if leave_str.lower() != "nan" and leave_str else 0.0
            
            staff.append({
                "name": str(row.get("Name", "")).replace("nan", ""),
                "id_number": str(row.get("ID Number", "")).replace("nan", ""),
                "rate": rate_val,
                "start_date": str(row.get("Start Date", "")).replace("nan", ""),
                "leave_credit": leave_val,
                "cell_number": str(row.get("Cell Number", "")).replace(".0", "").replace("nan", ""),
                "role": str(row.get("Role", "WAITER")).replace("nan", "WAITER")
            })
        return JSONResponse(content={"staff": staff})
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
        return JSONResponse(content={"error": "API_KEY_MISSING", "message": "Gemini API Key is missing. Please add your key to config.json to enable the AI Manager Assistant."})
    
    try:
        data = await request.json()
        user_msg = data.get("message", "")
        
        # 1. Gather Staff Context
        staff_summary = "No staff data found."
        if os.path.exists("data/templates/Staff_Details_Template.xlsx"):
            sdf = pd.read_excel("data/templates/Staff_Details_Template.xlsx")
            staff_summary = sdf[['Name', 'Role', 'Rate', 'Leave Credit']].to_string(index=False)
            
        # 2. Gather Active Roster Context
        active_roster = "No active roster found."
        if os.path.exists("data/input/latest_roster.xlsx"):
            rdf = pd.read_excel("data/input/latest_roster.xlsx")
            active_roster = rdf.to_string(index=False)
            
        system_prompt = f"""
        You are the 'Wimpy De Ville AI Manager'. You help the owner manage staff, rosters, and payroll.
        
        HERE IS YOUR STAFF DATA:
        {staff_summary}
        
        HERE IS THE CURRENT ACTIVE ROSTER:
        {active_roster}
        
        GUIDELINES:
        1. ROSTER AUTOMATION: If the user asks for a roster (e.g. 'Build a roster' or 'Someone is off'), 
           suggest shifts using the staff above. Match ROLES (Waiters do Waiter shifts, Grillers do Grillers).
           Standard shifts are: 07:00-15:30 (8.5h), 07:00-17:30 (10.5h), 09:00-17:30 (8.5h).
        2. FORMATTING: If you suggest a roster, format it clearly. 
        3. DATA LOOKUP: Answer questions about leave balances or roles instantly.
        4. TONE: Professional, efficient, restaurant manager style.
        """
        
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(f"{system_prompt}\n\nUSER REQUEST: {user_msg}")
        
        return JSONResponse(content={"reply": response.text})
    except Exception as e:
        return JSONResponse(content={"error": "AI_ERROR", "message": f"Gemini Error: {str(e)}"})


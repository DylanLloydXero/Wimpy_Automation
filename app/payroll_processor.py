import pandas as pd
import argparse
import sys
from fpdf import FPDF
from datetime import timedelta
import difflib
import re

# --- CONFIGURATION (Change these when exact CSV formats are known) ---
ROSTER_EMP_COL = 'Employee Name'
ROSTER_DATE_COL = 'Date'
ROSTER_START_COL = 'Rostered Start'
ROSTER_END_COL = 'Rostered End'

CLOCKIN_EMP_COL = 'Employee Name'
CLOCKIN_DATE_COL = 'Date'
CLOCKIN_START_COL = 'Clock In Time'
CLOCKIN_END_COL = 'Clock Out Time'
CLOCKIN_APPROVED_OT_COL = 'Approved Overtime' # Boolean or Yes/No

def preprocess_roster(df):
    name_col = next((c for c in df.columns if 'name' in str(c).lower()), df.columns[0])
    dates = [c for c in df.columns if c != name_col]
    rows = []
    for _, row in df.iterrows():
        emp_name = row[name_col]
        if pd.isna(emp_name) or str(emp_name).strip().lower() == 'name':
            continue
        emp_name = str(emp_name).strip()
        for date_col in dates:
            time_str = row[date_col]
            clean_date_str = str(date_col).replace('20206', '2026')
            
            try:
                pd.to_datetime(clean_date_str)
            except (ValueError, TypeError):
                continue
                
            start_str, end_str = None, None
            if not pd.isna(time_str) and '-' in str(time_str):
                try:
                    parts = [s.strip() for s in str(time_str).split('-')]
                    start_str = parts[0]
                    end_str = parts[1].split()[0] if ' ' in parts[1] else parts[1]
                except:
                    pass
                
            rows.append({
                'Employee Name': emp_name, 
                'Date': pd.to_datetime(clean_date_str).strftime('%Y-%m-%d'),
                'Rostered Start': start_str, 
                'Rostered End': end_str
            })
    return pd.DataFrame(rows)

def preprocess_clockin(df):
    rows = []
    for _, row in df.iterrows():
        server_name = row.get('Server Name')
        if pd.isna(server_name) or str(server_name).startswith(('Company:', 'Store:')):
            continue
        
        time_in = row.get('Time Logged In')
        time_out = row.get('Time Logged Out')
        
        if pd.isna(time_in):
            continue
            
        try:
            d_in = pd.to_datetime(time_in, dayfirst=True)
            if pd.isna(time_out):
                d_out = d_in
            else:
                d_out = pd.to_datetime(time_out, dayfirst=True)
                
            rows.append({
                'Employee Name': str(server_name).strip(), 
                'Date': d_in.strftime('%Y-%m-%d'),
                'Clock In Time': d_in.strftime('%H:%M:%S'), 
                'Clock Out Time': d_out.strftime('%H:%M:%S')
            })
        except Exception:
            pass
    return pd.DataFrame(rows)

def match_names(roster_df, clockin_df, manual_mappings=None):
    roster_names = roster_df['Employee Name'].unique() if not roster_df.empty else []
    clockin_names = clockin_df['Employee Name'].unique() if not clockin_df.empty else []
    mapping = {}
    unmatched = []
    manual_mappings = manual_mappings or {}
    
    for c_name in clockin_names:
        if str(c_name) in manual_mappings:
            mapping[c_name] = manual_mappings[str(c_name)]
            continue
        c_clean = re.split(r'\(|Total:', str(c_name), flags=re.IGNORECASE)[0].strip().split()[0].lower()
        best_match, best_score = None, 0
        for r_name in roster_names:
            r_clean = str(r_name).strip().lower()
            if c_clean.startswith(r_clean) or r_clean.startswith(c_clean):
                best_match = r_name
                break
            score = difflib.SequenceMatcher(None, r_clean, c_clean).ratio()
            if score > best_score and score > 0.6:
                best_match, best_score = r_name, score
        if best_match: mapping[c_name] = best_match
        else: unmatched.append(str(c_name))
            
    clockin_df['Employee Name'] = clockin_df['Employee Name'].map(lambda x: mapping.get(x, x))
    
    return roster_df, clockin_df, unmatched

def process_payroll(roster_path, clockin_path, holidays=None, output_format='excel', output_path='payable_hours', week_ending=None, manual_mappings=None):
    unmatched_names = []
    try:
        roster_df = preprocess_roster(pd.read_excel(roster_path))
        clockin_df = preprocess_clockin(pd.read_excel(clockin_path))
        
        if roster_df.empty or clockin_df.empty:
            return (None, pd.DataFrame(columns=['Employee Name']), []) if output_format == 'dataframe' else None
            
        roster_df, clockin_df, unmatched_names = match_names(roster_df, clockin_df, manual_mappings)
    except Exception as e:
        print(f"Error reading Excel files: {e}")
        return (None, None, []) if output_format == 'dataframe' else None

    # Merge dataframes on Employee and Date - USE RIGHT JOIN to capture unscheduled clock-ins
    merged = pd.merge(roster_df, clockin_df, 
                      left_on=[ROSTER_EMP_COL, ROSTER_DATE_COL], 
                      right_on=[CLOCKIN_EMP_COL, CLOCKIN_DATE_COL],
                      how='right')
    
    if week_ending:
        try:
            from datetime import datetime, timedelta
            we_date = pd.to_datetime(week_ending)
            start_we = we_date - timedelta(days=6)
            merged[CLOCKIN_DATE_COL] = pd.to_datetime(merged[CLOCKIN_DATE_COL])
            merged = merged[(merged[CLOCKIN_DATE_COL] >= start_we) & (merged[CLOCKIN_DATE_COL] <= we_date)]
            merged[CLOCKIN_DATE_COL] = merged[CLOCKIN_DATE_COL].dt.strftime('%Y-%m-%d')
        except Exception as e:
            print(f"Filtering by week_ending {week_ending} failed: {e}")

    if merged.empty:
        return (None, pd.DataFrame(columns=['Employee Name']), []) if output_format == 'dataframe' else None

    results = []
    overtime_flags = []
    import json, os
    approvals_path = "data/overtime_approvals.json"
    approvals = {}
    if os.path.exists(approvals_path):
        try:
            with open(approvals_path, "r") as f:
                approvals = json.load(f)
        except Exception:
            pass

    holidays = holidays or []

    for idx, row in merged.iterrows():
        try:
            date_str = str(row[CLOCKIN_DATE_COL])
            emp_name = str(row[CLOCKIN_EMP_COL])
            date_val = pd.to_datetime(date_str)
            
            is_holiday = date_str in holidays
            is_sunday = (date_val.weekday() == 6) and not is_holiday
            is_tuesday = date_val.weekday() == 1
            
            clock_in = pd.to_datetime(date_str + ' ' + row[CLOCKIN_START_COL])
            clock_out = pd.to_datetime(date_str + ' ' + row[CLOCKIN_END_COL])
            raw_clock_hours = max(0, (clock_out - clock_in).total_seconds() / 3600.0)
            
            # 1. Determine base start/end
            is_unscheduled = pd.isna(row['Rostered Start'])
            
            final_start = clock_in
            final_end = clock_out
            
            if is_unscheduled:
                # -- UNSCHEDULED SHIFT LOGIC --
                flag_key = f"{emp_name}_{date_str}"
                flag_status = approvals.get(flag_key, "PENDING")
                
                if flag_status == "APPROVED":
                    pass # Keep full clock hours
                else:
                    final_start = final_end # 0 hours
                
                overtime_flags.append({
                    "employee": emp_name,
                    "date": date_str,
                    "type": "UNSCHEDULED",
                    "rost_end": "OFF",
                    "clock_out": clock_out.strftime('%H:%M'),
                    "overtime_hours": round(raw_clock_hours, 2),
                    "status": flag_status
                })
            else:
                # Scheduled shift logic
                rost_start = pd.to_datetime(date_str + ' ' + row['Rostered Start'])
                rost_end = pd.to_datetime(date_str + ' ' + row['Rostered End'])
                
                final_start = max(rost_start, clock_in)
                final_end = min(rost_end, clock_out)
                
                # -- OVERTIME DETECTION --
                overtime_diff_mins = (clock_out - rost_end).total_seconds() / 60.0
                if overtime_diff_mins > 10:
                    flag_key = f"{emp_name}_{date_str}"
                    flag_status = approvals.get(flag_key, "PENDING")
                    
                    if flag_status == "APPROVED":
                        final_end = clock_out
                    elif flag_status == "DENIED":
                        final_end = rost_end
                    elif flag_status == "PENDING":
                        final_end = rost_end # cap until approved
                        
                    overtime_flags.append({
                        "employee": emp_name,
                        "date": date_str,
                        "type": "OVERTIME",
                        "rost_end": rost_end.strftime('%H:%M'),
                        "clock_out": clock_out.strftime('%H:%M'),
                        "overtime_hours": round(overtime_diff_mins / 60.0, 2),
                        "status": flag_status
                    })

            # 2. Final Hour Calculation
            hours = (final_end - final_start).total_seconds() / 3600.0
            
            # --- BREAK DEDUCTION RULE (30 mins if > 5.0 hours) ---
            if hours > 5.0:
                hours -= 0.5
                
            hours_val = max(0, hours)
            
            results.append({
                'Employee Name': emp_name,
                'Date': date_str,
                'Reg Hours': hours_val if (not is_sunday and not is_tuesday and not is_holiday) else 0,
                'Sun Hours': hours_val if is_sunday else 0,
                'Tue Hours': hours_val if is_tuesday else 0,
                'Hol Hours': hours_val if is_holiday else 0,
                'Total Payable Hours': hours_val,
                'Raw Rostered Hours': 0 if is_unscheduled else max(0, (pd.to_datetime(row['Rostered End']) - pd.to_datetime(row['Rostered Start'])).total_seconds() / 3600.0),
                'Raw Clocked Hours': raw_clock_hours
            })
        except Exception:
            pass
            
    if not results:
        return (None, pd.DataFrame(columns=['Employee Name']), []) if output_format == 'dataframe' else None
        
    df_results = pd.DataFrame(results)
    summary = df_results.groupby('Employee Name').agg({
        'Reg Hours': 'sum',
        'Sun Hours': 'sum',
        'Tue Hours': 'sum',
        'Hol Hours': 'sum',
        'Total Payable Hours': 'sum',
        'Raw Rostered Hours': 'sum',
        'Raw Clocked Hours': 'sum'
    }).reset_index()
    
    if output_format == 'dataframe':
        return df_results, summary, overtime_flags, unmatched_names
        
    if output_format == 'excel':
        summary.to_excel(f"{output_path}.xlsx", index=False)
        return f"{output_path}.xlsx", unmatched_names
    
    return None

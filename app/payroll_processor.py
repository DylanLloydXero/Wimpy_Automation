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
            
            # Skip non-date columns like "TOTAL"
            try:
                pd.to_datetime(clean_date_str)
            except (ValueError, TypeError):
                continue
                
            if pd.isna(time_str) or '-' not in str(time_str):
                continue
                
            try:
                start_str, end_str = [s.strip() for s in str(time_str).split('-')]
                # Clean up notes like "(8.5)" at the end of the time string
                end_str = end_str.split()[0] if ' ' in end_str else end_str
                
                rows.append({
                    'Employee Name': emp_name, 
                    'Date': pd.to_datetime(clean_date_str).strftime('%Y-%m-%d'),
                    'Rostered Start': start_str, 
                    'Rostered End': end_str
                })
            except Exception:
                pass
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
                # Missing out time = 0 hours for the day
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

def match_names(roster_df, clockin_df):
    roster_names = roster_df['Employee Name'].unique() if not roster_df.empty else []
    clockin_names = clockin_df['Employee Name'].unique() if not clockin_df.empty else []
    mapping = {}
    for c_name in clockin_names:
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
        if best_match:
            mapping[c_name] = best_match
    clockin_df['Employee Name'] = clockin_df['Employee Name'].map(lambda x: mapping.get(x, x))
    return roster_df, clockin_df

def process_payroll(roster_path, clockin_path, output_format='excel', output_path='payable_hours'):
    try:
        roster_df = preprocess_roster(pd.read_excel(roster_path))
        clockin_df = preprocess_clockin(pd.read_excel(clockin_path))
        
        if roster_df.empty or clockin_df.empty:
            print("Warning: One of the dataframes is empty after preprocessing.")
            return None, pd.DataFrame(columns=['Employee Name']) if output_format == 'dataframe' else None
            
        roster_df, clockin_df = match_names(roster_df, clockin_df)
    except Exception as e:
        print(f"Error reading Excel files: {e}")
        return None, None if output_format == 'dataframe' else None

    # Merge dataframes on Employee and Date
    merged = pd.merge(roster_df, clockin_df, 
                      left_on=[ROSTER_EMP_COL, ROSTER_DATE_COL], 
                      right_on=[CLOCKIN_EMP_COL, CLOCKIN_DATE_COL],
                      how='inner')
    
    if merged.empty:
        print("Warning: No matching employee/date records found between roster and clock-in.")
        return None, pd.DataFrame(columns=['Employee Name']) if output_format == 'dataframe' else None

    results = []
    for idx, row in merged.iterrows():
        try:
            # Parse Times
            date_val = pd.to_datetime(row[ROSTER_DATE_COL])
            is_sunday = date_val.weekday() == 6
            is_tuesday = date_val.weekday() == 1
            
            rost_start = pd.to_datetime(row[ROSTER_DATE_COL] + ' ' + row[ROSTER_START_COL])
            rost_end = pd.to_datetime(row[ROSTER_DATE_COL] + ' ' + row[ROSTER_END_COL])
            
            if is_tuesday:
                # Use full roster hours for Tuesday (Shift Hours)
                elapsed = (rost_end - rost_start).total_seconds() / 3600.0
                hours_val = max(0, elapsed)
            else:
                clock_in = pd.to_datetime(row[ROSTER_DATE_COL] + ' ' + row[CLOCKIN_START_COL])
                clock_out = pd.to_datetime(row[ROSTER_DATE_COL] + ' ' + row[CLOCKIN_END_COL])
                
                # Use later start and earlier finish
                start = max(rost_start, clock_in)
                end = min(rost_end, clock_out)
                
                hours = (end - start).total_seconds() / 3600.0
                hours_val = max(0, hours)
            
            results.append({
                'Employee Name': row[ROSTER_EMP_COL],
                'Date': row[ROSTER_DATE_COL],
                'Reg Hours': hours_val if (not is_sunday and not is_tuesday) else 0,
                'Sun Hours': hours_val if is_sunday else 0,
                'Tue Hours': hours_val if is_tuesday else 0,
                'Total Payable Hours': hours_val
            })
        except Exception:
            pass
            
    if not results:
        return None, pd.DataFrame(columns=['Employee Name']) if output_format == 'dataframe' else None
        
    df_results = pd.DataFrame(results)
    summary = df_results.groupby('Employee Name').agg({
        'Reg Hours': 'sum',
        'Sun Hours': 'sum',
        'Tue Hours': 'sum',
        'Total Payable Hours': 'sum'
    }).reset_index()
    
    if output_format == 'dataframe':
        return df_results, summary
        
    if output_format == 'excel':
        summary.to_excel(f"{output_path}.xlsx", index=False)
        return f"{output_path}.xlsx"
    
    return None

import os
from fpdf import FPDF
from datetime import datetime

class PayslipPDF(FPDF):
    def header(self):
        logo_path = os.path.join(os.getcwd(), 'data', 'templates', 'logo.png')
        if os.path.exists(logo_path):
            self.image(logo_path, 10, 8, 30)
            
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 10, 'P R I V A T E   &   C O N F I D E N T I A L', align='C', new_x="LMARGIN", new_y="NEXT")
        
        self.set_font('Helvetica', '', 10)
        self.cell(0, 5, "Store Name: WIMPY DE VILLE CENTER", align='C', new_x="LMARGIN", new_y="NEXT")
        self.cell(0, 5, "Address: Shop 03, De Ville Centre, Wellington Rd, Durbanville, Cape Town, 7550", align='C', new_x="LMARGIN", new_y="NEXT")
        self.cell(0, 5, "TEL: 021 975 9492", align='C', new_x="LMARGIN", new_y="NEXT")
        
        self.ln(10)

def generate_payslip(emp_data, output_dir="output/payslips"):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf = PayslipPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)

    # Employee Details Table
    def add_row(col1, col2, bold_col1=True, bold_col2=False):
        pdf.set_font("Helvetica", 'B' if bold_col1 else '', 10)
        pdf.cell(80, 8, col1, border=1)
        pdf.set_font("Helvetica", 'B' if bold_col2 else '', 10)
        pdf.cell(110, 8, col2, border=1, new_x="LMARGIN", new_y="NEXT", align='R')

    add_row("EMPLOYEE", emp_data.get('name', ''))
    add_row("ID NUMBER", emp_data.get('id_number', ''))
    add_row("POSITION", emp_data.get('role', 'WAITER').upper())
    add_row("START DATE", emp_data.get('start_date', ''))
    add_row("PAY PERIOD- WEEK ENDING", emp_data.get('week_ending', ''))
    add_row("BASIC | GROSS", "")
    add_row("PAYMENT MANNER", "EFT")
    add_row("LEAVE CREDIT: 1 Jan 2026", str(emp_data.get('leave_credit', '')))
    
    pdf.ln(5)

    # Breakdown Table
    pdf.set_font("Helvetica", 'B', 10)
    pdf.cell(70, 8, "DESCRIPTION", border=1)
    pdf.cell(40, 8, "RATE", border=1, align='C')
    pdf.cell(40, 8, "NO OF HOURS", border=1, align='C')
    pdf.cell(40, 8, "AMOUNT", border=1, new_x="LMARGIN", new_y="NEXT", align='C')

    pdf.set_font("Helvetica", '', 10)
    
    rate = float(emp_data.get('rate', 30.33))
    reg_hours = float(emp_data.get('reg_hours', 0.0))
    sun_hours = float(emp_data.get('sun_hours', 0.0))
    tue_hours = float(emp_data.get('tue_hours', 0.0))
    hol_hours = float(emp_data.get('hol_hours', 0.0))
    leave_days = float(emp_data.get('leave_days', 0.0))
    sick_days = float(emp_data.get('sick_days', 0.0))
    boh_share = float(emp_data.get('boh_tip_share', 0.0))
    shortfall = float(emp_data.get('tips_shortfall', 0.0))
    till_short = float(emp_data.get('till_short', 0.0))
    
    normal_amt = rate * reg_hours
    tue_amt = rate * tue_hours
    sunday_rate = rate * 1.5
    sunday_amt = sunday_rate * sun_hours
    hol_rate = rate * 2.0
    hol_amt = hol_rate * hol_hours
    leave_hrs = leave_days * 7
    leave_amt = rate * leave_hrs
    sick_hrs = sick_days * 7
    sick_amt = rate * sick_hrs
    
    def add_breakdown_row(desc, r, hrs, amt):
        pdf.cell(70, 8, desc, border=1)
        pdf.cell(40, 8, f"R {r:.2f}" if r else "", border=1, align='R')
        pdf.cell(40, 8, str(hrs) if hrs else "", border=1, align='C')
        pdf.cell(40, 8, f"R {amt:.2f}" if amt else "", border=1, new_x="LMARGIN", new_y="NEXT", align='R')

    add_breakdown_row("Normal", rate, reg_hours, normal_amt)
    if tue_hours > 0:
        add_breakdown_row("Tuesday Shift", rate, tue_hours, tue_amt)
    if sun_hours > 0:
        add_breakdown_row("SUNDAY TIME @ x 1,5", sunday_rate, sun_hours, sunday_amt)
    if hol_hours > 0:
        add_breakdown_row("PUBLIC HOLIDAYS @ x 2", hol_rate, hol_hours, hol_amt)
    if leave_days > 0:
        add_breakdown_row("ANNUAL LEAVE", rate, leave_hrs, leave_amt)
    if sick_days > 0:
        add_breakdown_row("SICK LEAVE", rate, sick_hrs, sick_amt)
    if boh_share > 0:
        add_breakdown_row("BOH Tips Pool Share", None, None, boh_share)
    bonus = float(emp_data.get('bonus', 0.0))
    if bonus > 0:
        add_breakdown_row("Bonus", None, None, bonus)
    gross = normal_amt + tue_amt + sunday_amt + hol_amt + leave_amt + sick_amt + bonus + boh_share
    
    pdf.set_font("Helvetica", 'B', 10)
    pdf.cell(150, 8, "TOTAL GROSS", border=1)
    pdf.cell(40, 8, f"R {gross:.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align='R')
    pdf.ln(5)

    # Deductions Table
    pdf.cell(190, 8, "DEDUCTIONS:", border=1, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", '', 10)
    
    uif = float(emp_data.get('uif', gross * 0.01)) # 1% default
    clothing = float(emp_data.get('clothing', 0.0))
    
    def add_deduction_row(desc, amt):
        if amt > 0:
            pdf.cell(150, 8, desc, border=1)
            pdf.cell(40, 8, f"R {amt:.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align='R')

    add_deduction_row("Tips Shortfall (Waiter Pool)", shortfall)
    add_deduction_row("Till Short", till_short)
    add_deduction_row("UIF", uif)
    add_deduction_row("Clothing", clothing)

    total_deductions = shortfall + till_short + uif + clothing
    
    pdf.set_font("Helvetica", 'B', 10)
    pdf.cell(150, 8, "TOTAL DEDUCTIONS", border=1)
    pdf.cell(40, 8, f"R {total_deductions:.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align='R')
    
    pdf.cell(150, 8, "Gross less Deductions Total", border=1)
    pdf.cell(40, 8, f"R {gross - total_deductions:.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align='R')
    
    pdf.set_font("Helvetica", 'B', 12)
    pdf.cell(150, 10, "TOTAL NETT SALARY", border=1)
    pdf.cell(40, 10, f"R {gross - total_deductions:.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align='R')

    import re
    # Scrub name for filename safety (e.g. remove slashes)
    safe_name = "".join(x for x in emp_data['name'] if x.isalnum() or x in "._- ")
    safe_name = safe_name.replace(' ', '_')
    
    filename = os.path.join(output_dir, f"Payslip_{safe_name}_{datetime.now().strftime('%Y%m%d')}.pdf")
    pdf.output(filename)
    return filename

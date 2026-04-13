# Wimpy Staff Management Portal

A powerful web-based automation portal for managing staff rosters, payroll, and back-of-house (BOH) operations. This tool automates the process of matching Ankerdata clock-ins with Excel-based rosters and generating payslips.

## Features
- **AI Roster Generation**: Uses Gemini AI to intelligently create weekly schedules based on staff roles and availability.
- **Payroll Automation**: Calculates regular, Sunday, and Tuesday hours automatically.
- **Break Rules**: Mandatory 30-minute unpaid break deduction for shifts over 5 hours.
- **Overtime & Off-Day Approval**: Manager dashboard for reviewing and approving unscheduled work.
- **Glassmorphism UI**: Modern, responsive dashboard with role-based staff grouping.

## Prerequisites
- Python 3.10+
- A Google Gemini API Key

## Setup Instructions

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/DylanLloydXero/Wimpy_Automation.git
   cd Wimpy_Automation
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure API Keys**:
   - Rename `config.json.example` to `config.json`.
   - Open `config.json` and enter your `GEMINI_API_KEY` and (optional) `NGROK_AUTHTOKEN`.

4. **Prepare Data (Templates)**:
   - Ensure `data/templates/Staff_Details_Template.xlsx` is populated with your staff list.
   - (A mock version is provided for testing).

5. **Run the App**:
   - On Windows: Double-click `start_app.bat`.
   - Manually: `python -m app.main_app`.

## Usage
- **Roster Tab**: Select a date and use "AI Auto-Generate" to build a schedule.
- **Payroll Tab**: Upload your latest roster and Ankerdata clock-in file to calculate wages and generate PDF payslips.
- **My Team**: View staff grouped by role (Barista, Griller, Kitchen, etc.).

## Security Note
This repository contains MOCK data in the `data/` folders for demonstration purposes. Never push your real `config.json` or personnel records to a public repository.

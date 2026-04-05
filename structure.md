# Project Structure

The Wimpy Automation project has been reorganized from a flat directory into a modern, hierarchical structure. This improves maintainability and makes it easier to find relevant files.

## Folder Map

```text
Wimpy_Automation/
├── app/                  # Core Application Logic
│   ├── api.py            # Main FastAPI Backend
│   ├── main_app.py        # Desktop UI (PyQT)
│   ├── ankerdata_bot.py   # Ankerdata Scraper
│   ├── wa_bot.py          # WhatsApp Automation
│   ├── payroll_processor.py
│   └── payslip_generator.py
├── data/                 # Application Data & Assets
│   ├── input/            # Roster and Clock-in Excel files
│   ├── templates/        # Staff Details & Brand Logo
│   ├── profiles/         # Selenium browser profiles
│   └── archives/         # Past Rosters and Payslips
├── output/               # Generated Files
│   ├── payslips/         # Current batch of PDF payslips
│   ├── summaries/        # EFT Excel sheets and Merge PDFs
│   └── whatsapp_payload.json
├── static/               # Web Application Frontend (HTML/JS/CSS)
├── tests/                # Testing & Mock Data
│   ├── test_docs/        # Mock Excel files for testing
│   └── run_tests.py      # Test execution script
└── utils/                # Maintenance & Utility Scripts
    └── download_logo.py  # Script to fetch Wimpy assets
```

## Key Changes
- **Running the App**: To start the application, continue using `start_app.bat` or `start_server.bat` from the root directory. They have been updated to point to the new internal paths.
- **Unified Data Storage**: All Excel inputs and the brand logo are now under `data/`, separated into logical subfolders.
- **Clean Output**: Generated payslips and summaries no longer clutter the root; they are neatly organized in the `output/` folder.
- **Modular Imports**: Code has been updated to use relative imports, allowing the `app/` folder to act as a proper Python package.

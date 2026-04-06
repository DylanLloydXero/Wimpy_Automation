@echo off
echo Starting Wimpy Payroll Management Server...
start "" "http://localhost:8000"
python -m uvicorn app.api:app --host 127.0.0.1 --port 8000 --reload
pause

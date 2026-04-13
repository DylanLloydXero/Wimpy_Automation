@echo off
echo Starting Wimpy Payroll Management Server...
python -m uvicorn app.api:app --host 0.0.0.0 --port 8000 --reload
pause

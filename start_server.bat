@echo off
echo Starting Wimpy Payroll Management Server...
python -m uvicorn api:app --host 127.0.0.1 --port 8000 --reload
pause

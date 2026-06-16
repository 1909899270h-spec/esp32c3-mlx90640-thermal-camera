@echo off
setlocal
cd /d "%~dp0"
python tools\mock_thermal_server.py --host 0.0.0.0 --port 8000
pause

@echo off
setlocal
cd /d "%~dp0"
python tools\mock_thermal_server.py --host 127.0.0.1 --port 8000
pause

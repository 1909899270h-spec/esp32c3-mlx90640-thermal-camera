@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*mock_thermal_server.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo Mock thermal server stopped.
pause

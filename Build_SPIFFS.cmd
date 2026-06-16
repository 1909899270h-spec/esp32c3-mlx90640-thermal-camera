@echo off
setlocal
cd /d "%~dp0"
pio run -e esp32-c3-devkitm-1 -t buildfs
pause

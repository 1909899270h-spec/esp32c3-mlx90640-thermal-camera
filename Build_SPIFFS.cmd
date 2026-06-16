@echo off
setlocal

set "PROJECT_DIR=%~dp0"
set "PATH=%PROJECT_DIR%.platformio-venv\Scripts;%PATH%"
set "PLATFORMIO_CORE_DIR=%PROJECT_DIR%.platformio-core"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

set "PYTHON=%PROJECT_DIR%.platformio-venv\Scripts\python.exe"
set "SPIFFSGEN=%PROJECT_DIR%.platformio-core\packages\framework-espidf\components\spiffs\spiffsgen.py"
set "OUT=%PROJECT_DIR%.pio\build\esp32-c3-devkitm-1\spiffs.bin"

if not exist "%PYTHON%" (
    echo Python venv not found: %PYTHON%
    exit /b 1
)

if not exist "%SPIFFSGEN%" (
    echo spiffsgen.py not found: %SPIFFSGEN%
    exit /b 1
)

if not exist "%PROJECT_DIR%.pio\build\esp32-c3-devkitm-1" (
    mkdir "%PROJECT_DIR%.pio\build\esp32-c3-devkitm-1"
)

"%PYTHON%" "%SPIFFSGEN%" 0x100000 "%PROJECT_DIR%data" "%OUT%" --page-size 256 --block-size 4096 --obj-name-len 32 --meta-len 4 --use-magic --use-magic-len
if errorlevel 1 exit /b %errorlevel%

for %%I in ("%OUT%") do echo Created %%~fI (%%~zI bytes)

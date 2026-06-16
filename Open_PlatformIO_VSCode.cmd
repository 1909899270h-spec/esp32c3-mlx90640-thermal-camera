@echo off
set "PROJECT_DIR=%~dp0"
set "PATH=%PROJECT_DIR%.platformio-venv\Scripts;%PATH%"
set "PLATFORMIO_CORE_DIR=%PROJECT_DIR%.platformio-core"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
code --extensions-dir "%PROJECT_DIR%.vscode-extensions" --user-data-dir "%PROJECT_DIR%.vscode-user" "%PROJECT_DIR%Thermal_Imaging.code-workspace"

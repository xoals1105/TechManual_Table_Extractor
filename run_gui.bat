@echo off
REM Launch GUI (double-click)
cd /d "%~dp0"

if not exist .venv\Scripts\pythonw.exe (
    echo [ERROR] venv not found. Run install_offline.bat first.
    pause
    exit /b 1
)

start "" /D "%~dp0" "%~dp0.venv\Scripts\pythonw.exe" -m src.gui

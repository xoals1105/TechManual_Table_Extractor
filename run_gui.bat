@echo off
REM Launch GUI (double-click)
cd /d "%~dp0"

if not exist .venv\Scripts\pythonw.exe (
    echo [ERROR] venv not found. Run install_offline.bat first.
    pause
    exit /b 1
)

start "" .venv\Scripts\pythonw.exe -m src.gui

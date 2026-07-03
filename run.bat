@echo off
REM Usage: run.bat INPUT_HWP_OR_FOLDER [options]
REM   ex : run.bat data\manual.hwpx
REM        run.bat data\ -o output
REM        run.bat data\manual.hwpx --all
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
    echo [ERROR] venv not found. Run install_offline.bat first.
    pause
    exit /b 1
)

if "%~1"=="" (
    echo Usage: run.bat INPUT_HWP_OR_FOLDER [options]
    pause
    exit /b 1
)

.venv\Scripts\python -m src.main %*

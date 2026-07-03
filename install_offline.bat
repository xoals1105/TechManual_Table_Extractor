@echo off
REM ============================================================
REM  Offline install script (no internet needed)
REM  - Installs dependencies from the wheelhouse folder only.
REM  - Requires Python 3.13 (64bit) to be installed first.
REM ============================================================
cd /d "%~dp0"

py -3.13 --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python 3.13 not found. Install Python 3.13 first.
    pause
    exit /b 1
)

echo [1/3] Creating venv (.venv)...
if not exist .venv (
    py -3.13 -m venv .venv
)

echo [2/3] Installing from wheelhouse (offline)...
.venv\Scripts\python -m pip install --no-index --find-links wheelhouse -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Install failed. Check that the wheelhouse folder exists.
    pause
    exit /b 1
)

echo [3/3] Verifying...
.venv\Scripts\python -c "import openpyxl, lxml, yaml, PyQt6; print('openpyxl', openpyxl.__version__, '/ lxml OK / PyYAML OK / PyQt6 OK')"

echo.
echo Done.
echo   - GUI : double-click run_gui.bat
echo   - CLI : run.bat INPUT_FILE_OR_FOLDER
pause

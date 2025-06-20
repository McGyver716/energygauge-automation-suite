@echo off
echo ==============================================
echo EnergyGauge USA Automation Suite
echo Modern GUI with Schematic Upload & Approval
echo ==============================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

REM Change to the script directory
cd /d "%~dp0"

REM Create virtual environment if it doesn't exist
if not exist "energy_gauge_env" (
    echo Creating virtual environment...
    python -m venv energy_gauge_env
)

REM Activate virtual environment
echo Activating virtual environment...
call energy_gauge_env\Scripts\activate.bat

REM Install/upgrade required packages
echo Installing required packages...
pip install --upgrade pip
pip install pywin32 opencv-python pytesseract easyocr watchdog pandas pillow

REM Create sample data if needed
echo Creating sample data...
python -c "
import os
from pathlib import Path
if not (Path('inputs').exists() and any(Path('inputs').glob('*.json'))):
    exec(open('code/energy_gauge_automation.py').read())
    create_sample_input()
    print('Sample data created!')
else:
    print('Sample data already exists.')
"

echo.
echo ==============================================
echo Starting EnergyGauge Automation Suite...
echo ==============================================
echo.
echo GUI Features:
echo - Upload schematics and JSON data files
echo - Configure project settings
echo - Approval workflow with preview
echo - COM interface discovery
echo - Results tracking and reporting
echo.

REM Launch the modern GUI
python code/energy_gauge_automation.py --mode modern

echo.
echo ==============================================
echo Session ended. Press any key to close...
pause >nul
@echo off
REM EnergyGauge USA Automation Setup Script for Windows
REM Run this script to set up the automation environment

echo Setting up EnergyGauge USA Automation...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or later from https://python.org
    pause
    exit /b 1
)

echo Python found. Creating virtual environment...

REM Create virtual environment
python -m venv energy_gauge_env
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

echo Activating virtual environment...
call energy_gauge_env\Scripts\activate

echo Installing required packages...
pip install -r requirements.txt
if errorlevel 1 (
    echo WARNING: Some packages may have failed to install
    echo This is normal on non-Windows systems
)

echo.
echo Setup complete!
echo.
echo To run the automation:
echo 1. Activate the virtual environment: energy_gauge_env\Scripts\activate
echo 2. Run the script: python code\energy_gauge_automation.py --mode gui
echo.
echo Make sure you have:
echo - EnergyGauge USA installed and licensed
echo - Tesseract OCR installed (https://github.com/UB-Mannheim/tesseract/wiki)
echo - Template file in templates\YourTemplate.egpj
echo - Input JSON files in inputs\ directory
echo.
pause
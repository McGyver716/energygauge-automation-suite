@echo off
REM Quick launcher for EnergyGauge USA Automation

echo Starting EnergyGauge USA Automation...
echo.

REM Check if virtual environment exists
if not exist "energy_gauge_env\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found
    echo Please run setup.bat first
    pause
    exit /b 1
)

REM Activate virtual environment
call energy_gauge_env\Scripts\activate

REM Run the automation in GUI mode
python code\energy_gauge_automation.py --mode gui

echo.
echo Automation finished.
pause
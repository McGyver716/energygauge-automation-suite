@echo off
echo ==============================================
echo EnergyGauge USA COM Interface Discovery
echo ==============================================
echo.
echo This script will discover the available COM interface
echo properties and methods in your EnergyGauge USA installation.
echo.
echo Prerequisites:
echo - EnergyGauge USA must be installed and licensed
echo - Run this script as Administrator for best results
echo.
pause

REM Change to the script directory
cd /d "%~dp0"

REM Activate virtual environment if it exists
if exist "energy_gauge_env\Scripts\activate.bat" (
    echo Activating virtual environment...
    call energy_gauge_env\Scripts\activate.bat
)

echo.
echo ==============================================
echo Connecting to EnergyGauge USA...
echo ==============================================
echo.

REM Run the discovery
python code/energy_gauge_automation.py --discover

echo.
echo ==============================================
echo Discovery complete!
echo ==============================================
echo.
echo Use the discovered properties and methods to customize
echo the COM interface integration in the source code by
echo replacing TODO comments with actual calls.
echo.
pause
@echo off
echo ========================================
echo  EnergyGauge Automation Suite
echo  GitHub Upload Script
echo ========================================
echo.

echo Checking Git status...
git status

echo.
echo Pushing to GitHub repository...
echo Repository: https://github.com/McGyver716/energygauge-automation-suite
echo.

git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo  SUCCESS! Repository uploaded to GitHub
    echo ========================================
    echo.
    echo Your complete EnergyGauge Automation Suite
    echo is now available at:
    echo https://github.com/McGyver716/energygauge-automation-suite
    echo.
) else (
    echo.
    echo ========================================
    echo  Push failed - Authentication needed
    echo ========================================
    echo.
    echo Please configure Git credentials:
    echo   git config --global user.name "Your Name"
    echo   git config --global user.email "your.email@example.com"
    echo.
    echo Then run this script again.
    echo.
)

pause
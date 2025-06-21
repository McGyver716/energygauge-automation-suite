@echo off
echo =========================================================
echo  EnergyGauge Automation Suite - GitHub Upload
echo  Email: aaronlbrown716@gmail.com
echo  GitHub: McGyver716
echo =========================================================
echo.

echo Step 1: Configuring Git with your email...
git config user.email "aaronlbrown716@gmail.com"
git config user.name "McGyver716"
echo ✅ Git configured

echo.
echo Step 2: Checking repository status...
git status
echo.

echo Step 3: Pushing to GitHub...
echo Repository: https://github.com/McGyver716/energygauge-automation-suite
echo.
echo When prompted for credentials:
echo   Username: McGyver716
echo   Password: [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]
echo.
echo ⚠️  IMPORTANT: You need a Personal Access Token from GitHub.com
echo    Go to: GitHub.com → Settings → Developer settings → Personal access tokens
echo    Generate new token with 'repo' permissions
echo.

pause
echo Starting upload...
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo =========================================================
    echo  🎉 SUCCESS! Your automation suite is now on GitHub!
    echo =========================================================
    echo.
    echo View your repository at:
    echo https://github.com/McGyver716/energygauge-automation-suite
    echo.
    echo What was uploaded:
    echo ✅ Complete automation engine (2,226+ lines)
    echo ✅ Modern GUI with tabbed interface
    echo ✅ OCR processing system (Tesseract + EasyOCR)
    echo ✅ COM interface automation
    echo ✅ Windows deployment scripts
    echo ✅ Comprehensive documentation
    echo ✅ Sample data and templates
    echo.
) else (
    echo.
    echo =========================================================
    echo  ❌ Upload failed - Need Personal Access Token
    echo =========================================================
    echo.
    echo Quick steps to get your token:
    echo 1. Go to GitHub.com and sign in with aaronlbrown716@gmail.com
    echo 2. Profile picture → Settings → Developer settings
    echo 3. Personal access tokens → Generate new token
    echo 4. Name: EnergyGauge Automation
    echo 5. Check 'repo' permissions
    echo 6. Copy the token and use it as your password
    echo.
    echo Then run this script again!
    echo.
)

pause

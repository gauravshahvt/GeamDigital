@echo off
title Geam Digital - District Voter Studio Desktop App
echo ===============================================================================
echo        GEAM DIGITAL - DISTRICT AUTOMATED VOTER EXCEL STUDIO (DESKTOP APP)
echo        Full District: Samiti -^> Panchayat -^> All Wards Auto Excel GUI
echo ===============================================================================
echo.
echo [1/2] Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not found on your system PATH!
    pause
    exit /b 1
)

echo [2/2] Checking dependencies...
python -c "import customtkinter, openpyxl, fitz, pandas" >nul 2>&1
if errorlevel 1 (
    echo Installing missing dependencies...
    pip install -r requirements.txt
)

echo.
echo Starting Desktop Application...
echo.

python district_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application exited with an error code.
    pause
)

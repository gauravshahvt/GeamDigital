@echo off
title Geam Digital - Voter List to Excel Generator
echo =========================================================
echo       GEAM DIGITAL - VOTER LIST TO EXCEL CONVERTER
echo =========================================================
echo.
echo [1/3] Python environment check...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not found on your system PATH!
    pause
    exit /b 1
)

echo [2/3] Checking dependencies...
python -c "import flask, pymupdf, openpyxl, pandas, pytesseract" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements.txt
)

echo [3/3] Starting Geam Digital Web App...
start "" http://127.0.0.1:5000
echo.
echo =========================================================
echo   Geam Digital is running at: http://127.0.0.1:5000
echo   Open your browser to upload Voter List PDFs.
echo =========================================================
echo.
python app.py
pause

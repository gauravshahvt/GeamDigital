@echo off
title Geam Digital - Color Voter Slip Studio Pro Desktop App
echo ===============================================================================
echo        GEAM DIGITAL - COLOR VOTER SLIP & POSTER STUDIO (DESKTOP APP 2026)
echo        A4 Color Voter Slips (8 / 12 per page) with Candidate Poster & Symbol
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
python -c "import customtkinter, openpyxl, pymupdf" >nul 2>&1
if errorlevel 1 (
    echo Installing missing dependencies...
    pip install customtkinter openpyxl pymupdf
)

echo.
echo Starting Color Voter Slip Studio Desktop App...
echo.

python color_parchi_desktop.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application exited with an error code.
    pause
)

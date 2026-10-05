@echo off
title Geam Digital - District Automated Voter Excel Studio (2026)
echo ===============================================================================
echo        GEAM DIGITAL - DISTRICT AUTOMATED VOTER EXCEL STUDIO (CLI)
echo        Full District: Samiti -^> Panchayat -^> All Wards Auto Excel
echo ===============================================================================
echo.

python automate_district.py

if errorlevel 1 (
    echo.
    echo [ERROR] Automation exited with an error code.
    pause
)

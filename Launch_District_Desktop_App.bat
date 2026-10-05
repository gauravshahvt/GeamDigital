@echo off
chcp 65001 > nul
title Geam Digital - District Voter Studio Desktop App
cls
echo ===============================================================================
echo        🏛️ GEAM DIGITAL - DISTRICT AUTOMATED VOTER EXCEL STUDIO (DESKTOP APP)
echo        सम्पूर्ण ज़िला: पंचायत समिति ➔ ग्राम पंचायत ➔ सभी वार्ड्स एक्सेल ऑटोमेशन GUI
echo ===============================================================================
echo.
echo ⏳ Starting District Desktop Application...
echo.

python district_app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Application exited with an error. Please see details above.
    pause
)

@echo off
chcp 65001 > nul
title Geam Digital - District Automated Voter Excel Studio (2026)
cls
echo ===============================================================================
echo        🏛️ GEAM DIGITAL - DISTRICT AUTOMATED VOTER EXCEL STUDIO (2026)
echo        सम्पूर्ण ज़िला: पंचायत समिति ➔ ग्राम पंचायत ➔ सभी वार्ड्स एक्सेल ऑटोमेशन
echo ===============================================================================
echo.

python automate_district.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ निष्पादन में त्रुटि हुई। कृपया ऊपर दिए गए संदेश को पढ़ें।
    pause
)

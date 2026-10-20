@echo off
title Start Geam Digital Server (Permanent Background)
cd /d "%~dp0"
echo =========================================================
echo       STARTING GEAM DIGITAL PERMANENT SERVER
echo =========================================================
echo.

echo [1/3] Starting server watchdog in background...
cscript //nologo start_server_hidden.vbs

echo [2/3] Waiting for server on http://127.0.0.1:5000...
powershell -Command "$ready = $false; for ($i=0; $i -lt 15; $i++) { if (Get-NetTCPConnection -LocalPort 5000 -ErrorAction SilentlyContinue) { $ready = $true; break }; Start-Sleep -Milliseconds 500 }; if ($ready) { Write-Host 'SUCCESS: Local server is ACTIVE on http://127.0.0.1:5000' -ForegroundColor Green } else { Write-Host 'Server starting... please check in browser.' -ForegroundColor Yellow }"

echo [3/3] Opening browser...
start "" "http://127.0.0.1:5000"

echo.
echo =========================================================
echo Server is running silently in the background with Auto-Restart!
echo It will NOT close even if this window is closed.
echo =========================================================
timeout /t 3 >nul

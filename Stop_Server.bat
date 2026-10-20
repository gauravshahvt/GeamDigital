@echo off
title Stop Geam Digital Server
echo =========================================================
echo       STOPPING GEAM DIGITAL LOCAL SERVER
echo =========================================================
echo.

echo [1/2] Terminating watchdog and python processes on port 5000 and 5005...
powershell -Command "$conns = Get-NetTCPConnection -LocalPort 5000, 5005 -ErrorAction SilentlyContinue; if ($conns) { foreach ($c in $conns) { Stop-Process -Id $c.OwningProcess -Force -ErrorAction SilentlyContinue } }"

echo [2/2] Checking if server stopped...
timeout /t 1 /nobreak >nul
powershell -Command "if (Get-NetTCPConnection -LocalPort 5000 -ErrorAction SilentlyContinue) { Write-Host 'Warning: Server still listening.' -ForegroundColor Red } else { Write-Host 'SUCCESS: Geam Digital Local Server stopped completely.' -ForegroundColor Green }"

echo.
pause

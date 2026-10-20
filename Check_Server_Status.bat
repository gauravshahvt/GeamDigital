@echo off
title Geam Digital Server Status
echo =========================================================
echo       GEAM DIGITAL LOCAL SERVER STATUS CHECK
echo =========================================================
echo.

powershell -Command "
$conn = Get-NetTCPConnection -LocalPort 5000 -ErrorAction SilentlyContinue
if ($conn) {
    Write-Host 'STATUS: [ACTIVE] Server is RUNNING on http://127.0.0.1:5000' -ForegroundColor Green
    Write-Host ('Process ID: ' + $conn.OwningProcess) -ForegroundColor Cyan
} else {
    Write-Host 'STATUS: [INACTIVE] Server is NOT running on port 5000.' -ForegroundColor Red
}
"

echo.
pause

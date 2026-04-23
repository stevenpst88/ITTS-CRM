@echo off
chcp 65001 >nul
echo.
echo  ========================================
echo   名片 CRM 系統啟動中...
echo  ========================================
echo.
cd /d "%~dp0"
start "" "http://localhost:3000"
node server.js
pause

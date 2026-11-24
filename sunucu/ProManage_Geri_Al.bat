@echo off
REM ============================================================
REM   ProManage_Kapat.bat'in yaptigini GERI ALIR: servislerin eski
REM   baslangic turu geri konur, calisanlar yeniden baslatilir.
REM   Yedek: C:\cofle\logs\promanage_yedek.json
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ProManage_Kapat.ps1" -GeriAl
echo.
pause

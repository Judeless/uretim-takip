@echo off
REM ============================================================
REM   Cofle sunucu ARIZA TESHISI - kilitlenme / reset sonrasi calistir.
REM   Yalniz okur. Sonuc: C:\cofle\logs\ariza_<tarih>\OZET.txt (Not Defteri'nde acilir)
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Ariza_Teshis.ps1"
echo.
echo OZET.txt Not Defteri'nde acildi - icerigini gonderin.
pause

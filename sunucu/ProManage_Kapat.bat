@echo off
REM ============================================================
REM   ProManage + Oracle servislerini DURDUR ve DEVRE DISI birak.
REM   Once liste gosterilir; KAPAT yazmadan hicbir sey degismez.
REM   Dosya / veritabani silinmez. Geri almak: ProManage_Geri_Al.bat
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ProManage_Kapat.ps1" -Kapat
echo.
pause

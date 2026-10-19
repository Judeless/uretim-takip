@echo off
REM Aspanda Flow - sunucu kurulumu. Yonetici olarak kendini acar, Aspanda_Kur.ps1'i calistirir.
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Aspanda_Kur.ps1"

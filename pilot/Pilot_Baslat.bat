@echo off
title Cofle Pilot Sayac - Backend
color 0D

echo ====================================================
echo        COFLE PILOT SAYAC BACKEND BASLATILIYOR...
echo ====================================================
echo.

echo [1/2] Port 5001 kontrol ediliyor...
:: Portta kalan eski surecleri temizle
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5001 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1

echo [2/2] Pilot sunucu (pilot_app.py) aciliyor...
cd /d "%~dp0"
start "Cofle Pilot" cmd /k "python pilot_app.py"

echo.
timeout /t 3 >nul

echo ====================================================
echo  PILOT YAYINDA — port 5001
echo.
echo  Canli UI    : http://localhost:5001/
echo  ESP32 URL   : http://%COMPUTERNAME%:5001/api/sinyal
echo  API Token   : cofle-pilot-2026
echo.
echo  Ana Cofle Manage sistemi (port 5000) etkilenmedi.
echo ====================================================
timeout /t 8

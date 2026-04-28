@echo off
title Cofle Manage - Otomatik Baslatici
color 0B

echo ====================================================
echo        COFLE MANAGE SISTEMI BASLATILIYOR...
echo ====================================================
echo.

echo [1/2] Port 5000 kontrol ediliyor...
:: Portta kalan eski surecleri temizle
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5000 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1

echo [2/2] Sunucu (app.py) aciliyor...
start "Cofle Manage Server" cmd /k "python app.py"

echo.
timeout /t 5 >nul

echo ====================================================
echo ISLEM TAMAMLANDI! SISTEM YAYINDA.
echo.
echo Adres: https://coflemanage.online
echo.
echo ====================================================
timeout /t 10

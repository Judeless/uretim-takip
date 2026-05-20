@echo off
title Cofle Manage - Otomatik Baslatici
color 0B

echo ====================================================
echo        COFLE MANAGE SISTEMI BASLATILIYOR...
echo ====================================================
echo.

echo [1/3] Port 5000 kontrol ediliyor...
:: Portta kalan eski surecleri temizle
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5000 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1

echo [2/3] Ana sunucu (app.py) aciliyor...
start "Cofle Manage Server" cmd /k "python app.py"

echo.
timeout /t 3 >nul

echo [3/3] Pilot sayac sunucu (port 5001) aciliyor...
:: Pilot_Baslat.bat kendi icinde port 5001 temizligi + pilot_app.py baslatma yapar
start "" "%~dp0pilot\Pilot_Baslat.bat"

echo.
timeout /t 3 >nul

echo ====================================================
echo ISLEM TAMAMLANDI! SISTEM YAYINDA.
echo.
echo Ana sistem  : https://coflemanage.online      (port 5000)
echo Pilot sayac : http://localhost:5001/          (port 5001)
echo.
echo ====================================================
timeout /t 10

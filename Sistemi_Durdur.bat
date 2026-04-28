@echo off
title Cofle Manage - Sistemi Durdur
color 0C

echo ====================================================
echo        COFLE MANAGE SISTEMI DURDURULUYOR...
echo ====================================================

echo [1/2] Python islemleri sonlandiriliyor...
taskkill /F /FI "WINDOWTITLE eq Cofle Manage Server*" /T >nul 2>&1
taskkill /F /IM python.exe /T >nul 2>&1

echo [2/2] Port 5000 temizleniyor...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5000 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1

echo.
echo ====================================================
echo SISTEM DURDURULDU!
echo ====================================================
timeout /t 3

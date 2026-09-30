@echo off
chcp 65001 >nul
title Cofle Forge - Sunucudan Veri Cek
cd /d "%~dp0"
echo.
echo  ================================================================
echo   CANLI SUNUCUNUN VERISI BU MAKINEYE INDIRILECEK (tek yon)
echo   Sunucuya HICBIR SEY yazilmaz. Yerel veriler yedeklenir.
echo  ================================================================
echo.
python Sunucudan_Veri_Cek.py %*
echo.
pause

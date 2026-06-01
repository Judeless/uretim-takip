@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title COFLE OTA Yukleyici

REM ============================================================
REM   COFLE PILOT - OTA Firmware Yukleyici (espota.py)
REM ============================================================
REM Arduino IDE 2.x'te su hata cikiyorsa bu script onu atlatir:
REM   espota.exe: error: argument -p/--port: invalid int value:
REM   '{upload.port.properties.port}'
REM
REM KULLANIM (tek seferlik hazirlik):
REM   1. Arduino IDE'de cofle_sayac_<cihaz>.ino dosyasini ac
REM   2. Board: ESP32-WROOM-32U sec (Tools > Board)
REM   3. Sketch > Export Compiled Binary  (Ctrl+Alt+S)
REM      -> sketch klasorunde build\... altinda .ino.bin olusur
REM   4. Bu .bat dosyasini cift tikla, cihaz adi + IP gir
REM
REM Network port secmeye, IDE'den upload etmeye GEREK YOK.
REM ============================================================

echo ============================================================
echo   COFLE PILOT - OTA Firmware Yukleyici
echo ============================================================
echo.

REM --- espota.py bul (ESP32 core icinde) ---
set "ESPOTA="
for /f "delims=" %%F in ('dir /b /s "%LOCALAPPDATA%\Arduino15\packages\esp32\hardware\esp32\espota.py" 2^>nul') do set "ESPOTA=%%F"
if "!ESPOTA!"=="" (
    echo [HATA] espota.py bulunamadi.
    echo        ESP32 board paketi kurulu mu? Arduino IDE ^> Boards Manager ^> esp32
    pause & exit /b 1
)
echo [OK] espota.py bulundu
echo.

REM --- Cihaz ve IP sor ---
set /p CIHAZ="Cihaz adi (ornek: abb5 / m1 / 400t): "
if "!CIHAZ!"=="" ( echo [HATA] Cihaz adi bos olamaz & pause & exit /b 1 )
set /p IP="Cihaz IP adresi (ornek: 192.168.21.140): "
if "!IP!"=="" ( echo [HATA] IP bos olamaz & pause & exit /b 1 )
set "PAROLA=cofle-ota-2026"

set "SKETCHDIR=%~dp0cofle_sayac_!CIHAZ!"
if not exist "!SKETCHDIR!" (
    echo [HATA] Klasor yok: !SKETCHDIR!
    echo        Cihaz adini kontrol et ^(abb5, m1, 400t gibi^)
    pause & exit /b 1
)

REM --- En yeni .ino.bin bul (bootloader/partitions haric) ---
set "BIN="
for /f "delims=" %%F in ('dir /b /s /o-d "!SKETCHDIR!\*.ino.bin" 2^>nul ^| findstr /v /i "bootloader partitions"') do (
    if "!BIN!"=="" set "BIN=%%F"
)
if "!BIN!"=="" (
    echo [HATA] Derlenmis .bin bulunamadi.
    echo        Arduino IDE'de once: Sketch ^> Export Compiled Binary
    pause & exit /b 1
)
echo [OK] Firmware: !BIN!
echo.
echo ------------------------------------------------------------
echo   Hedef : !CIHAZ! @ !IP! : 3232
echo   Yukleniyor...
echo ------------------------------------------------------------
python "!ESPOTA!" -i !IP! -p 3232 -a !PAROLA! -f "!BIN!"
set "RC=!errorlevel!"
echo.
if "!RC!"=="0" (
    echo [OK] Yukleme tamam! Cihaz reboot ediyor, yeni firmware aktif.
) else (
    echo [HATA] Yukleme basarisiz ^(kod !RC!^).
    echo        - Cihaz online mi? Pilot UI'da heartbeat geliyor mu?
    echo        - IP dogru mu? Dashboard Saha Cihazlari'nda kontrol et
    echo        - Ayni agda misin? ^(COFLE-TK^)
)
pause
endlocal

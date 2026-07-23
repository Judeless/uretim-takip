@echo off
chcp 65001 >nul
title Cofle Andon - Kaynakhane TV
REM ============================================================
REM  KAYNAKHANE ANDON - TV BASLATICI
REM  - Chrome'u TAM EKRAN (kiosk) acar. Adres cubugu/sekme yok.
REM  - Sayfa kendini ~30 dk'da bir yeniler (bellek hatasi/OOM onleme;
REM    andon_v5 sablonunun icinde). Tarayici kapatmaya gerek yok.
REM  - Chrome cokerse veya kapatilirsa 5 sn sonra OTOMATIK yeniden acar.
REM
REM  KULLANIM: Bu dosyayi masaustune kopyala, cift tikla.
REM  CIKIS:    Once bu siyah pencereyi kapat, sonra Chrome'da Alt+F4.
REM ============================================================

REM --- Andon adresi ---
REM Bu PC sunucunun KENDISI ise: http://localhost:5000/andon
REM Ayri bir TV/PC ise (ag uzerinden): http://192.168.21.155:5000/andon
REM Diger ekranlar:  Montaj=/andon_montaj  Metal=/andon_metal  TK1=/andon_tk1
set "ANDON_URL=http://192.168.21.155:5000/andon"

REM --- Ayri/temiz profil (normal taramayi etkilemez, her acilis temiz baslar) ---
set "PROFIL=%LOCALAPPDATA%\AndonKiosk"

REM --- Chrome'u bul ---
set "CHROME="
for %%P in (
  "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
  "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
  "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
) do if exist %%~P set "CHROME=%%~P"

if not defined CHROME (
  echo Google Chrome bulunamadi.
  echo Edge kullanmak icin asagidaki satirin basindaki REM'i kaldirin:
  REM set "CHROME=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
)
if not defined CHROME ( pause & exit /b 1 )

:dongu
start "" /wait "%CHROME%" --kiosk "%ANDON_URL%" ^
  --user-data-dir="%PROFIL%" ^
  --no-first-run --no-default-browser-check ^
  --disable-session-crashed-bubble --hide-crash-restore-bubble ^
  --disable-infobars --noerrdialogs ^
  --autoplay-policy=no-user-gesture-required
REM Chrome kapandi/coktu -> kisa bekle, tekrar ac (TV asla bos kalmasin)
timeout /t 5 /nobreak >nul
goto dongu

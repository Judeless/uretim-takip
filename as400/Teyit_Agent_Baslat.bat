@echo off
REM ═══════════════════════════════════════════════════════════════════
REM  Cofle Teyit Agent — PCOMM'un bulundugu (RDP/konsol) oturumda calisir
REM
REM  Baslangic klasorune KISAYOL koy (shell:startup) — logonda otomatik kalkar.
REM  Otomatik_Kalkis_Kur.bat bu kisayolu senin yerine olusturur.
REM
REM  KENDINI YENIDEN BASLATIR (2026-08-27): agent surecinin olmesi (PCOMM COM
REM  hatasi, elle kapatma, python cokmesi) tum otomasyonu sessizce durduruyordu
REM  — 16:45 transfer ve 17:10 teyit kosulari "teyit-agent kapali" diye atlaniyor,
REM  bunu ancak ertesi sabah fark ediyorduk. Artik surec olurse 10 sn sonra
REM  yeniden kalkar; her kalkis ekrana ve loga yazilir.
REM
REM  DURDURMAK ICIN: bu pencereyi kapat (Ctrl+C iki kez sorar).
REM ═══════════════════════════════════════════════════════════════════
title Cofle Teyit Agent
cd /d %~dp0
if not exist teyit_loglari mkdir teyit_loglari
set SAYAC=0

:dongu
set /a SAYAC+=1
echo.
echo ============================================================
echo  Teyit Agent baslatiliyor  (kalkis #%SAYAC%)  %DATE% %TIME%
echo ============================================================
echo [%DATE% %TIME%] agent kalkis #%SAYAC% >> teyit_loglari\agent_kalkis.log
python teyit_agent.py
echo.
echo  !!! Agent DURDU (cikis kodu %ERRORLEVEL%) — 10 sn sonra yeniden baslatiliyor.
echo  !!! Kapatmak istiyorsan bu pencereyi simdi kapat.
echo [%DATE% %TIME%] agent durdu, cikis=%ERRORLEVEL% >> teyit_loglari\agent_kalkis.log
timeout /t 10 /nobreak >nul
goto dongu

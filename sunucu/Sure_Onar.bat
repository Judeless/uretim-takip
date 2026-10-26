@echo off
REM ============================================================
REM   Sure onarimi (2026-10-26) - Excel otomatik doldurma hasari
REM   Once RAPOR gosterir, onay verirseniz uretim.db yedegini alip
REM   referans + gecmis uretim kaydi surelerini geri alir.
REM   ONCE Guncelle.bat calismis olmali (bu dosyalar oradan gelir).
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
cd /d C:\cofle\uretim_takip || (echo [HATA] C:\cofle\uretim_takip yok & pause & exit /b 1)
set "PY=C:\Program Files\Python313\python.exe"
set "VERI=sunucu\sure_onarim_20261026.json"
if not exist "%VERI%" (echo [HATA] %VERI% yok - once Guncelle.bat & pause & exit /b 1)

"%PY%" sunucu\sure_onar.py "%VERI%" --yalniz-referans
echo.
echo   924 = montaj 924 sn  ^|  120 = bos surelere yazilan 120 sn  ^|  test = pres test kodlari
echo   Uc grup da geri alinir; YALNIZ referans sureleri - gecmis uretim kayitlarina dokunulmaz.
echo.
set /p "CEVAP=Onarimi uygulamak icin E, vazgecmek icin H yazin: "
if /i "%CEVAP%"=="E" "%PY%" sunucu\sure_onar.py "%VERI%" --uygula --yalniz-referans
echo.
pause

@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM   Cofle Manage - Sunucu Guncelleme (VERI KORUMALI)
REM ============================================================
REM ESKI SURUM VERI KAYBETTIRIYORDU (2026-08-17):
REM   data\uretim_verileri.xlsx git tarafindan TAKIP EDILIYORDU (.gitignore'da
REM   data/ yazmasina ragmen — gitignore zaten takip edilen dosyaya islemez).
REM   'git reset --hard' bu dosyayi her guncellemede repo surumuyle EZIYORDU.
REM   Uygulama ise ayni dosyaya YAZIYOR (teyit PATCH auto-sync, "Excel'e Yaz")
REM   -> sunucuda Excel'e yazilan her sey bir sonraki guncellemede geri aliniyordu.
REM   Sessizdi; yalnizca dosya KILITLI oldugunda "Unlink of file failed" ile
REM   kendini gosterdi.
REM Bu surum: reset ONCESI data\ yedeklenir, SONRASI eksik dosyalar geri konur.
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)

set "APP=C:\cofle\uretim_takip"
set "GIT=C:\Program Files\Git\cmd\git.exe"
for /f "tokens=2 delims==" %%t in ('wmic os get localdatetime /value ^| find "="') do set "TS=%%t"
set "TS=!TS:~0,8!_!TS:~8,6!"
set "YEDEK=C:\cofle\yedek\data_!TS!"

cd /d "%APP%" || (echo [HATA] %APP% bulunamadi & pause & exit /b 1)

echo ============================================
echo   1/4  VERI YEDEKLENIYOR
echo ============================================
if exist "%APP%\data" (
    mkdir "!YEDEK!" 2>nul
    xcopy "%APP%\data" "!YEDEK!\" /E /I /Y /Q >nul
    if errorlevel 1 (
        echo [HATA] data\ yedeklenemedi - GUNCELLEME DURDURULDU.
        echo        Excel acik olabilir; kapatip tekrar deneyin.
        pause & exit /b 1
    )
    echo   [OK] data\ -^> !YEDEK!
) else (
    echo   [ATLA] data\ klasoru yok
)

echo.
echo ============================================
echo   2/4  GitHub'dan son kod cekiliyor
echo ============================================
"%GIT%" fetch origin main || (echo [HATA] fetch basarisiz & pause & exit /b 1)
REM --quiet degil: ne degistigi gorunsun. Kilitli dosya varsa git y/n sorar;
REM bu surumde data\ zaten takip disi oldugu icin o soru CIKMAMALI.
"%GIT%" reset --hard origin/main
if errorlevel 1 (
    echo [HATA] reset basarisiz. Yedek durUYOR: !YEDEK!
    pause & exit /b 1
)

echo.
echo ============================================
echo   3/4  VERI GERI KONUYOR (eksikse)
echo ============================================
set GERI=0
for %%F in ("!YEDEK!\*.xlsx") do (
    if not exist "%APP%\data\%%~nxF" (
        copy /Y "%%F" "%APP%\data\%%~nxF" >nul && (echo   [geri] %%~nxF & set /a GERI+=1)
    )
)
if "!GERI!"=="0" (echo   [OK] eksik dosya yok) else (echo   [OK] !GERI! dosya geri konuldu)

echo.
echo ============================================
echo   4/4  Servisler yeniden baslatiliyor
echo ============================================
C:\cofle\nssm.exe restart cofle-app
C:\cofle\nssm.exe restart cofle-pilot
echo.
echo   NOT: teyit-agent AYRI surectir, bu script onu yeniden BASLATMAZ.
echo        teyit_agent.py degistiyse agent penceresini kapatip
echo        Teyit_Agent_Baslat.bat ile tekrar acin.
echo.
echo   GUNCELLEME TAMAMLANDI.   Yedek: !YEDEK!
timeout /t 8
endlocal

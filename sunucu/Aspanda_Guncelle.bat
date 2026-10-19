@echo off
REM ============================================================
REM   Aspanda Flow - sunucuda kod guncelleme
REM   ONCE Cofle'nin Guncelle.bat'i calismis olmali: Aspanda kodu
REM   C:\cofle\uretim_takip'ten (yerel git) alinir. Veri (uretim.db,
REM   data\) git disidir, reset ona dokunmaz.
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
set "GIT=C:\Program Files\Git\cmd\git.exe"
cd /d C:\aspanda || (echo [HATA] C:\aspanda yok - once Aspanda_Kur.bat & pause & exit /b 1)
"%GIT%" -c safe.directory=* fetch origin || (echo [HATA] fetch basarisiz & pause & exit /b 1)
"%GIT%" -c safe.directory=* reset --hard origin/main || (echo [HATA] reset basarisiz & pause & exit /b 1)
if not exist "C:\aspanda\data\kurulum.json" (
    echo [HATA] data\kurulum.json YOK - servis yeniden baslatilmadi
    pause & exit /b 1
)
C:\cofle\nssm.exe restart aspanda-app
echo.
echo   ASPANDA GUNCELLENDI.
timeout /t 8

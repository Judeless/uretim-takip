@echo off
REM ============================================================
REM   Aspanda panel sifre sifirlama (2026-10-27)
REM   Yonetici sifresi unutulunca: yeni GECICI sifre uretir, yalniz
REM   bu ekrana bir kez yazar. Ilk giriste yeni sifre istenir.
REM   ONCE Guncelle.bat calismis olmali (bu dosya oradan gelir).
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
chcp 65001 >nul
if not exist "C:\aspanda\uretim.db" (echo [HATA] C:\aspanda\uretim.db yok & pause & exit /b 1)
"C:\Program Files\Python313\python.exe" "%~dp0sifre_sifirla.py" C:\aspanda
echo.
pause

@echo off
REM ============================================================
REM   Cofle sunucu saglik izlemeyi KURAR (bir kez calistir).
REM   Gorev Zamanlayici: "Cofle\Saglik_Izle" - SYSTEM, her 10 dakikada
REM   sunucu\Saglik_Izle.ps1 -> C:\cofle\logs\saglik\saglik_YYYY-MM.csv
REM   Kaldirmak icin: schtasks /delete /tn "Cofle\Saglik_Izle" /f
REM ============================================================
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)

set "PS1=C:\cofle\uretim_takip\sunucu\Saglik_Izle.ps1"
if not exist "%PS1%" (echo [HATA] %PS1% yok - once Guncelle.bat & pause & exit /b 1)

schtasks /create /tn "Cofle\Saglik_Izle" /sc minute /mo 10 /ru SYSTEM /rl HIGHEST /f ^
  /tr "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File %PS1%"
if errorlevel 1 (echo [HATA] gorev olusturulamadi & pause & exit /b 1)

schtasks /run /tn "Cofle\Saglik_Izle" >nul
echo.
echo [OK] Saglik izleme kuruldu - her 10 dakikada bir satir:
echo      C:\cofle\logs\saglik\saglik_YYYY-MM.csv
echo      Esik asilirsa: C:\cofle\logs\saglik\UYARI.txt
echo.
pause

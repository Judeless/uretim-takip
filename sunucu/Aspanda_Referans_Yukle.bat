@echo off
REM ============================================================
REM   Aspanda Flow - referans Excel'inden EKSIK referanslari ekler
REM   Kullanim: Excel dosyasini (.xlsb / .xlsx) bu dosyanin UZERINE
REM   surukleyip birakin. Var olan referanslara DOKUNMAZ; eski surumun
REM   200T sayfasindan metale yazdigi plastik urunleri plastige tasir.
REM   ONCE Aspanda_Guncelle.bat calismis olmali (yeni kod).
REM ============================================================
set "GECICI=%TEMP%\aspanda_referans"
if not "%~1"=="" (
    del /q "%GECICI%.xlsb" "%GECICI%.xlsx" 2>nul
    copy /Y "%~1" "%GECICI%%~x1" >nul || (echo [HATA] Excel kopyalanamadi & pause & exit /b 1)
)
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
set "EXCEL="
if exist "%GECICI%.xlsb" set "EXCEL=%GECICI%.xlsb"
if exist "%GECICI%.xlsx" set "EXCEL=%GECICI%.xlsx"
if not defined EXCEL (
    echo Excel dosyasini bu dosyanin UZERINE surukleyip birakin.
    pause & exit /b 1
)
cd /d C:\aspanda || (echo [HATA] C:\aspanda yok - once Aspanda_Kur.bat & pause & exit /b 1)
findstr /c:"sadece-referans" kurulum_hazirla.py >nul || (
    echo [HATA] C:\aspanda kodu eski - once Aspanda_Guncelle.bat calistirin.
    pause & exit /b 1
)
"C:\Program Files\Python313\python.exe" kurulum_hazirla.py --sadece-referans --excel "%EXCEL%"
del /q "%EXCEL%" 2>nul
echo.
pause

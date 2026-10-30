@echo off
REM ============================================================
REM   DENETIM (2026-10-30) - acilan / kapatilan launch listesi
REM   Ic kontrol talebi madde 1: 15-19.06.2026 ve 07-11.09.2026
REM   SALT OKUNUR: AS400'e yalniz SELECT gider, hicbir sey yazilmaz.
REM   promanage (RDP) oturumunda calistirin - AS400 sifresi kasadan alinir.
REM   ONCE Guncelle.bat calismis olmali (betik oradan gelir).
REM   Baska donem:  Launch_Listesi.bat 2026-06-15:2026-06-19 2026-09-07:2026-09-11
REM ============================================================
cd /d C:\cofle\uretim_takip || (echo [HATA] C:\cofle\uretim_takip yok & pause & exit /b 1)
set "PY=C:\Program Files\Python313\python.exe"
if not exist "as400\launch_listesi.py" (echo [HATA] as400\launch_listesi.py yok - once Guncelle.bat & pause & exit /b 1)

"%PY%" as400\launch_listesi.py %*
echo.
echo   Excel masaustune kopyalandi (Launch_Listesi_...xlsx). Laptopa alip
echo   Masaustu\uretim_takip\data\denetim klasorune koyun - kontrol edelim.
echo.
pause

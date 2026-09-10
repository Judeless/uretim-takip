@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Cofle - PCOMM Kurtarma
cd /d "%~dp0"
set "CS=C:\Windows\SysWOW64\cscript.exe"
set "PS=powershell -NoProfile -ExecutionPolicy Bypass -Command"
if not defined COFLE_BASLANGIC set "COFLE_BASLANGIC=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

REM ============================================================
REM   PCOMM KURTARMA - sunucuyu YENIDEN BASLATMADAN (2026-09-10)
REM ============================================================
REM Belirti: PCOMM "Personal Communications 5.8" acilis ekraninda takiliyor,
REM oturumlar acilmiyor; teyit-agent ECL37110 (emulasyon arayuzu yok) veriyor.
REM Sebep: eski pcsws.exe / pcscm.exe (Session Manager) sureci asili kalmis,
REM yeni oturum onu bekliyor. Cozum: pcs* sureclerini kapat, A sonra B'yi ac.
REM Sunucu reboot GEREKMEZ (reboot MES servisini de dusuruyordu).
REM
REM Bu betigi PROMANAGE RDP oturumunda, YONETICI OLMADAN calistir.
REM (Yonetici pencereden acilan PCOMM'u robot GOREMEZ - baska baglam.)
REM tasklist/taskkill KULLANILMAZ: bu sunucuda tasklist asiliyor (bkz. Robot_Tani.bat).
REM ============================================================

echo ============================================================
echo   PCOMM KURTARMA - sunucuyu yeniden baslatmadan
echo ============================================================
echo Bu pencere: oturum %SESSIONNAME%, kullanici %USERNAME%
echo [UYARI] Pencereye TIKLAMAYIN (QuickEdit). Donarsa ESC.
echo.
%PS% "if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Host '   [!!] Bu pencere YONETICI. PCOMM buradan acilirsa robot onu goremez.'; Write-Host '        Normal (yonetici olmayan) bir pencereden calistirin.'; exit 3 } else { exit 0 }"
if errorlevel 3 goto :son

echo 1/4  Calisan PCOMM ve agent surecleri:
%PS% "$p=Get-Process pcs*,python -ErrorAction SilentlyContinue; if($p){ $p | Select-Object Name,Id,SessionId,StartTime | Sort-Object Name | Format-Table -AutoSize | Out-String } else { '   surec yok' }"

echo 2/4  PCOMM surecleri kapatiliyor (pcsws, pcscm, pcsmon...):
%PS% "$p=Get-Process pcs* -ErrorAction SilentlyContinue; if(-not $p){ '   kapatilacak PCOMM sureci yok' }; foreach($x in $p){ try { Stop-Process -Id $x.Id -Force -ErrorAction Stop; '   kapatildi: ' + $x.Name + ' (' + $x.Id + ', oturum ' + $x.SessionId + ')' } catch { '   KAPATILAMADI: ' + $x.Name + ' (' + $x.Id + ', oturum ' + $x.SessionId + ') - baska kullanicinin oturumu olabilir' } }"
timeout /t 5 /nobreak >nul
%PS% "$k=Get-Process pcs* -ErrorAction SilentlyContinue; if($k){ '   [!!] Hala calisan: ' + (($k | ForEach-Object { $_.Name + '(' + $_.SessionId + ')' }) -join ', '); exit 2 } else { '   temiz - PCOMM sureci kalmadi'; exit 0 }"
if errorlevel 2 (
    echo    Kalan surecleri Gorev Yoneticisi ^> Ayrintilar'dan ^(yonetici^) sonlandirin,
    echo    sonra bu betigi NORMAL pencereden tekrar calistirin.
    goto :son
)

echo 3/4  Oturumlar aciliyor (Baslangic klasorundeki PCOMM kisayollari):
set ACILDI=0
for %%L in ("PCOMM A.lnk" "PCOMM B.lnk") do (
    if exist "%COFLE_BASLANGIC%\%%~L" (
        echo    aciliyor: %%~L
        start "" "%COFLE_BASLANGIC%\%%~L"
        set /a ACILDI+=1
        timeout /t 12 /nobreak >nul
    )
)
if "%ACILDI%"=="0" (
    echo    Baslangic'ta "PCOMM A.lnk" / "PCOMM B.lnk" yok ^(Otomatik_Kalkis_Kur.bat kurar^).
    echo    Simdi masaustunden ONCE A sonra B'yi acin ^(yonetici olmadan^), sonra ENTER.
    pause >nul
) else (
    echo    %ACILDI% oturum baslatildi, emulator otursun diye 15 sn bekleniyor...
    timeout /t 15 /nobreak >nul
)

echo 4/4  Otomasyon oturumlari goruyor mu?
> "%TEMP%\cofle_pcomm_liste.js" echo var l = new ActiveXObject("PCOMM.autECLConnList");
>>"%TEMP%\cofle_pcomm_liste.js" echo l.Refresh();
>>"%TEMP%\cofle_pcomm_liste.js" echo WScript.Echo("   Otomasyonun gordugu baglanti sayisi: " + l.Count);
>>"%TEMP%\cofle_pcomm_liste.js" echo for (var i = 1; i ^<= l.Count; i++) WScript.Echo("      ad=[" + l(i).Name + "]  baslatildi=" + l(i).Started);
>>"%TEMP%\cofle_pcomm_liste.js" echo if (l.Count ^< 2) WScript.Quit(2);
"%CS%" //nologo "%TEMP%\cofle_pcomm_liste.js"
set RC=%errorlevel%
del "%TEMP%\cofle_pcomm_liste.js" >nul 2>&1
echo.
if "%RC%"=="0" (
    echo   [OK] PCOMM A ve B otomasyona gorunuyor.
    echo        Session B sign-on ekranindaysa gozcu en gec 5 dk icinde girer
    echo        ^(hemen istersen agent penceresini kapatip Teyit_Agent_Baslat.bat ile ac^).
    echo        Session A'ya elle sign-on yapin.
) else (
    echo   [!!] Otomasyon 2 baglanti gormuyor. Robot_Tani.bat ile tani alin:
    echo        pencereler acik ama sayi 0 ise PCOMM baska baglamda ^(yonetici / baska oturum^) acilmistir.
)

:son
echo.
pause
endlocal

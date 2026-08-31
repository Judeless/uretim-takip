@echo off
setlocal enabledelayedexpansion
REM ===================================================================
REM  SUNUCU OTOMATIK KALKIS KURULUMU
REM  Bir kez calistirilir - promanage oturumunda, YONETICI OLARAK DEGIL.
REM
REM  NEDEN (kullanici 2026-08-27: "hafta sonu AS400 hataya gectiginde bu
REM  sabah geldigimde otomatik baglanmamisti"): PCOMM ve teyit-agent
REM  INTERAKTIF oturumda yasar. Sunucu yeniden baslarsa (Windows
REM  guncellemesi, elektrik) kimse oturum acmadigi icin ikisi de kalkmaz;
REM  oturum gozcusu de agent'in icinde oldugu icin olur. Sonuc: kimse fark
REM  edene kadar hicbir teyit gitmez.
REM
REM  Bu betik LOGON zincirini kurar:
REM     otomatik oturum acma -> PCOMM A + B -> teyit-agent -> gozcu sign-on
REM  Otomatik oturum acmayi betik YAPMAZ (sifre gerektirir) - durumunu
REM  raporlar ve nasil kurulacagini yazar. Kisayollari kurar.
REM
REM  PCOMM kisayollari icin SIRAYLA denenir:
REM     1) argumanla verilen .ws yollari
REM     2) SU AN CALISAN pcsws.exe oturumlari  <- en guvenilir, tahmin yok
REM     3) bilinen klasorlerde *.ws aramasi
REM
REM  Kullanim:
REM     Otomatik_Kalkis_Kur.bat                 -> kendisi bulur
REM     Otomatik_Kalkis_Kur.bat "A.ws" "B.ws"   -> profilleri sen verirsin
REM
REM  BAKIM NOTU: for /f ( '...' ) icindeki PowerShell komutunda VIRGUL ve
REM  BORU KULLANMA. cmd ikisini de bozuyor: virgul komutu parcaliyor (cikti
REM  bos doner), boru kacisi ^| PowerShell'e duz '^' argumani olarak geciyor
REM  ("A positional parameter cannot be found that accepts argument '^'").
REM  Dizi virgulle degil += ile kurulur, boru yerine foreach kullanilir.
REM ===================================================================
cd /d %~dp0
if not defined COFLE_BASLANGIC set COFLE_BASLANGIC=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set BASLANGIC=%COFLE_BASLANGIC%
set PS=powershell -NoProfile -ExecutionPolicy Bypass -Command
if not defined COFLE_PCOMM_EXE set COFLE_PCOMM_EXE=pcsws.exe
set PCOMM_KURULDU=0

echo.
echo ==========================================================
echo    COFLE - SUNUCU OTOMATIK KALKIS KURULUMU
echo ==========================================================
echo    Oturum    : %USERNAME%   (PCOMM bu oturumda calismali)
echo    Baslangic : %BASLANGIC%
echo.
if not exist "%BASLANGIC%" (
    echo    HATA: Baslangic klasoru bulunamadi. Win+R  shell:startup  ile
    echo          dogru yolu bul ve betigi o yolla calistir:
    echo             set COFLE_BASLANGIC=^<yol^>
    echo.
    pause
    exit /b 1
)

echo [1/4] PCOMM oturum kisayollari:

REM --- 1a) Yollar argumanla verildiyse -------------------------------
if not "%~1"=="" (
    call :kisayol "PCOMM A.lnk" "%~1" ""
    call :kisayol "PCOMM B.lnk" "%~2" ""
    set PCOMM_KURULDU=1
    goto ajan
)

REM --- 1b) SU AN CALISAN PCOMM oturumlarindan ------------------------
REM En guvenilir yol: emulator zaten aciksa onu NEYLE baslattiysak aynisini
REM Baslangic'a koyariz - .ws dosyasini aramaya gerek kalmaz. Kisayolu
REM PowerShell olusturur (for /f ayristirmasi yok), kac tane kurdugunu
REM cikis koduyla bildirir.
%PS% "$b='%BASLANGIC%'; $n=0; foreach($p in @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)){ if($p.Name -eq '%COFLE_PCOMM_EXE%'){ $n=$n+1; $exe=$p.ExecutablePath; $ws=''; foreach($q in ($p.CommandLine -split [char]34)){ if($q -like '*.ws'){ $ws=$q.Trim() } }; $ad='PCOMM ' + $n; if($ws -ne ''){ $ad='PCOMM ' + [System.IO.Path]::GetFileNameWithoutExtension($ws) }; $s=(New-Object -ComObject WScript.Shell).CreateShortcut($b + '\' + $ad + '.lnk'); $s.TargetPath=$exe; if($ws -ne ''){ $s.Arguments=[char]34 + $ws + [char]34 }; $s.WorkingDirectory=[System.IO.Path]::GetDirectoryName($exe); $s.Save(); ('       OK: ' + $ad + '.lnk  ->  ' + $exe + ' ' + $ws) } }; exit $n"
if %ERRORLEVEL% GEQ 1 (
    set PCOMM_KURULDU=1
    echo        Calisan %ERRORLEVEL% PCOMM oturumu Baslangic'a eklendi.
    goto ajan
)
echo        PCOMM (%COFLE_PCOMM_EXE%) su an calismiyor - profil dosyasi araniyor...

REM --- 1c) Bilinen klasorlerde *.ws aramasi -------------------------
REM PS 5.1: Get-ChildItem -Path'e COKLU yol verilip biri yoksa -Recurse
REM hicbir sonuc dondurmuyor -> yollar tek tek Test-Path ile dolasiliyor.
set BULUNAN=0
for /f "delims=" %%F in ('%PS% "$y=@(); $y+=$env:APPDATA+'\IBM'; $y+=$env:LOCALAPPDATA+'\IBM'; $y+=$env:ProgramData+'\IBM'; $y+=$env:USERPROFILE+'\Documents'; $y+=$env:USERPROFILE+'\Desktop'; $y+=$env:PUBLIC+'\Documents'; $y+=$env:PUBLIC+'\Desktop'; $y+=$env:ProgramFiles+'\IBM'; $y+='C:\Program Files (x86)\IBM'; $n=0; foreach($p in $y){ if(Test-Path -LiteralPath $p){ foreach($f in @(Get-ChildItem -LiteralPath $p -Filter *.ws -Recurse -ErrorAction SilentlyContinue)){ if($n -lt 20){ $f.FullName; $n=$n+1 } } } }"') do (
    if exist "%%F" (
        set /a BULUNAN+=1
        set "WS_!BULUNAN!=%%F"
        echo        !BULUNAN!. %%F
    )
)
if %BULUNAN%==0 (
    echo.
    echo        .ws profili bulunamadi. Iki secenek:
    echo          a^) PCOMM A ve B oturumlarini AC, betigi tekrar calistir -
    echo             calisan oturumlardan kendisi kurar. EN KOLAYI BU.
    echo          b^) Yolu kendin ver. Ogrenmek icin PCOMM'da "Dosya - Farkli
    echo             kaydet" penceresine bak, sonra:
    echo                Otomatik_Kalkis_Kur.bat "C:\yol\A.ws" "C:\yol\B.ws"
    goto ajan
)
echo.
set /p A_NO=   A oturumunun numarasi:
set /p B_NO=   B oturumunun numarasi:
REM Bosluklari at: kullanici " 1" yazarsa WS_ 1 diye var olmayan degisken aranir
set "A_NO=%A_NO: =%"
set "B_NO=%B_NO: =%"
set "A_WS=!WS_%A_NO%!"
set "B_WS=!WS_%B_NO%!"
if not exist "!A_WS!" (
    echo        HATA: %A_NO% numarali gecerli bir profil yok - kisayol kurulmadi.
    goto ajan
)
if not exist "!B_WS!" (
    echo        HATA: %B_NO% numarali gecerli bir profil yok - kisayol kurulmadi.
    goto ajan
)
call :kisayol "PCOMM A.lnk" "!A_WS!" ""
call :kisayol "PCOMM B.lnk" "!B_WS!" ""
set PCOMM_KURULDU=1

:ajan
echo.
echo [2/4] Teyit-agent kisayolu:
call :kisayol "Cofle Teyit Agent.lnk" "%~dp0Teyit_Agent_Baslat.bat" "%~dp0"

REM --- 3) Otomatik oturum acma durumu (SALT OKUNUR) ------------------
echo.
echo [3/4] Otomatik oturum acma (autologon):
REM Tek satirda coklu alan okumak kirilgan (bos alanlar birlesiyor, virgul
REM komutu bozuyor) -> her deger AYRI sorgu. Cikti bossa varsayilan kalir.
set W=HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon
set AAL=0
set DUN=(tanimsiz)
set DPW=0
for /f "delims=" %%A in ('%PS% "$p=Get-ItemProperty '%W%' -ErrorAction SilentlyContinue; if([string]$p.AutoAdminLogon -eq '1'){'1'}else{'0'}"') do set "AAL=%%A"
for /f "delims=" %%A in ('%PS% "$p=Get-ItemProperty '%W%' -ErrorAction SilentlyContinue; [string]$p.DefaultUserName"') do set "DUN=%%A"
for /f "delims=" %%A in ('%PS% "$p=Get-ItemProperty '%W%' -ErrorAction SilentlyContinue; if($p.DefaultPassword){'1'}else{'0'}"') do set "DPW=%%A"
if "%AAL%"=="1" (
    echo        AutoAdminLogon : ACIK    kullanici: %DUN%
) else (
    echo        AutoAdminLogon : KAPALI  ^<-- kurulmasi gerekiyor
)
if "%DPW%"=="1" (
    echo.
    echo        GUVENLIK UYARISI: Winlogon\DefaultPassword kayit defterinde
    echo        DUZ METIN sifre tutuyor. Sysinternals Autologon ile kurarsan
    echo        sifre LSA kasasinda sifreli durur. Duz metin girdiyi sil:
    echo           reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /f
)

REM --- 4) Ozet ------------------------------------------------------
echo.
echo [4/4] Sirada ne var:
if "%PCOMM_KURULDU%"=="0" (
    echo        - PCOMM kisayollari KURULMADI. En kolayi: emulator A+B acikken
    echo          bu betigi tekrar calistir.
)
if not "%AAL%"=="1" (
    echo        - Otomatik oturum acmayi kur: Sysinternals "Autologon"
    echo          [learn.microsoft.com/sysinternals/downloads/autologon]
    echo          Autologon64.exe calistir - kullanici %USERNAME%, sifre, Enable.
    echo          Sifre LSA kasasinda sifreli tutulur, kayit defterinde DEGIL.
)
echo        - Sunucuyu yeniden baslat, HICBIR SEYE DOKUNMADAN 5-6 dk bekle.
echo        - Panelde AS400 ayar blogunda sunu gor:
echo             "teyit-agent bagli"   ve   "gozcu acik"
echo        - Bir satir teyit gonderip dogrula.
echo.
echo    NOT: oturumu KAPATMA (logoff) - PCOMM + agent olur.
echo         RDP'den cikarken daima "Disconnect" kullan.
echo.
echo    Baslangic klasorunun icerigi:
dir /b "%BASLANGIC%"
echo.
pause
exit /b 0

:kisayol
REM %1 = kisayol adi, %2 = hedef, %3 = calisma klasoru
if "%~2"=="" echo        ATLANDI: %~1 (hedef bos) & exit /b 0
if not exist "%~2" echo        ATLANDI: %~1 (bulunamadi: %~2) & exit /b 0
%PS% "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%BASLANGIC%\%~1'); $s.TargetPath='%~2'; if ('%~3' -ne '') { $s.WorkingDirectory='%~3' }; $s.Save()" >nul 2>&1
if exist "%BASLANGIC%\%~1" (echo        OK: %~1) else (echo        HATA: %~1 olusturulamadi)
exit /b 0

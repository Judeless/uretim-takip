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
REM  Kullanim:
REM     Otomatik_Kalkis_Kur.bat                 -> .ws profillerini arar
REM     Otomatik_Kalkis_Kur.bat "A.ws" "B.ws"   -> profilleri sen verirsin
REM ===================================================================
cd /d %~dp0
if not defined COFLE_BASLANGIC set COFLE_BASLANGIC=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set BASLANGIC=%COFLE_BASLANGIC%
set PS=powershell -NoProfile -ExecutionPolicy Bypass -Command

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

REM --- 1) PCOMM oturum profilleri (.ws) -----------------------------
set A_WS=%~1
set B_WS=%~2
if not "%A_WS%"=="" goto kisayollar

echo [1/4] PCOMM oturum profilleri araniyor (*.ws)...
REM DIKKAT: for /f ( '...' ) icindeki PowerShell komutunda VIRGUL ve BORU
REM KULLANMA. cmd ikisini de bozuyor: virgul komutu parcaliyor (cikti bos
REM doner), boru ^| kacisi PowerShell'e duz '^' argumani olarak geciyor
REM ("A positional parameter cannot be found that accepts argument '^'").
REM Bu yuzden dizi virgulle degil += ile kuruluyor ve boru yerine foreach var.
REM Ayrica Get-ChildItem -Path'e COKLU yol verilip biri yoksa PS 5.1 hicbir
REM sonuc dondurmuyor -> yollar tek tek Test-Path ile dolasiliyor.
set BULUNAN=0
for /f "delims=" %%F in ('%PS% "$y=@(); $y+=$env:USERPROFILE+'\Documents'; $y+=$env:PUBLIC+'\Documents'; $y+=$env:APPDATA+'\IBM'; $y+=$env:ProgramData+'\IBM'; $n=0; foreach($p in $y){ if(Test-Path -LiteralPath $p){ foreach($f in @(Get-ChildItem -LiteralPath $p -Filter *.ws -Recurse -ErrorAction SilentlyContinue)){ if($n -lt 20){ $f.FullName; $n=$n+1 } } } }"') do (
    if exist "%%F" (
        set /a BULUNAN+=1
        set "WS_!BULUNAN!=%%F"
        echo        !BULUNAN!. %%F
    )
)
if %BULUNAN%==0 (
    echo.
    echo    .ws profili bulunamadi. PCOMM'da A ve B oturumlarini acip
    echo    "Dosya - Farkli kaydet" ile profilleri kaydet, sonra:
    echo        Otomatik_Kalkis_Kur.bat "C:\yol\A.ws" "C:\yol\B.ws"
    echo.
    pause
    exit /b 1
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
    echo.
    echo    HATA: %A_NO% numarali gecerli bir profil yok. Yollari kendin ver:
    echo        Otomatik_Kalkis_Kur.bat "C:\yol\A.ws" "C:\yol\B.ws"
    echo.
    pause
    exit /b 1
)
if not exist "!B_WS!" (
    echo.
    echo    HATA: %B_NO% numarali gecerli bir profil yok. Yollari kendin ver:
    echo        Otomatik_Kalkis_Kur.bat "C:\yol\A.ws" "C:\yol\B.ws"
    echo.
    pause
    exit /b 1
)

:kisayollar
echo.
echo [2/4] Baslangic kisayollari:
call :kisayol "PCOMM A.lnk"           "%A_WS%"                      ""
call :kisayol "PCOMM B.lnk"           "%B_WS%"                      ""
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
pause
exit /b 0

:kisayol
REM %1 = kisayol adi, %2 = hedef, %3 = calisma klasoru
if "%~2"=="" echo        ATLANDI: %~1 (hedef bos) & exit /b 0
if not exist "%~2" echo        ATLANDI: %~1 (bulunamadi: %~2) & exit /b 0
%PS% "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%BASLANGIC%\%~1'); $s.TargetPath='%~2'; if ('%~3' -ne '') { $s.WorkingDirectory='%~3' }; $s.Save()" >nul 2>&1
if exist "%BASLANGIC%\%~1" (echo        OK: %~1) else (echo        HATA: %~1 olusturulamadi)
exit /b 0

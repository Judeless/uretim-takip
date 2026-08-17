@echo off
chcp 65001 >nul
setlocal
title Cofle - Robot Tanisi (PCOMM / cscript)

REM ============================================================
REM   ROBOT TANISI — "Robot hic cikti uretmedi" hatasi icin
REM ============================================================
REM Bu dosyayi PROMANAGE RDP OTURUMUNDA calistir (PCOMM'un yaninda).
REM Servis oturumunda (Session 0) calistirma — orada PCOMM zaten gorunmez.
REM
REM Iki test yapar, HICBIRI AS400'e YAZMAZ:
REM   TEST 1: cscript + Windows Script Host calisiyor mu (PCOMM'a dokunmaz)
REM   TEST 2: PCOMM Session B'ye baglanilabiliyor mu (ekrani okur, yazmaz)
REM ============================================================

cd /d "%~dp0"
set "CS=C:\Windows\SysWOW64\cscript.exe"

REM QuickEdit tuzagi: pencereye TIKLAMA. Konsol "Sec/Isaretle" moduna girer ve
REM cikti DONAR (baslik "Sec Komut Istemi" olur). Donduysa ESC'e bas, devam eder.
echo ============================================================
echo   COFLE ROBOT TANISI
echo ============================================================
echo.
echo [UYARI] Bu pencereye TIKLAMAYIN - konsol "Sec" moduna girip cikti donar.
echo         Donarsa ESC'e basin, kaldigi yerden devam eder.
echo.

if not exist "%CS%" (
    echo [HATA] 32-bit cscript yok: %CS%
    echo        PCOMM COM arayuzu 32-bittir, SysWOW64 cscript sart.
    goto :son
)
echo [OK] cscript bulundu: %CS%
echo.

echo ------------------------------------------------------------
echo TEST 1/2 - cscript + Windows Script Host  (PCOMM'a DOKUNMAZ)
echo ------------------------------------------------------------
echo Beklenen: "HATA: arguman eksik ..." + "SONUC=IPTAL", cikis kodu 2
echo Cikti:
"%CS%" //nologo teyit_gir.js
echo   -^> cikis kodu: %errorlevel%
echo.
if "%errorlevel%"=="0" (
    echo [!!] Cikis kodu 0 ve/veya cikti YOK.
    echo      Sorun PCOMM DEGIL, script host tarafinda:
    echo        - Windows Script Host devre disi birakilmis olabilir
    echo          ^(HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings\Enabled^)
    echo        - .js dosyasi baska bir motora baglanmis olabilir
    echo        - antivirus/uygulama-kontrolu cscript'i sessizce engelliyor olabilir
    goto :son
)

echo ------------------------------------------------------------
echo OTURUM DURUMU  (ECL37110 hatasinin en sik sebebi)
echo ------------------------------------------------------------
echo Bu pencere su oturumda: %SESSIONNAME%  (kullanici: %USERNAME%)
echo.
echo Sunucudaki TUM oturumlar:
REM NOT: query session tabloyu bassa da bazen sifir-disi cikis kodu donuyor —
REM bu yuzden "|| echo calistirilamadi" YANILTICIYDI, kaldirildi.
query session 2>nul
echo.
echo PCOMM (pcsws) ve teyit-agent (python) hangi oturumda?
REM TASKLIST KULLANMA (2026-08-17): bu sunucuda tasklist ASILIYOR — teyit_agent.py
REM icindeki not da ayni seyi soyluyor ("bazi sunucularda tasklist asilip agent'i
REM startta kilitliyordu, PCOMM sayaci kaldirildi"). Get-Process saf Win32 surec
REM listesi kullanir, RPC/WMI beklemesi yoktur ve SessionId'yi dogrudan verir.
powershell -NoProfile -Command "$p = Get-Process pcsws,python -ErrorAction SilentlyContinue; if ($p) { $p | Select-Object Name,Id,SessionId | Sort-Object Name,SessionId | Format-Table -AutoSize | Out-String } else { 'HIC SUREC YOK' }"
echo.
powershell -NoProfile -Command "if (Get-Process pcsws -ErrorAction SilentlyContinue) { '' } else { '   *** pcsws YOK - PCOMM emulator HIC CALISMIYOR ***'; '   Cozum: bu oturumda once A sonra B oturumunu ac, elle sign-on yap.' }"
echo.
echo   ^>^> pcsws.exe ile teyit-agent AYNI oturum numarasinda olmali.
echo      Farkliysa robot PCOMM'u goremez - hata tam olarak ECL37110'dur.
echo.

echo ------------------------------------------------------------
echo PCOMM OTURUM LISTESI  (otomasyon KAC baglanti goruyor?)
echo ------------------------------------------------------------
REM Kritik ayrim (2026-08-17): pcsws.exe calisiyor olmasi YETMEZ. Otomasyon
REM katmani ayni BAGLAM'dan (ayni Windows oturumu + ayni yukseltme seviyesi)
REM bakmazsa baglanti listesi BOS gorunur ve SetConnectionByName("B") ECL37110
REM verir. Sahada tam olarak bu yasandi: pcsws x2 vardi ama Count=0 idi.
powershell -NoProfile -Command "'Bu pencere yonetici mi : ' + ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)"
> "%TEMP%\cofle_pcomm_liste.js" echo var l = new ActiveXObject("PCOMM.autECLConnList");
>>"%TEMP%\cofle_pcomm_liste.js" echo l.Refresh();
>>"%TEMP%\cofle_pcomm_liste.js" echo WScript.Echo("Otomasyonun gordugu baglanti sayisi: " + l.Count);
>>"%TEMP%\cofle_pcomm_liste.js" echo for (var i = 1; i ^<= l.Count; i++) WScript.Echo("   ad=[" + l(i).Name + "]  baslatildi=" + l(i).Started);
"%CS%" //nologo "%TEMP%\cofle_pcomm_liste.js"
del "%TEMP%\cofle_pcomm_liste.js" >nul 2>&1
echo.
echo   ^>^> Sayi 0 ise: pcsws calisiyor olsa bile otomasyon onlari GORMUYOR.
echo      Ya PCOMM baska bir baglamda acilmis (yonetici olarak / baska oturum),
echo      ya da pencereler gercek emulator oturumu degil (Session Manager).
echo      Cozum: PCOMM'u KAPAT, promanage oturumunda YONETICI OLMADAN
echo      once A sonra B'yi ac, sign-on yap, bu tanıyı tekrar calistir.
echo   ^>^> Sayi 2 ve adlar [A],[B] ise robot yolu hazir demektir.
echo.

echo ------------------------------------------------------------
echo TEST 2/2 - PCOMM Session B baglantisi  (SADECE OKUR, YAZMAZ)
echo ------------------------------------------------------------
> "%TEMP%\cofle_pcomm_test.js" echo var s = new ActiveXObject("PCOMM.autECLSession");
>>"%TEMP%\cofle_pcomm_test.js" echo s.SetConnectionByName("B");
>>"%TEMP%\cofle_pcomm_test.js" echo var ps = s.autECLPS;
>>"%TEMP%\cofle_pcomm_test.js" echo WScript.Echo("BAGLANDI - Session B, " + ps.NumRows + "x" + ps.NumCols);
>>"%TEMP%\cofle_pcomm_test.js" echo WScript.Echo("EKRANIN ILK SATIRI: " + ps.GetText(1, 1, ps.NumCols));
echo Beklenen: "BAGLANDI - Session B, 24x80" + ekranin ilk satiri
echo Cikti:
"%CS%" //nologo "%TEMP%\cofle_pcomm_test.js"
echo   -^> cikis kodu: %errorlevel%
del "%TEMP%\cofle_pcomm_test.js" >nul 2>&1
echo.
echo ------------------------------------------------------------
echo YORUM
echo ------------------------------------------------------------
echo  * TEST 2 "ECL37110 / emulazione non disponibile" derse:  ^<-- EN SIK
echo      PCOMM KURULU ama bu Windows oturumunda ACIK EMULATOR YOK.
echo      Yukaridaki OTURUM DURUMU bolumune bak:
echo        - pcsws.exe hic yoksa  -^> PCOMM'u bu oturumda ac (A sonra B) + sign-on
echo        - pcsws.exe VAR ama BASKA oturumdaysa -^> promanage'in birden fazla
echo          RDP oturumu var demektir. Fazlasini logoff et, PCOMM'u ve
echo          teyit-agent'i AYNI oturumda birak.
echo      NOT: bu hatada cscript yine de 0 ile cikar, aldanma.
echo  * TEST 2 "ActiveX component can't create object" derse:
echo      PCOMM kurulu degil ya da bu oturumda calismiyor.
echo  * TEST 2 "SetConnectionByName" hatasi verirse:
echo      PCOMM acik ama "B" adinda oturum YOK. Pencere acilis SIRASI onemli:
echo      ilk acilan = A, ikinci = B.
echo  * TEST 2 BAGLANDI derse ama ilk satir bos/sign-on ekraniysa:
echo      Oturum acilmamis. Elle sign-on yap (robot sifre girmez).
echo  * TEST 2 BAGLANDI ve ekranda ana menu varsa robot yolu SAGLAM;
echo      sorun teyit_agent tarafindadir (agent penceresi ayakta mi?).

:son
echo.
echo ============================================================
pause
endlocal

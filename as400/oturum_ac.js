// ═══════════════════════════════════════════════════════════════════
// AS400 OTURUM KURTARMA — Session B bagli degilse baglar ve sign-on yapar.
// Kullanim: cscript //nologo oturum_ac.js <kullanici> [DRYRUN]
//
// ŞİFRE ARGÜMANLA GECMEZ. Ortam degiskeninden okunur: COFLE_AS400_PW
//   Sebep: argumanlar agent konsoluna ve Windows surec listesine (Get-Process
//   /CommandLine, Sysmon) DUSER. Ortam degiskeni cocuk surece ozeldir, loglanmaz.
//   Sifre hicbir yere yazilmaz; log dosyasina da YALNIZ uzunlugu gecer.
//
// Neden var (kullanici 2026-08-17): sunucudaki Session B uzun sure islem
// yapilmayinca dusuyor (QINACTITV / baglanti kopmasi) ve teyit robotu bir daha
// calisamiyor. Kullanici izinde oldugunda elle sign-on yapacak kimse yok.
// NOT: bu, kurulum kilavuzundaki "sign-on INSAN isi" kuralini bilerek gevsetir —
// kullanicinin acik istegi. Sifre yine kasada (keyring), kodda/dosyada DEGIL.
//
// Cikti son satiri: SONUC=OK | SONUC=ZATEN | SONUC=IPTAL | SONUC=DRYRUN-OK
// Log: as400\teyit_loglari\<zaman>_OTURUM.txt   (SIFRE YAZILMAZ)
// ═══════════════════════════════════════════════════════════════════
var args = WScript.Arguments;
if (args.length < 1) { WScript.Echo("HATA: kullanici adi eksik"); WScript.Echo("SONUC=IPTAL"); WScript.Quit(2); }
var KULLANICI = ("" + args(0)).replace(/^\s+|\s+$/g, "");
var DRYRUN = false, TESTBAGLAN = false;
for (var ai = 1; ai < args.length; ai++) {
    var _a = ("" + args(ai)).toUpperCase();
    if (_a === "DRYRUN") DRYRUN = true;
    // TESTBAGLAN: baglanti KURAR, sign-on ekranini ve ALAN KONUMLARINI raporlar,
    // HICBIR TUS GONDERMEZ ve kendi kurdugu baglantiyi geri kapatir. Amac: alan
    // tespitini gercek ekranda dogrulamak. Yanlis sifre denemek AS400'de profili
    // kilitleyebilecegi icin (QMAXSIGN) "deneme sign-on" ASLA yapilmaz.
    if (_a === "TESTBAGLAN") { TESTBAGLAN = true; DRYRUN = true; }
}

if (!/^[A-Za-z0-9_$#@.\-]{1,20}$/.test(KULLANICI)) {
    WScript.Echo("HATA: gecersiz kullanici adi"); WScript.Echo("SONUC=IPTAL"); WScript.Quit(2);
}
var sh = new ActiveXObject("WScript.Shell");
var SIFRE = "";
try { SIFRE = "" + sh.Environment("PROCESS").Item("COFLE_AS400_PW"); } catch (e) { SIFRE = ""; }
if (!DRYRUN && !SIFRE) {
    WScript.Echo("HATA: COFLE_AS400_PW ortam degiskeni bos — kasada sifre yok mu?");
    WScript.Echo("SONUC=IPTAL"); WScript.Quit(2);
}

// LOG DOSYASI OPSIYONEL (2026-08-17): bu robot kullanici IZINDEYKEN gozetimsiz
// kosacak. Log dosyasi acilamazsa (klasor yok/kilitli/disk dolu ya da yolda ASCII
// disi karakter — cscript ANSI dosya API'si kullanir) OTURUM KURTARMA YAPILMADAN
// olmemeli. Diger robotlarda log fatal'dir; burada BILEREK degil.
var logDosya = null;
try {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var kok = fso.GetParentFolderName(WScript.ScriptFullName);
    var logKlasor = kok + "\teyit_loglari";
    if (!fso.FolderExists(logKlasor)) fso.CreateFolder(logKlasor);
    var dt = new Date();
    var p2 = function (n) { return (n < 10 ? "0" : "") + n; };
    var damga = dt.getFullYear() + p2(dt.getMonth() + 1) + p2(dt.getDate()) + "_" +
                p2(dt.getHours()) + p2(dt.getMinutes()) + p2(dt.getSeconds());
    logDosya = fso.OpenTextFile(logKlasor + "\\" + damga + "_OTURUM.txt", 2, true);
} catch (e) {
    logDosya = null;   // stdout yeterli — agent zaten yakaliyor
}
function log(m) {
    WScript.Echo(m);
    if (logDosya) { try { logDosya.WriteLine(m); } catch (e) {} }
}
function logKapat() { if (logDosya) { try { logKapat(); } catch (e) {} } }

var EKRAN_BEKLE_SN = 30;   // baglanti sonrasi ekranin boyanmasi icin azami bekleme
var bizBagladik = false;
var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName("B");
var ps = s.autECLPS, oia = s.autECLOIA;
function ekran() { var t = []; for (var r = 1; r <= ps.NumRows; r++) t.push(ps.GetText(r, 1, ps.NumCols)); return t; }
function ekranStr() { return ekran().join("\n"); }
function bekle() { try { oia.WaitForInputReady(8000); } catch (e) {} WScript.Sleep(400); }
function iptal(neden) {
    log("HATA: " + neden);
    log("SONUC=IPTAL"); logKapat(); WScript.Quit(2);
}
// Sign-on ekrani mi? (Italyanca PCOMM: Collegamento / Utente / Parola d'ordine)
function signOnMu(e) {
    return e.indexOf("Utente") >= 0 && e.indexOf("Parola") >= 0;
}
function anaMenuMu(e) { return e.indexOf("Menu S.I. CofleTk") >= 0 || e.indexOf("MenuIniziale") >= 0; }
// Satirdaki ilk GIRIS alani (korumasiz) — kolon varsaymak yerine ekrandan okunur
function ilkGirisAlani(satir) {
    try {
        var fl = ps.autECLFieldList; fl.Refresh();
        var en = null;
        for (var i = 1; i <= fl.Count; i++) {
            var f = fl(i);
            if (Number(f.StartRow) === Number(satir) && !f.Protected && Number(f.Length) > 0)
                if (!en || Number(f.StartCol) < en.kol) en = { kol: Number(f.StartCol), uz: Number(f.Length) };
        }
        return en;
    } catch (e) { return null; }
}
function satirBul(dizi, etiket) {
    for (var i = 0; i < dizi.length; i++) if (dizi[i].indexOf(etiket) >= 0) return i + 1;
    return -1;
}

log("Oturum kurtarma — kullanici=" + KULLANICI + " sifre=" + (SIFRE ? SIFRE.length + " karakter (kasadan)" : "YOK") + (DRYRUN ? " (DRYRUN)" : ""));

// ── 1) BAGLANTI (Comunicazioni > Connetti'nin programli karsiligi) ──
var bagliydi = false;
try { bagliydi = !!s.CommStarted; } catch (e) { bagliydi = false; }
log("Baslangic: CommStarted=" + bagliydi);
if (!bagliydi) {
    if (DRYRUN && !TESTBAGLAN) { log("DRYRUN: baglanti KURULMADI."); log("SONUC=DRYRUN-OK"); logKapat(); WScript.Quit(0); }
    log("Baglanti kuruluyor (StartCommunication)...");
    try { s.StartCommunication(); } catch (e) { iptal("StartCommunication basarisiz: " + e.message); }
    for (var i = 0; i < 30 && !s.CommStarted; i++) WScript.Sleep(500);
    if (!s.CommStarted) iptal("Baglanti kurulamadi (15 sn)");
    bizBagladik = true;
    log("Baglanti kuruldu.");
    WScript.Sleep(1500);
}

// ── 2) EKRAN NE DURUMDA? ──
// CommStarted=true OTURUMUN KULLANILABILIR OLDUGUNU GOSTERMEZ (2026-08-17 olcumu):
// laptopta baglanti kuruldu, OIA "hazir" dedi ama host 30 sn boyunca TEK SATIR
// gondermedi. Karar DAIMA EKRAN ICERIGINE gore verilir.
function bizimBaglantiyiKapat() {
    if (bizBagladik) {
        log("Bu kosuda kurulan baglanti geri kapatiliyor (yarim oturum birakma).");
        try { s.StopCommunication(); } catch (e2) { log("  uyari: " + e2.message); }
    }
}
try { oia.WaitForAppAvailable(10000); } catch (e0) {}
var e = "", dolu = 0;
for (var w = 0; w < EKRAN_BEKLE_SN * 2; w++) {
    e = ekranStr();
    dolu = e.replace(/\s/g, "").length;
    if (dolu > 20) break;
    WScript.Sleep(500);
}
if (dolu <= 20) {
    // Bagli ama BOS ekran: host oturum acmadi (cihaz baska yerde kilitli, profil
    // sorunu, 5250 anlasmasi yarim). "Sorun yok" DEME — bu bir arizadir.
    bizimBaglantiyiKapat();
    iptal("Baglanti var ama ekran " + EKRAN_BEKLE_SN + " sn boyunca BOS kaldi — host oturum vermedi. "
          + "Session B'yi elle kontrol edin (cihaz/WorkStationID cakismasi olabilir).");
}
if (anaMenuMu(e)) { log("Oturum ZATEN acik (ana menu) — sign-on gerekmedi."); log("SONUC=ZATEN"); logKapat(); WScript.Quit(0); }
if (!signOnMu(e)) {
    // Sign-on degil, ana menu de degil: teyit robotu bir ekranda birakmis olabilir.
    // DOKUNMA — yanlis ekranda tus gondermek veri girisi riskidir.
    log("Ekran ne sign-on ne ana menu — DOKUNULMADI (robot baska ekranda birakmis olabilir).");
    log("Ilk satir: " + ekran()[0]);
    log("SONUC=ZATEN"); logKapat(); WScript.Quit(0);
}
log("Sign-on ekrani algilandi.");
if (DRYRUN) {
    // Alan tespitini RAPORLA (tus gonderilmeden) — sahada dogrulama icin
    var _sat = ekran();
    var _u = satirBul(_sat, "Utente"), _p = satirBul(_sat, "Parola");
    log("TEST: Utente satiri=" + _u + "  Parola satiri=" + _p);
    if (_u > 0) { var _ua = ilkGirisAlani(_u); log("TEST: utente alani = " + (_ua ? ("kol " + _ua.kol + ", uzunluk " + _ua.uz) : "BULUNAMADI")); }
    if (_p > 0) { var _pa = ilkGirisAlani(_p); log("TEST: parola alani = " + (_pa ? ("kol " + _pa.kol + ", uzunluk " + _pa.uz) : "BULUNAMADI")); }
    bizimBaglantiyiKapat();
    log("DRYRUN: sign-on YAPILMADI (hicbir tus gonderilmedi).");
    log("SONUC=DRYRUN-OK"); logKapat(); WScript.Quit(0);
}

// ── 3) SIGN-ON ──
var sat = ekran();
var uSat = satirBul(sat, "Utente");
var pSat = satirBul(sat, "Parola");
if (uSat < 0 || pSat < 0) iptal("Utente/Parola satiri bulunamadi");
var uAlan = ilkGirisAlani(uSat), pAlan = ilkGirisAlani(pSat);
if (!uAlan || !pAlan) iptal("Giris alanlari okunamadi (utente=" + (uAlan ? uAlan.kol : "-") + " parola=" + (pAlan ? pAlan.kol : "-") + ")");
log("Alanlar: utente@" + uSat + ":" + uAlan.kol + "(" + uAlan.uz + ")  parola@" + pSat + ":" + pAlan.kol + "(" + pAlan.uz + ")");

ps.SetCursorPos(uSat, uAlan.kol); WScript.Sleep(200);
ps.SendKeys(KULLANICI); WScript.Sleep(300);
var yanki = ps.GetText(uSat, uAlan.kol, KULLANICI.length);
if (yanki.replace(/\s+$/, "").toUpperCase() !== KULLANICI.toUpperCase())
    iptal("Kullanici adi yazilamadi (yanki: '" + yanki + "')");

// Sifre: ayri alana imlecle git (Tab davranisina guvenme), yankisi OKUNMAZ/LOGLANMAZ
ps.SetCursorPos(pSat, pAlan.kol); WScript.Sleep(200);
ps.SendKeys(SIFRE); WScript.Sleep(300);
log("Kullanici + sifre alanlari dolduruldu (sifre loglanmaz).");
ps.SendKeys("[enter]"); bekle(); WScript.Sleep(2500);

// ── 4) DOGRULAMA ──
var son = "";
for (var k = 0; k < 12; k++) {
    son = ekranStr();
    if (anaMenuMu(son)) break;
    WScript.Sleep(1000);
}
if (anaMenuMu(son)) {
    log("SIGN-ON BASARILI — ana menu geldi.");
    log("SONUC=OK"); logKapat(); WScript.Quit(0);
}
if (signOnMu(son)) {
    // Hala sign-on: sifre yanlis / hesap kilitli / sifre suresi dolmus olabilir.
    var mesajSatiri = "";
    var ss = ekran();
    for (var m = ss.length - 4; m < ss.length; m++)
        if (m >= 0 && ss[m].replace(/\s/g, "").length) mesajSatiri += " | " + ss[m].replace(/\s+$/, "");
    iptal("Sign-on kabul edilmedi — ekran mesaji:" + mesajSatiri);
}
log("UYARI: sign-on sonrasi ekran taninmadi (ana menu degil). Ilk satir: " + ekran()[0]);
log("SONUC=IPTAL"); logKapat(); WScript.Quit(2);

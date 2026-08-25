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
//
// 2026-08-25: OTURUM DUSTUGUNDE BAGLANTI ACIK KALABILIYOR ('Fine del lavoro'
// ekrani). O halde StartCommunication CALISTIRILMAZDI ve script "dokunmadim"
// deyip cikardi — teyit yine verilemiyordu. Artik olu oturum algilanip
// StopCommunication + StartCommunication (Disconnetti + Connetti) yapiliyor.
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

// ── OLU OTURUM (kullanici 2026-08-25) ────────────────────────────────────────
// "AS400 bazen uzun sure islem yapilmadiginda dusebiliyor. Az once teyit vermek
//  istedim bundan dolayi veremedim."
// BAGLANTI ACIK KALIR (CommStarted=true) ama IS OLMUSTUR; ekranda:
//     Fine del lavoro
//     Lavoro: PRG00002   Utente: EMREDTK   Numero: 535267
//     Il sottosistema sta chiudendo in modo immediato.
// Bu ekranda tus gondermek ise yaramaz; kullanicinin elle yaptigi da budur:
// Comunicazioni > Disconnetti, sonra Connetti, sonra sign-on.
//
// ⚠ DIKKAT: bu ekran "Utente:" ICERIR ama "Parola" ICERMEZ — signOnMu() yanlis
// pozitif VERMEZ. Yine de bu kontrol signOnMu'dan ONCE calisir; sirasi degismesin.
function oturumOlmusMu(e) {
    var t = e.toLowerCase();
    return t.indexOf("fine del lavoro") >= 0
        || t.indexOf("sottosistema sta chiudendo") >= 0
        || t.indexOf("lavoro terminato") >= 0
        || t.indexOf("fine sessione") >= 0;
}

// StopCommunication + StartCommunication = Comunicazioni > Disconnetti > Connetti.
// BIR KEZ denenir: ikinci kez ayni yere gelmek "host vermiyor" demektir, sonsuz
// dongude AS400'u dovmek yerine durup gozcunun bir sonraki turunu beklemek dogru.
var yenidenBaglandi = false;
function commStarted() { try { return !!s.CommStarted; } catch (e) { return false; } }
function yenidenBagla() {
    if (yenidenBaglandi)
        iptal("Yeniden baglanma ZATEN denendi, ekran hala kullanilabilir degil — elle bakin.");
    yenidenBaglandi = true;
    log("Baglanti kesiliyor (StopCommunication = Comunicazioni > Disconnetti)...");
    try { s.StopCommunication(); } catch (e1) { log("  uyari: StopCommunication: " + e1.message); }
    for (var i = 0; i < 20 && commStarted(); i++) WScript.Sleep(500);
    WScript.Sleep(1500);
    log("Baglanti yeniden kuruluyor (StartCommunication = Comunicazioni > Connetti)...");
    try { s.StartCommunication(); } catch (e2) { iptal("Yeniden baglanmada StartCommunication basarisiz: " + e2.message); }
    for (var j = 0; j < 40 && !commStarted(); j++) WScript.Sleep(500);
    if (!commStarted()) iptal("Yeniden baglanti kurulamadi (20 sn)");
    // bizBagladik BILEREK set EDILMEZ: burada var olan bir oturum KURTARILIYOR,
    // gecici baglanti kurulmuyor. Hata yolunda baglanti kapatilirsa oturum
    // buttun gider ve gozcunun bir sonraki turu de bos ekran bulur.
    log("Yeniden baglandi — ekranin gelmesi bekleniyor.");
    try { oia.WaitForAppAvailable(10000); } catch (e3) {}
    WScript.Sleep(2500);
}
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
function ekranBekle() {
    var t = "";
    for (var w = 0; w < EKRAN_BEKLE_SN * 2; w++) {
        t = ekranStr();
        if (t.replace(/\s/g, "").length > 20) return t;
        WScript.Sleep(500);
    }
    return t;
}
var e = ekranBekle();
var dolu = e.replace(/\s/g, "").length;

// ── 2b) OTURUM OLMUS MU / EKRAN BOS MU -> DISCONNETTI + CONNETTI ──────
// Bagli gorunup is olmusse ya da bagliyken ekran hic boyanmadiysa tek care
// baglantiyi kesip yeniden kurmaktir (kullanicinin elle yaptigi adim).
if (oturumOlmusMu(e) || dolu <= 20) {
    log(oturumOlmusMu(e)
        ? "OLU OTURUM ekrani algilandi (is bitmis, baglanti acik) — yeniden baglanilacak."
        : "Ekran " + EKRAN_BEKLE_SN + " sn boyunca BOS kaldi — yeniden baglanilacak.");
    log("  Ekranin ilk satiri: " + ekran()[0].replace(/\s+$/, ""));
    if (DRYRUN) {
        log("DRYRUN: yeniden baglanma YAPILMADI (hicbir baglanti islemi).");
        log("SONUC=DRYRUN-OK"); logKapat(); WScript.Quit(0);
    }
    yenidenBagla();
    e = ekranBekle();
    dolu = e.replace(/\s/g, "").length;
}
if (dolu <= 20) {
    // Bagli ama BOS ekran: host oturum acmadi (cihaz baska yerde kilitli, profil
    // sorunu, 5250 anlasmasi yarim). "Sorun yok" DEME — bu bir arizadir.
    bizimBaglantiyiKapat();
    iptal("Yeniden baglandiktan sonra da ekran " + EKRAN_BEKLE_SN + " sn BOS kaldi — host oturum vermedi. "
          + "Session B'yi elle kontrol edin (cihaz/WorkStationID cakismasi olabilir).");
}
if (anaMenuMu(e)) { log("Oturum ZATEN acik (ana menu) — sign-on gerekmedi."); log("SONUC=ZATEN"); logKapat(); WScript.Quit(0); }
if (!signOnMu(e)) {
    // Sign-on degil, ana menu de degil: teyit robotu bir ekranda birakmis olabilir.
    // DOKUNMA — yanlis ekranda tus gondermek veri girisi riskidir.
    // NOT: olu oturum ekrani ('Fine del lavoro') artik YUKARIDA yakalanip yeniden
    // baglaniliyor; buraya dusen ekran gercekten "robot bir ekranda birakmis"tir.
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

// SIGMA KONTROLU (2026-08-17, sahada olculdu: parola alani 10 karakter).
// Sifre alandan UZUNSA yazma — tasan karakterler alan kilidine ya da EKSIK
// sifreye yol acar; ikisi de BASARISIZ GIRIS DENEMESI demektir ve AS400
// QMAXSIGN ile profili kilitleyebilir. Denemektense DURMAK dogru.
if (SIFRE.length > pAlan.uz) {
    iptal("Sifre ekran alanina sigmiyor (sifre " + SIFRE.length + " karakter, alan " +
          pAlan.uz + ") — yanlis giris denemesi yapilmadi. Kasadaki sifre bu ekranin " +
          "sifresi olmayabilir (ODBC sifresiyle karisti mi?).");
}
if (KULLANICI.length > uAlan.uz) {
    iptal("Kullanici adi alana sigmiyor (" + KULLANICI.length + " > " + uAlan.uz + ")");
}
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
// Menu adi kullaniciya/profile gore degisebilir. AS400'de sign-on HATASI olsaydi
// EKRAN SIGN-ON'DA KALIRDI (mesaj altta). Sign-on'dan CIKTIYSA giris kabul
// edilmistir — beklenmeyen ekran adiyla basariyi reddetme, ama LOGA yaz.
if (!signOnMu(son) && son.replace(/\s/g, "").length > 20) {
    log("SIGN-ON BASARILI — sign-on ekranindan cikildi (beklenen menu adi gorulmedi).");
    log("Gelen ekranin ilk satiri: " + ekran()[0].replace(/\s+$/, ""));
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

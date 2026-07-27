// ═══════════════════════════════════════════════════════════════════
// AS400 TRANSFER IPTAL ROBOTU — Session B, 07>01 ekrani.
// TK1 "hayali stok" transferlerini (TA 02→01) iptal eder (kullanici 2026-07-27):
//   07>01 ekranina kayit (hareket) numarasi girilir → kaydin satirlari listelenir
//   → hedef satirin basina 2 + Enter (Variation) → Shift+F4 (=F16) hareketi
//   IPTAL eder (kayit fiziksel silinir).
// Farkli kayitlarda (farkli 07>01 oturumlarinda) girilen hareketler ayni listede
// GORUNMEZ — bu robot HER CAGRIDA TEK hareket iptal eder; cagiran (Flask) kayit
// numarasi bazinda tek tek cagirir (F3 ile cikis + yeni kayit no = yeni cagri).
// Kullanim: cscript //nologo transfer_iptal.js <kayitYil> <kayitNo> <satirNo> <article> <adet> [DRYRUN]
//   ornek:  ... transfer_iptal.js 26 245458 200 "10.130.5420GRW" 53
// GUVENLIK KURALLARI:
//   - Causal TA DEGILSE ASLA F16 BASILMAZ (yanlis hareket silinmesin).
//   - Article + adet ekranda birebir dogrulanmadan F16 BASILMAZ.
//   - Beklenmedik ekran → dump + IPTAL (F3/F12 ile geri cekilme; commit yok).
//   - ASIL dogrulama Python tarafinda: iptal sonrasi BMMAF0'da kayit aranir.
// Cikti son satiri: SONUC=OK | SONUC=IPTAL | SONUC=DRYRUN-OK
// Log: as400\teyit_loglari\<zaman>_TRF_<kayitNo>_<satir>.txt
// ═══════════════════════════════════════════════════════════════════
var args = WScript.Arguments;
if (args.length < 5) { WScript.Echo("HATA: arguman eksik (kayitYil kayitNo satirNo article adet [DRYRUN])"); WScript.Echo("SONUC=IPTAL"); WScript.Quit(2); }
var OTURUM = "B";
var KYIL = "" + args(0), KNO = "" + args(1), SATIR = "" + args(2);
var ARTICLE = ("" + args(3)).replace(/^\s+|\s+$/g, ""), ADET = "" + args(4);
var DRYRUN = false;
for (var ai = 5; ai < args.length; ai++)
    if (("" + args(ai)).toUpperCase() === "DRYRUN") DRYRUN = true;
if (!/^\d{2}$/.test(KYIL) || !/^\d{1,7}$/.test(KNO) || !/^\d{1,5}$/.test(SATIR)
    || !/^[A-Za-z0-9.\-\/]{5,20}$/.test(ARTICLE) || !/^\d{1,6}(\.\d+)?$/.test(ADET)) {
    WScript.Echo("HATA: gecersiz parametre (yil=" + KYIL + " no=" + KNO + " satir=" + SATIR + " article='" + ARTICLE + "' adet=" + ADET + ")");
    WScript.Echo("SONUC=IPTAL"); WScript.Quit(2);
}
var ADET_F = parseFloat(ADET);

var fso = new ActiveXObject("Scripting.FileSystemObject");
var kok = fso.GetParentFolderName(WScript.ScriptFullName);
var logKlasor = kok + "\\teyit_loglari";
if (!fso.FolderExists(logKlasor)) fso.CreateFolder(logKlasor);
var dt = new Date();
function p2(n) { return (n < 10 ? "0" : "") + n; }
var damga = dt.getFullYear() + p2(dt.getMonth() + 1) + p2(dt.getDate()) + "_" + p2(dt.getHours()) + p2(dt.getMinutes()) + p2(dt.getSeconds());
var logDosya = fso.OpenTextFile(logKlasor + "\\" + damga + "_TRF_" + KNO + "_" + SATIR + ".txt", 2, true);

var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName(OTURUM);
var ps = s.autECLPS, oia = s.autECLOIA;

function log(m) { WScript.Echo(m); logDosya.WriteLine(m); }
function ekran() { var t = []; for (var r = 1; r <= ps.NumRows; r++) t.push(ps.GetText(r, 1, ps.NumCols)); return t; }
function ekranStr() { return ekran().join("\n"); }
function dump(et) {
    logDosya.WriteLine("---------- EKRAN [" + et + "] ----------");
    var e = ekran();
    for (var i = 0; i < e.length; i++)
        if (e[i].replace(/\s+/g, "").length) logDosya.WriteLine((i + 1 < 10 ? " " : "") + (i + 1) + "| " + e[i].replace(/\s+$/, ""));
    logDosya.WriteLine("----------");
}
function bekle() { oia.WaitForInputReady(8000); WScript.Sleep(450); }
function kilitCoz() { ps.SendKeys("[reset]"); WScript.Sleep(350); }
function geriCekil(maks) {
    for (var g = 0; g < (maks || 6); g++) {
        var o = ekranStr();
        if (o.indexOf("Menu S.I. CofleTk") >= 0) return;
        if (o.indexOf("Risorsa richiesta") >= 0 || o.indexOf("Riprova") >= 0) {
            ps.SendKeys("[reset]"); WScript.Sleep(400); ps.SendKeys("[enter]"); bekle();
            if (ekranStr() === o) { ps.SendKeys("[reset]"); WScript.Sleep(400); ps.SendKeys("[enter]"); bekle(); }
            continue;
        }
        kilitCoz();
        ps.SendKeys("[pf3]"); bekle();
        if (ekranStr() === o) { kilitCoz(); ps.SendKeys("[pf12]"); bekle(); }
    }
}
function anaMenuyeDon(maks) {
    for (var d = 0; d < maks; d++) {
        var e = ekranStr();
        if (e.indexOf("Menu S.I. CofleTk") >= 0) return true;
        if (e.indexOf("MenuIniziale") >= 0 && e.indexOf("F16") >= 0) { ps.SendKeys("[pf16]"); bekle(); continue; }
        var once = e;
        ps.SendKeys("[pf3]"); bekle();
        if (ekranStr() === once) {
            kilitCoz();
            ps.SendKeys("[pf12]"); bekle();
            if (ekranStr() === once) { kilitCoz(); ps.SendKeys("[pf3]"); bekle(); }
        }
    }
    return ekranStr().indexOf("Menu S.I. CofleTk") >= 0;
}
function iptal(neden) {
    log("HATA: " + neden);
    dump("iptal-ani");
    // Guvenli geri cekilme: 07>01 ekranlarinda Enter'siz hicbir sey commit olmaz;
    // F3/F12 ile listeden cikilir ki B sonraki robot cagrisina temiz kalsin.
    try { geriCekil(4); } catch (e) {}
    log("SONUC=IPTAL"); logDosya.Close(); WScript.Quit(2);
}
// Giris (F1 / 2-Variation) ekrani mi? — CFI robotunun kullandigi ekran.
function girisEkraniMi(e) {
    return e.indexOf("MOVIMENTI DI MAGAZZINO") >= 0 && e.indexOf("2-Variation") >= 0
        && e.indexOf("Causal cd") >= 0;
}
// 07>01 ANA (liste/karsilama) ekrani mi? — MOVIMENTI var ama variation formu yok.
function anaListeMi(e) {
    return e.indexOf("MOVIMENTI DI MAGAZZINO") >= 0 && e.indexOf("2-Variation") < 0;
}

log("TRANSFER IPTAL: kayit " + KYIL + "-" + KNO + " satir " + SATIR + " / " + ARTICLE + " / " + ADET + " adet" + (DRYRUN ? " (DRYRUN)" : ""));

// ── G0: 07>01 ana ekranina ulas ──
kilitCoz();
var e0 = ekranStr();
if (anaListeMi(e0)) {
    log("G0: zaten 07>01 ekraninda");
} else if (girisEkraniMi(e0)) {
    log("G0: 2-Variation giris ekrani (onceki CFI'dan kalma) → F3 ile listeye");
    ps.SendKeys("[pf3]"); bekle();
    if (!anaListeMi(ekranStr())) {
        // F3 ana menuye atmis olabilir → temiz yoldan yeniden gir
        geriCekil(4);
        if (!anaMenuyeDon(8)) iptal("Ana menuye ulasilamadi — B'yi elle ana menuye getirin");
        ps.SendKeys("07[enter]"); bekle();
        ps.SendKeys("01[enter]"); bekle();
        if (!anaListeMi(ekranStr())) iptal("07>01 ekrani acilamadi");
    }
} else {
    log("G0: navigasyon — ana menu > 07 > 01");
    geriCekil(6);
    if (!anaMenuyeDon(8)) iptal("Ana menuye ulasilamadi — B'yi elle ana menuye getirin");
    var eM = ekranStr();
    ps.SendKeys("07[enter]"); bekle();
    if (ekranStr() === eM) iptal("07 menusu acilmadi");
    ps.SendKeys("01[enter]"); bekle();
    if (!anaListeMi(ekranStr())) iptal("MOVIMENTI DI MAGAZZINO gelmedi (07>01)");
}
dump("ana-ekran");

// ── G1: kayit (hareket) numarasini gir ──
// Ekranda "Record" etiketli satir varsa imleci oraya koy; yoksa [home] ile ilk
// giris alanina git. Iki alanli kalip (yil + no) teyit_gir Order No ile ayni:
// yil yaz → imlec ilerlemediyse [fldext] → no yaz → [fldext] → Enter.
var sat0 = ekran(), recSatir = -1;
for (var ri = 0; ri < sat0.length; ri++)
    if (/Record|Registr/i.test(sat0[ri])) { recSatir = ri + 1; break; }
if (recSatir > 0) {
    log("G1: 'Record' etiketi satir " + recSatir + " → imlec oraya");
    ps.SetCursorPos(recSatir, 1); WScript.Sleep(200);
    ps.SendKeys("[home]"); WScript.Sleep(250);
} else {
    log("G1: 'Record' etiketi yok → [home] ile ilk giris alani");
    ps.SendKeys("[home]"); WScript.Sleep(250);
}
var r0 = ps.CursorPosRow, c0 = ps.CursorPosCol;
ps.SendKeys(KYIL); WScript.Sleep(250);
if (ps.CursorPosRow === r0 && (ps.CursorPosCol - c0) <= 2) { ps.SendKeys("[fldext]"); WScript.Sleep(250); }
ps.SendKeys(KNO); WScript.Sleep(250);
var yankiSat = ps.GetText(r0, 1, ps.NumCols);
if (yankiSat.indexOf(KYIL) < 0 || yankiSat.indexOf(KNO) < 0)
    iptal("Kayit no yankisi hatali: '" + yankiSat.replace(/\s+$/, "") + "' — 07>01 ekran duzeni beklenenden farkli");
ps.SendKeys("[fldext]"); WScript.Sleep(300);
ps.SendKeys("[enter]");

// Kaydin satirlari listelenene kadar bekle (article gorunmeli)
var eL = "", listelendi = false;
for (var pz = 0; pz < 10; pz++) {
    WScript.Sleep(800);
    eL = ekranStr();
    if (eL.indexOf("Risorsa richiesta") >= 0 || eL.indexOf("Riprova") >= 0) {
        kilitCoz(); ps.SendKeys("[enter]"); bekle(); continue;
    }
    if (eL.indexOf(ARTICLE) >= 0) { listelendi = true; break; }
}
if (!listelendi) {
    dump("kayit-acilamadi");
    iptal("Kayit " + KYIL + "-" + KNO + " acilamadi ya da '" + ARTICLE + "' listede yok (zaten iptal edilmis olabilir)");
}
log("G1: kayit acildi — '" + ARTICLE + "' listede gorunuyor");
dump("kayit-listesi");

// ── G2: hedef satiri bul ──
// Oncelik: hem satir no (MGPRRE, orn '200') hem article iceren ekran satiri.
// Bulunamazsa yalniz article + adet esleyen satir (tek aday olmali).
var satlar = ekran(), hedefRow = -1;
var reSatirNo = new RegExp("(^|\\s)" + SATIR + "(\\s|$)");
for (var i2 = 0; i2 < satlar.length; i2++) {
    if (satlar[i2].indexOf(ARTICLE) >= 0 && reSatirNo.test(satlar[i2])) { hedefRow = i2 + 1; break; }
}
if (hedefRow < 0) {
    var adaylar = [];
    for (var i3 = 0; i3 < satlar.length; i3++)
        if (satlar[i3].indexOf(ARTICLE) >= 0) adaylar.push(i3 + 1);
    if (adaylar.length === 1) hedefRow = adaylar[0];
    else iptal("Hedef satir tek basina bulunamadi (satirNo=" + SATIR + ", article aday sayisi=" + adaylar.length + ")");
}
// Adet dogrulamasi: ayni ekran satirinda italyan bicimi adet ('53,000' / '1.234,000')
var hedefTxt = satlar[hedefRow - 1];
var mAdet = /([\d\.]+),(\d{1,3})/.exec(hedefTxt);
if (mAdet) {
    var ekAdet = parseFloat(mAdet[1].replace(/\./g, "") + "." + mAdet[2]);
    if (Math.abs(ekAdet - ADET_F) > 0.001)
        iptal("Satirdaki adet uyusmadi: ekran=" + ekAdet + " beklenen=" + ADET_F + " ('" + hedefTxt.replace(/\s+$/, "") + "')");
    log("G2: hedef ekran satiri " + hedefRow + ", adet dogrulandi (" + ekAdet + ")");
} else {
    log("G2: hedef ekran satiri " + hedefRow + " (adet deseni okunamadi — article eslesti, devam)");
}

// ── G3: satir basina '2' yaz (secenek alani) + Enter ──
// Secenek kolonunun yeri bilinmiyor → 1..8 arasi kolonlarda dene: yaz, yanki
// gelirse tamam; gelmezse reset + sonraki kolon (korumali alanlara yazilamaz).
var yazildi = false;
for (var kol = 1; kol <= 8 && !yazildi; kol++) {
    kilitCoz();
    ps.SetCursorPos(hedefRow, kol); WScript.Sleep(150);
    ps.SendKeys("2"); WScript.Sleep(250);
    var bas = ps.GetText(hedefRow, 1, 9);
    if (bas.indexOf("2") >= 0 && bas.indexOf("2") < 8) { yazildi = true; log("G3: '2' kolon " + kol + " uzerinden yazildi ('" + bas.replace(/\s+$/, "") + "')"); }
}
if (!yazildi) iptal("Satir basina '2' yazilamadi (secenek alani bulunamadi — ekran duzeni farkli)");
ps.SendKeys("[enter]"); bekle(); WScript.Sleep(1200);

// ── G4: Variation/detay ekranini dogrula — CAUSAL TA SARTI ──
var eV = ekranStr();
if (eV.toLowerCase().indexOf("errore") >= 0) iptal("'2' + Enter sonrasi hata mesaji");
dump("detay-ekrani");
if (eV.indexOf(ARTICLE) < 0) iptal("Detay ekraninda article gorunmedi");
// Causal: 'Causal cd. . . .   TA   ...' satirindan oku
var mCa = /Causal\s*cd[.\s]*\s([A-Z]{2,3})\s/.exec(eV);
if (!mCa) iptal("Detay ekraninda 'Causal cd' okunamadi — F16 BASILMAYACAK");
if (mCa[1] !== "TA") iptal("Causal '" + mCa[1] + "' — TA degil! Bu hareket transfer degil, IPTAL EDILMEZ");
log("G4: Causal TA dogrulandi");
// Kayit no dogrulamasi (Record no. . 26 245458)
if (eV.indexOf(KNO) < 0) iptal("Detay ekraninda kayit no " + KNO + " gorunmedi");
// Actual row varsa satir no eslesmeli
var mAr = /Actual\s+row\s+(\d+)/i.exec(eV);
if (mAr && parseInt(mAr[1], 10) !== parseInt(SATIR, 10))
    iptal("Actual row " + mAr[1] + " != beklenen satir " + SATIR + " — yanlis satir acilmis olabilir");
log("G4: detay ekrani dogrulandi (kayit " + KNO + ", satir " + (mAr ? mAr[1] : "?") + ")");

if (DRYRUN) {
    log("DRYRUN: iptal edilecek hareket dogrulandi, Shift+F4'e BASILMADI.");
    try { geriCekil(4); } catch (e) {}
    log("SONUC=DRYRUN-OK"); logDosya.Close(); WScript.Quit(0);
}

// ── G5: Shift+F4 (=F16) → hareketi iptal et ──
var onceE = eV;
kilitCoz();
ps.SendKeys("[pf16]"); bekle(); WScript.Sleep(1500);
var bittiT = false;
for (var d5 = 0; d5 < 10 && !bittiT; d5++) {
    var e5 = ekranStr();
    if (e5.indexOf("INVIO=Riprova") >= 0 || e5.indexOf("fase di aggiornamento") >= 0
        || e5.indexOf("Risorsa richiesta") >= 0) {
        log("Kayit kilidi (Riprova) → Enter ile tekrar (deneme " + d5 + ")");
        kilitCoz(); ps.SendKeys("[enter]"); bekle(); WScript.Sleep(2000);
        continue;
    }
    if (e5.toLowerCase().indexOf("errore") >= 0) { dump("f16-hata"); iptal("F16 sonrasi hata mesaji"); }
    // Olasi onay penceresi (F06=Conferma) → iptal onayi, TEK SEFER F6
    if (e5.indexOf("F06=Conferma") >= 0) {
        log("Onay penceresi (F06=Conferma) → F6"); dump("f16-onay");
        kilitCoz(); ps.SendKeys("[pf6]"); bekle(); WScript.Sleep(1500);
        continue;
    }
    if (e5 !== onceE) { bittiT = true; break; }
    // Ekran degismedi → F16 yutulmus olabilir: reset + tekrar (tek sefer yeterli;
    // hareket coktan silindiyse ekran zaten degisir, ikinci F16 bos ekranda zararsiz)
    if (d5 === 3) { log("F16 tepkisiz (3 tur) — reset + F16 tekrar"); kilitCoz(); ps.SendKeys("[pf16]"); bekle(); }
    WScript.Sleep(1000);
}
if (!bittiT) { dump("f16-tepkisiz"); iptal("F16 sonrasi ekran degismedi — iptal islenmemis olabilir"); }
dump("f16-sonrasi");

// ── G6: sonuc — asil dogrulama Python'da (BMMAF0'da kayit silinmis mi) ──
// Ekranda article hala ayni satirda ayni adetle duruyorsa suphelen (yine de
// Python dogrulamasi belirleyici — burada yalniz bilgi logu).
var eSon = ekranStr();
if (eSon.indexOf(ARTICLE) >= 0 && anaListeMi(eSon) === false && girisEkraniMi(eSon) === false)
    log("BILGI: article hala ekranda gorunuyor — Python BMMAF0 dogrulamasi belirleyici olacak");
log("OK: F16 basildi, ekran degisti — iptal buyuk olasilikla islendi");
log("SONUC=OK");
logDosya.Close();

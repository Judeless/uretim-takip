// ═══════════════════════════════════════════════════════════════════
// AS400 TEYIT YAZICISI (uretim surumu) — Session B uzerinden tek teyit girer.
// Kullanim: C:\Windows\SysWOW64\cscript.exe //nologo teyit_gir.js <yil> <launchNo> <article> <adet>
//   ornek:  ... teyit_gir.js 26 140672 "10.300.1866-10W" 5
// Kurallar:
//   - YALNIZ A bayragi (launch acik kalir). S/kapatma bu scriptte YOKTUR (insan isi).
//   - Her adimda ekran dogrulanir; beklenmedik durumda reset+F3 ile cekilip IPTAL.
//   - Cikti son satiri: SONUC=OK | SONUC=IPTAL  (Flask bunu parse eder)
//   - Tam ekran dokumleri: as400\teyit_loglari\<zaman>_<launch>.txt
// 5250 dersleri (pilotta ogrenildi):
//   - Field Exit mnemoni [fldext] ([fieldexit] bu PCOMM 5.8'de GECERSIZ)
//   - Sayisal alan tam dolunca AID tusu operator-error kilidi yapar (OIA=5);
//     recete: [reset]→[fldext]→[reset]
// ═══════════════════════════════════════════════════════════════════
var args = WScript.Arguments;
if (args.length < 4) { WScript.Echo("HATA: arguman eksik (yil launchNo article adet [A|S])"); WScript.Echo("SONUC=IPTAL"); WScript.Quit(2); }
var OTURUM = "B";
var YIL = "" + args(0), NO = "" + args(1), ARTICLE = "" + args(2), ADET = "" + args(3);
// 5. arg: bayrak. A = devam (launch acik kalir), S = KAPAT (launch tamamlandi).
// Varsayilan A (guvenli). S yalnizca cagirandan acikca gelirse.
var BAYRAK = (args.length >= 5 && ("" + args(4)).toUpperCase() === "S") ? "S" : "A";
// 6. arg 'DRYRUN' → adet+bayrak yazilir, dogrulanir ama F6 BASILMAZ (IPTAL edilir).
// Guvenli test icin (ozellikle S mekanigini commit etmeden dogrulamak). Flask ASLA gondermez.
var DRYRUN = (args.length >= 6 && ("" + args(5)).toUpperCase() === "DRYRUN");
if (!/^\d{2}$/.test(YIL) || !/^\d{1,7}$/.test(NO) || !/^\d{1,5}$/.test(ADET) || parseInt(ADET, 10) <= 0) {
    WScript.Echo("HATA: gecersiz parametre (yil=" + YIL + " no=" + NO + " adet=" + ADET + ")");
    WScript.Echo("SONUC=IPTAL"); WScript.Quit(2);
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
var kok = fso.GetParentFolderName(WScript.ScriptFullName);
var logKlasor = kok + "\\teyit_loglari";
if (!fso.FolderExists(logKlasor)) fso.CreateFolder(logKlasor);
var dt = new Date();
function p2(n) { return (n < 10 ? "0" : "") + n; }
var damga = dt.getFullYear() + p2(dt.getMonth() + 1) + p2(dt.getDate()) + "_" + p2(dt.getHours()) + p2(dt.getMinutes()) + p2(dt.getSeconds());
var logDosya = fso.OpenTextFile(logKlasor + "\\" + damga + "_" + NO + ".txt", 2, true);

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
// KILIT COZME: SADECE reset. fldext KULLANMA — bir GIRIS ALANINDAYKEN (orn. onay
// penceresi 'Data movimento') fldext o alanin icerigini SILER (2026-07-20 bug: F6
// yutulunca iptal'deki fldext zorunlu alani sildi, F3 yasak → B takildi).
function kilitCoz() { ps.SendKeys("[reset]"); WScript.Sleep(350); }
// Guvenli geri cekilme: F12(Return, alt/onay ekranlari) once, degismezse F3.
// Hicbir alan silinmez; sadece ekranlardan geri cikilir (girilmemis satir commit OLMAZ).
function geriCekil(maks) {
    for (var g = 0; g < (maks || 6); g++) {
        var o = ekranStr();
        if (o.indexOf("Menu S.I. CofleTk") >= 0) return;
        // KAYIT KILIDI ekrani ("Risorsa richiesta" / INVIO=Riprova) → reset + Enter.
        // (kullanici 2026-07-20 sahada: Ctrl(reset)+Enter, gerekirse 2 kez.)
        if (o.indexOf("Risorsa richiesta") >= 0 || o.indexOf("Riprova") >= 0) {
            ps.SendKeys("[reset]"); WScript.Sleep(400); ps.SendKeys("[enter]"); bekle();
            if (ekranStr() === o) { ps.SendKeys("[reset]"); WScript.Sleep(400); ps.SendKeys("[enter]"); bekle(); }
            continue;
        }
        // F3=Exit onceki (Re-entry summary/cogu ekran) — 'In sistem' yarim satir
        // F3'te commit-siz terk edilir. Degismezse F12=Return fallback.
        kilitCoz();
        ps.SendKeys("[pf3]"); bekle();
        if (ekranStr() === o) { kilitCoz(); ps.SendKeys("[pf12]"); bekle(); }
    }
}
// Iptal cikisi: ANA MENUYE SURUNME (her F3'te X/hata — cok yavas, kullanici
// 2026-07-21). Limits/Summary orders/TP41 YETERLI: sonraki robot G0'da oradan
// devam eder. Summary materials'tan F12=Return, digerlerinden F3.
function guvenliBirak() {
    for (var g = 0; g < 5; g++) {
        var e = ekranStr();
        if (e.indexOf("Limits request") >= 0 || e.indexOf("Summary orders") >= 0
            || e.indexOf("Production Definitive Data") >= 0 || e.indexOf("Menu S.I. CofleTk") >= 0) return;
        if (e.indexOf("Risorsa richiesta") >= 0 || e.indexOf("Riprova") >= 0) {
            ps.SendKeys("[reset]"); WScript.Sleep(400); ps.SendKeys("[enter]"); bekle(); continue;
        }
        var once = e;
        if (e.indexOf("Summary materials") >= 0) { kilitCoz(); ps.SendKeys("[pf12]"); bekle(); }
        else {
            kilitCoz(); ps.SendKeys("[pf3]"); bekle();
            if (ekranStr() === once) { kilitCoz(); ps.SendKeys("[pf12]"); bekle(); }
        }
    }
}
function iptal(neden) {
    log("HATA: " + neden);
    dump("iptal-ani");
    try { guvenliBirak(); } catch (e) {}
    log("SONUC=IPTAL"); logDosya.Close(); WScript.Quit(2);
}
function bosluksuz(t) { return t.replace(/\s+/g, ""); }
function anaMenuyeDon(maks) {
    // Bazi ekranlar F3 degil F12 ister (orn. malzeme ozeti: F12=Return).
    // F3 ekrani degistirmezse kilit cozup F12 dene.
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

log("TEYIT: launch " + YIL + "-" + NO + " / " + ARTICLE + " / " + ADET + " adet / bayrak " + BAYRAK +
    (BAYRAK === "S" ? "  (!! LAUNCH KAPATILACAK)" : ""));

// G0-G1: Rientro Limits ekranina git.
// OPTIMIZASYON (kullanici 2026-07-20, 2 iterasyon): basarili F6 sonrasi sistem
// "Summary orders" LISTESINE doner (TP41 menusune degil!). En basa donme (F3
// spam) her adimda hata verip cok uzun suruyordu. Simdi ekrana gore kisa yol:
//   - Limits request       → zaten giris ekranindayiz, direkt devam
//   - Summary orders       → F3 ile bir adim geri → Limits/TP41'den devam
//   - TP41 (Definitive)    → direkt 02
//   - baska her sey        → eski tam yol (temizle + ana menu + 10>05>12>02)
function tamYolLimits() {
    geriCekil(6);
    if (!anaMenuyeDon(8)) iptal("Ana menuye ulasilamadi — B'yi elle ana menuye getirin");
    var adimlar = [["10", "TP00"], ["05", "TP40"], ["12", "TP41"]];
    for (var a = 0; a < adimlar.length; a++) {
        ps.SendKeys(adimlar[a][0] + "[enter]"); bekle();
        if (ekranStr().indexOf(adimlar[a][1]) < 0) iptal("Menu " + adimlar[a][0] + " -> " + adimlar[a][1] + " gelmedi");
    }
    ps.SendKeys("02[enter]"); bekle();
}
kilitCoz();
var e0 = ekranStr();
if (e0.indexOf("RIENTRO ORDINI PRODUZIONE") >= 0 && e0.indexOf("Limits request") >= 0) {
    log("G0: zaten Limits ekraninda → direkt devam");
} else if (e0.indexOf("Summary orders") >= 0) {
    log("G0: Summary orders listesi (onceki teyit sonrasi) → F3 ile geri");
    ps.SendKeys("[pf3]"); bekle();
    kilitCoz();
    var e1 = ekranStr();
    if (e1.indexOf("Limits request") >= 0) {
        log("G0: Limits ekranina donuldu");
    } else if (e1.indexOf("Production Definitive Data") >= 0) {
        log("G0: TP41'e donuldu → 02");
        ps.SendKeys("02[enter]"); bekle();
    } else {
        log("G0: F3 beklenmedik ekrana getirdi → tam yol");
        tamYolLimits();
    }
} else if (e0.indexOf("Production Definitive Data") >= 0 && e0.indexOf("Menu S.I. CofleTk") < 0) {
    log("G0: TP41'de → direkt 02, en basa donme YOK");
    ps.SendKeys("02[enter]"); bekle();
} else {
    // Ilk teyit / onceki iptal → temizle (record-lock/alt ekran) + ana menu + tam yol
    tamYolLimits();
}
var eL = ekranStr();
if (eL.indexOf("RIENTRO ORDINI PRODUZIONE") < 0 || eL.indexOf("Limits request") < 0) iptal("Limits ekrani gelmedi");

// G2: Order No (yil [fldext] no [fldext] enter)
ps.SendKeys("[home]"); WScript.Sleep(250);
var r0 = ps.CursorPosRow, c0 = ps.CursorPosCol;
ps.SendKeys(YIL); WScript.Sleep(250);
if (ps.CursorPosRow === r0 && (ps.CursorPosCol - c0) <= 2) { ps.SendKeys("[fldext]"); WScript.Sleep(250); }
ps.SendKeys(NO); WScript.Sleep(250);
var satirTxt = ps.GetText(r0, 1, ps.NumCols);
if (satirTxt.indexOf(YIL) < 0 || satirTxt.indexOf(NO) < 0) iptal("Order No yankisi hatali: '" + satirTxt.replace(/\s+$/, "") + "'");
ps.SendKeys("[fldext]"); WScript.Sleep(300);
ps.SendKeys("[enter]");
var eS = "";
for (var pz = 0; pz < 8; pz++) { WScript.Sleep(600); eS = ekranStr(); if (eS.indexOf("Re-entry summary") >= 0) break; }
if (eS.indexOf("Re-entry summary") < 0) iptal("Re-entry summary gelmedi (launch bulunamadi/kapali olabilir)");
if (eS.indexOf(NO) < 0) iptal("Launch no ozet ekraninda yok");
if (bosluksuz(eS).indexOf(bosluksuz(ARTICLE)) < 0) iptal("ARTICLE eslesmedi! Beklenen: " + ARTICLE);
// GUVENLIK: bu launch'ta kaydedilmemis ("In sistem.") yarim satir varsa DUR.
// Onceki yarim kalmis bir giris demektir — ustune yazmak mukerrer/yanlis olur.
if (eS.indexOf("In sistem") >= 0)
    iptal("Bu launch'ta kaydedilmemis satir var (Session B'de yarim kalmis giris). Once B'de kontrol edin.");
log("OK: ozet ekrani — launch + article dogrulandi, yarim satir yok");
dump("ozet");

// G3: (kosullu F10) + adet + [fldext] + bayrak (A/S)
// ONEMLI (kullanici 2026-07-17): launch'a HIC teyit verilmemisse bos ilk satir
// (Prg 100) ZATEN HAZIR — F10 BASMA (F10 fazladan bos satir acar, sonraki Enter'i
// 'Quantity obligatory' ile bozar → malzeme ekrani gelmez). Yalniz onceden teyit
// varsa (dolu satirlar) yeni satir icin F10 gerekir.
var _sat = ekran(), _bas = -1, oncekiTeyit = 0;
for (var _i = 0; _i < _sat.length; _i++) if (_sat[_i].indexOf("War.Rec.No") >= 0) { _bas = _i; break; }
if (_bas >= 0) for (var _j = _bas + 1; _j < _sat.length; _j++) {
    if (_sat[_j].indexOf("Managem") >= 0 || _sat[_j].indexOf("F03=Exit") >= 0) break;
    if (/\d{2}\s+\d{7}/.test(_sat[_j])) oncekiTeyit++;   // War.Rec.No dolu = kayitli teyit
}
log("Onceki teyit satiri: " + oncekiTeyit + (oncekiTeyit ? " → F10 (yeni satir)" : " → ILK teyit, F10 YOK"));
if (oncekiTeyit > 0) {
    ps.SendKeys("[pf10]"); bekle();
    if (ekranStr().indexOf("Re-entry summary") < 0) iptal("F10 sonrasi ekran degisti");
}
var qr = ps.CursorPosRow;
ps.SendKeys(ADET); WScript.Sleep(250);
ps.SendKeys("[fldext]"); WScript.Sleep(400);
var satirQ = ps.GetText(qr, 1, ps.NumCols).replace(/\s+$/, "");
if (satirQ.indexOf(ADET) < 0) iptal("Adet yankisi yok (imlec yanlis alanda olabilir): '" + satirQ + "'");
var mevcut = satirQ.charAt(satirQ.length - 1);
if (mevcut !== "A" && mevcut !== "S") iptal("Satir sonunda A/S bayragi yok: '" + satirQ + "'");
if (mevcut !== BAYRAK) {
    // A/S alanina imleci koy (satirin son non-space kolonu) ve istenen bayragi yaz.
    var asCol = satirQ.length;
    ps.SetCursorPos(qr, asCol); WScript.Sleep(200);
    ps.SendKeys(BAYRAK); WScript.Sleep(350);
    satirQ = ps.GetText(qr, 1, ps.NumCols).replace(/\s+$/, "");
    if (satirQ.charAt(satirQ.length - 1) !== BAYRAK)
        iptal("Bayrak '" + BAYRAK + "' yazilamadi: '" + satirQ + "'");
}
log("OK: adet " + ADET + " girildi, bayrak " + BAYRAK + ": '" + satirQ + "'");
dump("adet");

// DRYRUN: BURADA dur — adet+bayrak yazildi. Enter'lara BASMA, onay penceresine
// GIRME (o ekran zorunlu-alanli, cikisi zor; iptal orada takilir). Ozet
// ekranindan F3 ile guvenle geri cikilir (girilmemis satir commit OLMAZ).
// Onay penceresi + F6 zaten canli kanitli (2757W-10, 1866-10W).
if (DRYRUN) {
    log("DRYRUN: adet+bayrak yazildi (onay penceresine GIRILMEDI), F3 ile iptal (commit YOK).");
    try { geriCekil(6); anaMenuyeDon(4); } catch (e) {}
    log("SONUC=DRYRUN-OK");
    logDosya.Close();
    WScript.Quit(0);
}

// G4: EKRAN-TEPKILI teyit dizisi (2026-07-20 — kullanicinin ogrettigi tum durumlar).
// adet+bayrak yazildi. Simdi ilk Enter'i basip, gelen HER ekrana gore tepki ver.
// ONCELIK SIRASI KRITIK: onay penceresi (F06=Conferma) F23 kontrolunden ONCE gelir.
// Cunku F23-force onay penceresini acar AMA alt satirdaki "F23 to force" metni
// EKRANDA KALIR (overlay); F23'e once bakilirsa robot durmadan F23 spam eder
// (144724 vakasi: F23 x6 → 12-adim guard → iptal). Cozum: F23 SADECE onay
// penceresi icinde, TEK SEFER (f23Basildi bayragi), sonra F6.
// Data/Tipo/Costo alanlarina ASLA dokunulmaz (ERP otomatik doldurur).
ps.SendKeys("[enter]"); bekle(); WScript.Sleep(1800);
var confDogrulandi = false, f6Sayisi = 0, f23Basildi = false, sonEkran = "";
var bitti = false, duraklama = 0;
for (var adim = 0; adim < 20 && !bitti; adim++) {
    var e = ekranStr();
    // (0) KAYIT KILIDI / dosya guncelleniyor (2026-07-23: "BPRCF2 in fase di
    // aggiornamento / INVIO=Riprova" — 3852W/4473W iptalleri). Enter = Riprova;
    // kilit birkac saniye surebilir → sabirla tekrar, ayni-ekran sayacini sifirla.
    if (e.indexOf("INVIO=Riprova") >= 0 || e.indexOf("fase di aggiornamento") >= 0
        || e.indexOf("Risorsa richiesta") >= 0) {
        log("Kayit kilidi (Riprova) → Enter ile tekrar (adim " + adim + ")");
        kilitCoz(); ps.SendKeys("[enter]"); bekle(); WScript.Sleep(2500);
        sonEkran = ""; duraklama = 0; continue;
    }
    if (e.toLowerCase().indexOf("errore") >= 0 || e.toLowerCase().indexOf(" error") >= 0)
        iptal("Hata mesaji (adim " + adim + ")");

    // (1) CHECK QUADRATURE denge penceresi (ozellikle S kapatmada) → F6 (Confirm)
    if (e.indexOf("CHECK QUADRATURE") >= 0) {
        log("CHECK QUADRATURE → F6 (Confirm)"); dump("check-quadrature-adim" + adim);
        ps.SendKeys("[pf6]"); bekle(); WScript.Sleep(1800); f6Sayisi++; sonEkran = e; continue;
    }
    // (2) Onay penceresi (F06=Conferma). Stok yetersizse ONCE F23 (TEK SEFER), sonra F6.
    if (e.indexOf("F06=Conferma") >= 0) {
        if ((e.indexOf("F23 to force") >= 0 || e.indexOf("higher than the q.ty in stock") >= 0) && !f23Basildi) {
            log("Stok yetersiz → F23 (force, tek sefer)"); dump("f23-force-adim" + adim);
            ps.SendKeys("[pf23]"); bekle(); WScript.Sleep(1800); f23Basildi = true; sonEkran = e; continue;
        }
        if (e.indexOf("Data movimento") >= 0 && /Data movimento[^\d\n]*\d/.test(e) === false)
            iptal("Data movimento BOS — F6 reddedilir, teyit ISLENMEZ");
        dump("onay-penceresi-adim" + adim);
        var onceki = e;
        // F6 arada YUTULUYOR (2026-07-21 sabah batch: 4725W/6175W "ekran degismedi"
        // iptalleri — klavye kilidi/gec host). Once reset, sonra F6 + 8sn'e kadar
        // POLL; hala degismediyse reset+F6 BIR kez daha (8sn bekleme cift-commit
        // riskini onler: host normalde <2sn'de yanit verir).
        kilitCoz();
        ps.SendKeys("[pf6]"); bekle();
        var degisti = false;
        for (var b6 = 0; b6 < 8 && !degisti; b6++) { WScript.Sleep(1000); degisti = (ekranStr() !== onceki); }
        if (!degisti) {
            log("F6 tepkisiz (8sn) — reset + F6 tekrar");
            kilitCoz(); ps.SendKeys("[pf6]"); bekle();
            for (var b7 = 0; b7 < 6 && !degisti; b7++) { WScript.Sleep(1000); degisti = (ekranStr() !== onceki); }
        }
        if (!degisti) iptal("F6 sonrasi ekran DEGISMEDI (2 deneme, 14sn) — teyit islenmedi");
        f6Sayisi++;
        sonEkran = e; continue;
    }
    // (3) SAF Summary materials (F06 YOK). Stok yetersizse F23 (tek sefer); degilse Conf.qty + Enter.
    if (e.indexOf("Summary materials") >= 0) {
        if ((e.indexOf("F23 to force") >= 0 || e.indexOf("higher than the q.ty in stock") >= 0) && !f23Basildi) {
            log("Stok yetersiz (ozet) → F23 (force, tek sefer)"); dump("f23-ozet-adim" + adim);
            ps.SendKeys("[pf23]"); bekle(); WScript.Sleep(1800); f23Basildi = true; sonEkran = e; continue;
        }
        if (!confDogrulandi) {
            var mm = /Conf\.qty\s+([\d\.,]+)/.exec(e);
            if (mm) {
                var cq = parseFloat(mm[1].replace(/\./g, "").replace(",", "."));
                if (Math.abs(cq - parseInt(ADET, 10)) > 0.001) iptal("Conf.qty uyusmadi: " + mm[1] + " != " + ADET);
                log("OK: Conf.qty dogrulandi = " + cq);
            }
            confDogrulandi = true; dump("malzeme-ozeti");
        }
        ps.SendKeys("[enter]"); bekle(); WScript.Sleep(1800); sonEkran = e; continue;
    }
    // (4) Islem bitti: Summary orders LISTESI (F6 sonrasi normal donus ekrani,
    // kullanici fotografi 2026-07-20) / TP41 / Limits / ana menuye dondu
    if (e.indexOf("Summary orders") >= 0 || e.indexOf("Production Definitive Data") >= 0
        || e.indexOf("Limits request") >= 0 || e.indexOf("Menu S.I. CofleTk") >= 0) {
        if (f6Sayisi > 0) { log("OK: teyit islendi (F6 x" + f6Sayisi + "), islem tamam"); bitti = true; break; }
        iptal("Islem bitmis gorunuyor ama hic F6 basilmadi (adim " + adim + ")");
    }
    // (5) Ilerleme yok / beklenmeyen ekran → 2 duraklamaya kadar tolere et
    // (Enter arada yutulabiliyor — hemen iptal etme, kilit cozup tekrar dene)
    if (e === sonEkran) {
        duraklama++;
        if (duraklama >= 2) iptal("Ilerleme yok — beklenmeyen ekran (adim " + adim + ")");
        kilitCoz();
    } else { duraklama = 0; }
    sonEkran = e;
    ps.SendKeys("[enter]"); bekle(); WScript.Sleep(1500);   // bilinmeyen ara ekran → Enter dene
}
if (!bitti) iptal("20 adimda islem tamamlanamadi");

// G6: B'yi F6 sonrasi halinde (Summary orders / TP41) BIRAK — ana menuye DONME
// (kullanici: en basa donerken her adimda hata aliniyor, cok uzun suruyor).
// Sonraki teyit_gir cagrisi G0'da ekrani tanir: Summary orders→F3, TP41→02, Limits→direkt.
log("SONUC=OK");
logDosya.Close();

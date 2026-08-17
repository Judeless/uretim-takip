// ═══════════════════════════════════════════════════════════════════
// AS400 CFI DEPO GIRISI (launch'siz uretim) — Session B, 07>01>F1 ekrani.
// Kullanim: cscript //nologo cfi_gir.js <article> <adet> [DRYRUN]
//   ornek:  ... cfi_gir.js "10.130.5420GRW" 53
// Akis (kullanici 2026-07-20 ogretti; 5420GRW 53 adet elle girilerek kanitlandi):
//   07 > 01 > F1 → "2-Variation" giris ekrani.
//   Warehouse cd=01D, Causal cd=CFI, Counterpart War.cd=01D (bir kez yazilinca
//   EKRANDA KALIR), Article cd=kod (fldext GEREKMEZ), Qty=adet+[fldext]+[enter].
//   Enter sonrasi ust listede kod+adet gorunur; Article/Qty alanlari bosalir.
// Cikti son satiri: SONUC=OK | SONUC=IPTAL | SONUC=DRYRUN-OK
// Log: as400\teyit_loglari\<zaman>_CFI_<kod>.txt
// ═══════════════════════════════════════════════════════════════════
var args = WScript.Arguments;
if (args.length < 2) { WScript.Echo("HATA: arguman eksik (article adet [CFI|COP] [DRYRUN])"); WScript.Echo("SONUC=IPTAL"); WScript.Quit(2); }
var OTURUM = "B";
var ARTICLE = ("" + args(0)).replace(/^\s+|\s+$/g, "");
var ADET = "" + args(1);
// 3. arg: causal. CFI = launch'siz uretim girisi (counterpart 01D);
// COP = HURDA ayirma (kullanici 2026-07-23: counterpart BOS birakilir!).
var CAUSAL = (args.length >= 3 && ("" + args(2)).toUpperCase() !== "DRYRUN")
    ? ("" + args(2)).toUpperCase() : "CFI";
var DRYRUN = false;
for (var ai = 2; ai < args.length; ai++)
    if (("" + args(ai)).toUpperCase() === "DRYRUN") DRYRUN = true;
// ARTICLE FILTRESI (2026-07-31 duzeltmesi) — eskiden beyaz listeydi:
//   ^[A-Za-z0-9.\-\/]{5,20}$   → ALT CIZGI listede YOK.
// Sonuc: "10.300.1992A_LK" icin COP "gecersiz parametre" ile IPTAL oluyordu;
// kod ekrana hic yazilmadan robot cikiyordu. Ayni hatanin PYTHON tarafi
// 2026-07-30'da duzeltilmisti (ERP article master'inda _ , & ( ) $ * " iceren
// 318 GERCEK kod olculdu) — robot script'i atlanmisti.
// Beyaz liste DEGIL, gercek riskler:
//   [ ]           → PCOMM SendKeys mnemonik parantezi ("[enter]") — ekrana yazilmaz
//   bosluk / ASCII disi → arguman gecisi + ekran kodlamasi
//   21+ karakter  → ekran alani tasar (kuyrugu kilitleyen hata sinifi)
// Kodun ERP'de GERCEKTEN tanimli oldugu cagiran tarafta dogrulanir
// (_article_gecersiz_mi); ayrica alanYaz yazdigini ekrandan geri okuyup
// karsilastirir — yani bu filtre tek fren degil, ilk elemedir.
if (!/^[\x21-\x7E]{3,21}$/.test(ARTICLE) || /[\[\]]/.test(ARTICLE)
    || !/^\d{1,5}$/.test(ADET) || parseInt(ADET, 10) <= 0
    || (CAUSAL !== "CFI" && CAUSAL !== "COP")) {
    WScript.Echo("HATA: gecersiz parametre (article='" + ARTICLE + "' adet=" + ADET + " causal=" + CAUSAL + ")");
    WScript.Echo("SONUC=IPTAL"); WScript.Quit(2);
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
var kok = fso.GetParentFolderName(WScript.ScriptFullName);
var logKlasor = kok + "\\teyit_loglari";
if (!fso.FolderExists(logKlasor)) fso.CreateFolder(logKlasor);
var dt = new Date();
function p2(n) { return (n < 10 ? "0" : "") + n; }
var damga = dt.getFullYear() + p2(dt.getMonth() + 1) + p2(dt.getDate()) + "_" + p2(dt.getHours()) + p2(dt.getMinutes()) + p2(dt.getSeconds());
var logDosya = fso.OpenTextFile(logKlasor + "\\" + damga + "_" + CAUSAL + "_" + ARTICLE.replace(/[^A-Za-z0-9]/g, "") + ".txt", 2, true);

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
    // ── KUYRUGU KURTAR (2026-08-17, kullanici istegi) ────────────────────────
    // ESKIDEN: burada geri cekilme YOKTU; "B'yi oldugu yerde birak, kullanici
    // durumu gorsun" gerekcesiyle ekran yarim dolu KALIYORDU. Sonucu agir:
    // sonraki satirin kosusu ayni ekranda basliyor, "Article alani BOS DEGIL"
    // kontrolune carpip o da iptal oluyor ve KUYRUGUN GERI KALANI KOMPLE
    // dusuyordu (kullanici: "durup bekliyor, ilk sayfaya cikip devam etsin").
    // Launch tarafi (teyit_gir.js) bunu guvenliBirak ile zaten yapiyordu.
    // GUVENLIK: F3/F12 = Exit/Cancel; CFI satiri YALNIZ Enter ile commit olur,
    // geri cekilmek yarim satiri KAYDETMEZ. Enter basildiktan sonraki
    // dogrulama hatasinda ise satir gercekten girilmis OLABILIR — onu app
    // tarafi bagimsiz olarak BMMAF0'dan (_as400_cfi_bugun) kontrol ediyor,
    // yani geri cekilmek hicbir kaniti gizlemez.
    try {
        kilitCoz();
        geriCekil(6);
        var doneldu = anaMenuyeDon(4);   // TEK cagri: ikinci cagri bosuna PF3 gonderirdi
        log(doneldu ? "Kurtarma: ana menuye donuldu — sonraki satir devam edebilir"
                    : "UYARI: ana menuye donulemedi, sonraki satir de dusebilir");
    } catch (e) {
        log("UYARI: kurtarma sirasinda hata: " + e.message);
    }
    log("SONUC=IPTAL"); logDosya.Close(); WScript.Quit(2);
}

// Giris ekrani mi? (2-Variation + Causal cd gorunur)
function girisEkraniMi(e) {
    return e.indexOf("MOVIMENTI DI MAGAZZINO") >= 0 && e.indexOf("2-Variation") >= 0
        && e.indexOf("Causal cd") >= 0;
}
// Etiketli satirin numarasini bul (1-bazli); yoksa -1
function satirBul(dizi, etiket) {
    for (var i = 0; i < dizi.length; i++) if (dizi[i].indexOf(etiket) >= 0) return i + 1;
    return -1;
}

log(CAUSAL + " GIRIS: " + ARTICLE + " / " + ADET + " adet" + (CAUSAL === "COP" ? " (HURDA — counterpart BOS)" : "") + (DRYRUN ? " (DRYRUN)" : ""));

// ── G0: giris ekranina ulas ──
kilitCoz();
var e0 = ekranStr();
if (girisEkraniMi(e0)) {
    log("G0: zaten CFI giris ekraninda (2-Variation)");
} else {
    log("G0: navigasyon — ana menu > 07 > 01 > F1");
    geriCekil(6);
    if (!anaMenuyeDon(8)) iptal("Ana menuye ulasilamadi — B'yi elle ana menuye getirin");
    ps.SendKeys("07[enter]"); bekle();
    if (ekranStr() === e0) iptal("07 menusu acilmadi");
    ps.SendKeys("01[enter]"); bekle();
    if (ekranStr().indexOf("MOVIMENTI DI MAGAZZINO") < 0) iptal("MOVIMENTI DI MAGAZZINO gelmedi (07>01)");
    ps.SendKeys("[pf1]"); bekle();
    if (!girisEkraniMi(ekranStr())) iptal("2-Variation giris ekrani gelmedi (F1)");
    log("G0: 2-Variation giris ekrani acildi");
}
dump("giris-ekrani");

// ── G1: alan konumlari (runtime, ekrandan tureti) ──
var sat = ekran();
var whSat = satirBul(sat, "Warehouse cd");
var caSat = satirBul(sat, "Causal cd");
var arSat = satirBul(sat, "Article cd");
var qtSat = satirBul(sat, "Qty");
var cpSat = satirBul(sat, "Counterpart War.cd");
if (whSat < 0 || caSat < 0 || arSat < 0 || qtSat < 0 || cpSat < 0)
    iptal("Alan etiketleri bulunamadi (wh=" + whSat + " ca=" + caSat + " ar=" + arSat + " qt=" + qtSat + " cp=" + cpSat + ")");
// Deger sutunu: Warehouse satirindaki '01D' konumundan; yoksa varsayilan 25.
var kol = sat[whSat - 1].indexOf("01D") + 1;
if (kol <= 0) kol = 25;
log("G1: alan haritasi — wh@" + whSat + " ca@" + caSat + " ar@" + arSat + " qty@" + qtSat + " cp@" + cpSat + " kolon=" + kol);

// ── Satirdaki GIRIS alanlari (soldan saga) — autECLFieldList'ten OKUNUR ──
// NEDEN (2026-08-17): "Article cd" ekranda TEK alan DEGIL. 16 karakterlik ana alan,
// arada bosluk, sonra ek alan. 17 karakterlik kod (10.300.2757A-10-S) tek parca
// yazilinca 17. karakter alan sinirini asiyor; ustelik yanki kontrolu kol'dan
// itibaren 17 karakter okudugu icin araya dusen BOSLUGU kodun parcasi saniyor ve
// "Article cd yazilamadi" ile iptal ediyordu. Alan uzunlugunu VARSAYMA — ekrandan oku.
// NOT: bu PCOMM surumunde ps.GetFieldLength YOK; autECLFieldList kullaniliyor.
// JScript'te koleksiyona fl(i) ile erisilir (fl.Item(i) desteklenmiyor).
function satirGirisAlanlari(satir) {
    var a = [];
    try {
        var fl = ps.autECLFieldList;
        fl.Refresh();
        for (var i = 1; i <= fl.Count; i++) {
            var f = fl(i);
            if (Number(f.StartRow) === Number(satir) && !f.Protected && Number(f.Length) > 0)
                a.push({ kol: Number(f.StartCol), uz: Number(f.Length) });
        }
    } catch (e) {
        log("UYARI: alan listesi okunamadi (" + e.message + ") — tek alan varsayilacak");
        return [];
    }
    a.sort(function (x, y) { return x.kol - y.kol; });
    var s2 = [];
    for (var j = 0; j < a.length; j++) if (a[j].kol >= kol) s2.push(a[j]);
    return s2;
}

// Article kodunu alanlara BOLEREK yaz + her alani AYRI dogrula.
function articleYaz(satir, kod) {
    var al = satirGirisAlanlari(satir);
    // Alan listesi okunamadiysa eski davranis (tek alan) — kisa kodlar zaten boyle calisiyordu
    if (!al.length) { alanYaz(satir, kod, "Article cd"); return; }
    if (kod.length <= al[0].uz) { alanYaz(satir, kod, "Article cd"); return; }
    if (al.length < 2)
        iptal("Article kodu " + kod.length + " karakter ama ekranda tek " + al[0].uz +
              " karakterlik alan var — kod bu ekrana sigmiyor");
    var p1 = kod.substring(0, al[0].uz);
    var p2 = kod.substring(al[0].uz);
    if (p2.length > al[1].uz)
        iptal("Article kodu alanlara sigmiyor (" + al[0].uz + "+" + al[1].uz +
              " < " + kod.length + "): '" + kod + "'");
    log("Article BOLUNDU (kod " + kod.length + " karakter): alan1@" + al[0].kol + "(" +
        al[0].uz + ")='" + p1 + "'  alan2@" + al[1].kol + "(" + al[1].uz + ")='" + p2 + "'");
    ps.SetCursorPos(satir, al[0].kol); WScript.Sleep(200);
    ps.SendKeys(p1); WScript.Sleep(250);
    ps.SetCursorPos(satir, al[1].kol); WScript.Sleep(200);
    ps.SendKeys(p2); WScript.Sleep(250);
    // Dogrulama: alanlari AYRI oku, birlestir (aradaki bosluk kodun parcasi DEGIL)
    var g1 = ps.GetText(satir, al[0].kol, al[0].uz).replace(/\s+$/, "");
    var g2 = ps.GetText(satir, al[1].kol, p2.length);
    if (g1.toUpperCase() !== p1.toUpperCase() || g2.toUpperCase() !== p2.toUpperCase())
        iptal("Article cd yazilamadi (yanki: alan1='" + g1 + "' alan2='" + g2 + "')");
    log("Article yankisi dogrulandi: '" + g1 + "' + '" + g2 + "'");
}

if (DRYRUN) {
    // Article alan haritasini + planlanan bolmeyi RAPORLA (hicbir sey yazilmadan).
    // Boylece sahada uzun kodun nasil bolunecegi ONCE gorulur.
    var _al = satirGirisAlanlari(arSat);
    if (!_al.length) {
        log("DRYRUN: Article satirinda alan listesi okunamadi — tek alan varsayilir.");
    } else {
        var _tarif = [];
        for (var _i = 0; _i < _al.length && _i < 3; _i++)
            _tarif.push("alan" + (_i + 1) + "@" + _al[_i].kol + "(" + _al[_i].uz + ")");
        log("DRYRUN: Article giris alanlari -> " + _tarif.join("  "));
        if (ARTICLE.length <= _al[0].uz)
            log("DRYRUN: kod " + ARTICLE.length + " karakter, ilk alana SIGIYOR (bolme yok).");
        else if (_al.length >= 2)
            log("DRYRUN: kod " + ARTICLE.length + " karakter -> BOLUNECEK: '" +
                ARTICLE.substring(0, _al[0].uz) + "' + '" + ARTICLE.substring(_al[0].uz) + "'");
        else
            log("DRYRUN: DIKKAT — kod ilk alana sigmiyor ve ikinci alan YOK.");
    }
    log("DRYRUN: ekran + alan haritasi dogrulandi, HICBIR SEY YAZILMADI.");
    log("SONUC=DRYRUN-OK"); logDosya.Close(); WScript.Quit(0);
}

// ── G2: 01D / CFI / 01D (bos olanlari doldur — genelde onceki girdiden kalir) ──
function alanOku(satir, uz) { return ps.GetText(satir, kol, uz || 3).replace(/\s+/g, ""); }
function alanYaz(satir, deger, etiket) {
    ps.SetCursorPos(satir, kol); WScript.Sleep(200);
    ps.SendKeys(deger); WScript.Sleep(250);
    var geri = ps.GetText(satir, kol, deger.length);
    if (geri.toUpperCase() !== deger.toUpperCase())
        iptal(etiket + " yazilamadi (yanki: '" + geri + "')");
}
if (alanOku(whSat) !== "01D") alanYaz(whSat, "01D", "Warehouse cd");
if (alanOku(caSat) !== CAUSAL) alanYaz(caSat, CAUSAL, "Causal cd");
if (CAUSAL === "COP") {
    // HURDA: Counterpart War.cd BOS OLMALI (kullanici 2026-07-23). Onceki CFI
    // girisinden 01D kalmis olabilir → 3 bosluk yazarak temizle + dogrula.
    if (alanOku(cpSat) !== "") {
        ps.SetCursorPos(cpSat, kol); WScript.Sleep(200);
        ps.SendKeys("   "); WScript.Sleep(250);
        if (alanOku(cpSat) !== "") iptal("Counterpart War.cd temizlenemedi: '" + alanOku(cpSat) + "'");
        log("COP: Counterpart War.cd temizlendi (bos)");
    }
} else {
    if (alanOku(cpSat) !== "01D") alanYaz(cpSat, "01D", "Counterpart War.cd");
}

// GUVENLIK: Causal beklenenle birebir ayni olmali (yanlis hareket girilmesin)
if (alanOku(caSat) !== CAUSAL) iptal("Causal cd " + CAUSAL + " degil: '" + alanOku(caSat) + "'");

// ── G3: Article + Qty ──
// Article alani bos olmali (onceki girdiden kalan varsa Enter'lenmemis demektir → DUR)
var mevcutArticle = ps.GetText(arSat, kol, 20).replace(/\s+/g, "");
if (mevcutArticle.length) iptal("Article alani BOS DEGIL ('" + mevcutArticle + "') — B'de yarim giris olabilir, elle kontrol edin");
articleYaz(arSat, ARTICLE);   // uzun kodu alanlara boler (bkz. articleYaz)
ps.SetCursorPos(qtSat, kol); WScript.Sleep(200);
ps.SendKeys(ADET); WScript.Sleep(250);
ps.SendKeys("[fldext]"); WScript.Sleep(300);   // kucuk enter: saga yasla
// Yanki kontrolu: qty satirinda adet gorunmeli
if (ps.GetText(qtSat, 1, ps.NumCols).indexOf(ADET) < 0) iptal("Qty yankisi yok");
dump("giris-dolu");

// ── G4: Enter = kaydet + dogrula ──
ps.SendKeys("[enter]"); bekle(); WScript.Sleep(1200);
var basarili = false, sonE = "";
for (var d2 = 0; d2 < 10 && !basarili; d2++) {
    sonE = ekranStr();
    // KAYIT KILIDI (INVIO=Riprova / fase di aggiornamento) → Enter ile tekrar
    // (2026-07-23 dersi — magazzino dosyalari da kilitlenebilir)
    if (sonE.indexOf("INVIO=Riprova") >= 0 || sonE.indexOf("fase di aggiornamento") >= 0
        || sonE.indexOf("Risorsa richiesta") >= 0) {
        log("Kayit kilidi (Riprova) → Enter ile tekrar (deneme " + d2 + ")");
        kilitCoz(); ps.SendKeys("[enter]"); bekle(); WScript.Sleep(2000);
        continue;
    }
    var sats = ekran();
    // Basari: ust listede article gorunur VE Article giris alani BOSALDI
    var listede = false;
    for (var li = 2; li < 7 && li < sats.length; li++) {
        var sl = sats[li];
        // HARF DUYARSIZ (2026-07-30): operator kodu kucuk harfle yazabiliyor
        // (94.ltk.704); AS400 ekrani BUYUK gosterir. Harf duyarli arama, ERP'ye
        // BASARIYLA girilen hareketi 'gorunmedi' diye iptal ettiriyordu.
        if (sl.toUpperCase().indexOf(ARTICLE.toUpperCase()) >= 0) {
            // Ayni satirda adet (italyan bicimi: 53 → '53,000'; binlik nokta olabilir)
            var m = /([\d\.]+),000/.exec(sl);
            if (m && parseInt(m[1].replace(/\./g, ""), 10) === parseInt(ADET, 10)) { listede = true; break; }
        }
    }
    var alanBos = ps.GetText(arSat, kol, 20).replace(/\s+/g, "").length === 0;
    if (listede && alanBos) { basarili = true; break; }
    WScript.Sleep(900);
}
if (!basarili) {
    dump("enter-sonrasi-dogrulanamadi");
    iptal("Enter sonrasi dogrulama basarisiz — listede '" + ARTICLE + " " + ADET + "' gorunmedi (kod AS400'de tanimsiz olabilir)");
}
log("OK: ust listede " + ARTICLE + " / " + ADET + " goruldu, giris alanlari bosaldi");
dump("basarili");

// B'yi CFI ekraninda BIRAK (01D/CFI kalici) — sonraki cagri direkt kod+adet girer.
log("SONUC=OK");
logDosya.Close();

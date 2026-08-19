// ─────────────────────────────────────────────────────────────────────────────
// EOQ OKUYUCU — AS400 07/10/01 "SITUAZIONE ARTICOLO" (kullanıcı 2026-08-19)
//
// SALT OKUNUR: yalnız menü gezinme + ürün kodu yazma + Enter + PF3. Hiçbir veri
// DEĞİŞTİRMEZ, kaydetmez, onaylamaz.
//
// ⚠ 32-BİT cscript ŞART — düz "cscript" ile "Otomasyon sunucusu, nesne
//   oluşturamıyor" hatası verir (PCOMM COM 32-bit; 64-bit cscript göremez).
//   Ayrıca PCOMM oturumu AÇIK ve script SİZİN oturumunuzda koşmalı: Session 0'dan
//   (servis) başlatılan cscript PCOMM oturumlarını GÖREMEZ.
//
// KULLANIM
//   Tek kod (keşif/doğrulama — ekranı da döker):
//     C:\Windows\SysWOW64\cscript.exe //nologo as400\eoq_cek.js 10.300.3059W
//   Toplu (sunucudan kaynak referanslarını çeker, EOQ'ları geri yazar):
//     C:\Windows\SysWOW64\cscript.exe //nologo as400\eoq_cek.js --toplu http://192.168.20.210:5000
//   Toplu ama YAZMADAN (ne bulacağını gör):
//     ... --toplu http://192.168.20.210:5000 --deneme
//
// NOT: EOQ article master'da bir kolonsa bu betiğe HİÇ gerek yok —
// as400/eoq_kesif.py ODBC ile aynı işi saniyeler içinde yapar. Bu betik yalnız
// EOQ'nun SQL'den okunamadığı durum için yedektir.
//
// NOT (önemli): 07/10/01'den sonra ürün kodunun YAZILDIĞI ekranı görmedim —
// tek-kod modu bu yüzden ekranın TAMAMINI döker. İlk çalıştırmada çıktıyı
// paylaşın; gerekirse KOD_SATIR/KOD_KOLON ayarlanır. Akış tutmazsa script
// hiçbir şey bozmadan çıkar (temiz çıkış PF3 ile ana menüye döner).
// Kalıp: launch_ekran_oku.js (aynı bekleme/kilit-çözme/temiz-çıkış deseni).
// ─────────────────────────────────────────────────────────────────────────────

var OTURUM = "B";              // PCOMM bağlantı adı — diğer scriptlerle aynı
var MENU_YOLU = ["07", "10", "01"];
var EKRAN_BASLIK = "SITUAZIONE ARTICOLO";

var args = [];
for (var ai = 0; ai < WScript.Arguments.length; ai++) args.push("" + WScript.Arguments(ai));
var TOPLU = false, SUNUCU = "", DENEME = false, TEK_KOD = "";
for (var i = 0; i < args.length; i++) {
    if (args[i] === "--toplu") { TOPLU = true; SUNUCU = args[i + 1] || ""; i++; }
    else if (args[i] === "--deneme") { DENEME = true; }
    else if (!TEK_KOD) { TEK_KOD = args[i]; }
}
if (!TOPLU && !TEK_KOD) {
    WScript.Echo("Kullanim: eoq_cek.js <URUN_KODU>   |   eoq_cek.js --toplu <SUNUCU_URL> [--deneme]");
    WScript.Quit(1);
}

var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName(OTURUM);
var ps = s.autECLPS, oia = s.autECLOIA;

function estr() {
    var t = [];
    for (var r = 1; r <= ps.NumRows; r++) t.push(ps.GetText(r, 1, ps.NumCols));
    return t;
}
function ekran() { return estr().join("\n"); }
function bekle() { oia.WaitForInputReady(6000); WScript.Sleep(450); }
function kilitCoz() {
    ps.SendKeys("[reset]"); WScript.Sleep(250);
    ps.SendKeys("[fldext]"); WScript.Sleep(250);
    ps.SendKeys("[reset]"); WScript.Sleep(250);
}
function anaMenuyeDon() {
    for (var d = 0; d < 10; d++) {
        var e = ekran();
        if (e.indexOf("Menu S.I. CofleTk") >= 0) return true;
        if (e.indexOf("MenuIniziale") >= 0) ps.SendKeys("[pf16]");
        else {
            ps.SendKeys("[pf3]"); bekle();
            if (ekran() === e) { kilitCoz(); ps.SendKeys("[pf12]"); }
        }
        bekle();
    }
    return ekran().indexOf("Menu S.I. CofleTk") >= 0;
}
function dokum(baslik) {
    WScript.Echo("=== " + baslik + " ===");
    var e = estr();
    for (var i = 0; i < e.length; i++) {
        var sat = e[i].replace(/\s+$/, "");
        if (sat.replace(/\s+/g, "").length) WScript.Echo(((i + 1) < 10 ? " " : "") + (i + 1) + "| " + sat);
    }
}

// EOQ'yu ekran METNİNDEN çıkarır — kolon konumuna BAĞLI DEĞİL.
// Ekranda "EOQ . . . .   850" biçiminde duruyor (nokta dolgusu değişebilir).
function eoqOku(metin) {
    var m = metin.match(/EOQ[\s.:]*([0-9][0-9.,]*)/i);
    if (!m) return null;
    // 1.250 / 1,250 gibi binlik ayraçları at, ondalık yok (adet)
    return parseInt(m[1].replace(/[.,]/g, ""), 10);
}
// Ekrandaki ürün kodu — istediğimiz kodla aynı mı? (bayat ekran koruması)
function kodEslesti(metin, kod) {
    var a = metin.replace(/\s+/g, "").toUpperCase();
    return a.indexOf(kod.replace(/\s+/g, "").toUpperCase()) >= 0;
}

// 07/10/01 ekranına gider (ana menüden). Başarılıysa true.
function ekranaGit() {
    if (!anaMenuyeDon()) return false;
    for (var i = 0; i < MENU_YOLU.length; i++) { ps.SendKeys(MENU_YOLU[i] + "[enter]"); bekle(); }
    return true;
}

// Ekranda ürün kodunu sorgular; EOQ (sayı) ya da null döner.
function eoqSorgula(kod, dokumYap) {
    kilitCoz();
    ps.SendKeys("[home]"); WScript.Sleep(200);
    ps.SendKeys(kod); WScript.Sleep(200);
    ps.SendKeys("[enter]");
    for (var p = 0; p < 10; p++) {
        WScript.Sleep(500);
        if (ekran().indexOf(EKRAN_BASLIK) >= 0) break;
    }
    bekle();
    var e = ekran();
    if (dokumYap) dokum("EKRAN (" + kod + ")");
    if (e.indexOf(EKRAN_BASLIK) < 0) return { eoq: null, hata: "ekran acilmadi" };
    if (!kodEslesti(e, kod)) return { eoq: null, hata: "ekrandaki kod farkli" };
    var v = eoqOku(e);
    return { eoq: v, hata: (v === null ? "EOQ alani bulunamadi" : "") };
}

// Sonraki koda geçmek için sorgu alanına dön (PF3 = bir üst ekran)
function sorguyaDon() {
    kilitCoz();
    for (var d = 0; d < 4; d++) {
        if (ekran().indexOf(EKRAN_BASLIK) < 0) return true;
        ps.SendKeys("[pf3]"); bekle();
    }
    return ekran().indexOf(EKRAN_BASLIK) < 0;
}

// ── HTTP (sunucuyla konuşma) ────────────────────────────────────────────────
function httpGet(url) {
    var h = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
    h.setTimeouts(5000, 5000, 15000, 30000);
    h.open("GET", url, false); h.send();
    return h.status === 200 ? h.responseText : null;
}
function httpPost(url, govde) {
    var h = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
    h.setTimeouts(5000, 5000, 15000, 30000);
    h.open("POST", url, false);
    h.setRequestHeader("Content-Type", "application/json");
    h.send(govde);
    return h.status === 200 ? h.responseText : ("HTTP " + h.status);
}

// ── ÇALIŞTIR ────────────────────────────────────────────────────────────────
if (!ekranaGit()) {
    WScript.Echo("[HATA] 07/10/01 ekranina gidilemedi (ana menu bulunamadi).");
    dokum("SON EKRAN");
    WScript.Quit(2);
}

if (!TOPLU) {
    var tek = eoqSorgula(TEK_KOD, true);
    WScript.Echo("");
    WScript.Echo("KOD : " + TEK_KOD);
    WScript.Echo("EOQ : " + (tek.eoq === null ? "(okunamadi) " + tek.hata : tek.eoq));
    sorguyaDon(); anaMenuyeDon();
    WScript.Quit(0);
}

// Toplu: sunucudan kaynak referanslarini al
var ham = httpGet(SUNUCU + "/api/kaynak_eoq?bolum=kaynak&lokasyon=TK2");
if (!ham) { WScript.Echo("[HATA] Sunucudan referans listesi alinamadi: " + SUNUCU); WScript.Quit(3); }
var kodlar = [];
var re = /"referans_kodu"\s*:\s*"([^"]+)"/g, mm;
while ((mm = re.exec(ham)) !== null) kodlar.push(mm[1]);
WScript.Echo("[bilgi] " + kodlar.length + " kaynak referansi okunacak. Basla: " + new Date());

var sonuclar = [], basarisiz = 0;
for (var k = 0; k < kodlar.length; k++) {
    var r = eoqSorgula(kodlar[k], false);
    if (r.eoq === null) {
        basarisiz++;
        WScript.Echo("  [atlandi] " + kodlar[k] + " — " + r.hata);
    } else {
        sonuclar.push({ kod: kodlar[k], eoq: r.eoq });
        WScript.Echo("  " + kodlar[k] + " -> " + r.eoq);
    }
    if (!sorguyaDon()) { ekranaGit(); }   // akis bozulduysa bastan kur
}

WScript.Echo("[bilgi] okunan=" + sonuclar.length + " atlanan=" + basarisiz + " bitis: " + new Date());
anaMenuyeDon();

if (DENEME) { WScript.Echo("[deneme] Sunucuya YAZILMADI."); WScript.Quit(0); }
if (!sonuclar.length) { WScript.Echo("[bilgi] Yazilacak kayit yok."); WScript.Quit(0); }

var parca = [];
for (var j = 0; j < sonuclar.length; j++)
    parca.push('{"referans_kodu":"' + sonuclar[j].kod.replace(/"/g, '\\"') + '","eoq":' + sonuclar[j].eoq + '}');
var govde = '{"kaynak":"as400","lokasyon":"TK2","kayitlar":[' + parca.join(",") + ']}';
WScript.Echo("[sunucu] " + httpPost(SUNUCU + "/api/kaynak_eoq", govde));

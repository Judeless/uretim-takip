// ═══════════════════════════════════════════════════════════════════
// AS400 TEYIT PILOT — Session B, launch 26-140672 (10.300.1866-10W), 1 adet, bayrak A
// 32-bit cscript ile calisir:  C:\Windows\SysWOW64\cscript.exe //nologo teyit_pilot.js
// Guvenlik: her adimda ekran dogrulanir; beklenmedik ekranda F3 ile geri cekilip DURUR.
// S bayragina ASLA dokunulmaz (varsayilan A kalir).
// ═══════════════════════════════════════════════════════════════════
var OTURUM  = "B";
var YIL     = "26";
var NO      = "140672";
var ARTICLE = "10.300.1866-10W";   // ekranda birebir dogrulanacak
var ADET    = "1";

var fso = new ActiveXObject("Scripting.FileSystemObject");
// Log, scriptin kendi klasorune yazilir (Turkce karakterli yolu sistemden al —
// kaynak dosya kodlamasindan bagimsiz calissin)
var scriptKlasor = fso.GetParentFolderName(WScript.ScriptFullName);
var logDosya = fso.OpenTextFile(scriptKlasor + "\\pilot_log.txt", 2, true);

var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName(OTURUM);
var ps = s.autECLPS, oia = s.autECLOIA;

function log(m) { WScript.Echo(m); logDosya.WriteLine(m); }
function ekran() {
    var t = [];
    for (var r = 1; r <= ps.NumRows; r++) t.push(ps.GetText(r, 1, ps.NumCols));
    return t;
}
function ekranStr() { return ekran().join("\n"); }
function dump(etiket) {
    log("---------- EKRAN [" + etiket + "] ----------");
    var e = ekran();
    for (var i = 0; i < e.length; i++)
        if (e[i].replace(/\s+/g, "").length) log((i + 1 < 10 ? " " : "") + (i + 1) + "| " + e[i].replace(/\s+$/, ""));
    log("--------------------------------------");
}
function bekle() {
    oia.WaitForInputReady(8000);
    WScript.Sleep(450);
}
function iptal(neden) {
    log("!! IPTAL: " + neden);
    dump("iptal-ani");
    // guvenli geri cekilme: kilit coz (reset+fieldexit+reset), sonra 2x F3
    try {
        ps.SendKeys("[reset]"); WScript.Sleep(350);
        ps.SendKeys("[fldext]"); WScript.Sleep(350);
        ps.SendKeys("[reset]"); WScript.Sleep(350);
        ps.SendKeys("[pf3]"); bekle(); ps.SendKeys("[pf3]"); bekle();
    } catch (e) {}
    log("SONUC=IPTAL");
    logDosya.Close();
    WScript.Quit(2);
}
function bosluksuz(t) { return t.replace(/\s+/g, ""); }

log("PILOT BASLIYOR: launch " + YIL + "-" + NO + " / " + ARTICLE + " / " + ADET + " adet (A)");

// ── G0: Olasi klavye kilidini coz. 5250 kurali: sayisal alan tam dolunca
// "Field Exit bekliyor" durumuna girer; o haldeyken her AID tusu (Enter/F3)
// operator-error kilidi (OIA=5) tetikler. Recete: Reset → FieldExit → Reset. ──
ps.SendKeys("[reset]"); WScript.Sleep(350);
ps.SendKeys("[fldext]"); WScript.Sleep(350);
ps.SendKeys("[reset]"); WScript.Sleep(350);
log("G0: kilit cozme (reset+fieldexit+reset), OIA inhibit=" + oia.InputInhibited);

// ── G1: Ana menuye don (mevcut ekran neyse F3 ile geri cekil, menudeyse F16) ──
var tamam = false;
for (var d = 0; d < 7; d++) {
    var e = ekranStr();
    if (e.indexOf("Menu S.I. CofleTk") >= 0) { tamam = true; break; }
    if (e.indexOf("MenuIniziale") >= 0 && e.indexOf("F16") >= 0) ps.SendKeys("[pf16]");
    else ps.SendKeys("[pf3]");
    bekle();
}
if (!tamam) iptal("Ana menuye ulasilamadi (7 deneme)");
log("OK: Ana menu (TME1)");

// ── G2: 10 > 05 > 12 > 02 navigasyonu (her adimda ekran kodu dogrulanir) ──
var adimlar = [["10", "TP00"], ["05", "TP40"], ["12", "TP41"]];
for (var a = 0; a < adimlar.length; a++) {
    ps.SendKeys(adimlar[a][0] + "[enter]"); bekle();
    if (ekranStr().indexOf(adimlar[a][1]) < 0) iptal("Menu " + adimlar[a][0] + " sonrasi " + adimlar[a][1] + " gorunmedi");
    log("OK: menu " + adimlar[a][0] + " -> " + adimlar[a][1]);
}
ps.SendKeys("02[enter]"); bekle();
var e2 = ekranStr();
if (e2.indexOf("RIENTRO ORDINI PRODUZIONE") < 0 || e2.indexOf("Limits request") < 0)
    iptal("Rientro Limits ekrani gelmedi");
log("OK: RIENTRO / Limits request");

// ── G3: Order No doldur (yil + no; autoskip'e uyarlanir) + yankiyi dogrula ──
ps.SendKeys("[home]"); WScript.Sleep(250);
var r0 = ps.CursorPosRow, c0 = ps.CursorPosCol;
ps.SendKeys(YIL); WScript.Sleep(250);
var c1 = ps.CursorPosCol, r1 = ps.CursorPosRow;
// Sayisal (signed-numeric) alanlar autoskip YAPMAZ — alandan Field Exit ile cikilir.
// Cursor hala yil alanindaysa (≤2 kolon ilerlemis) Field Exit gonder.
if (r1 === r0 && (c1 - c0) <= 2) { ps.SendKeys("[fldext]"); WScript.Sleep(250); }
ps.SendKeys(NO); WScript.Sleep(250);
var satirTxt = ps.GetText(r0, 1, ps.NumCols);
if (satirTxt.indexOf(YIL) < 0 || satirTxt.indexOf(NO) < 0)
    iptal("Order No yankisi hatali: '" + satirTxt.replace(/\s+$/, "") + "'");
log("OK: Order No girildi: " + satirTxt.replace(/\s+$/, ""));
// Sayi alanindan Field Exit ile cik (numerik alan sagda yaslansin), sonra Enter.
ps.SendKeys("[fldext]"); WScript.Sleep(300);
ps.SendKeys("[enter]");
log("Enter sonrasi OIA inhibit: " + oia.InputInhibited);
// Ozet ekranini kademeli bekle (8 x 600ms) — ara durumlari logla
var e4 = "";
for (var p = 0; p < 8; p++) {
    WScript.Sleep(600);
    e4 = ekranStr();
    if (e4.indexOf("Re-entry summary") >= 0) break;
    if (p === 3) { log("... hala bekleniyor, OIA=" + oia.InputInhibited); dump("bekleme-4"); }
}

// ── G4: Re-entry summary dogrula (launch + ARTICLE birebir) ──
if (e4.indexOf("Re-entry summary") < 0) iptal("Re-entry summary gelmedi");
if (e4.indexOf(NO) < 0) iptal("Launch no ekranda yok");
if (bosluksuz(e4).indexOf(bosluksuz(ARTICLE)) < 0) iptal("ARTICLE eslesmedi! Beklenen: " + ARTICLE);
log("OK: Re-entry summary — launch " + NO + " / " + ARTICLE + " dogrulandi");
dump("re-entry-summary-ilk");

// ── G5: F10 = yeni satir, adet + kucuk Enter (FieldExit) ──
ps.SendKeys("[pf10]"); bekle();
if (ekranStr().indexOf("Re-entry summary") < 0) iptal("F10 sonrasi ekran degisti");
var qr = ps.CursorPosRow;
ps.SendKeys(ADET); WScript.Sleep(250);
ps.SendKeys("[fldext]"); WScript.Sleep(400);
var satirQ = ps.GetText(qr, 1, ps.NumCols);
if (satirQ.indexOf(ADET) < 0) iptal("Adet yankisi yok. Satir: '" + satirQ.replace(/\s+$/, "") + "'");
if (satirQ.replace(/\s+$/, "").charAt(satirQ.replace(/\s+$/, "").length - 1) !== "A")
    iptal("Satir sonunda A bayragi gorunmuyor: '" + satirQ.replace(/\s+$/, "") + "'");
log("OK: Adet " + ADET + " girildi, bayrak A: '" + satirQ.replace(/\s+$/, "") + "'");
dump("adet-girildi");

// ── G6: buyuk Enter x2 (her birinde hata kontrolu) ──
for (var n = 1; n <= 2; n++) {
    ps.SendKeys("[enter]"); bekle();
    var en = ekran();
    var alt = (en[21] + " " + en[22] + " " + en[23]).toLowerCase();
    if (alt.indexOf("error") >= 0 || alt.indexOf("errore") >= 0 || alt.indexOf("non valid") >= 0)
        iptal("Enter#" + n + " sonrasi hata mesaji");
    log("OK: Enter #" + n);
}
dump("enterlar-sonrasi");

// ── G7: F6 = kaydet ──
ps.SendKeys("[pf6]"); bekle();
dump("F6-sonrasi");
log("OK: F6 gonderildi");

// ── G8: temiz cikis — ana menuye don ──
for (var d2 = 0; d2 < 6; d2++) {
    var ec = ekranStr();
    if (ec.indexOf("Menu S.I. CofleTk") >= 0) break;
    if (ec.indexOf("MenuIniziale") >= 0 && ec.indexOf("F16") >= 0) ps.SendKeys("[pf16]");
    else ps.SendKeys("[pf3]");
    bekle();
}
log("OK: Session B ana menuye birakildi");
log("SONUC=OK");
logDosya.Close();

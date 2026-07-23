// Session B kurtarma: kilit coz + F3/F12 donusumlu geri cekilme → ana menu.
// Bekleyen (F6'lanmamis) satirlar kaydedilmeden terk edilir — guvenli.
var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName("B");
var ps = s.autECLPS, oia = s.autECLOIA;
function ekranStr() { var t = []; for (var r = 1; r <= ps.NumRows; r++) t.push(ps.GetText(r, 1, ps.NumCols)); return t.join("\n"); }
function bekle() { oia.WaitForInputReady(6000); WScript.Sleep(400); }
// SADECE reset — fldext KULLANMA (giris alanindayken alan icerigini siler).
function kilitCoz() { ps.SendKeys("[reset]"); WScript.Sleep(350); }
kilitCoz();
for (var d = 0; d < 10; d++) {
    var e = ekranStr();
    if (e.indexOf("Menu S.I. CofleTk") >= 0) { WScript.Echo("ANA MENUDE (adim " + d + ")"); WScript.Quit(0); }
    if (e.indexOf("MenuIniziale") >= 0 && e.indexOf("F16") >= 0) { ps.SendKeys("[pf16]"); bekle(); continue; }
    var once = e;
    ps.SendKeys("[pf3]"); bekle();
    if (ekranStr() === once) {           // F3 ise yaramadi → kilit coz + F12 dene
        kilitCoz();
        ps.SendKeys("[pf12]"); bekle();
        if (ekranStr() === once) { kilitCoz(); ps.SendKeys("[pf3]"); bekle(); }
    }
}
WScript.Echo(ekranStr().indexOf("Menu S.I. CofleTk") >= 0 ? "ANA MENUDE" : "ULASILAMADI — ekran basi: " + ps.GetText(1, 1, 80) + " / " + ps.GetText(2, 1, 80));

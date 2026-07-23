// Launch ozet ekranini OKU (salt-okuma). Kullanim: ... launch_ekran_oku.js 26 140672
var YIL = "" + WScript.Arguments(0), NO = "" + WScript.Arguments(1);
var s = new ActiveXObject("PCOMM.autECLSession"); s.SetConnectionByName("B");
var ps = s.autECLPS, oia = s.autECLOIA;
function estr() { var t = []; for (var r = 1; r <= ps.NumRows; r++) t.push(ps.GetText(r, 1, ps.NumCols)); return t; }
function bekle() { oia.WaitForInputReady(6000); WScript.Sleep(450); }
function kilitCoz() { ps.SendKeys("[reset]"); WScript.Sleep(250); ps.SendKeys("[fldext]"); WScript.Sleep(250); ps.SendKeys("[reset]"); WScript.Sleep(250); }
kilitCoz();
// ana menu
for (var d = 0; d < 8; d++) { var e = estr().join("\n"); if (e.indexOf("Menu S.I. CofleTk") >= 0) break; if (e.indexOf("MenuIniziale") >= 0) ps.SendKeys("[pf16]"); else ps.SendKeys("[pf3]"); bekle(); }
ps.SendKeys("10[enter]"); bekle(); ps.SendKeys("05[enter]"); bekle(); ps.SendKeys("12[enter]"); bekle(); ps.SendKeys("02[enter]"); bekle();
ps.SendKeys("[home]"); WScript.Sleep(200);
ps.SendKeys(YIL); WScript.Sleep(200); ps.SendKeys("[fldext]"); WScript.Sleep(200); ps.SendKeys(NO); WScript.Sleep(200); ps.SendKeys("[fldext]"); WScript.Sleep(200); ps.SendKeys("[enter]");
for (var p = 0; p < 8; p++) { WScript.Sleep(600); if (estr().join("\n").indexOf("Re-entry summary") >= 0) break; }
var e2 = estr();
WScript.Echo("=== OZET EKRANI (launch " + YIL + "-" + NO + ") ===");
for (var i = 0; i < e2.length; i++) if (e2[i].replace(/\s+/g, "").length) WScript.Echo(((i + 1) < 10 ? " " : "") + (i + 1) + "| " + e2[i].replace(/\s+$/, ""));
// temiz cikis (hicbir sey degistirmeden)
kilitCoz();
for (var d2 = 0; d2 < 8; d2++) { var e3 = estr().join("\n"); if (e3.indexOf("Menu S.I. CofleTk") >= 0) break; if (e3.indexOf("MenuIniziale") >= 0) ps.SendKeys("[pf16]"); else { ps.SendKeys("[pf3]"); bekle(); if (estr().join("\n") === e3) { kilitCoz(); ps.SendKeys("[pf12]"); bekle(); } } }
WScript.Echo("(cikildi)");

// Session B tanisi — SALT OKUMA + zararsiz imlec testi ([tab])
var cl = new ActiveXObject("PCOMM.autECLConnList");
cl.Refresh();
for (var i = 1; i <= cl.Count; i++) {
    var c = cl(i);
    WScript.Echo("oturum " + c.Name + " | CommStarted=" + c.CommStarted +
                 " | APIEnabled=" + c.APIEnabled + " | Ready=" + c.Ready);
}
var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName("B");
var ps = s.autECLPS, oia = s.autECLOIA;
WScript.Echo("OIA InputInhibited=" + oia.InputInhibited);
var r1 = ps.CursorPosRow, c1 = ps.CursorPosCol;
WScript.Echo("imlec once : " + r1 + "," + c1);
ps.SendKeys("[tab]");
WScript.Sleep(500);
var r2 = ps.CursorPosRow, c2 = ps.CursorPosCol;
WScript.Echo("imlec sonra: " + r2 + "," + c2 + "  → " + ((r1 !== r2 || c1 !== c2) ? "TUSLAR YEREL OLARAK ISLIYOR" : "IMLEC KIMILDAMADI (tuslar PS'e ulasmiyor)"));
WScript.Echo("ekran satir 1: " + ps.GetText(1, 1, 80));

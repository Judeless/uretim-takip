// Field Exit mnemonik adaylarini dene — kabul edileni bul.
// (Gecerli olan ayni zamanda B'deki bekleyen sayisal alani kurtarir.)
var s = new ActiveXObject("PCOMM.autECLSession");
s.SetConnectionByName("B");
var ps = s.autECLPS, oia = s.autECLOIA;
ps.SendKeys("[reset]"); WScript.Sleep(300);
var adaylar = ["[fieldexit]", "[field exit]", "[fldext]", "[field+]", "[fieldplus]"];
for (var i = 0; i < adaylar.length; i++) {
    try {
        ps.SendKeys(adaylar[i]);
        WScript.Sleep(350);
        WScript.Echo("KABUL: " + adaylar[i] + "  (OIA=" + oia.InputInhibited + ")");
        WScript.Quit(0);
    } catch (e) {
        WScript.Echo("red : " + adaylar[i] + " -> " + e.message.substring(0, 60));
    }
}
WScript.Echo("HICBIRI KABUL EDILMEDI");
WScript.Quit(1);

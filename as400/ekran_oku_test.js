// SALT-OKUMA testi — 32-bit cscript ile PCOMM autECL koprusu.
// Kullanim: C:\Windows\SysWOW64\cscript.exe //nologo ekran_oku_test.js
// HICBIR TUSA BASILMAZ; oturumlari listeler + ekranin ust satirlarini okur.
try {
    var cl = new ActiveXObject("PCOMM.autECLConnList");
    cl.Refresh();
    WScript.Echo("OTURUMLAR=" + cl.Count);
    var hedef = "";
    for (var i = 1; i <= cl.Count; i++) {
        var c = cl(i);
        WScript.Echo("  oturum " + c.Name + " baglanti=" + (c.CommStarted ? "ACIK" : "kapali"));
        if (c.Name == "B") hedef = "B";
        if (!hedef && c.Name == "A") hedef = "A";
    }
    if (!hedef) { WScript.Echo("HATA: A/B oturumu yok"); WScript.Quit(1); }

    var s = new ActiveXObject("PCOMM.autECLSession");
    s.SetConnectionByName(hedef);
    var ps = s.autECLPS;
    WScript.Echo("HEDEF=" + hedef + " EKRAN=" + ps.NumRows + "x" + ps.NumCols);
    for (var r = 1; r <= 6; r++) {
        var t = ps.GetText(r, 1, ps.NumCols);
        if (t.replace(/\s+/g, "").length > 0)
            WScript.Echo("SATIR" + r + ": " + t.replace(/\s+$/, ""));
    }
    WScript.Echo("OIA_InputInhibited=" + s.autECLOIA.InputInhibited);
    WScript.Echo("SONUC=OK");
} catch (e) {
    WScript.Echo("HATA: " + e.message);
    WScript.Quit(1);
}

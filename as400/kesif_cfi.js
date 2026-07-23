// kesif_cfi.js — CFI (2-Variation) ekrani ALAN KESFI. SALT-OKUNUR: hicbir tus basilmaz.
// Ekrani + giris alanlarinin (unprotected fields) satir/sutun/uzunluklarini doker.
var sess = new ActiveXObject("PCOMM.autECLSession");
sess.SetConnectionByName("B");
var ps = sess.autECLPS;
var rows = ps.NumRows, cols = ps.NumCols;
WScript.Echo("EKRAN=" + rows + "x" + cols);
for (var r = 1; r <= rows; r++) {
    var t = ps.GetText(r, 1, cols);
    if (t.replace(/\s+/g, "").length)
        WScript.Echo((r < 10 ? " " : "") + r + "| " + t.replace(/\s+$/, ""));
}
WScript.Echo("---- ALANLAR (unprotected giris alanlari) ----");
try {
    var fl = ps.autECLFieldList;
    fl.Refresh(1);   // 1 = tum alanlar
    WScript.Echo("AlanSayisi=" + fl.Count);
    for (var i = 1; i <= fl.Count; i++) {
        var f = fl(i);
        var korumali;
        try { korumali = f.Protected; } catch (e) { korumali = "?"; }
        if (korumali === false || korumali === 0) {
            var icerik = "";
            try { icerik = f.GetText(); } catch (e2) { icerik = "(okunamadi)"; }
            WScript.Echo("GIRIS: satir=" + f.StartRow + " sutun=" + f.StartCol +
                         " uzunluk=" + f.Length + " icerik='" + icerik.replace(/\s+$/, "") + "'");
        }
    }
} catch (e) {
    WScript.Echo("FieldList HATA: " + e.message + " — yalniz ekran dokumu kullanilacak");
}
WScript.Echo("SONUC=OK");

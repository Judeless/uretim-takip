var s=new ActiveXObject("PCOMM.autECLSession");s.SetConnectionByName("B");
var ps=s.autECLPS,oia=s.autECLOIA;
function estr(){var t=[];for(var r=1;r<=ps.NumRows;r++){var x=ps.GetText(r,1,ps.NumCols);if(x.replace(/\s+/g,"").length)t.push(x.replace(/\s+$/,""))}return t.join(" || ")}
function bekle(){oia.WaitForInputReady(6000);WScript.Sleep(600);}
// F03=Uscita ile kilit ekranindan cik (Riprova YAPMA)
ps.SendKeys("[reset]");WScript.Sleep(500);
ps.SendKeys("[pf3]");bekle();
WScript.Echo("F03 sonrasi s1-2: "+ps.GetText(1,1,55).replace(/\s+$/,"")+" | "+ps.GetText(2,1,40).replace(/\s+$/,"")+" | OIA="+oia.InputInhibited);
// Ana menuye kadar F3/F12 (fldext YOK)
for(var d=0;d<8;d++){var e=estr();if(e.indexOf("Menu S.I. CofleTk")>=0)break;
  ps.SendKeys("[reset]");WScript.Sleep(300);
  var o=e;ps.SendKeys("[pf3]");bekle();
  if(estr()===o){ps.SendKeys("[reset]");WScript.Sleep(300);ps.SendKeys("[pf12]");bekle();}}
WScript.Echo(estr().indexOf("Menu S.I. CofleTk")>=0?"ANA MENUDE ✓":"HALA: "+ps.GetText(1,1,55).replace(/\s+$/,"")+" | "+ps.GetText(2,1,40).replace(/\s+$/,""));

# ============================================================
#  ProManage + Oracle'i KAPAT / GERI AL (kullanici 2026-11-24)
#  Sunucu 15.10 ve 20.11'de kilitlendi. Teshis: ProManage'in SrvcOlay sureci
#  4 gunde 77.652 handle, NetPC 706 kapanmamis baglanti + 3,7 GB -> ~5 haftada
#  Windows yeni surec / oturum acamaz hale geliyor. Kullanici: "ProManage'i
#  kullanmiyoruz, server bizim". Bizim sistem (cofle-*, aspanda-*, Cloudflared,
#  PCOMM, teyit-agent) Oracle'i da ProManage'i da KULLANMAZ.
#
#  -Kapat  : listeyi gosterir, 'KAPAT' yazilinca: servisler durdurulur + Devre disi,
#            zamanlanmis gorevler kapatilir, Baslangic kisayollari yedek klasore
#            tasinir, HKLM Run degerleri kaldirilir, kalan surecler kapatilir.
#            Hepsi C:\cofle\logs\promanage_yedek.json'a yazilir.
#  -GeriAl : yedekteki her seyi eski haline getirir.
#  Hicbir program dosyasi / veritabani SILINMEZ.
#  'promanage' WINDOWS HESABINA DOKUNULMAZ: PCOMM, teyit-agent ve RDP o hesapla
#  calisir; 'C:\Users\promanage\' yolu ProManage YAZILIMI sayilmaz.
# ============================================================
param([switch]$Kapat, [switch]$GeriAl)
$ErrorActionPreference = 'Continue'
$logdir = 'C:\cofle\logs'
New-Item -ItemType Directory -Force -Path $logdir | Out-Null
$yedek = Join-Path $logdir 'promanage_yedek.json'
$lnk_yedek = 'C:\cofle\yedek\promanage_baslangic'
$kayit = Join-Path $logdir ("promanage_{0:yyyyMMdd_HHmm}.log" -f (Get-Date))
function Yaz([string]$s, [string]$renk = 'Gray') { Write-Host $s -ForegroundColor $renk; $s | Out-File $kayit -Append -Encoding utf8 }

# Bizim servisler - ASLA dokunulmaz
$KORUNAN = 'cofle-app', 'cofle-pilot', 'Cloudflared', 'aspanda-app', 'aspanda-tunel', 'TermService'
$desen = 'promanage|pro-manage|prmng|netpc|srvcolay|proview|oracle|tnslsnr'
# Bizim acilis zincirimiz (PCOMM, teyit-agent, NSSM, tunel) - kisayol / gorev / Run ASLA dokunulmaz
$KORUNAN_HEDEF = '(?i)pcsws|\\cofle\\|uretim_takip|cloudflared|nssm|teyit'
function Eslesir([string]$metin) {
    # 'C:\Users\promanage\...' kullanici klasorudur, ProManage yazilimi degil
    ($metin -replace '(?i)\\users\\promanage\\', '\users\_\') -match $desen
}
$RUN = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'

# ------------------------- GERI AL -------------------------
if ($GeriAl) {
    if (-not (Test-Path $yedek)) { Yaz "[HATA] Yedek yok: $yedek" Red; exit 1 }
    $y = Get-Content $yedek -Raw | ConvertFrom-Json
    Yaz "=== GERI ALINIYOR ($($y.zaman) yedegi) ===" Cyan
    foreach ($s in @($y.servisler)) {
        $tur = switch ($s.StartMode) { 'Auto' { if ($s.Gecikmeli) { 'delayed-auto' } else { 'auto' } } 'Manual' { 'demand' } default { 'disabled' } }
        sc.exe config "$($s.Name)" start= $tur | Out-Null
        Yaz ("  {0,-40} baslangic -> {1}" -f $s.Name, $tur)
    }
    # Oracle once (ProManage ona baglanir), sonra digerleri
    foreach ($s in @($y.servisler | Where-Object { $_.State -eq 'Running' } | Sort-Object { if ($_.Name -like 'Oracle*') { 0 } else { 1 } })) {
        Start-Service -Name $s.Name -ErrorAction SilentlyContinue
        Yaz ("  {0,-40} {1}" -f $s.Name, (Get-Service $s.Name).Status)
    }
    foreach ($g in @($y.gorevler)) {
        Enable-ScheduledTask -TaskPath $g.TaskPath -TaskName $g.TaskName -ErrorAction SilentlyContinue | Out-Null
        Yaz ("  gorev acildi: {0}{1}" -f $g.TaskPath, $g.TaskName)
    }
    foreach ($k in @($y.kisayollar)) {
        if (Test-Path $k.Yedek) { Move-Item -LiteralPath $k.Yedek -Destination $k.Yol -Force; Yaz "  kisayol geri kondu: $($k.Yol)" }
    }
    foreach ($r in @($y.run)) {
        New-ItemProperty -Path $r.Anahtar -Name $r.Ad -Value $r.Deger -PropertyType String -Force | Out-Null
        Yaz "  Run geri kondu: $($r.Ad)"
    }
    Yaz "[OK] Geri alindi. Programlar acilista / oturum acilinca yeniden baslar. Kayit: $kayit" Green
    exit 0
}

# ------------------------- ENVANTER -------------------------
$servisler = @(Get-CimInstance Win32_Service | Where-Object {
    $KORUNAN -notcontains $_.Name -and (Eslesir ($_.Name + ' ' + $_.DisplayName + ' ' + $_.PathName)) })
$surecler = @(Get-Process | Where-Object { $_.ProcessName -match $desen })
$klasorler = @(@($servisler | ForEach-Object { if ($_.PathName -match '^"?([^"]+\\)[^\\"]+\.exe') { $Matches[1] } }) +
               @($surecler | Where-Object { $_.Path } | ForEach-Object { (Split-Path $_.Path -Parent) + '\' }) |
               # EN AZ IKI KLASOR DERIN olmali: 'C:\' ya da 'C:\Program Files\' gibi genis bir kok, oradaki
               # HER kisayolu / gorevi (PCOMM dahil) ProManage sanardi
               Where-Object { $_ -and $_ -match '^[A-Za-z]:\\[^\\]+\\[^\\]+\\' -and $_ -notmatch '(?i)\\(windows|cofle|users)\\' } |
               Sort-Object -Unique)
function Klasorde([string]$yol) { foreach ($k in $klasorler) { if ($yol -and $yol -like "$k*") { return $true } }; $false }

$gorevler = @(Get-ScheduledTask | Where-Object { $_.TaskPath -notlike '\Microsoft\*' -and $_.TaskPath -notlike '\Cofle\*' } | Where-Object {
    $islem = ($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join ' '
    $islem -notmatch $KORUNAN_HEDEF -and ((Eslesir $islem) -or (Klasorde (($_.Actions | Select-Object -First 1).Execute -replace '"', ''))) })

$sh = New-Object -ComObject WScript.Shell
$bas_klasorleri = @('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp') +
                  @(Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' })
$kisayollar = @(foreach ($bk in $bas_klasorleri) {
    Get-ChildItem -LiteralPath $bk -File -ErrorAction SilentlyContinue | ForEach-Object {
        $hedef = if ($_.Extension -eq '.lnk') { $l = $sh.CreateShortcut($_.FullName); "$($l.TargetPath) $($l.Arguments)" } else { $_.FullName }
        if ($hedef -notmatch $KORUNAN_HEDEF -and ((Eslesir ($_.Name + ' ' + $hedef)) -or (Klasorde ($hedef.Trim())))) { [pscustomobject]@{ Yol = $_.FullName; Hedef = $hedef } }
    } })
$runlar = @(foreach ($a in $RUN) {
    $o = Get-ItemProperty -Path $a -ErrorAction SilentlyContinue
    if ($o) { $o.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' -and [string]$_.Value -notmatch $KORUNAN_HEDEF -and ((Eslesir ($_.Name + ' ' + $_.Value)) -or (Klasorde (($_.Value -replace '"', '').Split(' ')[0]))) } |
              ForEach-Object { [pscustomobject]@{ Anahtar = $a; Ad = $_.Name; Deger = [string]$_.Value } } } })

Yaz "=== SERVISLER ===" Cyan
if (-not $servisler.Count) { Yaz "  (eslesen servis yok)" }
foreach ($s in $servisler | Sort-Object Name) {
    $pr = if ($s.ProcessId) { Get-Process -Id $s.ProcessId -ErrorAction SilentlyContinue }
    $ek = if ($pr) { "{0:N0} MB, {1} handle" -f ($pr.PrivateMemorySize64 / 1MB), $pr.HandleCount } else { '' }
    Yaz ("  {0,-40} {1,-8} {2,-8} {3}" -f $s.Name, $s.State, $s.StartMode, $ek)
    Yaz ("      {0} | {1}" -f $s.DisplayName, $s.PathName) DarkGray
}
Yaz ""
Yaz "=== SURECLER ===" Cyan
foreach ($p in $surecler) { Yaz ("  {0,-20} pid {1,-6} {2:N0} MB  handle {3}  {4}" -f $p.ProcessName, $p.Id, ($p.PrivateMemorySize64 / 1MB), $p.HandleCount, $p.Path) }
Yaz ""
Yaz "=== ACILISTA BASLATANLAR (zamanlanmis gorev / Baslangic kisayolu / Run) ===" Cyan
foreach ($g in $gorevler) { Yaz ("  gorev   : {0}{1}  [{2}]" -f $g.TaskPath, $g.TaskName, $g.State) }
foreach ($k in $kisayollar) { Yaz ("  kisayol : {0}  ->  {1}" -f $k.Yol, $k.Hedef) }
foreach ($r in $runlar) { Yaz ("  Run     : {0}  ->  {1}" -f $r.Ad, $r.Deger) }
if (-not ($gorevler.Count + $kisayollar.Count + $runlar.Count)) { Yaz "  (yok)" }

# Hala ProManage / Oracle'a BASKA BILGISAYARDAN baglanan var mi?
$benim = @(Get-NetIPAddress | ForEach-Object { $_.IPAddress }) + '127.0.0.1', '::1', '0.0.0.0', '::'
$pidler = @(@($servisler | ForEach-Object { $_.ProcessId }) + @($surecler | ForEach-Object { $_.Id }) | Where-Object { $_ } | Sort-Object -Unique)
$dis = @(Get-NetTCPConnection -State Established | Where-Object { $pidler -contains $_.OwningProcess -and $benim -notcontains $_.RemoteAddress })
Yaz ""
Yaz "=== BU PROGRAMLARA BASKA BILGISAYARDAN BAGLANANLAR ===" Cyan
if (-not $dis.Count) { Yaz "  (yok - yalniz sunucunun kendi icinde kullaniliyor)" Green }
else {
    $dis | Group-Object RemoteAddress | ForEach-Object {
        $isim = try { [System.Net.Dns]::GetHostEntry($_.Name).HostName } catch { '' }
        $port = ($_.Group | ForEach-Object { $_.LocalPort } | Sort-Object -Unique) -join ','
        Yaz ("  {0,-16} {1,-30} {2} baglanti, sunucu portu {3}" -f $_.Name, $isim, $_.Count, $port) Yellow
    }
    Yaz "  !! Bu bilgisayarlar hala ProManage / Oracle ile konusuyor; kapatinca veri gonderemezler." Yellow
}

if (-not $Kapat) { Yaz ""; Yaz "Yalniz listelendi. Kapatmak icin ProManage_Kapat.bat."; exit 0 }
if (-not ($servisler.Count + $surecler.Count + $gorevler.Count + $kisayollar.Count + $runlar.Count)) { Yaz "Kapatilacak bir sey yok."; exit 0 }

Yaz ""
$cevap = Read-Host "Yukaridakiler DURDURULUP DEVRE DISI birakilsin mi? Onay icin KAPAT yazin"
if ($cevap -ne 'KAPAT') { Yaz "Vazgecildi - hicbir sey degismedi." Yellow; exit 0 }

# ------------------------- YEDEK -------------------------
New-Item -ItemType Directory -Force -Path $lnk_yedek | Out-Null
$i = 0
$k_yedek = @($kisayollar | ForEach-Object { $i++; @{ Yol = $_.Yol; Yedek = (Join-Path $lnk_yedek ("{0:D2}_{1}" -f $i, (Split-Path $_.Yol -Leaf))) } })
@{ zaman = (Get-Date -Format 'yyyy-MM-dd HH:mm')
   servisler = @($servisler | ForEach-Object {
       $v = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\$($_.Name)" -Name DelayedAutostart -ErrorAction SilentlyContinue).DelayedAutostart
       @{ Name = $_.Name; StartMode = $_.StartMode; State = $_.State; Gecikmeli = ($v -eq 1) } })
   gorevler = @($gorevler | Where-Object { $_.State -ne 'Disabled' } | ForEach-Object { @{ TaskPath = $_.TaskPath; TaskName = $_.TaskName } })
   kisayollar = $k_yedek
   run = @($runlar | ForEach-Object { @{ Anahtar = $_.Anahtar; Ad = $_.Ad; Deger = $_.Deger } })
} | ConvertTo-Json -Depth 4 | Out-File $yedek -Encoding utf8
Yaz "[OK] Onceki durum yedeklendi: $yedek" Green

# ------------- KAPAT: once ProManage, en son Oracle (ProManage ona baglidir) -------------
foreach ($g in $gorevler) {
    Disable-ScheduledTask -TaskPath $g.TaskPath -TaskName $g.TaskName -ErrorAction SilentlyContinue | Out-Null
    Yaz ("  gorev kapatildi: {0}{1}" -f $g.TaskPath, $g.TaskName)
}
foreach ($k in $k_yedek) { Move-Item -LiteralPath $k.Yol -Destination $k.Yedek -Force; Yaz "  kisayol tasindi: $($k.Yol)" }
foreach ($r in $runlar) { Remove-ItemProperty -Path $r.Anahtar -Name $r.Ad -ErrorAction SilentlyContinue; Yaz "  Run kaldirildi: $($r.Ad)" }
foreach ($s in @($servisler | Sort-Object { if ($_.Name -like 'Oracle*') { 1 } else { 0 } })) {
    Set-Service -Name $s.Name -StartupType Disabled -ErrorAction SilentlyContinue
    if ($s.State -eq 'Running') { Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue }
    $son = Get-Service $s.Name
    Yaz ("  {0,-40} {1,-8} {2}" -f $s.Name, $son.Status, $son.StartType) $(if ($son.Status -eq 'Stopped') { 'Green' } else { 'Yellow' })
}
Start-Sleep -Seconds 3
foreach ($p in @(Get-Process | Where-Object { $_.ProcessName -match $desen })) {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Yaz ("  surec kapatildi: {0} (pid {1})" -f $p.ProcessName, $p.Id)
}
$os = Get-CimInstance Win32_OperatingSystem
Yaz ""
Yaz ("[OK] Kapatildi. Bos RAM simdi {0:N1} GB. Geri almak icin: ProManage_Geri_Al.bat" -f ($os.FreePhysicalMemory / 1MB)) Green
Yaz "     Kayit: $kayit"

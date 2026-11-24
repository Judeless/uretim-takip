# ============================================================
#  Cofle sunucu ARIZA TESHISI (kullanici 2026-11-20)
#  Sunucu 15.10 ve 20.11'de kilitlendi: portlar acik, uygulama LAN'da
#  cevap veriyor ama RDP yok + Cloudflare 502 (tunel uygulamaya yeni
#  baglanti acamiyor) -> fiziki reset. Kok neden olculmedi.
#  Bu betik reset SONRASI da kalan Windows olay gunluklerinden iki arizanin
#  izini + su anki kaynak kullanimini C:\cofle\logs\ariza_<tarih>\'e yazar.
#  OZET.txt'yi gonder: RAM mi (2004), TCP port tukenmesi mi (4227/4231),
#  cekirdek havuzu mu (2019/2020), servis cokmesi mi (7031/7034) gorunur.
#  Yalniz OKUR, hicbir ayari degistirmez.
# ============================================================
$ErrorActionPreference = 'SilentlyContinue'
$ts  = Get-Date -Format 'yyyyMMdd_HHmm'
$dir = "C:\cofle\logs\ariza_$ts"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$ozet = Join-Path $dir 'OZET.txt'
function Yaz([string]$s) { $s | Out-File -FilePath $ozet -Append -Encoding utf8; Write-Host $s }

$os = Get-CimInstance Win32_OperatingSystem
$m  = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
Yaz "=== COFLE SUNUCU TESHIS $(Get-Date -Format 'yyyy-MM-dd HH:mm') ==="
Yaz ("Makine: {0} | {1} | son acilis: {2}" -f $env:COMPUTERNAME, $os.Caption, $os.LastBootUpTime)
Yaz ("RAM: toplam {0:N1} GB, bos {1:N1} GB" -f ($os.TotalVisibleMemorySize / 1MB), ($os.FreePhysicalMemory / 1MB))
Yaz ("Commit: {0:N1} / {1:N1} GB (%{2}) | NonPaged havuz {3:N0} MB | Paged havuz {4:N0} MB" -f `
     ($m.CommittedBytes / 1GB), ($m.CommitLimit / 1GB), $m.PercentCommittedBytesInUse, ($m.PoolNonpagedBytes / 1MB), ($m.PoolPagedBytes / 1MB))
$c = Get-PSDrive C
Yaz ("C: bos {0:N1} GB" -f ($c.Free / 1GB))

# -- 1) Son 45 gunun kritik sistem olaylari (15.10 dahil) --
$aciklama = @{
    41 = 'Beklenmeyen kapanma / reset (Kernel-Power)'; 6008 = 'Beklenmeyen kapanma'; 6005 = 'Acilis'
    6006 = 'Duzgun kapanis'; 1074 = 'Yeniden baslatma istegi (kim/neden)'; 1001 = 'Mavi ekran (bugcheck)'
    2004 = 'DUSUK SANAL BELLEK - RAM/commit doldu'; 26 = 'Uygulama uyarisi (dusuk bellek olabilir)'
    2019 = 'NonPaged cekirdek havuzu bitti'; 2020 = 'Paged cekirdek havuzu bitti'
    4227 = 'TCP PORT TUKENMESI (TIME_WAIT)'; 4231 = 'TCP PORT TUKENMESI'
    333  = 'Registry yazilamadi (bellek)'; 7031 = 'Servis coktu'; 7034 = 'Servis beklenmedik durdu'
    7023 = 'Servis hatayla durdu'; 7011 = 'Servis zaman asimi (sistem yavas)'; 7009 = 'Servis baglanti zaman asimi'
    10016 = 'DCOM izin'; 36 = 'Volume Shadow / disk'; 153 = 'Disk IO yeniden denendi'; 129 = 'Disk resetlendi'
}
$idler = @(41, 6008, 6005, 6006, 1074, 1001, 2004, 26, 2019, 2020, 4227, 4231, 333, 7031, 7034, 7023, 7011, 7009, 153, 129)
$bas = (Get-Date).AddDays(-45)
$ev = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = $idler; StartTime = $bas })
$ev | Select-Object TimeCreated, Id, ProviderName, LevelDisplayName, Message |
    Export-Csv (Join-Path $dir 'sistem_olaylari.csv') -NoTypeInformation -Encoding UTF8
Yaz ""
Yaz "=== KRITIK SISTEM OLAYLARI (son 45 gun) ==="
if (-not $ev.Count) { Yaz "  (kayit yok)" }
$ev | Group-Object Id | Sort-Object Count -Descending | ForEach-Object {
    $son = ($_.Group | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated
    Yaz ("{0,6} x  ID {1,-5} {2,-45} son: {3}" -f $_.Count, $_.Name, $aciklama[[int]$_.Name], $son)
}

Yaz ""
Yaz "=== ACILIS / KAPANIS ZAMAN CIZELGESI ==="
$ev | Where-Object { $_.Id -in 41, 6008, 6005, 6006, 1074, 1001 } | Sort-Object TimeCreated | ForEach-Object {
    $msg = ($_.Message -replace '\s+', ' ')
    Yaz ("  {0:yyyy-MM-dd HH:mm}  ID {1,-5} {2}" -f $_.TimeCreated, $_.Id, $msg.Substring(0, [Math]::Min(160, $msg.Length)))
}

Yaz ""
Yaz "=== BELLEK / PORT / SERVIS OLAYLARININ AYRINTISI (en yeni 12) ==="
$ev | Where-Object { $_.Id -in 2004, 2019, 2020, 4227, 4231, 333, 7031, 7034, 7023, 26 } |
    Sort-Object TimeCreated -Descending | Select-Object -First 12 | ForEach-Object {
    $msg = ($_.Message -replace '\s+', ' ')
    Yaz ("  {0:yyyy-MM-dd HH:mm}  ID {1}: {2}" -f $_.TimeCreated, $_.Id, $msg.Substring(0, [Math]::Min(500, $msg.Length)))
}

# -- 2) Her beklenmeyen kapanmadan ONCEKI 6 saat: hata/uyari ureten kaynaklar --
Yaz ""
Yaz "=== KILITLENME ONCESI 6 SAATTE HATA / UYARI KAYNAKLARI ==="
$resetler = $ev | Where-Object { $_.Id -eq 41 } | Sort-Object TimeCreated
foreach ($r in $resetler) {
    # Kernel-Power 41 acilista yazilir; kilitlenme ondan once. Acilistan onceki son kayda kadar bak.
    $onceki = Get-WinEvent -FilterHashtable @{ LogName = 'System'; EndTime = $r.TimeCreated.AddMinutes(-1) } -MaxEvents 1
    if (-not $onceki) { continue }
    $son_kayit = $onceki.TimeCreated
    Yaz ("-- reset {0:yyyy-MM-dd HH:mm} | kilitlenmeden onceki son kayit {1:yyyy-MM-dd HH:mm}" -f $r.TimeCreated, $son_kayit)
    foreach ($log in 'System', 'Application') {
        $w = @(Get-WinEvent -FilterHashtable @{ LogName = $log; Level = 1, 2, 3; StartTime = $son_kayit.AddHours(-6); EndTime = $son_kayit.AddMinutes(1) })
        $w | Select-Object TimeCreated, Id, ProviderName, LevelDisplayName, Message |
            Export-Csv (Join-Path $dir ("oncesi_{0:yyyyMMdd_HHmm}_{1}.csv" -f $r.TimeCreated, $log)) -NoTypeInformation -Encoding UTF8
        $w | Group-Object ProviderName, Id | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object {
            $ilk = $_.Group | Sort-Object TimeCreated | Select-Object -First 1
            $msg = ($ilk.Message -replace '\s+', ' ')
            Yaz ("   [{0}] {1,5} x {2}  ilk {3:HH:mm}: {4}" -f $log, $_.Count, $_.Name, $ilk.TimeCreated, $msg.Substring(0, [Math]::Min(140, $msg.Length)))
        }
    }
}

# -- 3) Su anki durum (karsilastirma tabani) --
Yaz ""
Yaz "=== SU ANKI TCP BAGLANTILARI ==="
$tcp = @(Get-NetTCPConnection)
$tcp | Group-Object State | Sort-Object Count -Descending | ForEach-Object { Yaz ("  {0,-12} {1}" -f $_.Name, $_.Count) }
$dp = (netsh int ipv4 show dynamicport tcp) -join ' '
Yaz ("  Dinamik port araligi: " + ($dp -replace '\s+', ' '))
$adlar = @{}; Get-Process | ForEach-Object { $adlar[$_.Id] = $_.ProcessName }
Yaz "  En cok baglanti tutan surecler:"
$tcp | Group-Object OwningProcess | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
    $durum = ($_.Group | Group-Object State | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' '
    Yaz ("   {0,6}  {1,-22} pid {2,-6} {3}" -f $_.Count, $adlar[[int]$_.Name], $_.Name, $durum)
}

Yaz ""
Yaz "=== SURECLER: BELLEK / HANDLE (ilk 12) ==="
$p = Get-Process
Yaz "  -- ozel bellek (MB):"
$p | Sort-Object PrivateMemorySize64 -Descending | Select-Object -First 12 | ForEach-Object {
    Yaz ("   {0,8:N0}  {1,-24} pid {2,-6} handle {3}" -f ($_.PrivateMemorySize64 / 1MB), $_.ProcessName, $_.Id, $_.HandleCount) }
Yaz "  -- handle sayisi:"
$p | Sort-Object HandleCount -Descending | Select-Object -First 8 | ForEach-Object {
    Yaz ("   {0,8}  {1,-24} pid {2}" -f $_.HandleCount, $_.ProcessName, $_.Id) }

Yaz ""
Yaz "=== SERVISLER ==="
Get-Service cofle-app, cofle-pilot, Cloudflared, aspanda-app, aspanda-tunel, TermService |
    ForEach-Object { Yaz ("  {0,-16} {1,-8} {2}" -f $_.Name, $_.Status, $_.StartType) }

Yaz ""
Yaz "Ayrintili CSV'ler: $dir"
Start-Process notepad.exe $ozet

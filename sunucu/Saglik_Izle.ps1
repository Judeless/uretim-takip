# ============================================================
#  Cofle sunucu SAGLIK IZLEME (kullanici 2026-11-20: "server 2. kez kilitlendi,
#  ram mi doluyor, olmamasi icin ne yapalim")
#  Gorev Zamanlayici her 10 dakikada calistirir (Saglik_Izle_Kur.bat kurar).
#  C:\cofle\logs\saglik\saglik_YYYY-MM.csv'ye BIR satir ekler: bos RAM, commit %,
#  cekirdek havuzlari, TCP baglantilari (durum + en cok tutan surec), en cok
#  bellek / handle tutan surec. Bir sonraki kilitlenmede son satirlar hangi
#  kaynagin, hangi surec yuzunden tukendigini gosterir.
#  Esik asilinca UYARI.txt'ye satir yazar. Hicbir ayari degistirmez.
# ============================================================
$ErrorActionPreference = 'SilentlyContinue'
$dir = 'C:\cofle\logs\saglik'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$simdi = Get-Date
$csv = Join-Path $dir ("saglik_{0:yyyy-MM}.csv" -f $simdi)

$os = Get-CimInstance Win32_OperatingSystem
$m  = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
$p  = Get-Process
$adlar = @{}; $p | ForEach-Object { $adlar[$_.Id] = $_.ProcessName }
$tcp = @(Get-NetTCPConnection)
$durum = @{}; $tcp | Group-Object State | ForEach-Object { $durum[$_.Name] = $_.Count }

function Ilk($liste, $n) { ($liste | Select-Object -First $n) -join ' ' }
$tcp_surec = Ilk ($tcp | Group-Object OwningProcess | Sort-Object Count -Descending |
                  ForEach-Object { "{0}={1}" -f $adlar[[int]$_.Name], $_.Count }) 3
$tw_surec  = Ilk ($tcp | Where-Object { $_.State -eq 'TimeWait' -or $_.State -eq 'CloseWait' } | Group-Object OwningProcess |
                  Sort-Object Count -Descending | ForEach-Object { "{0}={1}" -f $adlar[[int]$_.Name], $_.Count }) 3
$bellek    = Ilk ($p | Sort-Object PrivateMemorySize64 -Descending |
                  ForEach-Object { "{0}={1:N0}MB" -f $_.ProcessName, ($_.PrivateMemorySize64 / 1MB) }) 4
$handle    = Ilk ($p | Sort-Object HandleCount -Descending | ForEach-Object { "{0}={1}" -f $_.ProcessName, $_.HandleCount }) 3
$python    = ($p | Where-Object { $_.ProcessName -like 'python*' } | Measure-Object PrivateMemorySize64 -Sum).Sum / 1MB
$cfd       = ($p | Where-Object { $_.ProcessName -like 'cloudflared*' } | Measure-Object PrivateMemorySize64 -Sum).Sum / 1MB

$satir = [pscustomobject]@{
    zaman          = $simdi.ToString('yyyy-MM-dd HH:mm')
    acik_saat      = [Math]::Round(($simdi - $os.LastBootUpTime).TotalHours, 1)
    ram_bos_mb     = [Math]::Round($os.FreePhysicalMemory / 1KB)
    commit_yuzde   = $m.PercentCommittedBytesInUse
    commit_gb      = [Math]::Round($m.CommittedBytes / 1GB, 2)
    nonpaged_mb    = [Math]::Round($m.PoolNonpagedBytes / 1MB)
    paged_mb       = [Math]::Round($m.PoolPagedBytes / 1MB)
    handle_toplam  = ($p | Measure-Object HandleCount -Sum).Sum
    surec_sayisi   = $p.Count
    tcp_toplam     = $tcp.Count
    tcp_established = [int]$durum['Established']
    tcp_timewait   = [int]$durum['TimeWait']
    tcp_closewait  = [int]$durum['CloseWait']
    tcp_en_cok     = $tcp_surec
    timewait_closewait_en_cok = $tw_surec
    bellek_en_cok  = $bellek
    handle_en_cok  = $handle
    python_mb      = [Math]::Round($python)
    cloudflared_mb = [Math]::Round($cfd)
}
$satir | Export-Csv -Path $csv -Append -NoTypeInformation -Encoding UTF8

# Esikler: tukenmeye YAKLASILDIGINDA iz birakir (kilitlenme aninda betik calismayabilir)
$uyari = @()
if ($satir.ram_bos_mb -lt 1024)        { $uyari += "bos RAM $($satir.ram_bos_mb) MB" }
if ($satir.commit_yuzde -ge 90)        { $uyari += "commit %$($satir.commit_yuzde)" }
if ($satir.nonpaged_mb -ge 1024)       { $uyari += "nonpaged havuz $($satir.nonpaged_mb) MB" }
if ($satir.tcp_toplam -ge 10000)       { $uyari += "TCP $($satir.tcp_toplam) baglanti ($tcp_surec)" }
if ($satir.tcp_timewait -ge 5000)      { $uyari += "TIME_WAIT $($satir.tcp_timewait) ($tw_surec)" }
if ($satir.tcp_closewait -ge 500)      { $uyari += "CLOSE_WAIT $($satir.tcp_closewait) ($tw_surec)" }
if ($satir.handle_toplam -ge 500000)   { $uyari += "toplam handle $($satir.handle_toplam) ($handle)" }
if ($uyari.Count) {
    ("{0}  {1}" -f $satir.zaman, ($uyari -join ' | ')) | Out-File -FilePath (Join-Path $dir 'UYARI.txt') -Append -Encoding utf8
}

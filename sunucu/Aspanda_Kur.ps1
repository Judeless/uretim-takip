# ============================================================
#  Aspanda Flow - Cofle sunucusunda AYRI kurulum (idempotent)
#  Klasor C:\aspanda, servisler aspanda-app (127.0.0.1:5050) + aspanda-tunel.
#  Cofle'nin klasorune, servislerine (cofle-app, cofle-pilot, Cloudflared)
#  ve verisine DOKUNMAZ. Kod C:\cofle\uretim_takip'ten (yerel git) alinir.
#  Calistirma: sunucu\Aspanda_Kur.bat (yonetici olarak kendini acar)
# ============================================================
$ErrorActionPreference = 'Continue'   # native stderr (git/nssm) scripti durdurmasin
$PY    = 'C:\Program Files\Python313\python.exe'
$NSSM  = 'C:\cofle\nssm.exe'
$GIT   = 'C:\Program Files\Git\cmd\git.exe'
$COFLE = 'C:\cofle\uretim_takip'
$HEDEF = 'C:\aspanda'
$LOG   = 'C:\aspanda\logs'
$PORT  = 5050
$TOKEN = 'C:\aspanda\data\tunel_token.txt'

function Dur($m) {
    Write-Host ''
    Write-Host ('[DURDU] ' + $m) -ForegroundColor Red
    Read-Host 'Kapatmak icin Enter'
    exit 1
}
function Baslik($m) { Write-Host ''; Write-Host ('=== ' + $m + ' ===') -ForegroundColor Cyan }

Baslik '0/6  On kontrol'
foreach ($y in @($PY, $NSSM, $GIT, $COFLE)) { if (-not (Test-Path $y)) { Dur ($y + ' bulunamadi') } }
if (-not (Test-Path (Join-Path $COFLE 'ornek_kurulum\aspanda.json'))) {
    Dur 'Sunucudaki kod eski (aspanda.json yok). Once laptopta Gonder.bat, sonra sunucuda Guncelle.bat.'
}
$CF = @('C:\Program Files (x86)\cloudflared\cloudflared.exe', 'C:\Program Files\cloudflared\cloudflared.exe') |
      Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $CF) {
    $svc = Get-CimInstance Win32_Service -Filter "Name='Cloudflared'" -ErrorAction SilentlyContinue
    if ($svc -and $svc.PathName -match '"?([^"]*cloudflared\.exe)') { $CF = $Matches[1] }
}
if (-not $CF -or -not (Test-Path $CF)) { Dur 'cloudflared.exe bulunamadi' }
Write-Host ('  cloudflared: ' + $CF)
$dinleyen = Get-NetTCPConnection -LocalPort $PORT -State Listen -ErrorAction SilentlyContinue
if ($dinleyen -and -not (Get-Service 'aspanda-app' -ErrorAction SilentlyContinue)) {
    Dur ('Port ' + $PORT + ' baska bir program tarafindan kullaniliyor')
}
Write-Host ''
Write-Host '  ONEMLI: Laptoptaki Aspanda pencereleri (Sunucu + Tunel) KAPALI olmali.' -ForegroundColor Yellow
Write-Host '  Ikisi ayni anda acik kalirsa ziyaretciler laptopa ve sunucuya bolunur.' -ForegroundColor Yellow
$c = Read-Host '  Laptoptaki Aspanda pencereleri kapatildi mi? (E/H)'
if ($c -notmatch '^[EeYy]') { Dur 'Once laptoptaki iki Aspanda penceresini kapatin.' }

Baslik '1/6  Kod (C:\cofle\uretim_takip -> C:\aspanda)'
if (Test-Path (Join-Path $HEDEF '.git')) {
    & $GIT -c safe.directory=* -C $HEDEF fetch origin
    & $GIT -c safe.directory=* -C $HEDEF reset --hard origin/main
} else {
    & $GIT -c safe.directory=* clone $COFLE $HEDEF
    if ($LASTEXITCODE -ne 0) { Dur 'git clone basarisiz' }
}
New-Item -ItemType Directory -Force -Path (Join-Path $HEDEF 'data'), $LOG | Out-Null

Baslik '2/6  Veri (Aspanda_Tasima.zip)'
if (Test-Path (Join-Path $HEDEF 'uretim.db')) {
    Write-Host '  [ATLA] C:\aspanda\uretim.db zaten var - mevcut veri KORUNDU, zip kullanilmadi.'
} else {
    $masaustu = [Environment]::GetFolderPath('Desktop')
    $zip = @((Join-Path $masaustu 'Aspanda_Tasima.zip'), (Join-Path $env:USERPROFILE 'Downloads\Aspanda_Tasima.zip'),
             (Join-Path $PSScriptRoot 'Aspanda_Tasima.zip'), 'C:\Aspanda_Tasima.zip') |
           Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $zip) {
        $zip = (Read-Host '  Aspanda_Tasima.zip bulunamadi. Tam yolunu yazin').Trim('"', ' ')
        if (-not (Test-Path $zip)) { Dur ('Dosya yok: ' + $zip) }
    }
    Write-Host ('  zip: ' + $zip)
    $gecici = Join-Path $env:TEMP 'aspanda_tasima'
    Remove-Item $gecici -Recurse -Force -ErrorAction SilentlyContinue
    Expand-Archive -Path $zip -DestinationPath $gecici -Force
    foreach ($d in @('uretim.db', 'data\kurulum.json')) {
        if (-not (Test-Path (Join-Path $gecici $d))) { Dur ('Zip icinde ' + $d + ' yok') }
    }
    Copy-Item (Join-Path $gecici 'uretim.db') (Join-Path $HEDEF 'uretim.db')
    Copy-Item (Join-Path $gecici 'data\*') (Join-Path $HEDEF 'data') -Recurse -Force
    if (Test-Path (Join-Path $gecici 'ILK_GIRIS.txt')) { Copy-Item (Join-Path $gecici 'ILK_GIRIS.txt') $HEDEF -Force }
    Remove-Item $gecici -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host '  [OK] uretim.db + data\ kopyalandi'
}
# kurulum.json YOKSA uygulama kendini COFLE sanar (Cofle tohumlari Aspanda veritabanina yazilir)
if (-not (Test-Path (Join-Path $HEDEF 'data\kurulum.json'))) { Dur 'C:\aspanda\data\kurulum.json YOK - servis kurulmadi' }

Baslik '3/6  Tunel token'
if (-not (Test-Path $TOKEN)) {
    Write-Host '  Cloudflare > Tunnels > aspanda: kurulum komutunu KOPYALA dugmesiyle alip yapistirin.'
    $t = Read-Host '  Komut/Token'
    $t = ($t -replace '^.*install\s+', '' -replace '^.*--token\s+', '').Trim()
    if (-not $t.StartsWith('eyJ')) { Dur 'Token eyJ ile baslamiyor' }
    [IO.File]::WriteAllText($TOKEN, $t)
}
Write-Host '  [OK] token dosyasi hazir'

Baslik '4/6  waitress'
& $PY -c 'import waitress' 2>$null
if ($LASTEXITCODE -ne 0) { & $PY -m pip install --quiet 'waitress==3.0.2' }
& $PY -c 'import waitress' 2>$null
$WSGI = ($LASTEXITCODE -eq 0)
Write-Host ('  waitress: ' + $WSGI)

Baslik '5/6  Servisler'
function Servis($ad, $exe, $param, $ortam) {
    if (Get-Service $ad -ErrorAction SilentlyContinue) {
        & $NSSM stop $ad | Out-Null
        & $NSSM remove $ad confirm | Out-Null
    }
    & $NSSM install $ad $exe | Out-Null
    & $NSSM set $ad AppParameters $param | Out-Null
    & $NSSM set $ad AppDirectory $HEDEF | Out-Null
    & $NSSM set $ad DisplayName ('Aspanda Flow - ' + $ad) | Out-Null
    & $NSSM set $ad Description 'Aspanda Flow uretim takip (Cofle sunucusunda ayri kurulum)' | Out-Null
    & $NSSM set $ad Start SERVICE_AUTO_START | Out-Null
    & $NSSM set $ad AppStdout ($LOG + '\' + $ad + '.out.log') | Out-Null
    & $NSSM set $ad AppStderr ($LOG + '\' + $ad + '.err.log') | Out-Null
    & $NSSM set $ad AppRotateFiles 1 | Out-Null
    & $NSSM set $ad AppRotateBytes 10485760 | Out-Null
    & $NSSM set $ad AppStopMethodConsole 5000 | Out-Null
    & $NSSM set $ad AppExit Default Restart | Out-Null
    & $NSSM set $ad AppRestartDelay 3000 | Out-Null
    if ($ortam) { & $NSSM set $ad AppEnvironmentExtra @ortam | Out-Null }
    Write-Host ('  [OK] ' + $ad)
}
$ortam = @('COFLE_HOST=127.0.0.1', ('COFLE_PORT=' + $PORT), 'COFLE_THREADS=8', 'PYTHONIOENCODING=utf-8', 'PYTHONUNBUFFERED=1')
if ($WSGI) { $ortam += 'COFLE_WSGI=waitress' }
Servis 'aspanda-app' $PY 'app.py' $ortam
Servis 'aspanda-tunel' $CF ('tunnel run --protocol http2 --token-file ' + $TOKEN) $null
& $NSSM start 'aspanda-app' | Out-Null
Start-Sleep -Seconds 3
& $NSSM start 'aspanda-tunel' | Out-Null

Baslik '6/6  Test'
$ok = $false
for ($i = 0; $i -lt 20 -and -not $ok; $i++) {
    Start-Sleep -Seconds 2
    try {
        $r = Invoke-WebRequest ('http://127.0.0.1:' + $PORT + '/dashboard') -UseBasicParsing -TimeoutSec 5
        $ok = ($r.StatusCode -eq 200)
    } catch { }
}
Get-Service 'aspanda-app', 'aspanda-tunel', 'cofle-app', 'Cloudflared' -ErrorAction SilentlyContinue |
    Format-Table Name, Status -AutoSize
if ($ok) {
    Write-Host '  [OK] Aspanda Flow calisiyor: http://127.0.0.1:5050  (internet: https://aspandaflow.com.tr)' -ForegroundColor Green
} else {
    Write-Host ('  [UYARI] 5050 cevap vermedi. Log: ' + $LOG + '\aspanda-app.err.log') -ForegroundColor Yellow
}
Write-Host '  Tasima zip dosyasini (token iceriyor) masaustunden SILIN.'
Read-Host 'Kapatmak icin Enter'

# ============================================================
#  Cofle Manage - Sunucu servis kurulumu (idempotent)
#  192.168.20.210 (CTKPRMNGSERVER) uzerinde calisir.
#  Yonetici PS: powershell -ExecutionPolicy Bypass -File C:\cofle\_sunucu_kurulum.ps1
# ============================================================
$ErrorActionPreference = 'Continue'   # native stderr (nssm) scripti durdurmasin
$PY   = 'C:\Program Files\Python313\python.exe'
$APP  = 'C:\cofle\uretim_takip'
$NSSM = 'C:\cofle\nssm.exe'
$LOG  = 'C:\cofle\logs'

New-Item -ItemType Directory -Force -Path $LOG | Out-Null

# --- NSSM indir (yoksa) ---
if (-not (Test-Path $NSSM)) {
    Write-Host 'NSSM indiriliyor...'
    $ProgressPreference = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        Invoke-WebRequest 'https://nssm.cc/release/nssm-2.24.zip' -OutFile 'C:\cofle\nssm.zip' -UseBasicParsing -ErrorAction Stop
        Expand-Archive 'C:\cofle\nssm.zip' -DestinationPath 'C:\cofle\nssm_pkg' -Force
        Copy-Item 'C:\cofle\nssm_pkg\nssm-2.24\win64\nssm.exe' $NSSM -Force
    } catch { Write-Host ('NSSM indirilemedi: ' + $_.Exception.Message); return }
}
Write-Host ('NSSM hazir: ' + (Test-Path $NSSM))

function Kur-Servis {
    param($ad, $script, $calismaDizini)
    if (Get-Service -Name $ad -ErrorAction SilentlyContinue) {
        Write-Host "  ($ad zaten var, yeniden kuruluyor)"
        & $NSSM stop $ad | Out-Null
        & $NSSM remove $ad confirm | Out-Null
        Start-Sleep -Seconds 1
    }
    & $NSSM install $ad $PY $script | Out-Null
    & $NSSM set $ad AppDirectory $calismaDizini | Out-Null
    & $NSSM set $ad DisplayName "Cofle Manage - $ad" | Out-Null
    & $NSSM set $ad Description "Cofle Manage uretim takip ($ad)" | Out-Null
    & $NSSM set $ad Start SERVICE_AUTO_START | Out-Null
    & $NSSM set $ad AppStdout "$LOG\$ad.out.log" | Out-Null
    & $NSSM set $ad AppStderr "$LOG\$ad.err.log" | Out-Null
    & $NSSM set $ad AppRotateFiles 1 | Out-Null
    & $NSSM set $ad AppRotateBytes 10485760 | Out-Null
    & $NSSM set $ad AppStopMethodConsole 5000 | Out-Null
    & $NSSM set $ad AppExit Default Restart | Out-Null
    & $NSSM set $ad AppRestartDelay 3000 | Out-Null
    Write-Host "  -> servis kuruldu: $ad"
}

Kur-Servis 'cofle-app'   "$APP\app.py"              $APP
Kur-Servis 'cofle-pilot' "$APP\pilot\pilot_app.py"  "$APP\pilot"

# --- Guvenlik duvari (5000/5001 gelen izin) ---
foreach ($p in 5000,5001) {
    $kn = "Cofle Manage $p"
    if (-not (Get-NetFirewallRule -DisplayName $kn -ErrorAction SilentlyContinue)) {
        New-NetFirewallRule -DisplayName $kn -Direction Inbound -Action Allow -Protocol TCP -LocalPort $p | Out-Null
        Write-Host "  -> firewall acildi: $p"
    } else {
        Write-Host "  -> firewall zaten var: $p"
    }
}

Write-Host ''
Write-Host '--- SERVISLER (henuz baslatilmadi) ---'
Get-Service cofle-app,cofle-pilot -ErrorAction SilentlyContinue | Select-Object Name,Status,StartType | Format-Table -AutoSize
Write-Host 'KURULUM TAMAM. Test icin: Start-Service cofle-app,cofle-pilot'

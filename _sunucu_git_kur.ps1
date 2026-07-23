# ============================================================
#  Sunucuda Git kurar + GitHub repoyu baglar + Guncelle.bat olusturur
#  Yonetici PowerShell'de calisir.
# ============================================================
$ErrorActionPreference='Continue'
$APP='C:\cofle\uretim_takip'
$REPO='https://github.com/Judeless/uretim-takip.git'
$gitExe='C:\Program Files\Git\cmd\git.exe'

# --- 1) Git for Windows kur (yoksa) ---
if (-not (Test-Path $gitExe)) {
    Write-Host 'Git indiriliyor...'
    $ProgressPreference='SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
    $rel = Invoke-RestMethod 'https://api.github.com/repos/git-for-windows/git/releases/latest' -UseBasicParsing
    $asset = $rel.assets | Where-Object { $_.name -match '64-bit\.exe$' -and $_.name -notmatch 'Portable|Mini|Symbols|tag' } | Select-Object -First 1
    Write-Host ('  -> ' + $asset.name)
    Invoke-WebRequest $asset.browser_download_url -OutFile 'C:\cofle\git-installer.exe' -UseBasicParsing
    Write-Host 'Git kuruluyor (sessiz, ~1 dk)...'
    Start-Process 'C:\cofle\git-installer.exe' -ArgumentList '/VERYSILENT','/NORESTART','/NOCANCEL','/SP-','/SUPPRESSMSGBOXES' -Wait
}
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
Write-Host ('Git: ' + (& $gitExe --version))

# --- 2) Repo baglama (kod dizinini repoya cevir) ---
Set-Location $APP
if (-not (Test-Path (Join-Path $APP '.git'))) { & $gitExe init -b main | Out-Null }
& $gitExe config --local core.autocrlf true
& $gitExe config --local user.email 'server@coflemanage.local'
& $gitExe config --local user.name 'Cofle Server'
& $gitExe remote remove origin 2>$null
& $gitExe remote add origin $REPO
Write-Host ''
Write-Host '>>> Simdi GitHub girisi penceresi acilabilir - JUDELESS ile onayla <<<'
Write-Host ''

# --- 3) fetch (AUTH: GCM penceresi burada acilir) ---
& $gitExe fetch origin main
Write-Host ('fetch exit: ' + $LASTEXITCODE)

# --- 4) kod = repo (veri/DB/config DOKUNULMAZ, onlar gitignore) ---
& $gitExe reset --hard origin/main
& $gitExe branch --set-upstream-to=origin/main main 2>$null
Write-Host ('HEAD: ' + (& $gitExe log --oneline -1))

# --- 5) Guncelle.bat olustur ---
$bat = @'
@echo off
REM === Cofle Manage - Sunucu Guncelleme ===
REM GitHub main'den son kodu ceker ve servisleri yeniden baslatir.
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb RunAs" & exit /b)
cd /d C:\cofle\uretim_takip
echo ============================================
echo   GitHub'dan son kod cekiliyor...
echo ============================================
"C:\Program Files\Git\cmd\git.exe" fetch origin main
"C:\Program Files\Git\cmd\git.exe" reset --hard origin/main
echo.
echo   Servisler yeniden baslatiliyor...
C:\cofle\nssm.exe restart cofle-app
C:\cofle\nssm.exe restart cofle-pilot
echo.
echo   GUNCELLEME TAMAMLANDI.
timeout /t 5
'@
Set-Content -Path 'C:\cofle\Guncelle.bat' -Value $bat -Encoding ASCII
Write-Host 'Guncelle.bat olusturuldu -> C:\cofle\Guncelle.bat'
Write-Host '=== BITTI ==='

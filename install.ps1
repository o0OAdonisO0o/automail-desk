# Automail Desk - one-click installer for Windows
# Usage: irm https://raw.githubusercontent.com/o0OAdonisO0o/automail-desk/main/install.ps1 | iex
# What it does: download the repo zip -> extract to %USERPROFILE%\AutomailDesk
#               -> create a virtual env -> install requirements -> write config.json -> start panel

$ErrorActionPreference = "Stop"
$Repo = "o0OAdonisO0o/automail-desk"
$Dest = Join-Path $env:USERPROFILE "AutomailDesk"
$Zip  = Join-Path $env:TEMP "automail-desk.zip"
# github.com (web) is unreachable on some networks while codeload / raw still work,
# so try the CDN endpoint first and fall back to the web endpoint.
$Urls = @(
    "https://codeload.github.com/$Repo/zip/refs/heads/main",
    "https://github.com/$Repo/archive/refs/heads/main.zip"
)

Write-Host "== Automail Desk installer ==" -ForegroundColor Cyan
Write-Host "target: $Dest"

# 1. locate a Python 3.9+
$py = $null
foreach ($cand in @("python", "py -3")) {
    try {
        $ver = & cmd /c "$cand -c `"import sys;print(sys.version_info[0]*100+sys.version_info[1])`"" 2>$null
        if ([int]$ver -ge 309) { $py = $cand; break }
    } catch { }
}
if (-not $py) {
    Write-Host "[ERROR] Python 3.9+ not found." -ForegroundColor Red
    Write-Host "Install from https://www.python.org/downloads/ and check 'Add python.exe to PATH'."
    exit 1
}
Write-Host "python : $py"

# 2. download and extract
$got = $false
foreach ($u in $Urls) {
    Write-Host "downloading $u"
    try {
        Invoke-WebRequest -Uri $u -OutFile $Zip -UseBasicParsing -TimeoutSec 180
        $got = $true
        break
    } catch {
        Write-Host "  failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
if (-not $got) {
    Write-Host "[ERROR] download failed on all endpoints. Check network or proxy settings." -ForegroundColor Red
    Write-Host "Manual: open https://codeload.github.com/$Repo/zip/refs/heads/main and unzip by hand."
    exit 1
}
if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
$tmp = Join-Path $env:TEMP ("automail_x_" + [guid]::NewGuid().ToString("N"))
Expand-Archive -Path $Zip -DestinationPath $tmp -Force
$inner = Get-ChildItem $tmp -Directory | Select-Object -First 1
Move-Item $inner.FullName $Dest
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $Zip -Force -ErrorAction SilentlyContinue

# 3. virtual env + dependencies
$venv = Join-Path $Dest ".venv"
& cmd /c "$py -m venv `"$venv`""
$vpy = Join-Path $venv "Scripts\python.exe"
if (-not (Test-Path $vpy)) {
    Write-Host "[ERROR] failed to create virtual env: $venv" -ForegroundColor Red
    exit 1
}
& $vpy -m pip install -q --upgrade pip
& $vpy -m pip install -q -r (Join-Path $Dest "requirements.txt")

# 4. config
$cfg = Join-Path $Dest "config.json"
if (-not (Test-Path $cfg)) {
    Copy-Item (Join-Path $Dest "config.example.json") $cfg
    Write-Host ""
    Write-Host "Generated $cfg" -ForegroundColor Yellow
    Write-Host "Edit it: domain / password / account_prefix / account_count." -ForegroundColor Yellow
    Write-Host "For non-Aliyun mailboxes also change smtp_host / imap_host and ports."
}

# 5. desktop shortcut (best effort)
try {
    $ws = New-Object -ComObject WScript.Shell
    $lnk = Join-Path ([Environment]::GetFolderPath("Desktop")) "Automail Desk.lnk"
    $sc = $ws.CreateShortcut($lnk)
    $sc.TargetPath = $vpy
    $sc.Arguments = '"' + (Join-Path $Dest "start_silent.py") + '"'
    $sc.WorkingDirectory = $Dest
    $sc.IconLocation = (Join-Path $Dest "automail.ico") + ",0"
    $sc.Description = "Automail Desk panel"
    $sc.Save()
    Write-Host "shortcut: $lnk"
} catch {
    Write-Host "shortcut not created (create it manually if needed)."
}

# 6. start
$port = 8765
try { $port = [int]((Get-Content $cfg -Raw | ConvertFrom-Json).port) } catch { }
Write-Host "starting panel on http://127.0.0.1:$port" -ForegroundColor Green
Start-Process -FilePath $vpy -ArgumentList ('"' + (Join-Path $Dest "start_silent.py") + '"') -WorkingDirectory $Dest -WindowStyle Hidden
for ($i = 0; $i -lt 40; $i++) {
    Start-Sleep -Milliseconds 500
    if (Test-NetConnection -ComputerName 127.0.0.1 -Port $port -InformationLevel Quiet -WarningAction SilentlyContinue) { break }
}
Start-Process "http://127.0.0.1:$port"
Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Write-Host "Start again : $vpy `"$Dest\start_silent.py`""
Write-Host "Stop        : $vpy `"$Dest\stop.py`"   (or the Stop button in the panel)"

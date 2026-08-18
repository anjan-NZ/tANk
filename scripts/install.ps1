# tANk installer
# Puts the add-in into your own copy of Office. Nothing is installed system wide,
# nothing needs administrator rights, and no data leaves your machine.
#
#   Right click this file and choose "Run with PowerShell",
#   or run:  powershell -ExecutionPolicy Bypass -File install.ps1

$ErrorActionPreference = "Stop"

$BaseUrl    = "__BASE_URL__"
$AddinId    = "7f3c1a54-9e2b-4d61-8a77-2c5b9e0f41d3"
$InstallDir = Join-Path $env:LOCALAPPDATA "tANk"
$Manifest   = Join-Path $InstallDir "manifest.xml"

Write-Host ""
Write-Host "  tANk - AI chat inside Excel, Word and PowerPoint" -ForegroundColor Green
Write-Host "  ------------------------------------------------"
Write-Host ""

# Office has to be closed, or it will not notice the new add-in.
$open = Get-Process EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue
if ($open) {
  Write-Host "  Close Excel, Word and PowerPoint first, then run this again." -ForegroundColor Yellow
  Write-Host "  Currently open: $(($open | Select-Object -ExpandProperty ProcessName -Unique) -join ', ')"
  Write-Host ""
  Read-Host "  Press Enter to exit"
  exit 1
}

Write-Host "  Downloading the add-in description..."
if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }
try {
  Invoke-WebRequest -Uri "$BaseUrl/manifest.xml" -OutFile $Manifest -UseBasicParsing
} catch {
  Write-Host "  Could not reach $BaseUrl - check your internet connection." -ForegroundColor Red
  Read-Host "  Press Enter to exit"
  exit 1
}

# This single registry value is the whole installation as far as Office is concerned.
Write-Host "  Registering with Office..."
$Key = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }
New-ItemProperty -Path $Key -Name $AddinId -Value $Manifest -PropertyType String -Force | Out-Null

# Office caches the ribbon buttons, so a stale cache would hide a fresh install. Only these
# two folders are cleared, and Office rebuilds both on its next start. The neighbouring
# webview2 folder, where every add-in on this PC keeps its saved data, is never touched.
if ($env:LOCALAPPDATA) {
  $Wef = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
  foreach ($cache in @("AppCommands", "AddinInfo")) {
    $path = Join-Path $Wef $cache
    if (Test-Path $path) {
      try {
        Get-ChildItem $path -Recurse -Force -ErrorAction Stop |
          Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
      } catch {}
    }
  }
}

Write-Host ""
Write-Host "  Done." -ForegroundColor Green
Write-Host ""
Write-Host "  Open Excel, Word or PowerPoint, then look on the Home tab for the tANk button."
Write-Host "  First time: open Settings in the pane and paste a free API key."
Write-Host "  The Keys section has a button that opens each provider's signup page."
Write-Host ""
Write-Host "  To remove it later, run uninstall.ps1 from the same place."
Write-Host ""
Read-Host "  Press Enter to close"

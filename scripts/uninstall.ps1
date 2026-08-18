# Removes tANk from your copy of Office.
#
#   Right click this file and choose "Run with PowerShell",
#   or run:  powershell -ExecutionPolicy Bypass -File uninstall.ps1

$ErrorActionPreference = "Stop"

$AddinId    = "7f3c1a54-9e2b-4d61-8a77-2c5b9e0f41d3"
$InstallDir = Join-Path $env:LOCALAPPDATA "tANk"
$Key        = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"

Write-Host ""
Write-Host "  Removing tANk..."

$open = Get-Process EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue
if ($open) {
  Write-Host "  Close Excel, Word and PowerPoint first, then run this again." -ForegroundColor Yellow
  Read-Host "  Press Enter to exit"
  exit 1
}

# The registration.
if (Test-Path $Key) {
  try { Remove-ItemProperty -Path $Key -Name $AddinId -ErrorAction Stop } catch {}
}

# The files. Guarded, so an unusual environment cannot widen this into something else.
if ($env:LOCALAPPDATA -and (Split-Path $InstallDir -Leaf) -eq "tANk" -and (Test-Path $InstallDir)) {
  Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
}

# The cached ribbon button. Wef\webview2 is left alone: every add-in keeps its saved data there.
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

Write-Host "  Done. The button will be gone next time Office starts." -ForegroundColor Green
Write-Host ""
Write-Host "  One thing left: any API keys you pasted are stored by the pane itself, inside"
Write-Host "  Office's web storage, which this script cannot pick apart without wiping the"
Write-Host "  saved data of every other add-in on this PC."
Write-Host ""
Write-Host "  If you still have tANk installed, open it and press Settings, then Erase"
Write-Host "  everything. That clears the keys and every setting straight away."
Write-Host ""
Read-Host "  Press Enter to close"

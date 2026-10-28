<#
Development pipeline:  [export] -> check -> build dev copy -> open in Excel

  ./scripts/dev.ps1                 check, build dist\MayaTools-dev.xlam, open it in Excel
  ./scripts/dev.ps1 -NoOpen         same, without opening Excel
  ./scripts/dev.ps1 -Export         first pull your VBE edits from the dev copy into src\
  ./scripts/dev.ps1 -CloseExcel     close running Excel instances instead of stopping
  ./scripts/dev.ps1 -Force          rebuild even if the dev copy holds edits never exported
#>
param(
    [switch]$Export,
    [switch]$NoOpen,
    [switch]$Force,
    [switch]$CloseExcel
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

$devFile = Join-Path $root "dist\MayaTools-dev.xlam"
$state   = Join-Path $root "dist\.dev-state"

function Write-Step($text) { Write-Host "`n== $text" -ForegroundColor Cyan }

# Excel must be closed: it locks the dev file, and the scripts start their own hidden instance
$running = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($running) {
    if ($CloseExcel) { $running | Stop-Process -Force; Start-Sleep -Seconds 1 }
    else { throw "Excel is running. Save your work and close it (or use -CloseExcel). A hidden Excel process left over from an earlier run also counts." }
}

# 1. Optional: pull edits made in the VBE back into src\
if ($Export) {
    Write-Step "Export VBE edits -> src"
    if (-not (Test-Path $devFile)) { throw "No dev copy at $devFile. Run ./scripts/dev.ps1 first." }
    & "$PSScriptRoot/export.ps1" -Xlam $devFile
    Set-Content $state (Get-Item $devFile).LastWriteTimeUtc.Ticks
}

# 2. Source checks
Write-Step "Source checks"
$global:LASTEXITCODE = 0
& "$PSScriptRoot/check.ps1"
if ($LASTEXITCODE -ne 0) { throw "Source checks failed. Fix the problems above, then run dev.ps1 again." }

# 3. Build the dev copy (and open it)
Write-Step "Build dev copy"
& "$PSScriptRoot/build-dev.ps1" -NoOpen:$NoOpen -Force:$Force

Write-Host "`nAfter editing in the VBE: save, close Excel, then run ./scripts/dev.ps1 -Export" -ForegroundColor Yellow
<#
Builds dist\MayaTools-dev.xlam from src\ and opens it in Excel.
Refuses to overwrite the dev copy if it was saved after the last build or export,
because that means it holds VBE edits that are not in src\ yet.
#>
param(
    [switch]$NoOpen,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent

$ver     = (Get-Content (Join-Path $root "VERSION") -Raw).Trim()
$devFile = Join-Path $root "dist\MayaTools-dev.xlam"
$state   = Join-Path $root "dist\.dev-state"

# Guard: has the dev copy been saved since the last build/export?
if ((Test-Path $devFile) -and (Test-Path $state) -and -not $Force) {
    $saved = (Get-Content $state -Raw).Trim()
    $now   = (Get-Item $devFile).LastWriteTimeUtc.Ticks.ToString()
    if ($saved -ne $now) {
        throw "The dev copy was saved after the last build. Run ./scripts/dev.ps1 -Export to keep those edits, or -Force to discard them."
    }
}

& "$PSScriptRoot/build.ps1" -Version "$ver-dev" -Out "dist\MayaTools-dev.xlam"

Set-Content $state (Get-Item $devFile).LastWriteTimeUtc.Ticks

if (-not $NoOpen) {
    Start-Process $devFile
    Write-Host "Opened $devFile. Edit in the VBE (Alt+F11)." 
}
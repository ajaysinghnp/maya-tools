<#
Builds dist\MayaTools-dev.xlam from src\ and opens it in Excel.
Refuses to overwrite the dev copy if it was saved after the last build or export,
because that means it holds VBE edits that are not in src\ yet.

The dev build is labelled with the NEXT version, not the last released one:
VERSION 2.0.1 (released) -> the dev build is 2.0.2-dev.
Use -NextVersion when the next release will not be a patch bump, for example -NextVersion 2.1.0.
#>
param(
    [switch]$NoOpen,
    [switch]$Force,
    [string]$NextVersion
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent

$ver     = (Get-Content (Join-Path $root "VERSION") -Raw).Trim()
$devFile = Join-Path $root "dist\MayaTools-dev.xlam"
$state   = Join-Path $root "dist\.dev-state"

# Work out the dev version label
if ($NextVersion) {
    $base = $NextVersion.Trim().TrimStart("v")
    if ($base -notmatch '^\d+\.\d+\.\d+$') { throw "-NextVersion '$NextVersion' must look like 2.1.0" }
}
else {
    if ($ver -notmatch '^(\d+)\.(\d+)\.(\d+)(-.+)?$') { throw "VERSION '$ver' is not a valid version (expected 2.1.0 or 2.1.0-beta)." }
    if ($Matches[4]) {
        # last release was a pre-release such as 2.1.0-beta: the next release is 2.1.0 itself
        $base = "{0}.{1}.{2}" -f $Matches[1], $Matches[2], $Matches[3]
    }
    else {
        $base = "{0}.{1}.{2}" -f $Matches[1], $Matches[2], ([int]$Matches[3] + 1)
    }
}
$devVersion = "$base-dev"

# Guard: has the dev copy been saved since the last build/export?
if ((Test-Path $devFile) -and (Test-Path $state) -and -not $Force) {
    $saved = (Get-Content $state -Raw).Trim()
    $now   = (Get-Item $devFile).LastWriteTimeUtc.Ticks.ToString()
    if ($saved -ne $now) {
        throw "The dev copy was saved after the last build. Run ./scripts/dev.ps1 -Export to keep those edits, or -Force to discard them."
    }
}

& "$PSScriptRoot/build.ps1" -Version $devVersion -Out "dist\MayaTools-dev.xlam"

Set-Content $state (Get-Item $devFile).LastWriteTimeUtc.Ticks

if (-not $NoOpen) {
    Start-Process $devFile
    Write-Host "Opened $devFile. Edit in the VBE (Alt+F11)." 
}
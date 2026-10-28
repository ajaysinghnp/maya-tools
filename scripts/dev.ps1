<#
Development pipeline:
  [export] -> check -> build dev copy -> [git diff -> git add -> git commit] -> open in Excel

  ./scripts/dev.ps1                                  check, build dist\MayaTools-dev.xlam, open it
  ./scripts/dev.ps1 -Export                          first pull your VBE edits from the dev copy into src\
  ./scripts/dev.ps1 -Export -Message "aboutDeveloper: add version label"
                                                     ... and commit with that message after a green build
  ./scripts/dev.ps1 -Export -Commit                  ... and ask for the commit message
  ./scripts/dev.ps1 -Message "..." -Yes              commit without the confirmation prompt
  ./scripts/dev.ps1 -NoOpen                          do not open Excel at the end
  ./scripts/dev.ps1 -CloseExcel                      close running Excel instances instead of stopping
  ./scripts/dev.ps1 -Force                           rebuild even if the dev copy holds edits never exported

Commits cover src\, ribbon\ and references.txt only. Scripts, VERSION and workflows are committed separately.
#>
param(
    [switch]$Export,
    [switch]$Commit,
    [string]$Message,
    [switch]$Yes,
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

function Invoke-Git {
    & git @args
    if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') failed (exit code $LASTEXITCODE)" }
}

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

# 3. Build the dev copy (opening it is done at the end, after the commit step)
Write-Step "Build dev copy"
& "$PSScriptRoot/build-dev.ps1" -NoOpen -Force:$Force

# 4. Review and commit
if ($Commit -or $Message) {
    Write-Step "Review changes"
    $paths = @("src", "ribbon", "references.txt") | Where-Object { Test-Path (Join-Path $root $_) }

    Invoke-Git add -- @paths
    $changed = git diff --cached --name-only -- @paths

    if (-not $changed) {
        Write-Host "Nothing to commit."
    }
    else {
        Invoke-Git --no-pager diff --cached --stat -- @paths
        Write-Host ""
        Invoke-Git --no-pager diff --cached -- @paths

        if (-not $Message) {
            $Message = Read-Host "Commit message (for example: aboutDeveloper: <what changed>)"
        }

        if (-not $Message -or -not $Message.Trim()) {
            Write-Host "No commit message given. Changes stay staged, nothing was committed." -ForegroundColor Yellow
        }
        else {
            $go = $true
            if (-not $Yes) {
                $answer = Read-Host "Commit these changes as `"$Message`"? [Y/n]"
                if ($answer -match '^(n|no)$') { $go = $false }
            }
            if ($go) {
                Invoke-Git commit -m $Message -- @paths
                Write-Host "Committed. Push with: git push" -ForegroundColor Green
            }
            else {
                Write-Host "Not committed. Changes stay staged." -ForegroundColor Yellow
            }
        }
    }
}

# 5. Open the dev copy
if (-not $NoOpen) {
    Start-Process $devFile
    Write-Host "`nOpened $devFile. Edit in the VBE (Alt+F11), save, close Excel, then run ./scripts/dev.ps1 -Export" -ForegroundColor Yellow
}
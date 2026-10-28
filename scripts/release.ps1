param([Parameter(Mandatory)][string]$Version)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

$Version = $Version.TrimStart("v")
if ($Version -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.]+)?$') { throw "Version must look like 2.1.0 or 2.1.0-beta" }
$tag = "v$Version"

if (git tag --list $tag) { throw "Tag $tag already exists locally. Pick a new version (or delete the tag first)." }
if (git ls-remote --tags origin $tag) { throw "Tag $tag already exists on GitHub." }

# 1. working tree must be clean so the build matches the commit
if (git status --porcelain) { throw "Uncommitted changes. Commit or stash them first." }

# 2. VERSION file
if ((Get-Content VERSION -Raw).Trim() -ne $Version) {
    Set-Content VERSION $Version -NoNewline
}

# 3. checks, then build
& "$PSScriptRoot/check.ps1"
if ($LASTEXITCODE -ne 0) { throw "Source checks failed." }

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
& "$PSScriptRoot/build.ps1" -Version $Version
$built = Join-Path $root "dist/MayaTools-$tag.xlam"
if (-not (Test-Path $built)) { throw "Build output not found: $built" }

# 4. copy the build into the repo, commit, tag, push
New-Item -ItemType Directory -Force releases | Out-Null
Copy-Item $built "releases/MayaTools-$tag.xlam" -Force

git add VERSION releases
git commit -m "Release $Version"
git tag $tag
git push
git push origin $tag

Write-Host "Pushed $tag. GitHub Actions will publish the release."
Write-Host "Release process completed successfully." -ForegroundColor Green
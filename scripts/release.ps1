param(
    [Parameter(Mandatory)][string]$Version,
    [switch]$Draft,
    [switch]$SkipUpload      # build, tag and push only; attach the file in the browser
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

$Version = $Version.TrimStart("v")
if ($Version -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.]+)?$') { throw "Version must look like 2.1.0 or 2.1.0-beta" }
$tag = "v$Version"

# owner/repo from the git remote (https or ssh form)
$remote = git remote get-url origin
if ($remote -notmatch 'github\.com[:/](?<repo>[^/]+/[^/]+?)(\.git)?$') { throw "Cannot read owner/repo from remote: $remote" }
$repo = $Matches.repo

# 1. VERSION file
$current = (Get-Content VERSION -Raw).Trim()
if ($current -ne $Version) {
    Set-Content VERSION $Version -NoNewline
    git add VERSION
    git commit -m "Release $Version"
}

# 2. working tree must be clean so the release matches the commit
if (git status --porcelain) { throw "Uncommitted changes. Commit or stash them first." }

# 3. checks and build
& "$PSScriptRoot\check.ps1"
if ($LASTEXITCODE -ne 0) { throw "Source checks failed." }
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
& "$PSScriptRoot\build.ps1" -Version $Version
$asset = Join-Path $root "dist\MayaTools-$tag.xlam"
if (-not (Test-Path $asset)) { throw "Build output not found: $asset" }

# 4. tag and push
if (-not (git tag --list $tag)) { git tag $tag }
git push
git push origin $tag

if ($SkipUpload) {
    Write-Host "Tag pushed. Attach this file on the Releases page: $asset"
    return
}

# 5. create the release and upload the asset through the REST API
$token = $env:GITHUB_TOKEN
if (-not $token) {
    $secure = Read-Host "GitHub token (Enter to skip the upload)" -AsSecureString
    $token  = [Net.NetworkCredential]::new("", $secure).Password
}
if (-not $token) {
    Write-Host "Skipped upload. Attach this file on the Releases page: $asset"
    return
}

$headers = @{
    Authorization          = "Bearer $token"
    Accept                 = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
}
$body = @{
    tag_name               = $tag
    name                   = $tag
    draft                  = [bool]$Draft
    prerelease             = ($Version -match '-')
    generate_release_notes = $true
} | ConvertTo-Json

$release = Invoke-RestMethod -Method Post -Uri "https://api.github.com/repos/$repo/releases" `
    -Headers $headers -Body $body -ContentType "application/json"

$uploadUrl = ($release.upload_url -replace '\{.*$', '') + "?name=" + [Uri]::EscapeDataString((Split-Path $asset -Leaf))
Invoke-RestMethod -Method Post -Uri $uploadUrl -Headers $headers `
    -InFile $asset -ContentType "application/octet-stream" | Out-Null

Write-Host "Released $tag -> $($release.html_url)"
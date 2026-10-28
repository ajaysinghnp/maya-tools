$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root
$errors = @()

# VERSION format
$v = (Get-Content VERSION -Raw).Trim()
if ($v -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.]+)?$') { $errors += "VERSION is not valid: '$v'" }

# pure-ASCII VBA source
foreach ($f in Get-ChildItem src -Recurse -Include *.bas,*.cls,*.frm -File) {
    $n = ([IO.File]::ReadAllBytes($f.FullName) | Where-Object { $_ -gt 127 }).Count
    if ($n) { $errors += "$($f.Name): $n non-ASCII bytes (use ChrW)" }
}

# every .frm needs its .frx
foreach ($f in Get-ChildItem src -Recurse -Filter *.frm -File) {
    if (-not (Test-Path ([IO.Path]::ChangeExtension($f.FullName, ".frx")))) { $errors += "$($f.Name) has no .frx" }
}

# ribbon XML is valid and every callback has a procedure
$procs = Get-ChildItem src -Recurse -Include *.bas,*.cls -File |
    Select-String -Pattern '^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)' |
    ForEach-Object { $_.Matches[0].Groups[1].Value }

foreach ($x in "customUI.xml", "customUI14.xml") {
    $p = "ribbon/$x"
    if (-not (Test-Path $p)) { continue }
    try { $xml = [xml](Get-Content $p -Raw) } catch { $errors += "$x is not valid XML"; continue }
    $callbacks = $xml.SelectNodes("//@*") | Where-Object { $_.Name -match '^(on|get|load)' } |
        ForEach-Object { $_.Value } | Sort-Object -Unique
    foreach ($c in $callbacks) {
        if ($c -notin $procs) { $errors += "$x : callback '$c' has no matching procedure" }
    }
}

if ($errors) { $errors | ForEach-Object { Write-Host "FAIL: $_" -ForegroundColor Red }; exit 1 }
Write-Host "All checks passed." -ForegroundColor Green
# Checks the ribbon icons. Called by check.ps1; can also be run on its own.
#   - every image="..." in ribbon/customUI*.xml has a Relationship in ribbon/_rels/<file>.rels
#   - the target file exists and is a PNG
#   - the PNG is 32x32 and 32-bit RGBA (transparent background)
$ErrorActionPreference = 'Stop'

$root   = Split-Path $PSScriptRoot -Parent
$ribbon = Join-Path $root 'ribbon'
$errors = [System.Collections.Generic.List[string]]::new()
$pngSignature = [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
$expectedSize = 32

foreach ($xmlFile in Get-ChildItem -LiteralPath $ribbon -Filter 'customUI*.xml') {
    $name = $xmlFile.Name
    [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw

    $ids = @($xml.SelectNodes('//*[@image]') | ForEach-Object { $_.GetAttribute('image') } | Sort-Object -Unique)

    $relsPath = Join-Path $ribbon "_rels/$name.rels"
    if (-not (Test-Path -LiteralPath $relsPath)) {
        $errors.Add("$name`: missing ribbon/_rels/$name.rels")
        continue
    }
    [xml]$rels = Get-Content -LiteralPath $relsPath -Raw
    $targets = @{}
    foreach ($rel in $rels.DocumentElement.ChildNodes) {
        if ($rel.NodeType -eq 'Element') { $targets[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target') }
    }

    foreach ($id in $ids) {
        if (-not $targets.ContainsKey($id)) {
            $errors.Add("$name`: image id '$id' is not listed in $name.rels")
            continue
        }
        $file = Join-Path $ribbon $targets[$id]
        if (-not (Test-Path -LiteralPath $file)) {
            $errors.Add("$name`: '$id' points to a missing file ($($targets[$id]))")
            continue
        }

        $bytes = [System.IO.File]::ReadAllBytes($file)
        if ($bytes.Length -lt 26) {
            $errors.Add("$($targets[$id]): file is too small to be a PNG")
            continue
        }
        $isPng = $true
        for ($i = 0; $i -lt 8; $i++) { if ($bytes[$i] -ne $pngSignature[$i]) { $isPng = $false } }
        if (-not $isPng) {
            $errors.Add("$($targets[$id]): not a PNG file")
            continue
        }

        $width  = ([int]$bytes[16] -shl 24) -bor ([int]$bytes[17] -shl 16) -bor ([int]$bytes[18] -shl 8) -bor [int]$bytes[19]
        $height = ([int]$bytes[20] -shl 24) -bor ([int]$bytes[21] -shl 16) -bor ([int]$bytes[22] -shl 8) -bor [int]$bytes[23]
        $bitDepth  = $bytes[24]
        $colorType = $bytes[25]

        if ($width -ne $expectedSize -or $height -ne $expectedSize) {
            $errors.Add("$($targets[$id]): is ${width}x${height}, expected ${expectedSize}x${expectedSize}")
        }
        if ($colorType -ne 6 -or $bitDepth -ne 8) {
            $errors.Add("$($targets[$id]): must be 32-bit RGBA with transparency (found color type $colorType, bit depth $bitDepth)")
        }
    }
}

if ($errors.Count -gt 0) {
    Write-Host 'Ribbon icon checks failed:' -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

Write-Host 'Ribbon icon checks passed.' -ForegroundColor Green
exit 0
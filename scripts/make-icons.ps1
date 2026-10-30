<#
.SYNOPSIS
    Generates the ribbon icons from the Lucide icon set. PowerShell only, no Node.

.DESCRIPTION
    Reads scripts/icons.json (image id -> Lucide icon name), takes the SVG of each icon from
    scripts/icons-src/ (downloaded from the pinned Lucide version if missing), draws it in one
    colour on a transparent background and writes:

      ribbon/images/<id>.png            32x32, 32-bit RGBA
      ribbon/_rels/customUI.xml.rels    one Relationship per image id
      ribbon/_rels/customUI14.xml.rels  (same content)
      THIRD-PARTY-NOTICES.md            the Lucide license notice

    Drawing uses WPF, which is built into Windows (works in Windows PowerShell 5.1 and PowerShell 7).
    No Excel is needed.

    Run from the repo root:   ./scripts/make-icons.ps1
    If PowerShell 7 cannot load WPF: powershell -ExecutionPolicy Bypass -File scripts\make-icons.ps1

.PARAMETER NoPng
    Parse the SVGs and write the .rels and notice files, but do not draw PNGs. For testing on a machine
    without WPF.

.PARAMETER DebugSvgDir
    For testing: also write each icon, converted to plain paths, as an SVG into this folder.
#>
[CmdletBinding()]
param(
    [switch]$NoPng,
    [string]$DebugSvgDir
)

$ErrorActionPreference = 'Stop'
$inv      = [System.Globalization.CultureInfo]::InvariantCulture
$utf8     = New-Object System.Text.UTF8Encoding($false)
$root     = Split-Path $PSScriptRoot -Parent
$srcDir   = Join-Path $PSScriptRoot 'icons-src'
$ribbon   = Join-Path $root 'ribbon'
$imgDir   = Join-Path $ribbon 'images'
$relsDir  = Join-Path $ribbon '_rels'
$ribbonXml = @('customUI.xml', 'customUI14.xml')

$cfg = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'icons.json') -Raw | ConvertFrom-Json
$size    = [int]$cfg.size
$color   = [string]$cfg.color
$version = [string]$cfg.lucideVersion
$map     = [ordered]@{}
foreach ($p in $cfg.icons.PSObject.Properties) { $map[$p.Name] = [string]$p.Value }

# ---------------------------------------------------------------- helpers

function Format-Num([double]$Value) { return $Value.ToString('0.######', $inv) }

function Get-Attr([System.Xml.XmlElement]$El, [string]$Name, [double]$Default = 0) {
    $v = $El.GetAttribute($Name)
    if ([string]::IsNullOrEmpty($v)) { return $Default }
    return [double]::Parse($v, $inv)
}

# Rewrites SVG path data so every command and number is separated by a space.
# This also splits arc flags that are glued to the next number (for example "a2 2 0 012 2").
function ConvertTo-CleanPathData([string]$D) {
    $cmdRe  = [regex]'\G[\s,]*([MmLlHhVvCcSsQqTtAaZz])'
    $numRe  = [regex]'\G[\s,]*([-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?)'
    $flagRe = [regex]'\G[\s,]*([01])'
    $argCount = @{ M = 2; L = 2; H = 1; V = 1; C = 6; S = 4; Q = 4; T = 2; A = 7; Z = 0 }
    $pos = 0
    $sb = New-Object System.Text.StringBuilder
    while ($true) {
        $m = $cmdRe.Match($D, $pos)
        if (-not $m.Success) {
            if ($D.Substring($pos).Trim(" ,`t`r`n").Length -eq 0) { break }
            throw "Unexpected text in path data at position ${pos}: '$($D.Substring($pos))'"
        }
        $cmd = $m.Groups[1].Value
        $pos += $m.Length
        [void]$sb.Append($cmd)
        $n = $argCount[$cmd.ToUpperInvariant()]
        if ($n -eq 0) { [void]$sb.Append(' '); continue }
        while ($true) {
            $values = New-Object System.Collections.Generic.List[string]
            $start = $pos
            $i = 0
            $complete = $true
            for ($i = 0; $i -lt $n; $i++) {
                $re = $numRe
                if ($cmd -eq 'a' -or $cmd -eq 'A') { if ($i -eq 3 -or $i -eq 4) { $re = $flagRe } }
                $mm = $re.Match($D, $pos)
                if (-not $mm.Success) { $complete = $false; break }
                $values.Add($mm.Groups[1].Value)
                $pos += $mm.Length
            }
            if (-not $complete) {
                if ($i -ne 0) { throw "Incomplete '$cmd' command in path data near position $pos" }
                $pos = $start
                break
            }
            [void]$sb.Append(' ' + ($values -join ' '))
        }
        [void]$sb.Append(' ')
    }
    return $sb.ToString().Trim()
}

# Converts one SVG shape element to path data.
function Get-ShapePathData([System.Xml.XmlElement]$El) {
    $fill = $El.GetAttribute('fill')
    if ($fill -and $fill -ne 'none') { throw "<$($El.LocalName)> has fill='$fill'. Only stroked shapes are supported." }

    switch ($El.LocalName) {
        'path' { return ConvertTo-CleanPathData $El.GetAttribute('d') }
        'circle' {
            $cx = Get-Attr $El 'cx'; $cy = Get-Attr $El 'cy'; $r = Get-Attr $El 'r'
            return ('M {0} {1} A {2} {2} 0 1 0 {3} {1} A {2} {2} 0 1 0 {0} {1} Z' -f (Format-Num ($cx - $r)), (Format-Num $cy), (Format-Num $r), (Format-Num ($cx + $r)))
        }
        'ellipse' {
            $cx = Get-Attr $El 'cx'; $cy = Get-Attr $El 'cy'; $rx = Get-Attr $El 'rx'; $ry = Get-Attr $El 'ry'
            return ('M {0} {1} A {2} {3} 0 1 0 {4} {1} A {2} {3} 0 1 0 {0} {1} Z' -f (Format-Num ($cx - $rx)), (Format-Num $cy), (Format-Num $rx), (Format-Num $ry), (Format-Num ($cx + $rx)))
        }
        'line' {
            return ('M {0} {1} L {2} {3}' -f (Format-Num (Get-Attr $El 'x1')), (Format-Num (Get-Attr $El 'y1')), (Format-Num (Get-Attr $El 'x2')), (Format-Num (Get-Attr $El 'y2')))
        }
        'rect' {
            $x = Get-Attr $El 'x'; $y = Get-Attr $El 'y'; $w = Get-Attr $El 'width'; $h = Get-Attr $El 'height'
            $rx = Get-Attr $El 'rx' (-1); $ry = Get-Attr $El 'ry' (-1)
            if ($rx -lt 0 -and $ry -lt 0) { $rx = 0; $ry = 0 }
            elseif ($rx -lt 0) { $rx = $ry }
            elseif ($ry -lt 0) { $ry = $rx }
            if ($rx -gt $w / 2) { $rx = $w / 2 }
            if ($ry -gt $h / 2) { $ry = $h / 2 }
            $x0 = Format-Num $x; $y0 = Format-Num $y
            $x1 = Format-Num ($x + $rx); $x2 = Format-Num ($x + $w - $rx); $x3 = Format-Num ($x + $w)
            $y1 = Format-Num ($y + $ry); $y2 = Format-Num ($y + $h - $ry); $y3 = Format-Num ($y + $h)
            if ($rx -le 0 -or $ry -le 0) {
                return "M $x0 $y0 H $x3 V $y3 H $x0 Z"
            }
            $arc = 'A {0} {1} 0 0 1' -f (Format-Num $rx), (Format-Num $ry)
            return "M $x1 $y0 H $x2 $arc $x3 $y1 V $y2 $arc $x2 $y3 H $x1 $arc $x0 $y2 V $y1 $arc $x1 $y0 Z"
        }
        { $_ -eq 'polyline' -or $_ -eq 'polygon' } {
            $nums = @($El.GetAttribute('points').Trim() -split '[\s,]+' | Where-Object { $_ -ne '' } | ForEach-Object { [double]::Parse($_, $inv) })
            if ($nums.Count -lt 4 -or ($nums.Count % 2) -ne 0) { throw "<$($El.LocalName)> has invalid points" }
            $sb = New-Object System.Text.StringBuilder
            for ($i = 0; $i -lt $nums.Count; $i += 2) {
                $cmd = 'L'; if ($i -eq 0) { $cmd = 'M' }
                [void]$sb.Append(('{0} {1} {2} ' -f $cmd, (Format-Num $nums[$i]), (Format-Num $nums[$i + 1])))
            }
            if ($El.LocalName -eq 'polygon') { [void]$sb.Append('Z') }
            return $sb.ToString().Trim()
        }
        default { throw "Unsupported SVG element <$($El.LocalName)>" }
    }
}

# Returns @{ Paths = <string[]>; StrokeWidth = <double>; ViewBox = <double> } for an SVG file.
function Read-IconSvg([string]$File) {
    [xml]$svg = Get-Content -LiteralPath $File -Raw -Encoding UTF8
    $rootEl = $svg.DocumentElement
    $strokeWidth = 2.0
    if ($rootEl.GetAttribute('stroke-width')) { $strokeWidth = [double]::Parse($rootEl.GetAttribute('stroke-width'), $inv) }
    $viewBox = 24.0
    $vb = $rootEl.GetAttribute('viewBox')
    if ($vb) { $viewBox = [double]::Parse(($vb -split '[\s,]+')[2], $inv) }
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($node in $rootEl.ChildNodes) {
        if ($node.NodeType -eq 'Element') { $paths.Add((Get-ShapePathData $node)) }
    }
    if ($paths.Count -eq 0) { throw "$File contains no shapes" }
    return @{ Paths = $paths.ToArray(); StrokeWidth = $strokeWidth; ViewBox = $viewBox }
}

function Get-LucideFile([string]$RelativePath, [string]$Dest) {
    $url = "https://raw.githubusercontent.com/lucide-icons/lucide/$version/$RelativePath"
    Write-Host "  downloading $url"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $Dest -UseBasicParsing
    } catch {
        throw "Could not download $url ($($_.Exception.Message)). Check the icon name in scripts/icons.json and your internet connection."
    }
}

function Test-PngFile([string]$File, [int]$Expected) {
    $b = [System.IO.File]::ReadAllBytes($File)
    if ($b.Length -lt 26 -or $b[0] -ne 0x89 -or $b[1] -ne 0x50 -or $b[2] -ne 0x4E -or $b[3] -ne 0x47) { throw "$File is not a PNG" }
    $w = ([int]$b[16] -shl 24) -bor ([int]$b[17] -shl 16) -bor ([int]$b[18] -shl 8) -bor [int]$b[19]
    $h = ([int]$b[20] -shl 24) -bor ([int]$b[21] -shl 16) -bor ([int]$b[22] -shl 8) -bor [int]$b[23]
    if ($w -ne $Expected -or $h -ne $Expected) { throw "$File is ${w}x${h}, expected ${Expected}x${Expected}" }
    if ($b[25] -ne 6 -or $b[24] -ne 8) { throw "$File is not 32-bit RGBA (color type $($b[25]), bit depth $($b[24]))" }
}

# ---------------------------------------------------------------- 1. ids used by the ribbon

$used = New-Object System.Collections.Generic.HashSet[string]
foreach ($file in $ribbonXml) {
    $path = Join-Path $ribbon $file
    if (-not (Test-Path -LiteralPath $path)) { continue }
    foreach ($m in [regex]::Matches((Get-Content -LiteralPath $path -Raw), '\simage="([^"]+)"')) { [void]$used.Add($m.Groups[1].Value) }
}
$failed = $false
foreach ($id in $used) {
    if (-not $map.Contains($id)) {
        Write-Host "ERROR: the ribbon uses image '$id' but scripts/icons.json has no icon for it" -ForegroundColor Red
        $failed = $true
    }
}
foreach ($id in $map.Keys) {
    if ($used.Count -gt 0 -and -not $used.Contains($id)) { Write-Host "warning: icon '$id' is not used by any ribbon button" -ForegroundColor Yellow }
}
if ($failed) { exit 1 }

# ---------------------------------------------------------------- 2. make sure the SVG sources exist

New-Item -ItemType Directory -Force -Path $srcDir | Out-Null
foreach ($name in ($map.Values | Select-Object -Unique)) {
    $svgFile = Join-Path $srcDir "$name.svg"
    if (-not (Test-Path -LiteralPath $svgFile)) { Get-LucideFile "icons/$name.svg" $svgFile }
}
$licenseFile = Join-Path $srcDir 'LICENSE'
if (-not (Test-Path -LiteralPath $licenseFile)) { Get-LucideFile 'LICENSE' $licenseFile }

# ---------------------------------------------------------------- 3. draw the PNGs

New-Item -ItemType Directory -Force -Path $imgDir | Out-Null
if ($DebugSvgDir) { New-Item -ItemType Directory -Force -Path $DebugSvgDir | Out-Null }

$icons = @{}
foreach ($id in $map.Keys) {
    try { $icons[$id] = Read-IconSvg (Join-Path $srcDir "$($map[$id]).svg") }
    catch { throw "Icon '$id' ($($map[$id])): $($_.Exception.Message)" }
}

if ($DebugSvgDir) {
    foreach ($id in $map.Keys) {
        $ic = $icons[$id]
        $body = ($ic.Paths | ForEach-Object { "  <path d=`"$_`" />" }) -join "`n"
        $text = "<svg xmlns=`"http://www.w3.org/2000/svg`" width=`"24`" height=`"24`" viewBox=`"0 0 24 24`" fill=`"none`" stroke=`"currentColor`" stroke-width=`"$(Format-Num $ic.StrokeWidth)`" stroke-linecap=`"round`" stroke-linejoin=`"round`">`n$body`n</svg>`n"
        [System.IO.File]::WriteAllText((Join-Path $DebugSvgDir "$id.svg"), $text, $utf8)
    }
}

if (-not $NoPng) {
    try { Add-Type -AssemblyName PresentationCore, WindowsBase }
    catch {
        throw "Windows drawing (WPF) is not available in this PowerShell. Run: powershell -ExecutionPolicy Bypass -File scripts\make-icons.ps1"
    }

    $wpfColor = [System.Windows.Media.ColorConverter]::ConvertFromString($color)
    $brush = New-Object System.Windows.Media.SolidColorBrush -ArgumentList $wpfColor
    $brush.Freeze()

    foreach ($id in $map.Keys) {
        $ic = $icons[$id]
        $scale = $size / $ic.ViewBox
        $pen = New-Object System.Windows.Media.Pen -ArgumentList $brush, $ic.StrokeWidth
        $pen.StartLineCap = [System.Windows.Media.PenLineCap]::Round
        $pen.EndLineCap   = [System.Windows.Media.PenLineCap]::Round
        $pen.LineJoin     = [System.Windows.Media.PenLineJoin]::Round
        $pen.Freeze()

        $visual = New-Object System.Windows.Media.DrawingVisual
        $dc = $visual.RenderOpen()
        $dc.PushTransform((New-Object System.Windows.Media.ScaleTransform -ArgumentList $scale, $scale))
        foreach ($d in $ic.Paths) {
            $geometry = [System.Windows.Media.Geometry]::Parse($d)
            $dc.DrawGeometry($null, $pen, $geometry)
        }
        $dc.Pop()
        $dc.Close()

        $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap -ArgumentList $size, $size, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($visual)
        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))

        $out = Join-Path $imgDir "$id.png"
        $stream = [System.IO.File]::Create($out)
        try { $encoder.Save($stream) } finally { $stream.Dispose() }
        Test-PngFile $out $size
    }

    # Remove PNGs that are no longer in the map (old or renamed icons).
    foreach ($f in Get-ChildItem -LiteralPath $imgDir -Filter '*.png') {
        if (-not $map.Contains([System.IO.Path]::GetFileNameWithoutExtension($f.Name))) {
            Remove-Item -LiteralPath $f.FullName
            Write-Host "removed stale $($f.Name)"
        }
    }
}

# ---------------------------------------------------------------- 4. relationship files

New-Item -ItemType Directory -Force -Path $relsDir | Out-Null
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
$lines.Add('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
foreach ($id in $map.Keys) {
    $lines.Add("  <Relationship Id=`"$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"images/$id.png`"/>")
}
$lines.Add('</Relationships>')
$relsText = ($lines -join "`n") + "`n"
foreach ($file in $ribbonXml) { [System.IO.File]::WriteAllText((Join-Path $relsDir "$file.rels"), $relsText, $utf8) }

# ---------------------------------------------------------------- 5. third-party notice

$noticeTemplate = @'
# Third-party notices

## Lucide icons

The ribbon icons in `ribbon/images/` are generated from the [Lucide](https://lucide.dev) icon set (version {VERSION}), recoloured and rendered to PNG by `scripts/make-icons.ps1`. The source SVG files are in `scripts/icons-src/`. The mapping is in `scripts/icons.json`.

```text
{LICENSE}
```
'@
$licenseText = (Get-Content -LiteralPath $licenseFile -Raw).Trim() -replace "`r`n", "`n"
$notice = $noticeTemplate.Replace('{VERSION}', $version).Replace('{LICENSE}', $licenseText) -replace "`r`n", "`n"
[System.IO.File]::WriteAllText((Join-Path $root 'THIRD-PARTY-NOTICES.md'), $notice + "`n", $utf8)

if ($NoPng) { Write-Host "Parsed $($map.Count) icons (PNG drawing skipped)." }
else { Write-Host "Generated $($map.Count) icons ($size px, $color)." -ForegroundColor Green }
exit 0
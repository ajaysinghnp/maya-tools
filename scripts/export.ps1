<#
Exports the VBA modules, class modules and forms of an .xlam into src\ and writes
the non built-in library references to references.txt.

  ./scripts/export.ps1                                default: dist\MayaTools-dev.xlam
  ./scripts/export.ps1 -Xlam "dist\MayaTools.xlam"
  ./scripts/export.ps1 -Prune                         also delete src files whose component no longer exists

Document modules (Sheet1-3, ThisWorkbook) are not exported; a warning is shown if one has code.
The generated modVersion module is never exported (build.ps1 creates it).
Needs "Trust access to the VBA project object model" and an unlocked VBA project.
VBA writes exports in the Windows ANSI code page, so keep the source ASCII (check.ps1 enforces it).
#>
param(
    [string]$Xlam = "dist\MayaTools-dev.xlam",
    [switch]$Prune
)

$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$src  = Join-Path $root "src"

if ([IO.Path]::IsPathRooted($Xlam)) { $xlamPath = $Xlam } else { $xlamPath = Join-Path $root $Xlam }
if (-not (Test-Path $xlamPath)) { throw "File not found: $xlamPath" }

# component type -> extension and default folder
$types = @{
    1 = @{ Ext = "bas"; Dir = "modules" }   # standard module
    2 = @{ Ext = "cls"; Dir = "classes" }   # class module
    3 = @{ Ext = "frm"; Dir = "forms"   }   # user form (also writes the .frx)
}

$exported = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$xl = $null
$wb = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 1              # the VBA project only loads with macros allowed

    Write-Host "Opening: $xlamPath"
    $wb = $xl.Workbooks.Open($xlamPath, 0, $true)    # read-only, so the file is never modified

    try   { $proj = $wb.VBProject; $count = $proj.VBComponents.Count }
    catch { throw "Cannot access the VBA project: $($_.Exception.Message)" }

    if ($proj.Protection -eq 1) {
        throw "The VBA project is locked with a password. Unlock it first (VBE > Tools > VBAProject Properties > Protection)."
    }
    if ($count -eq 0) {
        throw "The VBA project loaded with 0 components. Check the Trust Center setting and that macros are not blocked for this file."
    }

    Write-Host "Project: $($proj.Name), components: $count"

    foreach ($c in $proj.VBComponents) {
        $name = $c.Name

        if ($name -eq "modVersion") { continue }

        # Document modules (sheets, ThisWorkbook): not exported, warn if they hold code
        if ($c.Type -eq 100) {
            if ($c.CodeModule.CountOfLines -gt 0) {
                Write-Warning "$name contains $($c.CodeModule.CountOfLines) lines of code and was NOT exported."
            }
            continue
        }

        if (-not $types.ContainsKey([int]$c.Type)) {
            Write-Warning "Skipped $name (unsupported component type $($c.Type))"
            continue
        }

        $ext = $types[[int]$c.Type].Ext
        $dir = $types[[int]$c.Type].Dir

        # Keep the file where it already lives under src\ (preserves manual subfolders)
        $existing = $null
        if (Test-Path $src) {
            $existing = Get-ChildItem $src -Recurse -Filter "$name.$ext" -File | Select-Object -First 1
        }

        if ($existing) {
            $target = $existing.FullName
        }
        else {
            $folder = Join-Path $src $dir
            New-Item -ItemType Directory -Force $folder | Out-Null
            $target = Join-Path $folder "$name.$ext"
        }

        $c.Export($target)
        [void]$exported.Add($target)
        Write-Host ("  Exported {0,-28} -> {1}" -f $name, $target.Substring($root.Length + 1))
    }

    # Library references (non built-in only)
    $refs = @()
    foreach ($r in $proj.References) {
        if (-not $r.BuiltIn) {
            $refs += ("{0}|{1}|{2}|{3}" -f $r.Name, $r.GUID, $r.Major, $r.Minor)
        }
    }
    [IO.File]::WriteAllLines((Join-Path $root "references.txt"), [string[]]$refs)
    Write-Host "References written: $($refs.Count)"
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($xl) {
        $xl.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
}

# Source files whose component no longer exists in the project
if ($exported.Count -gt 0) {
    $orphans = Get-ChildItem $src -Recurse -Include *.bas, *.cls, *.frm -File |
        Where-Object { -not $exported.Contains($_.FullName) }

    foreach ($o in $orphans) {
        $rel = $o.FullName.Substring($root.Length + 1)
        if ($Prune) {
            Remove-Item $o.FullName
            if ($o.Extension -eq ".frm") {
                $frx = [IO.Path]::ChangeExtension($o.FullName, ".frx")
                if (Test-Path $frx) { Remove-Item $frx }
            }
            Write-Host "  Removed $rel"
        }
        else {
            Write-Warning "$rel is not in the project any more (use -Prune to delete it)."
        }
    }
}

Write-Host "Done. Exported $($exported.Count) component(s)."
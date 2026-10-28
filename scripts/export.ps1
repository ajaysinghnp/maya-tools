<#
Exports VBA modules and forms from an .xlam into src/, plus references.txt.
Document modules (Sheet1-3, ThisWorkbook) are NOT exported; a warning is shown
if any of them contains code.
#>
param([string]$Xlam = "_oldversions\MayaTools v2.xlam")

$ErrorActionPreference = "Stop"

$root     = Split-Path $PSScriptRoot -Parent
$src      = Join-Path $root "src"
$xlamPath = Join-Path $root $Xlam

if (-not (Test-Path $xlamPath)) { throw "File not found: $xlamPath" }

# component type -> extension and default folder
$types = @{
    1 = @{ Ext = "bas"; Dir = "modules" }   # standard module
    2 = @{ Ext = "cls"; Dir = "classes" }   # class module
    3 = @{ Ext = "frm"; Dir = "forms"   }   # user form (also writes .frx)
}

$xl = $null
$wb = $null
$exported = 0

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1              # msoAutomationSecurityLow: load the VBA project

    Write-Host "Opening: $xlamPath"
    $wb = $xl.Workbooks.Open($xlamPath)

    try   { 
      $proj = $wb.VBProject; 
      $null = $proj.VBComponents.Count
      if ($proj.VBComponents.Count -eq 0) {
        throw "The VBA project loaded with 0 components. Macros are probably blocked for this file (see Trusted Locations / policy)."
      }
    }
    catch { throw "Cannot access the VBA project. Enable 'Trust access to the VBA project object model' (Trust Center > Macro Settings), or the setting is blocked by policy." }

    if ($proj.Protection -eq 1) {
        throw "The VBA project is locked with a password. Unlock it first (VBE > Tools > VBAProject Properties > Protection)."
    }

    Write-Host "Project: $($proj.Name), components: $($proj.VBComponents.Count)"

    foreach ($c in $proj.VBComponents) {
        $name = $c.Name

        # Document modules (sheets, ThisWorkbook): skip, warn if they hold code
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
            $existing = Get-ChildItem $src -Recurse -Filter "$name.$ext" | Select-Object -First 1
        }

        if ($existing) {
            $target = $existing.FullName
        }
        else {
            New-Item -ItemType Directory -Force (Join-Path $src $dir) | Out-Null
            $target = Join-Path (Join-Path $src $dir) "$name.$ext"
        }

        $c.Export($target)
        $exported++
        Write-Host ("  Exported {0,-28} -> {1}" -f $name, $target.Substring($root.Length + 1))
    }

    # Library references (non built-in only)
    $refs = @()
    foreach ($r in $proj.References) {
        if (-not $r.BuiltIn) {
            $refs += ("{0}|{1}|{2}|{3}" -f $r.Name, $r.GUID, $r.Major, $r.Minor)
        }
    }
    Set-Content -Path (Join-Path $root "references.txt") -Value $refs
    Write-Host "References written: $($refs.Count)"
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($xl) {
        $xl.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
}

Write-Host "Done. Exported $exported component(s)."
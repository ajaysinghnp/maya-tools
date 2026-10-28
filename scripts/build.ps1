param(
  [string]$Name = "MayaTools",
  [string]$Version = "",
  [string]$Out = "dist\$Name.xlam",
  [string]$Repo = "ajaysinghnp/MayaTools"   # change to the real owner/repo
)

$root = Split-Path $PSScriptRoot -Parent
$out = Join-Path $root $Out
$ansi = [Text.Encoding]::GetEncoding(1252)

# Version: parameter (from CI tag) wins, otherwise the VERSION file
if (-not $Version) {
  $Version = (Get-Content (Join-Path $root "VERSION") -Raw).Trim()
}
$Version = $Version.TrimStart("v")

New-Item -ItemType Directory -Force (Split-Path $out) | Out-Null
if (Test-Path $out) { Remove-Item $out }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $xl.SheetsInNewWorkbook = 1
  $wb = $xl.Workbooks.Add()
  $proj = $wb.VBProject

  # References first
  $refFile = Join-Path $root "references.txt"
  if (Test-Path $refFile) {
    Get-Content $refFile | Where-Object { $_ } | ForEach-Object {
        $n,$guid,$maj,$min = $_ -split '\|'
        try   { [void]$proj.References.AddFromGuid($guid, [int]$maj, [int]$min) }
        catch { Write-Warning "Reference $n not added: $($_.Exception.Message)" }
    }
  }

  # Source modules, classes, forms
  Get-ChildItem "$root\src" -Recurse -Include *.bas, *.cls, *.frm |
  Where-Object { $_.Name -ne "ThisWorkbook.bas" } |
  ForEach-Object { [void]$proj.VBComponents.Import($_.FullName) }

  # ThisWorkbook code
  $tw = Join-Path $root "src\document\ThisWorkbook.bas"
  if (Test-Path $tw) {
    $proj.VBComponents.Item("ThisWorkbook").CodeModule.AddFromString(
      [IO.File]::ReadAllText($tw, $ansi))
  }

  # Generated version module
  $ver = $proj.VBComponents.Add(1)      # 1 = standard module
  $ver.Name = "modVersion"
  $ver.CodeModule.AddFromString(@"
Option Explicit
Public Const ADDIN_NAME As String = "$Name"
Public Const ADDIN_VERSION As String = "$Version"
Public Const ADDIN_REPO As String = "$Repo"
"@)

  $proj.Name = $Name
  $wb.Title = $Name                  # name shown in the Add-ins dialog
  $wb.Comments = "$Name v$Version"
  $wb.IsAddin = $true
  $wb.SaveAs($out, 55)                  # xlOpenXMLAddIn
  $wb.Close($false)
}
finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host "Built $out (v$Version)"
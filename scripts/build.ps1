param(
    [string]$Name    = "MayaTools",
    [string]$Version = "",
    [string]$Out     = "",
    [string]$Repo    = "ajaysinghnp/maya-tools"   # your real owner/repo
)

$root = Split-Path $PSScriptRoot -Parent
$ansi = [Text.Encoding]::GetEncoding(1252)

# Version: parameter (from CI tag) wins, otherwise the VERSION file
if (-not $Version) {
    $Version = (Get-Content (Join-Path $root "VERSION") -Raw).Trim()
}
$Version = $Version.TrimStart("v")

# Output name: dist\MayaTools-v2.0.0.xlam unless -Out was given
if (-not $Out) { $Out = "dist\$Name-v$Version.xlam" }
$out = Join-Path $root $Out

Get-ChildItem (Split-Path $out) -Filter "$Name-v*.xlam" -ErrorAction SilentlyContinue | Remove-Item -Force

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

# ---- Inject ribbon (customUI) into the saved package ----
$ribbonDir = Join-Path $root "ribbon"
if (Test-Path $ribbonDir) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function Read-Entry($zip, $name) {
        $r = New-Object IO.StreamReader($zip.GetEntry($name).Open())
        try { $r.ReadToEnd() } finally { $r.Dispose() }
    }
    function Write-Entry($zip, $name, $text) {
        $old = $zip.GetEntry($name); if ($old) { $old.Delete() }
        $e = $zip.CreateEntry($name)
        $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false)))
        try { $w.Write($text) } finally { $w.Dispose() }
    }

    $zip = [IO.Compression.ZipFile]::Open($out, [IO.Compression.ZipArchiveMode]::Update)
    try {
        # 1. copy ribbon files into customUI/
        foreach ($f in Get-ChildItem $ribbonDir -Recurse -File) {
            $rel  = $f.FullName.Substring($ribbonDir.Length + 1).Replace('\', '/')
            $name = "customUI/$rel"
            $old  = $zip.GetEntry($name); if ($old) { $old.Delete() }
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $name)
        }

        # 2. root relationships: point the package at the ribbon XML
        $xml = New-Object Xml.XmlDocument
        $xml.LoadXml((Read-Entry $zip "_rels/.rels"))
        $ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        $wanted = @(
            @{ Id = "rIdUI07"; File = "customUI.xml";   Type = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" },
            @{ Id = "rIdUI10"; File = "customUI14.xml"; Type = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" }
        )
        foreach ($w in $wanted) {
            if (-not (Test-Path (Join-Path $ribbonDir $w.File))) { continue }
            $target = "customUI/$($w.File)"
            $found  = $xml.DocumentElement.ChildNodes | Where-Object { $_.Target -eq $target }
            if (-not $found) {
                $el = $xml.CreateElement("Relationship", $ns)
                $el.SetAttribute("Id", $w.Id)
                $el.SetAttribute("Type", $w.Type)
                $el.SetAttribute("Target", $target)
                [void]$xml.DocumentElement.AppendChild($el)
            }
        }
        $text = $xml.OuterXml
        if (-not $text.StartsWith("<?xml")) { $text = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $text }
        Write-Entry $zip "_rels/.rels" $text

        # 3. content types: make sure PNG is declared
        $ct = Read-Entry $zip "[Content_Types].xml"
        if ($ct -notmatch 'Extension="png"') {
            $ct = $ct -replace '(<Types[^>]*>)', '$1<Default Extension="png" ContentType="image/png"/>'
            Write-Entry $zip "[Content_Types].xml" $ct
        }
    }
    finally { $zip.Dispose() }
    Write-Host "Ribbon injected from $ribbonDir"
}

Write-Host "Built $out (v$Version)"
# MayaTools

An Excel add-in (`.xlam`) of VBA productivity tools with its own ribbon tab. It includes Nepali (Bikram Sambat) date functions, number-to-words conversion, and helpers for worksheets, workbooks, formatting, formulas, pivot tables and printing.

The source is kept as plain-text VBA modules in this repository and built into a release `.xlam` with the scripts in `scripts/`. Developers: see [dev.md](dev.md).

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Updating and uninstalling](#updating-and-uninstalling)
- [Project structure](#project-structure)
- [Building and contributing](#building-and-contributing)
- [Function reference](#function-reference)
- [Author](#author)

## Features

MayaTools adds a ribbon tab with buttons for common spreadsheet chores, and worksheet functions you can use in formulas. The main areas are:

- **Nepali dates**: functions for working with Bikram Sambat dates.
- **Spell numbers**: amounts in words, including Nepali.
- **Worksheets and workbooks**: insert many sheets at once, sort and reverse sheets, split sheets, build a table of contents, copy sheets to a new workbook, show hidden sheets, save sheets as PDF, close all workbooks, back up the file.
- **Formatting and colours**: highlight duplicates, greater or lesser values and negatives; fill formats; serial numbers; insert rows and columns.
- **Formulas and text**: extract numbers, convert text to numbers, append symbols to numeric cells.
- **Pivot tables, charts, printing, headers and footers**: shortcut routines for repetitive setup.
- **Tax calculation formatting** and an **About developer** form.

This list is a summary. Each procedure is documented in the [function reference](#function-reference) as the documentation grows.

## Requirements

- Windows with desktop Excel 2007 or later (the add-in is built in the `.xlam` format that Excel 2007 and later read).
- Macros enabled for the add-in (see the installation steps).

## Installation

1. Download the latest `MayaTools-v<version>.xlam` from the [Releases](../../releases) page.
2. **Unblock the file.** Files downloaded from the internet are blocked by Windows. Right-click the file, choose *Properties*, tick **Unblock**, and press OK. Or in PowerShell:

   ```powershell
   Unblock-File "$HOME\Downloads\MayaTools-v<version>.xlam"
   ```

3. Move the file to a permanent folder, for example `%APPDATA%\Microsoft\AddIns`.
4. In Excel: *File, Options, Add-ins*. At the bottom, set **Manage** to *Excel Add-ins* and press **Go**.
5. Press **Browse**, select the `.xlam` file, and make sure **MayaTools** is ticked. Press OK.
6. The **MayaTools** tab appears in the ribbon.

If Excel shows a security warning, either unblock the file (step 2) or place it in a Trusted Location (*File, Options, Trust Center, Trust Center Settings, Trusted Locations*).

## Updating and uninstalling

- **Update**: untick the old version in the Add-ins dialog, close Excel, delete or replace the old file, then install the new one as above. Do not keep two versions loaded at the same time.
- **Uninstall**: untick MayaTools in the Add-ins dialog and delete the file.

## Project structure

```
maya-tools/
?? src/                 VBA source: modules, classes and forms
?? ribbon/              custom ribbon definition and icons
?? scripts/             development, build and release scripts
?? releases/            built .xlam of each release
?? .github/workflows/   source checks and release publishing
?? VERSION              current version
?? README.md            this file
?? dev.md               developer guide
```

| Source module | Contents |
| --- | --- |
| `AdvanceSubroutines` | Advanced routines |
| `ChartSubroutines` | Chart helpers |
| `colorFunctions` | Colour functions |
| `FormattingSubroutines` | Cell and range formatting |
| `formulasFunctions` | Worksheet functions |
| `formulasSubroutines` | Formula and cell-content routines |
| `HeaderFooterSubroutines` | Page header and footer helpers |
| `HelperSubroutines` | Shared helper routines |
| `NepaliDatesFunctions` | Nepali (Bikram Sambat) date functions |
| `pivotTablesSubroutines` | Pivot table helpers |
| `printSubroutines` | Printing and page setup |
| `sheetsSubroutines` | Worksheet operations |
| `SpellNumbers` | Number-to-words functions |
| `TaxCalculationFormat` | Tax calculation formatting |
| `TextToNumber` | Text-to-number conversion |
| `workBookSubroutines` | Workbook-level operations |
| `aboutDeveloper` (form) | About developer dialog |

The one-line descriptions are provisional and will be refined as the functions below are documented.

## Building and contributing

Building needs desktop Excel, PowerShell 7 and Git. The whole workflow (edit, check, build, commit, release) is driven by `scripts/dev.ps1` and `scripts/release.ps1`; the full guide is in [dev.md](dev.md).

Quick start:

```powershell
./scripts/dev.ps1                                   # check, build the dev copy, open it in Excel
./scripts/dev.ps1 -Export -Message "module: change" # after editing in the VBA editor
```

## Function reference

Every public procedure gets an entry in the module sections below. Procedures that are called from the ribbon are marked with the button they belong to.

### Entry template

Copy this block for each procedure:

```markdown
#### ProcedureName

| | |
|---|---|
| Type | `Function` / `Sub` |
| Signature | `Function ProcedureName(arg1 As String, Optional arg2 As Long = 0) As String` |
| Ribbon button | Tab / group / button label, or "none" (callable from formulas or code) |
| Description | What it does, in one or two sentences |

**Parameters**

| Name | Type | Description |
|---|---|---|
| `arg1` | `String` | ... |

**Returns**: what a `Function` returns (omit for a `Sub`).

**Example**

    =ProcedureName("text")

**Notes**: edge cases, limitations, required references.
```

To list every procedure in a module as a starting point (run from the repo root):

```powershell
Get-ChildItem src -Recurse -Include *.bas,*.cls | ForEach-Object {
    $m = $_.BaseName
    Select-String -Path $_.FullName -Pattern '^\s*(Public\s+|Private\s+)?(Sub|Function)\s+\w+' |
        ForEach-Object { "$m | $($_.Line.Trim())" }
}
```

### AdvanceSubroutines

*To be documented.*

### ChartSubroutines

*To be documented.*

### colorFunctions

*To be documented.*

### FormattingSubroutines

*To be documented.*

### formulasFunctions

*To be documented.*

### formulasSubroutines

*To be documented.*

### HeaderFooterSubroutines

*To be documented.*

### HelperSubroutines

*To be documented.*

### NepaliDatesFunctions

*To be documented.*

### pivotTablesSubroutines

*To be documented.*

### printSubroutines

*To be documented.*

### sheetsSubroutines

*To be documented.*

### SpellNumbers

*To be documented.*

### TaxCalculationFormat

*To be documented.*

### TextToNumber

*To be documented.*

### workBookSubroutines

*To be documented.*

### aboutDeveloper (form)

*To be documented.*

## Author

Developed by Ajay ([@ajaysinghnp](https://github.com/ajaysinghnp)).

License: not specified yet.

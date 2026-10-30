# MayaTools User Guide

This guide is for people who use the released MayaTools add-in in Excel. For a short overview of the project, see the [README](README.md).

## Contents

- [MayaTools User Guide](#mayatools-user-guide)
  - [Contents](#contents)
  - [What MayaTools adds to Excel](#what-mayatools-adds-to-excel)
  - [Requirements](#requirements)
  - [Installation](#installation)
  - [Updating and uninstalling](#updating-and-uninstalling)
  - [Using MayaTools](#using-mayatools)
  - [Troubleshooting](#troubleshooting)
  - [Function reference](#function-reference)
    - [Nepali dates](#nepali-dates)
    - [Spell numbers](#spell-numbers)
    - [Worksheets](#worksheets)
    - [Workbooks](#workbooks)
    - [Formatting](#formatting)
    - [Colours](#colours)
    - [Formula functions](#formula-functions)
    - [Formula and cell-content tools](#formula-and-cell-content-tools)
    - [Text to number](#text-to-number)
    - [Pivot tables](#pivot-tables)
    - [Charts](#charts)
    - [Printing and page setup](#printing-and-page-setup)
    - [Headers and footers](#headers-and-footers)
    - [Tax calculation format](#tax-calculation-format)
    - [Advanced routines](#advanced-routines)
    - [Helper routines](#helper-routines)
    - [About developer](#about-developer)
  - [Support](#support)
  - [License and credit](#license-and-credit)

## What MayaTools adds to Excel

MayaTools adds a **MayaTools** tab to the ribbon with buttons for common spreadsheet chores, and worksheet functions you can type into formulas. The main areas are:

- **Nepali dates**: functions for working with Bikram Sambat dates.
- **Spell numbers**: amounts in words, including Nepali.
- **Worksheets and workbooks**: insert many sheets at once, sort and reverse sheets, split sheets, build a table of contents, copy sheets to a new workbook, show hidden sheets, save sheets as PDF, close all workbooks, back up the file.
- **Formatting and colours**: highlight duplicates, greater or lesser values and negatives; fill formats; serial numbers; insert rows and columns.
- **Formulas and text**: extract numbers, convert text to numbers, append symbols to numeric cells.
- **Pivot tables, charts, printing, headers and footers**: shortcut routines for repetitive setup.
- **Tax calculation formatting** and an **About developer** form.

## Requirements

- Windows with desktop Excel 2007 or later.
- Macros enabled for the add-in (see the installation steps).

Excel for the web and Excel for Mac do not load this add-in.

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

## Using MayaTools

- **Ribbon buttons**: open the **MayaTools** tab and click a button. Most buttons act on the current selection or the active sheet, so select the cells or sheet first.
- **Worksheet functions**: type them in a cell like any Excel function, for example `=ProcedureName("text")`. Functions are listed in the [function reference](#function-reference).
- **Undo**: macros clear Excel's undo history. Save your work, or use the backup button, before running a bulk change.
- **About**: the *About developer* button shows the installed version and author details.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| The MayaTools tab does not appear | Check that MayaTools is ticked in *File, Options, Add-ins* (Manage: *Excel Add-ins*). Make sure only one MayaTools file is loaded, then restart Excel |
| Security warning, or macros are disabled | Unblock the file (installation step 2) or place it in a Trusted Location |
| A function shows `#NAME?` | The add-in is not loaded. Tick it in the Add-ins dialog and recalculate with Ctrl+Alt+F9 |
| A button does nothing | Close Excel completely, reopen it, and try again. If it persists, report it with the version number from the About form |
| The file cannot be moved or deleted | Excel has it open. Untick it in the Add-ins dialog and close Excel first |

## Function reference

The functions and ribbon buttons are listed below by task, with the ribbon button each one belongs to. Sections marked *To be documented* are coming soon.

### Nepali dates

*To be documented.*

### Spell numbers

*To be documented.*

### Worksheets

*To be documented.*

### Workbooks

*To be documented.*

### Formatting

*To be documented.*

### Colours

*To be documented.*

### Formula functions

*To be documented.*

### Formula and cell-content tools

*To be documented.*

### Text to number

*To be documented.*

### Pivot tables

*To be documented.*

### Charts

*To be documented.*

### Printing and page setup

*To be documented.*

### Headers and footers

*To be documented.*

### Tax calculation format

*To be documented.*

### Advanced routines

*To be documented.*

### Helper routines

*To be documented.*

### About developer

*To be documented.*

## Support

Report problems or request features on the repository's issue tracker, and include your Excel version and the MayaTools version.

## License and credit

MayaTools is free to use, at home or at work, under the [MIT License](LICENSE).

- **You may**: use it, copy it, share it, modify it, and include it in your own work, including commercial work.
- **You must**: give credit to the author, Ajay ([@ajaysinghnp](https://github.com/ajaysinghnp)), by keeping the copyright notice and license text with any copy or substantial portion you distribute.
- **Icons**: the ribbon icons come from the [Lucide](https://lucide.dev) icon set (ISC License). Its notice is in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
- **No warranty**: the add-in is provided as is, and the author is not liable for problems arising from its use. Keep a backup of important workbooks before running bulk changes.

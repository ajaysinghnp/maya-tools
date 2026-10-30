# MayaTools

An Excel add-in (`.xlam`) of VBA productivity tools with its own ribbon tab. It includes Nepali (Bikram Sambat) date functions, number-to-words conversion, and helpers for worksheets, workbooks, formatting, formulas, pivot tables and printing.

## Where to go

| I want to... | Read |
| --- | --- |
| Install and use MayaTools | [DOCS.md](DOCS.md) |
| Look up a function or ribbon button | [Function reference](DOCS.md#function-reference) |
| Fix a problem with the add-in | [Troubleshooting](DOCS.md#troubleshooting) |
| Build it, change the code, or publish a release | [dev.md](dev.md) |
| Check the terms of use | [LICENSE](LICENSE) |

## Highlights

- **Nepali dates**: functions for working with Bikram Sambat dates.
- **Spell numbers**: amounts in words, including Nepali.
- **Worksheets and workbooks**: insert many sheets at once, sort and split sheets, build a table of contents, save sheets as PDF, back up the file.
- **Formatting and colours**: highlight duplicates, greater or lesser values and negatives; serial numbers; insert rows and columns.
- **Formulas and text**: extract numbers, convert text to numbers, append symbols to numeric cells.
- **Pivot tables, charts, printing, headers and footers**: shortcut routines for repetitive setup.

## Quick install

1. Download the latest `MayaTools-v<version>.xlam` from the [Releases](../../releases) page.
2. Unblock it (right-click, *Properties*, tick **Unblock**) and move it to a permanent folder.
3. In Excel: *File, Options, Add-ins*, set **Manage** to *Excel Add-ins*, press **Go**, **Browse** to the file and tick **MayaTools**.

Full steps, updating and uninstalling are in [DOCS.md](DOCS.md#installation).

## Requirements

- Windows with desktop Excel 2007 or later
- Macros enabled for the add-in

## For developers

The source is kept as plain-text VBA modules in this repository and built into a release `.xlam` with the scripts in `scripts/`. See [dev.md](dev.md) for the repository layout, build pipeline, source rules and release process.

## License and credit

MayaTools is released under the [MIT License](LICENSE). You may use, copy, modify, share and even sell it, free of charge, as long as you keep the copyright notice and the license text with every copy or substantial portion. In practice: credit the author.

Ribbon icons are from [Lucide](https://lucide.dev) (ISC License); see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

Developed by Ajay ([@ajaysinghnp](https://github.com/ajaysinghnp)).

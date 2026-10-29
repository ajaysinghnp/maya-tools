# Developer guide

How to work on MayaTools: edit the source, build and test the add-in, document it, commit, and publish a release.

This guide is for contributors. End users should read the [user guide](DOCS.md); the [README](README.md) is the short overview.

## Contents

- [Developer guide](#developer-guide)
  - [Contents](#contents)
  - [1. How the repo works](#1-how-the-repo-works)
  - [2. Prerequisites](#2-prerequisites)
  - [3. Repository layout](#3-repository-layout)
    - [Source modules](#source-modules)
  - [4. The scripts](#4-the-scripts)
    - [Pipeline order](#pipeline-order)
  - [5. Daily workflow](#5-daily-workflow)
    - [Edit in the VBA editor (forms, quick experiments)](#edit-in-the-vba-editor-forms-quick-experiments)
    - [Edit in a text editor (modules)](#edit-in-a-text-editor-modules)
    - [`dev.ps1` options](#devps1-options)
    - [The one rule that prevents data loss](#the-one-rule-that-prevents-data-loss)
    - [While you work](#while-you-work)
  - [6. Source rules](#6-source-rules)
  - [7. Common changes](#7-common-changes)
    - [Add a module, class or form](#add-a-module-class-or-form)
    - [Rename or delete a component](#rename-or-delete-a-component)
    - [Change the ribbon](#change-the-ribbon)
  - [8. Documenting procedures](#8-documenting-procedures)
    - [Entry template](#entry-template)
    - [List procedures (to seed the documentation)](#list-procedures-to-seed-the-documentation)
  - [9. Versioning and releases](#9-versioning-and-releases)
    - [Publish a release](#publish-a-release)
    - [Before you release](#before-you-release)
    - [Fix a bad release](#fix-a-bad-release)
  - [10. Commit conventions](#10-commit-conventions)
  - [11. Troubleshooting](#11-troubleshooting)
  - [12. Reference: the original add-in](#12-reference-the-original-add-in)

## 1. How the repo works

An `.xlam` is a zip that contains a binary `vbaProject.bin`. Git cannot diff or merge it, so this repo keeps the **exported text files as the source of truth** and treats the `.xlam` as a **build artifact**.

- VBA modules, classes and forms live in `src/` as `.bas`, `.cls`, `.frm` (+ `.frx`).
- The ribbon (custom tab, buttons, icons) lives in `ribbon/`.
- `scripts/build.ps1` creates a brand-new workbook through Excel automation, imports everything from `src/`, adds a generated `modVersion` module, saves it as an `.xlam` and injects the ribbon files.
- Building needs desktop Excel, so it only runs on a developer machine. GitHub-hosted runners have no Excel. The release flow ([section 9](#9-versioning-and-releases)) builds locally and lets GitHub publish the file.

Nothing from an older `.xlam` is carried over into a build, so no leftover sheets, names or other debris.

## 2. Prerequisites

| Requirement | Notes |
| --- | --- |
| Windows with desktop Excel 2007 or later | Excel automation (COM) is needed for build and export |
| PowerShell 7+ (`pwsh`) | The scripts and the GitHub workflows use it |
| Git | |
| **Trust access to the VBA project object model** | See below |

**Enable VBA project access** (once per Windows user):

- Excel: *File, Options, Trust Center, Trust Center Settings, Macro Settings*, tick **Trust access to the VBA project object model**; or
- PowerShell (`16.0` covers Excel 2016, 2019, 2021 and 365):

```powershell
$key = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
```

If a group policy forces the setting off, build and export cannot run on that machine.

**Allow scripts for the current PowerShell session** if they are blocked:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## 3. Repository layout

```text
maya-tools/
??? src/
?   ??? modules/            standard modules (.bas), may be grouped in subfolders
?   ??? classes/            class modules (.cls), created on demand
?   ??? forms/              user forms: .frm + .frx (always together)
??? ribbon/
?   ??? customUI.xml        ribbon for Excel 2007
?   ??? customUI14.xml      ribbon for Excel 2010 and later
?   ??? images/             button icons (.png)
?   ??? _rels/              maps image ids to files (needed for the icons)
??? scripts/                pipeline scripts (section 4)
??? releases/               built .xlam of each release, committed by release.ps1
??? _oldversions/           the original v2 add-in (password-locked, reference only)
??? .github/workflows/      check.yml, release.yml
??? references.txt          extra library references (Name|GUID|Major|Minor)
??? VERSION                 current version, for example 2.0.0
??? .gitattributes
??? .gitignore
??? LICENSE                 MIT License text
??? README.md               short overview and links
??? DOCS.md                 user guide and function reference
??? dev.md                  this guide
```

`dist/` (local build output) is git-ignored.

### Source modules

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

The one-line descriptions are provisional and will be refined as the functions are documented in [DOCS.md](DOCS.md#function-reference).

## 4. The scripts

Run everything from the repo root.

| Script | Purpose |
| --- | --- |
| `scripts/dev.ps1` | **The development pipeline.** Optionally exports your VBE edits, runs the checks, builds the dev copy, optionally commits, opens Excel |
| `scripts/build-dev.ps1` | A step of `dev.ps1`: builds `dist\MayaTools-dev.xlam` and refuses to overwrite VBE edits that were never exported |
| `scripts/export.ps1` | A step of `dev.ps1 -Export`: writes modules, classes, forms and references from an `.xlam` into `src/` |
| `scripts/check.ps1` | Source checks (also run by GitHub Actions) |
| `scripts/build.ps1` | Builds `dist\MayaTools-v<version>.xlam` from `src/` and `ribbon/` |
| `scripts/release.ps1` | Builds, commits the file under `releases/`, tags and pushes. GitHub publishes the release |

`check.ps1` verifies:

- the `VERSION` file is a valid version (`2.1.0` or `2.1.0-beta`);
- every `.bas`, `.cls` and `.frm` file is pure ASCII (see [section 6](#6-source-rules));
- every `.frm` has its `.frx`;
- the ribbon XML files are valid and every callback (`onAction`, `getImage`, and so on) has a matching `Sub` or `Function` in `src/`.

### Pipeline order

```text
dev.ps1
 ?? (-Export)            export.ps1   dev copy -> src/, references.txt
 ?? check.ps1                         source checks
 ?? build-dev.ps1        build.ps1    src/ + ribbon/ -> dist\MayaTools-dev.xlam
 ?? (-Message / -Commit) git add, diff, commit (src/, ribbon/, references.txt only)
 ?? open dist\MayaTools-dev.xlam in Excel
```

Each step runs only if the previous one succeeded, so nothing is committed from source that fails the checks or the build.

## 5. Daily workflow

### Edit in the VBA editor (forms, quick experiments)

The VBE is still the only form designer. Use it as a tool, and keep the repo as the source of truth.

```powershell
./scripts/dev.ps1
```

1. This builds `dist\MayaTools-dev.xlam` and opens it in Excel.
2. Press **Alt+F11**. Edit the form or module, test it, and save with **Ctrl+S**.
3. **Close Excel completely.** Excel locks the file while it is open.
4. Pull the edits into the repo, build again and commit:

```powershell
./scripts/dev.ps1 -Export -Message "aboutDeveloper: add version label"
```

The pipeline exports, checks, rebuilds, shows the diff, asks `Commit these changes? [Y/n]`, commits, and reopens Excel. Use `-Export -Commit` to be asked for the message, or leave both out to export and rebuild without committing.

### Edit in a text editor (modules)

Standard modules are plain text, so editing in VS Code works fine:

```powershell
# edit src\modules\SpellNumbers.bas ...
./scripts/dev.ps1 -Message "SpellNumbers: fix rounding" -Yes
```

Form code-behind can be edited as text too (below the `Attribute VB_...` lines in the `.frm`). Do not hand-edit the `Begin` / `End` control block or the header. Layout changes belong in the designer, because positions and `.frx` offsets are easy to corrupt.

### `dev.ps1` options

| Option | Effect |
| --- | --- |
| `-Export` | First export the VBE edits from the dev copy into `src/` |
| `-Message "text"` | Commit after a green build with this message |
| `-Commit` | Commit after a green build and ask for the message |
| `-Yes` | Skip the commit confirmation |
| `-NoOpen` | Do not open Excel at the end |
| `-CloseExcel` | Close running Excel instances instead of stopping with an error |
| `-Force` | Rebuild even if the dev copy holds edits that were never exported (they are lost) |

`dev.ps1` commits only `src/`, `ribbon/` and `references.txt`. Commit scripts, `VERSION`, workflows and documentation separately.

### The one rule that prevents data loss

**After changing anything in the VBE, export before building again.** Builds regenerate the dev copy from `src/`, so unexported VBE edits would be overwritten. `build-dev.ps1` refuses in that situation and tells you to use `-Export` (keep your edits) or `-Force` (discard them).

### While you work

- Untick the installed `MayaTools.xlam` in *File, Options, Add-ins* so two copies with the same ribbon and project name are not loaded together.
- Keep Excel closed while a script runs. The scripts start their own hidden Excel instance.
- `git push` is a separate, manual step.

## 6. Source rules

1. **Pure ASCII.** The VBE stores source in the Windows ANSI code page, so non-ASCII characters change meaning between machines and editors. Write special characters with `ChrW()`:

   ```vba
   ActiveCell.Value = ActiveCell.Value & ChrW(176)    ' degree sign
   ```

   Nepali text in this project uses a legacy font-encoded scheme (Latin characters drawn as Devanagari by a special font), which stays valid ASCII except for a few characters, written as `ChrW(164)` and `ChrW(182)`. `check.ps1` fails on any non-ASCII byte.
2. **CRLF line endings.** `.gitattributes` enforces them for `.bas`, `.cls` and `.frm`; `.frx` and `.png` are binary.
3. **Component name equals file name.** `HelperSubroutines` is `HelperSubroutines.bas`. The name is what the VBE uses on import.
4. **`.frm` and `.frx` always travel together**, in the same folder.
5. **No code in sheets or `ThisWorkbook`.** Document modules are not exported or built. `export.ps1` warns if one contains code.
6. **`modVersion` is generated** during the build (constants `ADDIN_NAME`, `ADDIN_VERSION`, `ADDIN_REPO`). Do not create it in `src/`; the export skips it.
7. **Library references** besides the built-in ones go in `references.txt`, one per line as `Name|GUID|Major|Minor`. `export.ps1` rewrites the file from the dev copy.
8. **No password on the VBA project.** The build produces an unlocked project, and export cannot read a locked one.

## 7. Common changes

### Add a module, class or form

Create it in the VBE (*Insert, Module / Class Module / UserForm*), name it in the Properties window (F4), save, close Excel, then run `./scripts/dev.ps1 -Export`. New files land in `src/modules`, `src/classes` or `src/forms`. Move them into a subfolder if you like; later exports update them where they live. Then add the new module to the [source modules table](#source-modules) and to the function reference in [DOCS.md](DOCS.md#function-reference).

### Rename or delete a component

Do it in the VBE, then `./scripts/dev.ps1 -Export`. The export warns about files in `src/` that no longer have a component. Add `-Prune` to `export.ps1` to delete them (run `./scripts/export.ps1 -Prune`), or delete them by hand. Update the documentation to match.

### Change the ribbon

Edit `ribbon/customUI14.xml` and keep `ribbon/customUI.xml` in step with it. Icons go in `ribbon/images/`, and each image id must be listed in the matching file under `ribbon/_rels/`. Every callback name in the XML needs a procedure with the right signature, for example:

```vba
Sub MyButtonClick(control As IRibbonControl)
```

`check.ps1` catches a missing procedure. If the tab does not appear, enable *File, Options, Advanced, General, Show add-in user interface errors* and restart Excel. Excel then reports the exact XML or relationship problem.

## 8. Documenting procedures

The end-user function reference lives in [DOCS.md](DOCS.md#function-reference). Every public procedure gets an entry in the section of its module, and procedures called from the ribbon are marked with the button they belong to. Write entries for the people who use the released file, not for contributors. DOCS.md groups entries by task, not by module name, so use the [source modules table](#source-modules) to decide which section a procedure belongs in.

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

### List procedures (to seed the documentation)

Run from the repo root to list every procedure in every module as a starting point:

```powershell
Get-ChildItem src -Recurse -Include *.bas,*.cls | ForEach-Object {
    $m = $_.BaseName
    Select-String -Path $_.FullName -Pattern '^\s*(Public\s+|Private\s+)?(Sub|Function)\s+\w+' |
        ForEach-Object { "$m | $($_.Line.Trim())" }
}
```

Private procedures appear in the output too; document only the public ones and the ribbon callbacks.

## 9. Versioning and releases

The version lives in one place, the `VERSION` file, for example `2.1.0`. The build writes it into the generated `modVersion` module, so the About form and a future update check read `ADDIN_VERSION`. Built files are named `MayaTools-v<version>.xlam`; dev builds use `<version>-dev`.

### Publish a release

1. Commit your work and make sure `git status` is clean.
2. Run:

   ```powershell
   ./scripts/release.ps1 -Version 2.1.0
   ```

The script:

1. refuses if the tag already exists or the working tree is dirty;
2. writes `VERSION`;
3. runs `check.ps1`, then builds `dist\MayaTools-v2.1.0.xlam` with Excel;
4. copies the file to `releases/MayaTools-v2.1.0.xlam`, commits it, tags `v2.1.0`, and pushes the commit and the tag.

No token is needed on your machine; it only uses `git push`. Pushing the tag starts `.github/workflows/release.yml` on a GitHub-hosted runner, which:

1. runs the source checks;
2. verifies `VERSION` matches the tag, the committed `.xlam` exists, and it contains `xl/vbaProject.bin` and the ribbon;
3. creates the GitHub Release with the `.xlam` attached, generated release notes, and the file's SHA-256 in the notes. A version with a hyphen (`2.1.0-beta`) is marked as a pre-release.

CI cannot prove that the binary was built from the source (only Excel can build it), so the recorded SHA-256 is what makes the asset traceable.

`.github/workflows/check.yml` also runs `check.ps1` on every push and pull request.

### Before you release

- The installation steps in [DOCS.md](DOCS.md#installation) still match how the file is installed and named.
- New or changed procedures are documented in the [function reference](DOCS.md#function-reference).

### Fix a bad release

Prefer a new patch version. If you must redo the same version:

```powershell
git tag -d v2.1.0
git push origin :refs/tags/v2.1.0
```

then delete the release on GitHub (*Releases, the release, Delete*) and run `release.ps1` again.

## 10. Commit conventions

- One topic per commit, in the form `<module or form>: <what changed>`, for example `aboutDeveloper: show version`.
- Commit `.frm` and `.frx` together.
- `.frm` diffs are readable (positions, captions, fonts). `.frx` shows as a binary change. A re-exported form can reorder a few lines, so commit only what you meant to change.
- Pipeline changes (`scripts/`, `.github/`) go in their own commits, separate from add-in changes.
- Documentation and license changes (`README.md`, `DOCS.md`, `dev.md`, `LICENSE`) go in their own commits, in the form `docs: <what changed>`.

## 11. Troubleshooting

These are problems when building or developing. For problems using the installed add-in, see the [user troubleshooting table](DOCS.md#troubleshooting).

| Symptom | Cause and fix |
| --- | --- |
| `Programmatic access to Visual Basic Project is not trusted` / "Cannot access the VBA project" | Enable the Trust Center setting ([section 2](#2-prerequisites)). If it keeps reverting, a policy controls it |
| "The VBA project is locked with a password" | Remove the password in the VBE (*Tools, VBAProject Properties, Protection*), save, close Excel |
| "0 components" | Macros blocked for the file, or the project is locked. Check the Trust Center and unblock the file |
| "Excel is running" | Close Excel, including hidden processes (`Get-Process EXCEL`), or use `-CloseExcel` |
| "The dev copy was saved after the last build" | You have unexported VBE edits. Run `dev.ps1 -Export`, or `-Force` to discard them |
| `check.ps1` reports non-ASCII bytes | Replace the characters with `ChrW(code)`. The report names the file |
| Compile error "User-defined type not defined" | A library reference is missing; add it to `references.txt` |
| Duplicate definition errors on compile | A component exists twice (for example `modVersion` in `src/`), or two files share a name in different folders |
| Form looks broken after a build | The `.frx` is missing or not next to the `.frm`, or was changed by line-ending conversion |
| Ribbon tab missing | Enable add-in UI errors ([section 7](#change-the-ribbon)). Check both XML files are valid and the old add-in is not loaded as well |
| Ribbon button does nothing | Callback name in the XML has no matching procedure (`check.ps1` reports it) |
| `release.ps1` says "Source checks failed" with no details above | A leftover exit code from an earlier command; add `exit 0` as the last line of `check.ps1` |
| Release workflow: "file is missing" | The tag was created by hand. Delete it and use `release.ps1` so the build is committed before the tag |
| Downloaded `.xlam` is blocked | Run `Unblock-File` on it, or right-click, Properties, Unblock |

## 12. Reference: the original add-in

`_oldversions/MayaTools v2.xlam` is the original, password-locked add-in. Its source was exported manually once to create `src/`. Keep it for reference only; do not use it as input to the scripts.

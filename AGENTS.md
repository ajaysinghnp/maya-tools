# AGENTS.md

Instructions for AI coding agents (GitHub Copilot, Claude Code, and others) working in this repository. Humans: see [dev.md](dev.md) for the full developer guide.

## Project

MayaTools is an Excel add-in (`.xlam`) of VBA tools with its own ribbon tab: Nepali (Bikram Sambat) date functions, number-to-words, and helpers for worksheets, workbooks, formatting, formulas, pivot tables and printing.

The `.xlam` is a **build artifact**. The **source of truth is the plain-text files** in this repo.

| Path | Contents |
| --- | --- |
| `src/modules/`, `src/classes/`, `src/forms/` | VBA source: `.bas`, `.cls`, `.frm` + `.frx` |
| `ribbon/` | `customUI.xml` (Excel 2007), `customUI14.xml` (2010+), `images/`, `_rels/` |
| `scripts/` | PowerShell pipeline: `dev.ps1`, `check.ps1`, `build.ps1`, `export.ps1`, `release.ps1` |
| `releases/` | Built `.xlam` files, committed by `release.ps1` only |
| `README.md`, `DOCS.md`, `dev.md` | Landing page, end-user guide and function reference, developer guide |
| `_oldversions/` | Original password-locked add-in. Reference only, never edit |

## What you can and cannot do

Building and exporting need **Windows with desktop Excel**, which you normally do not have. So:

- **Do**: edit `.bas` / `.cls` files, ribbon XML, scripts and Markdown docs.
- **Do**: run `./scripts/check.ps1` when PowerShell 7 (`pwsh`) is available. It needs no Excel.
- **Do not** run `make-icons.ps1` (it needs Windows drawing, WPF). After editing `scripts/icons.json`, tell the developer to run it.
- **Do not** run `build.ps1`, `build-dev.ps1`, `export.ps1`, `dev.ps1` or `release.ps1`. They drive Excel and can overwrite the developer's unexported edits.
- **Do not** commit, push, tag or publish a release unless asked.
- **Do not** edit anything in `releases/`, `dist/` or `_oldversions/`, and never touch `.xlam` files.

## Source rules for VBA (enforced by `check.ps1`)

1. **Pure ASCII only.** No accented letters, smart quotes, en dashes or emoji in `.bas`, `.cls` or `.frm`, including comments and strings. Write special characters with `ChrW()`, for example `ChrW(176)` for the degree sign. Nepali text uses a legacy font-encoded scheme that stays ASCII apart from `ChrW(164)` and `ChrW(182)`; do not "fix" it.
2. **CRLF line endings** in `.bas`, `.cls` and `.frm`. Preserve them when editing.
3. **Component name equals file name** (`HelperSubroutines` is `HelperSubroutines.bas`).
4. **Forms**: never hand-edit the `Begin` / `End` control block, the header, or any `.frx`. Layout changes belong in the VBE designer. Code-behind below the `Attribute VB_...` lines may be edited as text. Keep `.frm` and `.frx` together.
5. **No code in sheets or `ThisWorkbook`.**
6. **Do not create `modVersion`.** It is generated at build time.
7. **No password** on the VBA project. Extra library references go in `references.txt` as `Name|GUID|Major|Minor`.
8. Match the style of the surrounding module (naming, indentation, error handling). Do not reformat or rename code you were not asked to change.

## Ribbon rules

- Keep `ribbon/customUI.xml` and `ribbon/customUI14.xml` in step.
- Every callback (`onAction`, `getImage`, ...) needs a matching `Sub` or `Function` in `src/` with the right signature, for example `Sub Ribbon_WebLink(control As IRibbonControl)`.
- **All ribbon callbacks live in `src/modules/RibbonCallbacks.bas`**, named `Ribbon_<Action>`. A callback checks it can run (for example `RequireWorkbook()`, because `ActiveWorkbook` is `Nothing` when no workbook is open and causes run-time error 91) and then calls the procedure that does the work. Put one- or two-line actions directly in `RibbonCallbacks`. Never add a callback in another module. Use `ThisWorkbook` (the add-in), never `ActiveWorkbook`, for things that must work without an open workbook.
- Every image id must be listed in the matching file under `ribbon/_rels/`.
- **Icons are generated.** Never hand-edit `ribbon/images/*.png`, the files in `ribbon/_rels/` or `THIRD-PARTY-NOTICES.md`. To add or change an icon, edit `scripts/icons.json` (image id -> Lucide icon name) and ask the developer to run `./scripts/make-icons.ps1`. All icons stay one colour, 32x32, transparent. Use a unique image id per icon.

## Documentation rules

- End users read [DOCS.md](DOCS.md). When you add, rename or remove a public procedure or ribbon button, update its entry in the DOCS.md function reference, using the entry template in [dev.md](dev.md#8-documenting-procedures). Write for people using the released file, not for contributors.
- Keep the module table in [dev.md](dev.md#source-modules) in step when adding or removing a module.
- Do not invent behaviour. Read the VBA before documenting what a procedure does.
- The project is MIT licensed ([LICENSE](LICENSE)). Do not remove copyright notices.

## Commits (only when asked)

- One topic per commit: `<module or form>: <what changed>`, for example `SpellNumbers: fix rounding`.
- Documentation: `docs: <what changed>`.
- Pipeline changes (`scripts/`, `.github/`) go in their own commits.
- Commit `.frm` and `.frx` together. `git push` is always a manual step by the developer.

## Before you finish

1. Run `./scripts/check.ps1` if you can, or re-check by hand: ASCII only, CRLF, every ribbon callback has a procedure.
2. Tell the developer which files you changed and that they must run `./scripts/dev.ps1` on their Windows machine to build and test in Excel.

# Copilot instructions for MayaTools

MayaTools is an Excel add-in (`.xlam`) of VBA tools with a custom ribbon tab (Nepali dates, number-to-words, worksheet/workbook/formatting helpers). The `.xlam` is a build artifact; the plain-text files in this repo are the source of truth. Full rules are in `AGENTS.md`; the essentials are below.

## Where things live

- VBA source: `src/modules/*.bas`, `src/classes/*.cls`, `src/forms/*.frm` + `.frx`
- Ribbon: `ribbon/customUI.xml` (Excel 2007) and `ribbon/customUI14.xml` (2010+), icons in `ribbon/images/`
- Scripts: `scripts/*.ps1` (PowerShell 7)
- Docs: `README.md` (landing page), `DOCS.md` (end-user guide and function reference), `dev.md` (developer guide)

## VBA rules

- **ASCII only** in `.bas`, `.cls`, `.frm`, including comments and strings. Use `ChrW(code)` for special characters (for example `ChrW(176)`). Never add smart quotes, en dashes or accented letters.
- Keep **CRLF** line endings in VBA files.
- Component name must equal the file name.
- Do not hand-edit a form's `Begin` / `End` control block, its header, or any `.frx`. Edit code-behind only.
- No code in sheet modules or `ThisWorkbook`. Do not create `modVersion` (generated at build).
- Follow the style of the surrounding module. Do not reformat or rename code you were not asked to change.

## Ribbon rules

- Keep both `customUI` files in sync.
- Every callback in the XML must have a matching procedure in `src/` with the correct signature, for example `Sub Ribbon_WebLink(control As IRibbonControl)`.
- All ribbon callbacks live in `src/modules/RibbonCallbacks.bas`, named `Ribbon_<Action>`. They check they can run (for example `RequireWorkbook()`; `ActiveWorkbook` is `Nothing` with no open workbook, which gives run-time error 91) and call the procedure that does the work. Never add a callback in another module. Use `ThisWorkbook`, not `ActiveWorkbook`, for actions that must work without an open workbook.
- Every image id must be listed under `ribbon/_rels/`.
- Icons are generated from Lucide: edit `scripts/icons.json` and ask the developer to run `./scripts/make-icons.ps1` (needs Windows). Never hand-edit `ribbon/images/*.png`, `ribbon/_rels/*` or `THIRD-PARTY-NOTICES.md`.

## Docs

- When a public procedure or ribbon button changes, update its entry in `DOCS.md` (task-based sections, written for end users). The entry template is in `dev.md`, section 8.

## Do not

- Run `build.ps1`, `build-dev.ps1`, `export.ps1`, `dev.ps1` or `release.ps1` (they need desktop Excel and can overwrite unexported edits).
- Edit `releases/`, `dist/`, `_oldversions/` or any `.xlam`.
- Commit, push or tag unless asked.

You can validate source with `./scripts/check.ps1` (no Excel needed). Commit messages use `<module or form>: <what changed>`, or `docs: <what changed>` for documentation.

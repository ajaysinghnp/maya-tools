# CLAUDE.md

@AGENTS.md

## Claude Code notes

- The rules in AGENTS.md apply in full. Read them before editing.
- Excel is not available in this environment, so never attempt to build, export or open the add-in. Verify with `pwsh ./scripts/check.ps1` only.
- Prefer small, targeted edits to `.bas` files. Re-read a file before editing it again, and confirm CRLF line endings and ASCII-only content after each edit.
- When a change affects a public procedure, update DOCS.md in the same task and mention it in your summary.
- Do not run `git commit` or `git push` unless the developer explicitly asks.

# PowershellScripts

This is my personal collection of PowerShell scripts, built up over time to solve various practical cases in day-to-day administration and automation. It serves as a repository of proven tools and a reference for future use.

## Script header conventions

Every script uses PowerShell comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, `.NOTES`), so `Get-Help .\Script.ps1 -Full` works out of the box. The `.NOTES` section always ends with:

```text
Version: <MAJOR>.<MINOR> (<YYYY-MM-DD>)
Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
```

**Versioning (`MAJOR.MINOR`, e.g. `1.0`, `1.1`, `2.0`):**

- New scripts start at `1.0`.
- Bump **MINOR** (`1.0` → `1.1`) for a backward-compatible behavior change: bug fix, new optional parameter, new feature, expanded functionality. Pure documentation/comment/translation edits do **not** require a bump.
- Bump **MAJOR** (`1.x` → `2.0`) for a breaking change: a parameter is removed/renamed, a default behavior changes, the output shape changes, or the script is rewritten in a way that changes how it's invoked or what it does.
- The date in parentheses is updated whenever the file is edited for any reason (including doc-only changes) — it tracks "last touched," independent of whether the version number changed.

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file.

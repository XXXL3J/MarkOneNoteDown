# MarkOneNoteDown

Batch convert OneNote notebooks to Markdown via Pandoc.

## Requirements

| Software | Version | Notes |
|----------|---------|-------|
| Windows | ≥ 10 | |
| PowerShell | 5.x ~ 7.0.x | ≥ 7.1 **unsupported** ([why?](https://github.com/PowerShell/PowerShell/issues/13138)) |
| OneNote | ≥ 2016 | Desktop version, not Windows Store |
| Word | ≥ 2016 | Desktop version |
| Pandoc | ≥ 2.11 | `choco install pandoc` or [pandoc.org](https://pandoc.org/installing.html) |

## Quick Start

**Double-click `run.bat`** — it checks dependencies, then converts.

First-time setup:

| # | Action |
|---|--------|
| 1 | Import Onetastic macro to expand collapsed paragraphs: `assets/Onetastic-ExpandAllParagraphs.xml` |
| 2 | (Optional) `cp config.example.ps1 config.ps1` and edit — skips interactive prompts on next run |

## Common Errors

| Error | Fix |
|-------|-----|
| Script blocked | `Set-ExecutionPolicy Bypass -Scope Process -Force` |
| `HRESULT E_FAIL` | Use PS 5.x or 7.0.x |
| `0x80042006` | Run as Administrator, use absolute path in config |
| `0x800706BA` | Add `pandoc.exe` to Defender exclusions |
| `Class not registered` | Install Microsoft Word |

## Project Structure

```
├── run.bat, setup.ps1           # Launcher & dependency checker
├── MarkOneNoteDown.ps1           # Entry point
├── config.example.ps1            # Configuration template
├── MarkOneNoteDown.Tests.ps1     # Tests (.\test\test.ps1)
└── src/
    ├── Private/
    │   ├── Dependencies.ps1      # Validate-Dependencies + Test-Preflight
    │   ├── Config.ps1            # Configuration management
    │   ├── FileUtils.ps1         # Path & name utilities
    │   ├── OneNoteCOM.ps1        # OneNote COM interop
    │   └── Conversion.ps1        # Core conversion engine
    └── Public/
        ├── Convert-OneNote2MarkDown.ps1
        └── Print-ConversionErrors.ps1
```

## Credits

Forked from [ConvertOneNote2MarkDown](https://github.com/theohbrothers/ConvertOneNote2MarkDown) by [@theohbrothers](https://github.com/theohbrothers), originally by [@SjoerdV](https://github.com/SjoerdV). Restructured as a modular script.

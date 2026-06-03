[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $ConversionConfigurationExportPath
,
    [Parameter()]
    [switch]
    $Exit
)

# ============================================================
# Pre-flight: PowerShell 7.1+ does not support Win32 GAC
# assemblies required for OneNote COM interop. Fail fast.
# ============================================================
if (-not $Exit -and $PSVersionTable.PSVersion -ge [version]'7.1') {
    Write-Error @"
PowerShell $($PSVersionTable.PSVersion) is not supported.
This script requires Windows PowerShell 5.1 or PowerShell Core 6.x-7.0.x.
To install PowerShell 7.0.13 (portable, no conflict with existing versions):
  1. Download https://github.com/PowerShell/PowerShell/releases/download/v7.0.13/PowerShell-7.0.13-win-x64.zip
  2. Extract to C:\PowerShell-7.0.13-win-x64
  3. Run: C:\PowerShell-7.0.13-win-x64\pwsh.exe -File .\MarkOneNoteDown.ps1
See README.md for more details.
"@
    exit 1
}

# ============================================================
# Project root — used by all dot-sourced modules for path
# resolution, since $PSScriptRoot differs per module file.
# ============================================================
$Script:ProjectRoot = $PSScriptRoot

# ============================================================
# Dot-source all module functions
# ============================================================

# Private (internal) functions
. "$PSScriptRoot\src\Private\Dependencies.ps1"
. "$PSScriptRoot\src\Private\Config.ps1"
. "$PSScriptRoot\src\Private\FileUtils.ps1"
. "$PSScriptRoot\src\Private\OneNoteCOM.ps1"
. "$PSScriptRoot\src\Private\Conversion.ps1"

# Public functions
. "$PSScriptRoot\src\Public\Print-ConversionErrors.ps1"
. "$PSScriptRoot\src\Public\Convert-OneNote2MarkDown.ps1"

# ============================================================
# Entry-point: when invoked directly as a script
# ============================================================
if (-not $Exit) {
    # Quick preflight check (non-fatal: only the PS version check blocks execution)
    $preflightOk = Test-Preflight
    if (-not $preflightOk) {
        Write-Host "Fix the issues above before running.`nSee README.md #requirements for details." -ForegroundColor Red
        Write-Host "Run .\setup.ps1 to re-check dependencies at any time." -ForegroundColor Yellow
        exit 1
    }

    $params = @{
        ConversionConfigurationExportPath = $ConversionConfigurationExportPath
    }
    Convert-OneNote2MarkDown @params
}

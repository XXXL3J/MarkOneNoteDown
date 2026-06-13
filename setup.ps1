[CmdletBinding()]
param (
    [switch]$InstallPandoc
)

Write-Host "=== MarkOneNoteDown Setup ===" -ForegroundColor Cyan

$issues = @()

# 1. PowerShell version
Write-Host "`n[1/3] PowerShell version..." -ForegroundColor Yellow
$psVersion = $PSVersionTable.PSVersion
Write-Host "  Detected: $psVersion"
if ($psVersion -ge [version]'7.1') {
    Write-Host "  [WARN] PS 7.1+ is NOT supported. Use PS 5.x or 7.0.x." -ForegroundColor Red
    $issues += "PowerShell $psVersion unsupported (need 5.x-7.0.x)"
} elseif ($psVersion.Major -ge 5) {
    Write-Host "  [OK]   Supported version." -ForegroundColor Green
} else {
    Write-Host "  [WARN] PS $psVersion may not work. Upgrade to PS 5.x+." -ForegroundColor Red
    $issues += "PowerShell $psVersion too old (need 5.x+)"
}

# 2. Pandoc
Write-Host "`n[2/3] Pandoc..." -ForegroundColor Yellow
$pandoc = Get-Command pandoc.exe -ErrorAction SilentlyContinue
if (-not $pandoc) {
    # Fallback: check common user-install locations
    $userPaths = @(
        "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
        "$env:ProgramFiles\Pandoc\pandoc.exe"
        "${env:ProgramFiles(x86)}\Pandoc\pandoc.exe"
    )
    foreach ($p in $userPaths) {
        if (Test-Path $p) {
            $parentDir = Split-Path $p -Parent
            if ($env:Path -notlike "*$parentDir*") {
                $env:Path = "$parentDir;$env:Path"
            }
            $pandoc = Get-Command pandoc.exe -ErrorAction SilentlyContinue
            if (-not $pandoc) { $pandoc = $p }
            Write-Host "  [OK]   Found at: $p" -ForegroundColor Green
            break
        }
    }
}
if ($pandoc) {
    $ver = & pandoc --version | Select-Object -First 1
    $src = if ($pandoc -is [System.Management.Automation.CommandInfo]) { $pandoc.Source } else { $pandoc }
    Write-Host "  [OK]   $ver" -ForegroundColor Green
    Write-Host "  Path:  $src" -ForegroundColor DarkGray
} elseif ($InstallPandoc) {
    Write-Host "  Installing Pandoc via winget..." -ForegroundColor Yellow
    winget install pandoc --accept-source-agreements --accept-package-agreements
    # Refresh PATH (both Machine and User scope)
    $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machinePath;$userPath"
    $pandoc = Get-Command pandoc.exe -ErrorAction SilentlyContinue
    if ($pandoc) { Write-Host "  [OK]   Installed successfully." -ForegroundColor Green }
    else { $issues += "Pandoc not found" }
} else {
    Write-Host "  [MISS] Pandoc not found. Run with -InstallPandoc to install via winget." -ForegroundColor Red
    $issues += "Pandoc not found"
}

# 3. OneNote assemblies (Windows only)
Write-Host "`n[3/3] OneNote assemblies..." -ForegroundColor Yellow
if ($env:OS -imatch 'Windows') {
    if (Get-Item -Path "$env:windir\assembly\GAC_MSIL\*onenote*" -ErrorAction SilentlyContinue) {
        Write-Host "  [OK]   OneNote assemblies detected." -ForegroundColor Green
    } else {
        Write-Host "  [MISS] OneNote 2016+ (Desktop) required." -ForegroundColor Red
        $issues += "OneNote assemblies missing"
    }
} else {
    Write-Host "  [SKIP] Non-Windows system." -ForegroundColor DarkGray
}

# Summary
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
if ($issues.Count -eq 0) {
    Write-Host "All dependencies satisfied. Run .\MarkOneNoteDown.ps1 or double-click run.bat to start." -ForegroundColor Green
} else {
    Write-Host "$($issues.Count) issue(s) found:" -ForegroundColor Red
    $issues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host "`nFix the above, then re-run this script." -ForegroundColor Yellow
}

exit $issues.Count

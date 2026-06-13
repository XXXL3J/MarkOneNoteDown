Function Test-Preflight {
    [CmdletBinding()]
    [OutputType([bool])]
    param ()

    Write-Host "`nChecking dependencies..." -ForegroundColor Cyan

    $ok = $true

    # 1. PowerShell version
    $psOk = $PSVersionTable.PSVersion -lt [version]'7.1'
    if ($psOk) {
        "  [OK]   PowerShell $($PSVersionTable.PSVersion)" | Write-Host -ForegroundColor Green
    } else {
        "  [FAIL] PowerShell $($PSVersionTable.PSVersion) -- 7.1+ NOT supported. Use PS 5.x or 7.0.x." | Write-Host -ForegroundColor Red
        $ok = $false
    }

    # 2. Admin check -- OneNote COM may fail with 80080005 when running elevated
    $isAdmin = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        "  [WARN] Running as Administrator -- if OneNote COM fails, try running as normal user." | Write-Host -ForegroundColor Yellow
    }

    # 3. OneNote assemblies (Windows only)
    if ($env:OS -imatch 'Windows') {
        if (Get-Item -Path "$env:windir\assembly\GAC_MSIL\*onenote*" -ErrorAction SilentlyContinue) {
            "  [OK]   OneNote Desktop (assemblies detected)" | Write-Host -ForegroundColor Green
        } else {
            "  [WARN] OneNote 2016+ Desktop assemblies not found (needed for COM interop)" | Write-Host -ForegroundColor Yellow
        }
    }

    # 4. Pandoc
    $pandocPath = Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue
    if (-not $pandocPath) {
        # Fallback: check common user-install locations
        $userPaths = @(
            "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
            "$env:ProgramFiles\Pandoc\pandoc.exe"
            "${env:ProgramFiles(x86)}\Pandoc\pandoc.exe"
        )
        foreach ($p in $userPaths) {
            if (Test-Path $p) {
                $pandocPath = $p
                # Add to session PATH for downstream use
                $parentDir = Split-Path $p -Parent
                if ($env:Path -notlike "*$parentDir*") {
                    $env:Path = "$parentDir;$env:Path"
                }
                break
            }
        }
    }
    if ($pandocPath) {
        $ver = if ($pandocPath -is [System.Management.Automation.CommandInfo]) {
            (& pandoc --version 2>$null | Select-Object -First 1) -replace 'pandoc\s+', ''
        } else {
            (& "$pandocPath" --version 2>$null | Select-Object -First 1) -replace 'pandoc\s+', ''
        }
        $src = if ($pandocPath -is [System.Management.Automation.CommandInfo]) { $pandocPath.Source } else { $pandocPath }
        "  [OK]   Pandoc $ver -- $src" | Write-Host -ForegroundColor Green
    } else {
        "  [MISS] Pandoc not found in PATH. Install: https://pandoc.org/installing.html" | Write-Host -ForegroundColor Yellow
    }

    # 5. Word (check registry for Office 2016+)
    if ($env:OS -imatch 'Windows') {
        $wordPath = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe' -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty '(Default)' -ErrorAction SilentlyContinue
        if ($wordPath -and (Test-Path $wordPath)) {
            "  [OK]   Microsoft Word detected" | Write-Host -ForegroundColor Green
        } else {
            "  [WARN] Microsoft Word not detected (needed for .docx export)" | Write-Host -ForegroundColor Yellow
        }
    }

    Write-Host ""
    return $ok
}

Function Validate-Dependencies {
    [CmdletBinding()]
    param ()

    # Validate Powershell versions. Supported version are between 5.x (possibly lower) and 7.0.x
    if ($PSVersionTable.PSVersion -ge [version]'7.1') {
        throw "Unsupported Powershell version $( $PSVersionTable.PSVersion ), because it does not support importing Win32 GAC Assemblies. Supported versions are between Powershell 5.x and 7.0.x. See README.md for instructions to install Powershell 7.0.x"
    }

    # Validate assemblies
    if ( ($env:OS -imatch 'Windows') -and ! (Get-Item -Path $env:windir\assembly\GAC_MSIL\*onenote*) ) {
        "There are missing onenote assemblies. Please ensure the Desktop version of Onenote 2016 or above is installed." | Write-Warning
    }

    # Validate dependencies: pandoc
    $pandocPath = Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue
    if (-not $pandocPath) {
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
                $pandocPath = Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue
                if (-not $pandocPath) {
                    # If still not found by Get-Command, use the direct path
                    $pandocPath = $p
                }
                "Pandoc found at: $p" | Write-Host -ForegroundColor Green
                break
            }
        }
    }
    if (-not $pandocPath) {
        $title = "Pandoc not found"
        $msg = "Pandoc is required but was not found in PATH.`n`nWould you like to install Pandoc via winget (recommended)?"
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Install", "Install Pandoc via winget"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&Skip", "Skip installation (manual install required)"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $Host.UI.PromptForChoice($title, $msg, $options, 1)
        if ($result -eq 0) {
            Write-Host "Installing Pandoc via winget..." -ForegroundColor Yellow
            winget install pandoc --accept-source-agreements --accept-package-agreements | Out-Host
            # Refresh PATH (both Machine and User scope)
            $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
            $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
            $env:Path = "$machinePath;$userPath"
            $pandocPath = Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue
            if ($pandocPath) {
                "Pandoc installed successfully: $( $pandocPath.Source )" | Write-Host -ForegroundColor Green
            } else {
                throw "Pandoc installation may have failed. Please restart PowerShell and try again, or install manually from https://pandoc.org/installing.html"
            }
        } else {
            throw "Could not locate pandoc.exe. Please ensure pandoc is installed and available in PATH (Machine or User). If pandoc was just installed, you may need to restart Powershell or the computer for PATH to be updated."
        }
    }
}

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

    # 2. OneNote assemblies (Windows only)
    if ($env:OS -imatch 'Windows') {
        if (Get-Item -Path "$env:windir\assembly\GAC_MSIL\*onenote*" -ErrorAction SilentlyContinue) {
            "  [OK]   OneNote Desktop (assemblies detected)" | Write-Host -ForegroundColor Green
        } else {
            "  [WARN] OneNote 2016+ Desktop assemblies not found (needed for COM interop)" | Write-Host -ForegroundColor Yellow
        }
    }

    # 3. Pandoc
    $pandocPath = Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue
    if ($pandocPath) {
        $ver = (& pandoc --version 2>$null | Select-Object -First 1) -replace 'pandoc\s+', ''
        "  [OK]   Pandoc $ver -- $($pandocPath.Source)" | Write-Host -ForegroundColor Green
    } else {
        "  [MISS] Pandoc not found in PATH. Install: https://pandoc.org/installing.html" | Write-Host -ForegroundColor Yellow
    }

    # 4. Word (check registry for Office 2016+)
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
    if (! $pandocPath) {
        $title = "Pandoc not found"
        $msg = "Pandoc is required but was not found in PATH.`n`nWould you like to install Pandoc via winget (recommended)?"
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Install", "Install Pandoc via winget"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&Skip", "Skip installation (manual install required)"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $Host.UI.PromptForChoice($title, $msg, $options, 1)
        if ($result -eq 0) {
            Write-Host "Installing Pandoc via winget..." -ForegroundColor Yellow
            winget install pandoc --accept-source-agreements --accept-package-agreements | Out-Host
            # Refresh PATH
            $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')
            $pandocPath = Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue
            if ($pandocPath) {
                "Pandoc installed successfully: $( $pandocPath.Source )" | Write-Host -ForegroundColor Green
            } else {
                throw "Pandoc installation may have failed. Please restart PowerShell and try again, or install manually from https://pandoc.org/installing.html"
            }
        } else {
            throw "Could not locate pandoc.exe. Please ensure pandoc is installed for all users, and available in PATH. If pandoc was just installed using .msi or chocolatey, you may need to restart Powershell or the computer for pandoc to be set correctly in PATH."
        }
    }
}

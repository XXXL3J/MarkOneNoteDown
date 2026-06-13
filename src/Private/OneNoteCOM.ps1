Function New-OneNoteConnection {
    [CmdletBinding()]
    [OutputType([object])]
    param ()

    # Create a OneNote connection. See: https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
    if ($PSVersionTable.PSVersion.Major -le 5) {
        try {
            $OneNote = New-Object -ComObject OneNote.Application
            $OneNote
        } catch [System.Runtime.InteropServices.COMException] {
            $hresult = $_.Exception.HResult
            $msg = "Failed to connect to OneNote (COM error 0x{0:X8})." -f $hresult
            Write-Error $msg -ErrorAction Continue
            if ($hresult -eq 0x80080005) {
                Write-Host @"

=== CO_E_SERVER_EXEC_FAILURE (0x80080005) ===
This error means you are running PowerShell as Administrator.
OneNote COM does NOT work from an elevated process -- this is a Windows security restriction, not a bug.

FIX: Close this window, then open a NORMAL (non-admin) PowerShell:
  - Right-click in the folder -> "Open in Terminal" (NOT "Run as administrator")
  - Or: Win+R -> powershell -> cd to this folder -> run .\MarkOneNoteDown.ps1

OneNote 365 (Desktop) is fully supported -- just don't run as admin.
"@ -ForegroundColor Yellow
            }
            throw
        }
    }else {
        # Works between powershell 5.x (possibly lower) and 7.0, but not >= 7.1. 7.1 and above doesn't seem to support loading Win32 GAC Assemblies.
        try {
            if (Add-Type -Path $env:windir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll -PassThru) {
                $OneNote = [Microsoft.Office.Interop.OneNote.ApplicationClass]::new()
                $OneNote
            } else {
                Write-Error "Failed to make connection to OneNote." -ErrorAction Continue
                throw
            }
        } catch [System.Runtime.InteropServices.COMException] {
            $hresult = $_.Exception.HResult
            $msg = "Failed to connect to OneNote (COM error 0x{0:X8})." -f $hresult
            Write-Error $msg -ErrorAction Continue
            if ($hresult -eq 0x80080005) {
                Write-Host @"

=== CO_E_SERVER_EXEC_FAILURE (0x80080005) ===
This error means you are running PowerShell as Administrator.
OneNote COM does NOT work from an elevated process.

FIX: Close this window, then open a NORMAL (non-admin) PowerShell.
  - Right-click in the folder -> "Open in Terminal" (NOT "Run as administrator")
  - Or: Win+R -> powershell -> cd to this folder -> run .\MarkOneNoteDown.ps1
"@ -ForegroundColor Yellow
            }
            throw
        }
    }
}

Function Remove-OneNoteConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    )

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteConnection) | Out-Null
}

Function Get-OneNoteHierarchy {
    [CmdletBinding()]
    [OutputType([xml])]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    )

    # Open OneNote hierarchy
    [xml]$hierarchy = ""
    $OneNoteConnection.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

    $hierarchy
}

Function Get-OneNotePageContent {
    [CmdletBinding()]
    [OutputType([xml])]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $PageId
    )

    # Get page's xml content
    [xml]$page = ""
    $OneNoteConnection.GetPageContent($PageId, [ref]$page, 7)

    $page
}

Function Publish-OneNotePage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $PageId
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Destination
    ,
        [Parameter(Mandatory)]
        [ValidateSet('pfOneNote', 'pfPDF', 'pfXPS', 'pfWord', 'pfEMF', 'pfHTML', 'pfOneNote2007')]
        [ValidateNotNullOrEmpty()]
        [string]
        $PublishFormat
    )

    $OneNoteConnection.Publish($PageId, $Destination, $PublishFormat, "")
}

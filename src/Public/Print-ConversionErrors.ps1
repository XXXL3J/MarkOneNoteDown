Function Print-ConversionErrors {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorCollection
    )

    if ($null -ne $ErrorCollection -and $ErrorCollection.Count -gt 0) {
        "Conversion errors ($($ErrorCollection.Count)):" | Write-Host -ForegroundColor Yellow
        foreach ($err in $ErrorCollection) {
            if ($err -is [System.Management.Automation.ErrorRecord]) {
                "  $($err.Exception.Message)" | Write-Host -ForegroundColor Red
            }else {
                "  $err" | Write-Host -ForegroundColor Red
            }
        }
    }
}

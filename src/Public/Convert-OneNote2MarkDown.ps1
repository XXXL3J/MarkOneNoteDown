Function Convert-OneNote2MarkDown {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $ConversionConfigurationExportPath
    )

    try {
        Set-StrictMode -Version Latest

        # Fix encoding problems for languages other than English
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'

        $totalerr = @()

        # Validate dependencies
        Validate-Dependencies

        # Compile and validate configuration
        $config = Compile-Configuration | Validate-Configuration

        # Resolve relative paths to absolute (relative to script directory)
        foreach ($key in @('notesdestpath')) {
            if ($config[$key]['value'] -is [string] -and -not [io.path]::IsPathRooted($config[$key]['value'])) {
                $resolved = [io.path]::GetFullPath([io.path]::combine($Script:ProjectRoot, $config[$key]['value']))
                "Resolved '$key': $( $config[$key]['value'] ) → $resolved" | Write-Verbose
                $config[$key]['value'] = $resolved
            }
        }

        "Configuration:" | Write-Host -ForegroundColor Cyan
        $config | Print-Configuration

        # Connect to OneNote
        $OneNote = New-OneNoteConnection

        # Get the hierarchy of OneNote objects as xml
        $hierarchy = Get-OneNoteHierarchy -OneNoteConnection $OneNote

        # Get and validate the notebook(s) to convert
        $notebooks = @(
            if ($config['targetNotebook']['value']) {
                $hierarchy.Notebooks.Notebook | Where-Object { $_.Name -eq $config['targetNotebook']['value'] }
            }else {
                $hierarchy.Notebooks.Notebook
            }
        )
        if ($notebooks.Count -eq 0) {
            if ($config['targetNotebook']['value']) {
                throw "Could not find notebook of name '$( $config['targetNotebook']['value'] )'"
            }else {
                throw "Could not find notebooks"
            }
        }

        "`nNotebooks to convert:" | Write-Host -ForegroundColor Cyan
        $notebooks.name | Write-Host -ForegroundColor Green

        # Convert the notebook(s)
        $pageConversionConfigsAll = @()
        foreach ($notebook in $notebooks) {
            "`nConverting notebook '$( $notebook.name )'... (Ignoring deleted notes)" | Write-Host -ForegroundColor Cyan
            $pageConversionConfigs = New-SectionGroupConversionConfig -OneNoteConnection $OneNote -NotesDestination $config['notesdestpath']['value'] -Config $config -SectionGroups $notebook -LevelsFromRoot 0 -ErrorVariable +totalerr
            if ($pageConversionConfigs) {
                foreach ($pageCfg in $pageConversionConfigs) {
                    Convert-OneNotePage -OneNoteConnection $OneNote -Config $config -ConversionConfig $pageCfg -ErrorVariable +totalerr
                }
            }
            "`nDone converting notebook '$( $notebook.name )' with $( ($pageConversionConfigs | Measure-object).Count ) notes." | Write-Host -ForegroundColor Cyan
            $pageConversionConfigsAll += $pageConversionConfigs
        }

        # Export all Page Conversion Configuration objects as .json, which is useful for debugging
        if ($ConversionConfigurationExportPath) {
            "`nExporting Page Conversion Configuration as JSON file with $( $pageConversionConfigsAll.Count ) objects: $ConversionConfigurationExportPath" | Write-Host -ForegroundColor Cyan
            $pageConversionConfigsAll | ConvertTo-Json -Depth 100 | Out-File $ConversionConfigurationExportPath -Encoding utf8 -Force
        }
    }catch {
        if ($ErrorActionPreference -eq 'Stop') {
            throw
        }else {
            Write-Error -Message $_.Exception.Message
            Write-Error -Message $_.ScriptStackTrace
            "If you're seeing errors on the first run, check the FAQ in README.md to resolve them: " | Write-Host -ForegroundColor Yellow
            "  https://github.com/theohbrothers/ConvertOneNote2MarkDown#faq" | Write-Host -ForegroundColor Yellow
        }
    }finally {
        'Cleaning up...' | Write-Host -ForegroundColor Cyan

        # Disconnect OneNote connection
        if (Get-Variable -Name OneNote -ErrorAction SilentlyContinue) {
            Remove-OneNoteConnection -OneNoteConnection $OneNote
        }

        # Print any conversion errors
        Print-ConversionErrors -ErrorCollection $totalerr
        'Exiting.' | Write-Host -ForegroundColor Cyan
    }
}

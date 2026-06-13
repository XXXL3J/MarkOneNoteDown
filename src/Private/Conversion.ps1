Function New-SectionGroupConversionConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # The desired directory to store any converted Page(s) found in this Section Group's Section(s)
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $NotesDestination
    ,
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Section Group XML object(s)
        [Parameter(Mandatory)]
        [array]
        $SectionGroups
    ,
        [Parameter(Mandatory)]
        [int]
        $LevelsFromRoot
    ,
        [Parameter()]
        [switch]
        $AsArray
    )

    $sectionGroupConversionConfig = [System.Collections.ArrayList]@()

    # Build an object representing the conversion of a Section Group (treat a Notebook as a Section Group, it is no different)
    foreach ($sectionGroup in $SectionGroups) {
        # Skip over Section Groups in recycle bin
        if ((Get-Member -InputObject $sectionGroup -Name 'isRecycleBin') -and $sectionGroup.isRecycleBin -eq 'true') {
            continue
        }

        if ($LevelsFromRoot -eq 0) {
            "`nBuilding conversion configuration for $( $sectionGroup.name ) [Notebook]" | Write-Host -ForegroundColor DarkGreen
        }else {
            "`n$( '#' * ($LevelsFromRoot) ) Building conversion configuration for $( $sectionGroup.name ) [Section Group]" | Write-Host -ForegroundColor DarkGray
        }

        # Build this Section Group
        $cfg = [ordered]@{}
        $cfg['object'] = $sectionGroup # Keep a reference to the SectionGroup object
        $cfg['kind'] = 'SectionGroup'
        $cfg['id'] = $sectionGroup.ID # E.g. {9570CCF6-17C2-4DCE-83A0-F58AE8914E29}{1}{B0}
        $cfg['nameCompat'] = $sectionGroup.name | Remove-InvalidFileNameChars
        $cfg['levelsFromRoot'] = $LevelsFromRoot
        $cfg['uri'] = $sectionGroup.path # E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup
        $cfg['notesDirectory'] = [io.path]::combine( $NotesDestination.TrimEnd('/').TrimEnd('\').Replace('\', [io.path]::DirectorySeparatorChar), $cfg['nameCompat'] ) # No need to truncate. Section Group and Section names have a max length of 50, so we should never hit the absolute path, file name, or directory name limits on Windows
        $cfg['notesBaseDirectory'] = & {
            # E.g. 'c:\temp\notes\mynotebook\mysectiongroup'
            # E.g. levelsFromRoot: 1
            $split = $cfg['notesDirectory'].Split( [io.path]::DirectorySeparatorChar )
            # E.g. 5
            $totalLevels = $split.Count
            # E.g. 0..(5-1-1) -> 'c:\temp\notes\mynotebook'
            $split[0..($totalLevels - $cfg['levelsFromRoot'] - 1)] -join [io.path]::DirectorySeparatorChar
        }
        $cfg['notebookName'] = Split-Path $cfg['notesBaseDirectory'] -Leaf
        $cfg['pathFromRoot'] = $cfg['notesDirectory'].Replace($cfg['notesBaseDirectory'], '').Trim([io.path]::DirectorySeparatorChar)
        $cfg['pathFromRootCompat'] = $cfg['pathFromRoot'] | Remove-InvalidFileNameChars
        $cfg['notesDocxDirectory'] = [io.path]::combine( $cfg['notesBaseDirectory'], 'docx' )
        $cfg['directoriesToCreate'] = @()

        # Build this Section Group's sections
        $cfg['sections'] = [System.Collections.ArrayList]@()

        if (! (Get-Member -InputObject $sectionGroup -Name 'Section') -and ! (Get-Member -InputObject $sectionGroup -Name 'SectionGroup') ) {
            "Ignoring empty Section Group: $( $cfg['pathFromRoot'] )" | Write-Host -ForegroundColor DarkGray
        }

        if (Get-Member -InputObject $sectionGroup -Name 'Section') {
            foreach ($section in $sectionGroup.Section) {
                "$( '#' * ($LevelsFromRoot + 1) ) Building conversion configuration for $( $section.name ) [Section]" | Write-Host -ForegroundColor DarkGray

                $sectionCfg = [ordered]@{}
                $sectionCfg['notebookName'] = $cfg['notebookName']
                $sectionCfg['notesBaseDirectory'] = $cfg['notesBaseDirectory']
                $sectionCfg['notesDirectory'] = $cfg['notesDirectory']
                $sectionCfg['sectionGroupUri'] = $cfg['uri'] # Keep a reference to my Section Group Configuration object's uri
                $sectionCfg['sectionGroupName'] = $cfg['object'].name
                $sectionCfg['object'] = $section # Keep a reference to the Section object
                $sectionCfg['kind'] = 'Section'
                $sectionCfg['id'] = $section.ID # E.g {BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{B0}
                $sectionCfg['nameCompat'] = $section.name | Remove-InvalidFileNameChars
                $sectionCfg['levelsFromRoot'] = $cfg['levelsFromRoot'] + 1
                $sectionCfg['pathFromRoot'] = "$( $cfg['pathFromRoot'] )$( [io.path]::DirectorySeparatorChar )$( $sectionCfg['nameCompat'] )".Trim([io.path]::DirectorySeparatorChar) # No need to truncate. Section Group and Section names have a max length of 50, so we should never hit the absolute path, file name, or directory name limits on Windows
                $sectionCfg['pathFromRootCompat'] = $sectionCfg['pathFromRoot'] | Remove-InvalidFileNameChars
                $sectionCfg['uri'] = $section.path # E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup/mysection
                $sectionCfg['lastModifiedTime'] = [Datetime]::ParseExact($section.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $sectionCfg['lastModifiedTimeEpoch'] = [int][double]::Parse((Get-Date ((Get-Date $sectionCfg['lastModifiedTime']).ToUniversalTime()) -UFormat %s)) # Epoch

                $sectionCfg['pages'] = [System.Collections.ArrayList]@()

                # Build Section's pages
                if (Get-Member -InputObject $section -Name 'Page') {
                    foreach ($page in $section.Page) {
                        "$( '#' * ($LevelsFromRoot + 2) ) Building conversion configuration for $( $page.name ) [Page]" | Write-Host -ForegroundColor DarkGray

                        $previousPage = if ($sectionCfg['pages'].Count -gt 0) { $sectionCfg['pages'][$sectionCfg['pages'].Count - 1] } else { $null }
                        $pageCfg = [ordered]@{}
                        $pageCfg['notebookName'] = $cfg['notebookName']
                        $pageCfg['notesBaseDirectory'] = $cfg['notesBaseDirectory']
                        $pageCfg['notesDirectory'] = $cfg['notesDirectory']
                        $pageCfg['sectionGroupUri'] = $cfg['uri'] # Keep a reference to mt Section Group Configuration object's uri
                        $pageCfg['sectionGroupName'] = $cfg['object'].name
                        $pageCfg['sectionUri'] = $sectionCfg['uri'] # Keep a reference to my Section Configuration object's uri
                        $pageCfg['sectionName'] = $sectionCfg['object'].name
                        $pageCfg['object'] = $page # Keep a reference to my Page object
                        $pageCfg['kind'] = 'Page'
                        $pageCfg['id'] = $page.ID # E.g. {3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}
                        $pageCfg['nameCompat'] = $page.name | Remove-InvalidFileNameChars
                        $pageCfg['levelsFromRoot'] = $sectionCfg['levelsFromRoot']
                        $pageCfg['pathFromRoot'] = "$( $sectionCfg['pathFromRoot'] )$( [io.path]::DirectorySeparatorChar )$( $pageCfg['nameCompat'] )"
                        $pageCfg['pathFromRootCompat'] = $pageCfg['pathFromRoot'] | Remove-InvalidFileNameChars
                        $pageCfg['uri'] = "$( $sectionCfg['object'].path )/$( $page.name )" # There's no $page.path property, so we generate one. E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup/mysection/mypage
                        $pageCfg['dateTime'] = [Datetime]::ParseExact($page.dateTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                        $pageCfg['lastModifiedTime'] = [Datetime]::ParseExact($page.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                        $pageCfg['lastModifiedTimeEpoch'] = [int][double]::Parse((Get-Date ((Get-Date $pageCfg['lastModifiedTime']).ToUniversalTime()) -UFormat %s)) # Epoch
                        $pageCfg['pageLevel'] = $page.pageLevel -as [int]
                        $pageCfg['conversion'] = $config['conversion']['value']
                        $pageCfg['pagePrefix'] = & {
                            # 9 different scenarios
                            if ($pageCfg['pageLevel'] -eq 1) {
                                # 1 -> 1, 2 -> 1, or 3 -> 1
                                ''
                            }else {
                                if ($previousPage) {
                                    if ($previousPage['pageLevel'] -lt $pageCfg['pageLevel']) {
                                        # 1 -> 2, 1 -> 3, or 2 -> 3
                                        "$( $previousPage['filePathRel'] )$( [io.path]::DirectorySeparatorChar )"
                                    }elseif ($previousPage['pageLevel'] -eq $pageCfg['pageLevel']) {
                                        # 2 -> 2, or 3 -> 3
                                        "$( Split-Path $previousPage['filePathRel'] -Parent )$( [io.path]::DirectorySeparatorChar )"
                                    }else {
                                        # 3 -> 2 (or 4 -> 2, but 4th level subpages don't exist, but technically this supports it)
                                        $split = $previousPage['filePathRel'].Split([io.path]::DirectorySeparatorChar)
                                        $index = $pageCfg['pageLevel'] - 1 - 1 # If page level n, the prefix should be n-1
                                        if ($index -lt 0) {
                                            $index = 0 # The shallowest subpage must be a child of a first level page, i.e. $split[0]
                                        }
                                        "$( $split[0..$index] -join [io.path]::DirectorySeparatorChar )$( [io.path]::DirectorySeparatorChar )"
                                    }
                                }else {
                                    '' # Should never end up here
                                }
                            }
                        }
                        # Win32 path limits. E.g. 'C:\path\to\file' or 'C:\path\to\folder'
                        #   Absolute path:
                        #   - Win32: Max 259 characters for files, Max 247 characters for directories.
                        #   File or directory name:
                        #   - Max 255 characters long for file or folder names
                        # Non-Win32 path limits. E.g. '\\?\C:\path\to\file' or '\\?\C:\path\to\folder'. Prefixing with '\\?\' allows Windows Powershell <= 5 (based on Win32) to support long absolute paths.
                        #   Absolute path:
                        #   - N.A.
                        #   File or directory name:
                        #   - Max 255 characters long for file or folder names
                        # See: https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file?redirectedfrom=MSDN#maxpath

                        # Normalize the final .md file path. Page names can be very long, and can exceed the max absolute path length, or max file or folder name on a Windows system.
                        $pageCfg['filePathRel'] = & {
                            $filePathRel = "$( $pageCfg['pagePrefix'] )$( $pageCfg['nameCompat'] )"

                            # in case multiple pages with the same name exist in a section, postfix the filename
                            $recurrence = 0
                            foreach ($p in $sectionCfg['pages']) {
                                if ($p['pagePrefix'] -eq $pageCfg['pagePrefix'] -and $p['pathFromRoot'] -eq $pageCfg['pathFromRoot']) {
                                    $recurrence++
                                }
                            }
                            if ($recurrence -gt 0) {
                                $filePathRel = "$filePathRel-$recurrence"
                            }
                            $filePathRel | Truncate-PathFileName -Length $config['mdFileNameAndFolderNameMaxLength']['value'] # Truncate to no more than 255 characters so we don't hit the folder name limit on most file systems on Windows / Linux
                        }
                        $pageCfg['filePathRelUnderscore'] = $pageCfg['filePathRel'].Replace( [io.path]::DirectorySeparatorChar, '_' )
                        $pageCfg['filePathNormal'] = & {
                            $pathWithoutExtension = if ($config['prefixFolders']['value'] -eq 2) {
                                [io.path]::combine( $cfg['notesDirectory'], $sectionCfg['nameCompat'], "$( $pageCfg['filePathRelUnderscore'] )" )
                            }else {
                                [io.path]::combine( $cfg['notesDirectory'], $sectionCfg['nameCompat'], "$( $pageCfg['filePathRel'] )" )
                            }
                            "$( $pathWithoutExtension | Truncate-PathFileName -Length ($config['mdFileNameAndFolderNameMaxLength']['value'] - 3) ).md" # Truncate to no more than 255 characters so we don't hit the file name limit on Windows / Linux
                        }
                        $pageCfg['filePathLong'] = "\\?\$( $pageCfg['filePathNormal'] )" # A non-Win32 path. Prefixing with '\\?\' allows Windows Powershell <= 5 (based on Win32) to support long absolute paths.
                        $pageCfg['filePath'] = if ($PSVersionTable.PSVersion.Major -le 5) {
                            $pageCfg['filePathLong'] # Add support for long paths on Powershell 5
                        }else {
                            $pageCfg['filePathNormal'] # Powershell Core supports long file paths
                        }
                        $pageCfg['fileDirectory'] = Split-Path $pageCfg['filePathNormal'] -Parent
                        $pageCfg['fileName'] = Split-Path $pageCfg['filePathNormal'] -Leaf
                        $pageCfg['fileExtension'] = if ($pageCfg['filePathNormal'] -match '(\.[^.]+)$') { $matches[1] } else { '' }
                        $pageCfg['fileBaseName'] = $pageCfg['fileName'] -replace "$( [regex]::Escape($pageCfg['fileExtension']) )$", ''
                        $pageCfg['pdfExportFilePathTmp'] = [io.path]::combine( (Split-Path $pageCfg['filePath'] -Parent ), "$( $pageCfg['id'] )-$( $pageCfg['lastModifiedTimeEpoch'] ).pdf" ) # Publishing a .pdf seems to be limited to 204 characters. So we will export the .pdf to a unique file name, then rename it to the actual name
                        $pageCfg['pdfExportFilePath'] = if ( ($pageCfg['fileName'].Length + ('.pdf'.Length - '.md'.Length)) -le $config['mdFileNameAndFolderNameMaxLength']['value']) {
                            $pageCfg['filePath'] -replace '\.md$', '.pdf'
                        }else {
                            $pageCfg['filePath'] -replace '.\.md$', '.pdf' # Trim 1 character in the basename when replacing the extension
                        }
                        $pageCfg['levelsPrefix'] = if ($config['medialocation']['value'] -eq 2) {
                            ''
                        }else {
                            if ($config['prefixFolders']['value'] -eq 2) {
                                "$( '../' * ($pageCfg['levelsFromRoot'] + 1 - 1) )"
                            }else {
                                "$( '../' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel'] - 1) )"
                            }
                        }
                        $pageCfg['tmpPath'] = & {
                            $dateNs = Get-Date -Format "yyyy-MM-dd-HH-mm-ss-fffffff"
                            if ($env:OS -match 'windows') {
                                # Ensure $env:TEMP is not MSDOS 8.3 shortened name, but the actual full path. See: https://superuser.com/questions/1524767/powershell-uses-the-short-8-3-form-for-envtemp
                                [io.path]::combine((Get-Item $env:TEMP).FullName, $cfg['notebookName'], $dateNs)
                            }else {
                                [io.path]::combine('/tmp', $cfg['notebookName'], $dateNs)
                            }
                        }
                        $pageCfg['mediaParentPath'] = if ($config['medialocation']['value'] -eq 2) {
                            $pageCfg['fileDirectory']
                        }else {
                            $cfg['notesBaseDirectory']
                        }
                        $pageCfg['mediaPath'] = [io.path]::combine( $pageCfg['mediaParentPath'], 'media' )
                        $pageCfg['mediaParentPathPandoc'] = [io.path]::combine( $pageCfg['tmpPath'] ).Replace( [io.path]::DirectorySeparatorChar, '/' ) # Pandoc outputs paths in markdown with with front slahes after the supplied <mediaPath>, e.g. '<mediaPath>/media/image.png'. So let's use a front-slashed supplied mediaPath
                        $pageCfg['mediaPathPandoc'] = [io.path]::combine( $pageCfg['tmpPath'], 'media').Replace( [io.path]::DirectorySeparatorChar, '/' ) # Pandoc outputs paths in markdown with with front slahes after the supplied <mediaPath>, e.g. '<mediaPath>/media/image.png'. So let's use a front-slashed supplied mediaPath
                        $pageCfg['docxExportFilePath'] = if ($config['docxNamingConvention']['value'] -eq 1) {
                            [io.path]::combine( $cfg['notesDocxDirectory'], "$( $pageCfg['id'] )-$( $pageCfg['lastModifiedTimeEpoch'] ).docx" )
                        }else {
                            [io.path]::combine( $cfg['notesDocxDirectory'], "$( $pageCfg['pathFromRootCompat'] ).docx" )
                        }
                        $pageCfg['insertedAttachments'] = @(
                            & {
                                $pagexml = Get-OneNotePageContent -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID

                                # Search recursively for all attachment(s). This includes attachments nested in tables etc.
                                $ns = new-object Xml.XmlNamespaceManager $pagexml.NameTable
                                $ns.AddNamespace("one", $pagexml.DocumentElement.NamespaceURI)
                                $insertedFiles = $pagexml.SelectNodes("//one:InsertedFile", $ns)
                                foreach ($i in $insertedFiles) {
                                    $attachmentCfg = [ordered]@{}
                                    $attachmentCfg['object'] =  $i
                                    $attachmentCfg['nameCompat'] =  $i.preferredName | Remove-InvalidFileNameCharsInsertedFiles
                                    $attachmentCfg['markdownFileName'] =  $attachmentCfg['nameCompat'] | Encode-Markdown -Uri
                                    $attachmentCfg['source'] =  $i.pathCache
                                    $attachmentCfg['destination'] =  [io.path]::combine( $pageCfg['mediaPath'], $attachmentCfg['nameCompat'] )

                                    $attachmentCfg
                                }
                            }
                        )
                        $pageCfg['mutations'] = @(
                            # Markdown mutations. Each search and replace is done against a string containing the entire markdown content

                            foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                                @{
                                    description = 'Change inserted attachment(s) filename references'
                                    replacements = @(
                                        @{
                                            searchRegex = [regex]::Escape( $attachmentCfg['object'].preferredName )
                                            replacement = "[$( $attachmentCfg['markdownFileName'] )]($( $pageCfg['mediaPathPandoc'] )/$( $attachmentCfg['markdownFileName'] ))"
                                        }
                                    )
                                }
                            }
                            @{
                                description = 'Replace media (e.g. images, attachments) absolute paths with relative paths'
                                replacements = @(
                                    @{
                                        # E.g. 'C:/temp/notes/mynotebook/media/somepage-image1-timestamp.jpg' -> '../media/somepage-image1-timestamp.jpg'
                                        searchRegex = [regex]::Escape("$( $pageCfg['mediaParentPathPandoc'] )/") # Add a trailing front slash
                                        replacement = $pageCfg['levelsPrefix']
                                    }
                                )
                            }
                            @{
                                description = 'Add heading'
                                replacements = @(
                                    @{
                                        searchRegex = '^\s*'
                                        replacement = & {
                                            $heading = "# $( $pageCfg['object'].name )"
                                            if ($config['headerTimestampEnabled']['value'] -eq 1) {
                                                $heading += "`n`nCreated: $(  $pageCfg['dateTime'].ToString('yyyy-MM-dd HH:mm:ss zz00') )"
                                                $heading += "`n`nModified: $(  $pageCfg['lastModifiedTime'].ToString('yyyy-MM-dd HH:mm:ss zz00') )"
                                                $heading += "`n`n---`n`n"
                                            }
                                            $heading
                                        }
                                    }
                                )
                            }
                            if ($config['keepspaces']['value'] -eq 1) {
                                @{
                                    description = 'Clear extra newlines between unordered (bullet) and ordered (numbered) list items, non-breaking spaces from blank lines, and `>` after unordered lists'
                                    replacements = @(
                                        # Remove non-breaking spaces
                                        @{
                                            searchRegex = [regex]::Escape([char]0x00A0)
                                            replacement = ''
                                        }
                                        # Remove an extra newline between each occurrence of '- some unordered list item'
                                        @{
                                            searchRegex = '(\s*)- ([^\r\n]*)\r*\n\r*\n(?=\s*-)'
                                            replacement = "`$1- `$2`n"
                                        }
                                        # Remove an extra newline between each occurrence of '1. some ordered list item'
                                        @{
                                            searchRegex = '(\s*)(\d+\.) ([^\r\n]*)\r*\n\r*\n(?=\s*\d+\.)'
                                            replacement = "`$1`$2 `$3`n"
                                        }
                                        # Remove all '>' occurrences immediately following unordered lists
                                        @{
                                            searchRegex = '\n>[ ]*'
                                            replacement = "`n"
                                        }
                                    )
                                }
                            }
                            if ($config['keepescape']['value'] -eq 1) {
                                @{
                                    description = "Clear all '\' characters"
                                    replacements = @(
                                        @{
                                            searchRegex = [regex]::Escape('\')
                                            replacement = ''
                                        }
                                    )
                                }
                            }
                            elseif ($config['keepescape']['value'] -eq 2) {
                                @{
                                    description = "Clear all '\' characters except those preceding alphanumeric characters"
                                    replacements = @(
                                        @{
                                            searchRegex = '\\([^A-Za-z0-9])'
                                            replacement = '$1'
                                        }
                                    )
                                }
                            }
                            & {
                                if ($config['newlineCharacter']['value'] -eq 1) {
                                    @{
                                        description = "Use LF for newlines"
                                        replacements = @(
                                            @{
                                                searchRegex = '\r*\n'
                                                replacement = "`n"
                                            }
                                        )
                                    }
                                }else {
                                    @{
                                        description = "Use CRLF for newlines"
                                        replacements = @(
                                            @{
                                                searchRegex = '\r*\n'
                                                replacement = "`r`n"
                                            }
                                        )
                                    }
                                }
                            }
                        )
                        $pageCfg['directoriesToCreate'] = @(
                            # The directories to be created. These directories should never hit the absolute path, file name, or directory name limits on Windows
                            @(
                                $cfg['notesDocxDirectory']
                                $cfg['notesDirectory']
                                $pageCfg['tmpPath']
                                $pageCfg['fileDirectory']
                                $pageCfg['mediaPath']
                            ) | Select-Object -Unique
                        )
                        $pageCfg['directoriesToDelete'] = @(
                            $pageCfg['tmpPath']
                        )
                        $pageCfg['directorySeparatorChar'] = [io.path]::DirectorySeparatorChar

                        # Populate the pages array (needed even when -AsArray switch is not on, because we need this section's pages' state to know whether there are duplicate page names)
                        $sectionCfg['pages'].Add( $pageCfg ) > $null

                        if (!$AsArray) {
                            # Send the configuration immediately down the pipeline
                            $pageCfg
                        }
                    }
                }else {
                    "Ignoring empty Section: $( $sectionCfg['pathFromRoot'] )" | Write-Host -ForegroundColor DarkGray
                }

                # Populate the sections array
                if ($AsArray) {
                    $cfg['sections'].Add( $sectionCfg ) > $null
                }
            }
        }

        $cfg['sectionGroups'] = [System.Collections.ArrayList]@()

        # Build this Section Group's Section Groups
        if ((Get-Member -InputObject $sectionGroup -Name 'SectionGroup')) {
            if ($AsArray) {
                $cfg['sectionGroups'] = New-SectionGroupConversionConfig -OneNoteConnection $OneNoteConnection -NotesDestination $cfg['notesDirectory'] -Config $Config -SectionGroups $sectionGroup.SectionGroup -LevelsFromRoot ($LevelsFromRoot + 1) -AsArray:$AsArray
            }else {
                # Send the configuration immediately down the pipeline
                New-SectionGroupConversionConfig -OneNoteConnection $OneNoteConnection -NotesDestination $cfg['notesDirectory'] -Config $Config -SectionGroups $sectionGroup.SectionGroup -LevelsFromRoot ($LevelsFromRoot + 1)
            }
        }

        # Populate the conversion config
        if ($AsArray) {
            $sectionGroupConversionConfig.Add( $cfg ) > $null
        }
    }

    # Return the final conversion config
    if ($AsArray) {
        ,$sectionGroupConversionConfig # This syntax is needed to send an array down the pipeline without it being unwrapped. (It works by wrapping it in an array with a null sibling)
    }
}

Function Convert-OneNotePage {
    [CmdletBinding(DefaultParameterSetName='default')]
    [OutputType([void])]
    param (
        # Onenote connection object
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # ConvertOneNote2MarkDown configuration object
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Conversion object
        [Parameter(Mandatory,ParameterSetName='default')]
        [ValidateNotNullOrEmpty()]
        [object]
        $ConversionConfig
    ,
        [Parameter(Mandatory,ParameterSetName='pipeline',ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [object]
        $InputObject
    )

    process {
        if ($InputObject) {
            $ConversionConfig = $InputObject
        }
        if ($null -eq $ConversionConfig) {
            throw "No config specified."
        }

        try {
            $pageCfg = $ConversionConfig

            "$( '#' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel']) ) $( $pageCfg['object'].name ) [$( $pageCfg['kind'] )]" | Write-Host
            "Uri: $( $pageCfg['uri'] )" | Write-Verbose

            # Create directories
            foreach ($d in $pageCfg['directoriesToCreate']) {
                try {
                    "Directory: $( $d )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        $item = New-Item -Path $d -ItemType Directory -Force -ErrorAction Stop
                    }
                }catch {
                    Write-Error "Failed to create directory $d" -ErrorAction Continue
                    throw
                }
            }

            if ($config['usedocx']['value'] -eq 1) {
                # Remove any existing docx files, don't proceed if it fails
                try {
                    "Removing existing docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        if (Test-Path -LiteralPath $pageCfg['docxExportFilePath']) {
                            Remove-Item -LiteralPath $pageCfg['docxExportFilePath'] -Force -ErrorAction Stop
                        }
                    }
                }catch {
                    Write-Error "Failed to remove existing docx file $( $pageCfg['docxExportFilePath'] )" -ErrorAction Continue
                    throw
                }
            }

            # Publish OneNote page to Word, don't proceed if it fails
            if (! (Test-Path -LiteralPath $pageCfg['docxExportFilePath']) ) {
                try {
                    "Publishing new docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        Publish-OneNotePage -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID -Destination $pageCfg['docxExportFilePath'] -PublishFormat 'pfWord'
                    }
                }catch {
                    Write-Error "Failed to publish page to docx file $( $pageCfg['docxExportFilePath'] )" -ErrorAction Continue
                    throw
                }
            }else {
                "Existing docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
            }

            # Publish OneNote page to pdf, don't proceed if it fails
            if ($config['exportPdf']['value'] -eq 2) {
                if (! (Test-Path -LiteralPath $pageCfg['pdfExportFilePath']) ) {
                    try {
                        "Publishing new pdf file: $( $pageCfg['pdfExportFilePath'] )" | Write-Verbose
                        if (!$config['dryRun']['value']) {
                            Publish-OneNotePage -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID -Destination $pageCfg['pdfExportFilePathTmp'] -PublishFormat 'pfPdf'
                            Move-Item $pageCfg['pdfExportFilePathTmp'] $pageCfg['pdfExportFilePath']
                        }
                        "pdf file ready: $( $pageCfg['pdfExportFilePath'] )" | Write-Host -ForegroundColor Green
                    }catch {
                        Write-Error "Failed to publish page to pdf file $( $pageCfg['pdfExportFilePath'] )" -ErrorAction Continue
                        throw
                    }
                }else {
                    "Existing pdf file: $( $pageCfg['pdfExportFilePath'] )" | Write-Host -ForegroundColor Green
                }
            }

            # https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            # Convert .docx to .md, don't proceed if it fails
            $stderrFile = "$( $pageCfg['tmpPath'] )/pandoc-stderr.txt"
            try {
                # Start-Process has no way of capturing stderr / stdterr to variables, so we need to use temp files.
                "Converting docx file to markdown file: $( $pageCfg['filePath'] )" | Write-Verbose
                if (!$config['dryRun']['value']) {
                    $argumentList = @(
                        '-f'
                        'docx'
                        '-t'
                        $pageCfg['conversion']
                        '-i'
                        if ($pageCfg['docxExportFilePath'] -match ' ') {
                            "`"$( $pageCfg['docxExportFilePath'] )`"" # Add double-quotes to path containing spaces
                        }else {
                            $pageCfg['docxExportFilePath']
                        }
                        '-o'
                        if ($pageCfg['filePathNormal'] -match ' ') {
                            "`"$( $pageCfg['filePathNormal'] )`"" # Add double-quotes to path containing spaces
                        }else {
                            $pageCfg['filePathNormal']
                        }
                        '--wrap=none'
                        '--markdown-headings=atx'
                        if ($pageCfg['mediaParentPathPandoc'] -match ' ') {
                            "`"--extract-media=$( $pageCfg['mediaParentPathPandoc'] )`"" # Add double-quotes to path containing spaces
                        }else {
                            "--extract-media=$( $pageCfg['mediaParentPathPandoc'] )"
                        }
                    )
                    "Command line: pandoc.exe $argumentList" | Write-Verbose
                    $process = Start-Process -ErrorAction Stop -RedirectStandardError $stderrFile -PassThru -NoNewWindow -Wait -FilePath pandoc.exe -ArgumentList $argumentList # extracts into ./media of the supplied folder
                    if ($process.ExitCode -ne 0) {
                        $stderr = Get-Content $stderrFile -Raw
                        throw "pandoc error: $stderr"
                    }
                }
            }catch {
                Write-Error "Failed to convert docx file $( $pageCfg['docxExportFilePath'] ) to markdown file $( $pageCfg['filePathNormal'] )"
                throw
            }finally {
                if (Test-Path $stderrFile) {
                    Remove-Item $stderrFile -Force
                }
            }

            # Cleanup Word files
            if ($config['keepdocx']['value'] -eq 1) {
                try {
                    "Removing existing docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        if (Test-Path -LiteralPath $pageCfg['docxExportFilePath']) {
                            Remove-Item -LiteralPath $pageCfg['docxExportFilePath'] -Force -ErrorAction Stop
                        }
                    }
                }catch {
                    Write-Error "Failed to remove existing docx file $( $pageCfg['docxExportFilePath'] ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                }
            }

            # Save any attachments
            foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                try {
                    "Saving inserted attachment: $( $attachmentCfg['destination'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        Copy-Item -Path $attachmentCfg['source'] -Destination $attachmentCfg['destination'] -Force -ErrorAction Stop
                    }
                }catch {
                    Write-Error "Failed to save attachment from $( $attachmentCfg['source'] ) to $( $attachmentCfg['destination'] ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                }
            }

            # Rename images to have unique names - NoteName-Image#-HHmmssff.xyz
            if (!$config['dryRun']['value']) {
                $images = Get-ChildItem -Path $pageCfg['mediaPathPandoc'] -Recurse -Force -ErrorAction SilentlyContinue
                foreach ($image in $images) {
                    # Rename Image
                    try {
                        $newimageName = if ($config['medialocation']['value'] -eq 2) {
                            "$( $pageCfg['filePathRelUnderscore'] )-$($image.BaseName)$($image.Extension)"
                        }else {
                            "$( $pageCfg['pathFromRootCompat'] )-$($image.BaseName)$($image.Extension)"
                        }
                        $newimagePath = [io.path]::combine( $pageCfg['mediaPath'], $newimageName )
                        "Moving image: $( $image.FullName ) to $( $newimagePath )" | Write-Verbose
                        if (!$config['dryRun']['value']) {
                            $item = Move-Item -Path "$( $image.FullName )" -Destination $newimagePath -Force -ErrorAction Stop -PassThru
                        }
                    }catch {
                        Write-Error "Failed to rename image $( $image.FullName ) to $( $item.FullName ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                    }
                    # Mutate markdown content with new image references
                    try {
                        "Mutation of markdown: Rename image references. Find: '$( $image.Name )', Replacement: '$( $newimageName )'" | Write-Verbose
                        if (!$config['dryRun']['value']) {
                            $content = Get-Content -LiteralPath $pageCfg['filePath'] -Raw -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                            $content = $content.Replace("$($image.Name)", "$($newimageName)")
                            Set-ContentNoBom -LiteralPath $pageCfg['filePath'] -Value $content -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                        }
                    }catch {
                        Write-Error "Failed to rename image references to $( $newimageName ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                    }
                }
            }

            # Mutate markdown content
            try {
                if (!$config['dryRun']['value']) {
                    # Get markdown content
                    $content = @( Get-Content -LiteralPath $pageCfg['filePath'] -ErrorAction Stop ) # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                    $content = @(
                        if ($content.Count -gt 6) {
                            # Discard first 6 lines which contain a header, created date, and time. We are going to add our own header
                            $content[6..($content.Count - 1)]
                        }else {
                            # Empty page
                            ''
                        }
                    ) -join "`n"
                }

                # Mutate
                foreach ($m in ($pageCfg['mutations'] | Where-Object { $_ -and $_.ContainsKey('replacements') })) {
                    foreach ($r in $m['replacements']) {
                        try {
                            "Mutation of markdown: $( $m['description'] ). Regex: '$( $r['searchRegex'] )', Replacement: '$( $r['replacement'].Replace("`r", '\r').Replace("`n", '\n') )'" | Write-Verbose
                            if (!$config['dryRun']['value']) {
                                $content = $content -replace $r['searchRegex'], $r['replacement']
                            }
                        }catch {
                            Write-Error "Failed to mutate markdown content with mutation '$( $m['description'] )'. Exception: $( $_.Exception.Message )"
                        }
                    }
                }
                if (!$config['dryRun']['value']) {
                    Set-ContentNoBom -LiteralPath $pageCfg['filePath'] -Value $content -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                }
            }catch {
                Write-Error "Failed to mutate markdown content: $( $_.Exception.Message )"
            }

            "Markdown file ready: $( $pageCfg['filePathNormal'] )" | Write-Host -ForegroundColor Green
        }catch {
            # Check for specific OneNote publish errors
            $hresult = if ($_.Exception.InnerException) { $_.Exception.InnerException.HResult } else { 0 }
            $reason = if ($hresult -eq 0x80042006) { " (page may be empty or contain unsupported content)" } else { "" }
            "  [SKIP] $( $pageCfg['object'].name ) -- publish failed$reason" | Write-Host -ForegroundColor Yellow
            Write-Error "Failed to convert page: $( $pageCfg['pathFromRoot'] )" -ErrorAction Continue
            if ($ErrorActionPreference -eq 'Stop') {
                throw
            }
        }
    }
}

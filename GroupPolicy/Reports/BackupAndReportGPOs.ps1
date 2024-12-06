<#
.SYNOPSIS
    Script to back up Group Policy Objects (GPOs), generate reports, and export GPO link information.

.DESCRIPTION
    This script performs the following actions:
    - Backs up all GPOs to a specified location.
    - Generates HTML and XML reports for each GPO.
    - Extracts and exports GPO link and WMI filter information to CSV files.
    - Cleans up old backups based on a user-specified retention period.
    - Retrieves GPO link order for OUs.

.PARAMETER DaysToKeep
    Specifies the number of days to keep backup files before deleting them. Default is 7 days.

.PARAMETER BackupLocation
    Specifies the directory where GPO backup files, reports, and CSVs will be stored.
    The directory must already exist.

.EXAMPLE
    .\BackupAndReportGPOs.ps1 -DaysToKeep 30 -BackupLocation "C:\Backup"

.OUTPUTS
    - Creates a backup directory structure with:
      * Backups of all GPOs
      * Compressed HTML/XML reports for each GPO
      * CSV files listing GPO links and link order
    - Removes backups older than the specified retention period.

.NOTES
    Author: Micah
    Requires the GroupPolicy and ActiveDirectory modules.

.INPUTS
    DaysToKeep, BackupLocation

.LINK
    https://docs.microsoft.com/en-us/powershell/module/grouppolicy
    https://docs.microsoft.com/en-us/powershell/module/activedirectory
#>

[CmdletBinding()]
param
(
    [ValidateRange(1, 365)]
    [int]$DaysToKeep = 7,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (Test-Path $_ -PathType Container) { $true }
        else { throw "The path '$_' does not exist or is not a directory." }
    })]
    [string]$BackupLocation
)

#region Module Checks
# Ensure required modules are available
if (-not (Get-Module -Name 'GroupPolicy' -ListAvailable)) {
    Write-Error "The GroupPolicy module is not available on this system."
    exit 1
}
Import-Module GroupPolicy -ErrorAction Stop

if (-not (Get-Module -Name 'ActiveDirectory' -ListAvailable)) {
    Write-Error "The ActiveDirectory module is not available on this system."
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop
#endregion Module Checks

#region Functions

function Backup-GroupPolicy {
    <#
    .SYNOPSIS
        Backs up Group Policy Objects (GPOs), generates reports, and exports GPO link information.

    .DESCRIPTION
        This function performs a comprehensive backup of GPOs, creates HTML/XML reports for each GPO,
        and exports information about GPO links into CSV files. It then organizes and compresses the
        backup/report files and removes old backups based on the specified retention period.

    .PARAMETER BackupLocation
        The directory where backup files, reports, and CSVs will be stored.

    .PARAMETER DaysToKeep
        The number of days to retain old backup files.

    .EXAMPLE
        Backup-GroupPolicy -BackupLocation "C:\Backup" -DaysToKeep 30
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupLocation,

        [Parameter(Mandatory = $true)]
        [int]$DaysToKeep
    )

    try {
        Write-Verbose "Initializing backup process..."

        # Normalize path and initialize variables
        $BackupLocation = $BackupLocation.TrimEnd('\')
        $startTime = (Get-Date -Format "yyyy-MM-dd-HHmm")
        $todayBackupPath = Join-Path $BackupLocation ("GPO Backups from $startTime")

        $paths = @{
            'Temp'         = Join-Path $BackupLocation "GPOBackupsTemp"
            'TempBackups'  = Join-Path $BackupLocation "GPOBackupsTemp\Backup"
            'TempReports'  = Join-Path $BackupLocation "GPOBackupsTemp\Reports"
            'TodaysBackup' = $todayBackupPath
        }

        # Create required directories
        foreach ($dir in $paths.Values) {
            if (-not (Test-Path $dir)) {
                New-Item -ItemType Directory -Path $dir | Out-Null
                Write-Verbose "Created directory: $dir"
            }
        }

        # Clear Temp folder
        Write-Verbose "Clearing temp folders..."
        Get-ChildItem -Path $paths['Temp'] -Force -Recurse | Remove-Item -Force -Recurse -Confirm:$false

        # Get PDC Emulator DC
        Write-Verbose "Fetching PDC Emulator domain controller..."
        $pdcEmulator = (Get-ADDomainController -Filter { OperationMasterRoles -Like "PDCEmulator" }).HostName

        # Retrieve all GPOs
        Write-Verbose "Retrieving all GPOs from PDC Emulator '$pdcEmulator'..."
        $gpos = Get-GPO -All -Server $pdcEmulator

        $allGpoLinks = New-Object System.Collections.ArrayList

        foreach ($gpo in $gpos) {
            $gpoDisplayName = $gpo.DisplayName
            $gpoGuid = $gpo.Id.Guid
            $gpoBackupComment = "Backup of $gpoDisplayName on $startTime"

            Write-Verbose "Backing up GPO: $gpoDisplayName"
            Backup-GPO -Guid $gpoGuid -Path $paths['TempBackups'] -Comment $gpoBackupComment | Out-Null

            # Prepare file-safe name for reports
            $safeName = ($gpoDisplayName -replace '[\x2B\x2F\x22\x3A\x3C\x3E\x3F\x5C\x7C]', ' ')

            # Generate GPO reports
            $reportPathHtml = Join-Path $paths['TempReports'] ($safeName + ".html")
            Write-Verbose "Generating HTML report for $gpoDisplayName"
            Get-GPOReport -Name $gpoDisplayName -Path $reportPathHtml -ReportType HTML

            $reportPathXml = Join-Path $paths['TempReports'] ($safeName + ".xml")
            Write-Verbose "Generating XML report for $gpoDisplayName"
            Get-GPOReport -Name $gpoDisplayName -Path $reportPathXml -ReportType XML

            # Extract GPO link info from XML
            Write-Verbose "Extracting GPO link information..."
            $xmlContent = [xml](Get-Content $reportPathXml)
            $links = $xmlContent.GPO.LinksTo
            if ($links -and $links.SOMName) {
                # Convert links to objects
                foreach ($link in $links.SOMName) {
                    $linkObj = [PSCustomObject]@{
                        Policy     = $gpoDisplayName
                        OU         = $link
                        Path       = ($links.SOMPath | Where-Object {$_})[0]   # handle multiple entries if any
                        LinkStatus = ($links.Enabled | Where-Object {$_})[0]
                        Enforced   = ($links.NoOverride | Where-Object {$_})[0]
                    }
                    $allGpoLinks.Add($linkObj) | Out-Null
                }
            } else {
                # No links found, create a default object
                $allGpoLinks.Add([PSCustomObject]@{
                    Policy     = $gpoDisplayName
                    OU         = "NA"
                    Path       = "NA"
                    LinkStatus = "NA"
                    Enforced   = "NA"
                }) | Out-Null
            }
        }

        # Export GPO link info
        $linksCsv = Join-Path $paths['TempReports'] ("GPOLinks-$startTime.csv")
        Write-Verbose "Exporting GPO link data to $linksCsv"
        $allGpoLinks | Export-Csv -Delimiter "`t" -LiteralPath $linksCsv -NoTypeInformation

        # Export GPO Link Order
        Write-Verbose "Retrieving and exporting GPO link order..."
        $linkOrderCsv = Join-Path $paths['TempReports'] ("GPOLinkOrder-$startTime.csv")
        Get-GPOLinkOrder | Export-Csv -Delimiter "`t" -LiteralPath $linkOrderCsv -NoTypeInformation

        # Move backups to today's backup folder
        Write-Verbose "Organizing today's backup..."
        Move-Item -Path $paths['TempBackups'] -Destination $paths['TodaysBackup'] -Force

        # Compress reports
        $zipPath = Join-Path $paths['TodaysBackup'] ("GPOReports-$startTime.zip")
        Write-Verbose "Compressing reports to $zipPath"
        Compress-Archive -Path (Join-Path $paths['TempReports'] "*") -DestinationPath $zipPath -Force

        # Cleanup old backups
        Write-Verbose "Removing backups older than $DaysToKeep days..."
        Get-ChildItem -Path $BackupLocation | 
            Where-Object { $_.Name -like "GPO Backups from *" -and $_.CreationTime -lt (Get-Date).AddDays(-$DaysToKeep) } |
            Remove-Item -Recurse -Force -Confirm:$False

        Write-Host "GPO Backup and report generation completed successfully."
    }
    catch {
        Write-Error "An error occurred during GPO backup: $($_.Exception.Message)"
        exit 1
    }
}

function Get-GPOLinkOrder {
    <#
    .SYNOPSIS
        Gets the Group Policy Object (GPO) link order for Organizational Units (OUs).

    .DESCRIPTION
        This function retrieves OUs with associated GPO links from Active Directory and
        returns objects detailing the order and status of linked GPOs.

    .EXAMPLE
        Get-GPOLinkOrder

    .OUTPUTS
        PSCustomObjects with OU and GPO link order details.

    .NOTES
        Author: Micah
    #>
    [CmdletBinding()]
    param()

    try {
        # Retrieve OUs with gPlink property
        $organizationalUnits = Get-ADOrganizationalUnit -Filter * -Properties gPlink | Where-Object { $_.gPlink }

        foreach ($ou in $organizationalUnits) {
            $ouName = $ou.Name
            $ouDistinguishedName = $ou.DistinguishedName
            $gpoLinks = $ou.gPlink -split ']\[' | Where-Object { $_ }

            # The link order is reverse of the array index since GPOs are applied from bottom to top
            $gpoOrder = $gpoLinks.Count
            foreach ($gpoLink in $gpoLinks) {
                $parts = $gpoLink -split ';'
                $gpoDN = $parts[0].TrimEnd(']')
                $gpoStatus = $parts[1].TrimEnd(']')

                # Resolve GPO name from DN
                $gpoName = ([ADSI]$gpoDN).displayName
                $isLinked = $false
                $isEnforced = $false

                switch ($gpoStatus) {
                    "1" { # GPO Unlinked
                        $isLinked = $false
                        $isEnforced = $false
                    }
                    "2" { # GPO Linked & Enforced
                        $isLinked = $true
                        $isEnforced = $true
                    }
                    "3" { # GPO Unlinked & Enforced
                        $isLinked = $false
                        $isEnforced = $true
                    }
                    "0" { # GPO Linked, not enforced
                        $isLinked = $true
                        $isEnforced = $false
                    }
                }

                [PSCustomObject]@{
                    OUName              = $ouName
                    OUDistinguishedName = $ouDistinguishedName
                    GPOName             = $gpoName
                    IsLinked            = $isLinked
                    IsEnforced          = $isEnforced
                    GPOrder             = $gpoOrder
                }

                $gpoOrder--
            }
        }
    }
    catch {
        Write-Error "An error occurred retrieving GPO link order: $($_.Exception.Message)"
    }
}

#endregion Functions

#region Main Execution
Backup-GroupPolicy -BackupLocation $BackupLocation -DaysToKeep $DaysToKeep
#endregion Main Execution
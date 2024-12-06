<#
.SYNOPSIS
    Retrieves status information for all Group Policy Objects (GPOs).

.DESCRIPTION
    This script enumerates all GPOs in the environment and determines whether each one is:
    - Disabled (if all settings in the GPO are disabled)
    - Unlinked (if the GPO is not linked to any container)
    If a GPO does not meet these criteria, it is skipped.

.EXAMPLE
    PS> .\Get-GpoStatusInfo.ps1
    Retrieves all GPOs, determines their status, and displays a formatted table of results.

.OUTPUTS
    An array of PSCustomObjects containing:
    - Name:   The display name of the GPO
    - Status: The determined status of the GPO ("Disabled" or "Unlinked")

.NOTES
    Author: Micah
    This script requires the Group Policy Module (e.g., via RSAT or Active Directory management tools).

.PARAMETER None
    This script takes no parameters.

.INPUTS
    None
#>

[CmdletBinding()]
param()

# Ensure the GroupPolicy module is available.
try {
    Import-Module GroupPolicy -ErrorAction Stop
} catch {
    Write-Error "Failed to import the GroupPolicy module. Ensure it is installed."
    exit 1
}

$AllGPOs = Get-GPO -All -ErrorAction Stop
$Output = New-Object System.Collections.Generic.List[System.Object]

foreach ($GPO in $AllGPOs) {
    # Determine if the GPO is disabled or unlinked.
    # If it's disabled, GpoStatus will be 'AllSettingsDisabled'.
    # If not disabled, we need to check if it's unlinked by examining the XML report.
    $status = $null

    if ($GPO.GpoStatus -eq 'AllSettingsDisabled') {
        $status = 'Disabled'
    } else {
        try {
            [xml]$report = Get-GPOReport -Guid $GPO.Id -ReportType Xml -ErrorAction Stop
            # If 'LinksTo' is not present, the GPO is considered unlinked
            if ($report.GPO.PSObject.Properties.Name -notcontains 'LinksTo') {
                $status = 'Unlinked'
            }
        } catch {
            Write-Warning "Failed to retrieve report for GPO '$($GPO.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }

    if ($status) {
        # Construct a PSCustomObject with the GPO's name and status
        $GPOObject = [PSCustomObject]@{
            Name   = $GPO.DisplayName
            Status = $status
        }
        $Output.Add($GPOObject)
    }
}

# Write the count of GPOs that have a determined status
Write-Output $Output.Count

# Sort and display the output in a table
$Output | Sort-Object Name | Format-Table -AutoSize
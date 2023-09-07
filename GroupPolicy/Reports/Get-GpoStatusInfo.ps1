<#
.SYNOPSIS
    This script retrieves the status information of all Group Policy Objects (GPOs).
.DESCRIPTION
    The script collects all GPOs and determines their status as 'Disabled', 'Unlinked', or skips them.
.PARAMETER None
    This script takes no parameters.
.EXAMPLE
    PS> .\Get-GpoStatusInfo.ps1
    This will execute the script and display the GPOs with their status.
.NOTES
    Author: Micah
.INPUTS
    None
.OUTPUTS
    An array of PSCustomObjects containing GPO names and their statuses.
#>

$AllGPOs = Get-Gpo -All
$Output = @()

foreach ($GPO in $AllGPOs) {
    # Initialize variable to hold the GPO status
    $Status = $null

    if ($GPO.GpoStatus -eq 'AllSettingsDisabled') {
        # Set status to 'Disabled' if all settings are disabled
        $Status = 'Disabled'
    }
    else {
        # Generate an XML report for the GPO
        [xml]$GPOReport = Get-GPOReport -Guid $GPO.Id -ReportType 'Xml'
        
        # Check if the GPO is unlinked
        if ($GPOReport.GPO.PSObject.Properties.Name -notcontains 'LinksTo') {
            $Status = 'Unlinked'
        }
    }

    if ($Status) {
        # Create a custom object for GPO status and name
        $GPOOutput = [PSCustomObject]@{
            Status = $Status
            Name   = $GPO.DisplayName
        }
        
        # Add the custom object to the output array
        $Output += $GPOOutput
    }
}

# Display the count of GPOs with status
Write-Output $Output.Count

# Sort and format the output table
$Output | Sort-Object Name | Format-Table -AutoSize

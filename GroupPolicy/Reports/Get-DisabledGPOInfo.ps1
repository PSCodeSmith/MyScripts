<#
.SYNOPSIS
    Retrieves information about disabled Group Policy Objects (GPOs).

.DESCRIPTION
    This script queries all GPOs in the environment and identifies those with disabled settings.
    It returns an array of custom objects containing the GPO name and the category of settings that are disabled.

.EXAMPLE
    PS> .\Get-DisabledGpoInfo.ps1
    Returns disabled GPOs with their respective disabled categories.

.OUTPUTS
    PSCustomObject
    Each object includes:
    - Name: The display name of the GPO.
    - DisabledSettingsCategory: The category of GPO settings that are disabled (e.g., user, computer, or both).

.NOTES
    Author: Micah

.PARAMETER None
    No parameters required.

.INPUTS
    None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/grouppolicy
#>

[CmdletBinding()]
[OutputType([PSCustomObject])]
param ()

#region Initialization
$ErrorActionPreference = 'Stop'
#endregion Initialization

#region Main Logic
try {
    # Get all GPOs and filter those that are disabled
    $disabledGpos = Get-Gpo -All | Where-Object { $_.GpoStatus -like '*Disabled' }

    # Construct the output objects
    $results = foreach ($gpo in $disabledGpos) {
        [PSCustomObject]@{
            Name                     = $gpo.DisplayName
            DisabledSettingsCategory = $gpo.GpoStatus -replace 'Disabled'
        }
    }

    # Output the results
    $results
} catch {
    # If any error occurs, throw a terminating error
    Write-Error "An error occurred while retrieving disabled GPO information: $($_.Exception.Message)"
    $PSCmdlet.ThrowTerminatingError($_)
}
#endregion Main Logic
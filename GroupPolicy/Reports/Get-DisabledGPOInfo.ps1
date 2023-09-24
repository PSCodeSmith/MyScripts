<#
	.SYNOPSIS
		Retrieves information about disabled Group Policy Objects (GPOs).
	
	.DESCRIPTION
		This script fetches all GPOs, filters out the disabled ones, and returns their names along with the specific disabled settings.
	
	.EXAMPLE
		PS> .\Get-DisabledGpoInfo.ps1
		This will execute the script and return disabled GPOs with their respective categories.
	
	.OUTPUTS
		An array of PSCustomObjects containing the names of disabled GPOs and their disabled setting categories.
	
	.NOTES
		Author: Micah
	
	.PARAMETER None
		This script takes no parameters.
	
	.INPUTS
		None
#>
[CmdletBinding()]
[OutputType([PSCustomObject])]
param ()

Begin {
    # Set the error action preference to stop execution on errors
    $ErrorActionPreference = 'Stop'
}

Process {
    try {
        # Get all GPOs and filter those with disabled settings
        Get-Gpo -All | Where-Object { $_.GpoStatus -like '*Disabled' } | ForEach-Object {
            # Create a custom object for each GPO with disabled settings
            [PSCustomObject]@{
                Name                     = $_.DisplayName
                DisabledSettingsCategory = $_.GpoStatus -replace 'Disabled'
            }
        }
    } catch {
        # Terminate the script if an error occurs and display the error
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

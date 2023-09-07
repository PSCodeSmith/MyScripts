<#
.SYNOPSIS
    Identifies Group Policy Objects (GPOs) with unpopulated User or Computer links.
.DESCRIPTION
    The script scans all GPOs, filters out default GPOs, and identifies those with unpopulated User or Computer settings.
.PARAMETER Remediate
    When set to $true, automatically disables the unpopulated settings in the identified GPOs.
.EXAMPLE
    PS> Get-UnpopulatedGpoLinks -Remediate $false
    This will identify unpopulated GPO links without making any changes.
.NOTES
    Author: Micah
.INPUTS
    Remediate (Boolean)
.OUTPUTS
    PSCustomObject containing GPO names and their unpopulated links.
#>
[CmdletBinding()]
param (
    [Parameter()]
    [bool]$Remediate = $false
)

function Test-ObjectMember {
    <#
    .DESCRIPTION
        Tests if a specific member exists in a given object.
    .PARAMETER Object
        The object to test.
    .PARAMETER MemberName
        The name of the member to look for.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Object]$Object,
        [Parameter(Mandatory = $true)]
        [String]$MemberName
    )

    return $null -ne ($Object | Get-Member -Name $MemberName)
}

# Check if the Group Policy module is loaded
if (!(Get-Module -Name 'GroupPolicy')) {
    Write-Error 'One or more required modules not loaded'
    return
}

# Define an array of default Active Directory GPOs
$DefaultPolicyNames = @('Default Domain Controllers Policy', 'Default Domain Policy')

# Get all GPO reports in XML format
$AllGpoReports = Get-GPOReport -ReportType 'XML' -All

foreach ($SingleGpoReport in $AllGpoReports) {
    $ParsedGpo = ([xml]$SingleGpoReport).GPO

    # Skip default GPOs
    if ($DefaultPolicyNames -notcontains $ParsedGpo.Name) {
        $GpoOutput = [PSCustomObject]@{
            'GPOName' = $ParsedGpo.Name
        }

        # Check for unpopulated User link
        if ($ParsedGpo.User.Enabled -eq 'true' -and !(Test-ObjectMember $ParsedGpo.User 'ExtensionData')) {
            $GpoOutput | Add-Member -Type NoteProperty -Name 'UnpopulatedLink' -Value 'User' -Force
            if ($Remediate) {
                (Get-GPO -Name $ParsedGpo.Name).GPOStatus = 'UserSettingsDisabled'
                Write-Output "Disabled user settings on GPO $($ParsedGpo.Name)"
            } else {
                $GpoOutput
            }
        }

        # Check for unpopulated Computer link
        if ($ParsedGpo.Computer.Enabled -eq 'true' -and !(Test-ObjectMember $ParsedGpo.Computer 'ExtensionData')) {
            $GpoOutput | Add-Member -Type NoteProperty -Name 'UnpopulatedLink' -Value 'Computer' -Force
            if ($Remediate) {
                (Get-GPO -Name $ParsedGpo.Name).GPOStatus = 'ComputerSettingsDisabled'
                Write-Output "Disabled computer settings on GPO $($ParsedGpo.Name)"
            } else {
                $GpoOutput
            }
        }
    }
}

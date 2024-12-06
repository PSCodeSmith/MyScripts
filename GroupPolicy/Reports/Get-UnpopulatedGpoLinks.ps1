<#
.SYNOPSIS
    Identifies Group Policy Objects (GPOs) with unpopulated User or Computer links.

.DESCRIPTION
    This function retrieves all GPO reports in XML format, excluding the default GPOs
    ("Default Domain Controllers Policy" and "Default Domain Policy"). It then checks
    if the User or Computer configuration sections are enabled but have no linked settings
    (unpopulated links). If any such unpopulated links are found, it optionally disables
    them if the -Remediate parameter is set.

.PARAMETER Remediate
    When $true, disables the unpopulated settings in the identified GPOs. When $false,
    it only reports them without making changes.

.EXAMPLE
    PS> Get-UnpopulatedGpoLinks -Remediate $false
    Identifies GPOs with unpopulated links without making any changes.

.EXAMPLE
    PS> Get-UnpopulatedGpoLinks -Remediate $true
    Identifies GPOs with unpopulated links and disables those unpopulated settings.

.OUTPUTS
    PSCustomObject containing GPO names and their unpopulated links, unless -Remediate is used.
    If -Remediate is used, it outputs a string message confirming the remediation action.

.NOTES
    Author: Micah
    The script requires the GroupPolicy module.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param
(
    [Parameter(Mandatory = $false)]
    [bool]$Remediate = $false
)

function Test-ObjectMember {
    <#
    .SYNOPSIS
        Checks if a given object has a specified member.

    .DESCRIPTION
        Returns $true if the object has a specified member, $false otherwise.

    .PARAMETER Object
        The object to check.

    .PARAMETER MemberName
        The name of the member to look for.

    .EXAMPLE
        Test-ObjectMember -Object $ParsedGpo.User -MemberName 'ExtensionData'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Object]$Object,

        [Parameter(Mandatory = $true)]
        [String]$MemberName
    )

    return ($null -ne ($Object | Get-Member -Name $MemberName))
}

function Get-UnpopulatedGpoLinks {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory=$false)]
        [bool]$Remediate = $false
    )

    # Ensure the GroupPolicy module is loaded
    if (-not (Get-Module -Name 'GroupPolicy' -ListAvailable)) {
        Write-Error 'GroupPolicy module not found. Please ensure the GroupPolicy module is installed.'
        return
    }
    Import-Module GroupPolicy -ErrorAction Stop

    # Define default Active Directory GPOs to skip
    $DefaultPolicyNames = @('Default Domain Controllers Policy', 'Default Domain Policy')

    # Retrieve all GPO reports in XML format
    $AllGpoReports = Get-GPOReport -ReportType 'XML' -All

    foreach ($SingleGpoReport in $AllGpoReports) {
        $ParsedGpo = [xml]$SingleGpoReport
        $GpoNode = $ParsedGpo.GPO

        # Skip default GPOs
        if ($DefaultPolicyNames -contains $GpoNode.Name) {
            continue
        }

        # Prepare a base object to report on
        $GpoOutput = [PSCustomObject]@{
            GPOName = $GpoNode.Name
        }

        $foundUnpopulated = $false

        # Check for unpopulated User link
        if ($GpoNode.User.Enabled -eq 'true' -and -not (Test-ObjectMember -Object $GpoNode.User -MemberName 'ExtensionData')) {
            $foundUnpopulated = $true
            $GpoOutput | Add-Member -Type NoteProperty -Name 'UnpopulatedLink' -Value 'User' -Force

            if ($Remediate) {
                if ($PSCmdlet.ShouldProcess("GPO: $($GpoNode.Name)", "Disable User Settings")) {
                    (Get-GPO -Name $GpoNode.Name).GPOStatus = 'UserSettingsDisabled'
                    Write-Output "Disabled user settings on GPO '$($GpoNode.Name)'"
                }
            }
        }

        # Check for unpopulated Computer link
        if ($GpoNode.Computer.Enabled -eq 'true' -and -not (Test-ObjectMember -Object $GpoNode.Computer -MemberName 'ExtensionData')) {
            $foundUnpopulated = $true
            $GpoOutput | Add-Member -Type NoteProperty -Name 'UnpopulatedLink' -Value 'Computer' -Force

            if ($Remediate) {
                if ($PSCmdlet.ShouldProcess("GPO: $($GpoNode.Name)", "Disable Computer Settings")) {
                    (Get-GPO -Name $GpoNode.Name).GPOStatus = 'ComputerSettingsDisabled'
                    Write-Output "Disabled computer settings on GPO '$($GpoNode.Name)'"
                }
            }
        }

        # Output the GPO if we found unpopulated links and are not remediating
        # If we are remediating, we've already written output about our actions.
        if ($foundUnpopulated -and -not $Remediate) {
            Write-Output $GpoOutput
        }
    }
}

# Call the function with the provided Remediate parameter
Get-UnpopulatedGpoLinks -Remediate:$Remediate
<#
.SYNOPSIS
    Retrieves lockout events from the PDC emulator security logs for a specified domain.

.DESCRIPTION
    This script queries the PDC emulator of the given domain and retrieves Event ID 4740 (account lockout)
    entries from the Security event log. It filters events based on an optional UserName parameter and a 
    specified StartTime to limit the search window.

.PARAMETER DomainName
    The domain name to query. Defaults to the current user's domain.

.PARAMETER UserName
    The username for which to search lockout events. Wildcards are supported. Defaults to all locked out users.

.PARAMETER StartTime
    The start time from which to begin searching the event logs. Defaults to 3 days ago.

.PARAMETER Credential
    Credential to use when querying the remote PDC emulator's event logs. Defaults to the current user context.

.EXAMPLE
    .\Get-LockoutEvents.ps1 -DomainName "YourDomain" -UserName "jdoe"
    Retrieves lockout events for the user "jdoe" from the specified domain's PDC emulator.

.OUTPUTS
    Returns objects with TimeCreated, UserName, and ClientName properties.

.NOTES
    Author: Micah

.INPUTS
    DomainName, UserName, StartTime, Credential

.LINK
    https://docs.microsoft.com/en-us/windows/security/threat-protection/auditing/event-4740
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$DomainName = $env:USERDOMAIN,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$UserName = '*',

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [datetime]$StartTime = (Get-Date).AddDays(-3),

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
)

#region Helper Functions

function Get-PDCEmulator {
    <#
    .SYNOPSIS
        Retrieves the PDC emulator of the specified domain.

    .DESCRIPTION
        Uses the System.DirectoryServices.ActiveDirectory classes to find the PDC role owner
        for the given domain. Throws an error if it fails to retrieve the PDC.

    .PARAMETER DomainName
        The domain name for which to find the PDC emulator.

    .OUTPUTS
        [string]: PDC emulator hostname.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DomainName
    )

    try {
        $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName)
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
        return $domain.PdcRoleOwner.Name
    } catch {
        Write-Error "Unable to query the domain for PDC emulator. Verify access and try again. Error: $($_.Exception.Message)"
        throw
    }
}

function Get-LockoutEvents {
    <#
    .SYNOPSIS
        Retrieves Event ID 4740 (account lockout) events from the specified PDC emulator.

    .DESCRIPTION
        Connects to the PDC emulator, queries the Security event log for lockout events since the given StartTime,
        and filters by the specified UserName pattern.

    .PARAMETER PdcEmulator
        The PDC emulator hostname.

    .PARAMETER UserName
        The username wildcard pattern to filter lockout events. Defaults to '*'.

    .PARAMETER StartTime
        The start datetime from which events should be retrieved.

    .PARAMETER Credential
        Credential to use for remote query. If not provided, uses the current user context.

    .OUTPUTS
        Objects with TimeCreated, UserName, and ClientName properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$PdcEmulator,

        [Parameter(Mandatory=$true)]
        [string]$UserName,

        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,

        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$Credential
    )

    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('Credential') -and $Credential -ne [System.Management.Automation.PSCredential]::Empty) {
        $invokeParams.Credential = $Credential
    }

    try {
        Invoke-Command -ComputerName $PdcEmulator -ScriptBlock {
            param($UserName, $StartTime)
            Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4740; StartTime = $StartTime } |
                Where-Object { $_.Properties[0].Value -like $UserName } |
                Select-Object -Property TimeCreated,
                    @{ Label = 'UserName'; Expression = { $_.Properties[0].Value } },
                    @{ Label = 'ClientName'; Expression = { $_.Properties[1].Value } }
        } -ArgumentList $UserName, $StartTime @invokeParams
    } catch {
        Write-Error "Failed to retrieve lockout events from $PdcEmulator. Error: $($_.Exception.Message)"
        throw
    }
}
#endregion Helper Functions

#region Main Script
$pdcEmulatorName = Get-PDCEmulator -DomainName $DomainName
Write-Verbose "The PDC emulator for domain '$DomainName' is: $pdcEmulatorName"

$lockoutEvents = Get-LockoutEvents -PdcEmulator $pdcEmulatorName -UserName $UserName -StartTime $StartTime -Credential $Credential
$lockoutEvents | Format-Table -AutoSize
#endregion Main Script
<#
.SYNOPSIS
Retrieves logon sessions and associated users for a given computer.

.DESCRIPTION
Fetches logon session details using CIM and maps them to their respective users.

.PARAMETER ComputerName
The name of the computer for which to retrieve logon sessions.

.EXAMPLE
.\Get-LogonSessions.ps1 -ComputerName "Server1"

.NOTES
Author: Micah

.INPUTS
ComputerName: The name of the computer to be processed.

.OUTPUTS
Custom PowerShell objects containing logon session details.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ComputerName
)

# Define logon types
$logonType = @{
    "0" = "Local System"
    "2" = "Interactive"
    "3" = "Network"
    "4" = "Batch"
    "5" = "Service"
    "7" = "Unlock"
    "8" = "NetworkCleartext"
    "9" = "NewCredentials"
    "10" = "RemoteInteractive"
    "11" = "CachedInteractive"
}

# Retrieve logon sessions and logged-on users using CIM
$logonSessions = Get-CimInstance -ClassName Win32_LogonSession -ComputerName $ComputerName
$logonUsers = Get-CimInstance -ClassName Win32_LoggedOnUser -ComputerName $ComputerName

# Initialize a hashtable to map sessions to users
$sessionUser = @{}

foreach ($user in $logonUsers) {
    if ($user.Antecedent -match '.+Domain="(.+)",Name="(.+)"$') {
        $username = "${Matches[1]}\${Matches[2]}"
    }
    if ($user.Dependent -match '.+LogonId="(\d+)"$') {
        $session = $Matches[1]
    }
    $sessionUser[$session] = $username  # Assign the username to the session ID
}

foreach ($session in $logonSessions) {
    $startTime = [Management.ManagementDateTimeConverter]::ToDateTime($session.StartTime)
    
    $loggedOnUser = [PSCustomObject]@{
        Session   = $session.LogonId
        User      = $sessionUser[$session.LogonId]
        Type      = $logonType[$session.LogonType.ToString()]
        Auth      = $session.AuthenticationPackage
        StartTime = $startTime
    }

    $loggedOnUser
}












<#
.SYNOPSIS
    Retrieves logged on users and their session details on a specified computer.

.DESCRIPTION
    This script retrieves information about logged on users and their session details, including
    logon type, authentication package, and start time on a specified computer.

.PARAMETER ComputerName
    The name of the computer for which logged on user details should be retrieved.

.EXAMPLE
    Get-LoggedOnUser -ComputerName "Server01"
#>

function Get-LoggedOnUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $logonType = @{
        "0" = "Local System"
        "2" = "Interactive"
        "3" = "Network"
        "4" = "Batch"
        "5" = "Service"
        "7" = "Unlock"
        "8" = "NetworkCleartext"
        "9" = "NewCredentials"
        "10" = "RemoteInteractive"
        "11" = "CachedInteractive"
    }

    $logonSessions = Get-WmiObject -Class Win32_LogonSession -ComputerName $ComputerName
    $logonUsers = Get-WmiObject -Class Win32_LoggedOnUser -ComputerName $ComputerName

    $sessionUser = @{}

    foreach ($user in $logonUsers) {
        if ($user.Antecedent -match '.+Domain="(.+)",Name="(.+)"$') {
            $username = $Matches[1] + "\" + $Matches[2]
        }
        if ($user.Dependent -match '.+LogonId="(\d+)"$') {
            $session = $Matches[1]
        }
        $sessionUser[$session] += $username
    }

    foreach ($session in $logonSessions) {
        $startTime = [Management.ManagementDateTimeConverter]::ToDateTime($session.StartTime)

        $loggedOnUser = [PSCustomObject]@{
            Session   = $session.LogonId
            User      = $sessionUser[$session.LogonId]
            Type      = $logonType[$session.LogonType.ToString()]
            Auth      = $session.AuthenticationPackage
            StartTime = $startTime
        }

        $loggedOnUser
    }
}


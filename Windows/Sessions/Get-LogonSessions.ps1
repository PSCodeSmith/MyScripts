<#
.SYNOPSIS
    Retrieves logon session information and associated users for a given computer.

.DESCRIPTION
    This script queries a remote or local computer via CIM to retrieve information about
    logon sessions, their types, and the users associated with those sessions. It leverages
    Win32_LogonSession and Win32_LoggedOnUser WMI classes to correlate session IDs to user accounts.
    It outputs custom objects containing session details including user, authentication package,
    logon type, and start time.

.PARAMETER ComputerName
    The name of the computer to query for logon sessions. This can be a local or remote host.

.EXAMPLE
    .\Get-LogonSessions.ps1 -ComputerName "Server1"

    Retrieves all logon sessions for the computer "Server1".

.OUTPUTS
    PSCustomObject with the following properties:
    - Session: Logon session ID
    - User: Domain\User associated with the session
    - Type: Human-readable logon type
    - Auth: Authentication package used
    - StartTime: Session start time (DateTime)

.NOTES
    Author: Micah
    Updated By: [Your Name]
    Requires appropriate permissions to query the remote host.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$ComputerName
)

# Mapping of logon type integers to human-readable values
$LogonTypeMap = @{
    "0"  = "LocalSystem"
    "2"  = "Interactive"
    "3"  = "Network"
    "4"  = "Batch"
    "5"  = "Service"
    "7"  = "Unlock"
    "8"  = "NetworkCleartext"
    "9"  = "NewCredentials"
    "10" = "RemoteInteractive"
    "11" = "CachedInteractive"
}

try {
    # Retrieve CIM instances for logon sessions and logged-on users
    Write-Verbose "Querying Win32_LogonSession on $ComputerName..."
    $logonSessions = Get-CimInstance -ClassName Win32_LogonSession -ComputerName $ComputerName -ErrorAction Stop

    Write-Verbose "Querying Win32_LoggedOnUser on $ComputerName..."
    $logonUsers    = Get-CimInstance -ClassName Win32_LoggedOnUser -ComputerName $ComputerName -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve CIM data from $ComputerName: $($_.Exception.Message)"
    return
}

# Correlate sessions to users
$sessionUserMap = [System.Collections.Generic.Dictionary[string, string]]::new()

foreach ($userEntry in $logonUsers) {
    # Extract username and domain from Antecedent property
    # Antecedent looks like: \\<computer>\root\cimv2:Win32_Account.Domain="DOMAIN",Name="User"
    # Dependent looks like: \\<computer>\root\cimv2:Win32_LogonSession.LogonId="12345"
    if ($userEntry.Antecedent -match 'Domain="([^"]+)",Name="([^"]+)"') {
        $domain = $Matches[1]
        $user   = $Matches[2]
        $username = "$domain\$user"
    } else {
        # If parsing fails, skip
        continue
    }

    if ($userEntry.Dependent -match 'LogonId="(\d+)"') {
        $sessionId = $Matches[1]
        $sessionUserMap[$sessionId] = $username
    }
}

# Create output objects for each session
foreach ($session in $logonSessions) {
    $sessionId   = $session.LogonId
    $sessionType = $LogonTypeMap[$session.LogonType.ToString()] 
    if ([string]::IsNullOrWhiteSpace($sessionType)) {
        $sessionType = "Unknown($($session.LogonType))"
    }

    $startTime = [Management.ManagementDateTimeConverter]::ToDateTime($session.StartTime)
    $userName  = $sessionUserMap.ContainsKey($sessionId) ? $sessionUserMap[$sessionId] : $null

    [PSCustomObject]@{
        Session   = $sessionId
        User      = $userName
        Type      = $sessionType
        Auth      = $session.AuthenticationPackage
        StartTime = $startTime
    }
}
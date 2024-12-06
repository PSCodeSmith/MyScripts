<#
.SYNOPSIS
    This script gathers session-related events for specified computers.

.DESCRIPTION
    The script retrieves session start and stop events from the event logs of specified computers.
    It then correlates start/stop events to determine active session durations.

.PARAMETER ComputerName
    One or more computer names for which to gather session-related events. Defaults to the local computer.

.EXAMPLE
    .\Get-SessionEvents.ps1 -ComputerName "Server1","Server2"

.OUTPUTS
    Custom PowerShell objects containing session event details.

.NOTES
    Author: Micah (Original), Revised by: [Your Name]
    The script requires appropriate permissions to read event logs and to query CIM/WMI.
    It also assumes that the Windows Security and System logs contain the relevant event IDs.

.INPUTS
    [string[]]$ComputerName
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory=$false)]
    [string[]]$ComputerName = $Env:COMPUTERNAME
)

# Define session event metadata. Each key is the Event ID.
$sessionEvents = @{
    4624 = @{ ID = 4624; Label = 'Logon';                EventType = 'SessionStart'; LogName = 'Security' }
    4778 = @{ ID = 4778; Label = 'RdpSessionReconnect';  EventType = 'SessionStart'; LogName = 'Security' }
    4801 = @{ ID = 4801; Label = 'Unlocked';             EventType = 'SessionStart'; LogName = 'Security' }
    
    4647 = @{ ID = 4647; Label = 'Logoff';               EventType = 'SessionStop';  LogName = 'Security' }
    6005 = @{ ID = 6005; Label = 'Startup';             EventType = 'SessionStop';  LogName = 'System'   }
    4779 = @{ ID = 4779; Label = 'RdpSessionDisconnect'; EventType = 'SessionStop';  LogName = 'Security' }
    4800 = @{ ID = 4800; Label = 'Locked';              EventType = 'SessionStop';  LogName = 'Security' }
}

# Extract the start and stop event IDs
$sessionStartIds = ($sessionEvents.Values | Where-Object { $_.EventType -eq 'SessionStart' }).ID
$sessionStopIds  = ($sessionEvents.Values | Where-Object { $_.EventType -eq 'SessionStop' }).ID

foreach ($computer in $ComputerName) {
    Write-Verbose "Processing computer: $computer"
    
    # Gather unique log names needed
    $logNames = ($sessionEvents.Values | Select-Object -ExpandProperty LogName -Unique)

    # Build event filter hash table. 
    # We will retrieve events from the specified logs and filter by provider and event IDs.
    $eventsFilter = @{
        ComputerName    = $computer
        LogName         = $logNames
        ProviderName    = 'Microsoft-Windows-Security-Auditing'
        Level           = 0,1,2
        Id              = $sessionStartIds + $sessionStopIds
    }

    # Attempt to retrieve the events
    try {
        $events = Get-WinEvent @eventsFilter
    } catch {
        Write-Warning "Failed to retrieve events for computer [$computer]. Error: $($_.Exception.Message)"
        continue
    }

    Write-Verbose "Found [$($events.Count)] events to examine on [$computer]."

    # Determine currently logged-in user(s) on the target machine
    $loggedInUsers = try {
        # This attempts to retrieve the currently logged-on user.
        # On some servers, Win32_ComputerSystem might return a domain\username or possibly be null.
        (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $computer |
            Select-Object -ExpandProperty UserName) -replace '.*\\' # Extract the username after the backslash
    } catch {
        # If we cannot retrieve logged-in users, just set it to empty
        $null
    }

    # For each session start event, attempt to find the corresponding session stop event.
    $startEvents = $events | Where-Object { $_.Id -in $sessionStartIds }
    foreach ($evt in $startEvents) {
        try {
            $startXml = [xml]$evt.ToXml()

            # Extract username and logon ID from the event XML
            $username = ($startXml.Event.EventData.Data | Where-Object Name -eq 'TargetUserName').'#text'
            $logonId  = ($startXml.Event.EventData.Data | Where-Object Name -eq 'TargetLogonId').'#text'

            # Prepare the initial output object with session start info
            $output = [ordered]@{
                ComputerName          = $computer
                Username              = $username
                StartTime             = $evt.TimeCreated
                StartAction           = $sessionEvents[$evt.Id].Label
                StopTime              = $null
                StopAction            = $null
                'Session Active (Days)' = $null
                'Session Active (Min)'  = $null
            }

            # Attempt to find a matching stop event with the same logon ID that occurred after the start time
            $sessionEndEvent = $events | Where-Object {
                $_.Id -in $sessionStopIds -and
                $_.TimeCreated -gt $output.StartTime -and
                ([xml]$_.ToXml()).Event.EventData.Data |
                    Where-Object { $_.Name -eq 'TargetLogonId' } |
                    Select-Object -ExpandProperty '#text' -eq $logonId
            } | Select-Object -Last 1

            # If a stop event is found, use its details.
            # Otherwise, if the user is still logged in, set stop time as 'now' and action as 'Still logged in'.
            if ($sessionEndEvent) {
                $output.StopTime = $sessionEndEvent.TimeCreated
                $output.StopAction = $sessionEvents[$sessionEndEvent.Id].Label
            } else {
                if ($username -and ($username -in $loggedInUsers)) {
                    # User still logged on; current time as stop time
                    $output.StopTime = Get-Date
                    $output.StopAction = 'Still logged in'
                } else {
                    throw "No matching session stop event found for Logon ID [$logonId]."
                }
            }

            # Calculate session duration
            $sessionTimespan = New-TimeSpan -Start $output.StartTime -End $output.StopTime
            $output.'Session Active (Days)' = [math]::Round($sessionTimespan.TotalDays, 2)
            $output.'Session Active (Min)'  = [math]::Round($sessionTimespan.TotalMinutes, 2)

            # Emit the final object
            [PSCustomObject]$output

        } catch {
            # If anything goes wrong in processing a single event, log a warning and continue
            Write-Warning "Error processing event on [$computer]: $($_.Exception.Message)"
        }
    }
}
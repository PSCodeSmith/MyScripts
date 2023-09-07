<#
.SYNOPSIS
This script gathers session-related events for specified computers.

.DESCRIPTION
The script retrieves session start and stop events from the event logs of specified computers.
It then calculates the active session duration for each found session.

.PARAMETER ComputerName
The names of the computers to be processed. Defaults to the local computer.

.EXAMPLE
.\Get-SessionEvents.ps1 -ComputerName "Server1","Server2"

.NOTES
Author: Micah

.INPUTS
ComputerName: The names of the computers to be processed.

.OUTPUTS
Custom PowerShell objects containing session event details.
#>
[CmdletBinding()]
param (
    [string[]]$ComputerName = $Env:COMPUTERNAME
)

# Define session events
$sessionEvents = @{
    4624 = @{ Label = 'Logon'; EventType = 'SessionStart'; LogName = 'Security' } 
    4647 = @{ Label = 'Logoff'; EventType = 'SessionStop'; LogName = 'Security' }
    6005 = @{ Label = 'Startup'; EventType = 'SessionStop'; LogName = 'System' } 
    4778 = @{ Label = 'RdpSessionReconnect'; EventType = 'SessionStart'; LogName = 'Security' } 
    4779 = @{ Label = 'RdpSessionDisconnect'; EventType = 'SessionStop'; LogName = 'Security' } 
    4800 = @{ Label = 'Locked'; EventType = 'SessionStop'; LogName = 'Security' } 
    4801 = @{ Label = 'Unlocked'; EventType = 'SessionStart'; LogName = 'Security' } 
}

# Extract session start and stop event IDs
$sessionStartIds = $sessionEvents.Values | Where-Object { $_.EventType -eq 'SessionStart' } | Select-Object -ExpandProperty ID
$sessionStopIds = $sessionEvents.Values | Where-Object { $_.EventType -eq 'SessionStop' } | Select-Object -ExpandProperty ID

foreach ($computer in $ComputerName) {
    Write-Verbose -Message "Processing computer: $computer"
    
    # Define the filter for retrieving events
    $eventsFilter = @{
        LogName = ($sessionEvents.Values | Select-Object -ExpandProperty LogName | Select-Object -Unique)
        FilterHashtable = @{
            ProviderName = 'Microsoft-Windows-Security-Auditing'
            Level = 0,1,2
            ID = $sessionStartIds + $sessionStopIds
        }
    }
    
    try {
        $events = Get-WinEvent -ComputerName $computer @eventsFilter
    } catch {
        Write-Warning -Message "Failed to retrieve events for computer $computer. Error: $_"
        continue
    }
    
    Write-Verbose -Message "Found [$($events.Count)] events to examine"
    
    # Fetch the list of logged-in users
    $loggedInUsers = @(Get-CimInstance -ComputerName $computer -ClassName 'Win32_ComputerSystem' | 
                        Select-Object -ExpandProperty UserName | ForEach-Object { $_.split('\')[1] })

    # Process each session start event
    $events | Where-Object { $_.Id -in $sessionStartIds } | ForEach-Object {
        try {
            # Initialize the output object
            $output = [ordered]@{
                'ComputerName'          = $computer
                'Username'              = $null
                'StartTime'             = $_.TimeCreated
                'StartAction'           = $sessionEvents[$_.Id].Label
                'StopTime'              = $null
                'StopAction'            = $null
                'Session Active (Days)' = $null
                'Session Active (Min)'  = $null
            }
            
            # Extract username and logon ID
            $xEvt = [xml]$_.ToXml()
            $output.Username = ($xEvt.Event.EventData.Data | Where-Object Name -eq 'TargetUserName').'#text'
            $logonId = ($xEvt.Event.EventData.Data | Where-Object Name -eq 'TargetLogonId').'#text'
            
			# Find the corresponding session end event
			$sessionEndEvent = $events | Where-Object {
				$_.Id -in $sessionStopIds -and
				$_.TimeCreated -gt $output.StartTime -and
				([xml]$_.ToXml()).Event.EventData.Data | Where-Object Name -eq 'TargetLogonId' | Select-Object -ExpandProperty '#text' -eq $logonId
			} | Select-Object -Last 1

			# Populate StopTime and StopAction in the output object
			if ($sessionEndEvent) {
				$sessionEndDetails = $sessionEvents[$sessionEndEvent.Id]
				$output.StopTime = $sessionEndEvent.TimeCreated
				$output.StopAction = $sessionEndDetails.Label
				Write-Verbose -Message "Session stop ID is [$($sessionEndEvent.Id)]"
			} else {
				if ($output.Username -in $loggedInUsers) {
					$output.StopTime = Get-Date
					$output.StopAction = 'Still logged in'
				} else {
					throw "Could not find a session end event for logon ID [$($logonId)]."
				}
			}
            
            # Calculate session timespan and finalize output object
            $sessionTimespan = New-TimeSpan -Start $output.StartTime -End $output.StopTime
            $output.'Session Active (Days)' = [math]::Round($sessionTimespan.TotalDays, 2)
            $output.'Session Active (Min)' = [math]::Round($sessionTimespan.TotalMinutes, 2)
            
            [pscustomobject]$output
        } catch {
            Write-Warning -Message $_.Exception.Message
        }
    }
}
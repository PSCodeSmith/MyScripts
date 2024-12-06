<#
.SYNOPSIS
    Simulates keyboard activity to prevent inactivity timeouts.

.DESCRIPTION
    This script periodically simulates user input by toggling the SCROLLLOCK key. It is intended
    to keep the system active and prevent session timeouts due to inactivity. It includes an optional
    schedule that restricts when the simulation runs (e.g., only during work hours on certain days).
    Additionally, it checks if the device is on AC power before simulating activity.

.PARAMETER RunDuringWorkHours
    If $true, the script will only simulate activity during the defined work hours and days. 
    Outside of these times, it will wait and not simulate activity.

.PARAMETER StartDayOfWeek
    The first day of the workweek when activity simulation is allowed. Default is 'Monday'.

.PARAMETER EndDayOfWeek
    The last day of the workweek when activity simulation is allowed. Default is 'Friday'.

.PARAMETER StartTimeOfDay
    The start time of the workday (in HH:mm:ss format) when activity simulation is allowed. Default is '07:30:00'.

.PARAMETER EndTimeOfDay
    The end time of the workday (in HH:mm:ss format) after which activity simulation stops. Default is '17:00:00'.

.EXAMPLE
    .\KeepActive.ps1 -RunDuringWorkHours $true -StartDayOfWeek 'Tuesday' -StartTimeOfDay '08:00:00'

.NOTES
    - Requires Windows operating system and access to COM automation (WScript.Shell).
    - Adjust sleep intervals as needed.
    - The script runs indefinitely until manually stopped.

#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false,
               HelpMessage = 'Run only during defined work hours?')]
    [bool]$RunDuringWorkHours = $false,

    [Parameter(Mandatory = $false,
               HelpMessage = 'Start day of the workweek.')]
    [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
    [string]$StartDayOfWeek = 'Monday',

    [Parameter(Mandatory = $false,
               HelpMessage = 'End day of the workweek.')]
    [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
    [string]$EndDayOfWeek = 'Friday',

    [Parameter(Mandatory = $false,
               HelpMessage = 'Start time of the workday (HH:mm:ss).')]
    [string]$StartTimeOfDay = '07:30:00',

    [Parameter(Mandatory = $false,
               HelpMessage = 'End time of the workday (HH:mm:ss).')]
    [string]$EndTimeOfDay = '17:00:00'
)

function Convert-ToTimeSpan {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TimeString
    )
    try {
        return [TimeSpan]::Parse($TimeString)
    } catch {
        Write-Error "Unable to parse time string '$TimeString'. Ensure it's in HH:mm:ss format."
        exit 1
    }
}

function Test-WithinWorkHours {
    <#
    .SYNOPSIS
        Checks if the current time is within the specified work hours and days.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [DayOfWeek]$CurrentDay,
        
        [Parameter(Mandatory=$true)]
        [TimeSpan]$CurrentTime,
        
        [Parameter(Mandatory=$true)]
        [DayOfWeek]$StartDay,
        
        [Parameter(Mandatory=$true)]
        [DayOfWeek]$EndDay,
        
        [Parameter(Mandatory=$true)]
        [TimeSpan]$StartTime,
        
        [Parameter(Mandatory=$true)]
        [TimeSpan]$EndTime
    )

    # Check if current day is within the allowed range
    $dayInRange = ($CurrentDay.value__ -ge $StartDay.value__ -and $CurrentDay.value__ -le $EndDay.value__)

    # Check if current time is within the allowed range
    $timeInRange = ($CurrentTime -ge $StartTime -and $CurrentTime -le $EndTime)

    return ($dayInRange -and $timeInRange)
}

function Test-OnACPower {
    <#
    .SYNOPSIS
        Checks if the device is running on AC power.
    #>
    try {
        $batteries = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        # BatteryStatus = 2 means "On AC Power"
        # If no battery is found, assume AC power (e.g., desktop)
        if ($null -eq $batteries) {
            return $true
        } else {
            return ($batteries | Where-Object { $_.BatteryStatus -eq 2 }) -ne $null
        }
    } catch {
        # If we can't retrieve battery status, assume AC to avoid stopping the script
        Write-Verbose "Unable to determine battery status. Assuming AC power."
        return $true
    }
}

function Invoke-KeyPress {
    <#
    .SYNOPSIS
        Simulates a SCROLLLOCK key press and release to simulate user activity.
    #>
    [CmdletBinding()]
    param()

    $WShell = New-Object -ComObject Wscript.Shell
    $WShell.SendKeys("{SCROLLLOCK}")
    Start-Sleep -Milliseconds 100
    $WShell.SendKeys("{SCROLLLOCK}")
}

function Keep-Active {
    <#
    .SYNOPSIS
        Main loop that keeps the system active by toggling SCROLLLOCK at random intervals.
    #>

    # Convert provided start/end times to TimeSpan for easy comparison
    $parsedStartTime = Convert-ToTimeSpan $StartTimeOfDay
    $parsedEndTime   = Convert-ToTimeSpan $EndTimeOfDay

    $parsedStartDay = [System.DayOfWeek]::Parse($StartDayOfWeek)
    $parsedEndDay   = [System.DayOfWeek]::Parse($EndDayOfWeek)

    # Time interval (in seconds) to re-check conditions such as work hours and AC power
    $checkIntervalSec = 1800
    $nextCheckTime = (Get-Date).AddSeconds($checkIntervalSec)

    Clear-Host

    while ($true) {
        $currentTime = Get-Date

        # Periodically re-check conditions
        if ($currentTime -ge $nextCheckTime) {
            # If restricted to work hours, check if we should continue
            if ($RunDuringWorkHours) {
                $withinHours = Test-WithinWorkHours -CurrentDay $currentTime.DayOfWeek `
                                                       -CurrentTime $currentTime.TimeOfDay `
                                                       -StartDay $parsedStartDay `
                                                       -EndDay $parsedEndDay `
                                                       -StartTime $parsedStartTime `
                                                       -EndTime $parsedEndTime

                if (-not $withinHours) {
                    Write-Verbose "Outside of defined work hours. Waiting for $checkIntervalSec seconds before re-checking."
                    Start-Sleep -Seconds $checkIntervalSec
                    $nextCheckTime = (Get-Date).AddSeconds($checkIntervalSec)
                    continue
                }
            }

            # Check AC power status
            if (-not (Test-OnACPower)) {
                Write-Verbose "Not on AC power. Waiting for $checkIntervalSec seconds before re-checking."
                Start-Sleep -Seconds $checkIntervalSec
                $nextCheckTime = (Get-Date).AddSeconds($checkIntervalSec)
                continue
            }

            # Update the next scheduled check time
            $nextCheckTime = (Get-Date).AddSeconds($checkIntervalSec)
        }

        # Simulate activity by pressing SCROLLLOCK
        Invoke-KeyPress

        # Sleep for a random interval between keypresses
        $sleepTime = Get-Random -Minimum 140 -Maximum 280
        Write-Verbose "Simulated activity. Sleeping for $sleepTime seconds..."
        Start-Sleep -Seconds $sleepTime
    }
}

# Start the activity simulation
Keep-Active
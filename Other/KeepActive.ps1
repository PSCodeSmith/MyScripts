<#
	.SYNOPSIS
		Simulates keyboard activity to prevent inactivity timeouts.
	
	.DESCRIPTION
		The script sends a SCROLLLOCK key press and release at random intervals to simulate activity.
		It checks the current day, time, and battery status every half hour to decide whether to continue running.
	
	.PARAMETER RunDuringWorkHours
		If set to $true, the script will only run between the defined work hours and days.
	
	.PARAMETER StartDayOfWeek
		Defines the starting day of the workweek. Default is 'Monday'.
	
	.PARAMETER EndDayOfWeek
		Defines the ending day of the workweek. Default is 'Friday'.
	
	.PARAMETER StartTimeOfDay
		Defines the start time of the workday. Default is '07:30:00'.
	
	.PARAMETER EndTimeOfDay
		Defines the end time of the workday. Default is '17:00:00'.
	
	.EXAMPLE
		.\KeepActive.ps1 -RunDuringWorkHours $true -StartDayOfWeek 'Tuesday' -StartTimeOfDay '08:00:00'
	
	.NOTES
		Additional information about the file.
#>
[CmdletBinding()]
param
(
	[Parameter(Mandatory = $false,
			   HelpMessage = 'Run only during work hours?')]
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

function Keep-Active
{
	Clear-Host
	$WShell = New-Object -ComObject Wscript.Shell
	$nextCheckTime = (Get-Date).AddSeconds(1800) # Set the initial time for the next check
	
	while ($true)
	{
		$currentTime = Get-Date
		
		if ($currentTime -ge $nextCheckTime)
		{
			if ($RunDuringWorkHours -eq $true)
			{
				$dayOfWeek = $currentTime.DayOfWeek
				$timeOfDay = $currentTime.TimeOfDay
				
				# Check if it's outside the defined work hours or workweek
				if (
					$dayOfWeek.value__ -lt [System.Enum]::Parse([System.DayOfWeek], $StartDayOfWeek).value__ -or
					$dayOfWeek.value__ -gt [System.Enum]::Parse([System.DayOfWeek], $EndDayOfWeek).value__ -or
					$timeOfDay -lt $StartTimeOfDay -or
					$timeOfDay -gt $EndTimeOfDay
				)
				{
					Start-Sleep -Seconds 1800
					$nextCheckTime = $currentTime.AddSeconds(1800)
					continue
				}
			}
			
			$batteryStatus = Get-WmiObject Win32_Battery | Select-Object -ExpandProperty BatteryStatus
			if ($batteryStatus -ne 2)
			{
				# Not plugged in
				Start-Sleep -Seconds 1800
				$nextCheckTime = $currentTime.AddSeconds(1800)
				continue
			}
			
			# Update the time for the next check
			$nextCheckTime = $currentTime.AddSeconds(1800)
		}
		
		$time = Get-Random -Minimum 140 -Maximum 280
		$WShell.SendKeys("{SCROLLLOCK}")
		Start-Sleep -Milliseconds 100
		$WShell.SendKeys("{SCROLLLOCK}")
		Write-Verbose "sleeping for $time seconds"
		Start-Sleep -Seconds $time
	}
}

# Call the function to initiate the activity simulation
Keep-Active

<#
.SYNOPSIS
Exports Outlook calendar events to Obsidian markdown files.

.DESCRIPTION
This script retrieves calendar events from Microsoft Outlook and creates corresponding markdown files in a specified Obsidian folder structure. It supports filtering by date range and Outlook category.

.PARAMETER ObsidianFolder
The root folder where Obsidian markdown files will be created.

.PARAMETER DaysBack
Number of days in the past to retrieve events for. Default is -30.

.PARAMETER DaysForward
Number of days in the future to retrieve events for. Default is 0.

.PARAMETER OutlookCategory
The Outlook category to filter events by. If empty, all events are included.

.PARAMETER DefaultAttendees
An array of default attendees to add to each meeting note.

.EXAMPLE
.\Outlook-to-Obsidian.ps1 -ObsidianFolder "C:\ObsidianVault" -DaysBack -7 -DaysForward 7 -OutlookCategory "Work" -Verbose

.NOTES
Requires Microsoft Outlook to be installed and running.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string] $ObsidianFolder,

    [Parameter(Mandatory=$false)]
    [int] $DaysBack = -30,

    [Parameter(Mandatory=$false)]
    [int] $DaysForward = 0,

    [Parameter(Mandatory=$false)]
    [string] $OutlookCategory = "",

    [Parameter(Mandatory=$false)]
    [string[]] $DefaultAttendees = @()
)

function Get-OutlookCalendar {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [int] $OUTLOOK_CALENDAR_FOLDER = 9,

        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Container})]
        [string] $OBSIDIAN_FOLDER,

        [Parameter(Mandatory=$false)]
        [int] $DaysBack = -30,

        [Parameter(Mandatory=$false)]
        [int] $DaysForward = 0,

        [Parameter(Mandatory=$false)]
        [string] $OutlookCategory = "",

        [Parameter(Mandatory=$false)]
        [string[]] $DefaultAttendees = @()
    )

    Begin {
        Write-Verbose "Starting Get-OutlookCalendar function"
        Write-Verbose "Checking if Outlook is running..."
        if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
            throw "Outlook is not running. Please start Outlook and try again."
        }
        Write-Verbose "Outlook is running"

        Write-Verbose "Loading Microsoft.Office.Interop.Outlook assembly..."
        try {
            Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
            Write-Verbose "Microsoft.Office.Interop.Outlook assembly loaded successfully"
        }
        catch {
            throw "Failed to load Microsoft.Office.Interop.Outlook assembly: $_"
        }

        $outlook = $null
        $namespace = $null
        $folder = $null
    }

    Process {
        try {
            Write-Verbose "Creating Outlook COM objects..."
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $folder = $namespace.GetDefaultFolder($OUTLOOK_CALENDAR_FOLDER)
            Write-Verbose "Outlook COM objects created successfully"

            $startDate = (Get-Date).AddDays($DaysBack)
            $endDate = (Get-Date).AddDays($DaysForward + 1).AddSeconds(-1)
            Write-Verbose "Date range set: $startDate to $endDate"

            Write-Verbose "Retrieving and filtering meetings..."
            $items = $folder.Items
            $items.IncludeRecurrences = $true
            $items.Sort("[Start]")
            $filter = "[Start] >= '" + $startDate.ToString("g") + "' AND [End] <= '" + $endDate.ToString("g") + "'"
            $meetings = $items.Restrict($filter)
            Write-Verbose "Retrieved $($meetings.Count) meetings"

            foreach ($meeting in $meetings) {
                if ($OutlookCategory -and $meeting.Categories -notmatch $OutlookCategory) { 
                    Write-Verbose "Skipping meeting '$($meeting.Subject)' due to category mismatch"
                    continue 
                }

                Write-Verbose "Processing meeting: $($meeting.Subject)"
                $meetingStartTime = $meeting.Start.ToString("yyyy-MM-ddTHH:mm:ss")
                $meetingStartTimeYear = $meeting.Start.ToString("yyyy")
                $meetingStartTimeMonth = $meeting.Start.ToString("MM")
                $meetingSummary = Format-StringForObsidian $meeting.Subject

                try {
                    $meetingBody = Get-TruncatedMeetingBody ($meeting.Body -replace "`r`n", "`n")
                }
                catch {
                    Write-Warning "Error processing meeting body for '$($meeting.Subject)': $_"
                    $meetingBody = ""
                }

                Write-Verbose "Creating note content..."
                $noteContent = @"
---
meetingStartTime: $meetingStartTime
meetingEndTime: $($meeting.End.ToString("yyyy-MM-ddTHH:mm:ss"))
type: meeting
meetingSummary: "$meetingSummary"
meetingRecurring: $($meeting.IsRecurring)
meetingOrganizer: "$($meeting.Organizer)"
meetingLocation: "$($meeting.Location)"
meetingDuration: "$($meeting.Duration)"
Attendees:
$(($DefaultAttendees | ForEach-Object { "  - '$_'" }) -join "`r`n")
tags:
- meetings
---

> [!info]- Meeting Organizer
> $($meeting.Organizer)

> [!info]- Required Attendees
> $($meeting.RequiredAttendees -join ", ")

> [!info]- Optional Attendees
> $($meeting.OptionalAttendees -join ", ")

> [!info]- Location
> $($meeting.Location)

> [!info]- Meeting Agenda (from Outlook)
$(if ([string]::IsNullOrWhiteSpace($meetingBody)) { "> No agenda available" } else { $meetingBody -split "`n" | ForEach-Object { "> $_ " } | Out-String })

## ⭐Agenda/Questions

- 

---

## 📝 Notes

- 

---
## ✅ Action Items

- [ ] 

"@
                Write-Verbose "Creating folder structure..."
                $yearFolder = Join-Path $OBSIDIAN_FOLDER $meetingStartTimeYear
                $monthFolder = Join-Path $yearFolder $meetingStartTimeMonth

                New-Item -ItemType Directory -Path $yearFolder, $monthFolder -Force | Out-Null
                Write-Verbose "Folder structure created: $monthFolder"

                $fileName = "$($meeting.Start.ToString('yyyy-MM-dd')) $($meetingSummary).md"
                $filePath = Join-Path $monthFolder $fileName

                if (Test-Path $filePath) {
                    Write-Warning "File already exists, skipping: $filePath"
                }
                else {
                    $noteContent | Out-File -FilePath $filePath -Encoding utf8
                    Write-Verbose "Created note: $filePath"
                }
            }
        }
        catch {
            Write-Error "An error occurred: $_"
        }
    }

    End {
        Write-Verbose "Cleaning up COM objects..."
        if ($null -ne $namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
        if ($null -ne $outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Verbose "COM objects cleaned up"
        Write-Verbose "Get-OutlookCalendar function completed"
    }
}

function Format-StringForObsidian {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$InputString
    )

    process {
        Write-Verbose "Formatting string for Obsidian: '$InputString'"
        $formattedString = $InputString -replace '(\d{4})/(\d{2})/(\d{2})', '$1-$2-$3'
        $formattedString = $formattedString -replace '\s_(\w)', ' $1' -replace '(\w)_\s', '$1 ' -replace '(\w)_(\W)', '$1$2' -replace '(\W)_(\w)', '$1$2'
        $formattedString = $formattedString -replace '\*', '' `
                    -replace '"', "'" `
                    -replace '\\', '-' `
                    -replace '\bw/', 'with' `
                    -replace '/', '-' `
                    -replace '<', '(' `
                    -replace '>', ')' `
                    -replace ':', ' -' `
                    -replace '\|', '-' `
                    -replace '\?', '' `
                    -replace '\[', '(' `
                    -replace '\]', ')'
        $formattedString = $formattedString.Trim(' -')
        $formattedString = $formattedString -replace '\s+', ' ' `
                    -replace '-+', '-'
        Write-Verbose "Formatted string: '$formattedString'"
        return $formattedString
    }
}

function Get-TruncatedMeetingBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Body
    )

    process {
        Write-Verbose "Truncating meeting body..."
        if ([string]::IsNullOrWhiteSpace($Body)) {
            Write-Verbose "Meeting body is empty or null"
            return ""
        }

        $cutoffPhrases = @(
            "Microsoft Teams Need help?",
            "Join on your computer, mobile app or room device"
        )

        $indices = $cutoffPhrases | ForEach-Object { $Body.IndexOf($_) } | Where-Object { $_ -ge 0 }
        
        if ($indices.Count -gt 0) {
            $cutoffIndex = ($indices | Measure-Object -Minimum).Minimum
            $truncatedBody = $Body.Substring(0, $cutoffIndex).Trim()
            Write-Verbose "Meeting body truncated at index $cutoffIndex"
        } else {
            $truncatedBody = $Body
            Write-Verbose "No cutoff phrases found, returning full body"
        }
        return $truncatedBody
    }
}

# Main execution
Write-Verbose "Script started with parameters:"
Write-Verbose "ObsidianFolder: $ObsidianFolder"
Write-Verbose "DaysBack: $DaysBack"
Write-Verbose "DaysForward: $DaysForward"
Write-Verbose "OutlookCategory: $OutlookCategory"
Write-Verbose "DefaultAttendees: $($DefaultAttendees -join ', ')"

Get-OutlookCalendar -OBSIDIAN_FOLDER $ObsidianFolder -DaysBack $DaysBack -DaysForward $DaysForward -OutlookCategory $OutlookCategory -DefaultAttendees $DefaultAttendees -Verbose

Write-Verbose "Script completed"
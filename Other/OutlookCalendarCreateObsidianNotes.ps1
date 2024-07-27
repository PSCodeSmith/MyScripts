function Get-OutlookCalendar {
    <#
    .SYNOPSIS
    Exports Outlook calendar events to Obsidian markdown files.

    .DESCRIPTION
    This function retrieves calendar events from Microsoft Outlook and creates corresponding markdown files in a specified Obsidian folder structure. It supports filtering by date range and Outlook category.

    .PARAMETER OUTLOOK_CALENDAR_FOLDER
    The Outlook folder constant for the calendar. Default is 9 (olFolderCalendar).

    .PARAMETER OBSIDIAN_FOLDER
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
    Get-OutlookCalendar -OBSIDIAN_FOLDER "C:\ObsidianVault" -DaysBack -7 -DaysForward 7 -OutlookCategory "Work"

    .NOTES
    Requires Microsoft Outlook to be installed and running.
    #>
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
        # Check if Outlook is running
        if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
            throw "Outlook is not running. Please start Outlook and try again."
        }

        # Load required assembly
        try {
            Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        }
        catch {
            throw "Failed to load Microsoft.Office.Interop.Outlook assembly: $_"
        }

        # Initialize Outlook objects
        $outlook = $null
        $namespace = $null
        $folder = $null
    }

    Process {
        try {
            # Create Outlook objects
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $folder = $namespace.GetDefaultFolder($OUTLOOK_CALENDAR_FOLDER)

            # Set date range
            $startDate = (Get-Date).AddDays($DaysBack)
            $endDate = (Get-Date).AddDays($DaysForward + 1).AddSeconds(-1)

            # Retrieve and filter meetings
            $items = $folder.Items
            $items.IncludeRecurrences = $true
            $items.Sort("[Start]")
            $filter = "[Start] >= '" + $startDate.ToString("g") + "' AND [End] <= '" + $endDate.ToString("g") + "'"
            $meetings = $items.Restrict($filter)

            foreach ($meeting in $meetings) {
                # Apply category filter if specified
                if ($OutlookCategory -and $meeting.Categories -notmatch $OutlookCategory) { continue }

                # Process meeting details
                $meetingStartTime = $meeting.Start.ToString("yyyy-MM-ddTHH:mm:ss")
                $meetingStartTimeYear = $meeting.Start.ToString("yyyy")
                $meetingStartTimeMonth = $meeting.Start.ToString("MM")
                $meetingSummary = Clean-String $meeting.Subject
                $meetingBody = Truncate-MeetingBody $meeting.Body.Replace("`r`n", "`n")

                # Create note content
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
$($meetingBody -split "`n" | ForEach-Object { "> $_ " } | Out-String)

## ⭐Agenda/Questions

- 

---

## 📝 Notes

- 

---
## ✅ Action Items

- [ ] 

"@
                # Create folder structure
                $yearFolder = Join-Path $OBSIDIAN_FOLDER $meetingStartTimeYear
                $monthFolder = Join-Path $yearFolder $meetingStartTimeMonth

                New-Item -ItemType Directory -Path $yearFolder, $monthFolder -Force | Out-Null

                # Create file
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
        # Clean up COM objects
        if ($null -ne $namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
        if ($null -ne $outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Clean-String {
    <#
    .SYNOPSIS
    Cleans a string for use as an Obsidian note name, removing or replacing all illegal characters and preventing unintended formatting.

    .DESCRIPTION
    This function takes a string input and removes or replaces all characters that are illegal in Obsidian note names (* " \ / < > : | ?). 
    It preserves common date formats, replaces some common abbreviations with more readable alternatives, and prevents unintended formatting in Obsidian data views by removing underscores in specific positions.

    .PARAMETER str
    The input string to be cleaned.

    .EXAMPLE
    Clean-String "Meeting: [Team] * _Project Update_ * 2023/07/28 | Status w/ Management?"
    Returns: "Meeting - (Team) Project Update 2023-07-28 - Status with Management"
    #>
    [CmdletBinding()]
    param([string]$str)

    # First, protect date formats
    $str = $str -replace '(\d{4})/(\d{2})/(\d{2})', '$1-$2-$3'

    # Remove underscores that could cause unintended formatting (single line version)
    $str = $str -replace '\s_(\w)', ' $1' -replace '(\w)_\s', '$1 ' -replace '(\w)_(\W)', '$1$2' -replace '(\W)_(\w)', '$1$2'

    # Then perform other replacements
    $str = $str -replace '\*', '' `
                -replace '"', "'" `
                -replace '\\', '-' `
                -replace '\bw/', 'with' `  # Move this replacement before the general '/' replacement
                -replace '/', '-' `
                -replace '<', '(' `
                -replace '>', ')' `
                -replace ':', ' -' `
                -replace '\|', '-' `
                -replace '\?', '' `
                -replace '\[', '(' `
                -replace '\]', ')'

    # Trim any leading or trailing spaces and dashes
    $str = $str.Trim(' -')

    # Replace any multiple spaces or dashes with a single instance
    $str = $str -replace '\s+', ' ' `
                -replace '-+', '-'

    return $str
}

function Truncate-MeetingBody {
    <#
    .SYNOPSIS
    Truncates the body of a meeting invitation.

    .DESCRIPTION
    This function removes standard Microsoft Teams joining information from the meeting body to keep only the relevant content.

    .PARAMETER body
    The full body text of the meeting invitation.

    .EXAMPLE
    Truncate-MeetingBody $meetingBody
    #>
    [CmdletBinding()]
    param([string]$body)

    $cutoffPhrases = @(
        "Microsoft Teams Need help?",
        "Join on your computer, mobile app or room device"
    )

    $indices = $cutoffPhrases | ForEach-Object { $body.IndexOf($_) } | Where-Object { $_ -ge 0 }
    
    if ($indices.Count -gt 0) {
        $cutoffIndex = ($indices | Measure-Object -Minimum).Minimum
        return $body.Substring(0, $cutoffIndex).Trim()
    }
    return $body
}
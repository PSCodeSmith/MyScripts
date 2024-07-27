function Get-OutlookCalendar {
    [CmdletBinding()]
    param (
        [Parameter()]
        [int] $OUTLOOK_CALENDAR_FOLDER = 9,
        [Parameter()]
        [string] $OBSIDIAN_FOLDER = "",
        [Parameter()]
        [int] $daysBack = -30,
        [Parameter()]
        [int] $daysForward = 0,
        [Parameter()]
        [string] $outlookCategory = "",
        [Parameter()]
        [string[]] $defaultAttendees = @("")
    )

    if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
        throw "Outlook is not running. Please start Outlook and try again."
    }

    try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $folder = $namespace.GetDefaultFolder($OUTLOOK_CALENDAR_FOLDER)

        $startDate = (Get-Date).AddDays($daysBack)
        $endDate = (Get-Date).AddDays($daysForward + 1).AddSeconds(-1)

        $items = $folder.Items
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")
        $filter = "[Start] >= '" + $startDate.ToString("g") + "' AND [End] <= '" + $endDate.ToString("g") + "'"
        $meetings = $items.Restrict($filter)

        foreach ($meeting in $meetings) {
            if ($outlookCategory -and $meeting.Categories -notmatch $outlookCategory) { continue }

            $meetingStartTime = $meeting.Start.ToString("yyyy-MM-ddTHH:mm:ss")
            $meetingStartTimeYear = $meeting.Start.ToString("yyyy")
            $meetingStartTimeMonth = $meeting.Start.ToString("MM")
            $meetingSummary = Clean-String $meeting.Subject
            $meetingBody = Truncate-MeetingBody $meeting.Body.Replace("`r`n", "`n")

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
$(($defaultAttendees | ForEach-Object { "  - '$_'" }) -join "`r`n")
tags:
- meetings
---

## ⭐Agenda/Questions

> [!info]+ Meeting Details
**Organizer:** $($meeting.Organizer)
**Required Attendees:** $($meeting.RequiredAttendees -join ", ")
**Optional Attendees:** $($meeting.OptionalAttendees -join ", ")
**Location:** $($meeting.Location)
> 
$($meetingBody -split "`n" | ForEach-Object { "> $_ " } | Out-String)

---

## 📝 Notes

---
## ✅ Action Items

- [ ]

"@
            $yearFolder = Join-Path $OBSIDIAN_FOLDER $meetingStartTimeYear
            $monthFolder = Join-Path $yearFolder $meetingStartTimeMonth

            if (-not (Test-Path $yearFolder))
            {
                New-Item -ItemType Directory -Path $yearFolder | Out-Null
            }

            if (-not (Test-Path $monthFolder))
            {
                New-Item -ItemType Directory -Path $monthFolder | Out-Null
            }

            $fileName = "$($meeting.Start.ToString('yyyy-MM-dd')) $($meetingSummary).md"
            $filePath = Join-Path $monthFolder $fileName

            if (Test-Path $filePath) {
                Write-Host "File already exists, skipping: $filePath"
            }
            else {
                $noteContent | Out-File -FilePath $filePath -Encoding utf8
                Write-Host "Created note: $filePath"
            }
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }
    finally {
        if ($null -ne $namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
        if ($null -ne $outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Clean-String {
    param([string]$str)
    $str -replace '\[', '(' -replace '\]', ')' -replace '//', '&' -replace ':', ' -' -replace '[/\\]', '&' -replace '[\"*:<>?]', '' -replace '[^\w\-_\. ()&]', '_'
}

function Truncate-MeetingBody {
    param([string]$body)
    $indices = @(
        $body.IndexOf("Microsoft Teams Need help?"),
        $body.IndexOf("Join on your computer, mobile app or room device")
    ) | Where-Object { $_ -ge 0 }
    
    if ($indices.Count -gt 0) {
        $cutoffIndex = ($indices | Measure-Object -Minimum).Minimum
        return $body.Substring(0, $cutoffIndex).Trim()
    }
    return $body
}

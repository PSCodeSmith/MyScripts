<#
.SYNOPSIS
    Exports Outlook calendar events to Obsidian markdown files.

.DESCRIPTION
    This script retrieves calendar events from Microsoft Outlook within a specified date range
    and creates corresponding markdown files in a structured Obsidian folder layout (Year/Month).
    It supports filtering by Outlook category and allows specifying default attendees for every note.

.PARAMETER ObsidianFolder
    The root folder where Obsidian markdown files will be created.

.PARAMETER DaysBack
    Number of days in the past from today to retrieve events. Default is -30 (30 days in the past).

.PARAMETER DaysForward
    Number of days in the future from today to retrieve events. Default is 0 (no future events).

.PARAMETER OutlookCategory
    The Outlook category to filter events by. If empty, all events are included.

.PARAMETER DefaultAttendees
    An array of default attendees to add to each meeting note.

.EXAMPLE
    .\Outlook-to-Obsidian.ps1 -ObsidianFolder "C:\ObsidianVault" -DaysBack -7 -DaysForward 7 -OutlookCategory "Work" -Verbose

.NOTES
    - Requires Microsoft Outlook to be installed and running.
    - Requires the Microsoft.Office.Interop.Outlook assembly.

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

#region Helper Functions

function Test-OutlookRunning {
    <#
    .SYNOPSIS
        Checks if Outlook is running.

    .DESCRIPTION
        Returns $true if Outlook is running, otherwise $false.
    #>
    [CmdletBinding()]
    param()

    return [bool](Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
}

function Initialize-OutlookInterop {
    <#
    .SYNOPSIS
        Loads the Outlook Interop Assembly and returns the Outlook COM objects.

    .DESCRIPTION
        Loads the Microsoft.Office.Interop.Outlook assembly, creates the Outlook COM application object,
        and retrieves the MAPI namespace and calendar folder.

    .PARAMETER CalendarFolderConstant
        An integer representing the Outlook default folder for the calendar (usually 9).

    .OUTPUTS
        [Microsoft.Office.Interop.Outlook._Folder], [Microsoft.Office.Interop.Outlook._NameSpace], [Microsoft.Office.Interop.Outlook._Application]

    .NOTES
        Throws if Outlook is not running or the assembly fails to load.
    #>
    [CmdletBinding()]
    param(
        [int]$CalendarFolderConstant = 9
    )

    Write-Verbose "Checking if Outlook is running..."
    if (-not (Test-OutlookRunning)) {
        throw "Outlook is not running. Please start Outlook and try again."
    }

    Write-Verbose "Loading Microsoft.Office.Interop.Outlook assembly..."
    try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        Write-Verbose "Microsoft.Office.Interop.Outlook assembly loaded."
    } catch {
        throw "Failed to load Microsoft.Office.Interop.Outlook assembly: $($_.Exception.Message)"
    }

    Write-Verbose "Creating Outlook COM objects..."
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $folder = $namespace.GetDefaultFolder($CalendarFolderConstant)
    Write-Verbose "Outlook COM objects created and calendar folder retrieved."

    return $folder, $namespace, $outlook
}

function Format-StringForObsidian {
    <#
    .SYNOPSIS
        Formats a given string for safe usage as a note title and content in Obsidian.

    .DESCRIPTION
        Replaces or removes invalid or problematic characters and patterns that may cause issues
        in Obsidian file names or Markdown content.

    .PARAMETER InputString
        The string to be formatted.

    .EXAMPLE
        Format-StringForObsidian "Project/Meeting: Status?"
        Returns something like "Project-Meeting - Status"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$InputString
    )

    process {
        Write-Verbose "Formatting string for Obsidian: '$InputString'"

        # Convert dates like YYYY/MM/DD to YYYY-MM-DD
        $formattedString = $InputString -replace '(\d{4})/(\d{2})/(\d{2})', '$1-$2-$3'

        # Remove underscores around words
        $formattedString = $formattedString `
            -replace '\s_(\w)', ' $1' `
            -replace '(\w)_\s', '$1 ' `
            -replace '(\w)_(\W)', '$1$2' `
            -replace '(\W)_(\w)', '$1$2'

        # Replace problematic characters
        $formattedString = $formattedString `
            -replace '\*', '' `
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

        # Trim and normalize spacing/dashes
        $formattedString = $formattedString.Trim(' -')
        $formattedString = $formattedString `
            -replace '\s+', ' ' `
            -replace '-+', '-'

        Write-Verbose "Formatted string: '$formattedString'"
        return $formattedString
    }
}

function Get-TruncatedMeetingBody {
    <#
    .SYNOPSIS
        Truncates the meeting body at known cutoff phrases to remove unnecessary Teams join info.

    .DESCRIPTION
        Removes everything after certain known phrases that appear in meeting bodies
        (like Microsoft Teams links or boilerplate text).

    .PARAMETER Body
        The full meeting body text.

    .EXAMPLE
        Get-TruncatedMeetingBody "Agenda line 1`nMicrosoft Teams Need help?"
        Returns "Agenda line 1"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Body
    )

    process {
        Write-Verbose "Truncating meeting body..."

        if ([string]::IsNullOrWhiteSpace($Body)) {
            Write-Verbose "Meeting body is empty or null."
            return ""
        }

        $cutoffPhrases = @(
            "Microsoft Teams Need help?",
            "Join on your computer, mobile app or room device"
        )

        $indices = foreach ($phrase in $cutoffPhrases) {
            $idx = $Body.IndexOf($phrase)
            if ($idx -ge 0) { $idx }
        }

        if ($indices -and $indices.Count -gt 0) {
            $cutoffIndex = ($indices | Measure-Object -Minimum).Minimum
            $truncatedBody = $Body.Substring(0, $cutoffIndex).Trim()
            Write-Verbose "Meeting body truncated at index $cutoffIndex"
        } else {
            $truncatedBody = $Body
            Write-Verbose "No cutoff phrases found, returning full body."
        }
        return $truncatedBody
    }
}

#endregion

#region Main Function

function Export-OutlookCalendarToObsidian {
    <#
    .SYNOPSIS
        Retrieves Outlook calendar events and exports them as Obsidian notes.

    .DESCRIPTION
        Connects to Outlook (which must be running), retrieves calendar items within the specified date range,
        optionally filters by category, and creates a note in the specified Obsidian folder structure for each meeting.
        The folder structure is Year/Month, and each note contains metadata and placeholders for agenda, notes, and action items.

    .PARAMETER ObsidianFolder
        The root folder where Obsidian markdown files will be created.

    .PARAMETER DaysBack
        Number of days in the past from today to start retrieving events. Default is -30.

    .PARAMETER DaysForward
        Number of days in the future from today to retrieve events. Default is 0.

    .PARAMETER OutlookCategory
        If specified, only events matching this Outlook category are included.

    .PARAMETER DefaultAttendees
        Default attendees to include in each generated note.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Container})]
        [string]$ObsidianFolder,

        [Parameter(Mandatory=$false)]
        [int]$DaysBack = -30,

        [Parameter(Mandatory=$false)]
        [int]$DaysForward = 0,

        [Parameter(Mandatory=$false)]
        [string]$OutlookCategory = "",

        [Parameter(Mandatory=$false)]
        [string[]]$DefaultAttendees = @()
    )

    Write-Verbose "Initializing Outlook interop..."
    $CalendarFolder, $Namespace, $OutlookApp = Initialize-OutlookInterop

    try {
        $startDate = (Get-Date).AddDays($DaysBack)
        $endDate   = (Get-Date).AddDays($DaysForward + 1).AddSeconds(-1)
        Write-Verbose "Date range: $startDate to $endDate"

        Write-Verbose "Retrieving and filtering meetings..."
        $items = $CalendarFolder.Items
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")
        $filter = "[Start] >= '" + $startDate.ToString("g") + "' AND [End] <= '" + $endDate.ToString("g") + "'"
        $meetings = $items.Restrict($filter)

        Write-Verbose "Found $($meetings.Count) meetings in specified date range."

        foreach ($meeting in $meetings) {
            # Filter by category if specified
            if ([string]::IsNullOrWhiteSpace($OutlookCategory) -eq $false) {
                if ($meeting.Categories -notmatch $OutlookCategory) {
                    Write-Verbose "Skipping '$($meeting.Subject)' due to category mismatch."
                    continue
                }
            }

            Write-Verbose "Processing meeting: $($meeting.Subject)"
            $meetingStartTime = $meeting.Start.ToString("yyyy-MM-ddTHH:mm:ss")
            $yearFolder       = Join-Path $ObsidianFolder ($meeting.Start.ToString("yyyy"))
            $monthFolder      = Join-Path $yearFolder ($meeting.Start.ToString("MM"))

            # Ensure folder structure exists
            New-Item -ItemType Directory -Path $yearFolder -Force | Out-Null
            New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null
            Write-Verbose "Ensured folder structure: $monthFolder"

            $meetingSummary = $meeting.Subject | Format-StringForObsidian
            $fileName       = "{0} {1}.md" -f ($meeting.Start.ToString('yyyy-MM-dd')), $meetingSummary
            $filePath       = Join-Path $monthFolder $fileName

            if (Test-Path $filePath) {
                Write-Warning "File already exists, skipping: $filePath"
                continue
            }

            # Attempt to process meeting body
            try {
                $meetingBody = $meeting.Body -replace "`r`n", "`n"
                $meetingBody = Get-TruncatedMeetingBody $meetingBody
            } catch {
                Write-Warning "Error processing meeting body for '$($meeting.Subject)': $($_.Exception.Message)"
                $meetingBody = ""
            }

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
> $(if ($meeting.RequiredAttendees) { $meeting.RequiredAttendees -join ", " } else { "None" })

> [!info]- Optional Attendees
> $(if ($meeting.OptionalAttendees) { $meeting.OptionalAttendees -join ", " } else { "None" })

> [!info]- Location
> $($meeting.Location)

> [!info]- Meeting Agenda (from Outlook)
$(if ([string]::IsNullOrWhiteSpace($meetingBody)) { "> No agenda available" } else { $meetingBody -split "`n" | ForEach-Object { "> $_" } | Out-String })

## ⭐Agenda/Questions

- 

---

## 📝 Notes

- 

---

## ✅ Action Items

- [ ] 

"@

            Write-Verbose "Writing note to $filePath"
            $noteContent | Out-File -FilePath $filePath -Encoding utf8
            Write-Verbose "Created note: $filePath"
        }

    } catch {
        Write-Error "An error occurred during processing: $($_.Exception.Message)"
    } finally {
        # Clean up COM objects
        Write-Verbose "Cleaning up COM objects..."
        if ($null -ne $Namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null }
        if ($null -ne $OutlookApp) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutlookApp) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Verbose "COM objects cleaned up."
    }
}

#endregion Main Function

# Main execution
Write-Verbose "Script starting with parameters:"
Write-Verbose "ObsidianFolder: $ObsidianFolder"
Write-Verbose "DaysBack: $DaysBack"
Write-Verbose "DaysForward: $DaysForward"
Write-Verbose "OutlookCategory: $OutlookCategory"
Write-Verbose "DefaultAttendees: $($DefaultAttendees -join ', ')"

Export-OutlookCalendarToObsidian -ObsidianFolder $ObsidianFolder -DaysBack $DaysBack -DaysForward $DaysForward -OutlookCategory $OutlookCategory -DefaultAttendees $DefaultAttendees -Verbose

Write-Verbose "Script completed."
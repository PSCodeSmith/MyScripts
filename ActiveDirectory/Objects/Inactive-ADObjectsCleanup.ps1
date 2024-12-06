<#
.SYNOPSIS
    A script to manage inactive users, computers, and empty local admin groups in Active Directory.

.DESCRIPTION
    This script provides several functions to maintain an Active Directory environment by:
    - Disabling users who have been inactive beyond a certain threshold.
    - Optionally moving these disabled users to a 'pending deletion' OU.
    - Removing disabled users who have remained inactive beyond another threshold.
    - Removing inactive computers from AD.
    - Removing empty local admin groups.

    Each function uses provided search bases (OUs) and can exclude specific OUs to ensure critical accounts or
    groups remain untouched. The script relies on the Active Directory module cmdlets and requires appropriate
    permissions.

.PARAMETER userSearchBaseOU
    Specifies the OU for the initial user search (e.g. "OU=Users,DC=example,DC=com").

.PARAMETER pendingUserDeletionOU
    Specifies the OU to which disabled users should be moved before final deletion (e.g. "OU=Pending Deletion,OU=Users,DC=example,DC=com").

.PARAMETER computerSearchBaseOU
    Specifies the OU for the computer search (e.g. "OU=Computers,DC=example,DC=com").

.PARAMETER localAdminGroupSearchBaseOU
    Specifies the OU for local admin group search (e.g. "OU=Local Admin Groups,OU=Groups,DC=example,DC=com").

.EXAMPLE
    Disable-InactiveUsers -DaysInactive 30 -DescriptionPrefix "Disabled" -MoveToPendingDeletion

.EXAMPLE
    Remove-InactiveUsers -DaysInactive 30

.EXAMPLE
    Remove-InactiveComputers -DaysInactive 30

.EXAMPLE
    Remove-EmptyLocalAdminGroups

.NOTES
    - Requires the ActiveDirectory module.
    - Ensure that the provided OUs match your AD structure.
    - Proper permissions are needed to modify AD objects.

.LINK
    https://learn.microsoft.com/en-us/powershell/module/activedirectory
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^OU=.+')]
    [string]$userSearchBaseOU,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^OU=.+')]
    [string]$pendingUserDeletionOU,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^OU=.+')]
    [string]$computerSearchBaseOU,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^OU=.+')]
    [string]$localAdminGroupSearchBaseOU
)

# Ensure ActiveDirectory module is available
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "Failed to import ActiveDirectory module. Ensure the module is installed and accessible."
    exit 1
}

#region Helper Functions

function Get-ADInactiveObjects {
    <#
    .SYNOPSIS
        Retrieves inactive AD objects (users or computers) based on provided filters.

    .DESCRIPTION
        Given parameters for inactivity (DaysInactive), object type (User/Computer), a SearchBase, and excluded OUs,
        this function constructs and runs an AD query to return objects that meet the inactivity criteria.

    .PARAMETER ObjectType
        The type of AD object: 'User' or 'Computer'.

    .PARAMETER SearchBase
        The OU or container from which to start searching.

    .PARAMETER DaysInactive
        Number of days to use as the threshold for inactivity.

    .PARAMETER EnabledState
        Specifies whether to filter by Enabled = $true or $false objects. 
        Can be $true, $false, or $null (no filter on enabled property).

    .PARAMETER ExcludeOUs
        An array of OUs (e.g. "OU=Admins") to exclude from the search.

    .PARAMETER PendingDeletionOnly
        If specified, search is limited to the pendingUserDeletionOU. Used for Remove-InactiveUsers function.

    .OUTPUTS
        Returns an array of AD objects (users or computers) that meet the inactivity criteria.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('User','Computer')]
        [string]$ObjectType,

        [Parameter(Mandatory=$true)]
        [string]$SearchBase,

        [Parameter(Mandatory=$true)]
        [int]$DaysInactive,

        [bool]$EnabledState = $null,

        [string[]]$ExcludeOUs = @(),

        [switch]$PendingDeletionOnly
    )

    $logonThreshold = (Get-Date).AddDays(-$DaysInactive)

    # Base filters
    # Note: We are building a string for -Filter since AD cmdlets require a string-based filter.
    # For users and computers, LastLogonDate and whenCreated logic is similar.
    # We'll exclude objects with "DO NOT DELETE" in description to avoid accidental removal.
    
    # Construct the filter:
    # Start with EnabledState if provided
    $filterParts = @()
    if ($EnabledState -eq $true) {
        $filterParts += "(Enabled -eq $true)"
    } elseif ($EnabledState -eq $false) {
        $filterParts += "(Enabled -eq $false)"
    }

    # Exclude DO NOT DELETE objects
    $filterParts += "(Description -notlike '*DO NOT DELETE*')"

    # Inactivity conditions:
    # Object is considered inactive if LastLogonDate <= $logonThreshold, or if it never logged on and
    # its whenCreated is older than $logonThreshold.
    $thresholdString = $logonThreshold.ToString("MM/dd/yyyy HH:mm:ss")
    $filterParts += "((LastLogonDate -le '$thresholdString' -and LastLogonDate -ne $null) -or (whenCreated -le '$thresholdString' -and LastLogonDate -eq $null))"

    # Exclude specific OUs
    foreach ($ou in $ExcludeOUs) {
        $filterParts += "(DistinguishedName -notlike '*$ou*')"
    }

    # Combine all filter conditions with -and
    $finalFilter = $filterParts -join ' -and '

    # Determine which cmdlet to use
    $cmdlet = if ($ObjectType -eq 'User') { 'Get-ADUser' } else { 'Get-ADComputer' }

    # If PendingDeletionOnly is set, override SearchBase with $pendingUserDeletionOU for users
    if ($PendingDeletionOnly -and $ObjectType -eq 'User') {
        $SearchBase = $pendingUserDeletionOU
    }

    # Retrieve objects
    $properties = 'cn','lastlogondate','enabled','whencreated','whenChanged','Description','DistinguishedName'
    $inactiveObjects = & $cmdlet -Filter $finalFilter -SearchBase $SearchBase -Properties $properties -ErrorAction Stop

    return $inactiveObjects
}

#endregion Helper Functions

#region Functions

function Disable-InactiveUsers {
<#
.SYNOPSIS
    Disables inactive users in Active Directory based on a specified number of inactive days.

.DESCRIPTION
    This function identifies users who have been inactive for a specified number of days and disables them.
    It optionally moves these disabled users into a pending deletion OU for future removal. Additionally,
    it updates their description with a provided prefix and the current timestamp.

.PARAMETER DaysInactive
    Number of days of inactivity before disabling a user.

.PARAMETER DescriptionPrefix
    Prefix to add to the user's description upon disabling.

.PARAMETER MoveToPendingDeletion
    If specified, disabled users will be moved to the pending deletion OU.

.EXAMPLE
    Disable-InactiveUsers -DaysInactive 30 -DescriptionPrefix "Disabled" -MoveToPendingDeletion
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$DaysInactive,

        [Parameter(Mandatory=$true)]
        [string]$DescriptionPrefix,

        [switch]$MoveToPendingDeletion
    )

    $excludeOUs = @("OU=Admins", "OU=SERVICE ACCOUNTS", "OU=ServiceAccounts", "OU=DISABLED PER IA")

    $inactiveUsers = Get-ADInactiveObjects -ObjectType User -SearchBase $userSearchBaseOU -DaysInactive $DaysInactive -EnabledState $true -ExcludeOUs $excludeOUs

    foreach ($user in $inactiveUsers) {
        try {
            Disable-ADAccount -Identity $user.DistinguishedName -Confirm:$false -ErrorAction Stop
            $currentTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            
            if ($user.Description -notlike "* for $DaysInactive days inactivity*") {
                $newDescription = "$DescriptionPrefix on $currentTime by SYSTEM for $DaysInactive days inactivity"
                Set-ADUser -Identity $user.DistinguishedName -Description $newDescription -Confirm:$false -ErrorAction Stop
            }
            
            if ($MoveToPendingDeletion) {
                Move-ADObject -Identity $user.DistinguishedName -TargetPath $pendingUserDeletionOU -Confirm:$false -ErrorAction Stop
            }
        } catch {
            Write-Error "Failed to disable/move user $($user.DistinguishedName): $($_.Exception.Message)"
        }
    }
}

function Remove-InactiveUsers {
<#
.SYNOPSIS
    Removes inactive, disabled users from Active Directory.

.DESCRIPTION
    This function looks for users who have been disabled and inactive beyond a certain threshold
    and removes them from Active Directory. It should typically be run against the 'pending deletion' OU.

.PARAMETER DaysInactive
    The number of days of inactivity before removing a disabled user.

.EXAMPLE
    Remove-InactiveUsers -DaysInactive 90
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$DaysInactive
    )

    $excludeOUs = @("OU=Admins", "OU=SERVICE ACCOUNTS", "OU=ServiceAccounts", "OU=DISABLED PER IA")
    $inactiveUsers = Get-ADInactiveObjects -ObjectType User -SearchBase $pendingUserDeletionOU -DaysInactive $DaysInactive -EnabledState $false -ExcludeOUs $excludeOUs -PendingDeletionOnly

    foreach ($user in $inactiveUsers) {
        try {
            Remove-ADObject -Identity $user.DistinguishedName -Recursive -Confirm:$false -ErrorAction Stop
        } catch {
            Write-Error "Failed to remove user $($user.DistinguishedName): $($_.Exception.Message)"
        }
    }
}

function Remove-InactiveComputers {
<#
.SYNOPSIS
    Removes inactive computers from Active Directory.

.DESCRIPTION
    This function identifies and removes computer accounts that have been inactive beyond a specified 
    number of days. It excludes certain OUs from the search.

.PARAMETER DaysInactive
    The number of days of inactivity before removing a computer account.

.EXAMPLE
    Remove-InactiveComputers -DaysInactive 180
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$DaysInactive
    )

    $excludeOUs = @("OU=Admins", "OU=SERVICE ACCOUNTS", "OU=ServiceAccounts")
    $inactiveComputers = Get-ADInactiveObjects -ObjectType Computer -SearchBase $computerSearchBaseOU -DaysInactive $DaysInactive -EnabledState $true -ExcludeOUs $excludeOUs

    foreach ($computer in $inactiveComputers) {
        try {
            Remove-ADObject -Identity $computer.DistinguishedName -Recursive -Confirm:$false -ErrorAction Stop
        } catch {
            Write-Error "Failed to remove computer $($computer.DistinguishedName): $($_.Exception.Message)"
        }
    }
}

function Remove-EmptyLocalAdminGroups {
<#
.SYNOPSIS
    Removes empty local admin groups from Active Directory.

.DESCRIPTION
    This function searches for groups in the specified local admin groups OU that have no members and removes them.

.EXAMPLE
    Remove-EmptyLocalAdminGroups
#>

    [CmdletBinding()]
    param()

    try {
        # Retrieve groups with no members in the specified OU
        $emptyLocalAdminGroups = Get-ADGroup -Filter { member -eq $null } -SearchBase $localAdminGroupSearchBaseOU -Properties DistinguishedName -ErrorAction Stop
        
        foreach ($group in $emptyLocalAdminGroups) {
            try {
                Remove-ADGroup -Identity $group.DistinguishedName -Confirm:$false -ErrorAction Stop
            } catch {
                Write-Error "Failed to remove group $($group.DistinguishedName): $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Error "Failed to query empty local admin groups: $($_.Exception.Message)"
    }
}

#endregion Functions

# Example executions (adjust as needed):
Disable-InactiveUsers -DaysInactive 45 -DescriptionPrefix "Disabled"
Disable-InactiveUsers -DaysInactive 60 -DescriptionPrefix "Disabled" -MoveToPendingDeletion
Remove-InactiveUsers -DaysInactive 90
Remove-InactiveComputers -DaysInactive 180
Remove-EmptyLocalAdminGroups
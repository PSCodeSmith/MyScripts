<#
    .SYNOPSIS
    A script to manage inactive users, computers, and empty local admin groups in Active Directory.

    .DESCRIPTION
    This script contains functions to disable, remove, and manage inactive users, computers, and empty local 
    admin groups within a given Active Directory environment. The script provides flexibility to specify inactivity 
    duration and optional actions like moving users to a pending deletion OU.

    .PARAMETER userSearchBaseOU
    Specifies the OU for user search (example: "OU=Users,DC=example,DC=com").

    .PARAMETER pendingUserDeletionOU
    Specifies the OU for pending user deletion (example: "OU=Pending Deletion,OU=Users,DC=example,DC=com").

    .PARAMETER computerSearchBaseOU
    Specifies the OU for computer search (example: "OU=Computers,DC=example,DC=com").

    .PARAMETER localAdminGroupSearchBaseOU
    Specifies the OU for local admin group search (example: "OU=Local Admin Groups,OU=Groups,DC=example,DC=com").

    .EXAMPLE
    Disable-InactiveUsers -DaysInactive 30 -DescriptionPrefix "Disabled" -MoveToPendingDeletion

    .EXAMPLE
    Remove-InactiveUsers -DaysInactive 30

    .EXAMPLE
    Remove-InactiveComputers -DaysInactive 30

    .EXAMPLE
    Remove-EmptyLocalAdminGroups

    .NOTES
    - The script requires the Active Directory module to be imported.
    - Ensure that the provided OUs are accurate as per the AD structure.
    - Appropriate permissions are required to execute the actions on the Active Directory objects.

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/addsadministration
#>
param (
    [Parameter(Mandatory=$true)]
    [ValidatePattern('^OU=.+')]
    [string]$userSearchBaseOU,

    [Parameter(Mandatory=$true)]
    [ValidatePattern('^OU=.+')]
    [string]$pendingUserDeletionOU,

    [Parameter(Mandatory=$true)]
    [ValidatePattern('^OU=.+')]
    [string]$computerSearchBaseOU,

    [Parameter(Mandatory=$true)]
    [ValidatePattern('^OU=.+')]
    [string]$localAdminGroupSearchBaseOU
)

Import-Module ActiveDirectory

function Disable-InactiveUsers {
    <#
        .SYNOPSIS
        Disables inactive users in Active Directory based on their inactivity duration.

        .DESCRIPTION
        This function disables users in Active Directory that have been inactive for a specified number of days.
        Optionally, it moves the disabled users to a pending deletion organizational unit (OU).

        .PARAMETER DaysInactive
        Number of days of inactivity before disabling a user.

        .PARAMETER DescriptionPrefix
        Prefix for the description that will be added to the disabled users.

        .PARAMETER MoveToPendingDeletion
        If specified, disabled users will be moved to the pending deletion OU.

        .EXAMPLE
        Disable-InactiveUsers -DaysInactive 30 -DescriptionPrefix "Disabled" -MoveToPendingDeletion
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [int]$DaysInactive,
        [Parameter(Mandatory=$true)]
        [string]$DescriptionPrefix,
        
        [switch]$MoveToPendingDeletion
    )

    # Calculate the threshold date based on the inactivity days.
    $logonThreshold = (Get-Date).AddDays(-$DaysInactive)

    # Define an array of OUs to exclude
    $excludeOUs = @(
        "OU=Admins", 
        "OU=SERVICE ACCOUNTS", 
        "OU=ServiceAccounts", 
        "OU=DISABLED PER IA"
    )

    # Define the basic filter for enabled users and excluding specific descriptions
    $filter = "Enabled -eq '$true'"
    $filter += " -and Description -notlike '*DO NOT DELETE*'"

    # Add conditions for the logon threshold
    $filter += " -and ( (LastLogonDate -le '$logonThreshold' -and LastLogonDate -ne '$null')"
    $filter += " -or (whencreated -le '$logonThreshold' -and LastLogonDate -eq '$null') )"

    # Loop through the OUs to exclude, appending to the filter
    foreach ($ou in $excludeOUs) {
        $filter += " -and DistinguishedName -notlike '*$ou*'"
    }

    # Query Active Directory with the assembled filter
    $inactiveUsers = Get-ADUser -Filter $filter -SearchBase $userSearchBaseOU -Properties cn, lastlogondate, enabled, whencreated, whenChanged, Description, DistinguishedName

    # Iterate through the inactive users, disable them, update their description, and optionally move them to a specific OU.
    foreach ($user in $inactiveUsers) {
        Disable-ADAccount -Identity $user
        $currentTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        if ($user.Description -notlike "* for $DaysInactive days inactivity") {
            Set-ADUser -Identity $user -Description "$DescriptionPrefix on $currentTime by SYSTEM for $DaysInactive days inactivity" -Confirm:$false
        }
        if ($MoveToPendingDeletion) {
            Move-ADObject -Identity $user -TargetPath $pendingUserDeletionOU -Confirm:$false
        }
    }
}

function Remove-InactiveUsers {
    <#
        .SYNOPSIS
        Removes inactive users in Active Directory based on their inactivity duration.

        .DESCRIPTION
        This function removes users in Active Directory that have been inactive for a specified number of days and are disabled.

        .PARAMETER DaysInactive
        Number of days of inactivity before removing a user.

        .EXAMPLE
        Remove-InactiveUsers -DaysInactive 30
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [int]$DaysInactive
    )

    # Calculate the threshold date based on the inactivity days.
    $logonThreshold = (Get-Date).AddDays(-$DaysInactive)

    # Define an array of OUs to exclude
    $excludeOUs = @(
        "OU=Admins", 
        "OU=SERVICE ACCOUNTS", 
        "OU=ServiceAccounts", 
        "OU=DISABLED PER IA"
    )

    # Define the basic filter for disabled users and excluding specific descriptions
    $filter = "Enabled -eq '$false'"
    $filter += " -and Description -notlike '*DO NOT DELETE*'"

    # Add conditions for the logon threshold
    $filter += " -and ( (LastLogonDate -le '$logonThreshold' -and LastLogonDate -ne '$null')"
    $filter += " -or (whencreated -le '$logonThreshold' -and LastLogonDate -eq '$null') )"

    # Loop through the OUs to exclude, appending to the filter
    foreach ($ou in $excludeOUs) {
        $filter += " -and DistinguishedName -notlike '*$ou*'"
    }

    # Query Active Directory with the assembled filter
    $inactiveUsers = Get-ADUser -Filter $filter -SearchBase $pendingUserDeletionOU -Properties cn, lastlogondate, enabled, whencreated, whenChanged, Description, DistinguishedName

    # Iterate through the inactive users and remove them recursively
    foreach ($user in $inactiveUsers) {
        Remove-ADObject -Identity $user -Recursive -Confirm:$false
    }
}

function Remove-InactiveComputers {
    <#
        .SYNOPSIS
        Removes inactive computers in Active Directory based on their inactivity duration.

        .DESCRIPTION
        This function removes computers in Active Directory that have been inactive for a specified number of days and are enabled.

        .PARAMETER DaysInactive
        Number of days of inactivity before removing a computer.

        .EXAMPLE
        Remove-InactiveComputers -DaysInactive 30
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [int]$DaysInactive
    )

    # Calculate the threshold date based on the inactivity days.
    $logonThreshold = (Get-Date).AddDays(-$DaysInactive)

    # Define an array of OUs to exclude
    $excludeOUs = @(
        "OU=Admins", 
        "OU=SERVICE ACCOUNTS", 
        "OU=ServiceAccounts"
    )

    # Define the basic filter for enabled computers
    $filter = "Enabled -eq '$true'"
    $filter += " -and ( (LastLogonDate -le '$logonThreshold' -and LastLogonDate -ne '$null')"
    $filter += " -or (whencreated -le '$logonThreshold' -and LastLogonDate -eq '$null') )"

    # Loop through the OUs to exclude, appending to the filter
    foreach ($ou in $excludeOUs) {
        $filter += " -and DistinguishedName -notlike '*$ou*'"
    }

    # Query Active Directory with the assembled filter
    $inactiveComputers = Get-ADComputer -Filter $filter -SearchBase $computerSearchBaseOU -Properties cn, lastlogondate, enabled, whencreated, whenChanged, Description, DistinguishedName

    # Iterate through the inactive computers and remove them recursively
    foreach ($computer in $inactiveComputers) {
        Remove-ADObject -Identity $computer -Recursive -Confirm:$false
    }
}

function Remove-EmptyLocalAdminGroups {
    <#
        .SYNOPSIS
        Removes empty local admin groups in Active Directory.

        .DESCRIPTION
        This function removes local admin groups in Active Directory that have no members.

        .EXAMPLE
        Remove-EmptyLocalAdminGroups
    #>
    [CmdletBinding()]

    # Query Active Directory for local admin groups without members
    $emptyLocalAdminGroups = Get-ADGroup -Filter {
        DistinguishedName -like "OU=Local Admin Groups" -and member -eq $null
    } -SearchBase $localAdminGroupSearchBaseOU -Properties cn, Description, DistinguishedName

    # Iterate through the empty local admin groups and remove them
    foreach ($group in $emptyLocalAdminGroups) {
        Remove-ADGroup -Identity $group -Confirm:$false
    }
}

Disable-InactiveUsers -DaysInactive 45 -DescriptionPrefix "Disabled"
Disable-InactiveUsers -DaysInactive 60 -DescriptionPrefix "Disabled" -MoveToPendingDeletion
Remove-InactiveUsers -DaysInactive 90
Remove-InactiveComputers -DaysInactive 180
Remove-EmptyLocalAdminGroups
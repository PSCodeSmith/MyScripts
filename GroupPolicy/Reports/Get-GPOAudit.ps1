<#
.SYNOPSIS
    Audits Group Policy Objects (GPOs) in an Active Directory environment.

.DESCRIPTION
    This script:
    - Collects GPO and domain information.
    - Checks for orphaned GPOs in SYSVOL.
    - Evaluates various conditions (version mismatches, permissions, ownership, etc.).
    - Generates a report with identified issues and recommended remediation steps.

    The logic is reorganized to separate data collection, analysis, and reporting phases, making the script
    more maintainable and easier to understand, while still producing the same final report.

.EXAMPLE
    .\Get-GPOAudit.ps1

    Runs the script against the current domain.

.OUTPUTS
    A structured report of identified issues and recommendations.

.NOTES
    Author: Micah (Original)
    Modified By: [Your Name]
    Always test changes in a non-production environment first.
#>

[CmdletBinding()]
param()

#region Module Import
if (-not (Get-Module ActiveDirectory -ListAvailable)) {
    Write-Error "ActiveDirectory module not available."
    exit 1
}
if (-not (Get-Module GroupPolicy -ListAvailable)) {
    Write-Error "GroupPolicy module not available."
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module GroupPolicy -ErrorAction Stop
#endregion

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'

#region Data Collection Functions
function Get-DomainInfo {
    [CmdletBinding()]
    param()
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $domainInfo = [PSCustomObject]@{
        FQDN        = $domain.Name
        NetBIOSName = (Get-ADDomain).NetBIOSName
        SID515      = (Get-ADDomain).DomainSID.Value + "-515"
    }
    return $domainInfo
}

function Get-AllGPOData {
    [CmdletBinding()]
    param(
        [string]$DomainFQDN
    )

    $allGPOs = Get-GPO -All -Domain $DomainFQDN
    $gpoDataCollection = @()

    foreach ($gpo in $allGPOs) {
        [xml]$report = Get-GPOReport -Guid $gpo.Id -ReportType XML

        $compVersions = Get-GPOVersionData $gpo "Computer"
        $userVersions = Get-GPOVersionData $gpo "User"

        # Build a single object with all relevant properties
        $gpoObj = [PSCustomObject]@{
            GPO_Name                       = $report.GPO.Name
            GPO_Guid                       = $gpo.Id
            Computer_Versions_Inconsistent = $compVersions.Inconsistent
            Computer_AD_Version            = $compVersions.AD_Version
            Computer_SYSVOL_Version        = $compVersions.SYSVOL_Version
            Computer_Content               = [boolean]$report.GPO.Computer.ExtensionData
            Computer_Enabled               = $report.GPO.Computer.Enabled
            User_Versions_Inconsistent     = $userVersions.Inconsistent
            User_AD_Version                = $userVersions.AD_Version
            User_SYSVOL_Version            = $userVersions.SYSVOL_Version
            User_Content                   = [boolean]$report.GPO.User.ExtensionData
            User_Enabled                   = $report.GPO.User.Enabled
            UnknownSIDs                    = Test-GPOUnknownSIDs $gpo
            PermissionToUser               = Test-GPOPermissionToUser $gpo
            OwnerIsUser                    = Test-GPOOwnerIsUser $gpo $Global:domainInfo.NetBIOSName
            Missing_Permissions            = Test-GPOMissingPermissions $gpo $Global:domainInfo.SID515
            Security_Filtering             = Test-GPOEmptySecurityFiltering $gpo
        }

        # Handle links
        if ([string]::IsNullOrEmpty($report.GPO.LinksTo)) {
            $gpoObj | Add-Member -NotePropertyName Linked_to -NotePropertyValue "None"
            $gpoObj | Add-Member -NotePropertyName Link_Enabled -NotePropertyValue $false
            $gpoDataCollection += $gpoObj
        } else {
            foreach ($id in $report.GPO.LinksTo) {
                $linkedObj = $gpoObj | Select-Object * # clone
                $linkedObj | Add-Member -NotePropertyName Linked_to -NotePropertyValue $id.SOMPath
                $linkedObj | Add-Member -NotePropertyName Link_Enabled -NotePropertyValue $id.Enabled
                $gpoDataCollection += $linkedObj
            }
        }
    }
    return $gpoDataCollection
}

function Get-OrphanedGPOs {
    [CmdletBinding()]
    param(
        [array]$allGroupPolicyObjects,
        [string]$Domain
    )

    $gpoGuids = $allGroupPolicyObjects.GPO_Guid
    $sysvolPath = "\\$Domain\SYSVOL\$Domain\Policies"
    $sysvolGuids = (Get-ChildItem -Path $sysvolPath -Directory -Exclude 'PolicyDefinitions').Name -replace '{|}', ''

    $orphanedFolders = Compare-Object -ReferenceObject $sysvolGuids -DifferenceObject $gpoGuids |
        Select-Object -ExpandProperty InputObject -ErrorAction SilentlyContinue

    if ($orphanedFolders) {
        $orphanedFolders | ForEach-Object {
            [PSCustomObject]@{ GPO_Name = "Orphaned Policy: \\$Domain\SYSVOL\$Domain\Policies\{$_}" }
        }
    }
}
#endregion

#region Analysis Functions
function Get-GPOVersionData {
    param(
        [Microsoft.GroupPolicy.Gpo]$GPO,
        [ValidateSet("Computer","User")]
        $Section
    )

    $DSVersion = $GPO.$Section.DSVersion
    $SYSVOLVersion = $GPO.$Section.SysvolVersion
    [PSCustomObject]@{
        AD_Version     = $DSVersion
        SYSVOL_Version = $SYSVOLVersion
        Inconsistent   = ($DSVersion -ne $SYSVOLVersion)
    }
}

function Test-GPOUnknownSIDs {
    param(
        [Microsoft.GroupPolicy.Gpo]$GPO
    )
    $unknownSIDs = $GPO.GetSecurityInfo().Trustee | Where-Object { $_.SidType -eq 'Unknown' }
    return ([bool]$unknownSIDs)
}

function Test-GPOPermissionToUser {
    param(
        [Microsoft.GroupPolicy.Gpo]$GPO
    )
    $PermissionToUser = Get-GPPermission -Guid $GPO.Id -All |
        Where-Object { $_.Permission -eq 'GpoEditDeleteModifySecurity' -and $_.Trustee.SidType -eq 'User' }
    return ([bool]$PermissionToUser)
}

function Test-GPOMissingPermissions {
    param(
        [Microsoft.GroupPolicy.Gpo]$GPO,
        [string]$DomainComputersSID
    )

    $Permissions = Get-GPPermission -Guid $GPO.Id -All
    $AuthUsers = $Permissions | Where-Object { $_.Trustee.Sid.Value -eq 'S-1-5-11' }
    $DomainComputers = $Permissions | Where-Object { $_.Trustee.Sid.Value -eq $DomainComputersSID }

    $MissingPermissions = $true
    if ($AuthUsers -and ("GPOApply","GPORead" -contains $AuthUsers.Permission) -and -not $AuthUsers.Denied) {
        $MissingPermissions = $false
    } elseif ($DomainComputers -and ("GPOApply","GPORead" -contains $DomainComputers.Permission) -and -not $DomainComputers.Denied) {
        $MissingPermissions = $false
    }
    return $MissingPermissions
}

function Test-GPOEmptySecurityFiltering {
    param(
        [Microsoft.GroupPolicy.Gpo]$GPO
    )
    $Permissions = Get-GPPermission -Guid $GPO.Id -All
    return ($Permissions | Where-Object { $_.Permission -eq "GPOApply" } | ForEach-Object { $true }) -eq $true
}

function Test-GPOOwnerIsUser {
    param(
        [Microsoft.GroupPolicy.Gpo]$GPO,
        [string]$currentDomainNetBIOS
    )
    $OwnerName = $GPO.Owner -replace "$currentDomainNetBIOS\\", ''
    $OwnerADObject = Get-ADObject -Filter { sAMAccountName -eq $OwnerName } -ErrorAction SilentlyContinue
    return ($OwnerADObject.ObjectClass -eq 'user')
}
#endregion

#region Reporting

function Add-ToReport {
    param(
        [Parameter(Mandatory=$true)]$data,
        [Parameter(Mandatory=$true)][string]$urgency,
        [Parameter(Mandatory=$true)][string]$problem,
        [Parameter(Mandatory=$true)][string]$recommendation
    )

    foreach ($item in $data) {
        [PSCustomObject]@{
            Urgency        = $urgency
            Problem        = $problem
            GPO_Name       = $item.GPO_Name
            Recommendation = $recommendation
        }
    }
}

#endregion

#region Main Execution
try {
    $Global:domainInfo = Get-DomainInfo
    $allGPOData = Get-AllGPOData -DomainFQDN $domainInfo.FQDN

    # Generate report findings
    $report = @()

    $GPOComputerSectionNotEnabled = $allGPOData | Where-Object { $_.Computer_Content -and -not $_.Computer_Enabled }
    if ($GPOComputerSectionNotEnabled) {
        $report += Add-ToReport $GPOComputerSectionNotEnabled "MEDIUM" "Computer section has content but is disabled" "Enable the Computer Section if required."
    }

    $GPOComputerSectionEnabled = $allGPOData | Where-Object { -not $_.Computer_Content -and $_.Computer_Enabled }
    if ($GPOComputerSectionEnabled) {
        $report += Add-ToReport $GPOComputerSectionEnabled "LOW" "Computer section enabled but no content" "Disable the Computer Section."
    }

    $GPOAuthenticatedUsersProblem = $allGPOData | Where-Object { $_.Missing_Permissions }
    if ($GPOAuthenticatedUsersProblem) {
        $report += Add-ToReport $GPOAuthenticatedUsersProblem "LOW" "No read permissions for Authenticated Users/Domain Computers" "Add Read permissions to Authenticated Users or Domain Computers."
    }

    $GPOOwnerIsUser = $allGPOData | Where-Object { $_.OwnerIsUser }
    if ($GPOOwnerIsUser) {
        $report += Add-ToReport $GPOOwnerIsUser "MEDIUM" "GPO owner is a user account" "Change GPO ownership to an administrative group."
    }

    $GPOUnknownSIDs = $allGPOData | Where-Object { $_.UnknownSIDs }
    if ($GPOUnknownSIDs) {
        $report += Add-ToReport $GPOUnknownSIDs "LOW" "GPO has Unknown SIDs" "Remove or replace unknown SIDs."
    }

    $GPONoContent = $allGPOData | Where-Object { -not $_.Computer_Content -and -not $_.User_Content }
    if ($GPONoContent) {
        $report += Add-ToReport $GPONoContent "LOW" "GPO has no content" "Consider deleting this GPO."
    }

    $GPOPermissionToUser = $allGPOData | Where-Object { $_.PermissionToUser }
    if ($GPOPermissionToUser) {
        $report += Add-ToReport $GPOPermissionToUser "HIGH" "GPO grants edit/delete/modify to a user" "Assign permissions to an administrative group."
    }

    $GPOAllSettingsDisabled = $allGPOData | Where-Object { -not $_.Computer_Enabled -and -not $_.User_Enabled }
    if ($GPOAllSettingsDisabled) {
        $report += Add-ToReport $GPOAllSettingsDisabled "LOW" "GPO is disabled" "Enable a section if needed or remove it."
    }

    $GPOLinkDisabled = $allGPOData | Where-Object { $_.Linked_to -ne "None" -and -not $_.Link_Enabled }
    if ($GPOLinkDisabled) {
        $report += Add-ToReport $GPOLinkDisabled "MEDIUM" "GPO link is disabled" "Enable the link if needed or remove it."
    }

    $GPOSecurityFiltering = $allGPOData | Where-Object { $_.Security_Filtering -eq $false }
    if ($GPOSecurityFiltering) {
        $report += Add-ToReport $GPOSecurityFiltering "LOW" "GPO not applied due to no security filtering" "Add a user, computer, or group to security filtering."
    }

    $GPONotLinkedToOU = $allGPOData | Where-Object { $_.Linked_to -eq "None" }
    if ($GPONotLinkedToOU) {
        $report += Add-ToReport $GPONotLinkedToOU "LOW" "GPO not linked to OU, Site, or Domain" "Link it if needed or remove it."
    }

    $GPOUserSectionNotEnabled = $allGPOData | Where-Object { $_.User_Content -and -not $_.User_Enabled }
    if ($GPOUserSectionNotEnabled) {
        $report += Add-ToReport $GPOUserSectionNotEnabled "MEDIUM" "User section has content but is disabled" "Enable the User Section if required."
    }

    $GPOUserSectionEnabled = $allGPOData | Where-Object { -not $_.User_Content -and $_.User_Enabled }
    if ($GPOUserSectionEnabled) {
        $report += Add-ToReport $GPOUserSectionEnabled "LOW" "User section enabled but no content" "Disable the User Section."
    }

    $GPOComputerVersionMismatch = $allGPOData | Where-Object { $_.Computer_Versions_Inconsistent }
    if ($GPOComputerVersionMismatch) {
        $report += Add-ToReport $GPOComputerVersionMismatch "HIGH" "Computer SYSVOL rev doesn't match AD rev" "Make a minor change and revert to sync versions."
    }

    $GPOUserVersionMismatch = $allGPOData | Where-Object { $_.User_Versions_Inconsistent }
    if ($GPOUserVersionMismatch) {
        $report += Add-ToReport $GPOUserVersionMismatch "HIGH" "User SYSVOL rev doesn't match AD rev" "Make a minor change and revert to sync versions."
    }

    $GPOPhantom = Get-OrphanedGPOs -AllGroupPolicyObjects $allGPOData -Domain $domainInfo.FQDN
    if ($GPOPhantom) {
        $report += Add-ToReport $GPOPhantom "MEDIUM" "Phantom GPO in SYSVOL only" "Investigate and remove or recreate as needed."
    }

    # Output the final report
    $report
} catch {
    Write-Error "Error encountered: $($_.Exception.Message)"
}
#endregion
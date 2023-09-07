<#
.SYNOPSIS
    This script performs an audit on Group Policy Objects (GPOs) in an Active Directory environment.

.DESCRIPTION
    The script retrieves information about GPOs, checks for various issues, and generates a report. It's designed to provide insights 
	into potential misconfigurations, security risks, and inconsistencies within your GPO settings.

.EXAMPLE
    .\Get-GPOAudit.ps1
    Runs the script without any parameters, fetching all necessary information automatically.

.EXAMPLE
    .\Get-GPOAudit.ps1
    Runs the script for the current domain.

.NOTES
    Author: Micah
    This script is intended for auditing purposes. Always back up your GPOs and test changes in a non-production environment first.

.INPUTS
    None. You do not have to provide any input to run this script.

.OUTPUTS
    The script generates a structured report containing potential issues and recommendations for each GPO.
#>


# Retrieve the fully qualified domain name (FQDN) of the current domain
$currentDomainFQDN = $([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name

# Retrieve the NetBIOS name of the current domain
$currentDomainNetBIOS = $(Get-ADDomain).NetBIOSName

# Fetch all GPOs from the current domain
$allGroupPolicyObjects = Get-GPO -All -Domain $currentDomainFQDN

#region Functions

function Get-OrphanedGPO {
	<#
	.SYNOPSIS
		This function retrieves orphaned GPO folders in SYSVOL.

	.DESCRIPTION
		Get-OrphanedGPO identifies Group Policy Object (GPO) folders that exist in the SYSVOL share
		but are not linked to any Active Directory objects. It compares the list of GPOs in Active Directory
		with the GPO folders in the SYSVOL directory and identifies the ones that are orphaned.

	.PARAMETER allGroupPolicyObjects
		An array of all Group Policy Objects in the domain.

	.PARAMETER Domain
		The fully qualified domain name (FQDN) or NetBIOS name of the domain.

	.EXAMPLE
		Get-OrphanedGPO -allGroupPolicyObjects $allGroupPolicyObjects -Domain 'mydomain.com'

		This will return a list of orphaned GPO folders in the domain 'mydomain.com'.

	.NOTES
		Author: Micah

	.INPUTS
		Array of all Group Policy Objects, Domain name as string.

	.OUTPUTS
		PSCustomObject containing orphaned folder paths.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [array]$allGroupPolicyObjects,

        [Parameter(Mandatory=$true)]
        [string]$Domain
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-OrphanedGPO"
    }

    process {
        try {
            Write-Verbose "Fetching GPO GUIDs from the AllGPOs parameter"
            $gpoGuids = $allGroupPolicyObjects | Select-Object -ExpandProperty Id

            Write-Verbose "Constructing SYSVOL path"
            $sysvolPath = "\\$Domain\SYSVOL\$Domain\Policies"

            Write-Verbose "Fetching GUIDs from the SYSVOL directory"
            $sysvolGuids = Get-ChildItem -Path $sysvolPath -Directory -Exclude 'PolicyDefinitions' |
                ForEach-Object { $_.Name -replace '{|}', '' }

            Write-Verbose "Comparing SYSVOL GUIDs with GPO GUIDs"
            $orphanedFolders = Compare-Object -ReferenceObject $sysvolGuids -DifferenceObject $gpoGuids |
                Select-Object -ExpandProperty InputObject

            Write-Verbose "Creating output object for orphaned folders"
            $orphanedFolders | ForEach-Object {
                [PSCustomObject]@{
                    Folder = "\\$Domain\SYSVOL\$Domain\Policies\{$_}"
                }
            }
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

function Get-GPOUnknownSIDs {
	<#
	.SYNOPSIS
		Retrieves unknown Security Identifiers (SIDs) from a given GPO object.

	.DESCRIPTION
		The Get-GPOUnknownSIDs function takes a Group Policy Object (GPO) as an input parameter
		and identifies any unknown SIDs in its security information. It returns a boolean value
		indicating the presence of unknown SIDs.

	.PARAMETER GPO
		The Group Policy Object (Microsoft.GroupPolicy.Gpo) from which to fetch security information.

	.EXAMPLE
		Get-GPOUnknownSIDs -GPO $someGPOObject

		Checks for unknown SIDs in the given GPO object and returns a boolean value.

	.NOTES
		Author: Micah

	.INPUTS
		Microsoft.GroupPolicy.Gpo object.

	.OUTPUTS
		Boolean value indicating the presence of unknown SIDs.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.GroupPolicy.Gpo]$GPO
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-GPOUnknownSIDs"
    }

    process {
        try {
            Write-Verbose "Fetching unknown SIDs from the given GPO object"
            $unknownSIDs = $GPO.GetSecurityInfo().Trustee | Where-Object { $_.SidType -eq 'Unknown' }

            Write-Verbose "Checking if any unknown SIDs are found"
            $unknownSIDFound = $null -ne $unknownSIDs
            $unknownSIDFound
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

function Get-GPOPermissionToUser {
	<#
	.SYNOPSIS
		Retrieves information about GPO permissions assigned to user accounts.

	.DESCRIPTION
		The Get-GPOPermissionToUser function takes a Group Policy Object (GPO) as an input parameter.
		It checks if any user accounts have been granted 'GpoEditDeleteModifySecurity' permissions on the GPO
		and returns a boolean value indicating the presence of such permissions.

	.PARAMETER GPO
		The Group Policy Object (Microsoft.GroupPolicy.Gpo) from which to fetch permission information.

	.EXAMPLE
		Get-GPOPermissionToUser -GPO $someGPOObject

		Checks for user accounts with 'GpoEditDeleteModifySecurity' permissions on the given GPO object
		and returns a boolean value.

	.NOTES
		Author: Micah

	.INPUTS
		Microsoft.GroupPolicy.Gpo object.

	.OUTPUTS
		Boolean value indicating the presence of user accounts with specified permissions.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.GroupPolicy.Gpo]$GPO
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-GPOPermissionToUser"
    }

    process {
        try {
            Write-Verbose "Fetching permissions granted to user accounts on the given GPO object"
            $PermissionToUser = Get-GPPermission -Guid $GPO.Id -All |
                Where-Object { $_.Permission -eq 'GpoEditDeleteModifySecurity' -and $_.Trustee.SidType -eq 'User' }

            Write-Verbose "Checking if any user accounts have the specified permissions"
            $PermissionToUserFound = $null -ne $PermissionToUser
            $PermissionToUserFound
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

function Get-GPOMissingPermissions {
	<#
	.SYNOPSIS
		Checks for missing permissions for specified user groups on a GPO.

	.DESCRIPTION
		The Get-GPOMissingPermissions function takes a Group Policy Object (GPO) as input and evaluates
		if the 'Authenticated Users' or 'Domain Computers' groups have the required 'GPOApply' or 'GPORead' permissions.
		It returns a boolean value indicating whether such permissions are missing.

	.PARAMETER GPO
		The Group Policy Object (Microsoft.GroupPolicy.Gpo) to check.

	.EXAMPLE
		Get-GPOMissingPermissions -GPO $someGPOObject

		Returns a boolean indicating if the required permissions are missing on the specified GPO.

	.NOTES
		Author: Micah

	.INPUTS
		Microsoft.GroupPolicy.Gpo object.

	.OUTPUTS
		Boolean value indicating the presence or absence of required permissions.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.GroupPolicy.Gpo]$GPO
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-GPOMissingPermissions"
    }

    process {
        try {
            Write-Verbose "Fetching all permissions on the given GPO object"
            $Permissions = Get-GPPermission -Guid $GPO.Id -All

            Write-Verbose "Checking for 'Authenticated Users' and 'Domain Computers' permissions"
            $GPOPermissionForAuthUsers = $Permissions | Where-Object { $_.Trustee.Sid.Value -eq 'S-1-5-11' }

            # Assuming $DomainComputersSID is defined elsewhere in the script
            $GPOPermissionForDomainComputers = $Permissions | Where-Object { $_.Trustee.Sid.Value -eq $DomainComputersSID }

            $MissingPermissions = $true

            Write-Verbose "Evaluating permissions"
            if (($GPOPermissionForAuthUsers) -or ($GPOPermissionForDomainComputers)) {
                if (("GPOApply", "GPORead" -contains $GPOPermissionForAuthUsers.Permission) -and ($false -eq $GPOPermissionForAuthUsers.Denied)) {
                    $MissingPermissions = $false
                }
                elseif ($GPOPermissionForDomainComputers) {
                    if (("GPOApply", "GPORead" -contains $GPOPermissionForDomainComputers.Permission) -and ($false -eq $GPOPermissionForDomainComputers.Denied)) {
                        $MissingPermissions = $false
                    }
                }
            }
            $MissingPermissions
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

function Get-GPOEmptySecurityFiltering {
	<#
	.SYNOPSIS
		Checks if a GPO object has empty security filtering.

	.DESCRIPTION
		The Get-GPOEmptySecurityFiltering function takes a Group Policy Object (GPO) as an input and evaluates
		if the object has security filtering set for 'GPOApply'. It returns a boolean value indicating the status
		of security filtering.

	.PARAMETER GPO
		The Group Policy Object (Microsoft.GroupPolicy.Gpo) to check.

	.EXAMPLE
		Get-GPOEmptySecurityFiltering -GPO $someGPOObject

		Returns a boolean indicating if the GPO object has empty security filtering.

	.NOTES
		Author: Micah

	.INPUTS
		Microsoft.GroupPolicy.Gpo object.

	.OUTPUTS
		Boolean value indicating the presence or absence of 'GPOApply' security filtering.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.GroupPolicy.Gpo]$GPO
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-GPOEmptySecurityFiltering"
    }

    process {
        try {
            Write-Verbose "Fetching all permissions on the given GPO object"
            $Permissions = Get-GPPermission -Guid $GPO.Id -All

            $SecurityFiltering = $false

            Write-Verbose "Evaluating security filtering for 'GPOApply'"
            foreach ($Permission in $Permissions) {
                if ("GPOApply" -eq $Permission.Permission) {
                    $SecurityFiltering = $true
                    break
                }
            }
            $SecurityFiltering
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

function Get-GPOVersionConsistency {
	<#
	.SYNOPSIS
		Checks the version consistency between AD and SYSVOL for a given GPO section.

	.DESCRIPTION
		The Get-GPOVersionConsistency function takes a Group Policy Object (GPO) and a section ('Computer' or 'User') as inputs.
		It compares the DSVersion and SysvolVersion for the specified section and returns an object indicating the consistency status.

	.PARAMETER GPO
		The Group Policy Object (Microsoft.GroupPolicy.Gpo) to check.

	.PARAMETER Section
		The section ('Computer' or 'User') for which to check version consistency.

	.EXAMPLE
		Get-GPOVersionConsistency -GPO $someGPOObject -Section 'Computer'

		Returns an object with fields for AD version, SYSVOL version, and a boolean indicating if they are inconsistent.

	.NOTES
		Author: Micah

	.INPUTS
		Microsoft.GroupPolicy.Gpo object, Section as string ('Computer' or 'User').

	.OUTPUTS
		PSCustomObject containing AD_Version, SYSVOL_Version, and Inconsistent status.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.GroupPolicy.Gpo]$GPO,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Computer', 'User')]
        [string]$Section
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-GPOVersionConsistency"
    }

    process {
        try {
            Write-Verbose "Determining DS and SYSVOL versions for the $Section section"

            # Initialize versions based on the Section parameter
            $DSVersion = $GPO.$Section.DSVersion
            $SYSVOLVersion = $GPO.$Section.SysvolVersion

            Write-Verbose "Comparing DSVersion and SYSVOLVersion"
            $CompareResult = [boolean]$(Compare-Object -ReferenceObject $DSVersion -DifferenceObject $SYSVOLVersion | Select-Object -ExpandProperty InputObject)

            Write-Verbose "Creating output object"
            [PSCustomObject]@{
                AD_Version      = $DSVersion
                SYSVOL_Version  = $SYSVOLVersion
                Inconsistent    = $CompareResult
            }
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

function Get-GPOOwner {
	<#
	.SYNOPSIS
		Determines whether the owner of a given GPO is a user account.

	.DESCRIPTION
		The Get-GPOOwner function takes a Group Policy Object (GPO) and the current domain's NetBIOS name as inputs.
		It checks if the owner of the GPO is a user account and returns a boolean value indicating the same.

	.PARAMETER GPO
		The Group Policy Object (Microsoft.GroupPolicy.Gpo) to check.

	.PARAMETER currentDomainNetBIOS
		The NetBIOS name of the current domain.

	.EXAMPLE
		Get-GPOOwner -GPO $someGPOObject -currentDomainNetBIOS 'MYDOMAIN'

		Returns a boolean value indicating if the owner of the GPO is a user account.

	.NOTES
		Author: Micah

	.INPUTS
		Microsoft.GroupPolicy.Gpo object, currentDomainNetBIOS as string.

	.OUTPUTS
		Boolean value indicating whether the owner is a user account.
	#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.GroupPolicy.Gpo]$GPO,

        [Parameter(Mandatory = $true)]
        [String]$currentDomainNetBIOS
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Write-Verbose "Initializing function Get-GPOOwner"
    }

    process {
        try {
            Write-Verbose "Extracting the owner name from the GPO object"
            $OwnerName = $($GPO.Owner -replace "$currentDomainNetBIOS\\", '')

            Write-Verbose "Fetching the corresponding AD object for the owner"
            $OwnerADObject = Get-ADObject -Filter { sAMAccountName -eq $OwnerName }

            Write-Verbose "Determining if the owner is a user account"
            $OwnerIsUser = $false
            if ($OwnerADObject.ObjectClass -eq 'User') {
                $OwnerIsUser = $true
            }
            $OwnerIsUser
        }
        catch {
            Write-Error "An error occurred: $_"
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

# Function to add items to the report with a standardized structure
function Add-ToReport ($data, $urgency, $problem, $recommendation) {
	$report += $data | Select-Object @{ N = 'Urgency'; E = { $urgency } },
									@{ N = 'Problem'; E = { $problem } },
									'GPO_Name',
									@{ N = 'Recommendation'; E = { $recommendation } }
}

#endregion Functions

#region Main

# Fetch domain information
$DomainInformation = Get-ADDomain

# Generate the SID for domain computers
$DomainComputersSID = $('{0}-515' -f $DomainInformation.DomainSID.Value)

# Initialize an array to hold the GPO version information
$GPOVersionInfo = @()

# Loop through each GPO object to gather various details
foreach ($GPO in $allGroupPolicyObjects) {
    Write-Verbose "Processing GPO: $($GPO.DisplayName)"

    # Fetch GPO report in XML format
    [xml]$GPOInfo = Get-GPOReport -ReportType XML -Guid $GPO.id

    # Get version consistency information for Computer and User sections
    $GPOComputerVersions = Get-GPOVersionConsistency -GPO $GPO -Section 'Computer'
    $GPOUserVersions = Get-GPOVersionConsistency -GPO $GPO -Section 'User'

    # Create a custom object to hold the gathered information
    $GPOData = [PSCustomObject]@{
        GPO_Name                     = $GPOInfo.GPO.Name
        Computer_Versions_Inconsistent = $GPOComputerVersions.Inconsistent
        Computer_AD_Version          = $GPOComputerVersions.AD_Version
        Computer_SYSVOL_Version      = $GPOComputerVersions.SYSVOL_Version
        Computer_Content             = [boolean]$GPOInfo.GPO.Computer.ExtensionData
        Computer_Enabled             = $GPOInfo.GPO.Computer.Enabled
        User_Versions_Inconsistent   = $GPOUserVersions.Inconsistent
        User_AD_Version              = $GPOUserVersions.AD_Version
        User_SYSVOL_Version          = $GPOUserVersions.SYSVOL_Version
        User_Content                 = [boolean]$GPOInfo.GPO.User.ExtensionData
        User_Enabled                 = $GPOInfo.GPO.User.Enabled
        UnknownSIDs                  = Get-GPOUnknownSIDs -GPO $GPO
        PermissionToUser             = Get-GPOPermissionToUser -GPO $GPO
        OwnerIsUser                  = Get-GPOOwner -GPO $GPO -currentDomainNetBIOS $currentDomainNetBIOS
        Missing_Permissions          = Get-GPOMissingPermissions -GPO $GPO
        Security_Filtering           = Get-GPOEmptySecurityFiltering -GPO $GPO
    }
    $GPOVersionInfo += $GPOData

    # Check if the GPO is linked to any SOM (Scope of Management)
    if ([string]::IsNullOrEmpty($GPOInfo.GPO.LinksTo)) {
        $GPOVersionInfo += [PSCustomObject]@{
            GPO_Name     = $GPOInfo.GPO.Name
            Linked_to    = 'None'
            Link_Enabled = 'false'
        }
    } else {
        foreach ($id in $GPOInfo.GPO.LinksTo) {
            $GPOVersionInfo += [PSCustomObject]@{
                GPO_Name     = $GPOInfo.GPO.Name
                Linked_to    = $id.SOMPath
                Link_Enabled = $id.Enabled
            }
        }
    }
}

If ($GPOVersionInfo)
{
	# Initialize an empty array to hold the report data
	$report = @()

	# Check for GPOs where the Computer section has content but is disabled
	$GPOComputerSectionNotEnabled = $GPOVersionInfo | Where-Object {
		($_.Computer_Content) -and ($_.Computer_Enabled -eq $false)
	}

	if ($GPOComputerSectionNotEnabled) {
		Add-ToReport $GPOComputerSectionNotEnabled "MEDIUM" "Group Policy COMPUTER SECTION has content but is DISABLED" "Determine if the content in the Computer section is required. IF YES - Enable the Computer Section of this GPO"
	}

	# Check for GPOs where the Computer section is enabled but has no content
	$GPOComputerSectionEnabled = $GPOVersionInfo | Where-Object {
		($_.Computer_Content -eq $false) -and ($_.Computer_Enabled -eq $true)
	}

	if ($GPOComputerSectionEnabled) {
		Add-ToReport $GPOComputerSectionEnabled "LOW" "Group Policy COMPUTER SECTION has no content but is ENABLED" "DISABLE the COMPUTER SECTION"
	}

	# Check for GPOs that do not have read permissions for either Authenticated Users or Domain Computers
	$GPOAuthenticatedUsersProblem = $GPOVersionInfo | Where-Object {
		$_.Missing_Permissions -eq $true
	}

	if ($GPOAuthenticatedUsersProblem) {
		Add-ToReport $GPOAuthenticatedUsersProblem "LOW" "Group Policy does not have at least READ PERMISSIONS for either AUTHENTICATED USERS or DOMAIN COMPUTERS" "Add the Authenticated Users group with Read Permissions on the Group Policy Object"
	}
	
	# Check for GPOs where the owner is a user account
	$GPOOwnerIsUser = $GPOVersionInfo | Where-Object {
		$_.OwnerIsUser -eq $true
	}

	if ($GPOOwnerIsUser) {
		Add-ToReport $GPOOwnerIsUser "MEDIUM" "Group Policy has an OWNER that is a USER" "Assign the GPO to an Administrative Group, not a user"
	}

	# Check for GPOs that have unknown SIDs assigned permissions
	$GPOUnknownSIDs = $GPOVersionInfo | Where-Object {
		$_.UnknownSIDs -eq $true
	}

	if ($GPOUnknownSIDs) {
		Add-ToReport $GPOUnknownSIDs "LOW" "Group Policy has an Unknown SID that is assigned permissions" "Inspect the policy for the permissions with unknown SIDs and either replace the SIDs to match the new identity, or remove the SID"
	}

	# Check for GPOs that have no content
	$GPONoContent = $GPOVersionInfo | Where-Object {
		($_.Computer_Content -eq $false) -and ($_.User_Content -eq $false)
	}

	if ($GPONoContent) {
		Add-ToReport $GPONoContent "LOW" "Group Policy has NO CONTENT" "Delete this policy, as it serves no purpose"
	}

	# Check for GPOs that have permissions to edit, modify, or delete set to a user
	$GPOPermissionToUser = $GPOVersionInfo | Where-Object {
		$_.PermissionToUser -eq $true
	}

	if ($GPOPermissionToUser) {
		Add-ToReport $GPOPermissionToUser "HIGH" "Group Policy has permissions to EDIT, MODIFY or DELETE set to a USER" "Assign the permissions to EDIT, MODIFY or DELETE to an Administrative Group, not a user"
	}

	# Check for GPOs where all settings are disabled
	$GPOAllSettingsDisabled = $GPOVersionInfo | Where-Object {
		($_.Computer_Enabled -eq $false) -and ($_.User_Enabled -eq $false)
	}

	if ($GPOAllSettingsDisabled) {
		Add-ToReport $GPOAllSettingsDisabled "LOW" "Group Policy is DISABLED" "Determine if this GPO should be enabled. IF YES - Enable appropriate section. IF NO - Delete the GPO"
	}

	# Check for GPOs linked to an OU but the link is disabled
	if ($GPOLinkDisabled = $GPOVersionInfo | Where-Object { $_.Linked_to -ne "None" -and $_.Link_Enabled -eq $false }) {
		Add-ToReport $GPOLinkDisabled "MEDIUM" "Group Policy is linked to an OU; however, the LINK is DISABLED" "Determine if the link should be enabled. IF YES - Enable the link. IF NO - Remove the link"
	}

	# Check for GPOs not being applied due to Security Filtering
	if ($GPOSecurityFiltering = $GPOVersionInfo | Where-Object { $_.Security_Filtering -eq $false }) {
		Add-ToReport $GPOSecurityFiltering "LOW" "Group Policy is not being applied to a user, computer or security group" "Determine if this policy is still needed. IF YES - Apply it to a user, computer or group in Security Filtering. IF NO - Delete the GPO"
	}

	# Check for GPOs not linked to any OU, Site, or Domain
	if ($GPONotLinkedToOU = $GPOVersionInfo | Where-Object { $_.Linked_to -eq "None" }) {
		Add-ToReport $GPONotLinkedToOU "LOW" "Group Policy is NOT LINKED to an OU, Site or Domain" "Determine if this GPO is still needed. IF YES - Link it to the appropriate OU, Site or Domain. IF NO - Delete the GPO"
	}

	# Check for GPOs where the User section has content but is disabled
	if ($GPOUserSectionNotEnabled = $GPOVersionInfo | Where-Object { $_.User_Content -and $_.User_Enabled -eq $false }) {
		Add-ToReport $GPOUserSectionNotEnabled "MEDIUM" "Group Policy USER SECTION has content but is DISABLED" "Determine if the content in the User section is required. IF YES - Enable the User Section of this GPO"
	}

	# Check for GPOs where the User section is enabled but has no content
	if ($GPOUserSectionEnabled = $GPOVersionInfo | Where-Object { $_.User_Content -eq $false -and $_.User_Enabled -eq $true }) {
		Add-ToReport $GPOUserSectionEnabled "LOW" "Group Policy USER SECTION has no content but is ENABLED" "DISABLE the USER SECTION"
	}

	# Check for GPOs where the Computer version is inconsistent
	if ($GPOComputerVersionMismatch = $GPOVersionInfo | Where-Object { $_.Computer_Versions_Inconsistent -eq $true }) {
		Add-ToReport $GPOComputerVersionMismatch "HIGH" "Group Policy's Computer SYSVOL Revision DOES NOT MATCH Active Directory Revision" "Open the GPO with the GPMC or the GPO Editor. Make a minor change and save the settings to force a replication. Open again and undo the minor change"
	}

	# Check for GPOs where the User version is inconsistent
	if ($GPOUserVersionMismatch = $GPOVersionInfo | Where-Object { $_.User_Versions_Inconsistent -eq $true }) {
		Add-ToReport $GPOUserVersionMismatch "HIGH" "Group Policy's User SYSVOL Revision DOES NOT MATCH Active Directory Revision" "Open the GPO with the GPMC or the GPO Editor. Make a minor change and save the settings to force a replication. Open again and undo the minor change"
	}

	# Check for phantom GPOs that exist in SYSVOL but not in Active Directory
	if ($GPOPhantom = Get-OrphanedGPO -AllGPOs $allGroupPolicyObjects -Domain $currentDomainFQDN) {
		Add-ToReport $GPOPhantom "MEDIUM" "Phantom Group Policy is in SYSVOL, but doesn't exist in Active Directory" "Research what these policies are by exploring the files in the folders. IF STILL NEEDED - Recreate the policies. IF NOT NEEDED - Backup and delete the folders from SYSVOL"
	}
	
	$report
}
else
{
	Write-Verbose "Unable to get the required GPO information.  Please check that you have permissions to gather the correct information, by checkin 'Get-GPO' and 'Get-GPOReport' queries."
}

#endregion Main
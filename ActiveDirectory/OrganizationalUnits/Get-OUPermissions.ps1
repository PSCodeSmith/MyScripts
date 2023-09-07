<#
.SYNOPSIS
Generates a report of Active Directory permissions for Organizational Units (OUs).

.DESCRIPTION
Fetches the ACL entries for each OU and exports them to a CSV file.

.PARAMETER None

.EXAMPLE
.\Get-OUPermissions.ps1

.NOTES
Author: Micah

.INPUTS
None

.OUTPUTS
CSV file with OU permissions.
#>

[CmdletBinding()]
param()

# Import the ActiveDirectory module if not already loaded
if (-not (Get-Module -Name 'ActiveDirectory')) {
    Import-Module ActiveDirectory
}

# Initialize report array
$report = @()

# Cache schemaIDGUIDs for better performance
$schemaIDGUID = @{}
Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID |
    ForEach-Object { $schemaIDGUID.add([System.GUID]$_.schemaIDGUID, $_.name) }

# Cache Extended-Rights GUIDs
Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID |
    ForEach-Object { $schemaIDGUID.add([System.GUID]$_.rightsGUID, $_.name) }

# Fetch Organizational Units and Domain Containers
$OUs = @(Get-ADDomain | Select-Object -ExpandProperty DistinguishedName)
$OUs += Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName
$OUs += Get-ADObject -SearchBase (Get-ADDomain).DistinguishedName -SearchScope OneLevel -LDAPFilter '(objectClass=container)' | Select-Object -ExpandProperty DistinguishedName

foreach ($OU in $OUs) {
    try {
        $acl = Get-Acl -Path "AD:\$OU" -ErrorAction Stop
        $accessEntries = $acl.Access | Select-Object @{name='OrganizationalUnit'; expression={$OU}},
                                                       @{name='ObjectTypeName'; expression={if ($_.ObjectType.ToString() -eq '00000000-0000-0000-0000-000000000000') {'All'} else {$schemaIDGUID.Item($_.ObjectType)}}},
                                                       @{name='InheritedObjectTypeName'; expression={$schemaIDGUID.Item($_.InheritedObjectType)}},
                                                       *
        $report += $accessEntries
    } catch {
        Write-Warning -Message "Failed to fetch ACL for OU: $OU. Error: $_"
    }
}

# Export the report
$csvPath = ".\OU_Permissions.csv"
$report | Export-Csv -Path $csvPath -NoTypeInformation
Start-Process $csvPath

<# Additional code to filter and sort entries
$report |
    Where-Object {-not $_.IsInherited} |
    Select-Object IdentityReference, OrganizationalUnit -Unique |
    Sort-Object IdentityReference

$filter = Read-Host "Enter the user or group name to search in OU permissions"
$report |
    Where-Object {$_.IdentityReference -like "*$filter*"} |
    Select-Object IdentityReference, OrganizationalUnit, IsInherited -Unique |
    Sort-Object IdentityReference
#>
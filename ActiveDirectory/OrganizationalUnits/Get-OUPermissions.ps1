<#
.SYNOPSIS
    Generates a report of Active Directory permissions for Organizational Units (OUs).

.DESCRIPTION
    This script enumerates all OUs and certain domain containers in an Active Directory domain, 
    retrieves their Access Control Lists (ACLs), and maps the ObjectType and InheritedObjectType 
    GUIDs to their corresponding schema names. The resulting report is exported to a CSV file for 
    easy review. A CSV file named "OU_Permissions.csv" is generated in the script's directory.

.NOTES
    Author: Micah (Original), Updated by [Your Name]
    Requires: ActiveDirectory module and appropriate permissions to read ACLs in AD.
    Ensure that the script is run from a machine with RSAT tools or an AD environment where 
    ActiveDirectory module is available.

.PARAMETER None
    There are no parameters for this script.

.EXAMPLE
    .\Get-OUPermissions.ps1

    This will generate OU_Permissions.csv in the script directory and open it automatically.

.OUTPUTS
    A CSV file named "OU_Permissions.csv" containing OU permissions.

.LINK
    https://learn.microsoft.com/en-us/powershell/module/activedirectory
#>

[CmdletBinding()]
param()

# Ensure the ActiveDirectory module is available
try {
    if (-not (Get-Module -Name 'ActiveDirectory' -ListAvailable)) {
        Throw "ActiveDirectory module not found. Install RSAT or run on a domain controller."
    }
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "Failed to load ActiveDirectory module: $($_.Exception.Message)"
    exit 1
}

Write-Verbose "Fetching schema and extended rights mappings..."

# Dictionary to cache GUID-to-name mappings for schema classes, attributes, and extended rights
$schemaIDGUID = [System.Collections.Generic.Dictionary[System.Guid, string]]::new()

try {
    # Retrieve schema objects with schemaIDGUID
    Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID |
        ForEach-Object {
            $guid = [System.Guid]$_.schemaIDGUID
            if (-not $schemaIDGUID.ContainsKey($guid)) {
                $schemaIDGUID.Add($guid, $_.name)
            }
        }

    # Retrieve Extended-Rights objects
    Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).ConfigurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID |
        ForEach-Object {
            $guid = [System.Guid]$_.rightsGUID
            if (-not $schemaIDGUID.ContainsKey($guid)) {
                $schemaIDGUID.Add($guid, $_.name)
            }
        }
} catch {
    Write-Error "Failed to populate schema/extended rights GUID mappings: $($_.Exception.Message)"
    exit 1
}

Write-Verbose "Retrieving list of OUs and containers..."

# Collect distinguished names of the domain root, OUs, and top-level containers
$domainDN = (Get-ADDomain).DistinguishedName
$OUs = New-Object System.Collections.ArrayList

[void]$OUs.Add($domainDN)

try {
    Get-ADOrganizationalUnit -Filter * | ForEach-Object { [void]$OUs.Add($_.DistinguishedName) }

    Get-ADObject -SearchBase $domainDN -SearchScope OneLevel -LDAPFilter '(objectClass=container)' |
        ForEach-Object { [void]$OUs.Add($_.DistinguishedName) }
} catch {
    Write-Warning "Error retrieving OUs/Containers: $($_.Exception.Message)"
}

Write-Verbose "Retrieving ACLs for each OU..."

$report = New-Object System.Collections.ArrayList

foreach ($OU in $OUs) {
    try {
        # Get the ACL for the OU
        $acl = Get-Acl -Path "AD:\$OU" -ErrorAction Stop
        # Map each access entry, converting GUIDs to schema names if available
        foreach ($ace in $acl.Access) {
            $objectTypeName = if ($ace.ObjectType -eq [Guid]"00000000-0000-0000-0000-000000000000") {
                "All"
            } else {
                $schemaIDGUID[$ace.ObjectType] -as [string] -or $ace.ObjectType.ToString()
            }

            $inheritedObjectTypeName = if ($ace.InheritedObjectType -eq [Guid]"00000000-0000-0000-0000-000000000000") {
                $null
            } else {
                $schemaIDGUID[$ace.InheritedObjectType] -as [string] -or $ace.InheritedObjectType.ToString()
            }

            $reportEntry = [PSCustomObject]@{
                OrganizationalUnit       = $OU
                IdentityReference        = $ace.IdentityReference
                ActiveDirectoryRight     = $ace.ActiveDirectoryRights
                InheritanceType          = $ace.InheritanceType
                IsInherited              = $ace.IsInherited
                ObjectTypeGuid           = $ace.ObjectType
                ObjectTypeName           = $objectTypeName
                InheritedObjectTypeGuid  = $ace.InheritedObjectType
                InheritedObjectTypeName  = $inheritedObjectTypeName
            }
            [void]$report.Add($reportEntry)
        }
    } catch {
        Write-Warning "Failed to fetch ACL for OU: $OU. Error: $($_.Exception.Message)"
    }
}

Write-Verbose "Exporting results to CSV..."

# Define CSV output path
$csvPath = Join-Path (Split-Path -Parent $PSCommandPath) "OU_Permissions.csv"
try {
    $report | Export-Csv -Path $csvPath -NoTypeInformation -ErrorAction Stop
    Write-Host "Report exported to $csvPath"
    Start-Process $csvPath
} catch {
    Write-Error "Failed to export CSV: $($_.Exception.Message)"
}

Write-Verbose "Done."
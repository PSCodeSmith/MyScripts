<#
.SYNOPSIS
  Imports WMI filters from MOF files located in a specified folder into the current domain.

.DESCRIPTION
  This script takes a folder path containing MOF files that define WMI filters and imports them
  into the current Active Directory domain. It retrieves the current domain, enumerates all MOF files
  in the specified folder, and for each MOF file, it creates a corresponding WMI filter object in AD.

  The WMI filter metadata (Name, Query, QueryLanguage, TargetNameSpace) is extracted from the MOF file.
  Additional metadata (Author, CreationDate) are also set during the creation process.

.PARAMETER FolderPath
  The path to the folder containing the MOF files. This parameter is mandatory.

.EXAMPLE
  .\Import-WmiFilters.ps1 -FolderPath "C:\MOF_Files"
  Imports all WMI filters defined in MOF files located under C:\MOF_Files into the current AD domain.

.NOTES
  Author: Micah
  The script requires appropriate permissions to create WMI filters in AD.
  Ensure that you have the necessary modules and permissions loaded.

.INPUTS
  FolderPath (String)

.OUTPUTS
  None
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Folder not found at path: $_"
        }
        $true
    })]
    [string]$FolderPath
)

function Import-WmiFilter {
    <#
    .SYNOPSIS
      Imports a WMI filter from a specified MOF file into the given domain.

    .DESCRIPTION
      This function reads a MOF file, extracts the WMI filter properties (Name, Query,
      QueryLanguage, TargetNameSpace), and then creates a corresponding WMI filter in AD
      under CN=WMI Filters,CN=Policies,CN=System,<domain>. It also sets metadata such as
      Author and CreationDate. After creation, it validates the presence of the WMI filter
      to ensure successful import.

    .PARAMETER Path
      The full path to the MOF file from which to import the WMI filter.

    .PARAMETER Domain
      The FQDN of the domain in which the WMI filter is to be created.

    .EXAMPLE
      Import-WmiFilter -Path "C:\example.mof" -Domain "example.com"

    .NOTES
      Author: Micah
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "MOF file not found at path: $_"
            }
            $true
        })]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    try {
        # Read the entire MOF file
        $MofContent = Get-Content -Path $Path -Raw

        # Extract required properties
        # These assume a specific structure in the MOF file. If the MOF structure changes,
        # adjustments may be needed.
        $Name           = ($MofContent -split "`n" | Select-String -Pattern 'Name =').ToString().Split('"')[1]
        $Query          = ($MofContent -split "`n" | Select-String -Pattern 'Query =').ToString().Split('"')[1]
        $QueryLanguage  = ($MofContent -split "`n" | Select-String -Pattern 'QueryLanguage =').ToString().Split('"')[1]
        $TargetNameSpace= ($MofContent -split "`n" | Select-String -Pattern 'TargetNameSpace =').ToString().Split('"')[1]

        # Metadata
        $CreationDate = (Get-Date).ToString('yyyyMMddHHmmss.ffffff')
        $Author       = $env:USERNAME

        # LDAP Path for the domain root
        $ldapPath = "LDAP://$Domain"
        $rootDSE  = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)

        # Find the Policies container
        $policiesContainer = $rootDSE.Children.Find("CN=Policies,CN=System")

        # If container for WMI filters doesn't exist, attempt to find it first
        $wmiContainer = $policiesContainer.Children | Where-Object { $_.Name -eq "CN=WMI Filters" }
        if (-not $wmiContainer) {
            Write-Error "WMI Filters container not found in domain $Domain. Ensure your AD structure is correct."
            return
        }

        # Create a unique GUID for the new WMI filter object
        $ID = [guid]::NewGuid().ToString()
        $newFilter = $wmiContainer.Children.Add("CN=$ID", "msWMI-SomFilter")

        # Set the required properties
        $newFilter.Properties["displayName"].Value        = $Name
        $newFilter.Properties["msWMI-Query"].Value        = $Query
        $newFilter.Properties["msWMI-QueryLanguage"].Value= $QueryLanguage
        $newFilter.Properties["msWMI-TargetNamespace"].Value = $TargetNameSpace
        $newFilter.Properties["msWMI-Author"].Value       = $Author
        $newFilter.Properties["msWMI-CreationDate"].Value = $CreationDate

        if ($PSCmdlet.ShouldProcess("WMI Filter '$Name' in domain '$Domain'")) {
            # Commit the changes to the directory
            $newFilter.CommitChanges()

            # Validate creation
            $searcher = New-Object System.DirectoryServices.DirectorySearcher($rootDSE)
            $searcher.Filter = "(&(objectClass=msWMI-SomFilter)(displayName=$Name))"
            $searchResult = $searcher.FindOne()

            if ($null -ne $searchResult) {
                Write-Verbose "WMI Filter '$Name' imported successfully."
            }
            else {
                Write-Error "Failed to validate the imported WMI filter: '$Name'"
            }
        }
    }
    catch {
        Write-Error "Error importing WMI filter from '$Path': $($_.Exception.Message)"
    }
}

try {
    # Determine the current domain
    $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name

    # Retrieve all .mof files from the specified folder (recursively)
    $MofFiles = Get-ChildItem -Path $FolderPath -Filter "*.mof" -Recurse -ErrorAction Stop

    foreach ($MofFile in $MofFiles) {
        if ($PSCmdlet.ShouldProcess($MofFile.FullName, "Import WMI Filter")) {
            Import-WmiFilter -Path $MofFile.FullName -Domain $CurrentDomain
        }
    }
}
catch {
    Write-Error "An error occurred during the WMI filters import process: $($_.Exception.Message)"
}
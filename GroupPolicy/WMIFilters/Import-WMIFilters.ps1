<#
.SYNOPSIS
  This script imports WMI filters from MOF files located in a specified folder into the current domain.

.DESCRIPTION
  The script first validates the folder path and retrieves the current domain name. It then iterates through
  each MOF file in the folder, calling the Import-WmiFilter function to import the WMI filter into the domain.

.PARAMETER FolderPath
  Specifies the folder path where the MOF files are located. This is a mandatory parameter.

.EXAMPLE
  .\Your_Script_Name.ps1 -FolderPath "C:\MOF_Files"
  This example imports all WMI filters from MOF files located in C:\MOF_Files into the current domain.

.NOTES
  Author: Micah

.INPUTS
  FolderPath

.OUTPUTS
  None
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-Not (Test-Path $_)) {
            throw "Folder not found at path: $_"
        }
        $true
    })]
    [string]$FolderPath
)

function Import-WmiFilter {
    <#
    .SYNOPSIS
    This function imports a WMI filter using information parsed from a MOF file.

    .DESCRIPTION
    The function reads the MOF file specified by the -Path parameter and then creates
    a new WMI filter in the specified domain.

    .PARAMETER Path
    Specifies the path to the MOF file.

    .PARAMETER Domain
    Specifies the domain where the WMI filter will be created.

    .EXAMPLE
    Import-WmiFilter -Path "C:\example.mof" -Domain "example.com"
    This example imports a WMI filter from the MOF file located at C:\example.mof into the domain example.com.

    .NOTES
    Author: Micah

    .INPUTS
    Path, Domain

    .OUTPUTS
    None
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (-Not (Test-Path $_)) {
                throw "MOF file not found at path: $_"
            }
            $true
        })]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    try {
        # Read MOF content
        $MofContent = Get-Content -Path $Path -Raw

        # Parse MOF content
        $Name = ($MofContent -split "`n" | Select-String -Pattern 'Name =').ToString().Split('"')[1]
        $Query = ($MofContent -split "`n" | Select-String -Pattern 'Query =').ToString().Split('"')[1]
        $QueryLanguage = ($MofContent -split "`n" | Select-String -Pattern 'QueryLanguage =').ToString().Split('"')[1]
        $TargetNameSpace = ($MofContent -split "`n" | Select-String -Pattern 'TargetNameSpace =').ToString().Split('"')[1]

        # Metadata
        $CreationDate = (Get-Date).ToString('yyyyMMddHHmmss.ffffff')
        $Author = $env:USERNAME

        # LDAP Operations
        $ldapPath = "LDAP://$Domain"
        $rootDSE = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
        $gpmcContainer = $rootDSE.Children.Find("CN=Policies,CN=System")
        $ID = [guid]::NewGuid().ToString()
        $wmiFilterContainer = $gpmcContainer.Children.Add("CN=$ID,CN=WMI Filters", "msWMI-SomFilter")

        # Set properties
        $wmiFilterContainer.Properties["displayName"].Value = $Name
        $wmiFilterContainer.Properties["msWMI-Query"].Value = $Query
        $wmiFilterContainer.Properties["msWMI-QueryLanguage"].Value = $QueryLanguage
        $wmiFilterContainer.Properties["msWMI-TargetNamespace"].Value = $TargetNameSpace
        $wmiFilterContainer.Properties["msWMI-Author"].Value = $Author
        $wmiFilterContainer.Properties["msWMI-CreationDate"].Value = $CreationDate

        # Commit changes
        $wmiFilterContainer.CommitChanges()

        # Validation
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($rootDSE)
        $searcher.Filter = "(&(objectClass=msWMI-SomFilter)(displayName=$Name))"
        $searchResult = $searcher.FindOne()

        if ($null -ne $searchResult) {
            Write-Verbose "WMI Filter '$Name' has been imported successfully."
        }
        else {
            Write-Error "Failed to validate the imported WMI filter: '$Name'"
        }
    }
    catch {
        Write-Error "Error importing WMI filter: $_"
    }
}

# Validate the folder path
if (-not (Test-Path -Path $FolderPath)) {
    Write-Error "Folder not found at path: $FolderPath"
    return
}

try {
    # Get the current domain name using DirectoryServices.ActiveDirectory.Domain class
    $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name

    # Fetch .mof files from the specified folder
    $MofFiles = Get-ChildItem -Path $FolderPath -Filter "*.mof" -Recurse -ErrorAction Stop

    # Process each .mof file
    foreach ($MofFile in $MofFiles) {
        if ($PSCmdlet.ShouldProcess($MofFile.FullName, "Import WMI Filter")) {
            Import-WmiFilter -Path $MofFile.FullName -Domain $CurrentDomain
        }
    }
}
catch {
    Write-Error "An error occurred: $_"
}
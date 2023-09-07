#Requires -Version 3.0

# Parameters for the script
param (
    # Domain name to query. The default is the current user's domain
    [ValidateNotNullOrEmpty()] 
    [string]$DomainName = $env:USERDOMAIN,

    # User name to look for lockouts. The default is all locked out users
    [ValidateNotNullOrEmpty()] 
    [string]$UserName = '*',

    # Start time to search event logs from. The default is the past three days
    [ValidateNotNullOrEmpty()] 
    [datetime]$StartTime = (Get-Date).AddDays(-3),

    # Credential to use for reading the security event log. The default is the current user
    [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
)

try {
    $ErrorActionPreference = 'Stop'

    # Querying the domain for the PDC emulator name
    $PdcEmulator = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((
        New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName))
    ).PdcRoleOwner.name

    Write-Verbose -Message "The PDC emulator in your forest root domain is: $PdcEmulator"
    $ErrorActionPreference = 'Continue'
}
catch {
    # Handle errors if unable to query the domain
    Write-Error -Message 'Unable to query the domain. Verify the user running this script has read access to Active Directory and try again.'
}

# Preparing parameters for Invoke-Command
$Params = @{}
If ($PSBoundParameters['Credential']) {
    $Params.Credential = $Credential
}

# Invoking the command to retrieve lockout events from the security logs
Invoke-Command -ComputerName $PdcEmulator {
    Get-WinEvent -FilterHashtable @{LogName='Security';Id=4740;StartTime=$Using:StartTime} |
    Where-Object {$_.Properties[0].Value -like "$Using:UserName"} |
    Select-Object -Property TimeCreated,
                            @{Label='UserName';Expression={$_.Properties[0].Value}},
                            @{Label='ClientName';Expression={$_.Properties[1].Value}}
} @Params | 
Select-Object -Property TimeCreated, UserName, ClientName

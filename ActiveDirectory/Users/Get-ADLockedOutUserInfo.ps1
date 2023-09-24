<#
	.SYNOPSIS
		Retrieves lockout events from the PDC emulator security logs for a specified domain.
	
	.DESCRIPTION
		This script queries the PDC emulator of the given domain to fetch lockout events.
		It filters the events based on the UserName and StartTime parameters.
	
	.PARAMETER DomainName
		Domain name to query. The default is the current user's domain.
	
	.PARAMETER UserName
		User name to look for lockouts. The default is all locked out users.
	
	.PARAMETER StartTime
		Start time to search event logs from. The default is the past three days.
	
	.PARAMETER Credential
		Credential to use for reading the security event log. The default is the current user.
	
	.EXAMPLE
		.\YourScriptName.ps1 -DomainName "YourDomain" -UserName "jdoe"
	
	.OUTPUTS
		Outputs lockout events with TimeCreated, UserName, and ClientName properties.
	
	.NOTES
		Author: Micah
	
	.INPUTS
		DomainName, UserName, StartTime, Credential
#>
param
(
	[ValidateNotNullOrEmpty()]
	[string]$DomainName = $env:USERDOMAIN,
	[ValidateNotNullOrEmpty()]
	[string]$UserName = '*',
	[ValidateNotNullOrEmpty()]
	[datetime]$StartTime = (Get-Date).AddDays(-3),
	[PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
)

try
{
	$ErrorActionPreference = 'Stop'
	
	# Query the domain for the PDC emulator name
	$PdcEmulator = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((
			New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName))
	).PdcRoleOwner.name
	
	Write-Verbose -Message "The PDC emulator in your forest root domain is: $PdcEmulator"
	$ErrorActionPreference = 'Continue'
}
catch
{
	Write-Error -Message 'Unable to query the domain. Verify the user running this script has read access to Active Directory and try again.'
}

# Preparing parameters for Invoke-Command
$InvokeParams = @{ }
if ($PSBoundParameters.ContainsKey('Credential'))
{
	$InvokeParams.Credential = $Credential
}

# Invoke the command to retrieve lockout events from the security logs
Invoke-Command -ComputerName $PdcEmulator {
	Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4740; StartTime = $Using:StartTime } |
	Where-Object { $_.Properties[0].Value -like "$Using:UserName" } |
	Select-Object -Property TimeCreated,
				  @{ Label = 'UserName'; Expression = { $_.Properties[0].Value } },
				  @{ Label = 'ClientName'; Expression = { $_.Properties[1].Value } }
} @InvokeParams |
Select-Object -Property TimeCreated, UserName, ClientName

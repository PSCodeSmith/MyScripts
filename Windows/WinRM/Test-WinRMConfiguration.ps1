<#
	.SYNOPSIS
		Checks the WinRM service status, configuration settings, firewall rules, and SSL certificate (if required).
	
	.DESCRIPTION
		This script performs an audit of the WinRM configuration on the local system. It checks the WinRM service status,
		configuration settings, firewall rules, and SSL certificate if the -CheckHttps switch is provided.
	
	.PARAMETER CheckHttps
		A switch parameter to include or exclude the HTTPS portion check.
	
	.EXAMPLE
		.\Check-WinRMConfig.ps1
		
		This will run the script without checking the SSL certificate configuration.
	
	.EXAMPLE
		.\Check-WinRMConfig.ps1 -CheckHttps
		
		This will run the script and include the SSL certificate check.
	
	.OUTPUTS
		String messages indicating the status and remediation actions for each check.
	
	.NOTES
		Author: Micah
	
	.INPUTS
		None
#>
[CmdletBinding()]
param
(
	[switch]$CheckHttps
)

#region Functions

# Function to check the WinRM service status
function Test-WinRMService
{
	Write-Output "Checking WinRM service status..."
	$winrmService = Get-Service -Name WinRM
	if ($winrmService.Status -eq 'Running')
	{
		Write-Output "PASS: WinRM service is running."
	}
	else
	{
		Write-Error "FAIL: WinRM service is not running."
		Write-Warning "Remediation: Start the WinRM service using the following command:`nStart-Service -Name WinRM"
	}
}

# Function to check the WinRM configuration settings
function Test-WinRMConfig
{
	Write-Output "Checking WinRM configuration settings..."
	$winrmConfig = Invoke-Expression -Command 'winrm get winrm/config'
	if ($winrmConfig -match "AllowUnencrypted\s+=\s+false")
	{
		Write-Output "PASS: Unencrypted communication is disabled."
	}
	else
	{
		Write-Error "FAIL: Unencrypted communication is enabled."
		Write-Warning "Remediation: Disable unencrypted communication by running:`nwinrm set winrm/config/service @{AllowUnencrypted=`"false`"}"
	}
	
	if ($winrmConfig -match "Basic\s+=\s+true")
	{
		Write-Output "PASS: Basic authentication is enabled."
	}
	else
	{
		Write-Error "FAIL: Basic authentication is disabled."
		Write-Warning "Remediation: Enable basic authentication by running:`nwinrm set winrm/config/service @{Basic=`"true`"}"
	}
}

# Function to check the firewall rules for WinRM
function Test-FirewallRules
{
	Write-Output "Checking firewall rules for WinRM..."
	# Check HTTP rule for Domain profile
	$httpRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "Windows Remote Management (HTTP-In)" -and $_.Profile -like '*Domain*' }
	if ($null -eq $httpRule)
	{
		Write-Error "FAIL: Firewall rule for WinRM over HTTP is not found in the Domain profile."
		Write-Warning "Remediation: Create the required firewall rule in the Domain profile or consult your system documentation to configure WinRM over HTTP."
	}
	elseif ($httpRule.Enabled -eq $false)
	{
		Write-Error "FAIL: Firewall rule for WinRM over HTTP is disabled in the Domain profile."
		Write-Warning "Remediation: Enable the firewall rule in the Domain profile by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTP-In)'"
	}
	else
	{
		Write-Output "PASS: Firewall rule for WinRM over HTTP is enabled in the Domain profile."
	}
	
	# Check HTTPS rule for Domain profile if applicable
	If ($CheckHttps)
	{
		$httpsRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "Windows Remote Management (HTTPS-In)" -and $_.Profile -like '*Domain*' }
		if ($null -eq $httpsRule)
		{
			Write-Error "FAIL: Firewall rule for WinRM over HTTPS is not found in the Domain profile."
			Write-Warning "Remediation: Create the required firewall rule in the Domain profile or consult your system documentation to configure WinRM over HTTPS."
		}
		elseif ($httpsRule.Enabled -eq $false)
		{
			Write-Error "FAIL: Firewall rule for WinRM over HTTPS is disabled in the Domain profile."
			Write-Warning "Remediation: Enable the firewall rule in the Domain profile by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTPS-In)'"
		}
		else
		{
			Write-Output "PASS: Firewall rule for WinRM over HTTPS is enabled in the Domain profile."
		}
	}
}

# Function to check the SSL certificate for WinRM
function Test-SSLCertificate
{
	Write-Output "Checking SSL certificate for WinRM..."
	$httpsBinding = Get-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{ Address = "*"; Transport = "HTTPS" }
	if ($null -ne $httpsBinding)
	{
		Write-Output "PASS: HTTPS binding is configured."
	}
	else
	{
		Write-Error "FAIL: HTTPS binding is not configured."
		Write-Warning "Remediation: Follow these steps to configure HTTPS binding:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}"
	}
}

#endregion Functions

Write-Output "Starting WinRM configuration check..."
Test-WinRMService
Test-WinRMConfig
Test-FirewallRules
If ($CheckHttps) { Test-SSLCertificate }
Write-Output "WinRM configuration check completed."

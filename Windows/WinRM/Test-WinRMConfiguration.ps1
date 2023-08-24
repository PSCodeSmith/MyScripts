<#
.SYNOPSIS
    This script checks the WinRM configuration on a local Windows Server.

.DESCRIPTION
    The script performs checks on the WinRM service, WinRM configuration, firewall rules, and SSL certificate if HTTPS checks are enabled.
    It provides detailed output and suggests remediation paths for any detected issues.

.PARAMETER CheckHttps
    A boolean flag to include or exclude HTTPS portion checks. Defaults to true.

.EXAMPLE
    .\Check-WinRMConfiguration.ps1 -CheckHttps $false

    This command runs the script and skips the HTTPS checks.

.NOTES
    Author: Micah
    Version: 1.1
#>

[CmdletBinding()]
param (
    # Option to include or exclude HTTPS portion check
    [bool]$CheckHttps = $True
)

# Enable verbose logging
$VerbosePreference = 'Continue'

# Function to check the WinRM service status
function Test-WinRMService {
    Write-Verbose "Checking WinRM service status..."
    try {
        $winrmService = Get-Service -Name WinRM -ErrorAction Stop
        if ($winrmService.Status -eq 'Running') {
            Write-Verbose "WinRM service is running."
        } else {
            Write-Warning "WinRM service is not running."
            Write-Host "Remediation: Start the WinRM service using the following command:`nStart-Service -Name WinRM" -ForegroundColor Red
        }
    } catch [System.ServiceProcess.TimeoutException] {
        Write-Error "Timed out while checking WinRM service."
    } catch {
        Write-Error "An unexpected error occurred while checking WinRM service: $_"
    }
}

# Function to check the WinRM configuration settings
function Test-WinRMConfig {
    Write-Verbose "Checking WinRM configuration settings..."
    $winrmConfig = winrm get winrm/config
    if ($winrmConfig -match "AllowUnencrypted\s+=\s+false") {
        Write-Verbose "Unencrypted communication is disabled."
    } else {
        Write-Warning "Unencrypted communication is enabled."
        Write-Host "Remediation: Disable unencrypted communication by running:`nwinrm set winrm/config/service @{AllowUnencrypted=`"false`"}" -ForegroundColor Red
    }

    if ($winrmConfig -match "Basic\s+=\s+true") {
        Write-Verbose "Basic authentication is enabled."
    } else {
        Write-Warning "Basic authentication is disabled."
        Write-Host "Remediation: Enable basic authentication by running:`nwinrm set winrm/config/service @{Basic=`"true`"}" -ForegroundColor Red
    }
}

# Function to check the firewall rules for WinRM
function Test-FirewallRules {
    Write-Verbose "Checking firewall rules for WinRM..."
    # Check HTTP rule
    $httpRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "Windows Remote Management (HTTP-In)" }
    if ($null -eq $httpRule) {
        Write-Warning "Firewall rule for WinRM over HTTP is not found."
        Write-Host "Remediation: Create the required firewall rule or consult your system documentation to configure WinRM over HTTP." -ForegroundColor Red
    } elseif ($httpRule.Enabled -eq 'False') {
        Write-Warning "Firewall rule for WinRM over HTTP is disabled."
        Write-Host "Remediation: Enable the firewall rule by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTP-In)'" -ForegroundColor Red
    } else {
        Write-Verbose "Firewall rule for WinRM over HTTP is enabled."
    }

    # Check HTTPS rule if applicable
    If ($CheckHttps)
    {
        $httpsRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "Windows Remote Management (HTTPS-In)" }
        if ($null -eq $httpsRule) {
            Write-Warning "Firewall rule for WinRM over HTTPS is not found."
            Write-Host "Remediation: Create the required firewall rule or consult your system documentation to configure WinRM over HTTPS." -ForegroundColor Red
        } elseif ($httpsRule.Enabled -eq 'False') {
            Write-Warning "Firewall rule for WinRM over HTTPS is disabled."
            Write-Host "Remediation: Enable the firewall rule by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTPS-In)'" -ForegroundColor Red
        } else {
            Write-Verbose "Firewall rule for WinRM over HTTPS is enabled."
        }
    }
}

# Function to check the SSL certificate for WinRM
function Test-SSLCertificate {
    Write-Verbose "Checking SSL certificate for WinRM..."
    try {
        $httpsBinding = Get-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address="*";Transport="HTTPS"} -ErrorAction Stop
        if ($null -ne $httpsBinding) {
            Write-Verbose "HTTPS binding is configured."
        } else {
            Write-Warning "HTTPS binding is not configured."
            Write-Host "Remediation: Follow these steps to configure HTTPS binding:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}" -ForegroundColor Red
        }
    } catch {
        Write-Warning "Error checking HTTPS binding. HTTPS listener may not be configured."
        Write-Host "Remediation: If you need to configure WinRM for HTTPS, follow these steps:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}" -ForegroundColor Red
    }
}

try {
    Write-Host "Starting WinRM configuration check..."
    Test-WinRMService
    Test-WinRMConfig
    Test-FirewallRules
    If ($CheckHttps) { Test-SSLCertificate }
    Write-Host "WinRM configuration check completed." -ForegroundColor Green
} catch {
    Write-Error "An unexpected error occurred: $_"
    Write-Host "Remediation: Please review the error message and consult the documentation or support forums for assistance." -ForegroundColor Red
}

<#
.SYNOPSIS
    Checks the WinRM configuration on a local Windows Server.

.DESCRIPTION
    This script validates the Windows Remote Management (WinRM) configuration, including service status, configuration settings, firewall rules, and SSL certificate if HTTPS checks are enabled.
    It provides detailed output with color-coded results and suggests remediation paths for any detected issues.

.PARAMETER CheckHttps
    A switch that determines whether to include the HTTPS portion check.
    If present, the script checks the SSL certificate for WinRM.
    If absent, the script skips the HTTPS checks.

.EXAMPLE
    .\Check-WinRMConfiguration.ps1

    This command runs the script with the default parameters, including the HTTPS checks.

.EXAMPLE
    .\Check-WinRMConfiguration.ps1 -CheckHttps

    This command runs the script and performs the HTTPS checks.

.NOTES
    Author: Micah
    Version: 1.5
    This script assumes that all required cmdlets and functions are already installed and available.

.LINK
    WinRM Documentation: https://docs.microsoft.com/en-us/windows/win32/winrm/portal
    Get-NetFirewallRule: https://docs.microsoft.com/en-us/powershell/module/netsecurity/get-netfirewallrule

.INPUTS
    None. You cannot pipe input to this script.

.OUTPUTS
    The script writes the results of the checks to the console with color-coded messages. Passes are displayed in green, failures in red, and remediation steps in yellow.

#>
[CmdletBinding()]
param (
    # Option to include or exclude HTTPS portion check
    [switch]$CheckHttps
)

# Function to check the WinRM service status
function Test-WinRMService {
    Write-Host "Checking WinRM service status..." -ForegroundColor Green
    $winrmService = Get-Service -Name WinRM
    if ($winrmService.Status -eq 'Running') {
        Write-Host "PASS: WinRM service is running." -ForegroundColor Green
    } else {
        Write-Host "FAIL: WinRM service is not running." -ForegroundColor Red
        Write-Host "Remediation: Start the WinRM service using the following command:`nStart-Service -Name WinRM" -ForegroundColor Yellow
    }
}

# Function to check the WinRM configuration settings
function Test-WinRMConfig {
    Write-Host "Checking WinRM configuration settings..." -ForegroundColor Green
    $winrmConfig = winrm get winrm/config
    if ($winrmConfig -match "AllowUnencrypted\s+=\s+false") {
        Write-Host "PASS: Unencrypted communication is disabled." -ForegroundColor Green
    } else {
        Write-Host "FAIL: Unencrypted communication is enabled." -ForegroundColor Red
        Write-Host "Remediation: Disable unencrypted communication by running:`nwinrm set winrm/config/service @{AllowUnencrypted=`"false`"}" -ForegroundColor Yellow
    }

    if ($winrmConfig -match "Basic\s+=\s+true") {
        Write-Host "PASS: Basic authentication is enabled." -ForegroundColor Green
    } else {
        Write-Host "FAIL: Basic authentication is disabled." -ForegroundColor Red
        Write-Host "Remediation: Enable basic authentication by running:`nwinrm set winrm/config/service @{Basic=`"true`"}" -ForegroundColor Yellow
    }
}

# Function to check the firewall rules for WinRM
function Test-FirewallRules {
    Write-Host "Checking firewall rules for WinRM..." -ForegroundColor Green

    # Check HTTP rule for Domain profile
    $httpRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "Windows Remote Management (HTTP-In)" -and $_.Profile -like '*Domain*' }
    if ($null -eq $httpRule) {
        Write-Host "FAIL: Firewall rule for WinRM over HTTP is not found in the Domain profile." -ForegroundColor Red
        Write-Host "Remediation: Create the required firewall rule in the Domain profile or consult your system documentation to configure WinRM over HTTP." -ForegroundColor Yellow
    } elseif ($httpRule.Enabled -eq 'False') {
        Write-Host "FAIL: Firewall rule for WinRM over HTTP is disabled in the Domain profile." -ForegroundColor Red
        Write-Host "Remediation: Enable the firewall rule in the Domain profile by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTP-In)'" -ForegroundColor Yellow
    } else {
        Write-Host "PASS: Firewall rule for WinRM over HTTP is enabled in the Domain profile." -ForegroundColor Green
    }

    # Check HTTPS rule for Domain profile if applicable
    If ($CheckHttps)
    {
        $httpsRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "Windows Remote Management (HTTPS-In)" -and $_.Profile -like '*Domain*' }
        if ($null -eq $httpsRule) {
            Write-Host "FAIL: Firewall rule for WinRM over HTTPS is not found in the Domain profile." -ForegroundColor Red
            Write-Host "Remediation: Create the required firewall rule in the Domain profile or consult your system documentation to configure WinRM over HTTPS." -ForegroundColor Yellow
        } elseif ($httpsRule.Enabled -eq 'False') {
            Write-Host "FAIL: Firewall rule for WinRM over HTTPS is disabled in the Domain profile." -ForegroundColor Red
            Write-Host "Remediation: Enable the firewall rule in the Domain profile by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTPS-In)'" -ForegroundColor Yellow
        } else {
            Write-Host "PASS: Firewall rule for WinRM over HTTPS is enabled in the Domain profile." -ForegroundColor Green
        }
    }
}

# Function to check the SSL certificate for WinRM
function Test-SSLCertificate {
    Write-Host "Checking SSL certificate for WinRM..." -ForegroundColor Green
    $httpsBinding = Get-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address="*";Transport="HTTPS"}
    if ($null -ne $httpsBinding) {
        Write-Host "PASS: HTTPS binding is configured." -ForegroundColor Green
    } else {
        Write-Host "FAIL: HTTPS binding is not configured." -ForegroundColor Red
        Write-Host "Remediation: Follow these steps to configure HTTPS binding:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}" -ForegroundColor Yellow
    }
}

Write-Host "Starting WinRM configuration check..." -ForegroundColor Green
Test-WinRMService
Test-WinRMConfig
Test-FirewallRules
If ($CheckHttps) { Test-SSLCertificate }
Write-Host "WinRM configuration check completed." -ForegroundColor Green

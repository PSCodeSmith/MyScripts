<#
.SYNOPSIS
    Performs a comprehensive audit of the local WinRM configuration, including service status, configuration settings, 
    firewall rules, and (optionally) SSL certificate configuration for HTTPS.

.DESCRIPTION
    This script checks the following aspects of the WinRM configuration on the local machine:
    - WinRM service status
    - WinRM configuration settings (AllowUnencrypted and Basic authentication)
    - Firewall rules associated with WinRM (HTTP and optionally HTTPS)
    - SSL certificate binding for WinRM over HTTPS (if -CheckHttps is specified)

    The script provides PASS/FAIL results for each check, along with recommended remediation steps if any issues are identified.

.PARAMETER CheckHttps
    When specified, also checks the HTTPS firewall rules and SSL certificate binding.

.EXAMPLE
    .\Check-WinRMConfig.ps1

    Performs all checks except the HTTPS-specific certificate configuration.

.EXAMPLE
    .\Check-WinRMConfig.ps1 -CheckHttps

    Performs all checks, including the HTTPS firewall rules and SSL certificate configuration.

.NOTES
    Author: Micah (Original)
    Revised by: [Your Name]

.INPUTS
    None

.OUTPUTS
    Informational messages, warnings, and error messages describing the configuration state and any recommended remediation.
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
param
(
    [switch]$CheckHttps
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

#region Helper Functions

function Write-Pass([string]$Message)
{
    Write-Host ("PASS: {0}" -f $Message) -ForegroundColor Green
}

function Write-Fail([string]$Message)
{
    Write-Host ("FAIL: {0}" -f $Message) -ForegroundColor Red
}

function Write-Remediation([string]$Message)
{
    Write-Warning ("Remediation: {0}" -f $Message)
}

#endregion

#region Test Functions

function Test-WinRMService
{
    Write-Verbose "Checking WinRM service status..."
    try
    {
        $winrmService = Get-Service -Name 'WinRM' -ErrorAction Stop
        if ($winrmService.Status -eq 'Running')
        {
            Write-Pass "WinRM service is running."
        }
        else
        {
            Write-Fail "WinRM service is not running."
            Write-Remediation "Run: Start-Service -Name WinRM"
        }
    }
    catch
    {
        Write-Error "Unable to retrieve WinRM service status. Ensure that the 'WinRM' service exists."
    }
}

function Test-WinRMConfig
{
    Write-Verbose "Checking WinRM configuration settings..."
    try
    {
        $winrmConfig = winrm get winrm/config 2>&1
        if ($winrmConfig -match "AllowUnencrypted\s+=\s+false")
        {
            Write-Pass "Unencrypted communication is disabled."
        }
        else
        {
            Write-Fail "Unencrypted communication is enabled."
            Write-Remediation "Disable unencrypted communication: winrm set winrm/config/service @{AllowUnencrypted=`"false`"}"
        }

        if ($winrmConfig -match "Basic\s+=\s+true")
        {
            Write-Pass "Basic authentication is enabled."
        }
        else
        {
            Write-Fail "Basic authentication is disabled."
            Write-Remediation "Enable basic authentication: winrm set winrm/config/service @{Basic=`"true`"}"
        }
    }
    catch
    {
        Write-Error "Failed to retrieve WinRM configuration. Ensure that WinRM is installed and accessible."
    }
}

function Test-FirewallRules
{
    Write-Verbose "Checking firewall rules for WinRM..."

    # Helper function to validate firewall rules
    function Check-FirewallRule([string]$DisplayName, [string]$Protocol)
    {
        $rule = Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction SilentlyContinue | 
                Where-Object { $_.Profile -like '*Domain*' }

        if (-not $rule)
        {
            Write-Fail "Firewall rule for WinRM over $Protocol is not found in the Domain profile."
            Write-Remediation "Create or enable the firewall rule for WinRM ($Protocol) in the Domain profile."
            return
        }

        if ($rule.Enabled -eq $false)
        {
            Write-Fail "Firewall rule for WinRM over $Protocol is disabled in the Domain profile."
            Write-Remediation "Enable the firewall rule: Enable-NetFirewallRule -DisplayName '$DisplayName'"
        }
        else
        {
            Write-Pass "Firewall rule for WinRM over $Protocol is enabled in the Domain profile."
        }
    }

    # Check HTTP
    Check-FirewallRule "Windows Remote Management (HTTP-In)" "HTTP"

    # Check HTTPS if requested
    if ($CheckHttps)
    {
        Check-FirewallRule "Windows Remote Management (HTTPS-In)" "HTTPS"
    }
}

function Test-SSLCertificate
{
    Write-Verbose "Checking WinRM SSL certificate configuration..."
    try
    {
        $httpsBinding = Get-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{ Address="*"; Transport="HTTPS" } -ErrorAction Stop
        if ($httpsBinding -and $httpsBinding.CertificateThumbprint -and $httpsBinding.CertificateThumbprint -ne '')
        {
            Write-Pass "HTTPS binding is configured with a certificate."
        }
        else
        {
            Write-Fail "HTTPS binding is not properly configured."
            Write-Remediation "1. Obtain a valid SSL certificate. `n2. Bind the certificate to WinRM using: 
New-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}"
        }
    }
    catch
    {
        Write-Fail "Could not retrieve HTTPS binding configuration."
        Write-Remediation "Ensure WinRM is installed and configured properly. Then run the New-WSManInstance command as described."
    }
}

#endregion

Write-Host "Starting WinRM configuration check..." -ForegroundColor Cyan
Test-WinRMService
Test-WinRMConfig
Test-FirewallRules
if ($CheckHttps) { Test-SSLCertificate }
Write-Host "WinRM configuration check completed." -ForegroundColor Cyan
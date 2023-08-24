<#
.SYNOPSIS
    This script checks the WinRM configuration on a local Windows Server.
.DESCRIPTION
    It tests all necessary settings and components required for WinRM (both HTTP and HTTPS),
    provides detailed output, and suggests remediation paths for any detected issues.
#>

[CmdletBinding()]
param (
    # The maximum length of a line in the output
    [int]$MaxLineLength = 120
)

# Enable verbose logging
$VerbosePreference = 'Continue'

# Function to check WinRM service status
function Check-WinRMService {
    $winrmService = Get-Service -Name WinRM
    if ($winrmService.Status -eq 'Running') {
        Write-Verbose "WinRM service is running."
    } else {
        Write-Warning "WinRM service is not running."
        Write-Host "Remediation: Start the WinRM service using the following command:`nStart-Service -Name WinRM"
    }
}

# Function to check WinRM configuration
function Check-WinRMConfig {
    $winrmConfig = winrm get winrm/config
    if ($winrmConfig -match "AllowUnencrypted\s+=\s+false") {
        Write-Verbose "Unencrypted communication is disabled."
    } else {
        Write-Warning "Unencrypted communication is enabled."
        Write-Host "Remediation: Disable unencrypted communication by running:`nwinrm set winrm/config/service @{AllowUnencrypted=`"false`"}"
    }
    
    if ($winrmConfig -match "Basic\s+=\s+true") {
        Write-Verbose "Basic authentication is enabled."
    } else {
        Write-Warning "Basic authentication is disabled."
        Write-Host "Remediation: Enable basic authentication by running:`nwinrm set winrm/config/service @{Basic=`"true`"}"
    }
}

# Function to check firewall rules
function Check-FirewallRules {
    $httpRule = Get-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)"
    $httpsRule = Get-NetFirewallRule -DisplayName "Windows Remote Management (HTTPS-In)"
    
    if ($httpRule.Enabled -eq 'True' -and $httpsRule.Enabled -eq 'True') {
        Write-Verbose "Firewall rules for WinRM are enabled."
    } else {
        Write-Warning "Firewall rules for WinRM are not properly configured."
        Write-Host "Remediation: Enable the firewall rules by running:`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTP-In)'`nEnable-NetFirewallRule -DisplayName 'Windows Remote Management (HTTPS-In)'"
    }
}

# Function to check SSL certificate
function Check-SSLCertificate {
    $httpsBinding = Get-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address="*";Transport="HTTPS"}
    if ($null -ne $httpsBinding) {
        Write-Verbose "HTTPS binding is configured."
    } else {
        Write-Warning "HTTPS binding is not configured."
        Write-Host "Remediation: Follow these steps to configure HTTPS binding:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}"
    }
}

# Main script execution
try {
    Write-Host "Starting WinRM configuration check..."
    Check-WinRMService
    Check-WinRMConfig
    Check-FirewallRules
    Check-SSLCertificate
    Write-Host "WinRM configuration check completed." -ForegroundColor Green
} catch {
    Write-Error "An unexpected error occurred: $_"
    Write-Host "Remediation: Please review the error message and consult the documentation or support forums for assistance."
}

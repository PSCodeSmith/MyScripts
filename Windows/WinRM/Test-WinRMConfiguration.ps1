
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

function Check-WinRMService {
    $winrmService = Get-Service -Name WinRM
    if ($winrmService.Status -eq 'Running') {
        Write-Verbose "WinRM service is running."
    } else {
        Write-Warning "WinRM service is not running."
        Write-Host "Remediation: Start the WinRM service using the following command:`nStart-Service -Name WinRM"
    }
}

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

# Revised function to check firewall rules
function Check-FirewallRules {
    $portsAndProtocols = @(
        @{ Port = 5985; Protocol = "HTTP" },
        @{ Port = 5986; Protocol = "HTTPS" }
    )

    foreach ($item in $portsAndProtocols) {
        $port = $item.Port
        $protocol = $item.Protocol
        $rule = Get-NetFirewallRule | Where-Object { $_ | Get-NetFirewallPortFilter | Where-Object { $_.LocalPort -eq $port } }

        if ($null -eq $rule) {
            Write-Warning "Firewall rule for WinRM over $protocol (port $port) is not found."
            Write-Host ("Remediation: Create a firewall rule to allow traffic on port " + $port + " for WinRM over " + $protocol + ".") -ForegroundColor Yellow
        } elseif ($rule.Enabled -eq 'False') {
            Write-Warning "Firewall rule for WinRM over $protocol (port $port) is disabled."
            Write-Host ("Remediation: Enable the firewall rule using the command:`nEnable-NetFirewallRule -Name " + $($rule.Name) + "") -ForegroundColor Yellow
        } else {
            Write-Verbose "Firewall rule for WinRM over $protocol (port $port) is found and enabled."
        }
    }
}


function Check-SSLCertificate {
    try {
        $httpsBinding = Get-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address="*";Transport="HTTPS"}
        if ($null -ne $httpsBinding) {
            Write-Verbose "HTTPS binding is configured."
        } else {
            Write-Warning "HTTPS binding is not configured."
            Write-Host "Remediation: Follow these steps to configure HTTPS binding:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}"
        }
    } catch {
        Write-Warning "Error checking HTTPS binding. HTTPS listener may not be configured."
        Write-Host "Remediation: If you need to configure WinRM for HTTPS, follow these steps:`n1. Create or import a valid SSL certificate.`n2. Bind the certificate to WinRM using the following command:`nNew-WSManInstance -ResourceURI winrm/config/Listener -SelectorSet @{Address=`"*`"; Transport=`"HTTPS`"} -ValueSet @{CertificateThumbprint=`"Your-Certificate-Thumbprint`"}"
    }
}

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

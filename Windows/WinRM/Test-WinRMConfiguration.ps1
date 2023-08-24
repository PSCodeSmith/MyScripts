
<#
.SYNOPSIS
    Checks WinRM settings on the local computer and provides remediation paths.

.DESCRIPTION
    This function performs a comprehensive check of the WinRM settings on the local computer.
    It inspects the WinRM service, listeners, firewall rules, SSL certificates, and other settings.
    Detailed remediation advice is provided for any misconfigurations detected.

.PARAMETER Verbose
    Provides detailed information about the checking process.

.EXAMPLE
    Check-WinRMSettings

    Checks the WinRM settings and provides remediation advice if necessary.

.EXAMPLE
    Check-WinRMSettings -Verbose

    Performs the same check but includes verbose output detailing the checking process.

.INPUTS
    None.

.OUTPUTS
    PSCustomObject. An object containing the status of various WinRM components and any remediation advice.
#>
function Test-WinRMConfiguration {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param ()

    Write-Verbose "Starting WinRM settings check on the local computer"

    # Initialize the remediation knowledge base
    $remediation = @()

    # Check if WinRM service is running
    try {
        $winRMServiceStatus = (Get-Service -Name WinRM -ErrorAction Stop).Status
    } catch {
        Write-Warning "An error occurred while checking the WinRM service: $_"
        return
    }
    Write-Verbose "WinRM service status: $winRMServiceStatus"

    # Check listeners (HTTP and HTTPS)
    try {
        $winRMListeners = winrm enumerate winrm/config/Listener | Out-String
$httpListener = if ($winRMListeners -match 'Transport = HTTP') { "HTTP" } else { "Not Found" }
$httpsListener = if ($winRMListeners -match 'Transport = HTTPS') { "HTTPS" } else { "Not Found" }
$listenerInfo = @($httpListener, $httpsListener) -join ', '
        $listenerInfo = $winRMListeners | ForEach-Object { $_.Transport + ' : ' + $_.Hostname }
    } catch {
        Write-Warning "An error occurred while checking the WinRM listeners: $_"
        return
    }
    Write-Verbose "WinRM listeners: $listenerInfo"

    # Check firewall rules (for HTTP and HTTPS)
    try {
        $firewallRules = Get-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)" -ErrorAction Stop
        $firewallRuleInfo = $firewallRules | ForEach-Object { $_.DisplayName + " : " + $_.Enabled }
    } catch {
        Write-Warning "An error occurred while checking the WinRM firewall rules: $_"
        return
    }
    Write-Verbose "WinRM firewall rules: $firewallRuleInfo"

    # Check for an SSL certificate (for HTTPS)
    $certificateStatus = if (Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.EnhancedKeyUsageList.FriendlyName -eq "Server Authentication" }) { "Found" } else { "Not Found" }
    if ($certificateStatus -eq "Not Found") {
        $remediation += "No suitable SSL certificate found for HTTPS:`n- Import or create a certificate for secure communication.`n- Ensure the certificate has Server Authentication as a usage type.`n- Bind the certificate to the HTTPS listener."
    }
    Write-Verbose "Certificate status for HTTPS: $certificateStatus"

    # Check client settings
    $clientTrustedHosts = (Get-WSManInstance -ResourceURI winrm/config).Client.TrustedHosts
    Write-Verbose "WinRM client trusted hosts: $clientTrustedHosts"

    # Check service settings
    $serviceAuth = (Get-WSManInstance -ResourceURI winrm/config).Service.Authentication
    Write-Verbose "WinRM service authentication methods: $($serviceAuth -join ', ')"

    # Return result as an object with remediation advice
    return [PSCustomObject]@{
        "WinRMServiceStatus"  = $winRMServiceStatus
        "WinRMListeners"      = $listenerInfo -join ', '
        "FirewallRules"       = $firewallRuleInfo -join ', '
        "CertificateStatus"   = $certificateStatus
        "ClientTrustedHosts"  = $clientTrustedHosts
        "ServiceAuthMethods"  = $serviceAuth -join ', '
        "Remediation"         = $remediation -join "`n"
    }
}

# Call the function with Verbose output
Test-WinRMConfiguration -Verbose

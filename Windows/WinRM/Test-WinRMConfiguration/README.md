
# WinRM Configuration Check Script

## Purpose

This PowerShell script is designed to validate the Windows Remote Management (WinRM) configuration on a local Windows system. It checks various aspects including:

- WinRM service status
- WinRM configuration settings
- Firewall rules for WinRM (both HTTP and HTTPS)
- SSL certificate for WinRM (if HTTPS check is enabled)

The script provides detailed output with "PASS" and "FAIL" labels, and offers remediation steps for any detected issues using warnings.

## How to Use

You can run the script by downloading it to your local machine and executing it in a PowerShell session. The script has one optional switch parameter, `CheckHttps`, which enables or disables the HTTPS checks.

### Syntax

```powershell
.\Check-WinRMConfig.ps1 [-CheckHttps]
```

### Parameters

- `-CheckHttps`: Include this switch to perform the HTTPS checks. If absent, the script skips the HTTPS checks.

## Example Usage

To run the script with the default parameters (excluding the HTTPS checks):

```powershell
.\Check-WinRMConfig.ps1
```

To run the script and perform the HTTPS checks:

```powershell
.\Check-WinRMConfig.ps1 -CheckHttps
```

## Example Output

```
Starting WinRM configuration check...
Checking WinRM service status...
PASS: WinRM service is running.
FAIL: WinRM service is not running.
WARNING: Remediation: Start the WinRM service using the following command:
Start-Service -Name WinRM
Checking WinRM configuration settings...
PASS: Unencrypted communication is disabled.
FAIL: Unencrypted communication is enabled.
WARNING: Remediation: Disable unencrypted communication by running:
winrm set winrm/config/service @{AllowUnencrypted="false"}
...
WinRM configuration check completed.
```

The script writes the results of the checks to the console with clear messages. Passes are displayed with "PASS", failures with "FAIL", and remediation steps are given as warnings.

## Links and Resources

- [WinRM Documentation](https://docs.microsoft.com/en-us/windows/win32/winrm/portal)
- [Get-NetFirewallRule](https://docs.microsoft.com/en-us/powershell/module/netsecurity/get-netfirewallrule)

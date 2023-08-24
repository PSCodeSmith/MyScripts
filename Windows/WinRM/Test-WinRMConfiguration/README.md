
# WinRM Configuration Check Script

## Purpose

This PowerShell script is designed to validate the Windows Remote Management (WinRM) configuration on a local Windows Server. It checks various aspects including:

- WinRM service status
- WinRM configuration settings
- Firewall rules for WinRM (both HTTP and HTTPS)
- SSL certificate for WinRM (if HTTPS check is enabled)

The script provides color-coded output with "PASS" and "FAIL" labels, and offers remediation steps for any detected issues.

## How to Use

You can run the script by downloading it to your local machine and executing it in a PowerShell session. The script has one optional switch parameter, `CheckHttps`, which enables or disables the HTTPS checks.

### Syntax

```powershell
.\Test-WinRMConfiguration.ps1 [-CheckHttps]
```

### Parameters

- `-CheckHttps`: Include this switch to perform the HTTPS checks. If absent, the script skips the HTTPS checks.

## Example Usage

To run the script with the default parameters (including the HTTPS checks):

```powershell
.\Test-WinRMConfiguration.ps1
```

To run the script and perform the HTTPS checks:

```powershell
.\Test-WinRMConfiguration.ps1 -CheckHttps
```

## Example Output

<pre>
<font color="green">Starting WinRM configuration check...</font>
<font color="green">Checking WinRM service status...</font>
<font color="green">PASS: WinRM service is running.</font>
<font color="green">Checking WinRM configuration settings...</font>
<font color="green">PASS: Unencrypted communication is disabled.</font>
<font color="red">FAIL: Basic authentication is disabled.</font>
<font color="yellow">Remediation: Enable basic authentication by running:
winrm set winrm/config/service @{Basic="true"}</font>
<font color="green">Checking firewall rules for WinRM...</font>
<font color="green">PASS: Firewall rule for WinRM over HTTP is enabled in the Domain profile.</font>
<font color="red">FAIL: Firewall rule for WinRM over HTTPS is not found in the Domain profile.</font>
<font color="yellow">Remediation: Create the required firewall rule in the Domain profile or consult your system documentation to configure WinRM over HTTPS.</font>
<font color="green">Checking SSL certificate for WinRM...</font>
<font color="green">PASS: HTTPS binding is configured.</font>
<font color="green">WinRM configuration check completed.</font>
</pre>

The script writes the results of the checks to the console with color-coded messages. Passes are displayed in green, failures in red, and remediation steps in yellow.

## Links and Resources

- [WinRM Documentation](https://docs.microsoft.com/en-us/windows/win32/winrm/portal)
- [Get-NetFirewallRule](https://docs.microsoft.com/en-us/powershell/module/netsecurity/get-netfirewallrule)

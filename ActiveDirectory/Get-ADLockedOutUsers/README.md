
# Get-ADLockedOutUserInfo PowerShell Script

## Description
The `Get-ADLockedOutUserInfo.ps1` PowerShell script is designed to query and retrieve information about user lockouts in an Active Directory (AD) environment. It helps administrators quickly identify locked-out users and understand the causes of the lockouts by examining event logs within a specified timeframe.

## Parameters
- **DomainName**: The domain name to query. The default is the current user's domain.
- **UserName**: The username to look for lockouts. The default is all locked-out users.
- **StartTime**: The start time to search event logs from. The default is the past three days.

## Usage
1. **Open PowerShell**: Launch a PowerShell window with administrative privileges.
2. **Navigate to the script location**: Use the `cd` command to navigate to the directory where the script is located.
3. **Run the script**: Execute the script with the desired parameters. Below are some examples:

   - To query all locked-out users in the current domain for the past three days:
     ```powershell
     .\Get-ADLockedOutUserInfo.ps1
     ```

   - To query a specific user in a specific domain:
     ```powershell
     .\Get-ADLockedOutUserInfo.ps1 -DomainName "example.com" -UserName "john.doe"
     ```

   - To query all locked-out users from a specific date:
     ```powershell
     .\Get-ADLockedOutUserInfo.ps1 -StartTime "2023-08-10"
     ```

## Requirements
- PowerShell Version 3.0 or higher
- Appropriate permissions to query the AD and event logs

## License
Please refer to the licensing agreement for usage restrictions and details.

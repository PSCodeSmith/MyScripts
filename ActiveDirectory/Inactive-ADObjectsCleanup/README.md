# Inactive AD Objects Cleanup Script

## Overview
This PowerShell script is designed to manage inactive users, computers, and empty local admin groups in an Active Directory environment. It provides functionality to disable, remove, and manage these objects based on specified inactivity duration, with optional actions like moving users to a pending deletion organizational unit (OU).

## Prerequisites
- PowerShell version 5.1 or higher
- Necessary permissions to run scripts that modify Active Directory objects

## Parameters
Please refer to the script's inline documentation using `Get-Help ./Inactive-ADObjectsCleanup.ps1 -Full` for detailed information on each parameter.


## Usage
### Example 1: Basic Usage
```powershell
./Inactive-ADObjectsCleanup.ps1 -userSearchBaseOU "OU=Users,DC=example,DC=com" -pendingUserDeletionOU "OU=Pending,DC=example,DC=com" -computerSearchBaseOU "OU=Computers,DC=example,DC=com" -localAdminGroupSearchBaseOU "OU=AdminGroups,DC=example,DC=com"
```
Replace the parameters with appropriate values based on your environment.

### Example 2: Advanced Usage
```powershell
./Inactive-ADObjectsCleanup.ps1 -userSearchBaseOU "OU=Users,DC=example,DC=com" -pendingUserDeletionOU "OU=Pending,DC=example,DC=com" -computerSearchBaseOU "OU=Computers,DC=example,DC=com" -localAdminGroupSearchBaseOU "OU=AdminGroups,DC=example,DC=com" -inactivityDuration 60
```
Replace the parameters with appropriate values based on your environment.

## Verbose Logging
To enable verbose logging, use the `-Verbose` switch when running the script.

## Error Handling
The script includes error handling with helpful messages. If you encounter any issues, refer to the error messages and consult the script's documentation.

## Contribution
Feel free to contribute to this script by submitting pull requests or raising issues on GitHub.

## License
Please provide licensing information if applicable.

---
_Last Updated: 2023-08-18_

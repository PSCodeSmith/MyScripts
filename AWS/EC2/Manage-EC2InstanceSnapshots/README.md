
# Manage-EC2InstanceSnapshots.ps1

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Create New Snapshots](#create-new-snapshots)
  - [Restore Instances from Snapshots](#restore-instances-from-snapshots)
- [Function Details](#function-details)
  - [Restore-EC2InstanceFromSnapshots](#restore-ec2instancefromsnapshots)
  - [New-EC2InstanceVolumeSnapshots](#new-ec2instancevolumesnapshots)
- [License](#license)
- [Support](#support)
- [Download](#download)

## Overview

`Manage-EC2InstanceSnapshots.ps1` is a robust PowerShell script to manage Amazon EC2 instance snapshots within the AWS environment. It provides functionalities to create snapshots for EC2 instances' volumes and restore instances from those snapshots, allowing for easy backup and recovery.

## Features

- **Create Snapshots**: Easily create new snapshots for EC2 instances' volumes.
- **Restore Instances**: Restore EC2 instances from existing snapshots, with comprehensive error handling.
- **Customizable Tagging**: Utilize custom tags to associate volume and snapshot pairs.
- **Verbose Logging**: Detailed logging for better visibility and troubleshooting.

## Prerequisites

- PowerShell 5.1 or higher
- AWS PowerShell module
- AWS account with necessary permissions for EC2 instances, volumes, and snapshots

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/PSCodeSmith/MyScripts.git
   ```

2. Navigate to the directory containing the script:

   ```bash
   cd MyScripts/AWS/EC2/Manage-EC2InstanceSnapshots
   ```

3. Ensure that the AWS PowerShell module is installed:

   ```powershell
   Install-Module -Name AWSPowerShell -Force -SkipPublisherCheck
   ```

## Usage

### Create New Snapshots

```powershell
.\Manage-EC2InstanceSnapshots.ps1 -CreateSnapshots -InstanceIds i-12345678, i-23456789
```

### Restore Instances from Snapshots

```powershell
.\Manage-EC2InstanceSnapshots.ps1 -RestoreSnapshots -InstanceIds i-12345678, i-23456789 -VolumeSnapshotPairTagKey "CustomTag"
```

## Function Details

### Restore-EC2InstanceFromSnapshots

This function restores EC2 instances from existing snapshots. It handles volume detachment, new volume creation from snapshots, attachment, and instance state management.

### New-EC2InstanceVolumeSnapshots

This function creates new snapshots for EC2 instance volumes and assigns custom tags to associate volumes with snapshots.

## License

This project is licensed under the MIT License. See the LICENSE.md file for details.

## Support

For support or questions, please open an issue on the GitHub repository.

## Download

[Download README](https://github.com/PSCodeSmith/MyScripts/AWS/EC2/Manage-EC2InstanceSnapshots/README.md)


# Copy-SplitArchiveToS3.ps1

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Parameters](#parameters)
- [Usage](#usage)
- [Example](#example)
- [Supported Formats](#supported-formats)
- [Error Handling](#error-handling)
- [Links](#links)
- [License](#license)
- [Support](#support)

## Overview

`Copy-SplitArchiveToS3.ps1` is a PowerShell script designed to efficiently transfer split archive files created with 7-Zip to an Amazon S3 bucket. It supports parallel processing and integrates with either the AWS PowerShell Module or the AWS CLI, adapting to what's installed on the system.

## Prerequisites

- PowerShell 5.1 or higher
- AWS PowerShell Module or AWS CLI
- AWS account with necessary permissions to write objects to the specified S3 bucket
- Supported 7-Zip formats include 7z, zip, rar, gz, tar, and bz2

## Parameters

- **BucketName**: The name of the S3 bucket where the files will be copied. (Mandatory)
- **LocalPath**: The local directory path where the split archive files created with 7-Zip are located. (Mandatory)
- **MaxParallelFiles**: The maximum number of files to be transferred in parallel. The default value is 5.

## Usage

The script can be run from a PowerShell command line, specifying the required parameters. Ensure that the IAM user has the necessary permissions to write objects to the specified S3 bucket.

## Example

```powershell
.\Copy-SplitArchiveToS3.ps1 -BucketName 'your-bucket-name' -LocalPath 'C:\path\to\archive\files'
```

This command copies all split archive files from the specified local directory to the specified S3 bucket, processing up to 5 files in parallel.

## Supported Formats

The script supports the following 7-Zip formats:

- 7z
- zip
- rar
- gz
- tar
- bz2

## Error Handling

The script includes error handling and will provide verbose and error messages indicating the progress and status of the copy operation. It checks for the availability of the AWS PowerShell Module or AWS CLI and proceeds accordingly.

## Links

- [AWS PowerShell](https://aws.amazon.com/powershell/)
- [AWS CLI](https://aws.amazon.com/cli/)

## License

This project is licensed under the MIT License. See the LICENSE.md file for details.

## Support

For support or questions, please open an issue on the GitHub repository.

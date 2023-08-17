<#
.SYNOPSIS
    Copies split archive files created with 7-Zip to an S3 bucket in parallel.

.DESCRIPTION
    The Copy-SplitArchiveToS3 script is designed to efficiently transfer split archive files created with 7-Zip to an Amazon S3 bucket.
    It supports parallel processing and can utilize either the AWS PowerShell Module or the AWS CLI, depending on what's installed on the system.

.PARAMETER BucketName
    Specifies the name of the S3 bucket where the files will be copied. This parameter is mandatory.

.PARAMETER LocalPath
    Specifies the local directory path where the split archive files created with 7-Zip are located. This parameter is mandatory.

.PARAMETER MaxParallelFiles
    Specifies the maximum number of files to be transferred in parallel. The default value is 5.

.INPUTS
    None. You cannot pipe input to this script.

.OUTPUTS
    Verbose and error messages indicating the progress and status of the copy operation.

.EXAMPLE
    .\Copy-SplitArchiveToS3.ps1 -BucketName 'your-bucket-name' -LocalPath 'C:\path\to\archive\files'

    Copies all split archive files from the specified local directory to the specified S3 bucket, processing up to 5 files in parallel.

.NOTES
    - The AWS access key, secret key, and region must be configured as environment variables.
    - The script checks for the availability of the AWS PowerShell Module or AWS CLI and proceeds accordingly.
    - Supported 7-Zip formats include 7z, zip, rar, gz, tar, and bz2.
    - Ensure that the IAM user has the necessary permissions to write objects to the specified S3 bucket.

.LINK
    https://aws.amazon.com/powershell/
    https://aws.amazon.com/cli/

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$BucketName,

    [Parameter(Mandatory=$true)]
    [string]$LocalPath,

    [int]$MaxParallelFiles = 5
)

# Check if the AWS PowerShell Module is installed
$usingAWSPowerShell = $false
$usingAWSCLI = $false

if (Get-Module -ListAvailable -Name 'AWSPowerShell') {
    $usingAWSPowerShell = $true
    Import-Module AWSPowerShell
}
elseif (Get-Command aws -ErrorAction SilentlyContinue) {
    $usingAWSCLI = $true
}
else {
    Write-Error "Neither AWS PowerShell Module nor AWS CLI is installed. Please install one of them and try again."
    return
}

# Get the list of files to be copied (considering 7-Zip supported formats)
$archiveExtensions = @("*.7z.*", "*.zip.*", "*.rar.*", "*.gz.*", "*.tar.*", "*.bz2.*")
$files = $archiveExtensions | ForEach-Object { Get-ChildItem -Path $LocalPath -Filter $_ }

# Function to copy file to S3 using PowerShell Module
function CopyFileToS3_PowerShell {
    param (
        [System.IO.FileInfo]$file
    )

    Write-Verbose "Copying $($file.Name) to S3 bucket $BucketName"

    try {
        Write-S3Object -BucketName $BucketName -File $file.FullName -Key $file.Name
        Write-Verbose "Successfully copied $($file.Name) to S3 bucket $BucketName"
    }
    catch {
        Write-Error "Failed to copy $($file.Name) to S3. Error: $_"
    }
}

# Function to copy file to S3 using AWS CLI
function CopyFileToS3_CLI {
    param (
        [System.IO.FileInfo]$file
    )

    Write-Verbose "Copying $($file.Name) to S3 bucket $BucketName"

    try {
        aws s3 cp $file.FullName s3://$BucketName/$($file.Name)
        Write-Verbose "Successfully copied $($file.Name) to S3 bucket $BucketName"
    }
    catch {
        Write-Error "Failed to copy $($file.Name) to S3. Error: $_"
    }
}

# Copy files in parallel using the appropriate method
$files | ForEach-Object -Parallel {
    if ($usingAWSPowerShell) {
        CopyFileToS3_PowerShell -file $_
    }
    elseif ($usingAWSCLI) {
        CopyFileToS3_CLI -file $_
    }
} -ThrottleLimit $MaxParallelFiles

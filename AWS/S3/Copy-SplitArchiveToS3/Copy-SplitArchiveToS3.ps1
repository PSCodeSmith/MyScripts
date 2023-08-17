<#
.SYNOPSIS
    Copies split archive files created with 7-Zip to an S3 bucket in parallel.

.DESCRIPTION
    The Copy-SplitArchiveToS3 script copies split archive files created with 7-Zip to an Amazon S3 bucket in parallel. It supports various archive formats and can utilize either the AWS PowerShell Module or AWS CLI.

.PARAMETER BucketName
    Specifies the name of the S3 bucket where the files will be copied. This parameter is mandatory.

.PARAMETER LocalPath
    Specifies the local directory path where the split archive files created with 7-Zip are located. This parameter is mandatory.

.PARAMETER MaxParallelFiles
    Specifies the maximum number of files to be transferred in parallel. The default value is 5.

.PARAMETER Prefix
    Specifies an optional prefix (folder path) within the S3 bucket where the files will be copied. If provided, the script ensures that it ends with a backslash. This parameter is optional.

.NOTES
    - AWS credentials configured as environment variables are required.
    - AWS PowerShell Module or AWS CLI must be installed.
    - Supported 7-Zip formats include 7z, zip, rar, gz, tar, and bz2.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$BucketName,

    [Parameter(Mandatory=$true)]
    [string]$LocalPath,

    [int]$MaxParallelFiles = 5,

    [string]$Prefix = ""
)

# Ensure that Prefix ends with a backslash
if ($Prefix -ne "" -and -not $Prefix.EndsWith("/")) {
    $Prefix += "/"
}

# Ensure that LocalPath ends with a backslash and wildcard
$localPathWithWildcard = $LocalPath
if (-not $localPathWithWildcard.EndsWith("\*")) {
    if (-not $localPathWithWildcard.EndsWith("\")) {
        $localPathWithWildcard += "\"
    }
    $localPathWithWildcard += "*"
}

# Include patterns to capture all split files
$includePatterns = @('*.7z.*', '*.zip.*', '*.rar.*', '*.gz.*', '*.tar.*', '*.bz2.*')

# Get the list of split archive files created with 7-Zip formats
$archiveFiles = Get-ChildItem -Path $localPathWithWildcard -Include $includePatterns -Recurse

# Inform the user about the number of files to be transferred
Write-Output "Found $($archiveFiles.Count) files. Beginning to transfer files to S3 bucket $BucketName."

# Check for AWS PowerShell Module or AWS CLI
if (-not (Get-Module -ListAvailable -Name 'AWSPowerShell')) {
    if (-not (Get-Command aws -ErrorAction SilentlyContinue)) {
        Write-Error "Neither AWS PowerShell Module nor AWS CLI is installed. Please install one of them to proceed."
        return
    }
}

# Function to copy file to S3 using PowerShell Module
function CopyFileToS3_PowerShell {
    param (
        [System.IO.FileInfo]$file
    )

    $key = $Prefix + $file.Name

    Write-Verbose "Copying $($file.Name) to S3 bucket $BucketName with prefix $Prefix"

    try {
        Write-S3Object -BucketName $BucketName -File $file.FullName -Key $key
        Write-Verbose "Successfully copied $($file.Name) to S3 bucket $BucketName with prefix $Prefix"
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

    $key = $Prefix + $file.Name

    Write-Verbose "Copying $($file.Name) to S3 bucket $BucketName with prefix $Prefix"

    try {
        aws s3 cp $file.FullName s3://$BucketName/$key
        Write-Verbose "Successfully copied $($file.Name) to S3 bucket $BucketName with prefix $Prefix"
    }
    catch {
        Write-Error "Failed to copy $($file.Name) to S3. Error: $_"
    }
}

# Check and process files in parallel
$runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxParallelFiles)
$runspacePool.Open()

$runspaces = @()

foreach ($file in $archiveFiles) {
    $runspace = [powershell]::Create().AddScript({
        if (Get-Module -ListAvailable -Name 'AWSPowerShell') {
            CopyFileToS3_PowerShell -file $using:file
        }
        else {
            CopyFileToS3_CLI -file $using:file
        }
    })

    $runspace.RunspacePool = $runspacePool
    $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
}

# Wait for all runspaces to complete
$runspaces | ForEach-Object {
    $_.Pipe.EndInvoke($_.Status)
    $_.Pipe.Dispose()
}

$runspacePool.Close()
$runspacePool.Dispose()

# Inform the user about the number of files transferred
Write-Output "Completed copying $($archiveFiles.Count) files to S3 bucket $BucketName."

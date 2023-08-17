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
if (-not (Get-Command aws -ErrorAction SilentlyContinue)) {
    if (-not (Get-Module -ListAvailable -Name 'AWSPowerShell')) {
        Write-Error "Neither AWS PowerShell Module nor AWS CLI is installed. Please install one of them to proceed."
        return
    }
}

# Create a script block for the job
$jobScript = {
    param (
        $file,
        $BucketName,
        $Prefix
    )

    $key = $Prefix + $file.Name

    try {
        if (Get-Command aws -ErrorAction SilentlyContinue) {
            aws s3 cp $file.FullName s3://$BucketName/$key
        }
        elseif (Get-Module -ListAvailable -Name 'AWSPowerShell') {
            Write-S3Object -BucketName $BucketName -File $file.FullName -Key $key
        }
    }
    catch {
        Write-Error "Failed to copy $($file.Name) to S3. Error: $_"
    }
}

# Start jobs in parallel
$jobs = @()
foreach ($file in $archiveFiles) {
    while (@(Get-Job | Where-Object { $_.State -eq 'Running' }).Count -ge $MaxParallelFiles) {
        Start-Sleep -Milliseconds 100
    }

    Write-Output "Copying $($file.Name) to S3 bucket $BucketName/$Prefix"
    $jobs += Start-Job -ScriptBlock $jobScript -ArgumentList $file, $BucketName, $Prefix
}

# Wait for all jobs to complete
$jobs | Wait-Job

# Receive job output and remove jobs
$jobs | ForEach-Object {
    Receive-Job -Job $_
    Remove-Job -Job $_
}

# Inform the user about the number of files transferred
Write-Output "Completed copying $($archiveFiles.Count) files to S3 bucket $BucketName."

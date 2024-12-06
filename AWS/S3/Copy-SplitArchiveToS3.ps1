<#
    .SYNOPSIS
        Copies split archive files created with 7-Zip to an S3 bucket in parallel.

    .DESCRIPTION
        The Copy-SplitArchiveToS3 function copies split archive files created with 7-Zip to an Amazon S3 bucket in parallel.
        It supports various archive formats and can utilize either the AWS PowerShell Module or AWS CLI, depending on what
        is available in the environment.

    .PARAMETER BucketName
        Specifies the name of the S3 bucket where the files will be copied. This parameter is mandatory.

    .PARAMETER LocalPath
        Specifies the local directory path where the split archive files created with 7-Zip are located. This parameter is mandatory.

    .PARAMETER MaxParallelFiles
        Specifies the maximum number of files to be transferred in parallel. The default value is 5.

    .PARAMETER Prefix
        Specifies an optional prefix (folder path) within the S3 bucket where the files will be copied. If provided, the script ensures
        that it ends with a slash. This parameter is optional.

    .EXAMPLE
        Copy-SplitArchiveToS3 -BucketName 'my-bucket' -LocalPath 'C:\ArchiveFiles'
        
        This example copies all recognized split archive files from 'C:\ArchiveFiles' to the 'my-bucket' S3 bucket
        with the default parallelization level of 5.

    .EXAMPLE
        Copy-SplitArchiveToS3 -BucketName 'my-bucket' -LocalPath 'C:\ArchiveFiles' -MaxParallelFiles 10
        
        This example copies all recognized split archive files from 'C:\ArchiveFiles' to the 'my-bucket' S3 bucket
        with a parallelization level of 10.

    .EXAMPLE
        Copy-SplitArchiveToS3 -BucketName 'my-bucket' -LocalPath 'C:\ArchiveFiles' -Prefix 'archives/'
        
        This example copies all recognized split archive files from 'C:\ArchiveFiles' to the 'my-bucket' S3 bucket
        inside the 'archives/' folder.

    .NOTES
        - AWS credentials configured as environment variables or via a credentials profile are required.
        - Either the AWS PowerShell Module (AWSPowerShell) or AWS CLI must be installed and accessible.
        - Supported 7-Zip formats include 7z, zip, rar, gz, tar, and bz2.
#>
function Copy-SplitArchiveToS3 {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$BucketName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$LocalPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int]$MaxParallelFiles = 5,

        [Parameter(Mandatory = $false)]
        [string]$Prefix = ""
    )

    # Ensure Prefix ends with a forward slash if provided
    if ($Prefix -and -not $Prefix.EndsWith("/")) {
        $Prefix += "/"
    }

    # Construct a wildcard path to list files
    # Ensure that LocalPath ends with a backslash before adding "*"
    if (-not $LocalPath.EndsWith("\")) {
        $LocalPath += "\"
    }
    $LocalPathWithWildcard = Join-Path $LocalPath "*"

    # Define the supported split archive patterns
    $includePatterns = @('*.7z.*', '*.zip.*', '*.rar.*', '*.gz.*', '*.tar.*', '*.bz2.*')

    # Retrieve the files that match the given patterns
    $archiveFiles = Get-ChildItem -Path $LocalPathWithWildcard -Include $includePatterns -Recurse -ErrorAction Stop

    if ($archiveFiles.Count -eq 0) {
        Write-Verbose "No matching split archive files found in '$LocalPath'. Nothing to copy."
        return
    }

    Write-Verbose "Found $($archiveFiles.Count) files to transfer."

    # Check for AWS CLI or AWS PowerShell Module availability
    $awsCLI = Get-Command 'aws' -ErrorAction SilentlyContinue
    $awsModule = Get-Module -ListAvailable -Name 'AWSPowerShell'

    if (-not $awsCLI -and -not $awsModule) {
        Write-Error "Neither AWS CLI nor AWSPowerShell Module is installed. Please install one before proceeding."
        return
    }

    # Determine which tool to use
    $useAWSCLI = $false
    if ($awsCLI) { $useAWSCLI = $true }

    # The script block to copy files to S3
    $jobScript = {
        param (
            [Parameter(Mandatory=$true)][ValidateNotNull()]$file,
            [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$BucketName,
            [Parameter(Mandatory=$true)][ValidateNotNull()]$Prefix,
            [bool]$UseAWSCLI
        )

        $key = $Prefix + $file.Name

        try {
            if ($UseAWSCLI) {
                aws s3 cp $file.FullName "s3://$BucketName/$key" --quiet
            }
            else {
                # Fallback to AWS PowerShell Module
                Write-S3Object -BucketName $BucketName -Key $key -File $file.FullName -ErrorAction Stop
            }

            # If successful, output a status message
            "Successfully copied $($file.Name) to s3://$BucketName/$key"
        }
        catch {
            Write-Error "Failed to copy $($file.Name) to s3://$BucketName/$key. Error: $_"
        }
    }

    # Begin parallel copying using background jobs
    $jobs = @()
    foreach ($file in $archiveFiles) {
        # Throttle to MaxParallelFiles by waiting for active jobs to complete
        while ((Get-Job | Where-Object { $_.State -eq 'Running' }).Count -ge $MaxParallelFiles) {
            Start-Sleep -Milliseconds 200
        }

        if ($PSCmdlet.ShouldProcess("s3://$BucketName/$Prefix$file", "Copy")) {
            Write-Verbose "Queueing copy job for $($file.FullName) to s3://$BucketName/$Prefix"
            $jobs += Start-Job -ScriptBlock $jobScript -ArgumentList $file, $BucketName, $Prefix, $useAWSCLI
        }
    }

    # Wait for all jobs to complete
    if ($jobs) {
        Write-Verbose "Waiting for all copy jobs to finish..."
        $jobs | Wait-Job

        # Retrieve and display the job results
        foreach ($job in $jobs) {
            $output = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($output) {
                Write-Verbose $output
            }
            Remove-Job -Job $job
        }

        Write-Verbose "Completed copying $($archiveFiles.Count) files to S3 bucket '$BucketName'."
    }
    else {
        Write-Verbose "No copy operations were initiated."
    }
}
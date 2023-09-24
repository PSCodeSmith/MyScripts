<#
	.SYNOPSIS
		Script to back up Group Policy Objects, generate reports, and export GPO link information.
	
	.DESCRIPTION
		This script includes a function that performs a comprehensive backup of GPOs, generates HTML and XML reports,
		and exports information about GPO links and WMI filters into CSV files.
	
	.PARAMETER DaysToKeep
		Specifies the number of days to keep backup files. Default is 7 days.
	
	.PARAMETER BackupLocation
		Specifies the directory where the backup files will be stored.
	
	.EXAMPLE
		.\BackupAndReportGPOs.ps1 -DaysToKeep 30 -BackupLocation "C:\Backup"
	
	.OUTPUTS
		Backup files, Reports, and CSV files
	
	.NOTES
		Author: Micah
	
	.INPUTS
		DaysToKeep, BackupLocation
#>
param
(
	[ValidateRange(1, 365)]
	[int]$DaysToKeep = 7,
	[Parameter(Mandatory = $true)]
	[ValidateScript({
			if (Test-Path $_ -PathType Container)
			{
				$true
			}
			else
			{
				throw "The path $_ does not exist or is not a directory."
			}
		})]
	[string]$BackupLocation
)

#region Functions

function Backup-GroupPolicy
{
<#
	.SYNOPSIS
		Backs up Group Policy Objects (GPOs), generates GPO reports, and exports GPO link information.
	
	.DESCRIPTION
		This function performs a comprehensive backup of GPOs, generates HTML and XML reports, and exports information
		about GPO links and WMI filters into CSV files.
	
	.PARAMETER BackupLocation
		Specifies the directory where the backup files will be stored.
	
	.PARAMETER DaysToKeep
		Specifies the number of days to keep the backup files.
	
	.EXAMPLE
		Backup-GroupPolicy -BackupLocation "C:\Backup" -DaysToKeep 30
	
	.OUTPUTS
		Backup files, Reports, and CSV files
	
	.NOTES
		Author: Micah
	
	.INPUTS
		BackupLocation, DaysToKeep
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$BackupLocation,
		[Parameter(Mandatory = $true)]
		[int]$DaysToKeep
	)
	
	try
	{
		# Initialize variables
		$allGpoLinks = @()
		$startTime = (Get-Date -f "yyyy-MM-dd-hhmm")
		$backupLocation = $BackupLocation.TrimEnd('\')
		
		# Define backup paths
		$paths = @{
			'Temp' = "$backupLocation\GPOBackupsTemp"
			'TempBackups' = "$backupLocation\GPOBackupsTemp\Backup\"
			'TempReports' = "$backupLocation\GPOBackupsTemp\Reports\"
			'TodaysBackup' = "$backupLocation\GPO Backups from $startTime\"
		}
		
		# Create required directories if they don't exist
		$paths.Values | ForEach-Object {
			if (!(Test-Path $_))
			{
				New-Item -ItemType Directory -Path $_
			}
		}
		
		# Clear Temp folder
		Remove-Item ($paths['Temp'] + "\*") -Force -Recurse -Confirm:$False
		
		# Get PDC Emulator
		$pdcEmulator = Get-AdDomainController -Filter { OperationMasterRoles -Like "PDCEmulator" } | Select-Object -ExpandProperty HostName
		
		# Get all GPOs
		$gpos = Get-GPO -All -Server $pdcEmulator
		
		foreach ($gpo in $gpos)
		{
			# Backup each GPO and generate reports
			$gpoDisplayName = $gpo.DisplayName
			$gpoGuid = $gpo.Id.Guid
			$gpoBackupComment = "Backup of $gpoDisplayName on $startTime"
			
			Backup-GPO -Guid $gpoGuid -Path $paths['TempBackups'] -Comment $gpoBackupComment
			
			# Generate GPO reports in HTML and XML formats
			$reportPathHtml = ($paths['TempReports'] + ($gpoDisplayName -replace '[\x2B\x2F\x22\x3A\x3C\x3E\x3F\x5C\x7C]', ' ') + ".html")
			Get-GPOReport -Name $gpoDisplayName -Path $reportPathHtml -ReportType HTML
			
			$reportPathXml = ($paths['TempReports'] + ($gpoDisplayName -replace '[\x2B\x2F\x22\x3A\x3C\x3E\x3F\x5C\x7C]', ' ') + ".xml")
			Get-GPOReport -Name $gpoDisplayName -Path $reportPathXml -ReportType XML
			
			# Extract GPO link information from the XML report
			$gpoLinks = ([xml](Get-Content $reportPathXml)).GPO.LinksTo | Select-Object @{ n = "Policy"; e = { $gpoDisplayName } },
																						@{ n = "OU"; e = { $_.SOMName } },
																						@{ n = "Path"; e = { $_.SOMPath } },
																						@{ n = "LinkStatus"; e = { $_.Enabled } },
																						@{ n = "Enforced"; e = { $_.NoOverride } }
			
			# If no links are found, create a default object
			if ($null -eq $gpoLinks)
			{
				$gpoLinks = [PSCustomObject]@{
					Policy	   = $gpoDisplayName
					OU		   = "NA"
					Path	   = "NA"
					LinkStatus = "NA"
					Enforced   = "NA"
				}
			}
			
			$allGpoLinks += $gpoLinks
		}
		
		# Export all gathered data to CSV
		$allGpoLinks | Export-Csv -Delimiter "`t" -LiteralPath ($paths['TempReports'] + "GPOLinks-$startTime.csv") -NoTypeInformation
		Get-GPOLinkOrder | Export-Csv -Delimiter "`t" -LiteralPath ($paths['TempReports'] + "GPOLinkOrder-$startTime.csv") -NoTypeInformation
		
		# Move backups and compress reports
		Move-Item -Path $paths['TempBackups'] -Destination $paths['TodaysBackup']
		Compress-Archive -Path "$($paths['TempReports'])\*" -DestinationPath ($paths['TodaysBackup'] + "GPOReports-$startTime.zip")
		
		# Remove old backup files based on DaysToKeep
		Get-ChildItem -Path $backupLocation | Where-Object { $_.Name -like "GPO Backups from *" } |
		Where-Object { $_.CreationTime -lt $(Get-Date).AddDays(-$DaysToKeep) } | Remove-Item -Recurse -Force -Confirm:$False
	}
	catch
	{
		Write-Error "An error occurred: $($_.Exception.Message)"
		Exit
	}
}

function Get-GPOLinkOrder
{
<#
	.SYNOPSIS
		Gets the Group Policy Object (GPO) link order for Organizational Units (OUs) in Active Directory.
	
	.DESCRIPTION
		This function retrieves OUs with associated GPO links from Active Directory and details the order in which GPOs are applied.
	
	.EXAMPLE
		Get-GPOLinkOrder
		This example shows how to run the function without any parameters.
	
	.OUTPUTS
		A custom object with the properties OUName, OUDistinguishedName, GPOName, IsLinked, IsEnforced, and GPOrder.
	
	.NOTES
		Author: Micah
	
	.PARAMETER None
		This function takes no parameters.
	
	.INPUTS
		None
#>
	
	try
	{
		# Retrieve all OUs with gPlink property (i.e., those that have GPO links)
		$organizationalUnits = @(Get-ADOrganizationalUnit -Filter * -Properties gPlink | Where-Object { $null -ne $_.gpLink })
		
		foreach ($ou in $organizationalUnits)
		{
			$ouName = $ou.Name
			$ouDistinguishedName = $ou.DistinguishedName
			
			# Split the gPlink string to get individual GPO links
			$gpoLinks = $ou.gPlink.split("][")
			
			# Filter out any empty strings
			$gpoLinks = @($gpoLinks | Where-Object { $_ })
			
			$gpoOrder = $gpoLinks.Count
			
			foreach ($gpoLink in $gpoLinks)
			{
				$gpoName = [adsi]$gpoLink.split(";")[0] | Select-Object -ExpandProperty displayName
				$gpoStatus = $gpoLink.split(";")[1]
				
				$isLinked = $isEnforced = $false
				
				# Determine GPO status based on the number in the string
				switch ($gpoStatus)
				{
					"1" { $isLinked = $false; $isEnforced = $false }
					"2" { $isLinked = $true; $isEnforced = $true }
					"3" { $isLinked = $false; $isEnforced = $true }
					"0" { $isLinked = $true; $isEnforced = $false }
				}
				
				# Create output object and populate its properties
				$outputObject = New-Object -TypeName PSObject -Property @{
					OUName			    = $ouName
					OUDistinguishedName = $ouDistinguishedName
					GPOName			    = $gpoName
					IsLinked		    = $isLinked
					IsEnforced		    = $isEnforced
					GPOrder			    = $gpoOrder
				}
				
				# Output the object
				Write-Output $outputObject
				
				$gpoOrder--
			}
		}
	}
	catch
	{
		Write-Error "An error occurred: $_"
	}
}

#endregion Functions

# Run the Backup-GroupPolicy function
Backup-GroupPolicy -BackupLocation $BackupLocation -DaysToKeep $DaysToKeep
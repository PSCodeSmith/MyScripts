function Get-GPOMissingPermissions {
	<#
	.SYNOPSIS
		Retrieves GPOs missing specific permissions in a forest.
	
	.DESCRIPTION
		This function queries one or more domains within a forest and identifies Group Policy Objects (GPOs)
		that are missing certain permissions. By default, it can check for Authenticated Users, Domain Computers,
		or either. It supports excluding or including specific domains, optionally skipping RODCs, and gathering
		extended GPO information.
	
	.PARAMETER Forest
		Specifies the forest to query. If not provided, the current forest is assumed.
	
	.PARAMETER ExcludeDomains
		Specifies the domains to exclude from the search. Useful when you have multiple domains and you
		only want to focus on specific ones.
	
	.PARAMETER IncludeDomains
		Specifies the domains to include in the search. If this parameter is used, only the specified domains
		are queried.
	
	.PARAMETER SkipRODC
		If present, Read-Only Domain Controllers (RODCs) are skipped during the search process.
	
	.PARAMETER ExtendedForestInformation
		Allows you to provide pre-retrieved forest information (e.g. from Get-WinADForestDetails) to skip
		retrieving it again. Useful for performance when calling this function multiple times.
	
	.PARAMETER Mode
		Specifies the permission mode to check:
		- AuthenticatedUsers: Check if the GPO grants permissions to Authenticated Users (SID: S-1-5-11).
		- DomainComputers: Check if the GPO grants permissions to the Domain Computers group (SID: domain-515).
		- Either: Check if the GPO grants permissions to either Authenticated Users or Domain Computers.
	
	.PARAMETER Extended
		If specified, gathers extended information about the GPOs, including whether they are linked
		to any containers. This is done by retrieving XML-based GPO reports and annotating GPO objects
		with additional properties.
	
	.EXAMPLE
		Get-GPOMissingPermissions -Mode "Either"
	
		Retrieves all GPOs from the forest that do not have permissions for either Authenticated Users
		or Domain Computers.
	
	.OUTPUTS
		Returns an array of GPO objects that are missing the specified permissions.
	
	.NOTES
		Author: Micah
		Requires:
			- Get-WinADForestDetails (a custom helper function assumed to be available)
			- Get-GPO, Get-GPOReport, Get-GPPermission (from the GroupPolicy module)
			- Appropriate permissions to query and read GPO information.
	
	.INPUTS
		Forest, ExcludeDomains, IncludeDomains, SkipRODC, ExtendedForestInformation, Mode, Extended
	
	#>
	
		[CmdletBinding()]
		param
		(
			[Alias('ForestName')]
			[string]$Forest,
	
			[string[]]$ExcludeDomains,
	
			[Alias('Domain', 'Domains')]
			[string[]]$IncludeDomains,
	
			[switch]$SkipRODC,
	
			[System.Collections.IDictionary]$ExtendedForestInformation,
	
			[ValidateSet('AuthenticatedUsers', 'DomainComputers', 'Either')]
			[string]$Mode = 'Either',
	
			[switch]$Extended
		)
	
		# If forest information is not provided, retrieve it
		if (-not $ExtendedForestInformation) {
			$ForestInformation = Get-WinADForestDetails -Forest $Forest -IncludeDomains $IncludeDomains -ExcludeDomains $ExcludeDomains -SkipRODC:$SkipRODC
		} else {
			$ForestInformation = $ExtendedForestInformation
		}
	
		$allMissingPermissions = @()
	
		# Iterate through each domain in the forest information
		foreach ($Domain in $ForestInformation.Domains) {
			$QueryServer = $ForestInformation['QueryServers']["$Domain"].HostName[0]
	
			# Retrieve GPOs - either with extended info or basic info
			$GPOs = if ($Extended) {
				Get-AdvancedGPOInformation -QueryServer $QueryServer -Domain $Domain -ForestInformation $ForestInformation
			} else {
				Get-GPO -All -Domain $Domain -Server $QueryServer
			}
	
			$DomainInformation = Get-ADDomain -Server $QueryServer
			$DomainComputersSID = "{0}-515" -f $DomainInformation.DomainSID.Value
	
			# Identify GPOs missing permissions based on the chosen mode
			$MissingPermissions = Check-MissingPermissions -GPOs $GPOs -Mode $Mode -QueryServer $QueryServer -Domain $Domain -DomainComputersSID $DomainComputersSID
	
			$allMissingPermissions += $MissingPermissions
		}
	
		return $allMissingPermissions
	}
	
	# Helper function to gather advanced GPO information
	function Get-AdvancedGPOInformation {
		[CmdletBinding()]
		param
		(
			[Parameter(Mandatory=$true)]
			[string]$QueryServer,
	
			[Parameter(Mandatory=$true)]
			[string]$Domain,
	
			[Parameter(Mandatory=$true)]
			[System.Collections.IDictionary]$ForestInformation
		)
	
		# Retrieve all GPOs and add extended properties by parsing GPO XML reports
		Get-GPO -All -Domain $Domain -Server $QueryServer | ForEach-Object {
			$GPO = $_
			[xml]$XMLContent = Get-GPOReport -ID $GPO.Id.Guid -ReportType XML -Server $ForestInformation.QueryServers[$Domain].HostName[0] -Domain $Domain
	
			# Annotate GPO object with additional properties
			Add-Member -InputObject $GPO -MemberType NoteProperty -Name 'LinksTo' -Value $XMLContent.GPO.LinksTo
			Add-Member -InputObject $GPO -MemberType NoteProperty -Name 'Linked' -Value ([bool]($XMLContent.GPO.LinksTo))
			
			$GPO
		}
	}
	
	# Helper function to check for missing permissions on a list of GPOs
	function Check-MissingPermissions {
		[CmdletBinding()]
		param
		(
			[Parameter(Mandatory=$true)]
			[Object[]]$GPOs,
	
			[Parameter(Mandatory=$true)]
			[ValidateSet('AuthenticatedUsers','DomainComputers','Either')]
			[string]$Mode,
	
			[Parameter(Mandatory=$true)]
			[string]$QueryServer,
	
			[Parameter(Mandatory=$true)]
			[string]$Domain,
	
			[Parameter(Mandatory=$true)]
			[string]$DomainComputersSID
		)
	
		$MissingPermissions = @()
	
		foreach ($GPO in $GPOs) {
			# Get all trustees for this GPO
			$Trustees = Get-GPPermission -Guid $GPO.Id -All -Server $QueryServer -DomainName $Domain -ErrorAction Stop | Select-Object -ExpandProperty Trustee
	
			$HasAuthUsers = $false
			$HasDomainComputers = $false
	
			if ($Mode -eq 'Either' -or $Mode -eq 'AuthenticatedUsers') {
				$HasAuthUsers = $Trustees | Where-Object { $_.Sid.Value -eq 'S-1-5-11' } | ForEach-Object { $true }
			}
	
			if ($Mode -eq 'Either' -or $Mode -eq 'DomainComputers') {
				$HasDomainComputers = $Trustees | Where-Object { $_.Sid.Value -eq $DomainComputersSID } | ForEach-Object { $true }
			}
	
			switch ($Mode) {
				'Either' {
					if (-not $HasAuthUsers -and -not $HasDomainComputers) {
						$MissingPermissions += $GPO
					}
				}
				'AuthenticatedUsers' {
					if (-not $HasAuthUsers) {
						$MissingPermissions += $GPO
					}
				}
				'DomainComputers' {
					if (-not $HasDomainComputers) {
						$MissingPermissions += $GPO
					}
				}
			}
		}
	
		return $MissingPermissions
	}
	
	# Example usage:
	# $MissingPermissions = Get-GPOMissingPermissions -Mode "Either"
	# $MissingPermissions | Format-Table -AutoSize
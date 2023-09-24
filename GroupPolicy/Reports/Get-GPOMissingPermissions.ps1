function Get-GPOMissingPermissions
{
<#
	.SYNOPSIS
		Retrieves GPOs missing permissions in a forest.
	
	.DESCRIPTION
		This function retrieves information about GPOs that lack specific permissions.
	
	.PARAMETER Forest
		Specifies the forest to query.
	
	.PARAMETER ExcludeDomains
		Specifies the domains to exclude from the search.
	
	.PARAMETER IncludeDomains
		Specifies the domains to include in the search.
	
	.PARAMETER SkipRODC
		Specifies whether to skip RODCs during the search.
	
	.PARAMETER ExtendedForestInformation
		Allows providing extended forest information.
	
	.PARAMETER Mode
		Specifies the permissions mode to check: AuthenticatedUsers, DomainComputers, or Either.
	
	.PARAMETER Extended
		Specifies whether to gather extended information about GPOs.
	
	.EXAMPLE
		Get-GPOMissingPermissions -Mode "Either"
	
	.OUTPUTS
		Array of missing permissions
	
	.NOTES
		Author: Micah
	
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
	
	# Initialize Forest Information
	if (-not $ExtendedForestInformation)
	{
		$ForestInformation = Get-WinADForestDetails -Forest $Forest -IncludeDomains $IncludeDomains -ExcludeDomains $ExcludeDomains -SkipRODC:$SkipRODC
	}
	else
	{
		$ForestInformation = $ExtendedForestInformation
	}
	
	# Process each domain in the forest
	foreach ($Domain in $ForestInformation.Domains)
	{
		$QueryServer = $ForestInformation['QueryServers']["$Domain"].HostName[0]
		
		if ($Extended)
		{
			# Gather extended GPO information
			$GPOs = Get-AdvancedGPOInformation -QueryServer $QueryServer -Domain $Domain -ForestInformation $ForestInformation
		}
		else
		{
			$GPOs = Get-GPO -All -Domain $Domain -Server $QueryServer
		}
		
		$DomainInformation = Get-ADDomain -Server $QueryServer
		$DomainComputersSID = $('{0}-515' -f $DomainInformation.DomainSID.Value)
		
		# Check for missing permissions
		$MissingPermissions = Check-MissingPermissions -GPOs $GPOs -Mode $Mode -QueryServer $QueryServer -Domain $Domain -DomainComputersSID $DomainComputersSID
		
		$MissingPermissions
	}
}

# Helper function to gather advanced GPO information
function Get-AdvancedGPOInformation
{
	param
	(
		[string]$QueryServer,
		[string]$Domain,
		[System.Collections.IDictionary]$ForestInformation
	)
	
	# Gather and annotate GPOs
	Get-GPO -All -Domain $Domain -Server $QueryServer | ForEach-Object {
		[xml]$XMLContent = Get-GPOReport -ID $_.ID.Guid -ReportType XML -Server $ForestInformation.QueryServers[$Domain].HostName[0] -Domain $Domain
		Add-Member -InputObject $_ -MemberType NoteProperty -Name 'LinksTo' -Value $XMLContent.GPO.LinksTo
		Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Linked' -Value ($null -ne $XMLContent.GPO.LinksTo)
		$_
	}
}

# Helper function to check for missing permissions
function Check-MissingPermissions
{
	param
	(
		[Object[]]$GPOs,
		[string]$Mode,
		[string]$QueryServer,
		[string]$Domain,
		[string]$DomainComputersSID
	)
	
	$MissingPermissions = @()
	
	foreach ($GPO in $GPOs)
	{
		$Permissions = Get-GPPermission -Guid $GPO.Id -All -Server $QueryServer -DomainName $Domain | Select-Object -ExpandProperty Trustee
		
		$GPOPermissionForAuthUsers = $null
		$GPOPermissionForDomainComputers = $null
		
		if ($Mode -eq 'Either' -or $Mode -eq 'AuthenticatedUsers')
		{
			$GPOPermissionForAuthUsers = $Permissions | Where-Object { $_.Sid.Value -eq 'S-1-5-11' }
		}
		
		if ($Mode -eq 'Either' -or $Mode -eq 'DomainComputers')
		{
			$GPOPermissionForDomainComputers = $Permissions | Where-Object { $_.Sid.Value -eq $DomainComputersSID }
		}
		
		if ($Mode -eq 'Either')
		{
			if (-not $GPOPermissionForAuthUsers -and -not $GPOPermissionForDomainComputers)
			{
				$MissingPermissions += $GPO
			}
		}
		elseif ($Mode -eq 'AuthenticatedUsers')
		{
			if (-not $GPOPermissionForAuthUsers)
			{
				$MissingPermissions += $GPO
			}
		}
		elseif ($Mode -eq 'DomainComputers')
		{
			if (-not $GPOPermissionForDomainComputers)
			{
				$MissingPermissions += $GPO
			}
		}
	}
	
	return $MissingPermissions
}

# Main Execution
$MissingPermissions = Get-GPOMissingPermissions -Mode "Either"
$MissingPermissions | Format-Table -AutoSize

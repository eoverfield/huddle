<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Given a tenant, go and get list of tenant level sites, with filtering

.EXAMPLE
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1"

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Credentials $creds
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Url "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam21" -Credentials $creds
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "STS#0" -Credentials $creds
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "BLANKINTERNETCONTAINER#0" -Credentials $creds
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "DEV#0" -Credentials $creds
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "GROUP#0" -Credentials $creds
PS C:\> .\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "SITEPAGEPUBLISHING#0" -Credentials $creds
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Tenant Name")]
    [String]
    $Tenant,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Filter Url")]
    [String]
    $Url,

	# STS#0, BLANKINTERNETCONTAINER#0, DEV#0
	[Parameter(Mandatory = $false, HelpMessage="Filter Template")]
    [String]
    $Template
)

#includes
. "./utilities.ps1"

$Credentials = Verify-SPOCredentials -Credentials $Credentials -Url "https://$($Tenant)-admin.sharepoint.com"

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|          Get Tenant level site collection(s)          |"
Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target tenant: $($Tenant)"
Write-Host ""

try
{
	$TenantUrl = "https://" + $Tenant + "-admin.sharepoint.com"

	#connent to site
	Write-Host -ForegroundColor White "Connecting to the tenant: $($TenantUrl)"
	Connect-PnPOnline -Url $TenantUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to the tenant"

	$sites = Get-PnPTenantSite

	if ($sites -ne $null)
	{
		Write-Host ""
		Write-Host -ForegroundColor Green "Tenant sites retrieved"

		if(![string]::IsNullOrEmpty($Url))
		{
			$sites = Get-PnPTenantSite -Filter "Url -eq $($Url)"
		}

		if(![string]::IsNullOrEmpty($Template))
		{
			$sites = Get-PnPTenantSite -WebTemplate $Template
			#$sites = $sites | Where-Object {$_.Template -eq $Template}
		}

		if ($sites -ne $null)
		{
			$sites
		}
		else {
			Write-Host -ForegroundColor Yellow "After filtering, no valid tenant level sites found to match your request"
		}

    }
	else {
		Write-Host ""
		Write-Host -ForegroundColor Yellow "Unable to find tenant level sites, ignoring request"
	}
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
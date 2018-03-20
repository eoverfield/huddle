<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/Remove-PnPTenantSite.md

.SYNOPSIS
Remove Tenant level site collection, i.e. for classic team, publishing, etc sites

.EXAMPLE
PS C:\> .\Remove-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/classic-team" 

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> .\Remove-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/classic-team"  -Credentials $creds
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Tenant Name")]
    [String]
    $Tenant,

    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target site to remove, e.g. 'https://intranet.mydomain.com/sites/classic-team'")]
    [String]
    $TargetWebUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Force?")]
    [Switch]
    $Force,

	[Parameter(Mandatory = $false, HelpMessage="Wait?")]
    [Switch]
    $Wait,

	[Parameter(Mandatory = $false, HelpMessage="Remove from Recycle Bin as well?")]
    [Switch]
    $SkipRecycleBin
)

#includes
. "./utilities.ps1"

$Credentials = Verify-SPOCredentials -Credentials $Credentials -Url "https://$($Tenant)-admin.sharepoint.com"

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|          Remove Tenant Level Site Collection          |"
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

	Write-Host ""
	Write-Host -ForegroundColor White "Looking for site collection to remove from tenant: $($TargetWebUrl)"

	$web = Get-PnPTenantSite -Filter "Url -eq $($TargetWebUrl)"

	if ($web -ne $null)
	{
		Write-Host -ForegroundColor Green "Tenant site found and may be removed now: $($web.Url)"

		$exp = 'Remove-PnPTenantSite -Url $($web.Url)'

		#Force
		if ($Force)
		{
			$exp += ' -Force'
		}

		#Wait
		if ($Wait)
		{
			$exp += ' -Wait'
		}

		#SkipRecycleBin
		if ($SkipRecycleBin)
		{
			$exp += ' -SkipRecycleBin'
		}
		
		Write-Host -ForegroundColor White "Going to remove site using the following command:"
		Write-Host $exp

		Invoke-Expression $exp

		Write-Host -ForegroundColor Green "Site removed"
    }
	else {
		Write-Host ""
		Write-Host -ForegroundColor Yellow "Unable to find site to remove, ignoring request"
	}
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
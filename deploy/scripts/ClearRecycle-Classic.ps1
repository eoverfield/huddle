<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/Clear-PnPTenantRecycleBinItem.md

.SYNOPSIS
Clear a given removed tenant level site colllection from the recycle bin as well.

.EXAMPLE
PS C:\> .\ClearRecycle-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/classic-team" 

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> .\ClearRecycle-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/classic-team" -Credentials $creds
PS C:\> .\ClearRecycle-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/classic-team" -Force -Wait -Credentials $creds
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Tenant Name")]
    [String]
    $Tenant,

    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target site to clear from recycle bin, e.g. 'https://intranet.mydomain.com/sites/classic-team'")]
    [String]
    $TargetWebUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Force?")]
    [Switch]
    $Force,

	[Parameter(Mandatory = $false, HelpMessage="Wait?")]
    [Switch]
    $Wait
)

#includes
. "./utilities.ps1"

$Credentials = Verify-SPOCredentials -Credentials $Credentials -Url "https://$($Tenant)-admin.sharepoint.com"

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|  Clear Tenant Level Site Collection Recycle Bin Item  |"
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
	Write-Host -ForegroundColor White "Looking for site collection to clear from recycle bin: $($TargetWebUrl)"

	$recycleBin = Get-PnPTenantRecycleBinItem
	$item = $recycleBin | Where-Object {$_.Url -eq $TargetWebUrl}

	if ($item -ne $null)
	{
		Write-Host -ForegroundColor Green "Tenant site found in recycle bin and may be cleared now: $($item.Url)"

		$exp = 'Clear-PnPTenantRecycleBinItem -Url $($item.Url)'

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
		
		Write-Host -ForegroundColor White "Going to clear site using the following command:"
		Write-Host $exp

		Invoke-Expression $exp

		Write-Host -ForegroundColor Green "Site cleared"
    }
	else {
		Write-Host ""
		Write-Host -ForegroundColor Yellow "Unable to find site to clear, ignoring request"
	}
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
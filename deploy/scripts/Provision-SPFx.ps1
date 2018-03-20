<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Upload SPFx solution to tenant app catalog

.EXAMPLE
PS C:\> .\Provision-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" 

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Provision-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials $creds
PS C:\> .\Provision-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials "yourStoredCredentialName"
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target tenant app catalog web, e.g. 'https://intranet.mydomain.com/sites/apps'")]
    [String]
    $TargetCatalogWebUrl,

    [Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Optional SharePoint App Id with Full Control")]
    $AppId,

	[Parameter(Mandatory = $false, HelpMessage="Optional SharePoint App Secret for App Id with Full Control")]
    $AppSecret
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url $TargetCatalogWebUrl -AppId $AppId -AppSecret $AppSecret

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|        SPFx Provisioning to Tenant App Catalog        |"
Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target tenant app catalog web: $($TargetCatalogWebUrl)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to web"
	if ($Credentials -ne $null)
	{
		Write-Host -ForegroundColor Yellow "using credentials"
		Connect-PnPOnline -Url $TargetCatalogWebUrl -Credentials $Credentials
	}
	else 
	{
		Write-Host -ForegroundColor Yellow "using app id"
		Connect-PnPOnline -Url $TargetCatalogWebUrl -AppId $AppId -AppSecret $AppSecret
	}
	Write-Host -ForegroundColor Green "Connected to $($TargetCatalogWebUrl)"


	Write-Host ""
	Write-Host -ForegroundColor White "Verifying SPFX solution may be installed to tenant app catalog"
	$apps = Get-PnPApp 
	$app = $apps | Where-Object {$_.Title -eq 'Collaboration and Communication Blocks'}

	if ($app -ne $null)
	{
		Write-Host -ForegroundColor Yellow "SPFx solution already available, not re-provisioned"
		exit
	}

	#else we can add the solution as not yet found
	Write-Host -ForegroundColor Green "Verified availble for provisioning"
	Write-Host -ForegroundColor White "Provisioning SPFx solution"
	Add-PnPApp -Path .\assets\collab-comm-blocks.sppkg

	#get the app just installed
	$apps = Get-PnPApp 
	$app = $apps | Where-Object {$_.Title -eq 'Collaboration and Communication Blocks'}

	if ($app -ne $null) {
		Write-Host -ForegroundColor Green "SPFx solution deployed"

		#go ahead and publish app / solution to make it active
		Write-Host ""
		Write-Host -ForegroundColor White "Publishing tenant scoped SPFx solution: $($app.Id)"
		Publish-PnPApp -Identity $app.Id -SkipFeatureDeployment
		Write-Host -ForegroundColor Green "Tenant scoped SPFx solution deployed and published"	
	}
	else {
		Write-Host -ForegroundColor Red "SPFx solution deployment failed for unknown reason"
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Provisioning Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Upload SPFx assets to a given site, normally a CDN

.EXAMPLE
PS C:\> .\Provision-SPFx-Assets.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommCDN" 

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Provision-SPFx-Assets.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommCDN" -Credentials $creds
PS C:\> .\Provision-SPFx-Assets.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommCDN" -Credentials "yourStoredCredentialName"
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target CDN site, e.g. 'https://intranet.mydomain.com/sites/apps'")]
    [String]
    $TargetWebUrl,

    [Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url $TargetWebUrl

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|            Provisioning SPFx Assets to CDN            |"
Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target CDN web: $($TargetWebUrl)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to CDN web"
	Connect-PnPOnline $TargetWebUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to $($TargetWebUrl)"

	Write-Host ""
	Write-Host -ForegroundColor White "Provisioning SPFx assets to CDN"

	Apply-PnPProvisioningTemplate .\templates\Provision.SPFx.Assets.xml

	Write-Host -ForegroundColor Green "SPFx assets deployed to CDN"

	Write-Host ""
	Write-Host -ForegroundColor Green "Provisioning Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
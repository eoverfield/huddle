<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Remove SPFx solution from tenant app catalog

.EXAMPLE
PS C:\> .\Remove-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" 

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Remove-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials $creds
PS C:\> .\Remove-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials "yourStoredCredentialName"
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target tenant app catalog web, e.g. 'https://intranet.mydomain.com/sites/apps'")]
    [String]
    $TargetCatalogWebUrl,

    [Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url $TargetCatalogWebUrl

Write-Host -ForegroundColor White "----------------------------------------------------------"
Write-Host -ForegroundColor White "|    Remove SPFx Provisioning from Tenant App Catalog    |"
Write-Host -ForegroundColor White "----------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target tenant app catalog web: $($TargetCatalogWebUrl)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to web"
	Connect-PnPOnline $TargetCatalogWebUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to $($TargetCatalogWebUrl)"

	
	Write-Host ""
	Write-Host -ForegroundColor White "Looking for SPFx solution to remove from tenant app catalog"
	$apps = Get-PnPApp 
	$app = $apps | Where-Object {$_.Title -eq 'Collaboration and Communication Blocks'}

	if ($app -ne $null)
	{
		Write-Host -ForegroundColor White "SPFx solution found, removing now"
    	Remove-PnPApp -Identity $app.Id
		Write-Host -ForegroundColor Green "Solution removed"
    }
	else {
		Write-Host -ForegroundColor Yellow "No SPFx solution found that matches title/name, nothing to remove"
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "SPFx solution clean up completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
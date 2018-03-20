<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Remove setup for an existing site

.EXAMPLE
PS C:\> .\Remove-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -DeactivateExtensions -DeactivateSiteSettings -ResetLogo

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Remove-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials $creds -DeactivateExtensions -DeactivateSiteSettings -ResetLogo
PS C:\> .\Remove-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials "yourStoredCredentialName"
PS C:\> .\Remove-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/collabcomm194" -Credentials $creds
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target site to remove set up, e.g. 'https://intranet.mydomain.com/sites/apps'")]
    [String]
    $TargetWebUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Deactivate SPFx Extensions")]
    [Switch]
    $DeactivateExtensions,

	[Parameter(Mandatory = $false, HelpMessage="Deactivate Site Setttings")]
    [Switch]
    $DeactivateSiteSettings,

	[Parameter(Mandatory = $false, HelpMessage="Reset logo")]
    [Switch]
    $ResetLogo
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url $TargetWebUrl

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|          Remove Site Setup and Configuration          |"
Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target site: $($TargetWebUrl)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to web"
	Connect-PnPOnline $TargetWebUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to $($TargetWebUrl)"

	Write-Host ""
	Write-Host -ForegroundColor White "Removing site Settings and configuration"

	#Extensions
	$customAction = Get-PnPCustomAction -Scope Site | where { $_.Name -eq "CollabCommHeaderApplicationCustomizer" }
	if ($customAction -ne $null)
	{
		Write-Host -ForegroundColor White "Removing Header"
    	Remove-PnPCustomAction -Identity $customAction.Id -Scope Site -Force
		Write-Host -ForegroundColor Green "Header removed"
    }

	$customAction = Get-PnPCustomAction -Scope Site | where { $_.Name -eq "CollabCommFooterApplicationCustomizer" }
	if ($customAction -ne $null)
	{
		Write-Host -ForegroundColor White "Removing Footer"
    	Remove-PnPCustomAction -Identity $customAction.Id -Scope Site -Force
		Write-Host -ForegroundColor Green "Footer removed"
    }

	Write-Host -ForegroundColor White "Resetting homepage"
	Set-PnPHomePage -RootFolderRelativeUrl SitePages/Home.aspx
	Write-Host -ForegroundColor Green "Reseting homepage: complete"

	Write-Host -ForegroundColor White "Resetting logo"
	Set-PnPWeb -SiteLogoUrl ""
	Write-Host -ForegroundColor Green "Resetting logo: complete"

	Write-Host ""
	Write-Host -ForegroundColor Green "Site setup removed"

	Write-Host ""
	Write-Host -ForegroundColor Green "Proivisiong Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}